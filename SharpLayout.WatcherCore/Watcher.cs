using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Loader;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;
using Newtonsoft.Json;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using SharpLayout.WatcherCore.Ref;

namespace SharpLayout.WatcherCore
{
    public static class Watcher
    {
        public static void Start(string settingsPath, Assembly[] assemblies, Func<ParameterInfo, object> parameterFunc,
            string outputPath)
        {
            Document.CollectCallerInfo = true;

            var references = new[] {
                GetReference(typeof(netstandard).FullName + ".dll"),
                GetReference(typeof(netstandard).Namespace + ".System.Runtime.dll"),
                GetReference(typeof(netstandard).Namespace + ".System.Collections.dll"),
                AssemblyMetadata.CreateFromFile(typeof(PdfDocument).Assembly.Location).GetReference(),
                AssemblyMetadata.CreateFromFile(typeof(Document).Assembly.Location).GetReference(),
                AssemblyMetadata.CreateFromFile(typeof(WatcherSettings).Assembly.Location).GetReference()
            };
            var context = new Context(references, settingsPath, parameterFunc, outputPath);

            ProcessSettings(context, createPdf: false);
            StartWatcher(settingsPath, () => ProcessSettings(context, createPdf: false));
            Console.WriteLine("Press 'q' to quit.");
            while (true)
            {
                var keyChar = Console.ReadKey().KeyChar;
                if (keyChar == 'q')
                    break;
                if (keyChar == 'p')
                    ProcessSettings(context, createPdf: true);
            }
        }

        private static PortableExecutableReference GetReference(string fullName)
        {
            byte[] netstandardBytes;
            using (var stream = typeof(Watcher).Assembly.GetManifestResourceStream(fullName))
            {
                netstandardBytes = new byte[stream.Length];
                stream.Read(netstandardBytes, 0, netstandardBytes.Length);
            }
            return AssemblyMetadata.CreateFromImage(netstandardBytes).GetReference();
        }
        
        private static string GetOutputPath(Context context) => 
            Path.Combine(Path.GetDirectoryName(context.SettingsPath), context.OutputPath);
        
        private static void ProcessSettings(Context context, bool createPdf)
        {
            foreach (var watcher in context.Watchers) 
                watcher.Dispose();
            context.Watchers.Clear();
            var settingsChoice = GetSettings(context);
            if (settingsChoice.HasValue1)
            {
                var settings = settingsChoice.Value1;
                var reference1 = CompileAssembly(context, settings.SourceCodeFiles1);
                var reference2 = CompileAssembly(context, settings.SourceCodeFiles2);
                Compile(context, settings, newSettings: true, createPdf: createPdf, reference1, reference2);
                foreach (var sourceCodeFile in settings.SourceCodeFiles1.Select(_ => _.FullPath(context)))
                    context.Watchers.Add(
                        StartWatcher(sourceCodeFile, () => {
                            reference1 = CompileAssembly(context, settings.SourceCodeFiles1);
                            reference2 = CompileAssembly(context, settings.SourceCodeFiles2);
                            Compile(context, settings,
                                newSettings: false, createPdf: false, reference1: reference1, reference2);
                        }));
                foreach (var sourceCodeFile in settings.SourceCodeFiles2.Select(_ => _.FullPath(context)))
                    context.Watchers.Add(
                        StartWatcher(sourceCodeFile, () => {
                            reference2 = CompileAssembly(context, settings.SourceCodeFiles2);
                            Compile(context, settings,
                                newSettings: false, createPdf: false, reference1: reference1, reference2);
                        }));
                context.Watchers.Add(
                    StartWatcher(settings.SourceCodeFile.FullPath(context), () => {
                        Compile(context, settings,
                            newSettings: false, createPdf: false, reference1: reference1, reference2);
                    }));
            }
            else
            {
                WriteError(settingsChoice.Value2, context);
            }
        }

        private static Option<PortableExecutableReference> CompileAssembly(Context context, string[] sourceCodeFiles)
        {
            var compilation = CSharpCompilation.Create(
                Guid.NewGuid().ToString("N"),
                sourceCodeFiles.Select(_ => _.FullPath(context)).Select(
                    codeFile => SyntaxFactory.ParseSyntaxTree(File.ReadAllText(codeFile), path: codeFile)),
                options: new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary),
                references: context.References
            );
            byte[] bytes;
            using (var stream = new MemoryStream())
            {
                var emitResult = compilation.Emit(stream);
                if (!emitResult.Success)
                {
                    WriteError(GetErrorText(emitResult), context);
                    return new Option<PortableExecutableReference>();
                }
                bytes = stream.ToArray();
            }
            LoadAssembly(bytes);
            return AssemblyMetadata.CreateFromImage(bytes).GetReference();
        }

        private static void Compile(Context context, WatcherSettings settings, bool newSettings, bool createPdf,
            Option<PortableExecutableReference> reference1, Option<PortableExecutableReference> reference2)
        {
            if (!reference1.HasValue) return;
            if (!reference2.HasValue) return;
            try
            {
                var stopwatch = Stopwatch.StartNew();
                var codeFile = settings.SourceCodeFile.FullPath(context);
                var compilation = CSharpCompilation.Create(
                    Guid.NewGuid().ToString("N"),
                    new[] {SyntaxFactory.ParseSyntaxTree(File.ReadAllText(codeFile), path: codeFile)},
                    options: new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary),
                    references: context.References.Concat(new[] {reference1.Value, reference2.Value})
                );
                byte[] bytes;
                using (var stream = new MemoryStream())
                {
                    var emitResult = compilation.Emit(stream);
                    if (!emitResult.Success)
                    {
                        WriteError(GetErrorText(emitResult), context);
                        return;
                    }
                    bytes = stream.ToArray();
                }
                var assembly = LoadAssembly(bytes);
                var typeName = Path.GetFileNameWithoutExtension(settings.SourceCodeFile);
                var type = assembly.GetTypes().Single(_ => _.Name == typeName);
                var method = type.GetMethod("AddSection");
                var parameterType = method.GetParameters()[1].ParameterType;
                var dataPath = Path.Combine(Path.GetDirectoryName(GetOutputPath(context)),
                    $"{parameterType.FullName}.json");
                if (newSettings)
                    context.Watchers.Add(StartWatcher(dataPath, 
                        () => Compile(context, settings, newSettings: false, createPdf: false, reference1: reference1,
                            reference2)));
                var deserializeObject = JsonConvert.DeserializeObject(
                    File.ReadAllText(dataPath), parameterType);
                var document = settings.DocumentFunc();
                var parameters = method.GetParameters().Skip(2).Select(info => context.ParameterFunc(info));
                method.Invoke(null, new[] {document, deserializeObject}.Concat(parameters).ToArray());
                if (createPdf)
                {
                    var fileName = Path.Combine(
                        Path.GetDirectoryName(GetOutputPath(context)),
                        $"Temp_{Guid.NewGuid():N}.pdf");
                    document.SavePdf(fileName);
                    Process.Start(new ProcessStartInfo("cmd", $"/c start {fileName}"));
                }
                else
                    document.SavePng(
                        pageNumber: settings.PageNumber,
                        path: GetOutputPath(context),
                        resolution: settings.Resolution);
                stopwatch.Stop();
                Console.WriteLine($"Done. {DateTime.Now:HH:mm:ss} {stopwatch.ElapsedMilliseconds}ms");
            }
            catch (Exception e)
            {
                if (e is TargetInvocationException exception)
                    if (exception.InnerException != null)
                    {
                        WriteError(exception.InnerException.Message, context);
                        return;
                    }
                WriteError(e.Message, context);
            }
        }

        private static string FullPath(this string sourceCodeFile, Context context)
        {
            return Path.Combine(Path.GetDirectoryName(context.SettingsPath), sourceCodeFile);
        }

        private static Assembly LoadAssembly(byte[] bytes)
        {
            using (var stream = new MemoryStream(bytes))
                return AssemblyLoadContext.Default.LoadFromStream(stream);
        }

        private static FileSystemWatcher StartWatcher(string path, Action action)
        {
            var watcher = new FileSystemWatcher
            {
                Path = Path.GetDirectoryName(path),
                NotifyFilter = CombineAllNotifyFilters()
            };
            string oldText = null;
            var semaphore = new SemaphoreSlim(1, 1);

            async Task Handle(FileSystemEventArgs e)
            {
                if (string.Compare(e.FullPath, path, StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    async Task<string> GetText()
                    {
                        while (true)
                        {
                            try
                            {
                                return File.ReadAllText(e.FullPath);
                            }
                            catch (IOException)
                            {
                                await Task.Delay(20);
                            }
                        }
                    }

                    await semaphore.WaitAsync();
                    try
                    {
                        var text = await GetText();
                        if (oldText != text)
                        {
                            action();
                        }
                        oldText = text;
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                }
            }

            watcher.Changed += async (s, e) => await Handle(e);
            watcher.Renamed += async (s, e) => await Handle(e);
            watcher.Created += async (s, e) => await Handle(e);
            watcher.Deleted += async (s, e) => await Handle(e);
            watcher.Filter = Path.GetFileName(path);
            watcher.EnableRaisingEvents = true;
            return watcher;
        }
        
        private static Choice<WatcherSettings, string> GetSettings(Context context)
        {
            var settingsPath = context.SettingsPath;
            if (!File.Exists(settingsPath))
                return $"File not fount, {settingsPath}";
            var compilation = CSharpCompilation.Create(
                Guid.NewGuid().ToString("N"),
                new[] {SyntaxFactory.ParseSyntaxTree(File.ReadAllText(settingsPath), path: settingsPath)},
                options: new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary),
                references: context.References
            );
            byte[] settingsProviderBytes;
            using (var stream = new MemoryStream())
            {
                var emitResult = compilation.Emit(stream);
                if (!emitResult.Success)
                    return GetErrorText(emitResult);
                else
                    settingsProviderBytes = stream.ToArray();
            }
            var settingsProviderType = LoadAssembly(settingsProviderBytes)
                .GetTypes().Single(type => type.Name == "WatcherSettingsProvider");
            return (WatcherSettings) settingsProviderType.GetMethod("GetSettings")
                .Invoke(null, new object[]{Path.GetDirectoryName(settingsPath)});
        }
        
        private static string GetErrorText(EmitResult emitResult) => 
            $"{emitResult.Diagnostics.FirstOrDefault(diagnostic => diagnostic.Severity == DiagnosticSeverity.Error)}";

        private static void WriteError(string errorText, Context context)
        {
            var text = $"ERROR! {errorText}";
            Console.WriteLine(text);
            var document = new Document();
            var settings = new PageSettings();
            settings.LeftMargin = settings.TopMargin = settings.RightMargin = settings.BottomMargin = Util.Cm(0.5);
            document.Add(new Section(settings).Add(new Paragraph()
                .Add(text, new Font("Consolas", 9.5, XFontStyle.Regular, new XPdfFontOptions(PdfFontEncoding.Unicode)))));
            document.SavePng(0, GetOutputPath(context), 120);
        }
        
        private static NotifyFilters CombineAllNotifyFilters()
        {
            return Enum.GetValues(typeof(NotifyFilters)).Cast<NotifyFilters>()
                .Aggregate((filters, notifyFilters) => filters | notifyFilters);
        }
    }
    
    internal class Context
    {
        public PortableExecutableReference[] References { get; }
        public string SettingsPath { get; }
        public Func<ParameterInfo, object> ParameterFunc { get; }
        public string OutputPath { get; }
        public readonly List<FileSystemWatcher> Watchers = new List<FileSystemWatcher>();

        public Context(PortableExecutableReference[] references, string settingsPath, 
            Func<ParameterInfo, object> parameterFunc, string outputPath)
        {
            References = references;
            SettingsPath = settingsPath;
            ParameterFunc = parameterFunc;
            OutputPath = outputPath;
        }
    }
}
