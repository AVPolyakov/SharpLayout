using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
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
            
            byte[] netstandardBytes;
            using (var stream = typeof(Watcher).Assembly.GetManifestResourceStream(typeof(netstandard).FullName + ".dll"))
            {
                netstandardBytes = new byte[stream.Length];
                stream.Read(netstandardBytes, 0, netstandardBytes.Length);
            }
            var netstandardReference = AssemblyMetadata.CreateFromImage(netstandardBytes).GetReference();
            var pdfSharpReference = AssemblyMetadata.CreateFromFile(typeof(PdfDocument).Assembly.Location).GetReference();
            var sharpLayoutReference = AssemblyMetadata.CreateFromFile(typeof(Document).Assembly.Location).GetReference();
            var settingsReference = AssemblyMetadata.CreateFromFile(typeof(WatcherSettings).Assembly.Location).GetReference();
            var context = new Context(netstandardReference, settingsReference, pdfSharpReference, sharpLayoutReference,
                settingsPath, assemblies, parameterFunc, outputPath);

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

        private static IEnumerable<string> GetSourceCodeFileFullNames(WatcherSettings settings, Context context)
        {
            var directoryName = Path.GetDirectoryName(context.SettingsPath);
            return settings.SourceCodeFiles.Concat(new[] {settings.SourceCodeFile})
                .Select(codeFile => Path.Combine(directoryName, codeFile));
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
                var references = context.Assemblies
                    .Select(a => AssemblyMetadata.CreateFromFile(a.Location).GetReference())
                    .ToArray();
                var subContext = new SubContext(references);
                Compile(context, settings, newSettings: true, createPdf: createPdf, subContext);
                foreach (var sourceCodeFile in GetSourceCodeFileFullNames(settings, context))
                    context.Watchers.Add(
                        StartWatcher(sourceCodeFile, 
                            () => Compile(context, settings, 
                                newSettings: false, createPdf: false, subContext)));
            }
            else
            {
                WriteError(settingsChoice.Value2, context);
            }
        }

        private static void Compile(Context context, WatcherSettings settings, bool newSettings, bool createPdf, 
            SubContext subContext)
        {
            try
            {
                var stopwatch = Stopwatch.StartNew();
                var compilation = CSharpCompilation.Create(
                    Guid.NewGuid().ToString("N"),
                    GetSourceCodeFileFullNames(settings, context).Select(
                        codeFile => SyntaxFactory.ParseSyntaxTree(File.ReadAllText(codeFile), path: codeFile)),
                    options: new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary),
                    references: new[]
                    {
                        context.NetstandardReference,
                        context.PdfSharpReference,
                        context.SharpLayoutReference,
                    }.Concat(subContext.References)
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
                var assembly = Assembly.Load(bytes);
                var typeName = Path.GetFileNameWithoutExtension(settings.SourceCodeFile);
                var type = assembly.GetTypes().Single(_ => _.Name == typeName);
                var method = type.GetMethod("AddSection");
                var parameterType = method.GetParameters()[1].ParameterType;
                var dataPath = Path.Combine(Path.GetDirectoryName(GetOutputPath(context)),
                    $"{parameterType.FullName}.json");
                if (newSettings)
                    context.Watchers.Add(StartWatcher(dataPath, 
                        () => Compile(context, settings, newSettings: false, createPdf: false, subContext)));
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
                references: new[]
                {
                    context.NetstandardReference,
                    context.SharpLayoutReference,
                    context.SettingsReference
                }
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
            var settingsProviderType = Assembly.Load(settingsProviderBytes)
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

    internal class SubContext
    {
        public PortableExecutableReference[] References { get; }

        public SubContext(PortableExecutableReference[] references)
        {
            References = references;
        }
    }

    internal class Context
    {
        public PortableExecutableReference NetstandardReference { get; }
        public PortableExecutableReference SettingsReference { get; }
        public PortableExecutableReference PdfSharpReference { get; }
        public PortableExecutableReference SharpLayoutReference { get; }
        public string SettingsPath { get; }
        public Assembly[] Assemblies { get; }
        public Func<ParameterInfo, object> ParameterFunc { get; }
        public string OutputPath { get; }
        public readonly List<FileSystemWatcher> Watchers = new List<FileSystemWatcher>();

        public Context(PortableExecutableReference netstandardReference, PortableExecutableReference settingsReference,
            PortableExecutableReference pdfSharpReference, PortableExecutableReference sharpLayoutReference,
            string settingsPath, Assembly[] assemblies, Func<ParameterInfo, object> parameterFunc, string outputPath)
        {
            NetstandardReference = netstandardReference;
            SettingsReference = settingsReference;
            PdfSharpReference = pdfSharpReference;
            SharpLayoutReference = sharpLayoutReference;
            SettingsPath = settingsPath;
            Assemblies = assemblies;
            ParameterFunc = parameterFunc;
            OutputPath = outputPath;
        }
    }
}
