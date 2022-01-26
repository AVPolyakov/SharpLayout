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
using SharpLayout.ImageRendering;
using SharpLayout.WatcherCore.Ref;

namespace SharpLayout.WatcherCore
{
    public static class Watcher
    {
        private static readonly string[] referenceNames = {
            "Microsoft.CSharp.dll",
            "Microsoft.VisualBasic.Core.dll",
            "Microsoft.VisualBasic.dll",
            "Microsoft.Win32.Primitives.dll",
            "Microsoft.Win32.Registry.dll",
            "mscorlib.dll",
            "netstandard.dll",
            "System.AppContext.dll",
            "System.Buffers.dll",
            "System.Collections.Concurrent.dll",
            "System.Collections.dll",
            "System.Collections.Immutable.dll",
            "System.Collections.NonGeneric.dll",
            "System.Collections.Specialized.dll",
            "System.ComponentModel.Annotations.dll",
            "System.ComponentModel.DataAnnotations.dll",
            "System.ComponentModel.dll",
            "System.ComponentModel.EventBasedAsync.dll",
            "System.ComponentModel.Primitives.dll",
            "System.ComponentModel.TypeConverter.dll",
            "System.Configuration.dll",
            "System.Console.dll",
            "System.Core.dll",
            "System.Data.Common.dll",
            "System.Data.DataSetExtensions.dll",
            "System.Data.dll",
            "System.Diagnostics.Contracts.dll",
            "System.Diagnostics.Debug.dll",
            "System.Diagnostics.DiagnosticSource.dll",
            "System.Diagnostics.FileVersionInfo.dll",
            "System.Diagnostics.Process.dll",
            "System.Diagnostics.StackTrace.dll",
            "System.Diagnostics.TextWriterTraceListener.dll",
            "System.Diagnostics.Tools.dll",
            "System.Diagnostics.TraceSource.dll",
            "System.Diagnostics.Tracing.dll",
            "System.dll",
            "System.Drawing.dll",
            "System.Drawing.Primitives.dll",
            "System.Dynamic.Runtime.dll",
            "System.Formats.Asn1.dll",
            "System.Globalization.Calendars.dll",
            "System.Globalization.dll",
            "System.Globalization.Extensions.dll",
            "System.IO.Compression.Brotli.dll",
            "System.IO.Compression.dll",
            "System.IO.Compression.FileSystem.dll",
            "System.IO.Compression.ZipFile.dll",
            "System.IO.dll",
            "System.IO.FileSystem.AccessControl.dll",
            "System.IO.FileSystem.dll",
            "System.IO.FileSystem.DriveInfo.dll",
            "System.IO.FileSystem.Primitives.dll",
            "System.IO.FileSystem.Watcher.dll",
            "System.IO.IsolatedStorage.dll",
            "System.IO.MemoryMappedFiles.dll",
            "System.IO.Pipes.AccessControl.dll",
            "System.IO.Pipes.dll",
            "System.IO.UnmanagedMemoryStream.dll",
            "System.Linq.dll",
            "System.Linq.Expressions.dll",
            "System.Linq.Parallel.dll",
            "System.Linq.Queryable.dll",
            "System.Memory.dll",
            "System.Net.dll",
            "System.Net.Http.dll",
            "System.Net.Http.Json.dll",
            "System.Net.HttpListener.dll",
            "System.Net.Mail.dll",
            "System.Net.NameResolution.dll",
            "System.Net.NetworkInformation.dll",
            "System.Net.Ping.dll",
            "System.Net.Primitives.dll",
            "System.Net.Requests.dll",
            "System.Net.Security.dll",
            "System.Net.ServicePoint.dll",
            "System.Net.Sockets.dll",
            "System.Net.WebClient.dll",
            "System.Net.WebHeaderCollection.dll",
            "System.Net.WebProxy.dll",
            "System.Net.WebSockets.Client.dll",
            "System.Net.WebSockets.dll",
            "System.Numerics.dll",
            "System.Numerics.Vectors.dll",
            "System.ObjectModel.dll",
            "System.Reflection.DispatchProxy.dll",
            "System.Reflection.dll",
            "System.Reflection.Emit.dll",
            "System.Reflection.Emit.ILGeneration.dll",
            "System.Reflection.Emit.Lightweight.dll",
            "System.Reflection.Extensions.dll",
            "System.Reflection.Metadata.dll",
            "System.Reflection.Primitives.dll",
            "System.Reflection.TypeExtensions.dll",
            "System.Resources.Reader.dll",
            "System.Resources.ResourceManager.dll",
            "System.Resources.Writer.dll",
            "System.Runtime.CompilerServices.Unsafe.dll",
            "System.Runtime.CompilerServices.VisualC.dll",
            "System.Runtime.dll",
            "System.Runtime.Extensions.dll",
            "System.Runtime.Handles.dll",
            "System.Runtime.InteropServices.dll",
            "System.Runtime.InteropServices.RuntimeInformation.dll",
            "System.Runtime.Intrinsics.dll",
            "System.Runtime.Loader.dll",
            "System.Runtime.Numerics.dll",
            "System.Runtime.Serialization.dll",
            "System.Runtime.Serialization.Formatters.dll",
            "System.Runtime.Serialization.Json.dll",
            "System.Runtime.Serialization.Primitives.dll",
            "System.Runtime.Serialization.Xml.dll",
            "System.Security.AccessControl.dll",
            "System.Security.Claims.dll",
            "System.Security.Cryptography.Algorithms.dll",
            "System.Security.Cryptography.Cng.dll",
            "System.Security.Cryptography.Csp.dll",
            "System.Security.Cryptography.Encoding.dll",
            "System.Security.Cryptography.OpenSsl.dll",
            "System.Security.Cryptography.Primitives.dll",
            "System.Security.Cryptography.X509Certificates.dll",
            "System.Security.dll",
            "System.Security.Principal.dll",
            "System.Security.Principal.Windows.dll",
            "System.Security.SecureString.dll",
            "System.ServiceModel.Web.dll",
            "System.ServiceProcess.dll",
            "System.Text.Encoding.CodePages.dll",
            "System.Text.Encoding.dll",
            "System.Text.Encoding.Extensions.dll",
            "System.Text.Encodings.Web.dll",
            "System.Text.Json.dll",
            "System.Text.RegularExpressions.dll",
            "System.Threading.Channels.dll",
            "System.Threading.dll",
            "System.Threading.Overlapped.dll",
            "System.Threading.Tasks.Dataflow.dll",
            "System.Threading.Tasks.dll",
            "System.Threading.Tasks.Extensions.dll",
            "System.Threading.Tasks.Parallel.dll",
            "System.Threading.Thread.dll",
            "System.Threading.ThreadPool.dll",
            "System.Threading.Timer.dll",
            "System.Transactions.dll",
            "System.Transactions.Local.dll",
            "System.ValueTuple.dll",
            "System.Web.dll",
            "System.Web.HttpUtility.dll",
            "System.Windows.dll",
            "System.Xml.dll",
            "System.Xml.Linq.dll",
            "System.Xml.ReaderWriter.dll",
            "System.Xml.Serialization.dll",
            "System.Xml.XDocument.dll",
            "System.Xml.XmlDocument.dll",
            "System.Xml.XmlSerializer.dll",
            "System.Xml.XPath.dll",
            "System.Xml.XPath.XDocument.dll",
            "WindowsBase.dll",
        };

        public static void Start(string settingsPath, Assembly[] assemblies, Func<ParameterInfo, object> parameterFunc,
            string outputPath, bool startExternalProcess = false)
        {
            Document.CollectCallerInfo = true;

            var references = referenceNames
                .Select(file => GetReference($"{typeof(Refs).Namespace}.{file}"))
	            .Concat(new[]
	            {
		            AssemblyMetadata.CreateFromFile(typeof(PdfDocument).Assembly.Location).GetReference(),
		            AssemblyMetadata.CreateFromFile(typeof(Document).Assembly.Location).GetReference(),
		            AssemblyMetadata.CreateFromFile(typeof(WatcherSettings).Assembly.Location).GetReference()
	            })
	            .Concat(assemblies.Select(a => AssemblyMetadata.CreateFromFile(a.Location).GetReference()))
	            .ToArray();
            var context = new Context(references, settingsPath, parameterFunc, outputPath, startExternalProcess);
            ProcessSettings(context, createPdf: false);
            using var watcher = StartWatcher(settingsPath, () => ProcessSettings(context, createPdf: false));
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
            byte[] bytes;
            using (var stream = typeof(Watcher).Assembly.GetManifestResourceStream(fullName))
            {
                bytes = new byte[stream.Length];
                stream.Read(bytes, 0, bytes.Length);
            }
            return AssemblyMetadata.CreateFromImage(bytes).GetReference();
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
                var reference1 = CompileAssembly(context, settings.SourceCodeFiles1, new Option<PortableExecutableReference>(), settings);
                var reference2 = CompileAssembly(context, settings.SourceCodeFiles2, reference1, settings);
                Compile(context, settings, newSettings: true, createPdf: createPdf, reference1, reference2);
                foreach (var sourceCodeFile in settings.SourceCodeFiles1.Select(_ => _.FullPath(context)))
                    context.Watchers.Add(
                        StartWatcher(sourceCodeFile, () => {
                            reference1 = CompileAssembly(context, settings.SourceCodeFiles1, new Option<PortableExecutableReference>(), settings);
                            reference2 = CompileAssembly(context, settings.SourceCodeFiles2, reference1, settings);
                            Compile(context, settings,
                                newSettings: false, createPdf: false, reference1: reference1, reference2);
                        }));
                foreach (var sourceCodeFile in settings.SourceCodeFiles2.Select(_ => _.FullPath(context)))
                    context.Watchers.Add(
                        StartWatcher(sourceCodeFile, () => {
                            reference2 = CompileAssembly(context, settings.SourceCodeFiles2, reference1, settings);
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
                WriteError(settingsChoice.Value2, context, new Option<WatcherSettings>());
            }
        }

        private static Option<PortableExecutableReference> CompileAssembly(Context context, string[] sourceCodeFiles, Option<PortableExecutableReference> reference1,
            Option<WatcherSettings> watcherSettings)
        {
            try
            {
                var compilation = CSharpCompilation.Create(
                    Guid.NewGuid().ToString("N"),
                    sourceCodeFiles.Select(_ => _.FullPath(context)).Select(
                        codeFile => SyntaxFactory.ParseSyntaxTree(File.ReadAllText(codeFile), path: codeFile)),
                    options: new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary),
                    references: context.References.Concat(reference1.Match(_ => new[] {_}, () => new PortableExecutableReference[] { }))
                );
                byte[] bytes;
                using (var stream = new MemoryStream())
                {
                    var emitResult = compilation.Emit(stream);
                    if (!emitResult.Success)
                    {
                        WriteError(GetErrorText(emitResult), context, watcherSettings);
                        return new Option<PortableExecutableReference>();
                    }
                    bytes = stream.ToArray();
                }
                LoadAssembly(bytes);
                return AssemblyMetadata.CreateFromImage(bytes).GetReference();
            }
            catch (Exception e)
            {
                ProcessException(context, e, watcherSettings);
                return new Option<PortableExecutableReference>();
            }
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
                        WriteError(GetErrorText(emitResult), context, settings);
                        return;
                    }
                    bytes = stream.ToArray();
                }
                var assembly = LoadAssembly(bytes);
                var typeName = Path.GetFileNameWithoutExtension(settings.SourceCodeFile);
                var type = assembly.GetTypes().Single(_ => _.Name == typeName);
                var method = type.GetMethod("AddSection");
                var parameterType = method.GetParameters()[1].ParameterType;
                var dataPath = Path.Combine(GetDataSourceDirectory(context, settings),
                    $"{parameterType.FullName}_DataSource.json");
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
                    Process.Start("cmd", $"/c start {fileName}");
                }
                else
                    document.SavePng(
                        pageNumber: settings.PageNumber,
                        path: GetOutputPath(context),
                        resolution: settings.Resolution, watcherSettings: settings, context: context);
                stopwatch.Stop();
                Console.WriteLine($"Done. {DateTime.Now:HH:mm:ss} {stopwatch.ElapsedMilliseconds}ms");
            }
            catch (Exception e)
            {
                ProcessException(context, e, settings);
            }
        }

        private static string GetDataSourceDirectory(Context context, WatcherSettings settings)
        {
            return settings.DataSourceDirectory == null 
                ? Path.GetDirectoryName(GetOutputPath(context)) 
                : settings.DataSourceDirectory.FullPath(context);
        }

        private static void SavePng(this Document document, int pageNumber, string path, double resolution,
            Option<WatcherSettings> watcherSettings, Context context)
        {
            var result = document.SavePng(pageNumber, path, resolution);
            var startExternalProcess = watcherSettings.HasValue 
                ? watcherSettings.Value.StartExternalProcess 
                : context.StartExternalProcess;
            if (startExternalProcess)
                StartProcess(result);
        }

        private static void StartProcess(string fileName)
        {
            Process.Start(new ProcessStartInfo("cmd", $"/c start {fileName}") {CreateNoWindow = true});
        }

        private static void ProcessException(Context context, Exception e, Option<WatcherSettings> watcherSettings)
        {
            if (e is TargetInvocationException exception)
                if (exception.InnerException != null)
                {
                    WriteError(exception.InnerException.Message, context, watcherSettings, e);
                    return;
                }
            WriteError(e.Message, context, watcherSettings, e);
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

        private static readonly SemaphoreSlim semaphore = new SemaphoreSlim(1, 1);
        
        private static FileSystemWatcher StartWatcher(string path, Action action)
        {
            var watcher = new FileSystemWatcher
            {
                Path = Path.GetDirectoryName(path),
                NotifyFilter = CombineAllNotifyFilters()
            };
            string oldText = null;

            async Task Handle(FileSystemEventArgs e)
            {
                if (string.Compare(e.FullPath, path, StringComparison.InvariantCultureIgnoreCase) != 0)
                    return;

                var fileInfo = new FileInfo(path);
                if (!fileInfo.Exists || fileInfo.Length == 0)
                    return;

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

            watcher.Changed += async (_, e) => await Handle(e);
            watcher.Renamed += async (_, e) => await Handle(e);
            watcher.Created += async (_, e) => await Handle(e);
            watcher.Deleted += async (_, e) => await Handle(e);
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

        private static void WriteError(string errorText, Context context, Option<WatcherSettings> watcherSettings, Option<Exception> exception = default)
        {
            var text = $"ERROR! {errorText}";
            Console.WriteLine(text);
            if (exception.HasValue)
                Console.WriteLine(exception.Value);
            var document = new Document();
            var settings = new PageSettings();
            settings.LeftMargin = settings.TopMargin = settings.RightMargin = settings.BottomMargin = Util.Cm(0.5);
            document.Add(new Section(settings).Add(new Paragraph()
                .Add(text, new Font(DefaultFontFamilies.Roboto, 9.5, XFontStyle.Regular, new XPdfFontOptions(PdfFontEncoding.Unicode)))));
            document.SavePng(0, GetOutputPath(context), 120, watcherSettings, context);
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
        public bool StartExternalProcess { get; }
        public readonly List<FileSystemWatcher> Watchers = new();

        public Context(PortableExecutableReference[] references, string settingsPath,
            Func<ParameterInfo, object> parameterFunc, string outputPath, bool startExternalProcess)
        {
            References = references;
            SettingsPath = settingsPath;
            ParameterFunc = parameterFunc;
            OutputPath = outputPath;
            StartExternalProcess = startExternalProcess;
        }
    }
}