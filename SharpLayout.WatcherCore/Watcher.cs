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
        private static readonly string[] referenceNames = {
            "Microsoft.CSharp.dll",
            "Microsoft.VisualBasic.dll",
            "Microsoft.Win32.Primitives.dll",
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
            "System.Globalization.Calendars.dll",
            "System.Globalization.dll",
            "System.Globalization.Extensions.dll",
            "System.IO.Compression.Brotli.dll",
            "System.IO.Compression.dll",
            "System.IO.Compression.FileSystem.dll",
            "System.IO.Compression.ZipFile.dll",
            "System.IO.dll",
            "System.IO.FileSystem.dll",
            "System.IO.FileSystem.DriveInfo.dll",
            "System.IO.FileSystem.Primitives.dll",
            "System.IO.FileSystem.Watcher.dll",
            "System.IO.IsolatedStorage.dll",
            "System.IO.MemoryMappedFiles.dll",
            "System.IO.Pipes.dll",
            "System.IO.UnmanagedMemoryStream.dll",
            "System.Linq.dll",
            "System.Linq.Expressions.dll",
            "System.Linq.Parallel.dll",
            "System.Linq.Queryable.dll",
            "System.Memory.dll",
            "System.Net.dll",
            "System.Net.Http.dll",
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
            "System.Runtime.CompilerServices.VisualC.dll",
            "System.Runtime.dll",
            "System.Runtime.Extensions.dll",
            "System.Runtime.Handles.dll",
            "System.Runtime.InteropServices.dll",
            "System.Runtime.InteropServices.RuntimeInformation.dll",
            "System.Runtime.InteropServices.WindowsRuntime.dll",
            "System.Runtime.Loader.dll",
            "System.Runtime.Numerics.dll",
            "System.Runtime.Serialization.dll",
            "System.Runtime.Serialization.Formatters.dll",
            "System.Runtime.Serialization.Json.dll",
            "System.Runtime.Serialization.Primitives.dll",
            "System.Runtime.Serialization.Xml.dll",
            "System.Security.Claims.dll",
            "System.Security.Cryptography.Algorithms.dll",
            "System.Security.Cryptography.Csp.dll",
            "System.Security.Cryptography.Encoding.dll",
            "System.Security.Cryptography.Primitives.dll",
            "System.Security.Cryptography.X509Certificates.dll",
            "System.Security.dll",
            "System.Security.Principal.dll",
            "System.Security.SecureString.dll",
            "System.ServiceModel.Web.dll",
            "System.ServiceProcess.dll",
            "System.Text.Encoding.dll",
            "System.Text.Encoding.Extensions.dll",
            "System.Text.RegularExpressions.dll",
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
            string outputPath)
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
                var reference1 = CompileReference(context, settings.SourceCodeFiles1);
                var reference2 = CompileReference(context, settings.SourceCodeFiles2, reference1);
                var dataTypeReferenceTuple = CompileReferenceTuple(context, new []{settings.GetDataTypePath()});
                var data = GetData(context, settings, reference1, reference2, dataTypeReferenceTuple);
                Compile(context, settings, createPdf: createPdf, reference1, reference2, dataTypeReferenceTuple,
                    data);
                foreach (var sourceCodeFile in settings.SourceCodeFiles1.Select(_ => _.FullPath(context)))
                    context.Watchers.Add(
                        StartWatcher(sourceCodeFile, () => {
                            reference1 = CompileReference(context, settings.SourceCodeFiles1);
                            reference2 = CompileReference(context, settings.SourceCodeFiles2, reference1);
                            dataTypeReferenceTuple = CompileReferenceTuple(context, new []{settings.GetDataTypePath()});
                            data = GetData(context, settings, reference1, reference2, dataTypeReferenceTuple);
                            Compile(context, settings, createPdf: false, reference1: reference1, reference2: reference2,
                                dataTypeReferenceTuple: dataTypeReferenceTuple, data: data);
                        }));
                foreach (var sourceCodeFile in settings.SourceCodeFiles2.Select(_ => _.FullPath(context)))
                    context.Watchers.Add(
                        StartWatcher(sourceCodeFile, () => {
                            reference2 = CompileReference(context, settings.SourceCodeFiles2, reference1);
                            dataTypeReferenceTuple = CompileReferenceTuple(context, new []{settings.GetDataTypePath()});
                            data = GetData(context, settings, reference1, reference2, dataTypeReferenceTuple);
                            Compile(context, settings, createPdf: false, reference1: reference1, reference2: reference2,
                                dataTypeReferenceTuple: dataTypeReferenceTuple, data: data);
                        }));
                context.Watchers.Add(
                    StartWatcher(settings.GetDataTypePath().FullPath(context), () => {
                        dataTypeReferenceTuple = CompileReferenceTuple(context, new []{settings.GetDataTypePath()});
                        data = GetData(context, settings, reference1, reference2, dataTypeReferenceTuple);
                        Compile(context, settings, createPdf: false, reference1: reference1, reference2: reference2,
                            dataTypeReferenceTuple: dataTypeReferenceTuple, data: data);
                    }));
                context.Watchers.Add(
                    StartWatcher(settings.GetDataProviderPath(context).FullPath(context), () => {
                        data = GetData(context, settings, reference1, reference2, dataTypeReferenceTuple);
                        Compile(context, settings, createPdf: false, reference1: reference1, reference2: reference2,
                            dataTypeReferenceTuple: dataTypeReferenceTuple, data: data);
                    }));
                context.Watchers.Add(
                    StartWatcher(settings.GetQueryPath().FullPath(context), () => {
                        data = GetData(context, settings, reference1, reference2, dataTypeReferenceTuple);
                        Compile(context, settings, createPdf: false, reference1: reference1, reference2: reference2,
                            dataTypeReferenceTuple: dataTypeReferenceTuple, data: data);
                    }));
                context.Watchers.Add(
                    StartWatcher(settings.SourceCodeFile.FullPath(context), () => {
                        Compile(context, settings, createPdf: false, reference1: reference1, reference2: reference2,
                            dataTypeReferenceTuple: dataTypeReferenceTuple, data: data);
                    }));
                context.Watchers.Add(StartWatcher(context.GetDataPath(settings.GetDataType(dataTypeReferenceTuple)),
                    () => Compile(context, settings, createPdf: false, reference1: reference1,
                        reference2: reference2, dataTypeReferenceTuple: dataTypeReferenceTuple, data: data)));
            }
            else
            {
                WriteError(settingsChoice.Value2, context);
            }
        }

        private static Option<object> GetData(Context context, WatcherSettings settings, 
            Option<PortableExecutableReference> reference1, Option<PortableExecutableReference> reference2, 
            Option<ReferenceTuple> dataTypeReferenceTuple)
        {
            var dataProviderPath = settings.GetDataProviderPath(context);
            if (!File.Exists(dataProviderPath))
            {
                var dataType = settings.GetDataType(dataTypeReferenceTuple);
                return JsonConvert.DeserializeObject(
                    File.ReadAllText(context.GetDataPath(dataType)), dataType);
            }
            try
            {
                var sourceCodeFiles = new List<string>();
                var queryPath = settings.GetQueryPath();
                if (File.Exists(queryPath.FullPath(context)))
                    sourceCodeFiles.Add(queryPath);
                sourceCodeFiles.Add(dataProviderPath);
                return CompileAssembly(context, sourceCodeFiles, reference1, 
                        reference2, dataTypeReferenceTuple.Select(_ => _.Reference))
                    .Select(tuple => {
                        var sourceCodeFileName = Path.GetFileNameWithoutExtension(settings.SourceCodeFile);
                        var dataProviderType = tuple.Assembly
                            .GetTypes().Single(t => t.Name == $"{sourceCodeFileName}{DataProvider}");
                        var dataProvider = (IDataProvider) Activator.CreateInstance(dataProviderType);
                        return dataProvider.Create(() => {
                            var dataType = settings.GetDataType(dataTypeReferenceTuple);
                            return JsonConvert.DeserializeObject(
                                File.ReadAllText(context.GetDataPath(dataType)), dataType);
                        });
                    });
            }
            catch (Exception e)
            {
                ProcessException(context, e);
                return new Option<object>();
            }
        }

        private static string GetDataTypePath(this WatcherSettings settings)
        {
            var directoryName = Path.GetDirectoryName(settings.SourceCodeFile);
            var fileName = Path.GetFileNameWithoutExtension(settings.SourceCodeFile);
            var extension = Path.GetExtension(settings.SourceCodeFile);
            return Path.Combine(directoryName, $"{fileName}{Data}{extension}");
        }

        private static string GetQueryPath(this WatcherSettings settings)
        {
            var directoryName = Path.GetDirectoryName(settings.SourceCodeFile);
            var fileName = Path.GetFileNameWithoutExtension(settings.SourceCodeFile);
            var extension = Path.GetExtension(settings.SourceCodeFile);
            return Path.Combine(directoryName, $"{fileName}Query{extension}");
        }

        private static string GetDataProviderPath(this WatcherSettings settings, Context context)
        {
            var directoryName = Path.GetDirectoryName(context.SettingsPath);
            var fileName = Path.GetFileNameWithoutExtension(settings.SourceCodeFile);
            var extension = Path.GetExtension(settings.SourceCodeFile);
            return Path.Combine(directoryName, $"{fileName}{DataProvider}{extension}");
        }

        private const string DataProvider = "DataProvider";
        private const string Data = "Data";
        
        private static Option<PortableExecutableReference> CompileReference(Context context, string[] sourceCodeFiles,
            params Option<PortableExecutableReference>[] references)
        {
            return CompileReferenceTuple(context, sourceCodeFiles, references)
                .Select(_ => _.Reference);
        }

        private static Option<ReferenceTuple> CompileReferenceTuple(Context context, string[] sourceCodeFiles, 
            params Option<PortableExecutableReference>[] references)
        {
            return CompileAssembly(context, sourceCodeFiles, references)
                .Select(tuple => new ReferenceTuple(
                    reference: AssemblyMetadata.CreateFromImage(tuple.Bytes).GetReference(),
                    assemblyTuple: tuple));
        }

        private class ReferenceTuple
        {
            public PortableExecutableReference Reference { get; }
            public AssemblyTuple AssemblyTuple { get; }

            public ReferenceTuple(PortableExecutableReference reference, AssemblyTuple assemblyTuple)
            {
                Reference = reference;
                AssemblyTuple = assemblyTuple;
            }
        }

        private static Option<AssemblyTuple> CompileAssembly(Context context, IEnumerable<string> sourceCodeFiles,
            params Option<PortableExecutableReference>[] references)
        {
            if (references.Any(_ => !_.HasValue)) 
                return new Option<AssemblyTuple>();
            try
            {
                var compilation = CSharpCompilation.Create(
                    Guid.NewGuid().ToString("N"),
                    sourceCodeFiles.Select(_ => _.FullPath(context)).Select(
                        codeFile => SyntaxFactory.ParseSyntaxTree(File.ReadAllText(codeFile), path: codeFile)),
                    options: new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary),
                    references: context.References.Concat(references.Select(_ => _.Value))
                );
                byte[] bytes;
                using (var stream = new MemoryStream())
                {
                    var emitResult = compilation.Emit(stream);
                    if (!emitResult.Success)
                    {
                        WriteError(GetErrorText(emitResult), context);
                        return new Option<AssemblyTuple>();
                    }
                    bytes = stream.ToArray();
                }
                return new AssemblyTuple(LoadAssembly(bytes), bytes);
            }
            catch (Exception e)
            {
                ProcessException(context, e);
                return new Option<AssemblyTuple>();
            }
        }
        
        private class AssemblyTuple
        {
            public Assembly Assembly { get; }
            public byte[] Bytes { get; }

            public AssemblyTuple(Assembly assembly, byte[] bytes)
            {
                Assembly = assembly;
                Bytes = bytes;
            }
        }

        private static void Compile(Context context, WatcherSettings settings, bool createPdf,
            Option<PortableExecutableReference> reference1, Option<PortableExecutableReference> reference2,
            Option<ReferenceTuple> dataTypeReferenceTuple, Option<object> data)
        {
            if (!reference1.HasValue) return;
            if (!reference2.HasValue) return;
            if (!dataTypeReferenceTuple.HasValue) return;
            if (!data.HasValue) return;
            try
            {
                var stopwatch = Stopwatch.StartNew();
                var codeFile = FullPath(settings.SourceCodeFile, context);
                var compilation = CSharpCompilation.Create(
                    Guid.NewGuid().ToString("N"),
                    new[] {SyntaxFactory.ParseSyntaxTree(File.ReadAllText(codeFile), path: codeFile)},
                    options: new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary),
                    references: context.References.Concat(new[] {reference1.Value, reference2.Value, 
                        dataTypeReferenceTuple.Select(_ => _.Reference).Value})
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
                var document = settings.DocumentFunc();
                var parameters = method.GetParameters().Skip(2).Select(info => context.ParameterFunc(info));
                method.Invoke(null, new[] {document, data.Value}.Concat(parameters).ToArray());
                if (createPdf)
                {
                    var fileName = Path.Combine(
                        Path.GetDirectoryName(GetOutputPath(context)),
                        $"Temp_{Guid.NewGuid():N}.pdf");
                    document.SavePdf(fileName);
                    Process.Start(new ProcessStartInfo("cmd", $"/c start {fileName}"));
                }
                else
                {
                    const int numberOfRetries = 50;
                    const int delayOnRetry = 20;
                    for (var i = 1; i <= numberOfRetries; i++)
                        try
                        {
                            document.SavePng(
                                pageNumber: settings.PageNumber,
                                path: GetOutputPath(context),
                                resolution: settings.Resolution);
                            break;
                        }
                        catch (IOException)
                        {
                            Thread.Sleep(delayOnRetry);
                        }
                }
                stopwatch.Stop();
                Console.WriteLine($"Done. {DateTime.Now:HH:mm:ss} {stopwatch.ElapsedMilliseconds}ms");
            }
            catch (Exception e)
            {
                ProcessException(context, e);
            }
        }

        private static string GetDataPath(this Context context, Type dataType)
        {
            return Path.Combine(Path.GetDirectoryName(GetOutputPath(context)),
                $"{dataType.FullName}_DataSource.json");
        }

        private static Type GetDataType(this WatcherSettings settings, Option<ReferenceTuple> referenceTuple3)
        {
            var sourceCodeFileName = Path.GetFileNameWithoutExtension(settings.SourceCodeFile);
            var dataType = referenceTuple3.Value.AssemblyTuple.Assembly
                .GetTypes().Single(t => t.Name == $"{sourceCodeFileName}{Data}");
            return dataType;
        }

        private static void ProcessException(Context context, Exception e)
        {
            if (e is TargetInvocationException exception)
                if (exception.InnerException != null)
                {
                    WriteError(exception.InnerException.Message, context);
                    return;
                }
            WriteError(e.Message, context);
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

                if (new FileInfo(path).Length == 0)
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