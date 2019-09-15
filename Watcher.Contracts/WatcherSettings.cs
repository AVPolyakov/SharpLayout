using System;
using System.Reflection;
using SharpLayout;

namespace Watcher.Contracts
{
    public class WatcherSettings
    {
        public string[] SourceCodeFiles { get; }
        public string OutputPath { get; }
        public int PageNumber { get; }
        public int Resolution { get; }
        public Func<Document> DocumentFunc { get; }
        public Assembly[] Assemblies { get; }
        public Func<ParameterInfo, object> ParameterFunc { get; }

        public WatcherSettings(string[] sourceCodeFiles, string outputPath, int pageNumber,
            int resolution, Func<Document> documentFunc, Assembly[] assemblies,
            Func<ParameterInfo, object> parameterFunc)
        {
            SourceCodeFiles = sourceCodeFiles;
            OutputPath = outputPath;
            PageNumber = pageNumber;
            Resolution = resolution;
            DocumentFunc = documentFunc;
            Assemblies = assemblies;
            ParameterFunc = parameterFunc;
        }
    }
}