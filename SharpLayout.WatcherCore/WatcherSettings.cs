using System;

namespace SharpLayout.WatcherCore
{
    public class WatcherSettings
    {
        public string SourceCodeFile { get; }
        public string[] SourceCodeFiles { get; }
        public int PageNumber { get; }
        public int Resolution { get; }
        public Func<Document> DocumentFunc { get; }

        public WatcherSettings(string sourceCodeFile, string[] sourceCodeFiles, int pageNumber,
            int resolution, Func<Document> documentFunc)
        {
            SourceCodeFile = sourceCodeFile;
            SourceCodeFiles = sourceCodeFiles;
            PageNumber = pageNumber;
            Resolution = resolution;
            DocumentFunc = documentFunc;
        }
    }
}