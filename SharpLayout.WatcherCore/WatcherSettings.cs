using System;

namespace SharpLayout.WatcherCore
{
    public class WatcherSettings
    {
        public string SourceCodeFile { get; }
        public string[] SourceCodeFiles1 { get; }
        public string[] SourceCodeFiles2 { get; }
        public int PageNumber { get; }
        public int Resolution { get; }
        public Func<Document> DocumentFunc { get; }
        public string[] QueryFiles { get; }
        public bool StartExternalProcess { get; set; }

        public WatcherSettings(string sourceCodeFile, string[] sourceCodeFiles2, string[] sourceCodeFiles1, int pageNumber,
            int resolution, Func<Document> documentFunc, string[] queryFiles)
        {
            SourceCodeFile = sourceCodeFile;
            SourceCodeFiles1 = sourceCodeFiles1;
            SourceCodeFiles2 = sourceCodeFiles2;
            PageNumber = pageNumber;
            Resolution = resolution;
            DocumentFunc = documentFunc;
            QueryFiles = queryFiles;
        }
    }
}