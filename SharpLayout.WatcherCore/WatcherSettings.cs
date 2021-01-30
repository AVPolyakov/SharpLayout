using System;

namespace SharpLayout.WatcherCore
{
    public class WatcherSettings
    {
        public string SourceCodeFile { get; }
        public string[] SourceCodeFiles1 { get; }
        public string[] SourceCodeFiles2 { get; }
        public int PageNumber { get; }
        public double Resolution { get; }
        public Func<Document> DocumentFunc { get; }
        public bool StartExternalProcess { get; set; }
        public string DataSourceDirectory { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourceCodeFile"></param>
        /// <param name="sourceCodeFiles2">Source code files of level 2.</param>
        /// <param name="sourceCodeFiles1">Source code files of level 1.</param>
        /// <param name="pageNumber"></param>
        /// <param name="resolution"></param>
        /// <param name="documentFunc"></param>
        public WatcherSettings(string sourceCodeFile, string[] sourceCodeFiles2, string[] sourceCodeFiles1, int pageNumber,
            double resolution, Func<Document> documentFunc)
        {
            SourceCodeFile = sourceCodeFile;
            SourceCodeFiles1 = sourceCodeFiles1;
            SourceCodeFiles2 = sourceCodeFiles2;
            PageNumber = pageNumber;
            Resolution = resolution;
            DocumentFunc = documentFunc;
        }
    }
}