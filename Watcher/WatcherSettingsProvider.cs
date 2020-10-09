using System.IO;
using System.Runtime.CompilerServices;
using SharpLayout;
using SharpLayout.WatcherCore;

namespace Watcher
{
    public class WatcherSettingsProvider
    {
        private static DevSettings GetDevSettings()
        {
            return new DevSettings
            {
                SourceCodeFile = SourceCodeFiles.PaymentOrder,
                PageNumber = 1,
                //StartExternalProcess = true,
                Resolution = 120,
                //R1C1AreVisible = true,   
                //CellsAreHighlighted = true,
                //ParagraphsAreHighlighted = true 
            };
        }

        public static WatcherSettings GetSettings(string settingsDirectory)
        {
            var devSettings = GetDevSettings();
            return new WatcherSettings(
                sourceCodeFile: $@"..\Examples\{devSettings.SourceCodeFile}.cs",
                sourceCodeFiles2: new string[]
                {
                },
                sourceCodeFiles1: new[]
                {
                    @"..\Examples\DataModels\SqlServer.generated.cs",
                    @"..\Examples\Styles.cs",
                    @"..\Examples\Indexer.cs",
                    @"..\Examples\BandHelper.cs",
                },
                pageNumber: devSettings.PageNumber - 1,
                resolution: devSettings.Resolution,
                documentFunc: () => new Document
                {
                    R1C1AreVisible = devSettings.R1C1AreVisible,
                    CellsAreHighlighted = devSettings.CellsAreHighlighted,
                    ParagraphsAreHighlighted = devSettings.ParagraphsAreHighlighted
                },
                queryFiles: GetQueryFiles(settingsDirectory, devSettings))
            {
                StartExternalProcess = devSettings.StartExternalProcess
            };
        }

        private static string[] GetQueryFiles(string settingsDirectory, DevSettings devSettings)
        {
            var relativePath = $@"..\Examples\{devSettings.SourceCodeFile}Query.cs";
            var path = Path.Combine(settingsDirectory, relativePath);
            if (File.Exists(path))
                return new[] {relativePath};
            return new string[] { };
        }

        public static string FilePath => GetFilePath();
        private static string GetFilePath([CallerFilePath] string filePath = "") => filePath;
    }

    public enum SourceCodeFiles
    {
        PaymentOrder,
        ContractDealPassport,
        LoanAgreementDealPassport,
        Svo,
        LegalEntityCreation
    }

    public class DevSettings
    {
        public SourceCodeFiles SourceCodeFile { get; set; }
        public int PageNumber { get; set; }
        public int Resolution  { get; set; } = 120;
        public bool CellsAreHighlighted  { get; set; }
        public bool R1C1AreVisible  { get; set; }
        public bool ParagraphsAreHighlighted  { get; set; }
        public bool StartExternalProcess { get; set; }
    }
}