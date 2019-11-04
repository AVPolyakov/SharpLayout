using System.Runtime.CompilerServices;
using SharpLayout;
using SharpLayout.WatcherCore;

namespace Watcher
{
    public class WatcherSettingsProvider
    {
        private static DevSettings GetDevSettings()
        {
            return new DevSettings {
                SourceCodeFile = SourceCodeFiles.PaymentOrder,
                PageNumber = 1,
                Resolution = 120,
                //CellsAreHighlighted = true,
                //ParagraphsAreHighlighted = true
            };
        }

        public static WatcherSettings GetSettings(string settingsDirectory)
        {
            var devSettings = GetDevSettings();
            return new WatcherSettings(
                sourceCodeFile: $@"..\Examples\{devSettings.SourceCodeFile}.cs",
                sourceCodeFiles: new[] {
                    @"..\Examples\Styles.cs",
                },
                pageNumber: devSettings.PageNumber - 1,
                resolution: devSettings.Resolution,
                documentFunc: () => new Document {
                    CellsAreHighlighted = devSettings.CellsAreHighlighted,
                    R1C1AreVisible = devSettings.CellsAreHighlighted,
                    ParagraphsAreHighlighted = devSettings.ParagraphsAreHighlighted
                });
        }

        public static string FilePath => GetFilePath();
        private static string GetFilePath([CallerFilePath] string filePath = "") => filePath;
    }

    public enum SourceCodeFiles
    {
        PaymentOrder,
        ContractDealPassport,
        LoanAgreementDealPassport,
    }

    public class DevSettings
    {
        public SourceCodeFiles SourceCodeFile { get; set; }
        public int PageNumber { get; set; }
        public int Resolution  { get; set; } = 120;
        public bool CellsAreHighlighted  { get; set; }
        public bool ParagraphsAreHighlighted  { get; set; }
    }
}