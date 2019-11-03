using System;
using System.IO;
using System.Reflection;
using SharpLayout;
using Watcher.Contracts;

namespace Starter
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
            var otherAssembly1 = Assembly.LoadFrom(Path.Combine(settingsDirectory, @"bin\Debug\netcoreapp3.0\OtherAssembly1.dll"));
            var otherAssembly2 = Assembly.LoadFrom(Path.Combine(settingsDirectory, @"bin\Debug\netcoreapp3.0\OtherAssembly2.dll"));
            var a = Activator.CreateInstance(otherAssembly2.GetType("OtherAssembly2.A"));
            return new WatcherSettings(
                sourceCodeFile: $@"..\Examples\{devSettings.SourceCodeFile}.cs",
                sourceCodeFiles: new[] {
                    @"..\Examples\Styles.cs",
                },
                outputPath: @"bin\Debug\netcoreapp3.0\Temp.png",
                pageNumber: devSettings.PageNumber - 1,
                resolution: devSettings.Resolution,
                documentFunc: () => {
                    var document = new Document();
                    document.CellsAreHighlighted = document.R1C1AreVisible = devSettings.CellsAreHighlighted;
                    document.ParagraphsAreHighlighted = devSettings.ParagraphsAreHighlighted;
                    return document;
                },
                assemblies: new[] {
                    otherAssembly1,
                    otherAssembly2,
                },
                parameterFunc: info => {
                    if (info.ParameterType == otherAssembly1.GetType("OtherAssembly1.IA"))
                        return a;
                    throw new Exception();
                });
        }
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