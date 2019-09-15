using System;
using System.IO;
using System.Reflection;
using Watcher.Contracts;

namespace SharpLayout.Tests
{
    public class WatcherSettingsProvider
    {
        public static WatcherSettings GetSettings(string settingsDirectory)
        {
            var otherAssembly1 = Assembly.LoadFrom(Path.Combine(settingsDirectory, @"bin\Debug\OtherAssembly1.dll"));
            var otherAssembly2 = Assembly.LoadFrom(Path.Combine(settingsDirectory, @"bin\Debug\OtherAssembly2.dll"));
            var a = Activator.CreateInstance(otherAssembly2.GetType("OtherAssembly2.A"));
            return new WatcherSettings(
                sourceCodeFiles: new[]
                {
                    "Styles.cs",
                    "PaymentOrder.cs",
                },
                outputPath: @"bin\Debug\Temp.png",
                pageNumber: 0,
                resolution: 120,
                documentFunc: () =>
                {
                    var document = new Document();
                    document.CellsAreHighlighted = document.R1C1AreVisible = true;
                    //document.ParagraphsAreHighlighted = true;
                    return document;
                },
                assemblies: new[]
                {
                    otherAssembly1,
                    otherAssembly2,
                },
                parameterFunc: info =>
                {
                    if (info.ParameterType == otherAssembly1.GetType("OtherAssembly1.IA"))
                        return a;
                    throw new Exception();
                });
        }
    }
}