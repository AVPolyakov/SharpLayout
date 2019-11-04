using System;
using System.Linq;
using OtherAssembly1;
using OtherAssembly2;

namespace Watcher
{
    class Program
    {
        public static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            SharpLayout.WatcherCore.Watcher.Start(
                settingsPath: WatcherSettingsProvider.FilePath,
                assemblies: new[] {
                    typeof(IA),
                    typeof(A),
                }.Select(_ => _.Assembly).ToArray(),
                parameterFunc: info => {
                    if (info.ParameterType == typeof(IA))
                        return typeof(A);
                    throw new Exception();
                });
        }
    }
}
