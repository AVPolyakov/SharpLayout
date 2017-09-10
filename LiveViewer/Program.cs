using System;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualBasic.ApplicationServices;

namespace LiveViewer
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            var args = Environment.GetCommandLineArgs();
            var controller = new SingleInstanceController();
            controller.Run(args);
        }
    }

    /// <summary>
    /// http://www.hanselman.com/blog/TheWeeklySourceCode31SingleInstanceWinFormsAndMicrosoftVisualBasicdll.aspx
    /// https://stackoverflow.com/a/3808173
    /// </summary>
    public class SingleInstanceController : WindowsFormsApplicationBase
    {
        public SingleInstanceController()
        {
            IsSingleInstance = true;

            StartupNextInstance += this_StartupNextInstance;
        }

        void this_StartupNextInstance(object sender, StartupNextInstanceEventArgs e)
        {
            ((MainForm) MainForm).LoadFile(e.CommandLine.ToArray());
        }

        protected override void OnCreateMainForm()
        {
            MainForm = new MainForm();
        }
    }
}
