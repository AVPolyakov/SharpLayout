using System.Drawing;
using System.Windows.Forms;

namespace LiveViewer
{
    public class Settings
    {
        public Point? Location { get; set; }
        public Size? Size { get; set; }
        public FormWindowState WindowState { get; set; }
    }
}