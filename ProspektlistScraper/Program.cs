using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProspektlistScraper
{
    public class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        //static async Task Main()
        static void Main()
        {
            Application.SetHighDpiMode(HighDpiMode.SystemAware);

            //Comment out this line. Now the color of the progress bar will change as expected, but the style of your controls will also change a little.
            //Application.EnableVisualStyles();

            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
