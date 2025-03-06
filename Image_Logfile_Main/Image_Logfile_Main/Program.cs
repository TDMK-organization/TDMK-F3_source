using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using VHX;
namespace OK2SHIP
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
            myVar._frmMain = new FrmMain();
            Application.Run(myVar._frmMain);
            //Application.Run(new TestFunc());
            
        }
    }
}
