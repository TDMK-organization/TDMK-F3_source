using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Funtion_F3_SMT
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Setup_Spec_SMT());
            //Application.Run(new View_Data_3());
            //Application.Run(new  View_Data());
            Application.Run(new  MainScreen());


        }
    }
}
