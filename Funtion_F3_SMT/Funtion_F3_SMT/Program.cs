using OK2SHIP_SMT.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OK2SHIP_SMT.Views;

namespace Funtion_F3_SMT
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            if (args.Length > 0)
            {
                string receivedMessage = args[0];
                // Xử lý logic với receivedMessage ở đây
                Application.Run(new BendingWindow());
            }
            else
            {
                Application.Run(new MainScreen());
            }
        }
    }
}