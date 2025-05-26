using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

namespace Main_Interface
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void lblFAI_Click(object sender, EventArgs e)
        {
            string process_name = @"D:\Customer Projects\SEEV\SMT Project\Source Code\TestAreas\OK2SHIP_Measurements\OK2SHIP_Measurements(Ver_02).exe";
           // process_name = @"D:\Customer Projects\SEEV\SMT Project\Source Code\TestAreas\Bending Items\Bending_Items(Ver01).exe";
            Process.Start(process_name);
            
        }
    }
}
