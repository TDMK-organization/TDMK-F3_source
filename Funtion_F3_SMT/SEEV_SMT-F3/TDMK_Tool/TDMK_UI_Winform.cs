using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SEEV_SMT_F3.TDMK_Tool
{
    public static class TDMK_UI_Winform
    {
        
        // Open form large
        public static void OpenFormLarge(Form form)
        {
            form.WindowState = FormWindowState.Maximized;
            //form.Bounds = Screen.PrimaryScreen.Bounds;
        }
        // Open and close form
        public static void OpenAndCloseForm(Form currentForm, Form nextForm)
        {
            currentForm.Hide();
            nextForm.Show();
        }
    }
}
