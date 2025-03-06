using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VHX
{
    public class FindControl
    {
        Control c = null;
        Control f = null;
        public FindControl()
        {
        }
        public Control TheForm(string name)
        {
            FormCollection fc = Application.OpenForms;

            //---------------------------------------------------------------------------------
            for (int i = 0; i < fc.Count; i++)
            {

                c = null;
                if (fc[i].Name == name)
                {
                    f = fc[i];
                    break;
                }
            }
            return ((Control)f);
        }

        //---------------------------------------------------------------------------------
        public Control Ctrl(Control f, string name)
        {
            for (int i = 0; i < f.Controls.Count; i++)
            {
                if (f.Controls[i].Name == name)
                {
                    c = f.Controls[i];
                    break;
                }
                if (c == null)
                {
                    if (f.Controls[i].Controls.Count > 0)
                        Ctrl(f.Controls[i], name);
                }
                if (c != null)
                    break;
            }
            return (c);
        }
    }
}
