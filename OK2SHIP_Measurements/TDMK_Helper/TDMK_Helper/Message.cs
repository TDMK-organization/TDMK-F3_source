using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TDMK_Helper
{
    public class TDMK_Message
    {
        public static void MessageBoxTDMK_Info(string message)
        {
            MessageBox.Show(message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        public static void MessageBoxTDMK_Error(string message)
        {
            MessageBox.Show(message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        public static void MessageBoxTDMK_Warning(string message)
        {
            MessageBox.Show(message, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        public static DialogResult MessageBoxDialog(string message)
        {
            DialogResult result = MessageBox.Show(message, "Cảnh báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            return result;
        }
        public static bool MessageWarningLogin(string message)
        {
            using (FrmTDMKMessage f =
                new FrmTDMKMessage(
                    message,
                    "Cảnh báo",
                    allowAdminLogin: true))
            {
                f.ShowDialog();
                return f.AdminConfirmed;
            }
        }
    }
}
