using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.UserControls
{
    public partial class EditImageView : UserControl
    {
        public EditImageView()
        {
            InitializeComponent();
        }
        public Image image { get; set; }
        public bool save_status { get; set; } = false;
        public EditImageView(Image image, Action<object, EventArgs> btn_close_Click)
        {
            InitializeComponent();
            this.image = image;
            btnSave.Click += new System.EventHandler(btn_close_Click);
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void EditImageView_Load(object sender, EventArgs e)
        {
            this.mainForm1.ReplaceInitialImage(image, false, true);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            this.image = this.mainForm1.ExportToImage();
            this.save_status = true;

        }

        private void mainForm1_Load(object sender, EventArgs e)
        {

        }
    }
}
