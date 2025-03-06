using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP
{
    public partial class Image_history : Form
    {
        public Image_history()
        {
            InitializeComponent();
        }
        public byte[] data_image_before = null;
        public byte[] data_image_after = null;

        public byte[] data_image_before_
        {
            get { return data_image_before; }
            set { data_image_before = value; }
        }


        public byte[] data_image_after_
        {
            get { return data_image_after; }
            set { data_image_after = value; }
        }

        private void Image_history_Load(object sender, EventArgs e)
        {
            if (data_image_before != null && data_image_after != null)
            {
                using (MemoryStream ms = new MemoryStream(data_image_before))
                {
                    pic_before.Image = Image.FromStream(ms);
                }

                using (MemoryStream ms = new MemoryStream(data_image_after))
                {
                    pic_after.Image = Image.FromStream(ms);
                }
            }
        }
    }
}
