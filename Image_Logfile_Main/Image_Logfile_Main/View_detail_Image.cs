using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace OK2SHIP
{
    public partial class View_detail_Image : Form
    {
        public View_detail_Image()
        {
            InitializeComponent();
        }


        public byte[] data_image = null;
        public string region = "";
        public string pcs = "";

        public string choose_sheet = "";

        public byte[] data
        {
            get { return data_image; }
            set { data_image = value; }
        }

        public string region_
        {
            get { return region; }
            set { region = value; }
        }

        public string pcs_
        {
            get { return pcs; }
            set { pcs = value; }
        }

        public string choose_sheet_
        {
            get { return choose_sheet; }
            set { choose_sheet = value; }
        }
        private void View_detail_Image_Load(object sender, EventArgs e)
        {
            switch (choose_sheet)
            {
                case "Stackup":
                    if (data_image != null && region != "" && pcs != "")
                    {
                        using (MemoryStream ms = new MemoryStream(data_image))
                        {
                            image_detail.Image = Image.FromStream(ms);
                        }

                        lblImage_Graph.Text = "Image Details ---> Region: " + region + " / " + "Pcs_No: " + pcs;
                    }
                    break;

                case "Impedance":
                    if (data_image != null && region != "" && pcs != "")
                    {
                        using (MemoryStream ms = new MemoryStream(data_image))
                        {
                            image_detail.Image = Image.FromStream(ms);
                        }

                        lblImage_Graph.Text = "Graph Details ---> Impdedance: " + region + " / " + "Pcs_No: " + pcs;
                    }
                    break;

                case "Tracewidth":
                    if (data_image != null && region != "" && pcs != "")
                    {
                        using (MemoryStream ms = new MemoryStream(data_image))
                        {
                            image_detail.Image = Image.FromStream(ms);
                        }

                        lblImage_Graph.Text = "Image Details --->" + region + " / "  + pcs;
                    }
                    break;


            }
           
           
        }

    }
}
