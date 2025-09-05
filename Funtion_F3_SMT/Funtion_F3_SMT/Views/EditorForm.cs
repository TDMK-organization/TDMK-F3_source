using Microsoft.Office.Core;
using OK2SHIP_SMT.Services;
using OK2SHIP_SMT.UserControls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

namespace OK2SHIP_SMT.Views
{
    public partial class EditorForm : Form
    {

        public Image _image = null;
        public EditorForm(Image image)
        {

            this.WindowState = FormWindowState.Maximized;
            InitializeComponent();
            _image = image;
            pictureBox1.Image = image;
            if (this._image == null)
            {
                throw new Exception("Image cannot be null");
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            //here we declare mouse event handlers

            pictureBox1.MouseDown += new MouseEventHandler(pictureBox1_MouseDown);

            pictureBox1.MouseMove += new MouseEventHandler(pictureBox1_MouseMove);
            pictureBox1.MouseUp += new MouseEventHandler(pictureBox1_MouseUp);

            pictureBox1.MouseEnter += new EventHandler(pictureBox1_MouseEnter);
            //Controls.Add(pictureBox1);
        }

        //declare some variable for crop coordinates
        int crpX, crpY, rectW, rectH;
        // Declare crop pen for cropping image
        public Pen crpPen = new Pen(Color.White);
        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            base.OnMouseDown(e);

            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                Cursor = Cursors.Cross;
                crpPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dot;
                // set initial x,y co ordinates for crop rectangle
                //this is where we firstly click on image
                crpX = e.X;
                crpY = e.Y;

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_MouseUp(object sender, MouseEventArgs e)
        {
            Cursor = Cursors.Default;
            //Now we will draw the cropped image into pictureBox2
            Bitmap bmp2 = new Bitmap(pictureBox1.Width, pictureBox1.Height);
            pictureBox1.DrawToBitmap(bmp2, pictureBox1.ClientRectangle);

            Bitmap crpImg = new Bitmap(Math.Abs(rectW), Math.Abs(rectH));

            for (int i = 0; i < rectW; i++)
            {
                for (int y = 0; y < rectH; y++)
                {
                    Color pxlclr = bmp2.GetPixel(crpX + i, crpY + y);
                    crpImg.SetPixel(i, y, pxlclr);
                }
            }

            pictureBox1.Image = (Image)crpImg;
            pictureBox1.SizeMode = PictureBoxSizeMode.CenterImage;
        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {
            _image = pictureBox1.Image;
            this.Hide();
        }

        private void toolStripLabel4_Click(object sender, EventArgs e)
        {
            pictureBox1.Image = _image;
            pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
        }
        Form dialog = new Form();
        private void btn_close_Click(object sender, EventArgs e)
        {
            dialog.Hide();
        }
        private void toolStripLabel2_Click(object sender, EventArgs e)
        {
            Image image = pictureBox1.Image;
            EditImageView edit = new EditImageView((Image)image, btn_close_Click);
            dialog = new CommonForm("", edit, null);
            dialog.ShowDialog();
            if (!edit.save_status)
            {
                return;
            }
            pictureBox1.Image = edit.image;
        }

        private void toolStripLabel5_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";
            if (open.ShowDialog() == DialogResult.OK)
            {
                pictureBox1.Image = Image.FromFile(open.FileName);
            }
        }

        private void pictureBox1_MouseEnter(object sender, EventArgs e)
        {
            base.OnMouseEnter(e);
            Cursor = Cursors.Cross;
        }

        private void pictureBox1_MouseMove(object sender, MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                pictureBox1.Refresh();
                //set width and height for crop rectangle.
                rectW = e.X - crpX;
                rectH = e.Y - crpY;
                Graphics g = pictureBox1.CreateGraphics();
                g.DrawRectangle(crpPen, crpX, crpY, rectW, rectH);
                g.Dispose();
            }
        }
        protected override void OnMouseEnter(EventArgs e)
        {
            base.OnMouseEnter(e);
            Cursor = Cursors.Default;
        }

    }
}



