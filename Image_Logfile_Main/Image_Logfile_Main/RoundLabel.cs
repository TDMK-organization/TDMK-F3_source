using System;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace RGO_Main
{
    public class RoundedCornerButton : Button
    {
        protected override void OnPaint(PaintEventArgs pevent)
        {           
            GraphicsPath graphPath = new GraphicsPath();
            graphPath.AddArc(2, 2, 20, 20, 180, 90);
            graphPath.AddArc(Width - 23, 2, 20, 20, 270, 90);
            graphPath.AddArc(Width - 23, Height - 23, 20, 20, 0, 90);
            graphPath.AddArc(2, Height - 23, 20, 20, 90, 90);
            graphPath.CloseAllFigures();
            this.Region = new Region(graphPath);
            base.OnPaint(pevent);
        }
    }
    public class RoundedCornerLabel : Label
    {
        protected override void OnPaint(PaintEventArgs pevent)
        {
            GraphicsPath graphPath = new GraphicsPath();
            graphPath.AddArc(0, 0, 20, 20, 180, 90);
            graphPath.AddArc(Width - 21, 0, 20, 20, 270, 90);
            graphPath.AddArc(Width - 21, Height - 21, 20, 20, 0, 90);
            graphPath.AddArc(0, Height - 21, 20, 20, 90, 90);
            graphPath.CloseAllFigures();
            this.Region = new Region(graphPath);
            this.TextAlign = ContentAlignment.MiddleCenter;
            base.OnPaint(pevent);
        }
    }
    public class RoundedCornerForm : Form
    {
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics graphics = e.Graphics;
            Pen pen = new Pen(Color.Gray, 2);
            Rectangle rectangle = new Rectangle(0, 0, this.Width, this.Height);
            int radius = 10;
            GraphicsPath path = new GraphicsPath();
            path.AddArc(rectangle.X, rectangle.Y, radius, radius, 180, 90);
            path.AddArc(rectangle.X + rectangle.Width - radius, rectangle.Y, radius, radius, 270, 90);
            path.AddArc(rectangle.X + rectangle.Width - radius, rectangle.Y + rectangle.Height - radius, radius, radius, 0, 90);
            path.AddArc(rectangle.X, rectangle.Y + rectangle.Height - radius, radius, radius, 90, 90);
            path.CloseFigure();
            graphics.DrawPath(pen, path);
        }
    }

}
