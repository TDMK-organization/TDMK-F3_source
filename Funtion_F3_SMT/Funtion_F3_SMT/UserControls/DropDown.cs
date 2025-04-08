using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Funtion_F3_SMT.UserControls
{
    public class DropDown : ToolStripComboBox
    {
        private string PlaceHolder { get; set; }
        #region Constructor

        public DropDown()
        {
            this.PlaceHolder = "Chosse Your Process";
            Constructor();
            this.Text = PlaceHolder;

        }

        /// <summary>
        /// Constructor with placeholder
        /// </summary>
        /// <param name="placeHolder">string of placeholder</param>
        public DropDown(string placeHolder)
        {
            this.PlaceHolder = placeHolder;
            Constructor();

        }
        /// <summary>
        /// Constructor all
        /// </summary>
        private void Constructor()
        {
            this.Text = PlaceHolder;
            this.SelectedIndexChanged += DropDown_SelectedIndexChanged;
            this.DropDown += DropDown_Click;
            this.ForeColor = Color.Gray;
            this.DropDownClosed += DropDown_DropDownClosed;
        }

        #endregion
        /// <summary>
        /// When dropdown closed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DropDown_DropDownClosed(object sender, EventArgs e)
        {
            if(this.SelectedItem == null)
            {
                this.ForeColor = Color.Gray;
                this.Text = PlaceHolder;
            }
            else
            {
                this.ForeColor = Color.Black;
            }
        }
        /// <summary>
        /// Event when dropdown clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DropDown_Click(object sender, EventArgs e)
        {
            this.ForeColor = Color.Black;
        }

        /// <summary>
        /// Event when dropdown selected index changed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DropDown_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.SelectedItem != null)
            {
                this.ForeColor = Color.Black;
            }
        }

    }
}