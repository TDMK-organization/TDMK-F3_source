
namespace Main_Interface
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblFAI = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblFAI
            // 
            this.lblFAI.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblFAI.Location = new System.Drawing.Point(92, 51);
            this.lblFAI.Name = "lblFAI";
            this.lblFAI.Size = new System.Drawing.Size(161, 115);
            this.lblFAI.TabIndex = 0;
            this.lblFAI.Text = "FAI";
            this.lblFAI.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblFAI.Click += new System.EventHandler(this.lblFAI_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.lblFAI);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lblFAI;
    }
}

