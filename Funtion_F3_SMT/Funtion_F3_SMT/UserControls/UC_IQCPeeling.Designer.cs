using System.ComponentModel;

namespace OK2SHIP_SMT.UserControls
{
    partial class UC_IQCPeeling
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.gridPanel1 = new AntdUI.GridPanel();
            this.gridPanel2 = new AntdUI.GridPanel();
            this.gridPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // gridPanel1
            // 
            this.gridPanel1.Controls.Add(this.gridPanel2);
            this.gridPanel1.Location = new System.Drawing.Point(0, 0);
            this.gridPanel1.Name = "gridPanel1";
            this.gridPanel1.Size = new System.Drawing.Size(996, 628);
            this.gridPanel1.Span = " ";
            this.gridPanel1.TabIndex = 0;
            this.gridPanel1.Text = "gridPanel1";
            // 
            // gridPanel2
            // 
            this.gridPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gridPanel2.Location = new System.Drawing.Point(142, 90);
            this.gridPanel2.Name = "gridPanel2";
            this.gridPanel2.Size = new System.Drawing.Size(851, 538);
            this.gridPanel2.Span = "50% 50%;50% 50%;30% 30%";
            this.gridPanel2.TabIndex = 0;
            this.gridPanel2.Text = "gridPanel2";
            // 
            // UC_IQCPeeling
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.gridPanel1);
            this.Name = "UC_IQCPeeling";
            this.Size = new System.Drawing.Size(996, 628);
            this.gridPanel1.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        private AntdUI.GridPanel gridPanel2;

        private AntdUI.GridPanel gridPanel1;

        #endregion
    }
}