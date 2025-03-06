
namespace OK2SHIP
{
    partial class FrmMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmMain));
            this.lblType1 = new System.Windows.Forms.Label();
            this.lblType2 = new System.Windows.Forms.Label();
            this.lblType3 = new System.Windows.Forms.Label();
            this.lblType4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblType1
            // 
            this.lblType1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblType1.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblType1.Image = ((System.Drawing.Image)(resources.GetObject("lblType1.Image")));
            this.lblType1.Location = new System.Drawing.Point(12, 9);
            this.lblType1.Name = "lblType1";
            this.lblType1.Size = new System.Drawing.Size(250, 250);
            this.lblType1.TabIndex = 0;
            this.lblType1.Text = "ECheck Process";
            this.lblType1.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.lblType1.Click += new System.EventHandler(this.lblType1_Click);
            // 
            // lblType2
            // 
            this.lblType2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblType2.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblType2.Image = ((System.Drawing.Image)(resources.GetObject("lblType2.Image")));
            this.lblType2.Location = new System.Drawing.Point(12, 280);
            this.lblType2.Name = "lblType2";
            this.lblType2.Size = new System.Drawing.Size(250, 250);
            this.lblType2.TabIndex = 1;
            this.lblType2.Text = "Camera Process";
            this.lblType2.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.lblType2.Click += new System.EventHandler(this.lblType2_Click);
            // 
            // lblType3
            // 
            this.lblType3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblType3.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblType3.Image = ((System.Drawing.Image)(resources.GetObject("lblType3.Image")));
            this.lblType3.Location = new System.Drawing.Point(292, 9);
            this.lblType3.Name = "lblType3";
            this.lblType3.Size = new System.Drawing.Size(250, 250);
            this.lblType3.TabIndex = 2;
            this.lblType3.Text = "VHX Data Analysis";
            this.lblType3.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.lblType3.Click += new System.EventHandler(this.lblType3_Click);
            // 
            // lblType4
            // 
            this.lblType4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblType4.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblType4.Image = ((System.Drawing.Image)(resources.GetObject("lblType4.Image")));
            this.lblType4.Location = new System.Drawing.Point(292, 280);
            this.lblType4.Name = "lblType4";
            this.lblType4.Size = new System.Drawing.Size(250, 250);
            this.lblType4.TabIndex = 3;
            this.lblType4.Text = "Manual Input";
            this.lblType4.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.lblType4.Click += new System.EventHandler(this.lblType4_Click);
            // 
            // FrmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(560, 548);
            this.Controls.Add(this.lblType4);
            this.Controls.Add(this.lblType3);
            this.Controls.Add(this.lblType2);
            this.Controls.Add(this.lblType1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmMain";
            this.Text = "Main Screen";
            this.Load += new System.EventHandler(this.FrmMain_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lblType1;
        private System.Windows.Forms.Label lblType2;
        private System.Windows.Forms.Label lblType3;
        private System.Windows.Forms.Label lblType4;
    }
}

