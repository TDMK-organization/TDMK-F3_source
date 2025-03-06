namespace OK2SHIP
{
    partial class Image_history
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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.label23 = new System.Windows.Forms.Label();
            this.label22 = new System.Windows.Forms.Label();
            this.pic_before = new System.Windows.Forms.PictureBox();
            this.pic_after = new System.Windows.Forms.PictureBox();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pic_before)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_after)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.pic_after, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.pic_before, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.label23, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.label22, 1, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1111, 640);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label23.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label23.Font = new System.Drawing.Font("Arial", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label23.Location = new System.Drawing.Point(3, 0);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(549, 40);
            this.label23.TabIndex = 56;
            this.label23.Text = "Image Before";
            this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label22.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label22.Font = new System.Drawing.Font("Arial", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label22.Location = new System.Drawing.Point(558, 0);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(550, 40);
            this.label22.TabIndex = 57;
            this.label22.Text = "Image After";
            this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pic_before
            // 
            this.pic_before.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pic_before.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pic_before.Location = new System.Drawing.Point(558, 43);
            this.pic_before.Name = "pic_before";
            this.pic_before.Size = new System.Drawing.Size(550, 594);
            this.pic_before.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_before.TabIndex = 58;
            this.pic_before.TabStop = false;
            // 
            // pic_after
            // 
            this.pic_after.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pic_after.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pic_after.Location = new System.Drawing.Point(3, 43);
            this.pic_after.Name = "pic_after";
            this.pic_after.Size = new System.Drawing.Size(549, 594);
            this.pic_after.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_after.TabIndex = 59;
            this.pic_after.TabStop = false;
            // 
            // Image_history
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1111, 640);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "Image_history";
            this.Text = "Image_history";
            this.Load += new System.EventHandler(this.Image_history_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pic_before)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_after)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.PictureBox pic_before;
        private System.Windows.Forms.PictureBox pic_after;
    }
}