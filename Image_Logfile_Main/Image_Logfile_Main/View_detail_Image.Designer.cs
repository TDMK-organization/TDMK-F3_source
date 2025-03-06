namespace OK2SHIP
{
    partial class View_detail_Image
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
            this.image_detail = new System.Windows.Forms.PictureBox();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.lblImage_Graph = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.image_detail)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // image_detail
            // 
            this.image_detail.Dock = System.Windows.Forms.DockStyle.Fill;
            this.image_detail.Location = new System.Drawing.Point(3, 53);
            this.image_detail.Name = "image_detail";
            this.image_detail.Size = new System.Drawing.Size(878, 693);
            this.image_detail.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.image_detail.TabIndex = 0;
            this.image_detail.TabStop = false;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.lblImage_Graph, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.image_detail, 0, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(884, 749);
            this.tableLayoutPanel1.TabIndex = 1;
            // 
            // lblImage_Graph
            // 
            this.lblImage_Graph.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblImage_Graph.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblImage_Graph.Font = new System.Drawing.Font("Arial", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblImage_Graph.Location = new System.Drawing.Point(3, 0);
            this.lblImage_Graph.Name = "lblImage_Graph";
            this.lblImage_Graph.Size = new System.Drawing.Size(878, 50);
            this.lblImage_Graph.TabIndex = 5;
            this.lblImage_Graph.Text = "Image Details";
            this.lblImage_Graph.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // View_detail_Image
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(884, 749);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "View_detail_Image";
            this.Text = "View_detail_Image";
            this.Load += new System.EventHandler(this.View_detail_Image_Load);
            ((System.ComponentModel.ISupportInitialize)(this.image_detail)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox image_detail;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label lblImage_Graph;
    }
}