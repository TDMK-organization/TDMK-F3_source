namespace OK2SHIP_Software
{
    partial class FAI_Data
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
            this.label1 = new System.Windows.Forms.Label();
            this.DGV_Spec = new System.Windows.Forms.DataGridView();
            this.DGV_Data = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.DGV_CPK = new System.Windows.Forms.DataGridView();
            this.label3 = new System.Windows.Forms.Label();
            this.tblMain = new System.Windows.Forms.TableLayoutPanel();
            this.label4 = new System.Windows.Forms.Label();
            this.DGV_Histogram = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Spec)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_CPK)).BeginInit();
            this.tblMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Histogram)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Font = new System.Drawing.Font("Arial", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(1178, 22);
            this.label1.TabIndex = 0;
            this.label1.Text = "FAI Specification";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // DGV_Spec
            // 
            this.DGV_Spec.AllowUserToAddRows = false;
            this.DGV_Spec.AllowUserToDeleteRows = false;
            this.DGV_Spec.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Spec.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_Spec.Location = new System.Drawing.Point(3, 25);
            this.DGV_Spec.Name = "DGV_Spec";
            this.DGV_Spec.Size = new System.Drawing.Size(1178, 123);
            this.DGV_Spec.TabIndex = 1;
            // 
            // DGV_Data
            // 
            this.DGV_Data.AllowUserToAddRows = false;
            this.DGV_Data.AllowUserToDeleteRows = false;
            this.DGV_Data.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Data.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_Data.Location = new System.Drawing.Point(3, 478);
            this.DGV_Data.Name = "DGV_Data";
            this.DGV_Data.Size = new System.Drawing.Size(1178, 280);
            this.DGV_Data.TabIndex = 3;
            this.DGV_Data.ColumnHeaderMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.DGV_Data_ColumnHeaderMouseDoubleClick);
            // 
            // label2
            // 
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Font = new System.Drawing.Font("Arial", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(3, 453);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(1178, 22);
            this.label2.TabIndex = 2;
            this.label2.Text = "FAI Data";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // DGV_CPK
            // 
            this.DGV_CPK.AllowUserToAddRows = false;
            this.DGV_CPK.AllowUserToDeleteRows = false;
            this.DGV_CPK.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_CPK.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_CPK.Location = new System.Drawing.Point(3, 176);
            this.DGV_CPK.Name = "DGV_CPK";
            this.DGV_CPK.Size = new System.Drawing.Size(1178, 123);
            this.DGV_CPK.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label3.Font = new System.Drawing.Font("Arial", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(3, 151);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(1178, 22);
            this.label3.TabIndex = 4;
            this.label3.Text = "Calculate CPK";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tblMain
            // 
            this.tblMain.ColumnCount = 1;
            this.tblMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tblMain.Controls.Add(this.label1, 0, 0);
            this.tblMain.Controls.Add(this.DGV_CPK, 0, 3);
            this.tblMain.Controls.Add(this.DGV_Spec, 0, 1);
            this.tblMain.Controls.Add(this.label3, 0, 2);
            this.tblMain.Controls.Add(this.DGV_Data, 0, 7);
            this.tblMain.Controls.Add(this.label2, 0, 6);
            this.tblMain.Controls.Add(this.label4, 0, 4);
            this.tblMain.Controls.Add(this.DGV_Histogram, 0, 5);
            this.tblMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tblMain.Location = new System.Drawing.Point(0, 0);
            this.tblMain.Name = "tblMain";
            this.tblMain.RowCount = 8;
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 3F));
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 17F));
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 3F));
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 17F));
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 3F));
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 17F));
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 3F));
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 37F));
            this.tblMain.Size = new System.Drawing.Size(1184, 761);
            this.tblMain.TabIndex = 6;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label4.Font = new System.Drawing.Font("Arial", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(3, 302);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(1178, 22);
            this.label4.TabIndex = 6;
            this.label4.Text = "Histogram";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // DGV_Histogram
            // 
            this.DGV_Histogram.AllowUserToAddRows = false;
            this.DGV_Histogram.AllowUserToDeleteRows = false;
            this.DGV_Histogram.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Histogram.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_Histogram.Location = new System.Drawing.Point(3, 327);
            this.DGV_Histogram.Name = "DGV_Histogram";
            this.DGV_Histogram.Size = new System.Drawing.Size(1178, 123);
            this.DGV_Histogram.TabIndex = 7;
            // 
            // FAI_Data
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1184, 761);
            this.Controls.Add(this.tblMain);
            this.Name = "FAI_Data";
            this.Text = "FAI_Data Analysis";
            this.Load += new System.EventHandler(this.FAI_Data_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Spec)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_CPK)).EndInit();
            this.tblMain.ResumeLayout(false);
            this.tblMain.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Histogram)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView DGV_Spec;
        private System.Windows.Forms.DataGridView DGV_Data;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridView DGV_CPK;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TableLayoutPanel tblMain;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DataGridView DGV_Histogram;
    }
}