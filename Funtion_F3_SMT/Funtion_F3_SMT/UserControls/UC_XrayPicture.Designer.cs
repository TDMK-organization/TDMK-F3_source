namespace OK2SHIP_SMT.UserControls
{
    partial class UC_XrayPicture
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.pictureBox = new System.Windows.Forms.PictureBox();
            this.tbl_dta = new System.Windows.Forms.TableLayoutPanel();
            this.tdmK_Label1 = new OK2SHIP_SMT.ToolBoxs.TDMK_Label();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox)).BeginInit();
            this.tbl_dta.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.BackColor = System.Drawing.SystemColors.Control;
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.pictureBox, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.tbl_dta, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1103, 565);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // pictureBox
            // 
            this.pictureBox.BackColor = System.Drawing.Color.DodgerBlue;
            this.pictureBox.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.pictureBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pictureBox.InitialImage = null;
            this.pictureBox.Location = new System.Drawing.Point(554, 3);
            this.pictureBox.Name = "pictureBox";
            this.pictureBox.Size = new System.Drawing.Size(546, 559);
            this.pictureBox.TabIndex = 0;
            this.pictureBox.TabStop = false;
            // 
            // tbl_dta
            // 
            this.tbl_dta.ColumnCount = 1;
            this.tbl_dta.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tbl_dta.Controls.Add(this.tdmK_Label1, 0, 0);
            this.tbl_dta.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbl_dta.Location = new System.Drawing.Point(3, 3);
            this.tbl_dta.Name = "tbl_dta";
            this.tbl_dta.RowCount = 2;
            this.tbl_dta.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.55456F));
            this.tbl_dta.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 89.44543F));
            this.tbl_dta.Size = new System.Drawing.Size(545, 559);
            this.tbl_dta.TabIndex = 1;
            // 
            // tdmK_Label1
            // 
            this.tdmK_Label1.AutoSize = true;
            this.tdmK_Label1.BackColor = System.Drawing.Color.LightPink;
            this.tdmK_Label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tdmK_Label1.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold);
            this.tdmK_Label1.Location = new System.Drawing.Point(3, 0);
            this.tdmK_Label1.Name = "tdmK_Label1";
            this.tdmK_Label1.Size = new System.Drawing.Size(539, 59);
            this.tdmK_Label1.TabIndex = 0;
            this.tdmK_Label1.Text = "List Image";
            this.tdmK_Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // UC_XrayPicture
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "UC_XrayPicture";
            this.Size = new System.Drawing.Size(1103, 565);
            this.tableLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox)).EndInit();
            this.tbl_dta.ResumeLayout(false);
            this.tbl_dta.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.PictureBox pictureBox;
        private System.Windows.Forms.TableLayoutPanel tbl_dta;
        private ToolBoxs.TDMK_Label tdmK_Label1;
    }
}
