namespace OK2SHIP_SMT.UserControls.Logins
{
    partial class UC_Information
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
            this.Button_Area = new System.Windows.Forms.FlowLayoutPanel();
            this.TableLayout = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.Button_Area, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.TableLayout, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 80F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(923, 599);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // Button_Area
            // 
            this.Button_Area.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Button_Area.Location = new System.Drawing.Point(3, 482);
            this.Button_Area.Name = "Button_Area";
            this.Button_Area.Size = new System.Drawing.Size(917, 114);
            this.Button_Area.TabIndex = 0;
            // 
            // TableLayout
            // 
            this.TableLayout.ColumnCount = 2;
            this.TableLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30F));
            this.TableLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 70F));
            this.TableLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.TableLayout.Location = new System.Drawing.Point(3, 3);
            this.TableLayout.Name = "TableLayout";
            this.TableLayout.RowCount = 1;
            this.TableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.TableLayout.Size = new System.Drawing.Size(917, 473);
            this.TableLayout.TabIndex = 1;
            // 
            // UC_Information
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "UC_Information";
            this.Size = new System.Drawing.Size(923, 599);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.FlowLayoutPanel Button_Area;
        private System.Windows.Forms.TableLayoutPanel TableLayout;
    }
}
