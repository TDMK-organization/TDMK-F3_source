using System.ComponentModel;

namespace Export_FPCA_OK2ship_Auto_System.Views
{
    partial class WorksheetScreen
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WorksheetScreen));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.panel1 = new AntdUI.Panel();
            this.btn_scan = new AntdUI.Button();
            this.ip = new AntdUI.Input();
            this.label2 = new AntdUI.Label();
            this.dgv = new AntdUI.Table();
            this.tableLayoutPanel1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 14.38356F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 85.61644F));
            this.tableLayoutPanel1.Controls.Add(this.panel1, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.label2, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.dgv, 1, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11.39785F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 88.60215F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(730, 465);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btn_scan);
            this.panel1.Controls.Add(this.ip);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(106, 2);
            this.panel1.Margin = new System.Windows.Forms.Padding(2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(622, 49);
            this.panel1.TabIndex = 7;
            this.panel1.Text = "panel1";
            // 
            // btn_scan
            // 
            this.btn_scan.BorderWidth = 1F;
            this.btn_scan.Dock = System.Windows.Forms.DockStyle.Right;
            this.btn_scan.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_scan.IconSvg = resources.GetString("btn_scan.IconSvg");
            this.btn_scan.JoinLeft = true;
            this.btn_scan.Location = new System.Drawing.Point(573, 0);
            this.btn_scan.Margin = new System.Windows.Forms.Padding(2);
            this.btn_scan.Name = "btn_scan";
            this.btn_scan.Size = new System.Drawing.Size(49, 49);
            this.btn_scan.TabIndex = 4;
            this.btn_scan.Type = AntdUI.TTypeMini.Primary;
            this.btn_scan.WaveSize = 0;
            this.btn_scan.Click += new System.EventHandler(this.btn_scan_Click);
            // 
            // ip
            // 
            this.ip.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ip.JoinRight = true;
            this.ip.Location = new System.Drawing.Point(0, 0);
            this.ip.Margin = new System.Windows.Forms.Padding(2);
            this.ip.Name = "ip";
            this.ip.PlaceholderText = "Click chosse file";
            this.ip.Size = new System.Drawing.Size(622, 49);
            this.ip.TabIndex = 3;
            this.ip.WaveSize = 0;
            this.ip.DoubleClick += new System.EventHandler(this.ip_DoubleClick);
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(3, 3);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(98, 47);
            this.label2.TabIndex = 9;
            this.label2.Text = "Location";
            // 
            // dgv
            // 
            this.dgv.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv.Gap = 12;
            this.dgv.Location = new System.Drawing.Point(107, 56);
            this.dgv.Name = "dgv";
            this.dgv.Size = new System.Drawing.Size(620, 406);
            this.dgv.TabIndex = 10;
            this.dgv.Text = "table1";
            // 
            // WorksheetScreen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "WorksheetScreen";
            this.Size = new System.Drawing.Size(730, 465);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        private AntdUI.Table dgv;

        private AntdUI.Label label2;

        private AntdUI.Panel panel1;
        private AntdUI.Button btn_scan;
        public AntdUI.Input ip;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;

        #endregion
    }
}