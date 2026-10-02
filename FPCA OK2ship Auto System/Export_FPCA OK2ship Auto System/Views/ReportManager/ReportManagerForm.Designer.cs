using System.ComponentModel;

namespace Export_FPCA_OK2ship_Auto_System.Views
{
    partial class ReportManagerForm
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel4 = new System.Windows.Forms.TableLayoutPanel();
            this.dgv = new AntdUI.Table();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.ip_config = new AntdUI.Input();
            this.ip_build = new AntdUI.Input();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btn_load = new AntdUI.Button();
            this.btn_save = new AntdUI.Button();
            this.btn_delete = new AntdUI.Button();
            this.ip_itemname = new AntdUI.Input();
            this.titlebar = new AntdUI.PageHeader();
            this.button_color = new AntdUI.Button();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel4.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.titlebar.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel4, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel3, 0, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 38F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 148F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1172, 662);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // tableLayoutPanel4
            // 
            this.tableLayoutPanel4.ColumnCount = 1;
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel4.Controls.Add(this.dgv, 0, 1);
            this.tableLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel4.Location = new System.Drawing.Point(3, 189);
            this.tableLayoutPanel4.Name = "tableLayoutPanel4";
            this.tableLayoutPanel4.RowCount = 2;
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 32F));
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel4.Size = new System.Drawing.Size(1166, 470);
            this.tableLayoutPanel4.TabIndex = 2;
            // 
            // dgv
            // 
            this.dgv.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv.Gap = 12;
            this.dgv.Location = new System.Drawing.Point(3, 35);
            this.dgv.Name = "dgv";
            this.dgv.Size = new System.Drawing.Size(1160, 432);
            this.dgv.TabIndex = 2;
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.ColumnCount = 2;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 482F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel3.Controls.Add(this.groupBox1, 0, 0);
            this.tableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel3.Location = new System.Drawing.Point(3, 41);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 1;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel3.Size = new System.Drawing.Size(1166, 142);
            this.tableLayoutPanel3.TabIndex = 1;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.tableLayoutPanel2);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(3, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(476, 136);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Data";
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.tableLayoutPanel2.ColumnCount = 3;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 117F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 185F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 188F));
            this.tableLayoutPanel2.Controls.Add(this.ip_config, 1, 2);
            this.tableLayoutPanel2.Controls.Add(this.ip_build, 1, 1);
            this.tableLayoutPanel2.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.label2, 0, 1);
            this.tableLayoutPanel2.Controls.Add(this.label3, 0, 2);
            this.tableLayoutPanel2.Controls.Add(this.btn_load, 2, 0);
            this.tableLayoutPanel2.Controls.Add(this.btn_save, 2, 1);
            this.tableLayoutPanel2.Controls.Add(this.btn_delete, 2, 2);
            this.tableLayoutPanel2.Controls.Add(this.ip_itemname, 1, 0);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(3, 16);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 3;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(470, 117);
            this.tableLayoutPanel2.TabIndex = 0;
            // 
            // ip_config
            // 
            this.ip_config.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ip_config.Location = new System.Drawing.Point(122, 80);
            this.ip_config.Name = "ip_config";
            this.ip_config.Size = new System.Drawing.Size(179, 33);
            this.ip_config.TabIndex = 7;
            this.ip_config.Leave += new System.EventHandler(this.ip_config_Leave);
            // 
            // ip_build
            // 
            this.ip_build.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ip_build.Location = new System.Drawing.Point(122, 42);
            this.ip_build.Name = "ip_build";
            this.ip_build.Size = new System.Drawing.Size(179, 31);
            this.ip_build.TabIndex = 6;
            this.ip_build.Leave += new System.EventHandler(this.ip_build_Leave);
            // 
            // label1
            // 
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label1.Location = new System.Drawing.Point(4, 1);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(111, 37);
            this.label1.TabIndex = 0;
            this.label1.Text = "Item Name";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label2
            // 
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label2.Location = new System.Drawing.Point(4, 39);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(111, 37);
            this.label2.TabIndex = 0;
            this.label2.Text = "Build";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label3
            // 
            this.label3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(4, 77);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(111, 39);
            this.label3.TabIndex = 1;
            this.label3.Text = "Config";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btn_load
            // 
            this.btn_load.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_load.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_load.Location = new System.Drawing.Point(308, 4);
            this.btn_load.Name = "btn_load";
            this.btn_load.Size = new System.Drawing.Size(182, 31);
            this.btn_load.TabIndex = 2;
            this.btn_load.Text = "Load";
            this.btn_load.Click += new System.EventHandler(this.btn_load_Click);
            // 
            // btn_save
            // 
            this.btn_save.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_save.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_save.Location = new System.Drawing.Point(308, 42);
            this.btn_save.Name = "btn_save";
            this.btn_save.Size = new System.Drawing.Size(182, 31);
            this.btn_save.TabIndex = 3;
            this.btn_save.Text = "Create New";
            this.btn_save.Click += new System.EventHandler(this.btn_save_Click);
            // 
            // btn_delete
            // 
            this.btn_delete.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_delete.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_delete.Location = new System.Drawing.Point(308, 80);
            this.btn_delete.Name = "btn_delete";
            this.btn_delete.Size = new System.Drawing.Size(182, 33);
            this.btn_delete.TabIndex = 4;
            this.btn_delete.Text = "Delete";
            this.btn_delete.Type = AntdUI.TTypeMini.Error;
            this.btn_delete.Click += new System.EventHandler(this.btn_delete_Click);
            // 
            // ip_itemname
            // 
            this.ip_itemname.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ip_itemname.Location = new System.Drawing.Point(122, 4);
            this.ip_itemname.Name = "ip_itemname";
            this.ip_itemname.Size = new System.Drawing.Size(179, 31);
            this.ip_itemname.TabIndex = 5;
            this.ip_itemname.Leave += new System.EventHandler(this.ip_itemname_Leave);
            // 
            // titlebar
            // 
            this.titlebar.Controls.Add(this.button_color);
            this.titlebar.DividerShow = true;
            this.titlebar.Dock = System.Windows.Forms.DockStyle.Top;
            this.titlebar.Location = new System.Drawing.Point(0, 0);
            this.titlebar.Name = "titlebar";
            this.titlebar.ShowButton = true;
            this.titlebar.ShowIcon = true;
            this.titlebar.Size = new System.Drawing.Size(1172, 36);
            this.titlebar.SubText = "Demo";
            this.titlebar.TabIndex = 0;
            this.titlebar.Text = "AntdUI";
            // 
            // button_color
            // 
            this.button_color.Dock = System.Windows.Forms.DockStyle.Right;
            this.button_color.Ghost = true;
            this.button_color.IconSvg = "SunOutlined";
            this.button_color.Location = new System.Drawing.Point(978, 0);
            this.button_color.Name = "button_color";
            this.button_color.Radius = 0;
            this.button_color.Size = new System.Drawing.Size(50, 36);
            this.button_color.TabIndex = 3;
            this.button_color.ToggleIconSvg = "MoonOutlined";
            this.button_color.WaveSize = 0;
            // 
            // ReportManagerForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1172, 662);
            this.Controls.Add(this.titlebar);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "ReportManagerForm";
            this.Text = "ReportManagerForm";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel4.ResumeLayout(false);
            this.tableLayoutPanel3.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.tableLayoutPanel2.ResumeLayout(false);
            this.titlebar.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        private AntdUI.Input ip_itemname;
        private AntdUI.Input ip_build;
        private AntdUI.Input ip_config;

        private AntdUI.PageHeader titlebar;
        private AntdUI.Button button_color;

        private AntdUI.Button btn_load;
        private AntdUI.Button btn_save;
        private AntdUI.Button btn_delete;

        private System.Windows.Forms.GroupBox groupBox1;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel4;

        private AntdUI.Table dgv;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;

        #endregion
    }
}