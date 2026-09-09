using System.ComponentModel;
using System.Windows.Forms;

namespace OK2SHIP_SMT.UserControls.Bending
{
    partial class UC_BendingData
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UC_BendingData));
            AntdUI.Tabs.StyleLine styleLine1 = new AntdUI.Tabs.StyleLine();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel4 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel7 = new System.Windows.Forms.TableLayoutPanel();
            this.btn_Save = new AntdUI.Button();
            this.panel1 = new AntdUI.Panel();
            this.btn_login = new AntdUI.Button();
            this.tb_login = new AntdUI.Input();
            this.tableLayoutPanel12 = new System.Windows.Forms.TableLayoutPanel();
            this.label2 = new System.Windows.Forms.Label();
            this.tabs1 = new AntdUI.Tabs();
            this.tabPage1 = new AntdUI.TabPage();
            this.button_loading = new AntdUI.Button();
            this.tb_logfile = new AntdUI.Input();
            this.tabPage2 = new AntdUI.TabPage();
            this.tableLayoutPanel5 = new System.Windows.Forms.TableLayoutPanel();
            this.btn_loadData = new AntdUI.ButtonShadow();
            this.btn_export = new AntdUI.ButtonShadow();
            this.tableLayoutPanel6 = new System.Windows.Forms.TableLayoutPanel();
            this.tb_maker = new AntdUI.Input();
            this.label5 = new System.Windows.Forms.Label();
            this.tb_lotNo = new AntdUI.Input();
            this.label4 = new System.Windows.Forms.Label();
            this.tb_itemCode = new AntdUI.Input();
            this.label3 = new System.Windows.Forms.Label();
            this.tableLayoutPanel8 = new System.Windows.Forms.TableLayoutPanel();
            this.label1 = new AntdUI.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.tableLayoutPanel10 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel11 = new System.Windows.Forms.TableLayoutPanel();
            this.lb_rng = new AntdUI.Label();
            this.lb_usl = new AntdUI.Label();
            this.lb_total = new AntdUI.Label();
            this.lb_ok = new AntdUI.Label();
            this.lb_lsl = new AntdUI.Label();
            this.lb_T_NG = new AntdUI.Label();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.tableLayoutPanel9 = new System.Windows.Forms.TableLayoutPanel();
            this.menu_option = new AntdUI.Menu();
            this.dgv_error = new System.Windows.Forms.DataGridView();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            this.tableLayoutPanel4.SuspendLayout();
            this.tableLayoutPanel7.SuspendLayout();
            this.panel1.SuspendLayout();
            this.tableLayoutPanel12.SuspendLayout();
            this.tabs1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tableLayoutPanel5.SuspendLayout();
            this.tableLayoutPanel6.SuspendLayout();
            this.tableLayoutPanel8.SuspendLayout();
            this.tableLayoutPanel10.SuspendLayout();
            this.tableLayoutPanel11.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.tableLayoutPanel9.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_error)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel8, 0, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 120F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1107, 499);
            this.tableLayoutPanel1.TabIndex = 1;
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Inset;
            this.tableLayoutPanel2.ColumnCount = 2;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 22.32323F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 77.67677F));
            this.tableLayoutPanel2.Controls.Add(this.tableLayoutPanel3, 1, 0);
            this.tableLayoutPanel2.Controls.Add(this.tableLayoutPanel6, 0, 0);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel2.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 114F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(1101, 116);
            this.tableLayoutPanel2.TabIndex = 0;
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.tableLayoutPanel3.ColumnCount = 1;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel3.Controls.Add(this.tableLayoutPanel4, 0, 0);
            this.tableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel3.Location = new System.Drawing.Point(250, 4);
            this.tableLayoutPanel3.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 1;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.62069F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 107F));
            this.tableLayoutPanel3.Size = new System.Drawing.Size(847, 108);
            this.tableLayoutPanel3.TabIndex = 0;
            // 
            // tableLayoutPanel4
            // 
            this.tableLayoutPanel4.ColumnCount = 2;
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 72.42744F));
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 27.57256F));
            this.tableLayoutPanel4.Controls.Add(this.tableLayoutPanel7, 1, 0);
            this.tableLayoutPanel4.Controls.Add(this.tableLayoutPanel12, 0, 0);
            this.tableLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel4.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel4.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel4.Name = "tableLayoutPanel4";
            this.tableLayoutPanel4.RowCount = 1;
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 102F));
            this.tableLayoutPanel4.Size = new System.Drawing.Size(841, 102);
            this.tableLayoutPanel4.TabIndex = 1;
            // 
            // tableLayoutPanel7
            // 
            this.tableLayoutPanel7.ColumnCount = 1;
            this.tableLayoutPanel7.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel7.Controls.Add(this.btn_Save, 0, 1);
            this.tableLayoutPanel7.Controls.Add(this.panel1, 0, 0);
            this.tableLayoutPanel7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel7.Location = new System.Drawing.Point(611, 2);
            this.tableLayoutPanel7.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel7.Name = "tableLayoutPanel7";
            this.tableLayoutPanel7.RowCount = 2;
            this.tableLayoutPanel7.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tableLayoutPanel7.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel7.Size = new System.Drawing.Size(228, 98);
            this.tableLayoutPanel7.TabIndex = 1;
            // 
            // btn_Save
            // 
            this.btn_Save.BorderWidth = 2F;
            this.btn_Save.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_Save.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Save.Ghost = true;
            this.btn_Save.Location = new System.Drawing.Point(2, 42);
            this.btn_Save.Margin = new System.Windows.Forms.Padding(2);
            this.btn_Save.Name = "btn_Save";
            this.btn_Save.Size = new System.Drawing.Size(224, 54);
            this.btn_Save.TabIndex = 0;
            this.btn_Save.Text = "Save Data";
            this.btn_Save.Type = AntdUI.TTypeMini.Primary;
            this.btn_Save.Click += new System.EventHandler(this.btn_Save_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btn_login);
            this.panel1.Controls.Add(this.tb_login);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(2, 2);
            this.panel1.Margin = new System.Windows.Forms.Padding(2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(224, 36);
            this.panel1.TabIndex = 0;
            this.panel1.Text = "panel1";
            // 
            // btn_login
            // 
            this.btn_login.BorderWidth = 1F;
            this.btn_login.Dock = System.Windows.Forms.DockStyle.Right;
            this.btn_login.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_login.IconSvg = resources.GetString("btn_login.IconSvg");
            this.btn_login.JoinLeft = true;
            this.btn_login.Location = new System.Drawing.Point(175, 0);
            this.btn_login.Margin = new System.Windows.Forms.Padding(2);
            this.btn_login.Name = "btn_login";
            this.btn_login.Size = new System.Drawing.Size(49, 36);
            this.btn_login.TabIndex = 4;
            this.btn_login.Type = AntdUI.TTypeMini.Primary;
            this.btn_login.WaveSize = 0;
            this.btn_login.Click += new System.EventHandler(this.btn_login_Click);
            // 
            // tb_login
            // 
            this.tb_login.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tb_login.JoinRight = true;
            this.tb_login.Location = new System.Drawing.Point(0, 0);
            this.tb_login.Margin = new System.Windows.Forms.Padding(2);
            this.tb_login.Name = "tb_login";
            this.tb_login.PlaceholderText = "Login";
            this.tb_login.Size = new System.Drawing.Size(224, 36);
            this.tb_login.TabIndex = 3;
            this.tb_login.WaveSize = 0;
            // 
            // tableLayoutPanel12
            // 
            this.tableLayoutPanel12.ColumnCount = 1;
            this.tableLayoutPanel12.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel12.Controls.Add(this.label2, 0, 0);
            this.tableLayoutPanel12.Controls.Add(this.tabs1, 0, 1);
            this.tableLayoutPanel12.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel12.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel12.Name = "tableLayoutPanel12";
            this.tableLayoutPanel12.RowCount = 2;
            this.tableLayoutPanel12.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 39.79592F));
            this.tableLayoutPanel12.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 60.20408F));
            this.tableLayoutPanel12.Size = new System.Drawing.Size(603, 96);
            this.tableLayoutPanel12.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.LightSeaGreen;
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label2.Location = new System.Drawing.Point(0, 0);
            this.label2.Margin = new System.Windows.Forms.Padding(0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(603, 38);
            this.label2.TabIndex = 2;
            this.label2.Text = "Bending Data";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tabs1
            // 
            this.tabs1.Controls.Add(this.tabPage2);
            this.tabs1.Controls.Add(this.tabPage1);
            this.tabs1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.tabs1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabs1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabs1.Location = new System.Drawing.Point(2, 40);
            this.tabs1.Margin = new System.Windows.Forms.Padding(2);
            this.tabs1.Name = "tabs1";
            this.tabs1.Pages.Add(this.tabPage1);
            this.tabs1.Pages.Add(this.tabPage2);
            this.tabs1.SelectedIndex = 1;
            this.tabs1.Size = new System.Drawing.Size(599, 54);
            this.tabs1.Style = styleLine1;
            this.tabs1.TabIndex = 0;
            this.tabs1.Text = "tabs1";
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.button_loading);
            this.tabPage1.Controls.Add(this.tb_logfile);
            this.tabPage1.Location = new System.Drawing.Point(-1198, -46);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(2);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Size = new System.Drawing.Size(599, 23);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Logfile";
            // 
            // button_loading
            // 
            this.button_loading.BorderWidth = 1F;
            this.button_loading.Dock = System.Windows.Forms.DockStyle.Right;
            this.button_loading.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_loading.IconGap = 0F;
            this.button_loading.IconHoverSvg = "";
            this.button_loading.IconSvg = resources.GetString("button_loading.IconSvg");
            this.button_loading.JoinLeft = true;
            this.button_loading.Location = new System.Drawing.Point(474, 0);
            this.button_loading.Margin = new System.Windows.Forms.Padding(2);
            this.button_loading.Name = "button_loading";
            this.button_loading.Size = new System.Drawing.Size(125, 23);
            this.button_loading.TabIndex = 5;
            this.button_loading.Type = AntdUI.TTypeMini.Primary;
            this.button_loading.WaveSize = 0;
            this.button_loading.Click += new System.EventHandler(this.button_loading_Click);
            // 
            // tb_logfile
            // 
            this.tb_logfile.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tb_logfile.JoinRight = true;
            this.tb_logfile.Location = new System.Drawing.Point(0, 0);
            this.tb_logfile.Margin = new System.Windows.Forms.Padding(2);
            this.tb_logfile.Name = "tb_logfile";
            this.tb_logfile.PlaceholderText = "Double Click for select folder or copy on here";
            this.tb_logfile.Size = new System.Drawing.Size(599, 23);
            this.tb_logfile.TabIndex = 4;
            this.tb_logfile.WaveSize = 0;
            this.tb_logfile.TextChanged += new System.EventHandler(this.tb_logfile_TextChanged);
            this.tb_logfile.DoubleClick += new System.EventHandler(this.tb_logfile_DoubleClick);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.tableLayoutPanel5);
            this.tabPage2.Location = new System.Drawing.Point(0, 31);
            this.tabPage2.Margin = new System.Windows.Forms.Padding(2);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Size = new System.Drawing.Size(599, 23);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Database";
            // 
            // tableLayoutPanel5
            // 
            this.tableLayoutPanel5.ColumnCount = 2;
            this.tableLayoutPanel5.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel5.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel5.Controls.Add(this.btn_loadData, 0, 0);
            this.tableLayoutPanel5.Controls.Add(this.btn_export, 1, 0);
            this.tableLayoutPanel5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel5.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel5.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel5.Name = "tableLayoutPanel5";
            this.tableLayoutPanel5.RowCount = 1;
            this.tableLayoutPanel5.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel5.Size = new System.Drawing.Size(599, 23);
            this.tableLayoutPanel5.TabIndex = 0;
            // 
            // btn_loadData
            // 
            this.btn_loadData.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_loadData.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_loadData.Location = new System.Drawing.Point(2, 2);
            this.btn_loadData.Margin = new System.Windows.Forms.Padding(2);
            this.btn_loadData.Name = "btn_loadData";
            this.btn_loadData.Size = new System.Drawing.Size(295, 19);
            this.btn_loadData.TabIndex = 2;
            this.btn_loadData.Text = "Load Data";
            this.btn_loadData.Type = AntdUI.TTypeMini.Primary;
            this.btn_loadData.Click += new System.EventHandler(this.buttonShadow2_Click);
            // 
            // btn_export
            // 
            this.btn_export.BorderWidth = 2F;
            this.btn_export.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_export.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_export.Ghost = true;
            this.btn_export.Location = new System.Drawing.Point(301, 2);
            this.btn_export.Margin = new System.Windows.Forms.Padding(2);
            this.btn_export.Name = "btn_export";
            this.btn_export.Size = new System.Drawing.Size(296, 19);
            this.btn_export.TabIndex = 3;
            this.btn_export.Text = "Export Data";
            this.btn_export.Type = AntdUI.TTypeMini.Primary;
            this.btn_export.Click += new System.EventHandler(this.btn_export_Click);
            // 
            // tableLayoutPanel6
            // 
            this.tableLayoutPanel6.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.tableLayoutPanel6.ColumnCount = 2;
            this.tableLayoutPanel6.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 96F));
            this.tableLayoutPanel6.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel6.Controls.Add(this.tb_maker, 1, 2);
            this.tableLayoutPanel6.Controls.Add(this.label5, 0, 2);
            this.tableLayoutPanel6.Controls.Add(this.tb_lotNo, 1, 1);
            this.tableLayoutPanel6.Controls.Add(this.label4, 0, 1);
            this.tableLayoutPanel6.Controls.Add(this.tb_itemCode, 1, 0);
            this.tableLayoutPanel6.Controls.Add(this.label3, 0, 0);
            this.tableLayoutPanel6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel6.Location = new System.Drawing.Point(4, 4);
            this.tableLayoutPanel6.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel6.Name = "tableLayoutPanel6";
            this.tableLayoutPanel6.RowCount = 3;
            this.tableLayoutPanel6.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel6.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel6.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel6.Size = new System.Drawing.Size(240, 108);
            this.tableLayoutPanel6.TabIndex = 1;
            // 
            // tb_maker
            // 
            this.tb_maker.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tb_maker.Location = new System.Drawing.Point(100, 73);
            this.tb_maker.Margin = new System.Windows.Forms.Padding(2);
            this.tb_maker.Name = "tb_maker";
            this.tb_maker.Size = new System.Drawing.Size(137, 32);
            this.tb_maker.TabIndex = 5;
            // 
            // label5
            // 
            this.label5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(3, 71);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(92, 36);
            this.label5.TabIndex = 4;
            this.label5.Text = "Maker";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // tb_lotNo
            // 
            this.tb_lotNo.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tb_lotNo.Location = new System.Drawing.Point(100, 38);
            this.tb_lotNo.Margin = new System.Windows.Forms.Padding(2);
            this.tb_lotNo.Name = "tb_lotNo";
            this.tb_lotNo.Size = new System.Drawing.Size(137, 30);
            this.tb_lotNo.TabIndex = 3;
            this.tb_lotNo.Leave += new System.EventHandler(this.tb_lotNo_Leave);
            // 
            // label4
            // 
            this.label4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(3, 36);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(92, 34);
            this.label4.TabIndex = 2;
            this.label4.Text = "Lot No";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // tb_itemCode
            // 
            this.tb_itemCode.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tb_itemCode.Location = new System.Drawing.Point(100, 3);
            this.tb_itemCode.Margin = new System.Windows.Forms.Padding(2);
            this.tb_itemCode.Name = "tb_itemCode";
            this.tb_itemCode.Size = new System.Drawing.Size(137, 30);
            this.tb_itemCode.TabIndex = 0;
            this.tb_itemCode.StyleChanged += new System.EventHandler(this.tb_itemCode_StyleChanged);
            // 
            // label3
            // 
            this.label3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(3, 1);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(92, 34);
            this.label3.TabIndex = 1;
            this.label3.Text = "Item Code";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // tableLayoutPanel8
            // 
            this.tableLayoutPanel8.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.tableLayoutPanel8.ColumnCount = 2;
            this.tableLayoutPanel8.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 22.44692F));
            this.tableLayoutPanel8.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 77.55309F));
            this.tableLayoutPanel8.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel8.Controls.Add(this.label7, 1, 0);
            this.tableLayoutPanel8.Controls.Add(this.tableLayoutPanel10, 1, 1);
            this.tableLayoutPanel8.Controls.Add(this.tableLayoutPanel9, 0, 1);
            this.tableLayoutPanel8.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel8.Location = new System.Drawing.Point(4, 125);
            this.tableLayoutPanel8.Name = "tableLayoutPanel8";
            this.tableLayoutPanel8.RowCount = 2;
            this.tableLayoutPanel8.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 38F));
            this.tableLayoutPanel8.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel8.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel8.Size = new System.Drawing.Size(1099, 370);
            this.tableLayoutPanel8.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(4, 4);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(240, 32);
            this.label1.TabIndex = 1;
            this.label1.Text = "Category";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label7
            // 
            this.label7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold);
            this.label7.Location = new System.Drawing.Point(251, 1);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(844, 38);
            this.label7.TabIndex = 4;
            this.label7.Text = "Data";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tableLayoutPanel10
            // 
            this.tableLayoutPanel10.ColumnCount = 1;
            this.tableLayoutPanel10.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel10.Controls.Add(this.tableLayoutPanel11, 0, 1);
            this.tableLayoutPanel10.Controls.Add(this.dgv, 0, 0);
            this.tableLayoutPanel10.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel10.Location = new System.Drawing.Point(250, 42);
            this.tableLayoutPanel10.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel10.Name = "tableLayoutPanel10";
            this.tableLayoutPanel10.RowCount = 2;
            this.tableLayoutPanel10.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel10.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 36F));
            this.tableLayoutPanel10.Size = new System.Drawing.Size(846, 325);
            this.tableLayoutPanel10.TabIndex = 6;
            // 
            // tableLayoutPanel11
            // 
            this.tableLayoutPanel11.ColumnCount = 6;
            this.tableLayoutPanel11.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.tableLayoutPanel11.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.tableLayoutPanel11.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.tableLayoutPanel11.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.tableLayoutPanel11.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.tableLayoutPanel11.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.tableLayoutPanel11.Controls.Add(this.lb_rng, 3, 0);
            this.tableLayoutPanel11.Controls.Add(this.lb_usl, 1, 0);
            this.tableLayoutPanel11.Controls.Add(this.lb_total, 0, 0);
            this.tableLayoutPanel11.Controls.Add(this.lb_ok, 5, 0);
            this.tableLayoutPanel11.Controls.Add(this.lb_lsl, 1, 0);
            this.tableLayoutPanel11.Controls.Add(this.lb_T_NG, 4, 0);
            this.tableLayoutPanel11.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel11.Location = new System.Drawing.Point(2, 291);
            this.tableLayoutPanel11.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel11.Name = "tableLayoutPanel11";
            this.tableLayoutPanel11.RowCount = 1;
            this.tableLayoutPanel11.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel11.Size = new System.Drawing.Size(842, 32);
            this.tableLayoutPanel11.TabIndex = 3;
            // 
            // lb_rng
            // 
            this.lb_rng.BackColor = System.Drawing.Color.Red;
            this.lb_rng.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lb_rng.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lb_rng.Location = new System.Drawing.Point(422, 2);
            this.lb_rng.Margin = new System.Windows.Forms.Padding(2);
            this.lb_rng.Name = "lb_rng";
            this.lb_rng.Size = new System.Drawing.Size(136, 28);
            this.lb_rng.TabIndex = 3;
            this.lb_rng.Text = "R NG:";
            // 
            // lb_usl
            // 
            this.lb_usl.BackColor = System.Drawing.Color.Gold;
            this.lb_usl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lb_usl.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lb_usl.Location = new System.Drawing.Point(282, 2);
            this.lb_usl.Margin = new System.Windows.Forms.Padding(2);
            this.lb_usl.Name = "lb_usl";
            this.lb_usl.Size = new System.Drawing.Size(136, 28);
            this.lb_usl.TabIndex = 1;
            this.lb_usl.Text = "USL:";
            // 
            // lb_total
            // 
            this.lb_total.BackColor = System.Drawing.Color.Orchid;
            this.lb_total.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lb_total.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lb_total.Location = new System.Drawing.Point(2, 2);
            this.lb_total.Margin = new System.Windows.Forms.Padding(2);
            this.lb_total.Name = "lb_total";
            this.lb_total.Size = new System.Drawing.Size(136, 28);
            this.lb_total.TabIndex = 0;
            this.lb_total.Text = "Total item:";
            // 
            // lb_ok
            // 
            this.lb_ok.BackColor = System.Drawing.Color.DarkGreen;
            this.lb_ok.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lb_ok.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lb_ok.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.lb_ok.Location = new System.Drawing.Point(702, 2);
            this.lb_ok.Margin = new System.Windows.Forms.Padding(2);
            this.lb_ok.Name = "lb_ok";
            this.lb_ok.Size = new System.Drawing.Size(138, 28);
            this.lb_ok.TabIndex = 3;
            this.lb_ok.Text = "OK: ";
            // 
            // lb_lsl
            // 
            this.lb_lsl.BackColor = System.Drawing.Color.Cyan;
            this.lb_lsl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lb_lsl.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lb_lsl.Location = new System.Drawing.Point(142, 2);
            this.lb_lsl.Margin = new System.Windows.Forms.Padding(2);
            this.lb_lsl.Name = "lb_lsl";
            this.lb_lsl.Size = new System.Drawing.Size(136, 28);
            this.lb_lsl.TabIndex = 2;
            this.lb_lsl.Text = "LSL:";
            // 
            // lb_T_NG
            // 
            this.lb_T_NG.BackColor = System.Drawing.Color.DarkOrange;
            this.lb_T_NG.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lb_T_NG.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lb_T_NG.Location = new System.Drawing.Point(562, 2);
            this.lb_T_NG.Margin = new System.Windows.Forms.Padding(2);
            this.lb_T_NG.Name = "lb_T_NG";
            this.lb_T_NG.Size = new System.Drawing.Size(136, 28);
            this.lb_T_NG.TabIndex = 4;
            this.lb_T_NG.Text = "T_NG:";
            // 
            // dgv
            // 
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgv.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv.Location = new System.Drawing.Point(2, 2);
            this.dgv.Margin = new System.Windows.Forms.Padding(2);
            this.dgv.Name = "dgv";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dgv.RowTemplate.Height = 28;
            this.dgv.Size = new System.Drawing.Size(842, 285);
            this.dgv.TabIndex = 4;
            this.dgv.CellMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgv_CellMouseClick);
            this.dgv.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dgv_CellPainting);
            this.dgv.CellToolTipTextNeeded += new System.Windows.Forms.DataGridViewCellToolTipTextNeededEventHandler(this.dataGridView1_CellToolTipTextNeeded);
            // 
            // tableLayoutPanel9
            // 
            this.tableLayoutPanel9.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.tableLayoutPanel9.ColumnCount = 1;
            this.tableLayoutPanel9.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel9.Controls.Add(this.menu_option, 0, 0);
            this.tableLayoutPanel9.Controls.Add(this.dgv_error, 0, 1);
            this.tableLayoutPanel9.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel9.Location = new System.Drawing.Point(3, 42);
            this.tableLayoutPanel9.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel9.Name = "tableLayoutPanel9";
            this.tableLayoutPanel9.RowCount = 2;
            this.tableLayoutPanel9.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 74.50111F));
            this.tableLayoutPanel9.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25.49889F));
            this.tableLayoutPanel9.Size = new System.Drawing.Size(242, 325);
            this.tableLayoutPanel9.TabIndex = 7;
            // 
            // menu_option
            // 
            this.menu_option.Dock = System.Windows.Forms.DockStyle.Fill;
            this.menu_option.Location = new System.Drawing.Point(4, 4);
            this.menu_option.Name = "menu_option";
            this.menu_option.Size = new System.Drawing.Size(234, 233);
            this.menu_option.TabIndex = 0;
            this.menu_option.Text = "menu1";
            this.menu_option.SelectChanged += new AntdUI.SelectEventHandler(this.menu_option_SelectChanged);
            this.menu_option.SelectChanging += new AntdUI.SelectBoolEventHandler(this.menu_option_SelectChanging);
            // 
            // dgv_error
            // 
            this.dgv_error.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_error.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_error.Location = new System.Drawing.Point(4, 244);
            this.dgv_error.Name = "dgv_error";
            this.dgv_error.Size = new System.Drawing.Size(234, 77);
            this.dgv_error.TabIndex = 1;
            // 
            // UC_BendingData
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "UC_BendingData";
            this.Size = new System.Drawing.Size(1107, 499);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel3.ResumeLayout(false);
            this.tableLayoutPanel4.ResumeLayout(false);
            this.tableLayoutPanel7.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.tableLayoutPanel12.ResumeLayout(false);
            this.tabs1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tableLayoutPanel5.ResumeLayout(false);
            this.tableLayoutPanel6.ResumeLayout(false);
            this.tableLayoutPanel8.ResumeLayout(false);
            this.tableLayoutPanel10.ResumeLayout(false);
            this.tableLayoutPanel11.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.tableLayoutPanel9.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_error)).EndInit();
            this.ResumeLayout(false);
        }

        private System.Windows.Forms.DataGridView dgv_error;

        private AntdUI.Label lb_T_NG;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel12;

        private System.Windows.Forms.DataGridView dgv;

        private AntdUI.Label lb_ok;

        private AntdUI.Label lb_usl;
        private AntdUI.Label lb_lsl;
        private AntdUI.Label lb_rng;

        private AntdUI.Label lb_total;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel11;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel10;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel9;

        private System.Windows.Forms.Label label7;

        private AntdUI.Label label1;

        private AntdUI.Menu menu_option;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel8;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel4;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel7;
        private AntdUI.Button btn_Save;
        private AntdUI.Panel panel1;
        private AntdUI.Button btn_login;
        private AntdUI.Input tb_login;
        private AntdUI.Tabs tabs1;
        private AntdUI.TabPage tabPage1;
        private AntdUI.Button button_loading;
        private AntdUI.Input tb_logfile;
        private AntdUI.TabPage tabPage2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel5;
        private AntdUI.ButtonShadow btn_loadData;
        private AntdUI.ButtonShadow btn_export;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel6;
        private AntdUI.Input tb_maker;
        private System.Windows.Forms.Label label5;
        private AntdUI.Input tb_lotNo;
        private System.Windows.Forms.Label label4;
        private AntdUI.Input tb_itemCode;
        private System.Windows.Forms.Label label3;

        #endregion
    }
}