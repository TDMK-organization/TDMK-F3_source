namespace OK2SHIP
{
    partial class Check_Impedance_logfile
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
            this.lbltitle = new System.Windows.Forms.Label();
            this.DGV_impedance = new System.Windows.Forms.DataGridView();
            this.txt_nom_imp = new System.Windows.Forms.TextBox();
            this.txt_tol_imp = new System.Windows.Forms.TextBox();
            this.tblmain = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.dgv_spec_tracewidth = new System.Windows.Forms.DataGridView();
            this.label3 = new System.Windows.Forms.Label();
            this.DGV_spec_impedance = new System.Windows.Forms.DataGridView();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.label6 = new System.Windows.Forms.Label();
            this.btn_check = new System.Windows.Forms.Button();
            this.btn_get_infor_excel = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.tableLayoutPanel4 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel5 = new System.Windows.Forms.TableLayoutPanel();
            this.label8 = new System.Windows.Forms.Label();
            this.txt_tol_tracewidth = new System.Windows.Forms.TextBox();
            this.btn_check_tracewidth = new System.Windows.Forms.Button();
            this.txt_nom_tracewidth = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.tbldata = new System.Windows.Forms.TableLayoutPanel();
            this.DGV_tracewidth = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtLogfile = new System.Windows.Forms.TextBox();
            this.label17 = new System.Windows.Forms.Label();
            this.txtOperator = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.txtLotNo = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.txtItemCode = new System.Windows.Forms.TextBox();
            this.btnsave = new System.Windows.Forms.Button();
            this.tableLayoutPanel6 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel7 = new System.Windows.Forms.TableLayoutPanel();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_impedance)).BeginInit();
            this.tblmain.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_spec_tracewidth)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_spec_impedance)).BeginInit();
            this.tableLayoutPanel2.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            this.tableLayoutPanel4.SuspendLayout();
            this.tableLayoutPanel5.SuspendLayout();
            this.tbldata.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_tracewidth)).BeginInit();
            this.tableLayoutPanel6.SuspendLayout();
            this.tableLayoutPanel7.SuspendLayout();
            this.SuspendLayout();
            // 
            // lbltitle
            // 
            this.lbltitle.AutoSize = true;
            this.lbltitle.Location = new System.Drawing.Point(282, 52);
            this.lbltitle.Name = "lbltitle";
            this.lbltitle.Size = new System.Drawing.Size(0, 13);
            this.lbltitle.TabIndex = 6;
            // 
            // DGV_impedance
            // 
            this.DGV_impedance.AllowUserToAddRows = false;
            this.DGV_impedance.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.DGV_impedance.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_impedance.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_impedance.Location = new System.Drawing.Point(3, 43);
            this.DGV_impedance.Name = "DGV_impedance";
            this.DGV_impedance.Size = new System.Drawing.Size(737, 245);
            this.DGV_impedance.TabIndex = 5;
            this.DGV_impedance.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DGV_spec_CellContentClick);
            // 
            // txt_nom_imp
            // 
            this.txt_nom_imp.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txt_nom_imp.Location = new System.Drawing.Point(183, 3);
            this.txt_nom_imp.Name = "txt_nom_imp";
            this.txt_nom_imp.Size = new System.Drawing.Size(123, 20);
            this.txt_nom_imp.TabIndex = 11;
            // 
            // txt_tol_imp
            // 
            this.txt_tol_imp.Location = new System.Drawing.Point(3, 3);
            this.txt_tol_imp.Name = "txt_tol_imp";
            this.txt_tol_imp.Size = new System.Drawing.Size(50, 20);
            this.txt_tol_imp.TabIndex = 12;
            // 
            // tblmain
            // 
            this.tblmain.ColumnCount = 2;
            this.tblmain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 70F));
            this.tblmain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30F));
            this.tblmain.Controls.Add(this.tableLayoutPanel1, 1, 0);
            this.tblmain.Controls.Add(this.tbldata, 0, 0);
            this.tblmain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tblmain.Location = new System.Drawing.Point(3, 43);
            this.tblmain.Name = "tblmain";
            this.tblmain.RowCount = 1;
            this.tblmain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tblmain.Size = new System.Drawing.Size(1070, 589);
            this.tblmain.TabIndex = 13;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.dgv_spec_tracewidth, 0, 5);
            this.tableLayoutPanel1.Controls.Add(this.label3, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.DGV_spec_impedance, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.label7, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel4, 0, 4);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(752, 3);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 6;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 70F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 70F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(315, 583);
            this.tableLayoutPanel1.TabIndex = 1;
            // 
            // dgv_spec_tracewidth
            // 
            this.dgv_spec_tracewidth.AllowUserToAddRows = false;
            this.dgv_spec_tracewidth.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgv_spec_tracewidth.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_spec_tracewidth.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_spec_tracewidth.Location = new System.Drawing.Point(3, 404);
            this.dgv_spec_tracewidth.Name = "dgv_spec_tracewidth";
            this.dgv_spec_tracewidth.Size = new System.Drawing.Size(309, 176);
            this.dgv_spec_tracewidth.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label3.Font = new System.Drawing.Font("Arial", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(3, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(309, 40);
            this.label3.TabIndex = 6;
            this.label3.Text = "Impedance Spec";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // DGV_spec_impedance
            // 
            this.DGV_spec_impedance.AllowUserToAddRows = false;
            this.DGV_spec_impedance.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_spec_impedance.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_spec_impedance.Location = new System.Drawing.Point(3, 113);
            this.DGV_spec_impedance.Name = "DGV_spec_impedance";
            this.DGV_spec_impedance.Size = new System.Drawing.Size(309, 175);
            this.DGV_spec_impedance.TabIndex = 5;
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 3;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 129F));
            this.tableLayoutPanel2.Controls.Add(this.tableLayoutPanel3, 2, 1);
            this.tableLayoutPanel2.Controls.Add(this.btn_check, 0, 1);
            this.tableLayoutPanel2.Controls.Add(this.btn_get_infor_excel, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.txt_nom_imp, 2, 0);
            this.tableLayoutPanel2.Controls.Add(this.label4, 1, 0);
            this.tableLayoutPanel2.Controls.Add(this.label5, 1, 1);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(3, 43);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 2;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(309, 64);
            this.tableLayoutPanel2.TabIndex = 7;
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.ColumnCount = 2;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel3.Controls.Add(this.label6, 1, 0);
            this.tableLayoutPanel3.Controls.Add(this.txt_tol_imp, 0, 0);
            this.tableLayoutPanel3.Location = new System.Drawing.Point(183, 35);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 1;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 26F));
            this.tableLayoutPanel3.Size = new System.Drawing.Size(123, 26);
            this.tableLayoutPanel3.TabIndex = 14;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(64, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(56, 26);
            this.label6.TabIndex = 13;
            this.label6.Text = "%";
            // 
            // btn_check
            // 
            this.btn_check.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_check.Location = new System.Drawing.Point(3, 35);
            this.btn_check.Name = "btn_check";
            this.btn_check.Size = new System.Drawing.Size(84, 26);
            this.btn_check.TabIndex = 25;
            this.btn_check.Text = "CHECK";
            this.btn_check.UseVisualStyleBackColor = true;
            this.btn_check.Click += new System.EventHandler(this.btn_check_Click);
            // 
            // btn_get_infor_excel
            // 
            this.btn_get_infor_excel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_get_infor_excel.Location = new System.Drawing.Point(3, 3);
            this.btn_get_infor_excel.Name = "btn_get_infor_excel";
            this.btn_get_infor_excel.Size = new System.Drawing.Size(84, 26);
            this.btn_get_infor_excel.TabIndex = 24;
            this.btn_get_infor_excel.Text = "LOAD";
            this.btn_get_infor_excel.UseVisualStyleBackColor = true;
            this.btn_get_infor_excel.Click += new System.EventHandler(this.btn_get_infor_excel_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label4.Location = new System.Drawing.Point(93, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(84, 32);
            this.label4.TabIndex = 13;
            this.label4.Text = "NOM";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label5.Location = new System.Drawing.Point(93, 32);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(84, 32);
            this.label5.TabIndex = 14;
            this.label5.Text = "TOL";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label7.Font = new System.Drawing.Font("Arial", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(3, 291);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(309, 40);
            this.label7.TabIndex = 8;
            this.label7.Text = "Trace Width Spec";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tableLayoutPanel4
            // 
            this.tableLayoutPanel4.ColumnCount = 3;
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 129F));
            this.tableLayoutPanel4.Controls.Add(this.tableLayoutPanel5, 2, 1);
            this.tableLayoutPanel4.Controls.Add(this.btn_check_tracewidth, 0, 0);
            this.tableLayoutPanel4.Controls.Add(this.txt_nom_tracewidth, 2, 0);
            this.tableLayoutPanel4.Controls.Add(this.label9, 1, 0);
            this.tableLayoutPanel4.Controls.Add(this.label13, 1, 1);
            this.tableLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel4.Location = new System.Drawing.Point(3, 334);
            this.tableLayoutPanel4.Name = "tableLayoutPanel4";
            this.tableLayoutPanel4.RowCount = 2;
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel4.Size = new System.Drawing.Size(309, 64);
            this.tableLayoutPanel4.TabIndex = 9;
            // 
            // tableLayoutPanel5
            // 
            this.tableLayoutPanel5.ColumnCount = 2;
            this.tableLayoutPanel5.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel5.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel5.Controls.Add(this.label8, 1, 0);
            this.tableLayoutPanel5.Controls.Add(this.txt_tol_tracewidth, 0, 0);
            this.tableLayoutPanel5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel5.Location = new System.Drawing.Point(183, 35);
            this.tableLayoutPanel5.Name = "tableLayoutPanel5";
            this.tableLayoutPanel5.RowCount = 1;
            this.tableLayoutPanel5.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel5.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 26F));
            this.tableLayoutPanel5.Size = new System.Drawing.Size(123, 26);
            this.tableLayoutPanel5.TabIndex = 14;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(64, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(56, 26);
            this.label8.TabIndex = 13;
            this.label8.Text = "%";
            // 
            // txt_tol_tracewidth
            // 
            this.txt_tol_tracewidth.Location = new System.Drawing.Point(3, 3);
            this.txt_tol_tracewidth.Name = "txt_tol_tracewidth";
            this.txt_tol_tracewidth.Size = new System.Drawing.Size(50, 20);
            this.txt_tol_tracewidth.TabIndex = 12;
            // 
            // btn_check_tracewidth
            // 
            this.btn_check_tracewidth.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_check_tracewidth.Location = new System.Drawing.Point(3, 3);
            this.btn_check_tracewidth.Name = "btn_check_tracewidth";
            this.btn_check_tracewidth.Size = new System.Drawing.Size(84, 26);
            this.btn_check_tracewidth.TabIndex = 24;
            this.btn_check_tracewidth.Text = "CHECK";
            this.btn_check_tracewidth.UseVisualStyleBackColor = true;
            this.btn_check_tracewidth.Click += new System.EventHandler(this.btn_check_tracewidth_Click);
            // 
            // txt_nom_tracewidth
            // 
            this.txt_nom_tracewidth.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txt_nom_tracewidth.Location = new System.Drawing.Point(183, 3);
            this.txt_nom_tracewidth.Name = "txt_nom_tracewidth";
            this.txt_nom_tracewidth.Size = new System.Drawing.Size(123, 20);
            this.txt_nom_tracewidth.TabIndex = 11;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label9.Location = new System.Drawing.Point(93, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(84, 32);
            this.label9.TabIndex = 13;
            this.label9.Text = "NOM";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label13.Location = new System.Drawing.Point(93, 32);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(84, 32);
            this.label13.TabIndex = 14;
            this.label13.Text = "TOL";
            this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // tbldata
            // 
            this.tbldata.ColumnCount = 1;
            this.tbldata.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tbldata.Controls.Add(this.DGV_tracewidth, 0, 3);
            this.tbldata.Controls.Add(this.DGV_impedance, 0, 1);
            this.tbldata.Controls.Add(this.label2, 0, 0);
            this.tbldata.Controls.Add(this.label1, 0, 2);
            this.tbldata.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbldata.Location = new System.Drawing.Point(3, 3);
            this.tbldata.Name = "tbldata";
            this.tbldata.RowCount = 4;
            this.tbldata.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tbldata.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tbldata.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tbldata.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tbldata.Size = new System.Drawing.Size(743, 583);
            this.tbldata.TabIndex = 0;
            // 
            // DGV_tracewidth
            // 
            this.DGV_tracewidth.AllowUserToAddRows = false;
            this.DGV_tracewidth.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.DGV_tracewidth.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_tracewidth.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_tracewidth.Location = new System.Drawing.Point(3, 334);
            this.DGV_tracewidth.Name = "DGV_tracewidth";
            this.DGV_tracewidth.Size = new System.Drawing.Size(737, 246);
            this.DGV_tracewidth.TabIndex = 8;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Font = new System.Drawing.Font("Arial", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(3, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(737, 40);
            this.label2.TabIndex = 6;
            this.label2.Text = "Impedance Data";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Font = new System.Drawing.Font("Arial", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(3, 291);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(737, 40);
            this.label1.TabIndex = 7;
            this.label1.Text = "Trace Width Data";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtLogfile
            // 
            this.txtLogfile.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtLogfile.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtLogfile.Location = new System.Drawing.Point(669, 4);
            this.txtLogfile.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtLogfile.Name = "txtLogfile";
            this.txtLogfile.Size = new System.Drawing.Size(278, 22);
            this.txtLogfile.TabIndex = 23;
            this.txtLogfile.TextChanged += new System.EventHandler(this.txtLogfile_TextChanged);
            this.txtLogfile.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtLogfile_KeyDown);
            this.txtLogfile.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.txtLogfile_MouseDoubleClick);
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label17.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label17.Location = new System.Drawing.Point(564, 0);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(99, 34);
            this.label17.TabIndex = 22;
            this.label17.Text = "Log file location";
            this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtOperator
            // 
            this.txtOperator.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtOperator.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOperator.Location = new System.Drawing.Point(445, 4);
            this.txtOperator.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtOperator.Name = "txtOperator";
            this.txtOperator.Size = new System.Drawing.Size(113, 22);
            this.txtOperator.TabIndex = 19;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label12.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(377, 0);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(62, 34);
            this.label12.TabIndex = 21;
            this.label12.Text = "Operator";
            this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label11.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(198, 0);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(47, 34);
            this.label11.TabIndex = 20;
            this.label11.Text = "Lot No";
            this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtLotNo
            // 
            this.txtLotNo.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtLotNo.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtLotNo.Location = new System.Drawing.Point(251, 4);
            this.txtLotNo.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtLotNo.Name = "txtLotNo";
            this.txtLotNo.Size = new System.Drawing.Size(120, 22);
            this.txtLotNo.TabIndex = 18;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label10.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(3, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(65, 34);
            this.label10.TabIndex = 16;
            this.label10.Text = "ItemCode";
            this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtItemCode
            // 
            this.txtItemCode.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtItemCode.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtItemCode.Location = new System.Drawing.Point(74, 4);
            this.txtItemCode.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtItemCode.Name = "txtItemCode";
            this.txtItemCode.Size = new System.Drawing.Size(118, 22);
            this.txtItemCode.TabIndex = 17;
            // 
            // btnsave
            // 
            this.btnsave.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnsave.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnsave.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnsave.Location = new System.Drawing.Point(953, 3);
            this.btnsave.Name = "btnsave";
            this.btnsave.Size = new System.Drawing.Size(114, 28);
            this.btnsave.TabIndex = 24;
            this.btnsave.Text = "Load Data";
            this.btnsave.UseVisualStyleBackColor = false;
            this.btnsave.Click += new System.EventHandler(this.btnsave_Click);
            // 
            // tableLayoutPanel6
            // 
            this.tableLayoutPanel6.ColumnCount = 1;
            this.tableLayoutPanel6.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel6.Controls.Add(this.tblmain, 0, 1);
            this.tableLayoutPanel6.Controls.Add(this.tableLayoutPanel7, 0, 0);
            this.tableLayoutPanel6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel6.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel6.Name = "tableLayoutPanel6";
            this.tableLayoutPanel6.RowCount = 2;
            this.tableLayoutPanel6.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tableLayoutPanel6.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 99.99999F));
            this.tableLayoutPanel6.Size = new System.Drawing.Size(1076, 635);
            this.tableLayoutPanel6.TabIndex = 25;
            // 
            // tableLayoutPanel7
            // 
            this.tableLayoutPanel7.ColumnCount = 9;
            this.tableLayoutPanel7.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 6.635514F));
            this.tableLayoutPanel7.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 11.58879F));
            this.tableLayoutPanel7.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 4.953271F));
            this.tableLayoutPanel7.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 11.7757F));
            this.tableLayoutPanel7.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 6.35514F));
            this.tableLayoutPanel7.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 11.1215F));
            this.tableLayoutPanel7.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 9.813084F));
            this.tableLayoutPanel7.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 26.54206F));
            this.tableLayoutPanel7.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 11.11111F));
            this.tableLayoutPanel7.Controls.Add(this.label10, 0, 0);
            this.tableLayoutPanel7.Controls.Add(this.btnsave, 8, 0);
            this.tableLayoutPanel7.Controls.Add(this.txtItemCode, 1, 0);
            this.tableLayoutPanel7.Controls.Add(this.txtLogfile, 7, 0);
            this.tableLayoutPanel7.Controls.Add(this.label11, 2, 0);
            this.tableLayoutPanel7.Controls.Add(this.label17, 6, 0);
            this.tableLayoutPanel7.Controls.Add(this.txtLotNo, 3, 0);
            this.tableLayoutPanel7.Controls.Add(this.txtOperator, 5, 0);
            this.tableLayoutPanel7.Controls.Add(this.label12, 4, 0);
            this.tableLayoutPanel7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel7.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel7.Name = "tableLayoutPanel7";
            this.tableLayoutPanel7.RowCount = 1;
            this.tableLayoutPanel7.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel7.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34F));
            this.tableLayoutPanel7.Size = new System.Drawing.Size(1070, 34);
            this.tableLayoutPanel7.TabIndex = 14;
            // 
            // Check_Impedance_logfile
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1076, 635);
            this.Controls.Add(this.tableLayoutPanel6);
            this.Controls.Add(this.lbltitle);
            this.Name = "Check_Impedance_logfile";
            this.Text = "Check_Impedance_logfile";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Check_Impedance_logfile_FormClosed);
            this.Load += new System.EventHandler(this.Check_Impedance_logfile_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DGV_impedance)).EndInit();
            this.tblmain.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_spec_tracewidth)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_spec_impedance)).EndInit();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            this.tableLayoutPanel3.ResumeLayout(false);
            this.tableLayoutPanel3.PerformLayout();
            this.tableLayoutPanel4.ResumeLayout(false);
            this.tableLayoutPanel4.PerformLayout();
            this.tableLayoutPanel5.ResumeLayout(false);
            this.tableLayoutPanel5.PerformLayout();
            this.tbldata.ResumeLayout(false);
            this.tbldata.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_tracewidth)).EndInit();
            this.tableLayoutPanel6.ResumeLayout(false);
            this.tableLayoutPanel7.ResumeLayout(false);
            this.tableLayoutPanel7.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label lbltitle;
        private System.Windows.Forms.DataGridView DGV_impedance;
        private System.Windows.Forms.TextBox txt_nom_imp;
        private System.Windows.Forms.TextBox txt_tol_imp;
        private System.Windows.Forms.TableLayoutPanel tblmain;
        private System.Windows.Forms.TextBox txtLogfile;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.TextBox txtOperator;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox txtLotNo;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox txtItemCode;
        private System.Windows.Forms.TableLayoutPanel tbldata;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.DataGridView DGV_spec_impedance;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btn_get_infor_excel;
        private System.Windows.Forms.Button btn_check;
        private System.Windows.Forms.Button btnsave;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView DGV_tracewidth;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel4;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel5;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txt_tol_tracewidth;
        private System.Windows.Forms.Button btn_check_tracewidth;
        private System.Windows.Forms.TextBox txt_nom_tracewidth;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.DataGridView dgv_spec_tracewidth;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel6;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel7;
    }
}