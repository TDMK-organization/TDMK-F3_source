namespace Export_FPCA_OK2ship_Auto_System
{
    partial class Setup_Spec_SMT
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.lst_Item = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cbl_sheet = new System.Windows.Forms.CheckedListBox();
            this.GBInfo = new System.Windows.Forms.GroupBox();
            this.cb_Type = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnRUN = new System.Windows.Forms.Button();
            this.txtFormat = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.txtItemCode = new System.Windows.Forms.TextBox();
            this.lblItemCode = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btn_all = new System.Windows.Forms.Button();
            this.btn_filter = new System.Windows.Forms.Button();
            this.txt_ItemCode_filter = new System.Windows.Forms.TextBox();
            this.dgv_setup_detail = new System.Windows.Forms.DataGridView();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.cbAll = new System.Windows.Forms.CheckBox();
            this.GBInfo.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_setup_detail)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // lst_Item
            // 
            this.lst_Item.FormattingEnabled = true;
            this.lst_Item.ItemHeight = 20;
            this.lst_Item.Location = new System.Drawing.Point(748, 31);
            this.lst_Item.Name = "lst_Item";
            this.lst_Item.Size = new System.Drawing.Size(148, 4);
            this.lst_Item.TabIndex = 3;
            this.lst_Item.Visible = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(6, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(152, 20);
            this.label1.TabIndex = 4;
            this.label1.Text = "ItemCode đã cài đặt";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // cbl_sheet
            // 
            this.cbl_sheet.Cursor = System.Windows.Forms.Cursors.Hand;
            this.cbl_sheet.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbl_sheet.FormattingEnabled = true;
            this.cbl_sheet.Items.AddRange(new object[] {
            "FAI",
            "Cross section",
            "GAP Connector",
            "Peel Test",
            "(Mating) Pull Test",
            "(IQC Unmating) Pull Test",
            "Shear test",
            "IQC Liner peeling (Coupon)",
            "IQC PSA peeling (Coupon)",
            "Liner peel test (On product)",
            "PSA peel test (On product)",
            "ACF"});
            this.cbl_sheet.Location = new System.Drawing.Point(10, 81);
            this.cbl_sheet.Name = "cbl_sheet";
            this.cbl_sheet.Size = new System.Drawing.Size(234, 256);
            this.cbl_sheet.TabIndex = 12;
            this.cbl_sheet.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.cbl_sheet_ItemCheck);
            this.cbl_sheet.SelectedIndexChanged += new System.EventHandler(this.cbl_sheet_SelectedIndexChanged);
            // 
            // GBInfo
            // 
            this.GBInfo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.GBInfo.Controls.Add(this.cb_Type);
            this.GBInfo.Controls.Add(this.lst_Item);
            this.GBInfo.Controls.Add(this.label3);
            this.GBInfo.Controls.Add(this.btnRUN);
            this.GBInfo.Controls.Add(this.txtFormat);
            this.GBInfo.Controls.Add(this.btnBrowse);
            this.GBInfo.Controls.Add(this.txtItemCode);
            this.GBInfo.Controls.Add(this.lblItemCode);
            this.GBInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GBInfo.Location = new System.Drawing.Point(22, 13);
            this.GBInfo.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.GBInfo.Name = "GBInfo";
            this.GBInfo.Padding = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.GBInfo.Size = new System.Drawing.Size(902, 163);
            this.GBInfo.TabIndex = 14;
            this.GBInfo.TabStop = false;
            this.GBInfo.Text = "Nhập thông tin";
            // 
            // cb_Type
            // 
            this.cb_Type.FormattingEnabled = true;
            this.cb_Type.Items.AddRange(new object[] {
            "NPI",
            "MASS",
            "Other"});
            this.cb_Type.Location = new System.Drawing.Point(313, 32);
            this.cb_Type.Name = "cb_Type";
            this.cb_Type.Size = new System.Drawing.Size(121, 28);
            this.cb_Type.TabIndex = 18;
            this.cb_Type.SelectedIndexChanged += new System.EventHandler(this.cb_Type_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(264, 35);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(43, 20);
            this.label3.TabIndex = 17;
            this.label3.Text = "Type";
            // 
            // btnRUN
            // 
            this.btnRUN.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnRUN.Location = new System.Drawing.Point(10, 120);
            this.btnRUN.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnRUN.Name = "btnRUN";
            this.btnRUN.Size = new System.Drawing.Size(70, 35);
            this.btnRUN.TabIndex = 16;
            this.btnRUN.Text = "RUN";
            this.btnRUN.UseVisualStyleBackColor = true;
            this.btnRUN.Click += new System.EventHandler(this.btnRUN_Click);
            // 
            // txtFormat
            // 
            this.txtFormat.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtFormat.Location = new System.Drawing.Point(85, 77);
            this.txtFormat.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtFormat.Multiline = true;
            this.txtFormat.Name = "txtFormat";
            this.txtFormat.Size = new System.Drawing.Size(811, 78);
            this.txtFormat.TabIndex = 10;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnBrowse.Location = new System.Drawing.Point(9, 77);
            this.btnBrowse.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(70, 35);
            this.btnBrowse.TabIndex = 8;
            this.btnBrowse.Text = "Browse...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // txtItemCode
            // 
            this.txtItemCode.Location = new System.Drawing.Point(91, 34);
            this.txtItemCode.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtItemCode.Name = "txtItemCode";
            this.txtItemCode.Size = new System.Drawing.Size(117, 26);
            this.txtItemCode.TabIndex = 1;
            this.txtItemCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtItemCode.TextChanged += new System.EventHandler(this.txtItemCode_TextChanged);
            this.txtItemCode.Validated += new System.EventHandler(this.txtItemCode_Validated);
            // 
            // lblItemCode
            // 
            this.lblItemCode.AutoSize = true;
            this.lblItemCode.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblItemCode.Location = new System.Drawing.Point(6, 40);
            this.lblItemCode.Name = "lblItemCode";
            this.lblItemCode.Size = new System.Drawing.Size(79, 20);
            this.lblItemCode.TabIndex = 0;
            this.lblItemCode.Text = "ItemCode";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(12, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(142, 20);
            this.label2.TabIndex = 15;
            this.label2.Text = "Chọn sheet cài đặt";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.btn_all);
            this.groupBox1.Controls.Add(this.btn_filter);
            this.groupBox1.Controls.Add(this.txt_ItemCode_filter);
            this.groupBox1.Controls.Add(this.dgv_setup_detail);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(280, 197);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(644, 524);
            this.groupBox1.TabIndex = 16;
            this.groupBox1.TabStop = false;
            // 
            // btn_all
            // 
            this.btn_all.Location = new System.Drawing.Point(199, 47);
            this.btn_all.Name = "btn_all";
            this.btn_all.Size = new System.Drawing.Size(70, 27);
            this.btn_all.TabIndex = 21;
            this.btn_all.Text = "ALL";
            this.btn_all.UseVisualStyleBackColor = true;
            this.btn_all.Click += new System.EventHandler(this.btn_all_Click);
            // 
            // btn_filter
            // 
            this.btn_filter.Location = new System.Drawing.Point(129, 47);
            this.btn_filter.Name = "btn_filter";
            this.btn_filter.Size = new System.Drawing.Size(70, 28);
            this.btn_filter.TabIndex = 20;
            this.btn_filter.Text = "Filter";
            this.btn_filter.UseVisualStyleBackColor = true;
            this.btn_filter.Click += new System.EventHandler(this.btn_filter_Click);
            // 
            // txt_ItemCode_filter
            // 
            this.txt_ItemCode_filter.Location = new System.Drawing.Point(10, 48);
            this.txt_ItemCode_filter.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txt_ItemCode_filter.Name = "txt_ItemCode_filter";
            this.txt_ItemCode_filter.Size = new System.Drawing.Size(117, 26);
            this.txt_ItemCode_filter.TabIndex = 19;
            this.txt_ItemCode_filter.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txt_ItemCode_filter.TextChanged += new System.EventHandler(this.txt_ItemCode_filter_TextChanged);
            this.txt_ItemCode_filter.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txt_ItemCode_filter_KeyDown);
            this.txt_ItemCode_filter.Validated += new System.EventHandler(this.txt_ItemCode_filter_Validated);
            // 
            // dgv_setup_detail
            // 
            this.dgv_setup_detail.AllowUserToAddRows = false;
            this.dgv_setup_detail.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv_setup_detail.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv_setup_detail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_setup_detail.Location = new System.Drawing.Point(6, 81);
            this.dgv_setup_detail.Name = "dgv_setup_detail";
            this.dgv_setup_detail.ReadOnly = true;
            this.dgv_setup_detail.RowHeadersVisible = false;
            this.dgv_setup_detail.RowHeadersWidth = 82;
            this.dgv_setup_detail.RowTemplate.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.dgv_setup_detail.Size = new System.Drawing.Size(628, 437);
            this.dgv_setup_detail.TabIndex = 17;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.cbAll);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.cbl_sheet);
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(25, 183);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(249, 385);
            this.groupBox2.TabIndex = 17;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = " ";
            // 
            // cbAll
            // 
            this.cbAll.AutoSize = true;
            this.cbAll.Location = new System.Drawing.Point(13, 51);
            this.cbAll.Name = "cbAll";
            this.cbAll.Size = new System.Drawing.Size(57, 24);
            this.cbAll.TabIndex = 16;
            this.cbAll.Text = "ALL";
            this.cbAll.UseVisualStyleBackColor = true;
            this.cbAll.CheckedChanged += new System.EventHandler(this.cbAll_CheckedChanged);
            // 
            // Setup_Spec_SMT
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(932, 733);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.GBInfo);
            this.Name = "Setup_Spec_SMT";
            this.Text = "SETUP_FORMAT";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Setup_Spec_SMT_FormClosed);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.GBInfo.ResumeLayout(false);
            this.GBInfo.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_setup_detail)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.ListBox lst_Item;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckedListBox cbl_sheet;
        private System.Windows.Forms.GroupBox GBInfo;
        private System.Windows.Forms.Button btnRUN;
        private System.Windows.Forms.TextBox txtFormat;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.TextBox txtItemCode;
        private System.Windows.Forms.Label lblItemCode;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox cbAll;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cb_Type;
        private System.Windows.Forms.Button btn_filter;
        private System.Windows.Forms.TextBox txt_ItemCode_filter;
        private System.Windows.Forms.Button btn_all;
        private System.Windows.Forms.DataGridView dgv_setup_detail;
    }
}

