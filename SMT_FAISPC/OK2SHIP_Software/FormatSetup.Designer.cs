namespace OK2SHIP_Software
{
    partial class FormatSetup
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
            this.btnRUN = new System.Windows.Forms.Button();
            this.RB_Format = new System.Windows.Forms.RadioButton();
            this.RB_IPQC = new System.Windows.Forms.RadioButton();
            this.RB_FAI = new System.Windows.Forms.RadioButton();
            this.txtFormat = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.txtItemCode = new System.Windows.Forms.TextBox();
            this.lblItemCode = new System.Windows.Forms.Label();
            this.lblInfo = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lstSheet = new System.Windows.Forms.ListBox();
            this.DGV_Spec = new System.Windows.Forms.DataGridView();
            this.GBInfo = new System.Windows.Forms.GroupBox();
            this.GB_Type = new System.Windows.Forms.GroupBox();
            this.rbNPI = new System.Windows.Forms.RadioButton();
            this.rbMASS = new System.Windows.Forms.RadioButton();
            this.RBBreakLinks = new System.Windows.Forms.RadioButton();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Spec)).BeginInit();
            this.GBInfo.SuspendLayout();
            this.GB_Type.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnRUN
            // 
            this.btnRUN.Location = new System.Drawing.Point(9, 94);
            this.btnRUN.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnRUN.Name = "btnRUN";
            this.btnRUN.Size = new System.Drawing.Size(70, 35);
            this.btnRUN.TabIndex = 16;
            this.btnRUN.Text = "RUN";
            this.btnRUN.UseVisualStyleBackColor = true;
            this.btnRUN.Click += new System.EventHandler(this.btnRUN_Click);
            // 
            // RB_Format
            // 
            this.RB_Format.AutoSize = true;
            this.RB_Format.Location = new System.Drawing.Point(389, 24);
            this.RB_Format.Name = "RB_Format";
            this.RB_Format.Size = new System.Drawing.Size(103, 17);
            this.RB_Format.TabIndex = 15;
            this.RB_Format.TabStop = true;
            this.RB_Format.Text = "Format Structure";
            this.RB_Format.UseVisualStyleBackColor = true;
            this.RB_Format.CheckedChanged += new System.EventHandler(this.RB_Format_CheckedChanged);
            // 
            // RB_IPQC
            // 
            this.RB_IPQC.AutoSize = true;
            this.RB_IPQC.Location = new System.Drawing.Point(288, 24);
            this.RB_IPQC.Name = "RB_IPQC";
            this.RB_IPQC.Size = new System.Drawing.Size(78, 17);
            this.RB_IPQC.TabIndex = 14;
            this.RB_IPQC.TabStop = true;
            this.RB_IPQC.Text = "IPQC Spec";
            this.RB_IPQC.UseVisualStyleBackColor = true;
            this.RB_IPQC.CheckedChanged += new System.EventHandler(this.RB_IPQC_CheckedChanged);
            // 
            // RB_FAI
            // 
            this.RB_FAI.AutoSize = true;
            this.RB_FAI.Location = new System.Drawing.Point(196, 24);
            this.RB_FAI.Name = "RB_FAI";
            this.RB_FAI.Size = new System.Drawing.Size(69, 17);
            this.RB_FAI.TabIndex = 13;
            this.RB_FAI.TabStop = true;
            this.RB_FAI.Text = "FAI Spec";
            this.RB_FAI.UseVisualStyleBackColor = true;
            this.RB_FAI.CheckedChanged += new System.EventHandler(this.RB_FAI_CheckedChanged);
            // 
            // txtFormat
            // 
            this.txtFormat.Location = new System.Drawing.Point(84, 51);
            this.txtFormat.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtFormat.Multiline = true;
            this.txtFormat.Name = "txtFormat";
            this.txtFormat.Size = new System.Drawing.Size(427, 78);
            this.txtFormat.TabIndex = 10;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(9, 51);
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
            this.txtItemCode.Location = new System.Drawing.Point(87, 22);
            this.txtItemCode.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtItemCode.Name = "txtItemCode";
            this.txtItemCode.Size = new System.Drawing.Size(86, 20);
            this.txtItemCode.TabIndex = 1;
            this.txtItemCode.Text = "22B0413";
            // 
            // lblItemCode
            // 
            this.lblItemCode.AutoSize = true;
            this.lblItemCode.Location = new System.Drawing.Point(12, 26);
            this.lblItemCode.Name = "lblItemCode";
            this.lblItemCode.Size = new System.Drawing.Size(52, 13);
            this.lblItemCode.TabIndex = 0;
            this.lblItemCode.Text = "ItemCode";
            // 
            // lblInfo
            // 
            this.lblInfo.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.lblInfo.Font = new System.Drawing.Font("Arial", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInfo.Location = new System.Drawing.Point(12, 167);
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Size = new System.Drawing.Size(626, 39);
            this.lblInfo.TabIndex = 15;
            this.lblInfo.Text = "FAI/SPC Specification List";
            this.lblInfo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.label1.Font = new System.Drawing.Font("Arial", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(644, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(347, 39);
            this.label1.TabIndex = 14;
            this.label1.Text = "OK2SHIP Report Items";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lstSheet
            // 
            this.lstSheet.FormattingEnabled = true;
            this.lstSheet.Location = new System.Drawing.Point(647, 61);
            this.lstSheet.Name = "lstSheet";
            this.lstSheet.Size = new System.Drawing.Size(344, 511);
            this.lstSheet.TabIndex = 13;
            // 
            // DGV_Spec
            // 
            this.DGV_Spec.AllowUserToAddRows = false;
            this.DGV_Spec.AllowUserToDeleteRows = false;
            this.DGV_Spec.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Spec.Location = new System.Drawing.Point(12, 209);
            this.DGV_Spec.Name = "DGV_Spec";
            this.DGV_Spec.Size = new System.Drawing.Size(626, 368);
            this.DGV_Spec.TabIndex = 12;
            this.DGV_Spec.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DGV_Spec_CellContentClick);
            // 
            // GBInfo
            // 
            this.GBInfo.Controls.Add(this.GB_Type);
            this.GBInfo.Controls.Add(this.RBBreakLinks);
            this.GBInfo.Controls.Add(this.btnRUN);
            this.GBInfo.Controls.Add(this.RB_Format);
            this.GBInfo.Controls.Add(this.RB_IPQC);
            this.GBInfo.Controls.Add(this.RB_FAI);
            this.GBInfo.Controls.Add(this.txtFormat);
            this.GBInfo.Controls.Add(this.btnBrowse);
            this.GBInfo.Controls.Add(this.txtItemCode);
            this.GBInfo.Controls.Add(this.lblItemCode);
            this.GBInfo.Location = new System.Drawing.Point(12, 13);
            this.GBInfo.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.GBInfo.Name = "GBInfo";
            this.GBInfo.Padding = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.GBInfo.Size = new System.Drawing.Size(626, 150);
            this.GBInfo.TabIndex = 11;
            this.GBInfo.TabStop = false;
            this.GBInfo.Text = "Nhập thông tin";
            // 
            // GB_Type
            // 
            this.GB_Type.Controls.Add(this.rbNPI);
            this.GB_Type.Controls.Add(this.rbMASS);
            this.GB_Type.Location = new System.Drawing.Point(517, 47);
            this.GB_Type.Name = "GB_Type";
            this.GB_Type.Size = new System.Drawing.Size(103, 82);
            this.GB_Type.TabIndex = 18;
            this.GB_Type.TabStop = false;
            this.GB_Type.Text = "Format Type";
            // 
            // rbNPI
            // 
            this.rbNPI.AutoSize = true;
            this.rbNPI.Location = new System.Drawing.Point(12, 56);
            this.rbNPI.Name = "rbNPI";
            this.rbNPI.Size = new System.Drawing.Size(43, 17);
            this.rbNPI.TabIndex = 1;
            this.rbNPI.Text = "NPI";
            this.rbNPI.UseVisualStyleBackColor = true;
            // 
            // rbMASS
            // 
            this.rbMASS.AutoSize = true;
            this.rbMASS.Checked = true;
            this.rbMASS.Location = new System.Drawing.Point(12, 22);
            this.rbMASS.Name = "rbMASS";
            this.rbMASS.Size = new System.Drawing.Size(55, 17);
            this.rbMASS.TabIndex = 0;
            this.rbMASS.TabStop = true;
            this.rbMASS.Text = "MASS";
            this.rbMASS.UseVisualStyleBackColor = true;
            // 
            // RBBreakLinks
            // 
            this.RBBreakLinks.AutoSize = true;
            this.RBBreakLinks.Location = new System.Drawing.Point(515, 24);
            this.RBBreakLinks.Name = "RBBreakLinks";
            this.RBBreakLinks.Size = new System.Drawing.Size(93, 17);
            this.RBBreakLinks.TabIndex = 17;
            this.RBBreakLinks.TabStop = true;
            this.RBBreakLinks.Text = "Remove Links";
            this.RBBreakLinks.UseVisualStyleBackColor = true;
            this.RBBreakLinks.CheckedChanged += new System.EventHandler(this.RBBreakLinks_CheckedChanged);
            // 
            // FormatSetup
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1012, 584);
            this.Controls.Add(this.lblInfo);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lstSheet);
            this.Controls.Add(this.DGV_Spec);
            this.Controls.Add(this.GBInfo);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormatSetup";
            this.Text = "FormatSetup";
            this.Load += new System.EventHandler(this.FormatSetup_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Spec)).EndInit();
            this.GBInfo.ResumeLayout(false);
            this.GBInfo.PerformLayout();
            this.GB_Type.ResumeLayout(false);
            this.GB_Type.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnRUN;
        private System.Windows.Forms.RadioButton RB_Format;
        private System.Windows.Forms.RadioButton RB_IPQC;
        private System.Windows.Forms.RadioButton RB_FAI;
        private System.Windows.Forms.TextBox txtFormat;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.TextBox txtItemCode;
        private System.Windows.Forms.Label lblItemCode;
        private System.Windows.Forms.Label lblInfo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox lstSheet;
        private System.Windows.Forms.DataGridView DGV_Spec;
        private System.Windows.Forms.GroupBox GBInfo;
        private System.Windows.Forms.RadioButton RBBreakLinks;
        private System.Windows.Forms.GroupBox GB_Type;
        private System.Windows.Forms.RadioButton rbNPI;
        private System.Windows.Forms.RadioButton rbMASS;
    }
}