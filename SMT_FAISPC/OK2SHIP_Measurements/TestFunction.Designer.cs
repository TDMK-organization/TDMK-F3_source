namespace OK2SHIP_Measurements
{
    partial class TestFunction
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.txtFile = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.lstSTT = new System.Windows.Forms.ListBox();
            this.lstSa = new System.Windows.Forms.ListBox();
            this.lstSq = new System.Windows.Forms.ListBox();
            this.numMeasLoc = new System.Windows.Forms.NumericUpDown();
            this.lstSq2 = new System.Windows.Forms.ListBox();
            this.lstSa2 = new System.Windows.Forms.ListBox();
            this.btnTest = new System.Windows.Forms.Button();
            this.txtInput = new System.Windows.Forms.TextBox();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.DGV_Data = new System.Windows.Forms.DataGridView();
            this.DGV_CPK = new System.Windows.Forms.DataGridView();
            this.DGV_DataView = new System.Windows.Forms.DataGridView();
            this.DGV_Data_Plus = new System.Windows.Forms.DataGridView();
            this.DGV_SpecView = new System.Windows.Forms.DataGridView();
            this.btnLoadSpec = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.tsmloadSpec = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmTestFunction = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmTest2 = new System.Windows.Forms.ToolStripMenuItem();
            this.RBNormal = new System.Windows.Forms.RadioButton();
            this.RBEdit = new System.Windows.Forms.RadioButton();
            this.txtItemCode = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtLotNo = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.numMeasLoc)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_CPK)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_DataView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data_Plus)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_SpecView)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtFile
            // 
            this.txtFile.Location = new System.Drawing.Point(106, 28);
            this.txtFile.Multiline = true;
            this.txtFile.Name = "txtFile";
            this.txtFile.Size = new System.Drawing.Size(392, 49);
            this.txtFile.TabIndex = 0;
            this.txtFile.Text = "D:\\Customer Projects\\SEEV\\SMT Project\\Source Code\\TestAreas\\OK2SHIP Automation Sy" +
    "stem\\NPI\\Ok2ship_ D47 F-cam (SC-G) POR - P2 (C4.0)_7S0302.xlsm";
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(13, 27);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 54);
            this.btnBrowse.TabIndex = 1;
            this.btnBrowse.Text = "Browse...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // lstSTT
            // 
            this.lstSTT.FormattingEnabled = true;
            this.lstSTT.Location = new System.Drawing.Point(13, 244);
            this.lstSTT.Name = "lstSTT";
            this.lstSTT.Size = new System.Drawing.Size(97, 108);
            this.lstSTT.TabIndex = 2;
            this.lstSTT.SelectedIndexChanged += new System.EventHandler(this.lstSTT_SelectedIndexChanged);
            // 
            // lstSa
            // 
            this.lstSa.FormattingEnabled = true;
            this.lstSa.Location = new System.Drawing.Point(13, 129);
            this.lstSa.Name = "lstSa";
            this.lstSa.Size = new System.Drawing.Size(97, 108);
            this.lstSa.TabIndex = 3;
            // 
            // lstSq
            // 
            this.lstSq.FormattingEnabled = true;
            this.lstSq.Location = new System.Drawing.Point(12, 358);
            this.lstSq.Name = "lstSq";
            this.lstSq.Size = new System.Drawing.Size(97, 108);
            this.lstSq.TabIndex = 4;
            // 
            // numMeasLoc
            // 
            this.numMeasLoc.Location = new System.Drawing.Point(344, 83);
            this.numMeasLoc.Maximum = new decimal(new int[] {
            3,
            0,
            0,
            0});
            this.numMeasLoc.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numMeasLoc.Name = "numMeasLoc";
            this.numMeasLoc.Size = new System.Drawing.Size(46, 20);
            this.numMeasLoc.TabIndex = 5;
            this.numMeasLoc.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // lstSq2
            // 
            this.lstSq2.FormattingEnabled = true;
            this.lstSq2.Location = new System.Drawing.Point(12, 586);
            this.lstSq2.Name = "lstSq2";
            this.lstSq2.Size = new System.Drawing.Size(97, 108);
            this.lstSq2.TabIndex = 7;
            // 
            // lstSa2
            // 
            this.lstSa2.FormattingEnabled = true;
            this.lstSa2.Location = new System.Drawing.Point(12, 472);
            this.lstSa2.Name = "lstSa2";
            this.lstSa2.Size = new System.Drawing.Size(97, 108);
            this.lstSa2.TabIndex = 6;
            // 
            // btnTest
            // 
            this.btnTest.Location = new System.Drawing.Point(13, 82);
            this.btnTest.Name = "btnTest";
            this.btnTest.Size = new System.Drawing.Size(87, 41);
            this.btnTest.TabIndex = 8;
            this.btnTest.Text = "Test Function";
            this.btnTest.UseVisualStyleBackColor = true;
            this.btnTest.Click += new System.EventHandler(this.btnTest_Click);
            // 
            // txtInput
            // 
            this.txtInput.Location = new System.Drawing.Point(106, 82);
            this.txtInput.Name = "txtInput";
            this.txtInput.Size = new System.Drawing.Size(90, 20);
            this.txtInput.TabIndex = 9;
            // 
            // txtOutput
            // 
            this.txtOutput.Location = new System.Drawing.Point(223, 82);
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.Size = new System.Drawing.Size(90, 20);
            this.txtOutput.TabIndex = 10;
            // 
            // DGV_Data
            // 
            this.DGV_Data.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Data.Location = new System.Drawing.Point(129, 122);
            this.DGV_Data.Name = "DGV_Data";
            this.DGV_Data.Size = new System.Drawing.Size(737, 184);
            this.DGV_Data.TabIndex = 12;
            // 
            // DGV_CPK
            // 
            this.DGV_CPK.AllowUserToAddRows = false;
            this.DGV_CPK.AllowUserToDeleteRows = false;
            this.DGV_CPK.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_CPK.Location = new System.Drawing.Point(896, 203);
            this.DGV_CPK.Name = "DGV_CPK";
            this.DGV_CPK.ReadOnly = true;
            this.DGV_CPK.Size = new System.Drawing.Size(399, 103);
            this.DGV_CPK.TabIndex = 13;
            // 
            // DGV_DataView
            // 
            this.DGV_DataView.AllowUserToAddRows = false;
            this.DGV_DataView.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.DGV_DataView.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.DGV_DataView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.Format = "N3";
            dataGridViewCellStyle2.NullValue = null;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.DGV_DataView.DefaultCellStyle = dataGridViewCellStyle2;
            this.DGV_DataView.Location = new System.Drawing.Point(129, 313);
            this.DGV_DataView.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.DGV_DataView.Name = "DGV_DataView";
            this.DGV_DataView.Size = new System.Drawing.Size(737, 389);
            this.DGV_DataView.TabIndex = 14;
            this.DGV_DataView.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.DGV_DataView_CellFormatting);
            this.DGV_DataView.ColumnHeaderMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.DGV_DataView_ColumnHeaderMouseDoubleClick);
            // 
            // DGV_Data_Plus
            // 
            this.DGV_Data_Plus.AllowUserToAddRows = false;
            this.DGV_Data_Plus.AllowUserToDeleteRows = false;
            this.DGV_Data_Plus.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Data_Plus.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.DGV_Data_Plus.Location = new System.Drawing.Point(872, 318);
            this.DGV_Data_Plus.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.DGV_Data_Plus.Name = "DGV_Data_Plus";
            this.DGV_Data_Plus.Size = new System.Drawing.Size(423, 384);
            this.DGV_Data_Plus.TabIndex = 15;
            // 
            // DGV_SpecView
            // 
            this.DGV_SpecView.AllowUserToAddRows = false;
            this.DGV_SpecView.AllowUserToDeleteRows = false;
            this.DGV_SpecView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_SpecView.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.DGV_SpecView.Location = new System.Drawing.Point(896, 42);
            this.DGV_SpecView.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.DGV_SpecView.Name = "DGV_SpecView";
            this.DGV_SpecView.Size = new System.Drawing.Size(399, 154);
            this.DGV_SpecView.TabIndex = 16;
            // 
            // btnLoadSpec
            // 
            this.btnLoadSpec.Location = new System.Drawing.Point(411, 82);
            this.btnLoadSpec.Name = "btnLoadSpec";
            this.btnLoadSpec.Size = new System.Drawing.Size(87, 30);
            this.btnLoadSpec.TabIndex = 17;
            this.btnLoadSpec.Text = "Load Spec";
            this.btnLoadSpec.UseVisualStyleBackColor = true;
            this.btnLoadSpec.Click += new System.EventHandler(this.btnLoadSpec_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmloadSpec,
            this.tsmTestFunction,
            this.tsmTest2});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1307, 24);
            this.menuStrip1.TabIndex = 18;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // tsmloadSpec
            // 
            this.tsmloadSpec.Name = "tsmloadSpec";
            this.tsmloadSpec.Size = new System.Drawing.Size(73, 20);
            this.tsmloadSpec.Text = "Load Spec";
            this.tsmloadSpec.Click += new System.EventHandler(this.tsmloadSpec_Click);
            // 
            // tsmTestFunction
            // 
            this.tsmTestFunction.Name = "tsmTestFunction";
            this.tsmTestFunction.Size = new System.Drawing.Size(96, 20);
            this.tsmTestFunction.Text = "GetColumsList";
            this.tsmTestFunction.Click += new System.EventHandler(this.tsmTestFunction_Click);
            // 
            // tsmTest2
            // 
            this.tsmTest2.Name = "tsmTest2";
            this.tsmTest2.Size = new System.Drawing.Size(45, 20);
            this.tsmTest2.Text = "Test2";
            this.tsmTest2.Click += new System.EventHandler(this.tsmTest2_Click);
            // 
            // RBNormal
            // 
            this.RBNormal.AutoSize = true;
            this.RBNormal.Location = new System.Drawing.Point(517, 29);
            this.RBNormal.Name = "RBNormal";
            this.RBNormal.Size = new System.Drawing.Size(58, 17);
            this.RBNormal.TabIndex = 19;
            this.RBNormal.TabStop = true;
            this.RBNormal.Text = "Normal";
            this.RBNormal.UseVisualStyleBackColor = true;
            this.RBNormal.CheckedChanged += new System.EventHandler(this.RBNormal_CheckedChanged);
            // 
            // RBEdit
            // 
            this.RBEdit.AutoSize = true;
            this.RBEdit.Location = new System.Drawing.Point(517, 64);
            this.RBEdit.Name = "RBEdit";
            this.RBEdit.Size = new System.Drawing.Size(43, 17);
            this.RBEdit.TabIndex = 20;
            this.RBEdit.TabStop = true;
            this.RBEdit.Text = "Edit";
            this.RBEdit.UseVisualStyleBackColor = true;
            // 
            // txtItemCode
            // 
            this.txtItemCode.Location = new System.Drawing.Point(672, 29);
            this.txtItemCode.Name = "txtItemCode";
            this.txtItemCode.Size = new System.Drawing.Size(90, 20);
            this.txtItemCode.TabIndex = 21;
            this.txtItemCode.Text = "7S0011";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(614, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(52, 13);
            this.label1.TabIndex = 22;
            this.label1.Text = "ItemCode";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(614, 68);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(36, 13);
            this.label2.TabIndex = 24;
            this.label2.Text = "LotNo";
            // 
            // txtLotNo
            // 
            this.txtLotNo.Location = new System.Drawing.Point(672, 64);
            this.txtLotNo.Name = "txtLotNo";
            this.txtLotNo.Size = new System.Drawing.Size(90, 20);
            this.txtLotNo.TabIndex = 23;
            this.txtLotNo.Text = "00003";
            // 
            // TestFunction
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1307, 715);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtLotNo);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtItemCode);
            this.Controls.Add(this.RBEdit);
            this.Controls.Add(this.RBNormal);
            this.Controls.Add(this.btnLoadSpec);
            this.Controls.Add(this.DGV_SpecView);
            this.Controls.Add(this.DGV_Data_Plus);
            this.Controls.Add(this.DGV_DataView);
            this.Controls.Add(this.DGV_CPK);
            this.Controls.Add(this.DGV_Data);
            this.Controls.Add(this.txtOutput);
            this.Controls.Add(this.txtInput);
            this.Controls.Add(this.btnTest);
            this.Controls.Add(this.lstSq2);
            this.Controls.Add(this.lstSa2);
            this.Controls.Add(this.numMeasLoc);
            this.Controls.Add(this.lstSq);
            this.Controls.Add(this.lstSa);
            this.Controls.Add(this.lstSTT);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.txtFile);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "TestFunction";
            this.Text = "TestFunction";
            this.Load += new System.EventHandler(this.TestFunction_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numMeasLoc)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_CPK)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_DataView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data_Plus)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_SpecView)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtFile;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.ListBox lstSTT;
        private System.Windows.Forms.ListBox lstSa;
        private System.Windows.Forms.ListBox lstSq;
        private System.Windows.Forms.NumericUpDown numMeasLoc;
        private System.Windows.Forms.ListBox lstSq2;
        private System.Windows.Forms.ListBox lstSa2;
        private System.Windows.Forms.Button btnTest;
        private System.Windows.Forms.TextBox txtInput;
        private System.Windows.Forms.TextBox txtOutput;
        private System.Windows.Forms.DataGridView DGV_Data;
        private System.Windows.Forms.DataGridView DGV_CPK;
        private System.Windows.Forms.DataGridView DGV_DataView;
        private System.Windows.Forms.DataGridView DGV_Data_Plus;
        private System.Windows.Forms.DataGridView DGV_SpecView;
        private System.Windows.Forms.Button btnLoadSpec;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem tsmloadSpec;
        private System.Windows.Forms.ToolStripMenuItem tsmTestFunction;
        private System.Windows.Forms.ToolStripMenuItem tsmTest2;
        private System.Windows.Forms.RadioButton RBNormal;
        private System.Windows.Forms.RadioButton RBEdit;
        private System.Windows.Forms.TextBox txtItemCode;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtLotNo;
    }
}