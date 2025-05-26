
namespace OK2SHIP_Software
{
    partial class FrmAnalysis
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
            this.splMain = new System.Windows.Forms.SplitContainer();
            this.txtItemCode = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtLotNo = new System.Windows.Forms.TextBox();
            this.cbProcess = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnLoad = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.splMain)).BeginInit();
            this.splMain.Panel1.SuspendLayout();
            this.splMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // splMain
            // 
            this.splMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splMain.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splMain.Location = new System.Drawing.Point(0, 0);
            this.splMain.Name = "splMain";
            this.splMain.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splMain.Panel1
            // 
            this.splMain.Panel1.Controls.Add(this.btnExport);
            this.splMain.Panel1.Controls.Add(this.btnLoad);
            this.splMain.Panel1.Controls.Add(this.label3);
            this.splMain.Panel1.Controls.Add(this.cbProcess);
            this.splMain.Panel1.Controls.Add(this.label2);
            this.splMain.Panel1.Controls.Add(this.txtLotNo);
            this.splMain.Panel1.Controls.Add(this.label1);
            this.splMain.Panel1.Controls.Add(this.txtItemCode);
            this.splMain.Size = new System.Drawing.Size(1184, 861);
            this.splMain.SplitterDistance = 62;
            this.splMain.TabIndex = 0;
            // 
            // txtItemCode
            // 
            this.txtItemCode.Location = new System.Drawing.Point(87, 21);
            this.txtItemCode.Name = "txtItemCode";
            this.txtItemCode.Size = new System.Drawing.Size(100, 20);
            this.txtItemCode.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(52, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "ItemCode";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(209, 25);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(36, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "LotNo";
            // 
            // txtLotNo
            // 
            this.txtLotNo.Location = new System.Drawing.Point(267, 21);
            this.txtLotNo.Name = "txtLotNo";
            this.txtLotNo.Size = new System.Drawing.Size(100, 20);
            this.txtLotNo.TabIndex = 2;
            // 
            // cbProcess
            // 
            this.cbProcess.FormattingEnabled = true;
            this.cbProcess.Items.AddRange(new object[] {
            "FAI",
            "Dimension without parentheses",
            "Dimension with parentheses",
            "Cross section",
            "GAP Connector",
            "Peel Test",
            "(Mating) Pull Test",
            "Shear test",
            "IQC Liner peeling (Coupon)",
            "IQC PSA peeling (Coupon) ",
            "PSA peel test (On product)",
            "Flex bending",
            "Thermal Cycling & bending",
            "Heat Soak & bending",
            "ACF"});
            this.cbProcess.Location = new System.Drawing.Point(456, 21);
            this.cbProcess.Name = "cbProcess";
            this.cbProcess.Size = new System.Drawing.Size(244, 21);
            this.cbProcess.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(389, 25);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(45, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Process";
            // 
            // btnLoad
            // 
            this.btnLoad.Location = new System.Drawing.Point(725, 15);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(84, 31);
            this.btnLoad.TabIndex = 6;
            this.btnLoad.Text = "Load Data";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(831, 15);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(84, 31);
            this.btnExport.TabIndex = 7;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // FrmAnalysis
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1184, 861);
            this.Controls.Add(this.splMain);
            this.Name = "FrmAnalysis";
            this.Text = "FrmAnalysis";
            this.Load += new System.EventHandler(this.FrmAnalysis_Load);
            this.splMain.Panel1.ResumeLayout(false);
            this.splMain.Panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splMain)).EndInit();
            this.splMain.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splMain;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbProcess;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtLotNo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtItemCode;
    }
}