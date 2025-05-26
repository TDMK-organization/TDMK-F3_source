
namespace OK2SHIP_Software
{
    partial class FrmExport
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
            this.splitContainer_main = new System.Windows.Forms.SplitContainer();
            this.label1 = new System.Windows.Forms.Label();
            this.txtItemCode = new System.Windows.Forms.TextBox();
            this.txtLotNo = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cbProcess = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnLoadData = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer_main)).BeginInit();
            this.splitContainer_main.Panel1.SuspendLayout();
            this.splitContainer_main.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer_main
            // 
            this.splitContainer_main.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer_main.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer_main.Location = new System.Drawing.Point(0, 0);
            this.splitContainer_main.Name = "splitContainer_main";
            this.splitContainer_main.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer_main.Panel1
            // 
            this.splitContainer_main.Panel1.Controls.Add(this.btnLoadData);
            this.splitContainer_main.Panel1.Controls.Add(this.label3);
            this.splitContainer_main.Panel1.Controls.Add(this.cbProcess);
            this.splitContainer_main.Panel1.Controls.Add(this.txtLotNo);
            this.splitContainer_main.Panel1.Controls.Add(this.label2);
            this.splitContainer_main.Panel1.Controls.Add(this.txtItemCode);
            this.splitContainer_main.Panel1.Controls.Add(this.label1);
            this.splitContainer_main.Size = new System.Drawing.Size(800, 708);
            this.splitContainer_main.SplitterDistance = 55;
            this.splitContainer_main.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(31, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(52, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "ItemCode";
            // 
            // txtItemCode
            // 
            this.txtItemCode.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.txtItemCode.Location = new System.Drawing.Point(114, 19);
            this.txtItemCode.Name = "txtItemCode";
            this.txtItemCode.Size = new System.Drawing.Size(98, 20);
            this.txtItemCode.TabIndex = 1;
            // 
            // txtLotNo
            // 
            this.txtLotNo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.txtLotNo.Location = new System.Drawing.Point(310, 19);
            this.txtLotNo.Name = "txtLotNo";
            this.txtLotNo.Size = new System.Drawing.Size(98, 20);
            this.txtLotNo.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(243, 22);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(36, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "LotNo";
            // 
            // cbProcess
            // 
            this.cbProcess.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.cbProcess.FormattingEnabled = true;
            this.cbProcess.Items.AddRange(new object[] {
            "FAI",
            "FLEX_BENDING",
            "HEAT_SOAK_AND_BEND",
            "THERMAL_CYCLING_AND_BEND"});
            this.cbProcess.Location = new System.Drawing.Point(515, 18);
            this.cbProcess.Name = "cbProcess";
            this.cbProcess.Size = new System.Drawing.Size(152, 21);
            this.cbProcess.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(439, 21);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(45, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Process";
            // 
            // btnLoadData
            // 
            this.btnLoadData.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.btnLoadData.Location = new System.Drawing.Point(698, 7);
            this.btnLoadData.Name = "btnLoadData";
            this.btnLoadData.Size = new System.Drawing.Size(68, 42);
            this.btnLoadData.TabIndex = 6;
            this.btnLoadData.Text = "Load Data";
            this.btnLoadData.UseVisualStyleBackColor = true;
            this.btnLoadData.Click += new System.EventHandler(this.btnLoadData_Click);
            // 
            // FrmExport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 708);
            this.Controls.Add(this.splitContainer_main);
            this.Name = "FrmExport";
            this.Text = "FrmExport";
            this.Load += new System.EventHandler(this.FrmExport_Load);
            this.splitContainer_main.Panel1.ResumeLayout(false);
            this.splitContainer_main.Panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer_main)).EndInit();
            this.splitContainer_main.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer_main;
        private System.Windows.Forms.Button btnLoadData;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbProcess;
        private System.Windows.Forms.TextBox txtLotNo;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtItemCode;
        private System.Windows.Forms.Label label1;
    }
}