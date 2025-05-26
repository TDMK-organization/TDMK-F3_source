
namespace NET_SPEC_Setup
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.txtItemCode = new System.Windows.Forms.TextBox();
            this.lblItemCode = new System.Windows.Forms.Label();
            this.DGV_Data = new System.Windows.Forms.DataGridView();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.tsmLoad_data = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmExport = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmImport = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmSave = new System.Windows.Forms.ToolStripMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.cbProcess = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtItemCode
            // 
            this.txtItemCode.Location = new System.Drawing.Point(100, 37);
            this.txtItemCode.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtItemCode.Name = "txtItemCode";
            this.txtItemCode.Size = new System.Drawing.Size(86, 20);
            this.txtItemCode.TabIndex = 3;
            this.txtItemCode.Text = "22B0500";
            // 
            // lblItemCode
            // 
            this.lblItemCode.AutoSize = true;
            this.lblItemCode.Location = new System.Drawing.Point(28, 40);
            this.lblItemCode.Name = "lblItemCode";
            this.lblItemCode.Size = new System.Drawing.Size(52, 13);
            this.lblItemCode.TabIndex = 2;
            this.lblItemCode.Text = "ItemCode";
            // 
            // DGV_Data
            // 
            this.DGV_Data.AllowUserToAddRows = false;
            this.DGV_Data.AllowUserToDeleteRows = false;
            this.DGV_Data.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Data.Location = new System.Drawing.Point(12, 64);
            this.DGV_Data.Name = "DGV_Data";
            this.DGV_Data.Size = new System.Drawing.Size(776, 504);
            this.DGV_Data.TabIndex = 4;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmLoad_data,
            this.tsmExport,
            this.tsmImport,
            this.tsmSave});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(800, 24);
            this.menuStrip1.TabIndex = 6;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // tsmLoad_data
            // 
            this.tsmLoad_data.Name = "tsmLoad_data";
            this.tsmLoad_data.Size = new System.Drawing.Size(72, 20);
            this.tsmLoad_data.Text = "Load Data";
            this.tsmLoad_data.Click += new System.EventHandler(this.tsmLoad_data_Click);
            // 
            // tsmExport
            // 
            this.tsmExport.Name = "tsmExport";
            this.tsmExport.Size = new System.Drawing.Size(97, 20);
            this.tsmExport.Text = "Export to Excel";
            this.tsmExport.Click += new System.EventHandler(this.tsmExport_Click);
            // 
            // tsmImport
            // 
            this.tsmImport.Name = "tsmImport";
            this.tsmImport.Size = new System.Drawing.Size(114, 20);
            this.tsmImport.Text = "Import from Excel";
            this.tsmImport.Click += new System.EventHandler(this.tsmImport_Click);
            // 
            // tsmSave
            // 
            this.tsmSave.Name = "tsmSave";
            this.tsmSave.Size = new System.Drawing.Size(108, 20);
            this.tsmSave.Text = "Save to Database";
            this.tsmSave.Click += new System.EventHandler(this.tsmSave_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(222, 40);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(45, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Process";
            // 
            // cbProcess
            // 
            this.cbProcess.FormattingEnabled = true;
            this.cbProcess.Location = new System.Drawing.Point(273, 36);
            this.cbProcess.Name = "cbProcess";
            this.cbProcess.Size = new System.Drawing.Size(257, 21);
            this.cbProcess.TabIndex = 8;
            this.cbProcess.SelectedIndexChanged += new System.EventHandler(this.cbProcess_SelectedIndexChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 590);
            this.Controls.Add(this.cbProcess);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.DGV_Data);
            this.Controls.Add(this.txtItemCode);
            this.Controls.Add(this.lblItemCode);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtItemCode;
        private System.Windows.Forms.Label lblItemCode;
        private System.Windows.Forms.DataGridView DGV_Data;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem tsmLoad_data;
        private System.Windows.Forms.ToolStripMenuItem tsmExport;
        private System.Windows.Forms.ToolStripMenuItem tsmImport;
        private System.Windows.Forms.ToolStripMenuItem tsmSave;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbProcess;
    }
}

