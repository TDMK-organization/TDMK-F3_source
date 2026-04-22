namespace OK2SHIP_Measurements
{
    partial class OK2SHIP_Measurements
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OK2SHIP_Measurements));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.tsmCB_Items = new System.Windows.Forms.ToolStripComboBox();
            this.tsmCB_Mode = new System.Windows.Forms.ToolStripComboBox();
            this.tsmDevice = new System.Windows.Forms.ToolStripComboBox();
            this.tsmLoad = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmReset = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmExit = new System.Windows.Forms.ToolStripMenuItem();
            this.pMain = new System.Windows.Forms.Panel();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmCB_Items,
            this.tsmCB_Mode,
            this.tsmDevice,
            this.tsmLoad,
            this.tsmReset,
            this.tsmExit});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(9, 3, 0, 3);
            this.menuStrip1.Size = new System.Drawing.Size(1264, 33);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // tsmCB_Items
            // 
            this.tsmCB_Items.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tsmCB_Items.Items.AddRange(new object[] {
            "FAI/SPC",
            "IPQC",
            "PTH_Diameter"});
            this.tsmCB_Items.Name = "tsmCB_Items";
            this.tsmCB_Items.Size = new System.Drawing.Size(150, 27);
            this.tsmCB_Items.SelectedIndexChanged += new System.EventHandler(this.tsmCB_Items_SelectedIndexChanged);
            this.tsmCB_Items.Click += new System.EventHandler(this.tsmCB_Items_Click);
            // 
            // tsmCB_Mode
            // 
            this.tsmCB_Mode.AutoToolTip = true;
            this.tsmCB_Mode.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tsmCB_Mode.Items.AddRange(new object[] {
            "Auto",
            "Manual"});
            this.tsmCB_Mode.Name = "tsmCB_Mode";
            this.tsmCB_Mode.Size = new System.Drawing.Size(100, 27);
            this.tsmCB_Mode.ToolTipText = "Select Mode";
            this.tsmCB_Mode.SelectedIndexChanged += new System.EventHandler(this.tsmCB_Mode_SelectedIndexChanged);
            // 
            // tsmDevice
            // 
            this.tsmDevice.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tsmDevice.Name = "tsmDevice";
            this.tsmDevice.Size = new System.Drawing.Size(150, 27);
            this.tsmDevice.SelectedIndexChanged += new System.EventHandler(this.tsmDevice_SelectedIndexChanged);
            // 
            // tsmLoad
            // 
            this.tsmLoad.Name = "tsmLoad";
            this.tsmLoad.Size = new System.Drawing.Size(56, 27);
            this.tsmLoad.Text = "Load";
            this.tsmLoad.Click += new System.EventHandler(this.tsmLoad_Click);
            // 
            // tsmReset
            // 
            this.tsmReset.Name = "tsmReset";
            this.tsmReset.Size = new System.Drawing.Size(61, 27);
            this.tsmReset.Text = "Reset";
            this.tsmReset.Click += new System.EventHandler(this.tsmReset_Click);
            // 
            // tsmExit
            // 
            this.tsmExit.Name = "tsmExit";
            this.tsmExit.Size = new System.Drawing.Size(46, 27);
            this.tsmExit.Text = "Exit";
            this.tsmExit.Click += new System.EventHandler(this.tsmExit_Click);
            // 
            // pMain
            // 
            this.pMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pMain.Location = new System.Drawing.Point(0, 33);
            this.pMain.Name = "pMain";
            this.pMain.Size = new System.Drawing.Size(1264, 728);
            this.pMain.TabIndex = 1;
            // 
            // OK2SHIP_Measurements
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1264, 761);
            this.Controls.Add(this.pMain);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "OK2SHIP_Measurements";
            this.Text = "OK2SHIP_Measurements";
            this.Load += new System.EventHandler(this.OK2SHIP_Measurements_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripComboBox tsmCB_Items;
        private System.Windows.Forms.ToolStripMenuItem tsmLoad;
        private System.Windows.Forms.ToolStripMenuItem tsmExit;
        private System.Windows.Forms.ToolStripComboBox tsmCB_Mode;
        private System.Windows.Forms.ToolStripMenuItem tsmReset;
        private System.Windows.Forms.Panel pMain;
        private System.Windows.Forms.ToolStripComboBox tsmDevice;
    }
}

