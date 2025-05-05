namespace OK2SHIP_SMT.UserControls
{
    partial class UC_NewFeature
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.slctn_Location = new System.Windows.Forms.SplitContainer();
            this.tbl_Loaction = new System.Windows.Forms.TableLayoutPanel();
            this.tb_locationFolder = new System.Windows.Forms.TextBox();
            this.lb_loactionFile = new System.Windows.Forms.Label();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.btnGetData = new System.Windows.Forms.Button();
            this.btn_runProcess = new System.Windows.Forms.Button();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.slctn_Location)).BeginInit();
            this.slctn_Location.Panel1.SuspendLayout();
            this.slctn_Location.Panel2.SuspendLayout();
            this.slctn_Location.SuspendLayout();
            this.tbl_Loaction.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).BeginInit();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // slctn_Location
            // 
            this.slctn_Location.Dock = System.Windows.Forms.DockStyle.Fill;
            this.slctn_Location.Location = new System.Drawing.Point(0, 0);
            this.slctn_Location.Name = "slctn_Location";
            // 
            // slctn_Location.Panel1
            // 
            this.slctn_Location.Panel1.Controls.Add(this.tbl_Loaction);
            // 
            // slctn_Location.Panel2
            // 
            this.slctn_Location.Panel2.Controls.Add(this.splitContainer2);
            this.slctn_Location.Size = new System.Drawing.Size(1044, 77);
            this.slctn_Location.SplitterDistance = 823;
            this.slctn_Location.TabIndex = 2;
            // 
            // tbl_Loaction
            // 
            this.tbl_Loaction.ColumnCount = 1;
            this.tbl_Loaction.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tbl_Loaction.Controls.Add(this.tb_locationFolder, 0, 1);
            this.tbl_Loaction.Controls.Add(this.lb_loactionFile, 0, 0);
            this.tbl_Loaction.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbl_Loaction.Location = new System.Drawing.Point(0, 0);
            this.tbl_Loaction.Name = "tbl_Loaction";
            this.tbl_Loaction.RowCount = 2;
            this.tbl_Loaction.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 23.28767F));
            this.tbl_Loaction.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 76.71233F));
            this.tbl_Loaction.Size = new System.Drawing.Size(823, 77);
            this.tbl_Loaction.TabIndex = 0;
            // 
            // tb_locationFolder
            // 
            this.tb_locationFolder.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tb_locationFolder.Location = new System.Drawing.Point(3, 20);
            this.tb_locationFolder.Multiline = true;
            this.tb_locationFolder.Name = "tb_locationFolder";
            this.tb_locationFolder.Size = new System.Drawing.Size(817, 54);
            this.tb_locationFolder.TabIndex = 5;
            this.tb_locationFolder.DoubleClick += new System.EventHandler(this.tb_locationFolder_DoubleClick);
            // 
            // lb_loactionFile
            // 
            this.lb_loactionFile.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lb_loactionFile.Location = new System.Drawing.Point(3, 0);
            this.lb_loactionFile.Name = "lb_loactionFile";
            this.lb_loactionFile.Size = new System.Drawing.Size(817, 17);
            this.lb_loactionFile.TabIndex = 0;
            this.lb_loactionFile.Text = "Loaction ";
            this.lb_loactionFile.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // splitContainer2
            // 
            this.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer2.Location = new System.Drawing.Point(0, 0);
            this.splitContainer2.Name = "splitContainer2";
            this.splitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.Controls.Add(this.btnGetData);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.btn_runProcess);
            this.splitContainer2.Size = new System.Drawing.Size(217, 77);
            this.splitContainer2.SplitterDistance = 38;
            this.splitContainer2.TabIndex = 6;
            // 
            // btnGetData
            // 
            this.btnGetData.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnGetData.Location = new System.Drawing.Point(0, 0);
            this.btnGetData.Name = "btnGetData";
            this.btnGetData.Size = new System.Drawing.Size(217, 38);
            this.btnGetData.TabIndex = 5;
            this.btnGetData.Text = "Scan Folder";
            this.btnGetData.UseVisualStyleBackColor = true;
            this.btnGetData.Click += new System.EventHandler(this.btnGetData_Click);
            // 
            // btn_runProcess
            // 
            this.btn_runProcess.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_runProcess.Location = new System.Drawing.Point(0, 0);
            this.btn_runProcess.Name = "btn_runProcess";
            this.btn_runProcess.Size = new System.Drawing.Size(217, 35);
            this.btn_runProcess.TabIndex = 6;
            this.btn_runProcess.Text = "Run Process";
            this.btn_runProcess.UseVisualStyleBackColor = true;
            this.btn_runProcess.Click += new System.EventHandler(this.btn_runProcess_Click);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.slctn_Location);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.dataGridView);
            this.splitContainer1.Size = new System.Drawing.Size(1044, 538);
            this.splitContainer1.SplitterDistance = 77;
            this.splitContainer1.TabIndex = 3;
            // 
            // dataGridView
            // 
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView.Location = new System.Drawing.Point(0, 0);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.Size = new System.Drawing.Size(1044, 457);
            this.dataGridView.TabIndex = 0;
            // 
            // UC_NewFeature
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.splitContainer1);
            this.Name = "UC_NewFeature";
            this.Size = new System.Drawing.Size(1044, 538);
            this.slctn_Location.Panel1.ResumeLayout(false);
            this.slctn_Location.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.slctn_Location)).EndInit();
            this.slctn_Location.ResumeLayout(false);
            this.tbl_Loaction.ResumeLayout(false);
            this.tbl_Loaction.PerformLayout();
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).EndInit();
            this.splitContainer2.ResumeLayout(false);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer slctn_Location;
        private System.Windows.Forms.TableLayoutPanel tbl_Loaction;
        private System.Windows.Forms.TextBox tb_locationFolder;
        private System.Windows.Forms.Label lb_loactionFile;
        private System.Windows.Forms.Button btnGetData;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.SplitContainer splitContainer2;
        private System.Windows.Forms.Button btn_runProcess;
    }
}
