namespace OK2SHIP_Software
{
    partial class DE_Data_Info
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
            this.components = new System.ComponentModel.Container();
            this.lblItemCode = new System.Windows.Forms.Label();
            this.txtItemCode = new System.Windows.Forms.TextBox();
            this.RBDeclare = new System.Windows.Forms.RadioButton();
            this.RBMaterials = new System.Windows.Forms.RadioButton();
            this.RBRecycle = new System.Windows.Forms.RadioButton();
            this.GBDatabase = new System.Windows.Forms.GroupBox();
            this.lblItems = new System.Windows.Forms.Label();
            this.lstTable = new System.Windows.Forms.ListBox();
            this.btnExit = new System.Windows.Forms.Button();
            this.btnLoad = new System.Windows.Forms.Button();
            this.DGV_Data = new System.Windows.Forms.DataGridView();
            this.mnuAction = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.mnuAddNew = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuDelete = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuUpdate = new System.Windows.Forms.ToolStripMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.txtMasterList = new System.Windows.Forms.TextBox();
            this.GBDatabase.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data)).BeginInit();
            this.mnuAction.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblItemCode
            // 
            this.lblItemCode.AutoSize = true;
            this.lblItemCode.Location = new System.Drawing.Point(8, 194);
            this.lblItemCode.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblItemCode.Name = "lblItemCode";
            this.lblItemCode.Size = new System.Drawing.Size(76, 18);
            this.lblItemCode.TabIndex = 0;
            this.lblItemCode.Text = "ItemCode";
            // 
            // txtItemCode
            // 
            this.txtItemCode.Location = new System.Drawing.Point(100, 188);
            this.txtItemCode.Margin = new System.Windows.Forms.Padding(4);
            this.txtItemCode.Name = "txtItemCode";
            this.txtItemCode.Size = new System.Drawing.Size(180, 26);
            this.txtItemCode.TabIndex = 1;
            // 
            // RBDeclare
            // 
            this.RBDeclare.AutoSize = true;
            this.RBDeclare.Location = new System.Drawing.Point(40, 47);
            this.RBDeclare.Margin = new System.Windows.Forms.Padding(4);
            this.RBDeclare.Name = "RBDeclare";
            this.RBDeclare.Size = new System.Drawing.Size(106, 22);
            this.RBDeclare.TabIndex = 2;
            this.RBDeclare.TabStop = true;
            this.RBDeclare.Text = "Declaration";
            this.RBDeclare.UseVisualStyleBackColor = true;
            this.RBDeclare.CheckedChanged += new System.EventHandler(this.RBDeclare_CheckedChanged);
            // 
            // RBMaterials
            // 
            this.RBMaterials.AutoSize = true;
            this.RBMaterials.Location = new System.Drawing.Point(40, 97);
            this.RBMaterials.Margin = new System.Windows.Forms.Padding(4);
            this.RBMaterials.Name = "RBMaterials";
            this.RBMaterials.Size = new System.Drawing.Size(90, 22);
            this.RBMaterials.TabIndex = 3;
            this.RBMaterials.TabStop = true;
            this.RBMaterials.Text = "Materials";
            this.RBMaterials.UseVisualStyleBackColor = true;
            this.RBMaterials.CheckedChanged += new System.EventHandler(this.RBMaterials_CheckedChanged);
            // 
            // RBRecycle
            // 
            this.RBRecycle.AutoSize = true;
            this.RBRecycle.Location = new System.Drawing.Point(40, 144);
            this.RBRecycle.Margin = new System.Windows.Forms.Padding(4);
            this.RBRecycle.Name = "RBRecycle";
            this.RBRecycle.Size = new System.Drawing.Size(187, 22);
            this.RBRecycle.TabIndex = 5;
            this.RBRecycle.TabStop = true;
            this.RBRecycle.Text = "Recycle_PGC_Copper";
            this.RBRecycle.UseVisualStyleBackColor = true;
            this.RBRecycle.CheckedChanged += new System.EventHandler(this.RBRecycle_CheckedChanged);
            // 
            // GBDatabase
            // 
            this.GBDatabase.Controls.Add(this.txtMasterList);
            this.GBDatabase.Controls.Add(this.btnUpdate);
            this.GBDatabase.Controls.Add(this.lblItems);
            this.GBDatabase.Controls.Add(this.lstTable);
            this.GBDatabase.Controls.Add(this.btnExit);
            this.GBDatabase.Controls.Add(this.btnLoad);
            this.GBDatabase.Controls.Add(this.RBDeclare);
            this.GBDatabase.Controls.Add(this.txtItemCode);
            this.GBDatabase.Controls.Add(this.lblItemCode);
            this.GBDatabase.Controls.Add(this.RBRecycle);
            this.GBDatabase.Controls.Add(this.RBMaterials);
            this.GBDatabase.Location = new System.Drawing.Point(13, 13);
            this.GBDatabase.Margin = new System.Windows.Forms.Padding(4);
            this.GBDatabase.Name = "GBDatabase";
            this.GBDatabase.Padding = new System.Windows.Forms.Padding(4);
            this.GBDatabase.Size = new System.Drawing.Size(454, 379);
            this.GBDatabase.TabIndex = 6;
            this.GBDatabase.TabStop = false;
            this.GBDatabase.Text = "Items_Database";
            // 
            // lblItems
            // 
            this.lblItems.AutoSize = true;
            this.lblItems.Location = new System.Drawing.Point(286, 22);
            this.lblItems.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblItems.Name = "lblItems";
            this.lblItems.Size = new System.Drawing.Size(93, 18);
            this.lblItems.TabIndex = 10;
            this.lblItems.Text = "Select Items";
            // 
            // lstTable
            // 
            this.lstTable.FormattingEnabled = true;
            this.lstTable.ItemHeight = 18;
            this.lstTable.Location = new System.Drawing.Point(291, 47);
            this.lstTable.Margin = new System.Windows.Forms.Padding(4);
            this.lstTable.Name = "lstTable";
            this.lstTable.Size = new System.Drawing.Size(145, 238);
            this.lstTable.TabIndex = 9;
            this.lstTable.SelectedIndexChanged += new System.EventHandler(this.lstTable_SelectedIndexChanged);
            this.lstTable.RightToLeftChanged += new System.EventHandler(this.lstTable_RightToLeftChanged);
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(166, 244);
            this.btnExit.Margin = new System.Windows.Forms.Padding(4);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(116, 43);
            this.btnExit.TabIndex = 8;
            this.btnExit.Text = "Exit";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnLoad
            // 
            this.btnLoad.Location = new System.Drawing.Point(9, 244);
            this.btnLoad.Margin = new System.Windows.Forms.Padding(4);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(116, 43);
            this.btnLoad.TabIndex = 7;
            this.btnLoad.Text = "Load Data";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // DGV_Data
            // 
            this.DGV_Data.AllowUserToAddRows = false;
            this.DGV_Data.AllowUserToDeleteRows = false;
            this.DGV_Data.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Data.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.DGV_Data.Location = new System.Drawing.Point(482, 60);
            this.DGV_Data.Margin = new System.Windows.Forms.Padding(4);
            this.DGV_Data.Name = "DGV_Data";
            this.DGV_Data.Size = new System.Drawing.Size(834, 332);
            this.DGV_Data.TabIndex = 7;
            this.DGV_Data.CellMouseDown += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.DGV_Data_CellMouseDown);
            this.DGV_Data.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.DGV_Data_ColumnHeaderMouseClick);
            // 
            // mnuAction
            // 
            this.mnuAction.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuAddNew,
            this.mnuDelete,
            this.mnuUpdate});
            this.mnuAction.Name = "mnuAction";
            this.mnuAction.Size = new System.Drawing.Size(124, 70);
            // 
            // mnuAddNew
            // 
            this.mnuAddNew.Name = "mnuAddNew";
            this.mnuAddNew.Size = new System.Drawing.Size(123, 22);
            this.mnuAddNew.Text = "Add New";
            this.mnuAddNew.Click += new System.EventHandler(this.mnuAddNew_Click);
            // 
            // mnuDelete
            // 
            this.mnuDelete.Name = "mnuDelete";
            this.mnuDelete.Size = new System.Drawing.Size(123, 22);
            this.mnuDelete.Text = "Delete";
            this.mnuDelete.Click += new System.EventHandler(this.mnuDelete_Click);
            // 
            // mnuUpdate
            // 
            this.mnuUpdate.Name = "mnuUpdate";
            this.mnuUpdate.Size = new System.Drawing.Size(123, 22);
            this.mnuUpdate.Text = "Update";
            this.mnuUpdate.Click += new System.EventHandler(this.mnuUpdate_Click);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label1.Font = new System.Drawing.Font("Arial", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(482, 22);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(833, 31);
            this.label1.TabIndex = 9;
            this.label1.Text = "Details Data";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnUpdate
            // 
            this.btnUpdate.Location = new System.Drawing.Point(11, 308);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(114, 52);
            this.btnUpdate.TabIndex = 11;
            this.btnUpdate.Text = "Update from Master list";
            this.btnUpdate.UseVisualStyleBackColor = true;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // txtMasterList
            // 
            this.txtMasterList.Location = new System.Drawing.Point(131, 308);
            this.txtMasterList.Multiline = true;
            this.txtMasterList.Name = "txtMasterList";
            this.txtMasterList.Size = new System.Drawing.Size(305, 52);
            this.txtMasterList.TabIndex = 12;
            // 
            // DE_Data_Info
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1336, 423);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.DGV_Data);
            this.Controls.Add(this.GBDatabase);
            this.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "DE_Data_Info";
            this.Text = "DE_Data_Info";
            this.Load += new System.EventHandler(this.DE_Data_Info_Load);
            this.GBDatabase.ResumeLayout(false);
            this.GBDatabase.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data)).EndInit();
            this.mnuAction.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lblItemCode;
        private System.Windows.Forms.TextBox txtItemCode;
        private System.Windows.Forms.RadioButton RBDeclare;
        private System.Windows.Forms.RadioButton RBMaterials;
        private System.Windows.Forms.RadioButton RBRecycle;
        private System.Windows.Forms.GroupBox GBDatabase;
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.DataGridView DGV_Data;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.ListBox lstTable;
        private System.Windows.Forms.Label lblItems;
        private System.Windows.Forms.ContextMenuStrip mnuAction;
        private System.Windows.Forms.ToolStripMenuItem mnuAddNew;
        private System.Windows.Forms.ToolStripMenuItem mnuDelete;
        private System.Windows.Forms.ToolStripMenuItem mnuUpdate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtMasterList;
        private System.Windows.Forms.Button btnUpdate;
    }
}