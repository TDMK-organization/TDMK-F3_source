namespace OK2SHIP_Software
{
    partial class FrmMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmMain));
            this.label1 = new System.Windows.Forms.Label();
            this.txtItemCode = new System.Windows.Forms.TextBox();
            this.btnLoadData = new System.Windows.Forms.Button();
            this.tvItemDetails = new System.Windows.Forms.TreeView();
            this.label3 = new System.Windows.Forms.Label();
            this.DGV_Data = new System.Windows.Forms.DataGridView();
            this.tblLayout_Main = new System.Windows.Forms.TableLayoutPanel();
            this.tblLayout_Detail = new System.Windows.Forms.TableLayoutPanel();
            this.label2 = new System.Windows.Forms.Label();
            this.tblLayout_ItemCode = new System.Windows.Forms.TableLayoutPanel();
            this.btnExport = new System.Windows.Forms.Button();
            this.cmsResult = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.tsmPass = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmFail = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmBlank = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmNA = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data)).BeginInit();
            this.tblLayout_Main.SuspendLayout();
            this.tblLayout_Detail.SuspendLayout();
            this.tblLayout_ItemCode.SuspendLayout();
            this.cmsResult.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Location = new System.Drawing.Point(3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(101, 33);
            this.label1.TabIndex = 0;
            this.label1.Text = "Production Code";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtItemCode
            // 
            this.txtItemCode.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtItemCode.Location = new System.Drawing.Point(3, 36);
            this.txtItemCode.Name = "txtItemCode";
            this.txtItemCode.Size = new System.Drawing.Size(101, 20);
            this.txtItemCode.TabIndex = 1;
            this.txtItemCode.Text = "22B0358";
            // 
            // btnLoadData
            // 
            this.btnLoadData.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnLoadData.Location = new System.Drawing.Point(3, 69);
            this.btnLoadData.Name = "btnLoadData";
            this.btnLoadData.Size = new System.Drawing.Size(101, 60);
            this.btnLoadData.TabIndex = 2;
            this.btnLoadData.Text = "Load Data";
            this.btnLoadData.UseVisualStyleBackColor = true;
            this.btnLoadData.Click += new System.EventHandler(this.btnLoadData_Click);
            // 
            // tvItemDetails
            // 
            this.tvItemDetails.CheckBoxes = true;
            this.tvItemDetails.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tvItemDetails.Location = new System.Drawing.Point(3, 168);
            this.tvItemDetails.Name = "tvItemDetails";
            this.tvItemDetails.Size = new System.Drawing.Size(101, 423);
            this.tvItemDetails.TabIndex = 9;
            this.tvItemDetails.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.tvItemDetails_AfterCheck);
            this.tvItemDetails.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvItemDetails_AfterSelect);
            this.tvItemDetails.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.tvItemDetails_NodeMouseDoubleClick);
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.label3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label3.Location = new System.Drawing.Point(3, 132);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(101, 33);
            this.label3.TabIndex = 12;
            this.label3.Text = "LotNo List";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // DGV_Data
            // 
            this.DGV_Data.AllowUserToAddRows = false;
            this.DGV_Data.AllowUserToDeleteRows = false;
            this.DGV_Data.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.DGV_Data.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.DGV_Data.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Data.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_Data.Location = new System.Drawing.Point(3, 69);
            this.DGV_Data.Name = "DGV_Data";
            this.DGV_Data.RowHeadersVisible = false;
            this.DGV_Data.Size = new System.Drawing.Size(1014, 588);
            this.DGV_Data.TabIndex = 15;
            this.DGV_Data.CellMouseDown += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.DGV_temp_CellMouseDown);
            this.DGV_Data.CellToolTipTextNeeded += new System.Windows.Forms.DataGridViewCellToolTipTextNeededEventHandler(this.DGV_Data_CellToolTipTextNeeded);
            this.DGV_Data.ColumnHeaderMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.DGV_Data_ColumnHeaderMouseDoubleClick);
            // 
            // tblLayout_Main
            // 
            this.tblLayout_Main.ColumnCount = 2;
            this.tblLayout_Main.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tblLayout_Main.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 90F));
            this.tblLayout_Main.Controls.Add(this.tblLayout_Detail, 1, 0);
            this.tblLayout_Main.Controls.Add(this.tblLayout_ItemCode, 0, 0);
            this.tblLayout_Main.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tblLayout_Main.Location = new System.Drawing.Point(0, 0);
            this.tblLayout_Main.Name = "tblLayout_Main";
            this.tblLayout_Main.RowCount = 1;
            this.tblLayout_Main.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tblLayout_Main.Size = new System.Drawing.Size(1139, 666);
            this.tblLayout_Main.TabIndex = 16;
            // 
            // tblLayout_Detail
            // 
            this.tblLayout_Detail.ColumnCount = 1;
            this.tblLayout_Detail.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tblLayout_Detail.Controls.Add(this.DGV_Data, 0, 1);
            this.tblLayout_Detail.Controls.Add(this.label2, 0, 0);
            this.tblLayout_Detail.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tblLayout_Detail.Location = new System.Drawing.Point(116, 3);
            this.tblLayout_Detail.Name = "tblLayout_Detail";
            this.tblLayout_Detail.RowCount = 2;
            this.tblLayout_Detail.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tblLayout_Detail.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 90F));
            this.tblLayout_Detail.Size = new System.Drawing.Size(1020, 660);
            this.tblLayout_Detail.TabIndex = 0;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Font = new System.Drawing.Font("Arial", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(3, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(1014, 66);
            this.label2.TabIndex = 16;
            this.label2.Text = "OK2SHIP DashBoard";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tblLayout_ItemCode
            // 
            this.tblLayout_ItemCode.ColumnCount = 1;
            this.tblLayout_ItemCode.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tblLayout_ItemCode.Controls.Add(this.btnLoadData, 0, 2);
            this.tblLayout_ItemCode.Controls.Add(this.label3, 0, 3);
            this.tblLayout_ItemCode.Controls.Add(this.txtItemCode, 0, 1);
            this.tblLayout_ItemCode.Controls.Add(this.label1, 0, 0);
            this.tblLayout_ItemCode.Controls.Add(this.tvItemDetails, 0, 4);
            this.tblLayout_ItemCode.Controls.Add(this.btnExport, 0, 5);
            this.tblLayout_ItemCode.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tblLayout_ItemCode.Location = new System.Drawing.Point(3, 3);
            this.tblLayout_ItemCode.Name = "tblLayout_ItemCode";
            this.tblLayout_ItemCode.RowCount = 6;
            this.tblLayout_ItemCode.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5F));
            this.tblLayout_ItemCode.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5F));
            this.tblLayout_ItemCode.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tblLayout_ItemCode.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5F));
            this.tblLayout_ItemCode.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 65F));
            this.tblLayout_ItemCode.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tblLayout_ItemCode.Size = new System.Drawing.Size(107, 660);
            this.tblLayout_ItemCode.TabIndex = 1;
            // 
            // btnExport
            // 
            this.btnExport.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnExport.Location = new System.Drawing.Point(3, 597);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(101, 60);
            this.btnExport.TabIndex = 13;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // cmsResult
            // 
            this.cmsResult.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmPass,
            this.tsmFail,
            this.tsmBlank,
            this.tsmNA});
            this.cmsResult.Name = "cmsResult";
            this.cmsResult.Size = new System.Drawing.Size(104, 92);
            // 
            // tsmPass
            // 
            this.tsmPass.Name = "tsmPass";
            this.tsmPass.Size = new System.Drawing.Size(103, 22);
            this.tsmPass.Text = "Pass";
            this.tsmPass.Click += new System.EventHandler(this.tsmPass_Click);
            // 
            // tsmFail
            // 
            this.tsmFail.Name = "tsmFail";
            this.tsmFail.Size = new System.Drawing.Size(103, 22);
            this.tsmFail.Text = "Fail";
            this.tsmFail.Click += new System.EventHandler(this.tsmFail_Click);
            // 
            // tsmBlank
            // 
            this.tsmBlank.Name = "tsmBlank";
            this.tsmBlank.Size = new System.Drawing.Size(103, 22);
            this.tsmBlank.Text = "Blank";
            this.tsmBlank.Click += new System.EventHandler(this.tsmBlank_Click);
            // 
            // tsmNA
            // 
            this.tsmNA.Name = "tsmNA";
            this.tsmNA.Size = new System.Drawing.Size(103, 22);
            this.tsmNA.Text = "NA";
            this.tsmNA.Click += new System.EventHandler(this.tsmNA_Click);
            // 
            // FrmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1139, 666);
            this.Controls.Add(this.tblLayout_Main);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmMain";
            this.Text = "OK2SHIP-Overview";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FrmMain_FormClosed);
            this.Load += new System.EventHandler(this.FrmMain_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data)).EndInit();
            this.tblLayout_Main.ResumeLayout(false);
            this.tblLayout_Detail.ResumeLayout(false);
            this.tblLayout_ItemCode.ResumeLayout(false);
            this.tblLayout_ItemCode.PerformLayout();
            this.cmsResult.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtItemCode;
        private System.Windows.Forms.Button btnLoadData;
        private System.Windows.Forms.TreeView tvItemDetails;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DataGridView DGV_Data;
        private System.Windows.Forms.TableLayoutPanel tblLayout_Main;
        private System.Windows.Forms.TableLayoutPanel tblLayout_Detail;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TableLayoutPanel tblLayout_ItemCode;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.ContextMenuStrip cmsResult;
        private System.Windows.Forms.ToolStripMenuItem tsmPass;
        private System.Windows.Forms.ToolStripMenuItem tsmFail;
        private System.Windows.Forms.ToolStripMenuItem tsmBlank;
        private System.Windows.Forms.ToolStripMenuItem tsmNA;
    }
}