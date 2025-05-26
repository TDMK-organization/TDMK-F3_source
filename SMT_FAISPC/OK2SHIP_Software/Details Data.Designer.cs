namespace OK2SHIP_Software
{
    partial class Details_Data
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
            this.tblMain = new System.Windows.Forms.TableLayoutPanel();
            this.lblDetail = new System.Windows.Forms.Label();
            this.tbldetail = new System.Windows.Forms.TableLayoutPanel();
            this.tvItem_Details = new System.Windows.Forms.TreeView();
            this.tblRight = new System.Windows.Forms.TableLayoutPanel();
            this.DGV_Results = new System.Windows.Forms.DataGridView();
            this.btnConfirm = new System.Windows.Forms.Button();
            this.lblInfo = new System.Windows.Forms.Label();
            this.tblMain.SuspendLayout();
            this.tbldetail.SuspendLayout();
            this.tblRight.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Results)).BeginInit();
            this.SuspendLayout();
            // 
            // tblMain
            // 
            this.tblMain.ColumnCount = 1;
            this.tblMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tblMain.Controls.Add(this.lblDetail, 0, 0);
            this.tblMain.Controls.Add(this.tbldetail, 0, 1);
            this.tblMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tblMain.Location = new System.Drawing.Point(0, 0);
            this.tblMain.Name = "tblMain";
            this.tblMain.RowCount = 2;
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 90F));
            this.tblMain.Size = new System.Drawing.Size(1252, 655);
            this.tblMain.TabIndex = 0;
            // 
            // lblDetail
            // 
            this.lblDetail.BackColor = System.Drawing.Color.CornflowerBlue;
            this.lblDetail.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblDetail.Font = new System.Drawing.Font("Arial", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDetail.Location = new System.Drawing.Point(3, 0);
            this.lblDetail.Name = "lblDetail";
            this.lblDetail.Size = new System.Drawing.Size(1246, 65);
            this.lblDetail.TabIndex = 1;
            this.lblDetail.Text = "OK2SHIP Report Data Contents";
            this.lblDetail.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tbldetail
            // 
            this.tbldetail.ColumnCount = 2;
            this.tbldetail.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30F));
            this.tbldetail.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 70F));
            this.tbldetail.Controls.Add(this.tvItem_Details, 0, 0);
            this.tbldetail.Controls.Add(this.tblRight, 1, 0);
            this.tbldetail.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbldetail.Location = new System.Drawing.Point(3, 68);
            this.tbldetail.Name = "tbldetail";
            this.tbldetail.RowCount = 1;
            this.tbldetail.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tbldetail.Size = new System.Drawing.Size(1246, 584);
            this.tbldetail.TabIndex = 2;
            // 
            // tvItem_Details
            // 
            this.tvItem_Details.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tvItem_Details.Location = new System.Drawing.Point(3, 3);
            this.tvItem_Details.Name = "tvItem_Details";
            this.tvItem_Details.Size = new System.Drawing.Size(367, 578);
            this.tvItem_Details.TabIndex = 0;
            this.tvItem_Details.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvItem_Details_AfterSelect);
            this.tvItem_Details.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.tvItem_Details_NodeMouseDoubleClick);
            // 
            // tblRight
            // 
            this.tblRight.ColumnCount = 1;
            this.tblRight.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tblRight.Controls.Add(this.DGV_Results, 0, 1);
            this.tblRight.Controls.Add(this.btnConfirm, 0, 2);
            this.tblRight.Controls.Add(this.lblInfo, 0, 0);
            this.tblRight.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tblRight.Location = new System.Drawing.Point(376, 3);
            this.tblRight.Name = "tblRight";
            this.tblRight.RowCount = 3;
            this.tblRight.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tblRight.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 80F));
            this.tblRight.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tblRight.Size = new System.Drawing.Size(867, 578);
            this.tblRight.TabIndex = 1;
            // 
            // DGV_Results
            // 
            this.DGV_Results.AllowUserToAddRows = false;
            this.DGV_Results.AllowUserToDeleteRows = false;
            this.DGV_Results.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Results.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_Results.Location = new System.Drawing.Point(3, 60);
            this.DGV_Results.Name = "DGV_Results";
            this.DGV_Results.Size = new System.Drawing.Size(861, 456);
            this.DGV_Results.TabIndex = 5;
            // 
            // btnConfirm
            // 
            this.btnConfirm.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnConfirm.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnConfirm.Location = new System.Drawing.Point(3, 522);
            this.btnConfirm.Name = "btnConfirm";
            this.btnConfirm.Size = new System.Drawing.Size(861, 53);
            this.btnConfirm.TabIndex = 2;
            this.btnConfirm.Text = "Confirm";
            this.btnConfirm.UseVisualStyleBackColor = true;
            this.btnConfirm.Click += new System.EventHandler(this.btnConfirm_Click);
            // 
            // lblInfo
            // 
            this.lblInfo.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.lblInfo.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblInfo.Font = new System.Drawing.Font("Arial", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInfo.Location = new System.Drawing.Point(3, 0);
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Size = new System.Drawing.Size(861, 57);
            this.lblInfo.TabIndex = 0;
            this.lblInfo.Text = "Confirm Result";
            this.lblInfo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Details_Data
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1252, 655);
            this.Controls.Add(this.tblMain);
            this.Name = "Details_Data";
            this.Text = "OK2SHIP-Details_Data";
            this.Load += new System.EventHandler(this.Details_Data_Load);
            this.tblMain.ResumeLayout(false);
            this.tbldetail.ResumeLayout(false);
            this.tblRight.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Results)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tblMain;
        private System.Windows.Forms.Label lblDetail;
        private System.Windows.Forms.TableLayoutPanel tbldetail;
        private System.Windows.Forms.TreeView tvItem_Details;
        private System.Windows.Forms.TableLayoutPanel tblRight;
        private System.Windows.Forms.DataGridView DGV_Results;
        private System.Windows.Forms.Button btnConfirm;
        private System.Windows.Forms.Label lblInfo;
    }
}