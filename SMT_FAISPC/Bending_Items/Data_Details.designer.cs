
namespace OK2SHIP
{
    partial class Data_Details
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Data_Details));
            this.DGV_CPK_NET = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.DGV_NG_NET = new System.Windows.Forms.DataGridView();
            this.zedHisto = new ZedGraph.ZedGraphControl();
            this.zedChart = new ZedGraph.ZedGraphControl();
            this.label13 = new System.Windows.Forms.Label();
            this.txtCPK = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtMin = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtMax = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtMean = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtSTDEV = new System.Windows.Forms.TextBox();
            this.lblLL = new System.Windows.Forms.Label();
            this.lblUL = new System.Windows.Forms.Label();
            this.txtLL = new System.Windows.Forms.TextBox();
            this.txtUL = new System.Windows.Forms.TextBox();
            this.splMain = new System.Windows.Forms.SplitContainer();
            this.GBCalculate = new System.Windows.Forms.GroupBox();
            this.tblHisto_Graph = new System.Windows.Forms.TableLayoutPanel();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_CPK_NET)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_NG_NET)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splMain)).BeginInit();
            this.splMain.Panel1.SuspendLayout();
            this.splMain.Panel2.SuspendLayout();
            this.splMain.SuspendLayout();
            this.GBCalculate.SuspendLayout();
            this.tblHisto_Graph.SuspendLayout();
            this.SuspendLayout();
            // 
            // DGV_CPK_NET
            // 
            this.DGV_CPK_NET.AllowUserToAddRows = false;
            this.DGV_CPK_NET.AllowUserToDeleteRows = false;
            this.DGV_CPK_NET.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_CPK_NET.Location = new System.Drawing.Point(15, 35);
            this.DGV_CPK_NET.Name = "DGV_CPK_NET";
            this.DGV_CPK_NET.Size = new System.Drawing.Size(240, 386);
            this.DGV_CPK_NET.TabIndex = 0;
            this.DGV_CPK_NET.RowHeaderMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.DGV_CPK_NET_RowHeaderMouseDoubleClick);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 18F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Blue;
            this.label1.Location = new System.Drawing.Point(10, 4);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(231, 28);
            this.label1.TabIndex = 1;
            this.label1.Text = "Details CPK of NET";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial", 18F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.Red;
            this.label2.Location = new System.Drawing.Point(10, 424);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(188, 28);
            this.label2.TabIndex = 3;
            this.label2.Text = "Details NET NG";
            // 
            // DGV_NG_NET
            // 
            this.DGV_NG_NET.AllowUserToAddRows = false;
            this.DGV_NG_NET.AllowUserToDeleteRows = false;
            this.DGV_NG_NET.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_NG_NET.Location = new System.Drawing.Point(15, 455);
            this.DGV_NG_NET.Name = "DGV_NG_NET";
            this.DGV_NG_NET.Size = new System.Drawing.Size(240, 131);
            this.DGV_NG_NET.TabIndex = 2;
            // 
            // zedHisto
            // 
            this.zedHisto.Dock = System.Windows.Forms.DockStyle.Fill;
            this.zedHisto.Location = new System.Drawing.Point(3, 3);
            this.zedHisto.Name = "zedHisto";
            this.zedHisto.ScrollGrace = 0D;
            this.zedHisto.ScrollMaxX = 0D;
            this.zedHisto.ScrollMaxY = 0D;
            this.zedHisto.ScrollMaxY2 = 0D;
            this.zedHisto.ScrollMinX = 0D;
            this.zedHisto.ScrollMinY = 0D;
            this.zedHisto.ScrollMinY2 = 0D;
            this.zedHisto.Size = new System.Drawing.Size(961, 374);
            this.zedHisto.TabIndex = 17;
            // 
            // zedChart
            // 
            this.zedChart.Dock = System.Windows.Forms.DockStyle.Fill;
            this.zedChart.Location = new System.Drawing.Point(3, 383);
            this.zedChart.Name = "zedChart";
            this.zedChart.ScrollGrace = 0D;
            this.zedChart.ScrollMaxX = 0D;
            this.zedChart.ScrollMaxY = 0D;
            this.zedChart.ScrollMaxY2 = 0D;
            this.zedChart.ScrollMinX = 0D;
            this.zedChart.ScrollMinY = 0D;
            this.zedChart.ScrollMinY2 = 0D;
            this.zedChart.Size = new System.Drawing.Size(961, 375);
            this.zedChart.TabIndex = 18;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Arial Narrow", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(30, 102);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(30, 16);
            this.label13.TabIndex = 82;
            this.label13.Text = "CPK";
            // 
            // txtCPK
            // 
            this.txtCPK.Font = new System.Drawing.Font("Arial Narrow", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCPK.Location = new System.Drawing.Point(66, 99);
            this.txtCPK.Name = "txtCPK";
            this.txtCPK.Size = new System.Drawing.Size(50, 22);
            this.txtCPK.TabIndex = 81;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Arial Narrow", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(145, 67);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(28, 16);
            this.label6.TabIndex = 80;
            this.label6.Text = "MIN";
            // 
            // txtMin
            // 
            this.txtMin.Font = new System.Drawing.Font("Arial Narrow", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMin.Location = new System.Drawing.Point(178, 64);
            this.txtMin.Name = "txtMin";
            this.txtMin.Size = new System.Drawing.Size(50, 22);
            this.txtMin.TabIndex = 79;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Arial Narrow", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(29, 67);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(31, 16);
            this.label5.TabIndex = 78;
            this.label5.Text = "MAX";
            // 
            // txtMax
            // 
            this.txtMax.Font = new System.Drawing.Font("Arial Narrow", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMax.Location = new System.Drawing.Point(66, 64);
            this.txtMax.Name = "txtMax";
            this.txtMax.Size = new System.Drawing.Size(50, 22);
            this.txtMax.TabIndex = 77;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Arial Narrow", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(138, 32);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(35, 16);
            this.label4.TabIndex = 76;
            this.label4.Text = "Mean";
            // 
            // txtMean
            // 
            this.txtMean.Font = new System.Drawing.Font("Arial Narrow", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMean.Location = new System.Drawing.Point(178, 29);
            this.txtMean.Name = "txtMean";
            this.txtMean.Size = new System.Drawing.Size(50, 22);
            this.txtMean.TabIndex = 75;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arial Narrow", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(16, 32);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(44, 16);
            this.label3.TabIndex = 74;
            this.label3.Text = "STDEV";
            // 
            // txtSTDEV
            // 
            this.txtSTDEV.Font = new System.Drawing.Font("Arial Narrow", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSTDEV.Location = new System.Drawing.Point(66, 29);
            this.txtSTDEV.Name = "txtSTDEV";
            this.txtSTDEV.Size = new System.Drawing.Size(50, 22);
            this.txtSTDEV.TabIndex = 73;
            // 
            // lblLL
            // 
            this.lblLL.AutoSize = true;
            this.lblLL.Font = new System.Drawing.Font("Arial Narrow", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLL.Location = new System.Drawing.Point(146, 136);
            this.lblLL.Name = "lblLL";
            this.lblLL.Size = new System.Drawing.Size(27, 16);
            this.lblLL.TabIndex = 86;
            this.lblLL.Text = "LSL";
            // 
            // lblUL
            // 
            this.lblUL.AutoSize = true;
            this.lblUL.Font = new System.Drawing.Font("Arial Narrow", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblUL.Location = new System.Drawing.Point(31, 136);
            this.lblUL.Name = "lblUL";
            this.lblUL.Size = new System.Drawing.Size(29, 16);
            this.lblUL.TabIndex = 85;
            this.lblUL.Text = "USL";
            // 
            // txtLL
            // 
            this.txtLL.Font = new System.Drawing.Font("Arial Narrow", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtLL.Location = new System.Drawing.Point(178, 133);
            this.txtLL.Name = "txtLL";
            this.txtLL.Size = new System.Drawing.Size(50, 22);
            this.txtLL.TabIndex = 84;
            // 
            // txtUL
            // 
            this.txtUL.Font = new System.Drawing.Font("Arial Narrow", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtUL.Location = new System.Drawing.Point(66, 133);
            this.txtUL.Name = "txtUL";
            this.txtUL.Size = new System.Drawing.Size(50, 22);
            this.txtUL.TabIndex = 83;
            // 
            // splMain
            // 
            this.splMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splMain.Location = new System.Drawing.Point(0, 0);
            this.splMain.Name = "splMain";
            // 
            // splMain.Panel1
            // 
            this.splMain.Panel1.Controls.Add(this.GBCalculate);
            this.splMain.Panel1.Controls.Add(this.DGV_CPK_NET);
            this.splMain.Panel1.Controls.Add(this.label1);
            this.splMain.Panel1.Controls.Add(this.label2);
            this.splMain.Panel1.Controls.Add(this.DGV_NG_NET);
            // 
            // splMain.Panel2
            // 
            this.splMain.Panel2.Controls.Add(this.tblHisto_Graph);
            this.splMain.Size = new System.Drawing.Size(1234, 761);
            this.splMain.SplitterDistance = 263;
            this.splMain.TabIndex = 87;
            // 
            // GBCalculate
            // 
            this.GBCalculate.Controls.Add(this.txtSTDEV);
            this.GBCalculate.Controls.Add(this.label3);
            this.GBCalculate.Controls.Add(this.lblLL);
            this.GBCalculate.Controls.Add(this.txtMean);
            this.GBCalculate.Controls.Add(this.lblUL);
            this.GBCalculate.Controls.Add(this.label4);
            this.GBCalculate.Controls.Add(this.txtLL);
            this.GBCalculate.Controls.Add(this.txtMax);
            this.GBCalculate.Controls.Add(this.txtUL);
            this.GBCalculate.Controls.Add(this.label5);
            this.GBCalculate.Controls.Add(this.label13);
            this.GBCalculate.Controls.Add(this.txtMin);
            this.GBCalculate.Controls.Add(this.txtCPK);
            this.GBCalculate.Controls.Add(this.label6);
            this.GBCalculate.Location = new System.Drawing.Point(15, 592);
            this.GBCalculate.Name = "GBCalculate";
            this.GBCalculate.Size = new System.Drawing.Size(240, 166);
            this.GBCalculate.TabIndex = 2;
            this.GBCalculate.TabStop = false;
            this.GBCalculate.Text = "Calculate";
            // 
            // tblHisto_Graph
            // 
            this.tblHisto_Graph.ColumnCount = 1;
            this.tblHisto_Graph.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tblHisto_Graph.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tblHisto_Graph.Controls.Add(this.zedChart, 0, 1);
            this.tblHisto_Graph.Controls.Add(this.zedHisto, 0, 0);
            this.tblHisto_Graph.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tblHisto_Graph.Location = new System.Drawing.Point(0, 0);
            this.tblHisto_Graph.Name = "tblHisto_Graph";
            this.tblHisto_Graph.RowCount = 2;
            this.tblHisto_Graph.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tblHisto_Graph.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tblHisto_Graph.Size = new System.Drawing.Size(967, 761);
            this.tblHisto_Graph.TabIndex = 19;
            // 
            // Data_Details
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1234, 761);
            this.Controls.Add(this.splMain);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Data_Details";
            this.Text = "CPK and Graph Details";
            this.Load += new System.EventHandler(this.Electrical_Details_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DGV_CPK_NET)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_NG_NET)).EndInit();
            this.splMain.Panel1.ResumeLayout(false);
            this.splMain.Panel1.PerformLayout();
            this.splMain.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splMain)).EndInit();
            this.splMain.ResumeLayout(false);
            this.GBCalculate.ResumeLayout(false);
            this.GBCalculate.PerformLayout();
            this.tblHisto_Graph.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView DGV_CPK_NET;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridView DGV_NG_NET;
        private ZedGraph.ZedGraphControl zedHisto;
        private ZedGraph.ZedGraphControl zedChart;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox txtCPK;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtMin;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtMax;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtMean;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtSTDEV;
        private System.Windows.Forms.Label lblLL;
        private System.Windows.Forms.Label lblUL;
        private System.Windows.Forms.TextBox txtLL;
        private System.Windows.Forms.TextBox txtUL;
        private System.Windows.Forms.SplitContainer splMain;
        private System.Windows.Forms.GroupBox GBCalculate;
        private System.Windows.Forms.TableLayoutPanel tblHisto_Graph;
    }
}