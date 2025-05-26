namespace OK2SHIP_Software
{
    partial class FAI_Chart
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
            this.lblFAI = new System.Windows.Forms.Label();
            this.txtSideCheck = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.txtCPKM = new System.Windows.Forms.TextBox();
            this.txtCPK = new System.Windows.Forms.TextBox();
            this.txtCPKU = new System.Windows.Forms.TextBox();
            this.txtCPKL = new System.Windows.Forms.TextBox();
            this.txtCP = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtMin = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtMax = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtMean = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtSTDEV = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtNorminalDim = new System.Windows.Forms.TextBox();
            this.lblLL = new System.Windows.Forms.Label();
            this.lblUL = new System.Windows.Forms.Label();
            this.txtLL = new System.Windows.Forms.TextBox();
            this.txtUL = new System.Windows.Forms.TextBox();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.tabFAI = new System.Windows.Forms.TabPage();
            this.tabCPK_Calcu = new System.Windows.Forms.TabPage();
            this.tabChart = new System.Windows.Forms.TabControl();
            this.zgChartFAI1 = new ZedGraph.ZedGraphControl();
            this.zgHistogramFAI1 = new ZedGraph.ZedGraphControl();
            this.tableLayoutPanel1.SuspendLayout();
            this.tabFAI.SuspendLayout();
            this.tabCPK_Calcu.SuspendLayout();
            this.tabChart.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblFAI
            // 
            this.lblFAI.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.lblFAI.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblFAI.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblFAI.Font = new System.Drawing.Font("Arial", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFAI.Location = new System.Drawing.Point(3, 3);
            this.lblFAI.Name = "lblFAI";
            this.lblFAI.Size = new System.Drawing.Size(786, 48);
            this.lblFAI.TabIndex = 75;
            this.lblFAI.Text = "CPK Result: FAI";
            this.lblFAI.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtSideCheck
            // 
            this.txtSideCheck.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSideCheck.Location = new System.Drawing.Point(199, 63);
            this.txtSideCheck.Name = "txtSideCheck";
            this.txtSideCheck.Size = new System.Drawing.Size(120, 38);
            this.txtSideCheck.TabIndex = 74;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(417, 400);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(78, 31);
            this.label12.TabIndex = 73;
            this.label12.Text = "CPKM";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(435, 328);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(60, 31);
            this.label13.TabIndex = 72;
            this.label13.Text = "CPK";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.Location = new System.Drawing.Point(100, 469);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(76, 31);
            this.label14.TabIndex = 71;
            this.label14.Text = "CPKU";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label15.Location = new System.Drawing.Point(104, 396);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(72, 31);
            this.label15.TabIndex = 70;
            this.label15.Text = "CPKL";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label16.Location = new System.Drawing.Point(131, 324);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(45, 31);
            this.label16.TabIndex = 69;
            this.label16.Text = "CP";
            // 
            // txtCPKM
            // 
            this.txtCPKM.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCPKM.Location = new System.Drawing.Point(521, 396);
            this.txtCPKM.Name = "txtCPKM";
            this.txtCPKM.Size = new System.Drawing.Size(120, 38);
            this.txtCPKM.TabIndex = 68;
            // 
            // txtCPK
            // 
            this.txtCPK.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCPK.Location = new System.Drawing.Point(521, 324);
            this.txtCPK.Name = "txtCPK";
            this.txtCPK.Size = new System.Drawing.Size(120, 38);
            this.txtCPK.TabIndex = 67;
            // 
            // txtCPKU
            // 
            this.txtCPKU.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCPKU.Location = new System.Drawing.Point(199, 469);
            this.txtCPKU.Name = "txtCPKU";
            this.txtCPKU.Size = new System.Drawing.Size(120, 38);
            this.txtCPKU.TabIndex = 66;
            // 
            // txtCPKL
            // 
            this.txtCPKL.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCPKL.Location = new System.Drawing.Point(199, 396);
            this.txtCPKL.Name = "txtCPKL";
            this.txtCPKL.Size = new System.Drawing.Size(120, 38);
            this.txtCPKL.TabIndex = 65;
            // 
            // txtCP
            // 
            this.txtCP.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCP.Location = new System.Drawing.Point(199, 324);
            this.txtCP.Name = "txtCP";
            this.txtCP.Size = new System.Drawing.Size(120, 38);
            this.txtCP.TabIndex = 64;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(441, 240);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(54, 31);
            this.label6.TabIndex = 63;
            this.label6.Text = "MIN";
            // 
            // txtMin
            // 
            this.txtMin.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMin.Location = new System.Drawing.Point(521, 240);
            this.txtMin.Name = "txtMin";
            this.txtMin.Size = new System.Drawing.Size(120, 38);
            this.txtMin.TabIndex = 62;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(433, 181);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(62, 31);
            this.label5.TabIndex = 61;
            this.label5.Text = "MAX";
            // 
            // txtMax
            // 
            this.txtMax.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMax.Location = new System.Drawing.Point(521, 181);
            this.txtMax.Name = "txtMax";
            this.txtMax.Size = new System.Drawing.Size(120, 38);
            this.txtMax.TabIndex = 60;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(427, 122);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(68, 31);
            this.label4.TabIndex = 59;
            this.label4.Text = "Mean";
            // 
            // txtMean
            // 
            this.txtMean.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMean.Location = new System.Drawing.Point(521, 122);
            this.txtMean.Name = "txtMean";
            this.txtMean.Size = new System.Drawing.Size(120, 38);
            this.txtMean.TabIndex = 58;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(406, 63);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(89, 31);
            this.label3.TabIndex = 57;
            this.label3.Text = "STDEV";
            // 
            // txtSTDEV
            // 
            this.txtSTDEV.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSTDEV.Location = new System.Drawing.Point(521, 63);
            this.txtSTDEV.Name = "txtSTDEV";
            this.txtSTDEV.Size = new System.Drawing.Size(120, 38);
            this.txtSTDEV.TabIndex = 56;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(44, 66);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(132, 31);
            this.label2.TabIndex = 55;
            this.label2.Text = "Side_Check";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(24, 125);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(152, 31);
            this.label1.TabIndex = 53;
            this.label1.Text = "Norminal_Dim";
            // 
            // txtNorminalDim
            // 
            this.txtNorminalDim.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtNorminalDim.Location = new System.Drawing.Point(199, 122);
            this.txtNorminalDim.Name = "txtNorminalDim";
            this.txtNorminalDim.Size = new System.Drawing.Size(120, 38);
            this.txtNorminalDim.TabIndex = 52;
            // 
            // lblLL
            // 
            this.lblLL.AutoSize = true;
            this.lblLL.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLL.Location = new System.Drawing.Point(138, 243);
            this.lblLL.Name = "lblLL";
            this.lblLL.Size = new System.Drawing.Size(53, 31);
            this.lblLL.TabIndex = 51;
            this.lblLL.Text = "LSL";
            // 
            // lblUL
            // 
            this.lblUL.AutoSize = true;
            this.lblUL.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblUL.Location = new System.Drawing.Point(134, 184);
            this.lblUL.Name = "lblUL";
            this.lblUL.Size = new System.Drawing.Size(57, 31);
            this.lblUL.TabIndex = 50;
            this.lblUL.Text = "USL";
            // 
            // txtLL
            // 
            this.txtLL.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtLL.Location = new System.Drawing.Point(199, 240);
            this.txtLL.Name = "txtLL";
            this.txtLL.Size = new System.Drawing.Size(120, 38);
            this.txtLL.TabIndex = 49;
            // 
            // txtUL
            // 
            this.txtUL.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtUL.Location = new System.Drawing.Point(199, 181);
            this.txtUL.Name = "txtUL";
            this.txtUL.Size = new System.Drawing.Size(120, 38);
            this.txtUL.TabIndex = 48;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Controls.Add(this.zgHistogramFAI1, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.zgChartFAI1, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(786, 603);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // tabFAI
            // 
            this.tabFAI.Controls.Add(this.tableLayoutPanel1);
            this.tabFAI.Location = new System.Drawing.Point(4, 22);
            this.tabFAI.Name = "tabFAI";
            this.tabFAI.Padding = new System.Windows.Forms.Padding(3);
            this.tabFAI.Size = new System.Drawing.Size(792, 609);
            this.tabFAI.TabIndex = 0;
            this.tabFAI.Text = "FAI_Histogram";
            this.tabFAI.UseVisualStyleBackColor = true;
            // 
            // tabCPK_Calcu
            // 
            this.tabCPK_Calcu.Controls.Add(this.lblFAI);
            this.tabCPK_Calcu.Controls.Add(this.txtSideCheck);
            this.tabCPK_Calcu.Controls.Add(this.label12);
            this.tabCPK_Calcu.Controls.Add(this.label13);
            this.tabCPK_Calcu.Controls.Add(this.label14);
            this.tabCPK_Calcu.Controls.Add(this.label15);
            this.tabCPK_Calcu.Controls.Add(this.label16);
            this.tabCPK_Calcu.Controls.Add(this.txtCPKM);
            this.tabCPK_Calcu.Controls.Add(this.txtCPK);
            this.tabCPK_Calcu.Controls.Add(this.txtCPKU);
            this.tabCPK_Calcu.Controls.Add(this.txtCPKL);
            this.tabCPK_Calcu.Controls.Add(this.txtCP);
            this.tabCPK_Calcu.Controls.Add(this.label6);
            this.tabCPK_Calcu.Controls.Add(this.txtMin);
            this.tabCPK_Calcu.Controls.Add(this.label5);
            this.tabCPK_Calcu.Controls.Add(this.txtMax);
            this.tabCPK_Calcu.Controls.Add(this.label4);
            this.tabCPK_Calcu.Controls.Add(this.txtMean);
            this.tabCPK_Calcu.Controls.Add(this.label3);
            this.tabCPK_Calcu.Controls.Add(this.txtSTDEV);
            this.tabCPK_Calcu.Controls.Add(this.label2);
            this.tabCPK_Calcu.Controls.Add(this.label1);
            this.tabCPK_Calcu.Controls.Add(this.txtNorminalDim);
            this.tabCPK_Calcu.Controls.Add(this.lblLL);
            this.tabCPK_Calcu.Controls.Add(this.lblUL);
            this.tabCPK_Calcu.Controls.Add(this.txtLL);
            this.tabCPK_Calcu.Controls.Add(this.txtUL);
            this.tabCPK_Calcu.Location = new System.Drawing.Point(4, 22);
            this.tabCPK_Calcu.Name = "tabCPK_Calcu";
            this.tabCPK_Calcu.Padding = new System.Windows.Forms.Padding(3);
            this.tabCPK_Calcu.Size = new System.Drawing.Size(792, 609);
            this.tabCPK_Calcu.TabIndex = 4;
            this.tabCPK_Calcu.Text = "Calculate CPK";
            this.tabCPK_Calcu.UseVisualStyleBackColor = true;
            // 
            // tabChart
            // 
            this.tabChart.Controls.Add(this.tabFAI);
            this.tabChart.Controls.Add(this.tabCPK_Calcu);
            this.tabChart.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabChart.Location = new System.Drawing.Point(0, 0);
            this.tabChart.Name = "tabChart";
            this.tabChart.SelectedIndex = 0;
            this.tabChart.Size = new System.Drawing.Size(800, 635);
            this.tabChart.TabIndex = 1;
            // 
            // zgChartFAI1
            // 
            this.zgChartFAI1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.zgChartFAI1.Location = new System.Drawing.Point(3, 3);
            this.zgChartFAI1.Name = "zgChartFAI1";
            this.zgChartFAI1.ScrollGrace = 0D;
            this.zgChartFAI1.ScrollMaxX = 0D;
            this.zgChartFAI1.ScrollMaxY = 0D;
            this.zgChartFAI1.ScrollMaxY2 = 0D;
            this.zgChartFAI1.ScrollMinX = 0D;
            this.zgChartFAI1.ScrollMinY = 0D;
            this.zgChartFAI1.ScrollMinY2 = 0D;
            this.zgChartFAI1.Size = new System.Drawing.Size(780, 295);
            this.zgChartFAI1.TabIndex = 1;
            // 
            // zgHistogramFAI1
            // 
            this.zgHistogramFAI1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.zgHistogramFAI1.Location = new System.Drawing.Point(3, 304);
            this.zgHistogramFAI1.Name = "zgHistogramFAI1";
            this.zgHistogramFAI1.ScrollGrace = 0D;
            this.zgHistogramFAI1.ScrollMaxX = 0D;
            this.zgHistogramFAI1.ScrollMaxY = 0D;
            this.zgHistogramFAI1.ScrollMaxY2 = 0D;
            this.zgHistogramFAI1.ScrollMinX = 0D;
            this.zgHistogramFAI1.ScrollMinY = 0D;
            this.zgHistogramFAI1.ScrollMinY2 = 0D;
            this.zgHistogramFAI1.Size = new System.Drawing.Size(780, 296);
            this.zgHistogramFAI1.TabIndex = 2;
            // 
            // FAI_Chart
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 635);
            this.Controls.Add(this.tabChart);
            this.Name = "FAI_Chart";
            this.Text = "FAI_Chart";
            this.Load += new System.EventHandler(this.FAI_Chart_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tabFAI.ResumeLayout(false);
            this.tabCPK_Calcu.ResumeLayout(false);
            this.tabCPK_Calcu.PerformLayout();
            this.tabChart.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lblFAI;
        private System.Windows.Forms.TextBox txtSideCheck;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.TextBox txtCPKM;
        private System.Windows.Forms.TextBox txtCPK;
        private System.Windows.Forms.TextBox txtCPKU;
        private System.Windows.Forms.TextBox txtCPKL;
        private System.Windows.Forms.TextBox txtCP;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtMin;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtMax;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtMean;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtSTDEV;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtNorminalDim;
        private System.Windows.Forms.Label lblLL;
        private System.Windows.Forms.Label lblUL;
        private System.Windows.Forms.TextBox txtLL;
        private System.Windows.Forms.TextBox txtUL;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TabPage tabFAI;
        private System.Windows.Forms.TabPage tabCPK_Calcu;
        private System.Windows.Forms.TabControl tabChart;
        private ZedGraph.ZedGraphControl zgHistogramFAI1;
        private ZedGraph.ZedGraphControl zgChartFAI1;
    }
}