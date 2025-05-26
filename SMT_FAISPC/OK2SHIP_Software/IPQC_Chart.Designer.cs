namespace OK2SHIP_Software
{
    partial class IPQC_Chart
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
            this.zgChartIPQC = new ZedGraph.ZedGraphControl();
            this.SuspendLayout();
            // 
            // zgChartIPQC
            // 
            this.zgChartIPQC.Dock = System.Windows.Forms.DockStyle.Fill;
            this.zgChartIPQC.Location = new System.Drawing.Point(0, 0);
            this.zgChartIPQC.Name = "zgChartIPQC";
            this.zgChartIPQC.ScrollGrace = 0D;
            this.zgChartIPQC.ScrollMaxX = 0D;
            this.zgChartIPQC.ScrollMaxY = 0D;
            this.zgChartIPQC.ScrollMaxY2 = 0D;
            this.zgChartIPQC.ScrollMinX = 0D;
            this.zgChartIPQC.ScrollMinY = 0D;
            this.zgChartIPQC.ScrollMinY2 = 0D;
            this.zgChartIPQC.Size = new System.Drawing.Size(1184, 761);
            this.zgChartIPQC.TabIndex = 2;
            // 
            // IPQC_Chart
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1184, 761);
            this.Controls.Add(this.zgChartIPQC);
            this.Name = "IPQC_Chart";
            this.Text = "IPQC_Chart";
            this.Load += new System.EventHandler(this.IPQC_Chart_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private ZedGraph.ZedGraphControl zgChartIPQC;
    }
}