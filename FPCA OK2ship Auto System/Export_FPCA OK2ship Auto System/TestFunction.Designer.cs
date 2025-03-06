namespace Export_FPCA_OK2ship_Auto_System
{
    partial class TestFunction
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
            this.DGV_Spec = new System.Windows.Forms.DataGridView();
            this.DGV_FAI = new System.Windows.Forms.DataGridView();
            this.DGV_CPK = new System.Windows.Forms.DataGridView();
            this.txtItemCode = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtLotNo = new System.Windows.Forms.TextBox();
            this.btnTest = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Spec)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_FAI)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_CPK)).BeginInit();
            this.SuspendLayout();
            // 
            // DGV_Spec
            // 
            this.DGV_Spec.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Spec.Location = new System.Drawing.Point(11, 13);
            this.DGV_Spec.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.DGV_Spec.Name = "DGV_Spec";
            this.DGV_Spec.RowTemplate.Height = 33;
            this.DGV_Spec.Size = new System.Drawing.Size(418, 268);
            this.DGV_Spec.TabIndex = 0;
            // 
            // DGV_FAI
            // 
            this.DGV_FAI.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_FAI.Location = new System.Drawing.Point(11, 298);
            this.DGV_FAI.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.DGV_FAI.Name = "DGV_FAI";
            this.DGV_FAI.RowTemplate.Height = 33;
            this.DGV_FAI.Size = new System.Drawing.Size(418, 224);
            this.DGV_FAI.TabIndex = 1;
            // 
            // DGV_CPK
            // 
            this.DGV_CPK.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_CPK.Location = new System.Drawing.Point(444, 13);
            this.DGV_CPK.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.DGV_CPK.Name = "DGV_CPK";
            this.DGV_CPK.RowTemplate.Height = 33;
            this.DGV_CPK.Size = new System.Drawing.Size(442, 449);
            this.DGV_CPK.TabIndex = 2;
            // 
            // txtItemCode
            // 
            this.txtItemCode.Location = new System.Drawing.Point(578, 488);
            this.txtItemCode.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtItemCode.Name = "txtItemCode";
            this.txtItemCode.Size = new System.Drawing.Size(52, 20);
            this.txtItemCode.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(515, 491);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(52, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "ItemCode";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(647, 491);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(36, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "LotNo";
            // 
            // txtLotNo
            // 
            this.txtLotNo.Location = new System.Drawing.Point(710, 491);
            this.txtLotNo.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtLotNo.Name = "txtLotNo";
            this.txtLotNo.Size = new System.Drawing.Size(52, 20);
            this.txtLotNo.TabIndex = 5;
            // 
            // btnTest
            // 
            this.btnTest.Location = new System.Drawing.Point(792, 491);
            this.btnTest.Name = "btnTest";
            this.btnTest.Size = new System.Drawing.Size(75, 23);
            this.btnTest.TabIndex = 7;
            this.btnTest.Text = "Test";
            this.btnTest.UseVisualStyleBackColor = true;
            this.btnTest.Click += new System.EventHandler(this.btnTest_Click);
            // 
            // TestFunction
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(908, 551);
            this.Controls.Add(this.btnTest);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtLotNo);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtItemCode);
            this.Controls.Add(this.DGV_CPK);
            this.Controls.Add(this.DGV_FAI);
            this.Controls.Add(this.DGV_Spec);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "TestFunction";
            this.Text = "TestFunction";
            this.Load += new System.EventHandler(this.TestFunction_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Spec)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_FAI)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_CPK)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView DGV_Spec;
        private System.Windows.Forms.DataGridView DGV_FAI;
        private System.Windows.Forms.DataGridView DGV_CPK;
        private System.Windows.Forms.TextBox txtItemCode;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtLotNo;
        private System.Windows.Forms.Button btnTest;
    }
}