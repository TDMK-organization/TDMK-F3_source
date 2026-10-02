using System.ComponentModel;

namespace Export_FPCA_OK2ship_Auto_System.Views
{
    partial class DBScreen
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.label4 = new AntdUI.Label();
            this.dgv = new AntdUI.Table();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.label3 = new AntdUI.Label();
            this.label2 = new AntdUI.Label();
            this.ip_itemCode = new AntdUI.Input();
            this.ip_lotno = new AntdUI.Input();
            this.ip_maker = new AntdUI.Input();
            this.label1 = new AntdUI.Label();
            this.btn_load = new AntdUI.Button();
            this.label5 = new AntdUI.Label();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 190F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.Controls.Add(this.label4, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.dgv, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.label5, 1, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 448F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(870, 489);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // label4
            // 
            this.label4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(4, 4);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(184, 32);
            this.label4.TabIndex = 6;
            this.label4.Text = "Input";
            // 
            // dgv
            // 
            this.dgv.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv.Gap = 12;
            this.dgv.Location = new System.Drawing.Point(195, 43);
            this.dgv.Name = "dgv";
            this.dgv.Size = new System.Drawing.Size(713, 442);
            this.dgv.TabIndex = 0;
            this.dgv.Text = "table1";
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 1;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Controls.Add(this.tableLayoutPanel3, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.btn_load, 0, 1);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(4, 43);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 2;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(184, 442);
            this.tableLayoutPanel2.TabIndex = 1;
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.ColumnCount = 1;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel3.Controls.Add(this.label3, 0, 4);
            this.tableLayoutPanel3.Controls.Add(this.label2, 0, 2);
            this.tableLayoutPanel3.Controls.Add(this.ip_itemCode, 0, 1);
            this.tableLayoutPanel3.Controls.Add(this.ip_lotno, 0, 3);
            this.tableLayoutPanel3.Controls.Add(this.ip_maker, 0, 5);
            this.tableLayoutPanel3.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel3.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 8;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel3.Size = new System.Drawing.Size(178, 386);
            this.tableLayoutPanel3.TabIndex = 0;
            // 
            // label3
            // 
            this.label3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label3.Location = new System.Drawing.Point(3, 143);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(172, 14);
            this.label3.TabIndex = 5;
            this.label3.Text = "Maker";
            // 
            // label2
            // 
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Location = new System.Drawing.Point(3, 73);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(172, 14);
            this.label2.TabIndex = 4;
            this.label2.Text = "LotNo";
            // 
            // ip_itemCode
            // 
            this.ip_itemCode.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ip_itemCode.Location = new System.Drawing.Point(3, 23);
            this.ip_itemCode.Name = "ip_itemCode";
            this.ip_itemCode.Size = new System.Drawing.Size(172, 44);
            this.ip_itemCode.TabIndex = 0;
            // 
            // ip_lotno
            // 
            this.ip_lotno.Location = new System.Drawing.Point(3, 93);
            this.ip_lotno.Name = "ip_lotno";
            this.ip_lotno.Size = new System.Drawing.Size(172, 44);
            this.ip_lotno.TabIndex = 1;
            // 
            // ip_maker
            // 
            this.ip_maker.Location = new System.Drawing.Point(3, 163);
            this.ip_maker.Name = "ip_maker";
            this.ip_maker.Size = new System.Drawing.Size(172, 44);
            this.ip_maker.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Location = new System.Drawing.Point(3, 3);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(172, 14);
            this.label1.TabIndex = 3;
            this.label1.Text = "ItemCode";
            // 
            // btn_load
            // 
            this.btn_load.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_load.Location = new System.Drawing.Point(3, 395);
            this.btn_load.Name = "btn_load";
            this.btn_load.Size = new System.Drawing.Size(178, 44);
            this.btn_load.TabIndex = 1;
            this.btn_load.Text = "load";
            this.btn_load.Click += new System.EventHandler(this.btn_load_Click);
            // 
            // label5
            // 
            this.label5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(195, 4);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(713, 32);
            this.label5.TabIndex = 6;
            this.label5.Text = "Data";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // DBScreen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.Controls.Add(this.tableLayoutPanel1);
            this.Location = new System.Drawing.Point(15, 15);
            this.Name = "DBScreen";
            this.Size = new System.Drawing.Size(870, 489);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel3.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        private AntdUI.Label label5;

        private AntdUI.Label label4;

        private AntdUI.Input ip_maker;
        private AntdUI.Label label1;
        private AntdUI.Label label2;
        private AntdUI.Label label3;

        private AntdUI.Input ip_lotno;

        private AntdUI.Input ip_itemCode;

        private AntdUI.Button btn_load;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;

        private AntdUI.Table dgv;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;

        #endregion
    }
}