namespace TESTImageFactory
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            textBox1 = new TextBox();
            label1 = new Label();
            button1 = new Button();
            dataGridView1 = new DataGridView();
            button2 = new Button();
            label3 = new Label();
            tableLayoutPanel1 = new TableLayoutPanel();
            tb_statusAPI = new Label();
            label5 = new Label();
            label4 = new Label();
            button3 = new Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            tableLayoutPanel1.SuspendLayout();
            SuspendLayout();
            // 
            // textBox1
            // 
            textBox1.Location = new Point(115, 21);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(675, 23);
            textBox1.TabIndex = 0;
            textBox1.Text = "G:\\GD 2 F3\\DATA 3.1\\DATA 3.1\\10.shear\\5CCEV-720555-00001_78219 DLK F4 32PT";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(12, 36);
            label1.Name = "label1";
            label1.Size = new Size(38, 15);
            label1.TabIndex = 1;
            label1.Text = "label1";
            // 
            // button1
            // 
            button1.Location = new Point(796, 12);
            button1.Name = "button1";
            button1.Size = new Size(83, 39);
            button1.TabIndex = 2;
            button1.Text = "Location";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(12, 66);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.Size = new Size(778, 468);
            dataGridView1.TabIndex = 3;
            // 
            // button2
            // 
            button2.Location = new Point(885, 12);
            button2.Name = "button2";
            button2.Size = new Size(83, 39);
            button2.TabIndex = 4;
            button2.Text = "Save";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Dock = DockStyle.Fill;
            label3.Font = new Font("Segoe UI", 20.25F, FontStyle.Bold);
            label3.ForeColor = Color.Black;
            label3.Location = new Point(3, 0);
            label3.Name = "label3";
            label3.Size = new Size(146, 50);
            label3.TabIndex = 6;
            label3.Text = "API";
            label3.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 2;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34.6153831F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65.38461F));
            tableLayoutPanel1.Controls.Add(tb_statusAPI, 1, 0);
            tableLayoutPanel1.Controls.Add(label5, 1, 1);
            tableLayoutPanel1.Controls.Add(label4, 0, 1);
            tableLayoutPanel1.Controls.Add(label3, 0, 0);
            tableLayoutPanel1.Location = new Point(796, 66);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 2;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel1.Size = new Size(442, 100);
            tableLayoutPanel1.TabIndex = 7;
            // 
            // tb_statusAPI
            // 
            tb_statusAPI.AutoSize = true;
            tb_statusAPI.Dock = DockStyle.Fill;
            tb_statusAPI.Font = new Font("Segoe UI", 20.25F, FontStyle.Bold);
            tb_statusAPI.ForeColor = Color.Silver;
            tb_statusAPI.Location = new Point(155, 0);
            tb_statusAPI.Name = "tb_statusAPI";
            tb_statusAPI.Size = new Size(284, 50);
            tb_statusAPI.TabIndex = 9;
            tb_statusAPI.Text = "None";
            tb_statusAPI.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Dock = DockStyle.Fill;
            label5.Font = new Font("Segoe UI", 20.25F, FontStyle.Bold);
            label5.ForeColor = Color.Silver;
            label5.Location = new Point(155, 50);
            label5.Name = "label5";
            label5.Size = new Size(284, 50);
            label5.TabIndex = 8;
            label5.Text = "None";
            label5.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Dock = DockStyle.Fill;
            label4.Font = new Font("Segoe UI", 20.25F, FontStyle.Bold);
            label4.ForeColor = Color.Black;
            label4.Location = new Point(3, 50);
            label4.Name = "label4";
            label4.Size = new Size(146, 50);
            label4.TabIndex = 7;
            label4.Text = "Database";
            label4.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // button3
            // 
            button3.Location = new Point(796, 172);
            button3.Name = "button3";
            button3.Size = new Size(301, 39);
            button3.TabIndex = 8;
            button3.Text = "Check Status";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1250, 561);
            Controls.Add(button3);
            Controls.Add(tableLayoutPanel1);
            Controls.Add(button2);
            Controls.Add(dataGridView1);
            Controls.Add(button1);
            Controls.Add(label1);
            Controls.Add(textBox1);
            Name = "Form1";
            Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            tableLayoutPanel1.ResumeLayout(false);
            tableLayoutPanel1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox textBox1;
        private Label label1;
        private Button button1;
        private DataGridView dataGridView1;
        private Button button2;
        private Label label3;
        private TableLayoutPanel tableLayoutPanel1;
        private Label tb_statusAPI;
        private Label label5;
        private Label label4;
        private Button button3;
    }
}
