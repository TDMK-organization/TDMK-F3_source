using System.ComponentModel;

namespace Export_FPCA_OK2ship_Auto_System.Views
{
    partial class SetupForm
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            AntdUI.StepsItem stepsItem1 = new AntdUI.StepsItem();
            AntdUI.StepsItem stepsItem2 = new AntdUI.StepsItem();
            AntdUI.StepsItem stepsItem3 = new AntdUI.StepsItem();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SetupForm));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.step = new AntdUI.Steps();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.tbl_step1 = new System.Windows.Forms.TableLayoutPanel();
            this.label1 = new AntdUI.Label();
            this.sl_type = new AntdUI.Select();
            this.Category = new AntdUI.Label();
            this.tableLayoutPanel4 = new System.Windows.Forms.TableLayoutPanel();
            this.panel1 = new AntdUI.Panel();
            this.btn_addCateogry = new AntdUI.Button();
            this.ip = new AntdUI.Input();
            this.selectMultiple = new AntdUI.SelectMultiple();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.button1 = new System.Windows.Forms.Button();
            this.btn_back = new System.Windows.Forms.Button();
            this.btn_next = new System.Windows.Forms.Button();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.tbl_step1.SuspendLayout();
            this.tableLayoutPanel4.SuspendLayout();
            this.panel1.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 902F));
            this.tableLayoutPanel1.Controls.Add(this.step, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel3, 0, 2);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 78F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 53F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(902, 635);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // step
            // 
            this.step.Dock = System.Windows.Forms.DockStyle.Fill;
            this.step.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            stepsItem1.Description = "Select Type Setup";
            stepsItem1.Title = "Select";
            stepsItem2.Description = "Setup Report";
            stepsItem2.Title = "Setup";
            stepsItem3.Title = "Run Progess";
            this.step.Items.Add(stepsItem1);
            this.step.Items.Add(stepsItem2);
            this.step.Items.Add(stepsItem3);
            this.step.Location = new System.Drawing.Point(3, 3);
            this.step.Name = "step";
            this.step.Size = new System.Drawing.Size(896, 72);
            this.step.TabIndex = 1;
            this.step.Text = "steps1";
            this.step.TabIndexChanged += new System.EventHandler(this.step_TabIndexChanged);
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 1;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.81505F));
            this.tableLayoutPanel2.Controls.Add(this.tbl_step1, 0, 0);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(3, 81);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(896, 498);
            this.tableLayoutPanel2.TabIndex = 1;
            // 
            // tbl_step1
            // 
            this.tbl_step1.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.tbl_step1.ColumnCount = 2;
            this.tbl_step1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 14.23402F));
            this.tbl_step1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 85.76598F));
            this.tbl_step1.Controls.Add(this.label1, 0, 0);
            this.tbl_step1.Controls.Add(this.sl_type, 1, 0);
            this.tbl_step1.Controls.Add(this.Category, 0, 1);
            this.tbl_step1.Controls.Add(this.tableLayoutPanel4, 1, 1);
            this.tbl_step1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbl_step1.Location = new System.Drawing.Point(3, 3);
            this.tbl_step1.Name = "tbl_step1";
            this.tbl_step1.RowCount = 4;
            this.tbl_step1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 55F));
            this.tbl_step1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 55F));
            this.tbl_step1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 46F));
            this.tbl_step1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 162F));
            this.tbl_step1.Size = new System.Drawing.Size(890, 492);
            this.tbl_step1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(3, 3);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(120, 49);
            this.label1.TabIndex = 0;
            this.label1.Text = "Your Type:\r\n";
            // 
            // sl_type
            // 
            this.sl_type.Dock = System.Windows.Forms.DockStyle.Fill;
            this.sl_type.Location = new System.Drawing.Point(129, 3);
            this.sl_type.Name = "sl_type";
            this.sl_type.PlaceholderText = "Select your type";
            this.sl_type.Size = new System.Drawing.Size(758, 49);
            this.sl_type.TabIndex = 1;
            this.sl_type.SelectedIndexChanged += new AntdUI.IntEventHandler(this.sl_type_SelectedIndexChanged);
            // 
            // Category
            // 
            this.Category.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Category.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Category.Location = new System.Drawing.Point(3, 58);
            this.Category.Name = "Category";
            this.Category.Size = new System.Drawing.Size(120, 49);
            this.Category.TabIndex = 2;
            this.Category.Text = "Category:";
            // 
            // tableLayoutPanel4
            // 
            this.tableLayoutPanel4.ColumnCount = 2;
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 70.78014F));
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 29.21986F));
            this.tableLayoutPanel4.Controls.Add(this.panel1, 1, 0);
            this.tableLayoutPanel4.Controls.Add(this.selectMultiple, 0, 0);
            this.tableLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel4.Location = new System.Drawing.Point(129, 58);
            this.tableLayoutPanel4.Name = "tableLayoutPanel4";
            this.tableLayoutPanel4.RowCount = 1;
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel4.Size = new System.Drawing.Size(758, 49);
            this.tableLayoutPanel4.TabIndex = 4;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btn_addCateogry);
            this.panel1.Controls.Add(this.ip);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(538, 2);
            this.panel1.Margin = new System.Windows.Forms.Padding(2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(218, 45);
            this.panel1.TabIndex = 6;
            this.panel1.Text = "panel1";
            // 
            // btn_addCateogry
            // 
            this.btn_addCateogry.BorderWidth = 1F;
            this.btn_addCateogry.Dock = System.Windows.Forms.DockStyle.Right;
            this.btn_addCateogry.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_addCateogry.IconSvg = resources.GetString("btn_addCateogry.IconSvg");
            this.btn_addCateogry.JoinLeft = true;
            this.btn_addCateogry.Location = new System.Drawing.Point(169, 0);
            this.btn_addCateogry.Margin = new System.Windows.Forms.Padding(2);
            this.btn_addCateogry.Name = "btn_addCateogry";
            this.btn_addCateogry.Size = new System.Drawing.Size(49, 45);
            this.btn_addCateogry.TabIndex = 4;
            this.btn_addCateogry.Type = AntdUI.TTypeMini.Primary;
            this.btn_addCateogry.WaveSize = 0;
            this.btn_addCateogry.Click += new System.EventHandler(this.btn_addCateogry_Click);
            // 
            // ip
            // 
            this.ip.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ip.JoinRight = true;
            this.ip.Location = new System.Drawing.Point(0, 0);
            this.ip.Margin = new System.Windows.Forms.Padding(2);
            this.ip.Name = "ip";
            this.ip.PlaceholderText = "Add new Category";
            this.ip.Size = new System.Drawing.Size(218, 45);
            this.ip.TabIndex = 3;
            this.ip.WaveSize = 0;
            // 
            // selectMultiple
            // 
            this.selectMultiple.Location = new System.Drawing.Point(3, 3);
            this.selectMultiple.Name = "selectMultiple";
            this.selectMultiple.PlaceholderText = "Select Category";
            this.selectMultiple.Size = new System.Drawing.Size(493, 43);
            this.selectMultiple.TabIndex = 3;
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.ColumnCount = 3;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel3.Controls.Add(this.button1, 0, 0);
            this.tableLayoutPanel3.Controls.Add(this.btn_back, 0, 0);
            this.tableLayoutPanel3.Controls.Add(this.btn_next, 2, 0);
            this.tableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel3.Location = new System.Drawing.Point(3, 585);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 1;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel3.Size = new System.Drawing.Size(896, 47);
            this.tableLayoutPanel3.TabIndex = 2;
            // 
            // button1
            // 
            this.button1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button1.Location = new System.Drawing.Point(301, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(292, 41);
            this.button1.TabIndex = 2;
            this.button1.Text = "Close";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btn_back
            // 
            this.btn_back.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_back.Location = new System.Drawing.Point(3, 3);
            this.btn_back.Name = "btn_back";
            this.btn_back.Size = new System.Drawing.Size(292, 41);
            this.btn_back.TabIndex = 1;
            this.btn_back.Text = "<< Previous";
            this.btn_back.UseVisualStyleBackColor = true;
            this.btn_back.Click += new System.EventHandler(this.btn_back_Click);
            // 
            // btn_next
            // 
            this.btn_next.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_next.Location = new System.Drawing.Point(599, 3);
            this.btn_next.Name = "btn_next";
            this.btn_next.Size = new System.Drawing.Size(294, 41);
            this.btn_next.TabIndex = 0;
            this.btn_next.Text = "Next >>";
            this.btn_next.UseVisualStyleBackColor = true;
            this.btn_next.Click += new System.EventHandler(this.btn_next_Click);
            // 
            // SetupForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(902, 635);
            this.ControlBox = false;
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "SetupForm";
            this.Text = "Setup Form";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tbl_step1.ResumeLayout(false);
            this.tableLayoutPanel4.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.tableLayoutPanel3.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        private System.Windows.Forms.Button button1;

        private AntdUI.Panel panel1;
        private AntdUI.Button btn_addCateogry;
        private AntdUI.Input ip;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel4;

        private AntdUI.SelectMultiple selectMultiple;

        private AntdUI.Label Category;

        private AntdUI.Select sl_type;

        private AntdUI.Label label1;

        private System.Windows.Forms.TableLayoutPanel tbl_step1;

        private System.Windows.Forms.Button btn_next;
        private System.Windows.Forms.Button btn_back;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;

        private AntdUI.Steps step;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;

        #endregion
    }
}