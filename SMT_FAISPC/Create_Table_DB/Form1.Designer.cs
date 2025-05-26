
namespace Create_Table_DB
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.txtFile = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.btnCreate = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.GB_Table = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.lstTable_Struct = new System.Windows.Forms.ListBox();
            this.btnGetAllTable = new System.Windows.Forms.Button();
            this.btnDeleteAll = new System.Windows.Forms.Button();
            this.lstTable = new System.Windows.Forms.ListBox();
            this.GB_Data = new System.Windows.Forms.GroupBox();
            this.btnDelAll = new System.Windows.Forms.Button();
            this.btnDeleteData = new System.Windows.Forms.Button();
            this.txtLotNo = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtItemCode = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnCreate_Col = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.cbType = new System.Windows.Forms.ComboBox();
            this.btnDelete_Col = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.txtColName = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.GB_Database = new System.Windows.Forms.GroupBox();
            this.rbOK2SHIP_Periods2 = new System.Windows.Forms.RadioButton();
            this.rbOK2SHIP_SMT = new System.Windows.Forms.RadioButton();
            this.rbMaterials = new System.Windows.Forms.RadioButton();
            this.rbRecycled_PGC_Copper = new System.Windows.Forms.RadioButton();
            this.rbDeclaration = new System.Windows.Forms.RadioButton();
            this.rbOK2SHIP_Items = new System.Windows.Forms.RadioButton();
            this.rbIPQC_Data = new System.Windows.Forms.RadioButton();
            this.rbSEI_FAI = new System.Windows.Forms.RadioButton();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.GB_Table.SuspendLayout();
            this.GB_Data.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.GB_Database.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtFile
            // 
            this.txtFile.Location = new System.Drawing.Point(122, 12);
            this.txtFile.Name = "txtFile";
            this.txtFile.Size = new System.Drawing.Size(820, 20);
            this.txtFile.TabIndex = 0;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(12, 8);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(104, 27);
            this.btnBrowse.TabIndex = 1;
            this.btnBrowse.Text = "Browse...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // btnCreate
            // 
            this.btnCreate.Location = new System.Drawing.Point(6, 28);
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.Size = new System.Drawing.Size(104, 37);
            this.btnCreate.TabIndex = 2;
            this.btnCreate.Text = "Create Selected Table";
            this.btnCreate.UseVisualStyleBackColor = true;
            this.btnCreate.Click += new System.EventHandler(this.btnCreate_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(6, 82);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(104, 37);
            this.btnDelete.TabIndex = 3;
            this.btnDelete.Text = "Delete Selected Table";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // GB_Table
            // 
            this.GB_Table.Controls.Add(this.button2);
            this.GB_Table.Controls.Add(this.label5);
            this.GB_Table.Controls.Add(this.lstTable_Struct);
            this.GB_Table.Controls.Add(this.button1);
            this.GB_Table.Controls.Add(this.btnGetAllTable);
            this.GB_Table.Controls.Add(this.btnDeleteAll);
            this.GB_Table.Controls.Add(this.lstTable);
            this.GB_Table.Controls.Add(this.btnDelete);
            this.GB_Table.Controls.Add(this.btnCreate);
            this.GB_Table.Location = new System.Drawing.Point(12, 50);
            this.GB_Table.Name = "GB_Table";
            this.GB_Table.Size = new System.Drawing.Size(545, 326);
            this.GB_Table.TabIndex = 4;
            this.GB_Table.TabStop = false;
            this.GB_Table.Text = "Table Modification";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(404, 28);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(80, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "Table Structure";
            // 
            // lstTable_Struct
            // 
            this.lstTable_Struct.FormattingEnabled = true;
            this.lstTable_Struct.Location = new System.Drawing.Point(404, 54);
            this.lstTable_Struct.Name = "lstTable_Struct";
            this.lstTable_Struct.Size = new System.Drawing.Size(121, 264);
            this.lstTable_Struct.TabIndex = 7;
            this.lstTable_Struct.SelectedIndexChanged += new System.EventHandler(this.lstTable_Struct_SelectedIndexChanged);
            // 
            // btnGetAllTable
            // 
            this.btnGetAllTable.Location = new System.Drawing.Point(6, 227);
            this.btnGetAllTable.Name = "btnGetAllTable";
            this.btnGetAllTable.Size = new System.Drawing.Size(104, 37);
            this.btnGetAllTable.TabIndex = 6;
            this.btnGetAllTable.Text = "Get All Table";
            this.btnGetAllTable.UseVisualStyleBackColor = true;
            this.btnGetAllTable.Click += new System.EventHandler(this.btnGetAllTable_Click);
            // 
            // btnDeleteAll
            // 
            this.btnDeleteAll.Location = new System.Drawing.Point(6, 281);
            this.btnDeleteAll.Name = "btnDeleteAll";
            this.btnDeleteAll.Size = new System.Drawing.Size(104, 37);
            this.btnDeleteAll.TabIndex = 5;
            this.btnDeleteAll.Text = "Delete All Table";
            this.btnDeleteAll.UseVisualStyleBackColor = true;
            this.btnDeleteAll.Click += new System.EventHandler(this.btnDeleteAll_Click);
            // 
            // lstTable
            // 
            this.lstTable.FormattingEnabled = true;
            this.lstTable.Location = new System.Drawing.Point(116, 28);
            this.lstTable.Name = "lstTable";
            this.lstTable.Size = new System.Drawing.Size(282, 290);
            this.lstTable.TabIndex = 4;
            this.lstTable.SelectedIndexChanged += new System.EventHandler(this.lstTable_SelectedIndexChanged);
            // 
            // GB_Data
            // 
            this.GB_Data.Controls.Add(this.btnDelAll);
            this.GB_Data.Controls.Add(this.btnDeleteData);
            this.GB_Data.Controls.Add(this.txtLotNo);
            this.GB_Data.Controls.Add(this.label2);
            this.GB_Data.Controls.Add(this.txtItemCode);
            this.GB_Data.Controls.Add(this.label1);
            this.GB_Data.Location = new System.Drawing.Point(563, 50);
            this.GB_Data.Name = "GB_Data";
            this.GB_Data.Size = new System.Drawing.Size(215, 184);
            this.GB_Data.TabIndex = 5;
            this.GB_Data.TabStop = false;
            this.GB_Data.Text = "Data Delete";
            // 
            // btnDelAll
            // 
            this.btnDelAll.Location = new System.Drawing.Point(124, 128);
            this.btnDelAll.Name = "btnDelAll";
            this.btnDelAll.Size = new System.Drawing.Size(78, 37);
            this.btnDelAll.TabIndex = 7;
            this.btnDelAll.Text = "Delete from all table";
            this.btnDelAll.UseVisualStyleBackColor = true;
            this.btnDelAll.Click += new System.EventHandler(this.btnDelAll_Click);
            // 
            // btnDeleteData
            // 
            this.btnDeleteData.Location = new System.Drawing.Point(18, 128);
            this.btnDeleteData.Name = "btnDeleteData";
            this.btnDeleteData.Size = new System.Drawing.Size(78, 37);
            this.btnDeleteData.TabIndex = 6;
            this.btnDeleteData.Text = "Delete from selected table";
            this.btnDeleteData.UseVisualStyleBackColor = true;
            this.btnDeleteData.Click += new System.EventHandler(this.btnDeleteData_Click);
            // 
            // txtLotNo
            // 
            this.txtLotNo.Location = new System.Drawing.Point(73, 82);
            this.txtLotNo.Name = "txtLotNo";
            this.txtLotNo.Size = new System.Drawing.Size(129, 20);
            this.txtLotNo.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 85);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(36, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "LotNo";
            // 
            // txtItemCode
            // 
            this.txtItemCode.Location = new System.Drawing.Point(73, 45);
            this.txtItemCode.Name = "txtItemCode";
            this.txtItemCode.Size = new System.Drawing.Size(129, 20);
            this.txtItemCode.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 48);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(52, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "ItemCode";
            // 
            // btnCreate_Col
            // 
            this.btnCreate_Col.Location = new System.Drawing.Point(18, 95);
            this.btnCreate_Col.Name = "btnCreate_Col";
            this.btnCreate_Col.Size = new System.Drawing.Size(78, 33);
            this.btnCreate_Col.TabIndex = 7;
            this.btnCreate_Col.Text = "Create";
            this.btnCreate_Col.UseVisualStyleBackColor = true;
            this.btnCreate_Col.Click += new System.EventHandler(this.btnCreate_Col_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.cbType);
            this.groupBox1.Controls.Add(this.btnDelete_Col);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.txtColName);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.btnCreate_Col);
            this.groupBox1.Location = new System.Drawing.Point(563, 240);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(215, 136);
            this.groupBox1.TabIndex = 8;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Table Structure";
            // 
            // cbType
            // 
            this.cbType.FormattingEnabled = true;
            this.cbType.Items.AddRange(new object[] {
            "NVARCHAR(MAX)",
            "VARBINARY(MAX)",
            "INT"});
            this.cbType.Location = new System.Drawing.Point(94, 53);
            this.cbType.Name = "cbType";
            this.cbType.Size = new System.Drawing.Size(111, 21);
            this.cbType.TabIndex = 13;
            // 
            // btnDelete_Col
            // 
            this.btnDelete_Col.Location = new System.Drawing.Point(128, 95);
            this.btnDelete_Col.Name = "btnDelete_Col";
            this.btnDelete_Col.Size = new System.Drawing.Size(78, 33);
            this.btnDelete_Col.TabIndex = 12;
            this.btnDelete_Col.Text = "Delete";
            this.btnDelete_Col.UseVisualStyleBackColor = true;
            this.btnDelete_Col.Click += new System.EventHandler(this.btnDelete_Col_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(15, 55);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(69, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Column Type";
            // 
            // txtColName
            // 
            this.txtColName.Location = new System.Drawing.Point(94, 21);
            this.txtColName.Name = "txtColName";
            this.txtColName.Size = new System.Drawing.Size(112, 20);
            this.txtColName.TabIndex = 9;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 24);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(73, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Column Name";
            // 
            // GB_Database
            // 
            this.GB_Database.Controls.Add(this.rbOK2SHIP_Periods2);
            this.GB_Database.Controls.Add(this.rbOK2SHIP_SMT);
            this.GB_Database.Controls.Add(this.rbMaterials);
            this.GB_Database.Controls.Add(this.rbRecycled_PGC_Copper);
            this.GB_Database.Controls.Add(this.rbDeclaration);
            this.GB_Database.Controls.Add(this.rbOK2SHIP_Items);
            this.GB_Database.Controls.Add(this.rbIPQC_Data);
            this.GB_Database.Controls.Add(this.rbSEI_FAI);
            this.GB_Database.Location = new System.Drawing.Point(792, 50);
            this.GB_Database.Name = "GB_Database";
            this.GB_Database.Size = new System.Drawing.Size(164, 325);
            this.GB_Database.TabIndex = 9;
            this.GB_Database.TabStop = false;
            this.GB_Database.Text = "Database";
            // 
            // rbOK2SHIP_Periods2
            // 
            this.rbOK2SHIP_Periods2.AutoSize = true;
            this.rbOK2SHIP_Periods2.Location = new System.Drawing.Point(24, 287);
            this.rbOK2SHIP_Periods2.Name = "rbOK2SHIP_Periods2";
            this.rbOK2SHIP_Periods2.Size = new System.Drawing.Size(113, 17);
            this.rbOK2SHIP_Periods2.TabIndex = 8;
            this.rbOK2SHIP_Periods2.Text = "OK2SHIP_Period2";
            this.rbOK2SHIP_Periods2.UseVisualStyleBackColor = true;
            // 
            // rbOK2SHIP_SMT
            // 
            this.rbOK2SHIP_SMT.AutoSize = true;
            this.rbOK2SHIP_SMT.Checked = true;
            this.rbOK2SHIP_SMT.Location = new System.Drawing.Point(24, 250);
            this.rbOK2SHIP_SMT.Name = "rbOK2SHIP_SMT";
            this.rbOK2SHIP_SMT.Size = new System.Drawing.Size(100, 17);
            this.rbOK2SHIP_SMT.TabIndex = 6;
            this.rbOK2SHIP_SMT.TabStop = true;
            this.rbOK2SHIP_SMT.Text = "OK2SHIP_SMT";
            this.rbOK2SHIP_SMT.UseVisualStyleBackColor = true;
            this.rbOK2SHIP_SMT.CheckedChanged += new System.EventHandler(this.DB_SelectedChanged);
            // 
            // rbMaterials
            // 
            this.rbMaterials.AutoSize = true;
            this.rbMaterials.Location = new System.Drawing.Point(24, 176);
            this.rbMaterials.Name = "rbMaterials";
            this.rbMaterials.Size = new System.Drawing.Size(67, 17);
            this.rbMaterials.TabIndex = 5;
            this.rbMaterials.Text = "Materials";
            this.rbMaterials.UseVisualStyleBackColor = true;
            this.rbMaterials.CheckedChanged += new System.EventHandler(this.DB_SelectedChanged);
            // 
            // rbRecycled_PGC_Copper
            // 
            this.rbRecycled_PGC_Copper.AutoSize = true;
            this.rbRecycled_PGC_Copper.Location = new System.Drawing.Point(24, 139);
            this.rbRecycled_PGC_Copper.Name = "rbRecycled_PGC_Copper";
            this.rbRecycled_PGC_Copper.Size = new System.Drawing.Size(138, 17);
            this.rbRecycled_PGC_Copper.TabIndex = 4;
            this.rbRecycled_PGC_Copper.Text = "Recycled_PGC_Copper";
            this.rbRecycled_PGC_Copper.UseVisualStyleBackColor = true;
            this.rbRecycled_PGC_Copper.CheckedChanged += new System.EventHandler(this.DB_SelectedChanged);
            // 
            // rbDeclaration
            // 
            this.rbDeclaration.AutoSize = true;
            this.rbDeclaration.Location = new System.Drawing.Point(24, 102);
            this.rbDeclaration.Name = "rbDeclaration";
            this.rbDeclaration.Size = new System.Drawing.Size(79, 17);
            this.rbDeclaration.TabIndex = 3;
            this.rbDeclaration.Text = "Declaration";
            this.rbDeclaration.UseVisualStyleBackColor = true;
            this.rbDeclaration.CheckedChanged += new System.EventHandler(this.DB_SelectedChanged);
            // 
            // rbOK2SHIP_Items
            // 
            this.rbOK2SHIP_Items.AutoSize = true;
            this.rbOK2SHIP_Items.Location = new System.Drawing.Point(24, 213);
            this.rbOK2SHIP_Items.Name = "rbOK2SHIP_Items";
            this.rbOK2SHIP_Items.Size = new System.Drawing.Size(102, 17);
            this.rbOK2SHIP_Items.TabIndex = 2;
            this.rbOK2SHIP_Items.Text = "OK2SHIP_Items";
            this.rbOK2SHIP_Items.UseVisualStyleBackColor = true;
            this.rbOK2SHIP_Items.CheckedChanged += new System.EventHandler(this.DB_SelectedChanged);
            // 
            // rbIPQC_Data
            // 
            this.rbIPQC_Data.AutoSize = true;
            this.rbIPQC_Data.Location = new System.Drawing.Point(24, 65);
            this.rbIPQC_Data.Name = "rbIPQC_Data";
            this.rbIPQC_Data.Size = new System.Drawing.Size(79, 17);
            this.rbIPQC_Data.TabIndex = 1;
            this.rbIPQC_Data.Text = "IPQC_Data";
            this.rbIPQC_Data.UseVisualStyleBackColor = true;
            this.rbIPQC_Data.CheckedChanged += new System.EventHandler(this.DB_SelectedChanged);
            // 
            // rbSEI_FAI
            // 
            this.rbSEI_FAI.AutoSize = true;
            this.rbSEI_FAI.Location = new System.Drawing.Point(24, 28);
            this.rbSEI_FAI.Name = "rbSEI_FAI";
            this.rbSEI_FAI.Size = new System.Drawing.Size(64, 17);
            this.rbSEI_FAI.TabIndex = 0;
            this.rbSEI_FAI.Text = "SEI_FAI";
            this.rbSEI_FAI.UseVisualStyleBackColor = true;
            this.rbSEI_FAI.CheckedChanged += new System.EventHandler(this.DB_SelectedChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(6, 176);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(104, 37);
            this.button1.TabIndex = 6;
            this.button1.Text = "Clone DB";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(6, 128);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(104, 37);
            this.button2.TabIndex = 9;
            this.button2.Text = "Read File";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(968, 382);
            this.Controls.Add(this.GB_Database);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.GB_Data);
            this.Controls.Add(this.GB_Table);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.txtFile);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Database Modification";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.GB_Table.ResumeLayout(false);
            this.GB_Table.PerformLayout();
            this.GB_Data.ResumeLayout(false);
            this.GB_Data.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.GB_Database.ResumeLayout(false);
            this.GB_Database.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtFile;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Button btnCreate;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.GroupBox GB_Table;
        private System.Windows.Forms.Button btnGetAllTable;
        private System.Windows.Forms.Button btnDeleteAll;
        private System.Windows.Forms.ListBox lstTable;
        private System.Windows.Forms.GroupBox GB_Data;
        private System.Windows.Forms.Button btnDeleteData;
        private System.Windows.Forms.TextBox txtLotNo;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtItemCode;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnDelAll;
        private System.Windows.Forms.Button btnCreate_Col;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox cbType;
        private System.Windows.Forms.Button btnDelete_Col;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtColName;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ListBox lstTable_Struct;
        private System.Windows.Forms.GroupBox GB_Database;
        private System.Windows.Forms.RadioButton rbOK2SHIP_SMT;
        private System.Windows.Forms.RadioButton rbMaterials;
        private System.Windows.Forms.RadioButton rbRecycled_PGC_Copper;
        private System.Windows.Forms.RadioButton rbDeclaration;
        private System.Windows.Forms.RadioButton rbOK2SHIP_Items;
        private System.Windows.Forms.RadioButton rbIPQC_Data;
        private System.Windows.Forms.RadioButton rbSEI_FAI;
        private System.Windows.Forms.RadioButton rbOK2SHIP_Periods2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
    }
}

