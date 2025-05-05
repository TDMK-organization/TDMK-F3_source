using OfficeOpenXml;
using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;

namespace OK2SHIP_SMT.UserControls
{
    public partial class UC_NewFeature : UserControl
    {
        public UC_NewFeature()
        {
            InitializeComponent();
        }

        private void tb_locationFolder_DoubleClick(object sender, EventArgs e)
        {
            OpenFolderDialog();
        }
        private void btn_runProcess_Click(object sender, EventArgs e)
        {
            DataTable dataTable = (DataTable)dataGridView.DataSource;
            try
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    if (!row["File Name"].ToString().Contains("-PROCESS"))
                    {
                        row["Status"] = solveFileProcess(row["File Path"].ToString()) ? "Done" : "Fail";
                    }
                }
            }
            catch
            {
                MessageBox.Show("Plse scan data");
            }
        }
        private void btnGetData_Click(object sender, EventArgs e)
        {
            scanFolder(tb_locationFolder.Text.Trim());
        }

        private void scanFolder(string location)
        {
            location = location.Trim();
            if (string.IsNullOrEmpty(location))
            {
                MessageBox.Show("Please select a folder to scan.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (!System.IO.Directory.Exists(location))
            {
                MessageBox.Show("The specified folder does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            List<string> xlsxFiles = GetAllXlsxFiles(location);
            if (xlsxFiles.Count <= 0)
            {
                MessageBox.Show("No .xlsx files found in the selected folder.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;

            }
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("File Name", typeof(string));
            dataTable.Columns.Add("File Path", typeof(string));
            dataTable.Columns.Add("Status", typeof(string));

            foreach (string item in xlsxFiles)
            {
                dataTable.Rows.Add(Path.GetFileName(item), item, "Not Processed");
            }
            dataGridView.DataSource = dataTable;
        }

        private bool solveFileProcess(string locationFile)
        {
            try
            {
                ExportProcess exportProcess = new ExportProcess();
                using (ExcelPackage package = new ExcelPackage(locationFile))
                {
                    using (ExcelWorksheet worksheet = package.Workbook.Worksheets[0])
                    {
                        string[] sheet = new[] { "ĐK thêm dòng", "Vlo", "Tháng làm hố sơ", "Công ty mua", "STT in BKLS", "STT lô làm hs" };
                        IDictionary<string, string> dic = ExportProcess.FindAddressByText(worksheet, sheet);
                        if (dic.TryGetValue("STT lô làm hs", out string addZ))
                        {

                        }
                        else
                        {
                            MessageBox.Show("Không tìm thấy cột STT lô làm hs");
                        }
                        if (dic.TryGetValue("ĐK thêm dòng", out string address))
                        {
                            address = ExportProcess.AddRow(address, 1);
                            while (true)
                            {
                                string value = worksheet.Cells[address].Text;
                                if (string.IsNullOrEmpty(value))
                                {
                                    break;
                                }
                                if (int.TryParse(value, out int colAdd))
                                {
                                    worksheet.InsertRow(worksheet.Cells[address].End.Row + 1, colAdd - 1);
                                    string addRoot = $"{address}:{worksheet.Cells[worksheet.Cells[address].End.Row, worksheet.Cells[addZ].End.Column]}";
                                    address = ExportProcess.AddRow(address, 1);
                                    for (int i = 1; i < colAdd; i++)
                                    {
                                        exportProcess.CopyAndInsert(worksheet, addRoot, ref address);
                                    }
                                }
                            }
                            FileInfo file = new FileInfo($"{Path.GetDirectoryName(locationFile)}\\{Path.GetFileNameWithoutExtension(locationFile)}-PROCESS.xlsx");
                            package.SaveAs(file);
                            return true;
                        }
                        else
                        {
                            MessageBox.Show("Không tìm thấy địa chỉ ĐK thêm dòng", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return false;
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Debugger.Break();
                return false;
            }
        }

        public static List<string> GetAllXlsxFiles(string folderPath)
        {
            if (Directory.Exists(folderPath))
            {
                return Directory.GetFiles(folderPath, "*.xlsx", SearchOption.AllDirectories).ToList();
            }
            else
            {
                return new List<string>(); // Trả về danh sách rỗng nếu thư mục không tồn tại
            }
        }
        private void OpenFolderDialog()
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = "Select a folder";
            folderBrowserDialog.ShowNewFolderButton = true;
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                tb_locationFolder.Text = folderBrowserDialog.SelectedPath;
            }

        }


    }
}
