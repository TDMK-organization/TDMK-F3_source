using Export_FPCA_OK2ship_Auto_System.Repositories;
using Export_FPCA_OK2ship_Auto_System.Services.TDMK_services;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Diagnostics.PerformanceData;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

namespace Export_FPCA_OK2ship_Auto_System.Services
{
    class SEMServices
    {

        private DBContext _context = new DBContext();
        private List<KeyValuePair<int, Image>> getDataFolder(string dir, List<bool> ConverterString = null)
        {
            List<KeyValuePair<int, Image>> imageList = new List<KeyValuePair<int, Image>>();
            string[] imageExtensions = { ".jpg", ".jpeg", ".png", ".gif", ".bmp" }; // Các phần mở rộng ảnh phổ biến
            bool prime = ConverterString != null;
            try
            {
                foreach (string extension in imageExtensions)
                {
                    string[] files = Directory.GetFiles(dir, $"*{extension}");
                    foreach (string file in files)
                    {
                        try
                        {
                            string filename = Path.GetFileNameWithoutExtension(file).Split('-')[0].ToUpper();
                            if (filename.Contains("PT"))
                            {
                                int id;
                                if (int.TryParse(filename.Substring(2), out id))
                                {
                                    Image image = Image.FromFile(file);
                                    imageList.Add(new KeyValuePair<int, Image>(id, image));
                                    if (prime)
                                    {
                                        ConverterString.Add(Path.GetFileNameWithoutExtension(file).Contains("CONVERTED"));
                                    }
                                }
                            }

                        }
                        catch (OutOfMemoryException)
                        {
                            Console.WriteLine($"Không thể tải ảnh '{file}' vì không đủ bộ nhớ.");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Lỗi khi tải ảnh '{file}': {ex.Message}");
                        }
                    }
                }
            }
            catch (DirectoryNotFoundException)
            {
                throw; // Chuyển tiếp ngoại lệ để xử lý ở lớp gọi
            }
            catch (Exception)
            {
                throw; // Chuyển tiếp ngoại lệ để xử lý ở lớp gọi
            }

            return imageList;
        }
        public DataTable LoadDataProcess(string itemCode, string lotNo)
        {
            try
            {
                DataTable table = new DataTable();
                //1. Kiểm tra ItemCode Lotno
                if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
                {
                    throw new Exception("ItemCode hoặc LotNo không được để trống");
                }

                table = _context.LoadDataTable("SEM_BSE_Binarization_Logfile", new string[] { "ItemCode", "LotNo" }, new string[] { itemCode, lotNo });
                int i = 0;
                foreach (DataRow item in table.Rows)
                {
                    i++;
                    item["Id"] = i;
                }
                return table;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return null;
        }
        public static byte[] ImageToByteArrayFunc(Image imageIn)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                imageIn.Save(ms, ImageFormat.Png); // Or another format like Jpeg
                return ms.ToArray();
            }
        }
        public DataTable JudgementCheck(DataTable dataTable)
        {
            if (!dataTable.Columns.Contains("CheckResults"))
            {
                dataTable.Columns.Add("CheckResults");
            }
            foreach (DataRow row in dataTable.Rows)
            {
                string s1 = row["Black200250"].ToString();
                string s2 = row["Black500700"].ToString();
                string jud = row["Judgement"].ToString();
                if (jud.Equals("Level 3") && double.TryParse(s1, out double d1) && double.TryParse(s2, out double d2))
                {
                    row["CheckResults"] = "OK";
                }
                else
                {

                    row["CheckResults"] = "NG";
                }
            }
            return dataTable;
        }
        public DataTable SEMProcessRead(string location, string itemcode, string lotno, bool prime = false)
        {
            itemcode = itemcode.Trim();
            lotno = lotno.Trim();
            if (!prime)
            {
                DataTable table = _context.LoadDataTable("SEM_BSE_Binarization_Logfile", new string[] { "ItemCode", "LotNo" }, new string[] { itemcode, lotno }, new[] { "ItemCode" });

                if (table.Rows.Count > 0)
                {
                    throw new Exception("12344-Cảnh báo TDMK");
                }
            }
            DataTable dt = new DataTable();
            string[] subdirectories = Directory.GetDirectories(location);
            Dictionary<string, IList<KeyValuePair<int, Image>>> dic = new Dictionary<string, IList<KeyValuePair<int, Image>>>();

            Dictionary<string, DataTable> dataBlack = new Dictionary<string, DataTable>();

            foreach (string dir in subdirectories)
            {
                string s = "";

                IList<Image> list = new List<Image>();
                string[] sub = Directory.GetDirectories(dir);
                if (sub.Length > 0)
                {
                    foreach (string subz in sub)
                    {
                        //get Image
                        s = Path.GetFileName(dir) + "-" + Path.GetFileName(subz);
                        List<bool> converList = new List<bool>();
                        List<KeyValuePair<int, Image>> convertedImage = new List<KeyValuePair<int, Image>>();
                        List<KeyValuePair<int, Image>> normalImage = new List<KeyValuePair<int, Image>>();
                        List<KeyValuePair<int, Image>> res = getDataFolder(subz + "\\READ_PASS", converList);
                        foreach (var item in res)
                        {
                            if (converList[res.IndexOf(item)])
                            {
                                convertedImage.Add(item);
                            }
                            else
                            {
                                normalImage.Add(item);
                            }
                        }
                        dic.Add(s, normalImage);
                        dic.Add(s + "- CONVERTED", convertedImage);

                        // get % black
                        IList<string[]> listBlack = new List<string[]>();
                        string[] files = Directory.GetFiles(subz, "*.csv");
                        DataTable dtBlack = FileFolderRepository.ConvertCsvToDataTable(files[0]);
                        dataBlack.Add(s + "-CONVERTED", dtBlack);
                    }
                }
                else
                {
                    s = Path.GetFileName(dir);
                    dic.Add(s, getDataFolder(dir, new List<bool> { false }));
                }
            }
            dt = _context.GetTableStructure("SEM_BSE_Binarization_Logfile");
            //int id = TDMK_SQL.SQL_MAX("SEM_BSE_Binarization_Logfile", "Id", sqlConneciton);
            // row trên datatable
            Dictionary<int, int> coupleOf = new Dictionary<int, int>();
            foreach (var jtem in dic.Keys)
            {
                IList<KeyValuePair<int, Image>> list = new List<KeyValuePair<int, Image>>();
                if (dic.TryGetValue(jtem, out list))
                {
                    int itemz = getNumber(jtem);
                    string name = "";

                    if (itemz <= 25000 && itemz >= 20000)
                    {
                        name = "Binarization200250";
                    }
                    if (itemz <= 70000 && itemz >= 50000)
                    {
                        name = "Binarization500700";
                    }

                    if (itemz <= 250 && itemz >= 200)
                    {
                        name = "SEM200250";
                    }
                    if (itemz <= 700 && itemz >= 500)
                    {
                        name = "SEM500700";
                    }
                    if (itemz == 5000)
                    {
                        name = "SEM5K";
                    }
                    if (!string.IsNullOrEmpty(name))
                    {
                        addDT(dt, list, coupleOf, name);
                    }
                }
            }
            //Insert Black %
            foreach (var item in dataBlack.Keys)
            {
                string name = "";
                int itemz = getNumber(item);
                if (itemz <= 25000 && itemz >= 20000)
                {
                    name = "Black200250";
                }
                if (itemz <= 70000 && itemz >= 50000)
                {
                    name = "Black500700";
                }
                DataTable dtz = new DataTable();
                if (!string.IsNullOrEmpty(name) && dataBlack.TryGetValue(item, out dtz))
                {
                    addDT(dt, dtz, coupleOf, name);
                }
            }

            // Insert itemcode lotno in datatable
            foreach (DataRow item in dt.Rows)
            {
                item["itemcode"] = itemcode;
                item["LotNo"] = lotno;
            }
            return dt.AsEnumerable().OrderBy(row => row.Field<int>("id")).CopyToDataTable(); ;
        }
        private void addDT(DataTable datatable, DataTable blackValue, Dictionary<int, int> couple, string colName)
        {
            int i = 1;
            foreach (DataRow item in blackValue.Rows)
            {
                int id = i, row;
                if (couple.TryGetValue(i, out row))
                {
                    datatable.Rows[row][colName] = double.Parse(item[4].ToString().Replace("%", ""));
                }
                else
                {
                    couple.Add(id, datatable.Rows.Count);
                    var newRow = datatable.NewRow();
                    newRow["Id"] = id;
                    newRow[colName] = double.Parse(item[4].ToString().Replace("%", ""));
                    datatable.Rows.Add(newRow);
                }
                i++;
            }
        }
        private void addDT(DataTable datatable, IList<KeyValuePair<int, Image>> list, Dictionary<int, int> couple, string colName)
        {
            foreach (var item in list)
            {
                int id = item.Key, row;
                if (couple.TryGetValue(id, out row))
                {
                    datatable.Rows[row][colName] = TDMK_ImageConverter.ImageToByteArray(item.Value, ImageFormat.Jpeg);
                }
                else
                {
                    couple.Add(id, datatable.Rows.Count);
                    var newRow = datatable.NewRow();
                    newRow["Id"] = id;
                    newRow[colName] = TDMK_ImageConverter.ImageToByteArray(item.Value, ImageFormat.Jpeg);
                    datatable.Rows.Add(newRow);
                }
            }
        }
        private int getNumber(string key)
        {
            key = key.ToUpper();
            if (key.Contains("DK"))
            {
                string[] str = key.Split('-');
                if (str.Length == 3)
                {
                    if (str[1].Contains("K"))
                    {
                        return int.Parse(str[1].Replace("K", "000")) * 100;
                    }
                    return int.Parse(str[1].Trim()) * 100;
                }
                return 0;
            }
            if (key.Contains("K"))
            {
                return int.Parse(key.Trim().Replace("K", "000"));
            }
            return int.Parse(key);
        }
        /// <summary>
        /// Save process to database
        /// </summary>
        /// <param name="dataTable">dữ liệu</param>
        /// <param name="prime">Xác nhận ghi đè dữ liệu</param>
        /// <exception cref="NotImplementedException"></exception>
        public void SaveProcess(DataTable dataTable, bool prime)
        {
            if (dataTable.Rows.Count == 0)
            {
                throw new Exception("Không có dữ liệu để lưu");
            }
            //1. Check data had in database (itemcode, lotno)
            // 1.1 Get ItemCode lotno in datatable
            string itemCode = dataTable.Rows[0]["ItemCode"].ToString();
            string lotno = dataTable.Rows[0]["LotNo"].ToString();

            // 1.2 Load data from database
            DataTable dt = _context.LoadDataTable("SEM_BSE_Binarization_Logfile", new string[] { "ItemCode", "LotNo" }, new string[] { itemCode, lotno }, new string[] { "ItemCode", "LotNo" });

            // 1.3 Nếu dữ liệu đã tồn tại và không được phép lưu thì trả về thông báo để hỏi người dùng


            if (!prime && dt.Rows.Count > 0)
            {
                throw new Exception("2267 - Dữ liệu đã tồn tại");
            }

            //2. Save data to database if prime = true
            //2.1 Lưu và thay thế
            int count = _context.BuckDataTable(dataTable, "SEM_BSE_Binarization_Logfile", new string[] { "ItemCode", "LotNo" }, null, "Id");

        }
        public static void Export(SqlConnection sql, ExcelWorksheet worksheet, string itemCode, string lotNo)
        {
            DBContext _context = new DBContext(sql);
            DataTable dt = _context.LoadDataTable("SEM_BSE_Binarization_Logfile", new string[] { "ItemCode", "LotNo" }, new string[] { itemCode, lotNo });
            int countData = dt.Rows.Count;




            string[] colName = new string[] { "SEM BSE 500-700", "SEM 500-700", "Sample", "Binarization", "% Black", "Judgement Level", "SEM 200-250", "SEM 200-300", "SEM 5000", "Binarization Check  Results", "Min", "Max", "Aver" };
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(worksheet, colName, true);

            //pre_process Binarization
            if (dic.TryGetValue("Binarization", out string valueBin))
            {
                string[] valueList = valueBin.Split('-');
                dic.Remove("Binarization");
                for (int i = 1; i < valueList.Length; i++)
                {
                    dic.Add($"Binarization {i + 1}", valueList[i].Trim());
                }
            }
            string addressRow = "", addressCol;
            string addressNum = "";
            if (dic.TryGetValue("Sample", out addressRow))
            {
                addressRow = ExportProcess.AddRow(addressRow, 2);
                int num = -1;
                int iz = 0;
                while (int.TryParse(worksheet.Cells[addressRow].Text, out num) && countData > iz)
                {
                    // Insert Judgement
                    string Jd = (string)dt.Rows[iz]["Judgement"].ToString();
                    if (dic.TryGetValue("Judgement Level", out addressCol))
                    {
                        addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[addressCol].Start.Column].Address;
                        worksheet.Cells[addressCol].Value = Jd;


                    }
                    // Insert SEM5000
                    if (dic.TryGetValue("SEM 5000", out addressCol))
                    {
                        addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[addressCol].Start.Column].Address;
                        byte[] imgData = (byte[])dt.Rows[iz]["SEM5K"];
                        ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[addressCol], imgData, $"SEM500{iz}");
                    }


                    /// Insert SEM200-250
                    byte[] imgData250 = (byte[])dt.Rows[iz]["SEM200250"];


                    if (dic.TryGetValue("SEM 200-250", out addressCol))
                    {
                        string[] add = addressCol.Split('-');
                        for (int i = 0; i < add.Count(); i++)
                        {
                            addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[add[i]].Start.Column].Address;
                            ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[addressCol], imgData250, $"SEM20025020{i}{iz}");
                        }
                    }
                    if (dic.TryGetValue("SEM 200-300", out addressCol))
                    {
                        string[] add = addressCol.Split('-');
                        for (int i = 0; i < add.Count(); i++)
                        {
                            addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[add[i]].Start.Column].Address;
                            ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[addressCol], imgData250, $"SEM200300{i}{iz}");
                        }
                    }

                    /// Insert SEM500-700

                    byte[] imgData500 = (byte[])dt.Rows[iz]["SEM500700"];
                    if (dic.TryGetValue("SEM 500-700", out addressCol))
                    {
                        addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[addressCol].Start.Column].Address;
                        ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[addressCol], imgData500, $"SEM500700{iz}");
                    }
                    if (dic.TryGetValue("SEM BSE 500-700", out addressCol))
                    {
                        addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[addressCol].Start.Column].Address;
                        ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[addressCol], imgData500, $"SEM500700BIN{iz}");
                    }

                    //Insert Binarization Image
                    byte[] imgData200bin = (byte[])dt.Rows[iz]["Binarization500700"];

                    if (dic.TryGetValue("Binarization 2", out addressCol))
                    {
                        addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[addressCol].Start.Column].Address;
                        ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[addressCol], imgData200bin, $"Bin2001{iz}");
                        ExportProcess.AddBorderToImage(worksheet, $"Bin2001{iz}", Color.Green);
                    }

                    try
                    {

                        byte[] imgData500bin = (byte[])dt.Rows[iz]["Binarization200250"];
                        if (dic.TryGetValue("Binarization 3", out addressCol))
                        {
                            addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[addressCol].Start.Column].Address;
                            ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[addressCol], imgData500bin, $"Bin5001{iz}");
                            ExportProcess.AddBorderToImage(worksheet, $"Bin5001{iz}", Color.Green);
                        }
                    }
                    catch
                    {

                    }

                    //Insert black %
                    double x2 = (double)dt.Rows[iz]["Black500700"] / 100;
                    if (dic.TryGetValue("% Black", out addressCol))
                    {
                        addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[addressCol].Start.Column].Address;
                        worksheet.Cells[addressCol].Value = x2;
                        addressNum += addressCol + ',';

                    }


                    //Insert Binarization Check  Results

                    if (dic.TryGetValue("Binarization Check  Results", out addressCol))
                    {
                        addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[addressCol].Start.Column].Address;
                        worksheet.Cells[addressCol].Value = (Jd.Equals("Level 3") && x2 < 0.43) ? "OK" : "NG";
                    }
                    //add row sample
                    addressRow = ExportProcess.AddRow(addressRow, 1);
                    iz++;
                }
                if (dic.TryGetValue("Max", out string addressRowz))
                {
                    if (dic.TryGetValue("% Black", out string addressColz))
                    {
                        addressCol = worksheet.Cells[worksheet.Cells[addressRowz].Start.Row, worksheet.Cells[addressColz].Start.Column].Address;
                        worksheet.Cells[addressCol].Formula = $"Max({addressNum.Trim().TrimEnd(',')})";
                    }

                }
                if (dic.TryGetValue("Min", out addressRowz))
                {
                    if (dic.TryGetValue("% Black", out string addressColz))
                    {
                        addressCol = worksheet.Cells[worksheet.Cells[addressRowz].Start.Row, worksheet.Cells[addressColz].Start.Column].Address;
                        worksheet.Cells[addressCol].Formula = $"Min({addressNum.Trim().TrimEnd(',')})"; ;
                    }

                }
                if (dic.TryGetValue("Aver", out addressRowz))
                {
                    if (dic.TryGetValue("% Black", out string addressColz))
                    {
                        addressCol = worksheet.Cells[worksheet.Cells[addressRowz].Start.Row, worksheet.Cells[addressColz].Start.Column].Address;
                        worksheet.Cells[addressCol].Formula = $"Average({addressNum.Trim().TrimEnd(',')})"; ;
                    }
                }
            }
        }
        private DataTable getDataTableStructor()
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("ID", typeof(int));
            dataTable.Columns.Add("ItemCode", typeof(string));
            dataTable.Columns.Add("LotNo", typeof(string));
            dataTable.Columns.Add("SEM200250", typeof(byte[]));
            dataTable.Columns.Add("SEM500700", typeof(byte[]));
            dataTable.Columns.Add("SEM5K", typeof(byte[]));
            dataTable.Columns.Add("Judgement", typeof(string));
            dataTable.Columns.Add("Binarization200250", typeof(byte[]));
            dataTable.Columns.Add("Black200250", typeof(double));
            dataTable.Columns.Add("Binarization500700", typeof(byte[]));
            dataTable.Columns.Add("Black500700", typeof(double));
            dataTable.Columns.Add("CheckResults", typeof(string));
            return dataTable;
        }
        public DataTable LoadDataProcess(string itemCode, string lotNo, bool prime = false)
        {
            DataTable table = new DataTable();
            try
            {
                //1. Kiểm tra ItemCode Lotno
                if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
                {
                    throw new Exception("ItemCode hoặc LotNo không được để trống");
                }

                table = _context.LoadDataTable("SEM_BSE_Binarization_Logfile" + (prime ? "" : "_NAS"), new string[] { "ItemCode", "LotNo" }, new string[] { itemCode, lotNo });
                if (table.Rows.Count <= 0)
                {
                    table = getDataTableStructor();
                }
                else
                {
                    if (!prime)
                    {
                        try
                        {
                            string json = table.Rows[0]["Data"].ToString();
                            string address = table.Rows[0]["AddressImg"].ToString();
                            table = TDMK_ConverterService.JsonToDataTable(json);
                            NasRepository nas = new NasRepository();
                            nas.MergeDataTable(table, "SEM_BSE_Binarization_Logfile", itemCode, lotNo, address);
                        }
                        catch (Exception ex)
                        {
                            throw new Exception($"Lỗi khi đọc dữ liệu: {ex.Message}");
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return table;
        }
        public string Export(string itemCode, string lotNo, bool prime = false)
        {
            //1. Kiểm tra ItemCode Lotno
            if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
            {
                throw new Exception("ItemCode hoặc LotNo không được để trống");
            }

            //2. Load data from database
            DataTable dt = LoadDataProcess(itemCode, lotNo, false);
            int countData = dt.Rows.Count;
            if (countData <= 0)
            {
                throw new Exception($"{itemCode}-{lotNo} Không tồn tại dữ liệu");
            }
            if (!prime)
            {
                if (countData < 22)
                {
                    throw new Exception($"ATPX4869-{itemCode}-{lotNo} Dữ liệu không đủ");
                }
            }

            //3. Export data to excel
            ExportProcess exportProcess = new ExportProcess();
            using (ExcelPackage ex = exportProcess.FindFormatProcess("SEM BSE & Binarization", itemCode, lotNo))
            {
                using (ExcelWorksheet worksheet = exportProcess.FindSheet(ex, "SEM BSE & Binarization"))
                {

                    string[] colName = new string[] { "SEM BSE 500-700", "SEM 500-700", "Sample", "Binarization", "% Black", "Judgement Level", "SEM 200-250", "SEM 200-300", "SEM 5000", "Binarization Check  Results", "Min", "Max", "Aver" };
                    IDictionary<string, string> dic = ExportProcess.FindAddressByText(worksheet, colName, true);

                    //pre_process Binarization
                    if (dic.TryGetValue("Binarization", out string valueBin))
                    {
                        string[] valueList = valueBin.Split('-');
                        dic.Remove("Binarization");
                        for (int i = 1; i < valueList.Length; i++)
                        {
                            dic.Add($"Binarization {i + 1}", valueList[i].Trim());
                        }
                    }
                    string addressRow = "", addressCol;
                    string addressNum = "";
                    if (dic.TryGetValue("Sample", out addressRow))
                    {
                        addressRow = ExportProcess.AddRow(addressRow, 2);
                        int num = -1;
                        int iz = 0;
                        while (int.TryParse(worksheet.Cells[addressRow].Text, out num) && countData > iz)
                        {
                            // Insert Judgement
                            string Jd = (string)dt.Rows[iz]["Judgement"].ToString();
                            if (dic.TryGetValue("Judgement Level", out addressCol))
                            {
                                addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[addressCol].Start.Column].Address;
                                worksheet.Cells[addressCol].Value = Jd;


                            }
                            // Insert SEM5000
                            if (dic.TryGetValue("SEM 5000", out addressCol))
                            {
                                addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[addressCol].Start.Column].Address;
                                byte[] imgData = (byte[])dt.Rows[iz]["SEM5K"];
                                ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[addressCol], imgData, $"SEM500{iz}");
                            }


                            /// Insert SEM200-250
                            byte[] imgData250 = (byte[])dt.Rows[iz]["SEM200250"];


                            if (dic.TryGetValue("SEM 200-250", out addressCol))
                            {
                                string[] add = addressCol.Split('-');
                                for (int i = 0; i < add.Count(); i++)
                                {
                                    addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[add[i]].Start.Column].Address;
                                    ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[addressCol], imgData250, $"SEM20025020{i}{iz}");
                                }
                            }
                            if (dic.TryGetValue("SEM 200-300", out addressCol))
                            {
                                string[] add = addressCol.Split('-');
                                for (int i = 0; i < add.Count(); i++)
                                {
                                    addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[add[i]].Start.Column].Address;
                                    ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[addressCol], imgData250, $"SEM200300{i}{iz}");
                                }
                            }

                            /// Insert SEM500-700
                            try
                            {

                                byte[] imgData500 = (byte[])dt.Rows[iz]["SEM500700"];
                                if (dic.TryGetValue("SEM 500-700", out addressCol))
                                {
                                    addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[addressCol].Start.Column].Address;
                                    ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[addressCol], imgData500, $"SEM500700{iz}");
                                }
                                if (dic.TryGetValue("SEM BSE 500-700", out addressCol))
                                {
                                    addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[addressCol].Start.Column].Address;
                                    ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[addressCol], imgData500, $"SEM500700BIN{iz}");
                                }
                            }
                            catch
                            {


                            }
                            try
                            {
                                //Insert Binarization Image
                                byte[] imgData200bin = (byte[])dt.Rows[iz]["Binarization500700"];

                                if (dic.TryGetValue("Binarization 2", out addressCol))
                                {
                                    addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[addressCol].Start.Column].Address;
                                    ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[addressCol], imgData200bin, $"Bin2001{iz}");
                                    ExportProcess.AddBorderToImage(worksheet, $"Bin2001{iz}", Color.Green);
                                }
                            }
                            catch
                            {

                            }
                            try
                            {

                                byte[] imgData500bin = (byte[])dt.Rows[iz]["Binarization200250"];
                                if (dic.TryGetValue("Binarization 3", out addressCol))
                                {
                                    addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[addressCol].Start.Column].Address;
                                    ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[addressCol], imgData500bin, $"Bin5001{iz}");
                                    ExportProcess.AddBorderToImage(worksheet, $"Bin5001{iz}", Color.Green);
                                }
                            }
                            catch
                            {

                            }

                            //Insert black %
                            if (double.TryParse(dt.Rows[iz]["Black500700"].ToString(), out double d))
                            {
                                double x2 = d / 100;
                                if (dic.TryGetValue("% Black", out addressCol))
                                {
                                    addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[addressCol].Start.Column].Address;
                                    worksheet.Cells[addressCol].Value = x2;
                                    addressNum += addressCol + ',';

                                }


                                //Insert Binarization Check  Results

                                if (dic.TryGetValue("Binarization Check  Results", out addressCol))
                                {
                                    addressCol = worksheet.Cells[worksheet.Cells[addressRow].Start.Row, worksheet.Cells[addressCol].Start.Column].Address;
                                    worksheet.Cells[addressCol].Value = (Jd.Equals("Level 3") && x2 < 0.43) ? "OK" : "NG";
                                }
                                //add row sample
                                addressRow = ExportProcess.AddRow(addressRow, 1);
                                iz++;
                            }
                        }
                        if (dic.TryGetValue("Max", out string addressRowz))
                        {
                            if (dic.TryGetValue("% Black", out string addressColz))
                            {
                                addressCol = worksheet.Cells[worksheet.Cells[addressRowz].Start.Row, worksheet.Cells[addressColz].Start.Column].Address;
                                worksheet.Cells[addressCol].Formula = $"Max({addressNum.Trim().TrimEnd(',')})";
                            }

                        }
                        if (dic.TryGetValue("Min", out addressRowz))
                        {
                            if (dic.TryGetValue("% Black", out string addressColz))
                            {
                                addressCol = worksheet.Cells[worksheet.Cells[addressRowz].Start.Row, worksheet.Cells[addressColz].Start.Column].Address;
                                worksheet.Cells[addressCol].Formula = $"Min({addressNum.Trim().TrimEnd(',')})"; ;
                            }

                        }
                        if (dic.TryGetValue("Aver", out addressRowz))
                        {
                            if (dic.TryGetValue("% Black", out string addressColz))
                            {
                                addressCol = worksheet.Cells[worksheet.Cells[addressRowz].Start.Row, worksheet.Cells[addressColz].Start.Column].Address;
                                worksheet.Cells[addressCol].Formula = $"Average({addressNum.Trim().TrimEnd(',')})"; ;
                            }
                        }
                    }
                    try
                    {
                        exportProcess.SaveExcelWorksheet(ex, "SEM BSE & Binarization", $"{itemCode.Trim()}_{lotNo.Trim()}", "NPI", false);
                    }
                    catch (Exception exz)
                    {
                        return $"Lỗi khi mở file export: bạn có thể đang mở file export! {exz}";
                    }
                    return "OK";
                }
            }
            throw new NotImplementedException();
        }

        public static string setup_spec(ExcelWorksheet worksheet)
        {

            return "";
        }
    }
}
