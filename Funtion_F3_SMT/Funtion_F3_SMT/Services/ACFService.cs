using OfficeOpenXml;
using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.Services
{
    public class ACFService
    {
        DBContext _dbContext = new DBContext();
        public string loadItemNamebyItemCode(string itemCode)
        {


            itemCode = itemCode.Trim();
            DataTable dataTable = _dbContext.LoadDataTable("TABLE_OF_CONTENT_SETTING", new[] { "ItemCode" }, new[] { itemCode }, new[] { "ItemName" });
            if (dataTable.Rows.Count <= 0)
            {
                throw new Exception("Item code not found in the database.");
            }
            DataRow row = dataTable.Rows[0];
            return row["ItemName"].ToString().Trim();
        }


        public static DataTable ReadWettingContactAngle(DataTable dataTable, string data)
        {
            if (dataTable == null || dataTable.Columns.Count < 0)
            {
                dataTable = new DataTable();
                dataTable.Columns.Add("Machine");
                dataTable.Columns.Add("WettingAngle");
                dataTable.Columns.Add("Side");
                dataTable.Columns.Add("Spec");
                dataTable.Columns.Add("Judge");
                dataTable.Columns.Add("Data");
                for (int i = 1; i <= 32; i++)
                {
                    dataTable.Columns.Add($"Sample {i}");
                }

            }
            string[] dataZ = data.Split('\n');
            foreach (string z in dataZ)
            {
                string[] zRay = z.Split('\t');
                if (zRay.Count() >= 32)
                {
                    int i = 0;
                    DataRow row = dataTable.NewRow();
                    foreach (DataColumn col in dataTable.Columns)
                    {
                        if (i >= zRay.Count())
                        {
                            break;
                        }
                        if (col.ColumnName.Contains("Sample"))
                        {

                            row[col.ColumnName] = zRay[i++];
                        }
                    }
                    dataTable.Rows.Add(row);
                }
            }
            return dataTable;
        }
        public string packagingData(Dictionary<string, DataTable> dic, Dictionary<string, List<DateTime>> dic_list)
        {
            var result = new StringBuilder();
            foreach (string machine in dic.Keys)
            {
                string dataTable = ConverterService.DataTableToJson(dic[machine]);
                string[] dateTimes = dic_list[machine].Select(dt => dt.ToString("yyyy-MM-dd HH:mm:ss")).ToArray();
                string dateTimeJoined = string.Join(",", dateTimes);
                string z = $"{machine}|{dateTimeJoined}|{dataTable}";
                result.AppendLine(z);
            }
            return result.ToString();
        }
        public DataTable loadBase(string itemCode, string lotNo, string type)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            type = type.Trim();
            if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo) || string.IsNullOrEmpty(type))
            {
                throw new Exception("Không được để trống itemcode, lotno hoặc type!");
            }
            DataTable dataTable = _dbContext.LoadDataTable("ACF_WCA", new[] { "ItemCode", "LotNo", "Type" }, new[] { itemCode, lotNo, type });
            return dataTable;
        }
        public void Save(string itemCode, string lotno, Dictionary<string, DataTable> dic, Dictionary<string, List<DateTime>> dic_list, string type, bool prime = false)
        {
            DataTable dataTable = new DataTable();
            if (!prime)
            {
                dataTable = loadBase(itemCode, lotno, type);
                if (dataTable.Rows.Count > 0)
                {
                    throw new Exception($"{1231}_Bạn có muốn lưu đè dữ liệu");
                }
            }
            else
            {
                dataTable = _dbContext.GetTableStructure("ACF_WCA");
            }
            DataRow row = dataTable.NewRow();
            string data = packagingData(dic, dic_list);
            row["ItemCode"] = itemCode;
            row["LotNo"] = lotno;
            row["Data"] = data;
            row["Type"] = type;
            row["Operator"] = UserSession.Instance.User_ID;
            dataTable.Rows.Add(row);
            int i = _dbContext.BuckDataTable(dataTable, "ACF_WCA", new[] { "ItemCode", "LotNo", "Type" }, null, "Id");
        }

        public KeyValuePair<Dictionary<string, DataTable>, Dictionary<string, List<DateTime>>> LoadWCA(string itemCode, string lotNo, string type)
        {
            if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo) || string.IsNullOrEmpty(type))
            {
                throw new Exception("Không được để trống itemcode, lotno hoặc type!");
            }
            DataTable dataTable = loadBase(itemCode, lotNo, type);
            if (dataTable.Rows.Count <= 0)
            {
                throw new Exception("Không có dữ liệu nào trong cơ sở dữ liệu.");
            }
            DataRow row = dataTable.Rows[0];
            string data = row["Data"].ToString();
            string[] dataLines = data.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            Dictionary<string, DataTable> dic = new Dictionary<string, DataTable>();
            Dictionary<string, List<DateTime>> dic_list = new Dictionary<string, List<DateTime>>();
            foreach (string line in dataLines)
            {
                string[] parts = line.Split('|');
                if (parts.Length < 3)
                {
                    continue; // Skip invalid lines
                }
                string machine = parts[0].Trim();
                string[] dateTimeStrings = parts[1].Split(',');
                List<DateTime> dateTimes = new List<DateTime>();
                foreach (string dateTimeString in dateTimeStrings)
                {
                    if (DateTime.TryParse(dateTimeString, out DateTime dateTime))
                    {
                        dateTimes.Add(dateTime);
                    }
                }
                DataTable dt = ConverterService.JsonToDataTable(parts[2]);

                if (!dic.ContainsKey(machine))
                {
                    dic[machine] = dt;
                    dic_list[machine] = dateTimes;
                }
            }
            return new KeyValuePair<Dictionary<string, DataTable>, Dictionary<string, List<DateTime>>>(dic, dic_list);
        }
        public static DataTable getStructorPeel()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("ItemCode", typeof(string));
            dt.Columns.Add("LotNo", typeof(string));
            dt.Columns.Add("Pcs_No", typeof(string));
            dt.Columns.Add("Image_Before", typeof(Image));
            dt.Columns.Add("Image_After", typeof(Image));
            dt.Columns.Add("Graph", typeof(Image));
            dt.Columns.Add("Data", typeof(string));
            dt.Columns.Add("Operator", typeof(string));
            dt.Columns.Add("Time_Update", typeof(string));
            dt.Columns.Add("Remark", typeof(string));

            return dt;
        }
        public void GetProductID(string itemCode, string lotNo, string location, string type, DataTable data_tbl)
        {
            if (string.IsNullOrEmpty(location) || string.IsNullOrEmpty(type) || data_tbl == null || data_tbl.Columns.Count == 0)
            {
                throw new Exception("Không được để trống location, type hoặc data_tbl!");
            }
            ProductIDService productIDService = new ProductIDService(itemCode, lotNo, location, new[] { "OQC", "ACF" }, new[] { "Flatness" });
            if (productIDService._listFile.Count <= 0)
            {
                throw new Exception("PRODUCTID: Không có file nào trong thư mục đã chọn.");
            }
            List<string> list = productIDService.getListProductID(productIDService._listFile["Flatness"]);
            data_tbl.Columns.Add("ProductID");
            int i = 0;
            foreach (DataRow row in data_tbl.Rows)
            {
                if (i < list.Count)
                {
                    row["ProductID"] = list[i++];
                }
            }
        }

        public void ExportWCA(ExcelWorksheet ws, string itemCode, string lotNo)
        {
            KeyValuePair<Dictionary<string, DataTable>, Dictionary<string, List<DateTime>>> diz = LoadWCA(itemCode, lotNo, "NPI");
            if (diz.Key.Keys.Count <= 0)
            {
                throw new Exception("No Data");
            }
            string[] strZ = new[] { "Result data machine", "After Packing 60 days", "Sample 32" };
            #region Insert col
            IDictionary<string, string> dicZ = ExportProcess.FindAddressByText(ws, strZ);
            ExportProcess exportProcess = new ExportProcess();
            string enPoint = ws.Cells[ws.Cells[ExportProcess.AddRow(dicZ["After Packing 60 days"], 1)].End.Row, 50].Address;
            string addressRange = $"{dicZ["Result data machine"]}:{enPoint}";
            string addressPointer = ExportProcess.AddRow(dicZ["After Packing 60 days"], 3);
            for (int i = 1; i < diz.Key.Keys.Count; i++)
            {
                exportProcess.CopyAndInsert(ws, addressRange, ref addressPointer, true);
            }
            #endregion
            #region  FillData
            string[] str = new[] { "Result data machine", "Wetting angle" };
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(ws, str);
            string[] strList = dic["Result data machine"].Split('-');
            for (int i = 0; i < strList.Count(); i++)
            {
                string key = diz.Key.Keys.ToArray()[i];
                if (diz.Key.TryGetValue(key, out DataTable dataTable))
                {
                    ws.Cells[strList[i]].Value = $"Result data machine No.{key}";
                    if (dic.TryGetValue("Wetting angle", out string address))
                    {
                        address = ExportProcess.AddRow(address.Split('-')[i], 1);
                        string valuez = ws.Cells[address].Text;
                        int zA = 0;
                        while (!string.IsNullOrEmpty(valuez))
                        {
                            foreach (DataRow row in dataTable.Rows)
                            {
                                if (row["WettingAngle"].ToString().Trim().Contains(valuez))
                                {
                                    #region Spec
                                    string addZ = ExportProcess.AddColumn(address, 2);
                                    if (row["Side"].ToString().ToString().Contains("ACF"))
                                    {
                                        addZ = ExportProcess.AddRow(addZ, 0);
                                        ws.Cells[addZ].Value = $"<{row["Spec"]}°";
                                    }
                                    if (row["Side"].ToString().ToString().Contains("GND"))
                                    {
                                        addZ = ExportProcess.AddRow(addZ, 1);
                                        ws.Cells[addZ].Value = $"<{row["Spec"]}°";
                                    }
                                    #endregion
                                    #region Sample
                                    int z = 1;
                                    addZ = ExportProcess.AddColumn(addZ, 6);
                                    while (true)
                                    {
                                        if (z > 32)
                                        {
                                            break;
                                        }
                                        try
                                        {

                                            ws.Cells[ExportProcess.AddColumn(addZ, z)].Value = double.Parse(row[$"Sample {z}"].ToString());
                                        }
                                        catch
                                        {

                                        }

                                        z++;
                                    }

                                    addZ = ExportProcess.AddColumn(addZ, -5);
                                    ws.Cells[addZ].FormulaR1C1 = $"=MIN(RC[6]:RC[37])";
                                    addZ = ExportProcess.AddColumn(addZ, 1);
                                    ws.Cells[addZ].FormulaR1C1 = $"=MAX(rc[5]:rc[36])";
                                    addZ = ExportProcess.AddColumn(addZ, 1);
                                    ws.Cells[addZ].FormulaR1C1 = $"=AVERAGE(rc[4]:rc[35])";
                                    addZ = ExportProcess.AddColumn(addZ, 1);
                                    ws.Cells[addZ].FormulaR1C1 = $"=STDEV(rc[3]:rc[34])";
                                    addZ = ExportProcess.AddColumn(addZ, 1);
                                    ws.Cells[addZ].FormulaR1C1 = $"=({row["spec"]}-rc[-2])/(3*rc[-1])";
                                    addZ = ExportProcess.AddColumn(addZ, 1);
                                    ws.Cells[addZ].FormulaR1C1 = $"=IF(AND(rc[-4]<={row["spec"]},rc[-1] >1.33),\"OK\",\"NG\")";
                                    #endregion
                                }
                            }
                            if (valuez.Contains("After Packing") && diz.Value.TryGetValue(key, out List<DateTime> list))
                            {
                                string addressRangeX = ExportProcess.getRangeBaseAddressByCellAddress(ws, ExportProcess.AddColumn(address, 8));
                                ws.Cells[addressRangeX].Value = list[zA].ToString("'Plan update' dd/MMM");
                                zA++;

                            }
                            address = ExportProcess.AddRow(address, 2);
                            valuez = ws.Cells[address].Text;
                        }
                    }
                }
            }
            #endregion
            return;
        }
        public DataTable loadBonding(string itemCode, string lotNo, bool nas_status)
        {
            DataTable dataTable = _dbContext.LoadDataTable("Roughness" + (nas_status ? "_NAS" : ""), new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            if (nas_status)
            {
                Debugger.Break();
                string json = dataTable.Rows[0]["Data"].ToString();
                dataTable = ConverterService.JsonToDataTable(json);
            }
            try
            {
                ProductIDService.FillProductID(dataTable, itemCode, lotNo, "Roughness");
            }
            catch { }
            int i = 1;
            foreach (DataRow row in dataTable.Rows)
            {
                row["ID"] = i++;
            }
            return dataTable;
        }
        public DataTable loadPeel(string itemCode, string lotNo, string type, bool prime = true)
        {
            try
            {
                DataTable dataTable = _dbContext.LoadDataTable("ACF_BONDING" + (!prime ? "" : "_NAS"), new[] { "ItemCode", "LotNo", "Remark" }, new[] { itemCode, lotNo , type});
                if (prime)
                {

                    string json = dataTable.Rows[0]["Data"].ToString();
                    string locationImg = dataTable.Rows[0]["LocationImg"].ToString();
                    dataTable = ConverterService.JsonToDataTable(json);
                    NasRepository nas = new NasRepository();
                    nas.MergeDataTable(dataTable, "ACF_BONDING_NAS", itemCode.PadRight(10), lotNo.PadRight(10), locationImg);
                }
                try
                {

                    ProductIDService.FillProductID(dataTable, itemCode, lotNo, "ACF_BONDING");
                }
                catch
                {

                }
                return dataTable;
            }
            catch
            {
                throw new Exception("Không có dữ liệu ACF BONDING trong database");
            }
        }
        public void ExportRoughness(string itemCode, string lotNo, ExcelWorksheet ws, bool nas_mode)
        {
            List<string> list = new[] { "Surface Roughness Measurement", "Max", "Min", "Mean", "Std Dev", "Cpk" }.ToList();
            for (int i = 1; i <= 32; i++)
            {
                list.Add($"Sample {i}");
            }
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(ws, list.ToArray(), true);
            if (dic.TryGetValue("Surface Roughness Measurement", out string address))
            {
                int rowSRM = ws.Cells[address].End.Row;
                List<string> sa1 = new List<string>();
                List<string> sq1 = new List<string>();
                List<string> sdr1 = new List<string>();
                List<string> sa2 = new List<string>();
                List<string> sq2 = new List<string>();
                List<string> sdr2 = new List<string>();
                List<string> sa3 = new List<string>();
                List<string> sq3 = new List<string>();
                List<string> sdr3 = new List<string>();
                DataTable dataTable = loadBonding(itemCode, lotNo, nas_mode);
                if (dataTable.Rows.Count <= 0)
                {
                    throw new Exception("No Data");
                }
                foreach (DataRow row in dataTable.Rows)
                {
                    if (dic.TryGetValue($"Sample {row["ID"]}", out string addressX))
                    {
                        int min = int.MaxValue;

                        foreach (string item in addressX.Split('-'))
                        {
                            int rowX = ws.Cells[item].End.Row;
                            if (rowX - rowSRM > 0 && rowX - rowSRM < min)
                            {
                                min = rowX;
                                address = item;
                            }
                        }
                        if (dataTable.Columns.Contains("ProductID"))
                        {
                            ws.Cells[ExportProcess.AddRow(address, 1)].Value = row["ProductID"];
                        }
                        sa1.Add(ExportProcess.AddRow(address, 2));
                        if (double.TryParse(row["L1_Roughness_Sa"].ToString(), out double o))
                        {
                            ws.Cells[ExportProcess.AddRow(address, 2)].Value = o;
                        }
                        else
                        {
                            ws.Cells[ExportProcess.AddRow(address, 2)].Value = "NA";
                        }
                        sq1.Add(ExportProcess.AddRow(address, 3));
                        if (double.TryParse(row["L1_Roughness_Sq"].ToString(), out o))
                        {
                            ws.Cells[ExportProcess.AddRow(address, 3)].Value = o;
                        }
                        else
                        {
                            ws.Cells[ExportProcess.AddRow(address, 3)].Value = "NA";
                        }

                        sdr1.Add(ExportProcess.AddRow(address, 4));
                        if (double.TryParse(row["L1_Roughness_Sdr"].ToString(), out o))
                        {
                            ws.Cells[ExportProcess.AddRow(address, 4)].Value = o;
                        }
                        else
                        {
                            ws.Cells[ExportProcess.AddRow(address, 4)].Value = "NA";
                        }

                        sa2.Add(ExportProcess.AddRow(address, 5));
                        if (double.TryParse(row["L2_Roughness_Sa"].ToString(), out o))
                        {
                            ws.Cells[ExportProcess.AddRow(address, 5)].Value = o;
                        }
                        else
                        {
                            ws.Cells[ExportProcess.AddRow(address, 5)].Value = "NA";
                        }

                        sq2.Add(ExportProcess.AddRow(address, 6));
                        if (double.TryParse(row["L2_Roughness_Sq"].ToString(), out o))
                        {
                            ws.Cells[ExportProcess.AddRow(address, 6)].Value = o;
                        }
                        else
                        {
                            ws.Cells[ExportProcess.AddRow(address, 6)].Value = "NA";
                        }

                        sdr2.Add(ExportProcess.AddRow(address, 7));
                        if (double.TryParse(row["L2_Roughness_Sdr"].ToString(), out o))
                        {
                            ws.Cells[ExportProcess.AddRow(address, 7)].Value = o;
                        }
                        else
                        {
                            ws.Cells[ExportProcess.AddRow(address, 7)].Value = "NA";
                        }


                        sa3.Add(ExportProcess.AddRow(address, 8));
                        if (double.TryParse(row["L3_Roughness_Sa"].ToString(), out o))
                        {
                            ws.Cells[ExportProcess.AddRow(address, 8)].Value = o;
                        }
                        else
                        {
                            ws.Cells[ExportProcess.AddRow(address, 8)].Value = "NA";
                        }

                        sq3.Add(ExportProcess.AddRow(address, 9));
                        if (double.TryParse(row["L3_Roughness_Sq"].ToString(), out o))
                        {
                            ws.Cells[ExportProcess.AddRow(address, 9)].Value = o;
                        }
                        else
                        {
                            ws.Cells[ExportProcess.AddRow(address, 9)].Value = "NA";
                        }

                        sdr3.Add(ExportProcess.AddRow(address, 10));
                        if (double.TryParse(row["L3_Roughness_Sdr"].ToString(), out o))
                        {
                            ws.Cells[ExportProcess.AddRow(address, 10)].Value = o;
                        }
                        else
                        {
                            ws.Cells[ExportProcess.AddRow(address, 10)].Value = "NA";
                        }
                    }
                }
                foreach (var itemZ in new[] { "Max", "Min", "Mean", "Std Dev", "Cpk" })
                {
                    string fun = "";
                    switch (itemZ)
                    {
                        case "Max":
                        case "Min":
                            fun = itemZ;
                            break;
                        case "Mean":
                            fun = "AVERAGE";
                            break;
                        case "Std Dev":
                            fun = "STDEV";
                            break;
                        case "Cpk":
                            fun = "THUY";
                            break;
                        default:

                            break;
                    }
                    if (dic.TryGetValue(itemZ, out string addressX))
                    {
                        int min = int.MaxValue;
                        foreach (string item in addressX.Split('-'))
                        {
                            int rowX = ws.Cells[item].End.Row;
                            if (rowX - rowSRM > 0 && rowX - rowSRM < min)
                            {
                                min = rowX;
                                address = item;
                            }
                        }
                        if (fun.Equals("THUY"))
                        {
                            for (int i = 0; i < 9; i++)
                            {
                                ws.Cells[ExportProcess.AddRow(address, i + 1)].FormulaR1C1 = $"=(0.5-rc[-2])/(3*rc[-1])";
                            }
                        }
                        else
                        {

                            Dictionary<string, string> ddd = GetMerge(sa1);
                            ws.Cells[ExportProcess.AddRow(address, 1)].Formula = $"={fun}({string.Join(",", ddd.Values.ToArray())})";
                            ddd = GetMerge(sq1);
                            ws.Cells[ExportProcess.AddRow(address, 2)].Formula = $"={fun}({string.Join(",", ddd.Values.ToArray())})";
                            ddd = GetMerge(sdr1);
                            ws.Cells[ExportProcess.AddRow(address, 3)].Formula = $"={fun}({string.Join(",", ddd.Values.ToArray())})";
                            ddd = GetMerge(sa2);
                            ws.Cells[ExportProcess.AddRow(address, 4)].Formula = $"={fun}({string.Join(",", ddd.Values.ToArray())})";
                            ddd = GetMerge(sq2);
                            ws.Cells[ExportProcess.AddRow(address, 5)].Formula = $"={fun}({string.Join(",", ddd.Values.ToArray())})";
                            ddd = GetMerge(sdr2);
                            ws.Cells[ExportProcess.AddRow(address, 6)].Formula = $"={fun}({string.Join(",", ddd.Values.ToArray())})";
                            ddd = GetMerge(sa3);
                            ws.Cells[ExportProcess.AddRow(address, 7)].Formula = $"={fun}({string.Join(",", ddd.Values.ToArray())})";
                            ddd = GetMerge(sq3);
                            ws.Cells[ExportProcess.AddRow(address, 8)].Formula = $"={fun}({string.Join(",", ddd.Values.ToArray())})";
                            ddd = GetMerge(sdr3);
                            ws.Cells[ExportProcess.AddRow(address, 9)].Formula = $"={fun}({string.Join(",", ddd.Values.ToArray())})";

                        }
                        //Debugger.Break();

                    }
                }

            }
        }

        private Dictionary<string, string> GetMerge(List<string> lists)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            List<string> list = new List<string>();
            string point = "";
            lists.Add("D00000");
            foreach (string item in lists)
            {
                if (string.IsNullOrEmpty(point))
                {
                    //get new point
                    point = ConverterService.getDigit(item).ToString();
                }
                if (item.Contains(point))
                {
                    list.Add(item);
                }
                else
                {

                    // merge address
                    list.Sort();

                    dic.Add(point, $"{list[0]}:{list[list.Count() - 1]}");
                    // meke new list
                    list = new List<string>();
                    list.Add(item);
                    //get new point
                    point = "";
                }
            }
            return dic;
        }

        public void ExportBoding(ExcelWorksheet ws, string itemCode, string lotNo, bool nas_mode, string type = "NPI")
        {
            List<string> list = new List<string>();
            for (int i = 1; i <= 32; i++)
            {
                list.Add($"Sample {i}");
            }
            list.Add("MAX");
            list.Add("MIN");
            list.Add("AVERAGE");
            string[] listZ = new[] { "ACF Bonding", "Glass coupon", "Flex side", "Load-travel curve", "Peeling Force", "Fail mode" };
            IDictionary<string, string> dic = DictionaryService.MergeDictionaries(ExportProcess.FindAddressByText(ws, list.ToArray(), true), ExportProcess.FindAddressByText(ws, listZ));
            if (dic.TryGetValue("ACF Bonding", out string address))
            {
                int rowSRM = ws.Cells[address].End.Row;
                DataTable dataTable = loadPeel(itemCode, lotNo, type);
                if (dataTable.Rows.Count <= 0)
                {
                    throw new Exception("No Data");
                }
                List<string> dataList = new List<string>();
                foreach (DataRow row in dataTable.Rows)
                {
                    if (dic.TryGetValue($"Sample {row["Pcs_No"]}", out string addressX))
                    {
                        int min = int.MaxValue;
                        foreach (string item in addressX.Split('-'))
                        {
                            int rowX = ws.Cells[item].End.Row;
                            if (rowX - rowSRM >= 0 && rowX - rowSRM < min)
                            {
                                min = rowX - rowSRM;
                                address = item;
                            }
                        }
                        if (dataTable.Columns.Contains("ProductID"))
                        {
                            ws.Cells[ExportProcess.AddRow(address, -1)].Value = row["ProductID"];
                        }
                        ExportProcess.InsertImageToCell(ws, ws.Cells[ExportProcess.AddRow(address, 1)], ExportProcess.ConvertDataRowToByte(row, "Image_Before"), $"{Guid.NewGuid()}");
                        ExportProcess.InsertImageToCell(ws, ws.Cells[ExportProcess.AddRow(address, 2)], ExportProcess.ConvertDataRowToByte(row, "Image_After"), $"{Guid.NewGuid()}");
                        ExportProcess.InsertImageToCell(ws, ws.Cells[ExportProcess.AddRow(address, 3)], ExportProcess.ConvertDataRowToByte(row, "Graph"), $"{Guid.NewGuid()}");
                        try
                        {
                            ws.Cells[ExportProcess.AddRow(address, 4)].Value = double.Parse(row["Data"].ToString());
                        }
                        catch
                        {
                            ws.Cells[ExportProcess.AddRow(address, 4)].Value = "NA";
                        }
                        dataList.Add(ExportProcess.AddRow(address, 4));
                    }
                }
                Dictionary<string, string> diZc = GetMerge(dataList);
                string addList = string.Join(",", diZc.Values.ToArray());
                if (dic.TryGetValue("MAX", out string addressZ))
                {
                    int min = int.MaxValue;
                    foreach (string item in addressZ.Split('-'))
                    {
                        int rowX = ws.Cells[item].End.Row;
                        if (rowX - rowSRM >= 0 && rowX - rowSRM < min)
                        {
                            min = rowX - rowSRM;
                            address = item;
                        }
                    }
                    ws.Cells[ExportProcess.AddColumn(address, 1)].Formula = $"=MAX({addList})";
                }
                if (dic.TryGetValue("MIN", out addressZ))
                {
                    int min = int.MaxValue;
                    foreach (string item in addressZ.Split('-'))
                    {
                        int rowX = ws.Cells[item].End.Row;
                        if (rowX - rowSRM >= 0 && rowX - rowSRM < min)
                        {
                            min = rowX - rowSRM;
                            address = item;
                        }
                    }
                    ws.Cells[ExportProcess.AddColumn(address, 1)].Formula = $"=MIN({addList})";
                }
                if (dic.TryGetValue("AVERAGE", out addressZ))
                {
                    int min = int.MaxValue;
                    foreach (string item in addressZ.Split('-'))
                    {
                        int rowX = ws.Cells[item].End.Row;
                        if (rowX - rowSRM >= 0 && rowX - rowSRM < min)
                        {
                            min = rowX - rowSRM;
                            address = item;
                        }
                    }
                    ws.Cells[ExportProcess.AddColumn(address, 1)].Formula = $"=AVERAGE({addList})";
                }
            }
        }
    }
}

