using OfficeOpenXml;
using OfficeOpenXml.Style;
using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net.WebSockets;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZedGraph;

namespace OK2SHIP_SMT.Services
{
    public class AirBubbleService
    {
        private DBContext _dBContext = new DBContext();
        public string _itemCode { get; set; }
        public string _lotNo { get; set; }

        public Dictionary<string, DataTable> _beforeImage { get; set; } = new Dictionary<string, DataTable>();
        public Dictionary<string, KeyValuePair<double, byte[]>> _dicTONG = new Dictionary<string, KeyValuePair<double, byte[]>>();
        public Dictionary<string, DataTable> _dic = new Dictionary<string, DataTable>();

        private List<KeyValuePair<string, string>> listRefer = new List<KeyValuePair<string, string>>();
        public AirBubbleService()
        {
        }

        public AirBubbleService(string itemCode, string lotNo)
        {
            _itemCode = itemCode;
            _lotNo = lotNo;
            makeNotification(new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
        }


        private void makeNotification(string[] name, string[] value)
        {
            List<string> err = new List<string>();
            for (int i = 0; i < name.Count(); i++)
            {
                if (string.IsNullOrEmpty(value[i]))
                {
                    err.Add($"[{name[i]}]");
                }
            }
            if (err.Count > 0)
            {
                throw new Exception($"Không được để trống {string.Join(", ", err)}");
            }
            else
            {
                return;
            }
        }
        public void ReadData(string location, string pid)
        {
                location = location.Trim();
            pid = pid.Trim();
                List<string> pidList = new List<string>();
            try
            {

                ProductIDService pidService = new ProductIDService(_itemCode, _lotNo, pid, new[] { "Air Bubble" }, new[] { _itemCode });
                if (pidService._listFile.Count() <= 0)
                {
                    throw new Exception("Không có product ID");
                }
                if (pidService._listFile.TryGetValue(_itemCode, out string Value))
                {
                    pidList = pidService.getListProductID(pidService._listFile[_itemCode]);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"ProductID: {ex}");
            }
            //"5CCEV-720211-00012-CPL5-MIYAGI-JPOND-DO KHI-78888"
            string maker = location.Split('\\')[location.Split('\\').Count() - 1].Split('-')[_lotNo.Split('-').Count() >= 2 ? 4 : 3];

            string[] listTape = FileFolderRepository.GetSubFolders(location);
            foreach (string item in listTape)
            {
                string pai = item.Split('\\')[item.Split('\\').Count() - 1];
                switch (pai)
                {
                    case "BF":
                        _beforeImage.Add($"{_itemCode} - {_lotNo}", getDataInBF(item));
                        if (!_beforeImage[$"{_itemCode} - {_lotNo}"].Columns.Contains("ProductID"))
                        {
                            _beforeImage[$"{_itemCode} - {_lotNo}"].Columns.Add("ProductID");

                        }
                        int i = 0;
                        foreach (DataRow row in _beforeImage[$"{_itemCode} - {_lotNo}"].Rows)
                        {
                            if (i >= pidList.Count)
                            {
                                break;
                            }
                            row["ProductID"] = pidList[i++];

                        }
                        break;
                    case "PSA":
                    case "LINER":
                        SolveFolderPSALINER(item, maker, pai);
                        break;
                }
            }
            foreach (string key in _dic.Keys)
            {
                CalculateArea(_dic[key], _dicTONG[key]);
            }
        }

        private void SolveFolderPSALINER(string location, string maker, string pai)
        {
            string[] subs = FileFolderRepository.GetSubFolders(location);
            foreach (string s in subs)
            {
                int num = ConverterService.getDigit(s.Split('\\')[s.Split('\\').Count() - 1]);
                string name = $"{pai}_{maker} (TAPE {num})";
                DataTable dataTable = SolveTapeFile(s, maker);
                _dic.Add(name, dataTable);
            }
        }
        private DataTable SolveTapeFile(string location, string maker)
        {

            DataTable dataTable = _dBContext.GetTableStructure("AIR_BUBBLE_DATA");
            //dataTable.Columns.Add("ID", typeof(int));

            //dataTable.Columns.Add("TONG", typeof(Image));
            //dataTable.Columns.Add("Measure TONG (mm2)", typeof(double));
            //dataTable.Columns.Add("% Area TONG", typeof(double));

            //dataTable.Columns.Add("TRU", typeof(Image));
            //dataTable.Columns.Add("Measure TRU(mm2)", typeof(double));
            //dataTable.Columns.Add("% Area TRU", typeof(double));
            //dataTable.Columns.Add("Judgement", typeof(string));
            string[] subs = FileFolderRepository.GetSubFolders(location);
            if (subs.Length <= 0)
            {
                solveData(maker, location, dataTable);
            }
            else
            {
                foreach (string item in subs)
                {
                    solveData(maker, item, dataTable);
                }
            }

            return dataTable;
        }
        private void solveData(string maker, string item, DataTable dataTable)
        {
            string pai = item.Split('\\')[item.Split('\\').Count() - 1].Trim();
            IList<KeyValuePair<Image, string>> listImage = FileFolderRepository.ListAllPictureInAFolder(item, "jpg");


            //LIST IMAGE TONG IN TONG FOLDER
            Dictionary<int, KeyValuePair<double, Image>> dic = new Dictionary<int, KeyValuePair<double, Image>>();
            foreach (KeyValuePair<Image, string> itemZ in listImage)
            {
                try
                {
                    int psc = ConverterService.getDigit(itemZ.Value);

                    if (dic.TryGetValue(psc, out KeyValuePair<double, Image> pair))
                    {
                        dic[psc] = new KeyValuePair<double, Image>(double.NaN, itemZ.Key);
                    }
                    else
                    {
                        pair = new KeyValuePair<double, Image>(double.NaN, itemZ.Key);
                        dic.Add(psc, pair);
                    }
                }
                catch
                {
                    if (itemZ.Value.Split('/')[itemZ.Value.Split('/').Count() - 1].Contains("TONG"))
                    {
                        string l = item.Split('\\')[item.Split('\\').Count() - 3 + (pai.Equals("TONG") || pai.Contains("TRU") ? 0 : 1)];
                        string tape = item.Split('\\')[item.Split('\\').Count() - 2 + (pai.Equals("TONG") || pai.Contains("TRU") ? 0 : 1)];
                        _dicTONG.Add($"{l}_{maker} (TAPE {tape})", new KeyValuePair<double, byte[]>(double.NaN, TDMK_ImageConverter.ImageToByteArray(itemZ.Key, ImageFormat.Jpeg)));
                    }
                }
            }
            //get meansure 
            string[] listFile = FileFolderRepository.GetFileByExtension(item, "csv");
            foreach (string file in listFile)
            {
                try
                {

                    int psc = ConverterService.getDigit(file.Split('\\')[file.Split('\\').Count() - 1].Trim());

                    using (StreamReader reader = new StreamReader(file))
                    {
                        while (!reader.EndOfStream)
                        {
                            string dataLine = reader.ReadLine();
                            if (dataLine.Contains("Total area"))
                            {
                                string st = dataLine.Split(',')[1];
                                if (double.TryParse(st, out double d))
                                {
                                    d = Math.Round(d / 1000000, 2);
                                    dic[psc] = new KeyValuePair<double, Image>(d, dic[psc].Value);
                                }
                                break;
                            }
                        }
                    }
                }
                catch
                {
                    if (file.Split('\\')[file.Split('\\').Count() - 1].Trim().Contains("TONG"))
                    {
                        string l = item.Split('\\')[item.Split('\\').Count() - 3 + (pai.Equals("TONG") || pai.Contains("TRU") ? 0 : 1)];
                        string tape = item.Split('\\')[item.Split('\\').Count() - 2 + (pai.Equals("TONG") || pai.Contains("TRU") ? 0 : 1)];
                        using (StreamReader reader = new StreamReader(file))
                        {
                            while (!reader.EndOfStream)
                            {
                                string dataLine = reader.ReadLine();
                                if (dataLine.Contains("Total area"))
                                {
                                    string st = dataLine.Split(',')[1];
                                    if (double.TryParse(st, out double d))
                                    {
                                        try
                                        {
                                            d = Math.Round(d / 1000000, 2);
                                            _dicTONG[$"{l}_{maker} (TAPE {tape})"] = new KeyValuePair<double, byte[]>(d, _dicTONG[$"{l}_{maker} (TAPE {tape})"].Value);
                                        }
                                        catch
                                        {

                                        }
                                    }
                                    break;
                                }
                            }
                        }

                    }
                }
            }
            List<int> listKey = dic.Keys.ToList();
            listKey.Sort();
            int iz = dataTable.Rows.Count <= 0 ? 1 : 0;
            bool prime = dataTable.Rows.Count <= 0;
            foreach (int i in listKey)
            {
                if (prime)
                {
                    DataRow row = dataTable.NewRow();
                    row["ID"] = iz++;
                    if (pai.Equals("TONG"))
                    {
                        row["TONG"] = TDMK_ImageConverter.ImageToByteArray(dic[i].Value, ImageFormat.Jpeg);
                        row["Measure TONG (mm2)"] = dic[i].Key;
                    }
                    else if (pai.Contains("TRU"))
                    {
                        row["TRU"] = TDMK_ImageConverter.ImageToByteArray(dic[i].Value, ImageFormat.Jpeg);
                        row["Measure TRU (mm2)"] = dic[i].Key;
                    }
                    else
                    {
                        row["TONG"] = TDMK_ImageConverter.ImageToByteArray(dic[i].Value, ImageFormat.Jpeg);
                        row["Measure TONG (mm2)"] = dic[i].Key;
                    }
                    row["ItemCode"] = _itemCode;
                    row["LotNo"] = _lotNo;
                    row["Adhesive"] = "No Adhesive Squeeze Out";
                    row["Measure(mm)"] = "0";
                    dataTable.Rows.Add(row);
                }
                else
                {
                    if (pai.Equals("TONG"))
                    {
                        dataTable.Rows[iz]["TONG"] = TDMK_ImageConverter.ImageToByteArray(dic[i].Value, ImageFormat.Jpeg);
                        dataTable.Rows[iz]["Measure TONG (mm2)"] = dic[i].Key;
                    }
                    else if (pai.Contains("TRU"))
                    {
                        dataTable.Rows[iz]["TRU"] = TDMK_ImageConverter.ImageToByteArray(dic[i].Value, ImageFormat.Jpeg);
                        dataTable.Rows[iz]["Measure TRU (mm2)"] = dic[i].Key;
                    }
                    else
                    {
                        dataTable.Rows[iz]["TONG"] = TDMK_ImageConverter.ImageToByteArray(dic[i].Value, ImageFormat.Jpeg);
                        dataTable.Rows[iz]["Measure TONG (mm2)"] = dic[i].Key;

                    }
                    iz++;
                }
            }
        }


        private DataTable getDataInBF(string location)
        {
            IList<KeyValuePair<Image, string>> list = FileFolderRepository.ListAllPictureInAFolder(location, ".jpg");
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("ID");
            dataTable.Columns.Add("Image Before", typeof(Image));
            dataTable.Columns.Add("Image Before+", typeof(Image));
            Dictionary<int, KeyValuePair<Image, Image>> dic = new Dictionary<int, KeyValuePair<Image, Image>>();
            foreach (KeyValuePair<Image, string> item in list)
            {
                try
                {
                    bool prime = item.Value.Contains('+');
                    int psc = ConverterService.getDigit(item.Value);

                    if (dic.TryGetValue(psc, out KeyValuePair<Image, Image> pair))
                    {
                        if (prime)
                        {
                            dic[psc] = new KeyValuePair<Image, Image>(dic[psc].Key, item.Key);
                        }
                        else
                        {
                            dic[psc] = new KeyValuePair<Image, Image>(item.Key, dic[psc].Value);
                        }
                    }
                    else
                    {
                        if (prime)
                        {
                            pair = new KeyValuePair<Image, Image>(new Bitmap(1, 1), item.Key);
                        }
                        else
                        {
                            pair = new KeyValuePair<Image, Image>(item.Key, new Bitmap(1, 1));
                        }
                        dic.Add(psc, pair);
                    }
                }
                catch
                {

                }
            }
            List<int> listKey = dic.Keys.ToList();
            listKey.Sort();
            int iz = 1;
            foreach (int i in listKey)
            {
                DataRow row = dataTable.NewRow();
                row["ID"] = iz++;
                row["Image Before"] = dic[i].Key;
                row["Image Before+"] = dic[i].Value;
                dataTable.Rows.Add(row);

            }
            return dataTable;
        }
        public KeyValuePair<string, string> getItemCodeLotNo(string folderName)
        {
            string itemCode = "", lotNo = "";
            string[] s = folderName.Split('-');
            if (s.Length > 5)
            {
                itemCode = s[1];
                try
                {
                    lotNo = ValidateService.lotNoHandle($"{s[2]}-{ValidateService.isDigit(s[3])}");
                }
                catch
                {
                    lotNo = ValidateService.lotNoHandle($"{s[2]}");
                }

            }
            else
            {

            }
            return new KeyValuePair<string, string>(itemCode, lotNo);
        }
        private const string _NAMETABLE = "AIR_BUBBLE_BEFORE";
        public void SaveData(Guid? area = null, bool prime = false)
        {
            DataTable dataTable = new DataTable();
            DataTable dataTable_DATA = new DataTable();
            DataTable datatable_BEFORE = new DataTable();
            if (!prime)
            {
                dataTable = _dBContext.LoadDataTable(_NAMETABLE, new[] { "ItemCode", "LotNo" }, new[] { _itemCode, _lotNo });
                dataTable_DATA = _dBContext.LoadDataTable("AIR_BUBBLE_DATA", new[] { "ItemCode", "LotNo" }, new[] { _itemCode, _lotNo });
                datatable_BEFORE = _dBContext.LoadDataTable("AIR_BUBBLE_BEFORE", new[] { "ItemCode", "LotNo" }, new[] { _itemCode, _lotNo });

                if (dataTable.Rows.Count > 0 || dataTable_DATA.Rows.Count > 0 || datatable_BEFORE.Rows.Count > 0)
                {
                    throw new Exception($"1402 - Dữ liệu tồn tại bạn có muốn tiếp tục lưu?");
                }
            }
            else
            {
                dataTable = _dBContext.GetTableStructure(_NAMETABLE);
                dataTable_DATA = _dBContext.GetTableStructure("AIR_BUBBLE_DATA");
                datatable_BEFORE = _dBContext.GetTableStructure("AIR_BUBBLE_BEFORE");
            }
            int res = 0;
            fillPID(_beforeImage[$"{_itemCode} - {_lotNo}"], new[] { "ItemCode", "LotNo" }, new[] { _itemCode, _lotNo });
            convertBitmap(_beforeImage[$"{_itemCode} - {_lotNo}"], new[] { "Image Before", "Image Before+" });
            try
            {
                res = _dBContext.BuckDataTable(_beforeImage[$"{_itemCode} - {_lotNo}"], "AIR_BUBBLE_BEFORE_PIC", new[] { "ItemCode", "LotNo" }, null, "ID");
            }
            catch
            {
                _beforeImage[$"{_itemCode} - {_lotNo}"].Columns.Remove("ItemCode");
                _beforeImage[$"{_itemCode} - {_lotNo}"].Columns.Remove("LotNo");
            }

            foreach (string key in _dic.Keys)
            {
                DataTable dataZ = _dic[key];

                fillPID(dataZ, new[] { "ItemCode", "LotNo", "KeyDIC" }, new[] { _itemCode, _lotNo, key });
                foreach (DataRow row in dataZ.Rows)
                {
                    DataRow r = dataTable_DATA.NewRow();
                    foreach (DataColumn col in dataZ.Columns)
                    {
                        r[col.ColumnName] = row[col.ColumnName];
                    }
                    dataTable_DATA.Rows.Add(r);
                }
                //convertBitmap(dataZ, new[] { "TONG", "TRU" });
            }
            try
            {
                res += _dBContext.BuckDataTable(dataTable_DATA, "AIR_BUBBLE_DATA", new[] { "ItemCode", "LotNo" }, null, "ID");
            }
            catch
            {

            }
            foreach (string item in _dicTONG.Keys)
            {
                DataRow row = datatable_BEFORE.NewRow();
                row["ID"] = 0;
                row["ItemCode"] = _itemCode;
                row["LotNo"] = _lotNo;
                row["Picture"] = _dicTONG[item].Value;
                row["Data"] = _dicTONG[item].Key.ToString();
                row["KeyDic"] = item;
                datatable_BEFORE.Rows.Add(row);
            }
            res += _dBContext.BuckDataTable(datatable_BEFORE, "AIR_BUBBLE_BEFORE", new[] { "ItemCode", "LotNo" }, null, "ID");

            DataTable dtZ = _dBContext.LoadDataTable("AIR_BUBBLE_REFER", new[] { "ItemCodeRefer", "LotNoRefer" }, new[] { _itemCode, _lotNo });
            DataTable dtZ2 = _dBContext.LoadDataTable("AIR_BUBBLE_REFER", new[] { "ItemMain", "LotMain" }, new[] { _itemCode, _lotNo });
            if (dtZ.Rows.Count <= 0 && dtZ2.Rows.Count <= 0)
            {
                DataTable dt = _dBContext.GetTableStructure("AIR_BUBBLE_AVAILABLE");
                DataRow rowZ = dt.NewRow();
                rowZ["ItemCode"] = _itemCode;
                rowZ["LotNo"] = _lotNo;
                rowZ["Status"] = 1;
                dt.Rows.Add(rowZ);
                res += _dBContext.BuckDataTable(dt, "AIR_BUBBLE_AVAILABLE", new[] { "ItemCode", "LotNo" }, null, "ID");
            }
            throw new Exception($"{res} dòng đã lưu thành công");
        }
        private void convertBitmap(DataTable dataTable, string[] columnSel)
        {
            foreach (string col in columnSel)
            {
                if (dataTable.Columns.Contains(col))
                {
                    dataTable.Columns.Add($"{col}$CONVERTER", typeof(byte[]));
                }
            }
            foreach (DataRow row in dataTable.Rows)
            {
                foreach (string col in columnSel)
                {
                    row[$"{col}$CONVERTER"] = TDMK_ImageConverter.ImageToByteArray((Image)row[col], ImageFormat.Jpeg);
                }
            }
            foreach (string col in columnSel)
            {
                dataTable.Columns.Remove(col);
            }
        }
        private void fillPID(DataTable dataTable, string[] nameCol, string[] valueCol)
        {
            for (int i = 0; i < nameCol.Count(); i++)
            {
                // nếu đã có thì thêm 
                if (!dataTable.Columns.Contains(nameCol[i]))
                {
                    dataTable.Columns.Add(nameCol[i]);
                }
                foreach (DataRow row in dataTable.Rows)
                {
                    row[nameCol[i]] = valueCol[i];
                }



            }
        }
        public void LoadData()
        {
            _itemCode = _itemCode.Trim();
            _lotNo = _lotNo.Trim();

            if (string.IsNullOrEmpty(_itemCode) || string.IsNullOrEmpty(_lotNo))
            {
                throw new Exception("Item Code and Lot No cannot be empty.");
            }
            //AIR_BUBBLE_DATA
            DataTable dataTable_data = _dBContext.LoadDataTable("AIR_BUBBLE_DATA", new[] { "ItemCode", "LotNo" }, new[] { _itemCode, _lotNo });
            foreach (DataRow row in dataTable_data.Rows)
            {
                string keyDic = row["KeyDic"].ToString();
                if (_dic.ContainsKey(keyDic))
                {
                    DataRow newRow = _dic[keyDic].NewRow();
                    newRow.ItemArray = row.ItemArray.Clone() as object[];
                    _dic[keyDic].Rows.Add(newRow);
                }
                else
                {
                    // Clone the structure (columns, constraints) of dataTable_data but not the data
                    DataTable clonedTable = dataTable_data.Clone();
                    DataRow newRow = clonedTable.NewRow();
                    newRow.ItemArray = row.ItemArray.Clone() as object[];
                    clonedTable.Rows.Add(newRow);
                    _dic.Add(keyDic, clonedTable);
                }
            }
            //"AIR_BUBBLE_BEFORE"
            DataTable dataTable_Before = _dBContext.LoadDataTable("AIR_BUBBLE_BEFORE", new[] { "ItemCode", "LotNo" }, new[] { _itemCode, _lotNo });
            foreach (DataRow row in dataTable_Before.Rows)
            {
                string keyDic = row["KeyDic"].ToString();
                if (_dicTONG.ContainsKey(keyDic))
                {
                    _dicTONG[keyDic] = new KeyValuePair<double, byte[]>(Convert.ToDouble(row["Data"]), (byte[])row["Picture"]);
                }
                else
                {
                    _dicTONG.Add(keyDic, new KeyValuePair<double, byte[]>(Convert.ToDouble(row["Data"]), (byte[])row["Picture"]));
                }
            }
            //AIR_BUBBLE_BEFORE_PIC
            DataTable dataTable_Before_PIC = _dBContext.LoadDataTable("AIR_BUBBLE_BEFORE_PIC", new[] { "ItemCode", "LotNo" }, new[] { _itemCode, _lotNo });
            _beforeImage[$"{_itemCode} - {_lotNo}"] = dataTable_Before_PIC;
            return;
        }
        public void LoadData(string itemCode, string lotNo, bool isRefer = false)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();

            if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
            {
                throw new Exception("Item Code and Lot No cannot be empty.");
            }
            //AIR_BUBBLE_DATA
            DataTable dataTable_data = _dBContext.LoadDataTable("AIR_BUBBLE_DATA", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            foreach (DataRow row in dataTable_data.Rows)
            {
                string keyDic = row["KeyDic"].ToString() + (isRefer ? $"_Refer: {itemCode}-{lotNo}" : "");
                if (_dic.ContainsKey(keyDic))
                {
                    DataRow newRow = _dic[keyDic].NewRow();
                    newRow.ItemArray = row.ItemArray.Clone() as object[];
                    _dic[keyDic].Rows.Add(newRow);
                }
                else
                {
                    // Clone the structure (columns, constraints) of dataTable_data but not the data
                    DataTable clonedTable = dataTable_data.Clone();
                    DataRow newRow = clonedTable.NewRow();
                    newRow.ItemArray = row.ItemArray.Clone() as object[];
                    clonedTable.Rows.Add(newRow);
                    _dic.Add(keyDic, clonedTable);
                }
            }
            //"AIR_BUBBLE_BEFORE"
            if (!isRefer)
            {
                DataTable dataTable_Before = _dBContext.LoadDataTable("AIR_BUBBLE_BEFORE", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
                foreach (DataRow row in dataTable_Before.Rows)
                {
                    string keyDic = row["KeyDic"].ToString() + (isRefer ? $"_Refer: {itemCode}-{lotNo}" : "");
                    if (_dicTONG.ContainsKey(keyDic))
                    {
                        _dicTONG[keyDic] = new KeyValuePair<double, byte[]>(Convert.ToDouble(row["Data"]), (byte[])row["Picture"]);
                    }
                    else
                    {
                        _dicTONG.Add(keyDic, new KeyValuePair<double, byte[]>(Convert.ToDouble(row["Data"]), (byte[])row["Picture"]));
                    }
                }
            }
            else
            {

            }
            //AIR_BUBBLE_BEFORE_PIC
            DataTable dataTable_Before_PIC = _dBContext.LoadDataTable("AIR_BUBBLE_BEFORE_PIC", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            _beforeImage[$"{itemCode} - {lotNo}"] = dataTable_Before_PIC;
            return;
        }
        public List<KeyValuePair<string, string>> LoadRefer()
        {
            DataTable dataTable = _dBContext.LoadDataTable("AIR_BUBBLE_AVAILABLE", new[] { "Status" }, new[] { "1" });
            List<KeyValuePair<string, string>> list = new List<KeyValuePair<string, string>>();
            foreach (DataRow row in dataTable.Rows)
            {
                string itemCode = row["ItemCode"].ToString().Trim();
                string lotNo = row["LotNo"].ToString().Trim();
                if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
                {
                    throw new Exception("itemcode lotno cant not null");
                }
                if (!itemCode.Contains(_itemCode) || !lotNo.Contains(_lotNo))
                {
                    list.Add(new KeyValuePair<string, string>(itemCode.Trim(), lotNo.Trim()));
                }
            }
            return list;
        }
        public List<KeyValuePair<string, string>> LoadRefered()
        {
            DBContext dBContext = new DBContext();
            DataTable dataTable = dBContext.LoadDataTable("AIR_BUBBLE_REFER", new[] { "ItemMain", "LotMain" }, new[] { _itemCode, _lotNo });
            List<KeyValuePair<string, string>> list = new List<KeyValuePair<string, string>>();
            foreach (DataRow row in dataTable.Rows)
            {
                list.Add(new KeyValuePair<string, string>(row["ItemCodeRefer"].ToString().Trim(), row["LotNoRefer"].ToString().Trim()));
            }
            listRefer = list;
            return list;
        }

        public void AddRefer(KeyValuePair<string, string> keyValuePair)
        {
            string itemCode = keyValuePair.Key;
            string lotNo = keyValuePair.Value;
            _dBContext.DeleteData("AIR_BUBBLE_AVAILABLE", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            string json = "";
            DataTable dataTable = _dBContext.LoadDataTable("AIR_BUBBLE_REFER", new[] { "ItemMain", "LotMain" }, new[] { _itemCode, _lotNo });

            DataRow row = dataTable.NewRow();
            row["ID"] = 1;
            row["ItemMain"] = _itemCode;
            row["LotMain"] = _lotNo;
            row["ItemCodeRefer"] = itemCode;
            row["LotNoRefer"] = lotNo;
            dataTable.Rows.Add(row);

            _dBContext.BuckDataTable(dataTable, "AIR_BUBBLE_REFER", new[] { "ItemMain", "LotMain" }, null, "ID");
        }
        private DataTable dt_refer;

        private void ExportLinerSheet(ExcelWorksheet ws, string refer)
        {
            string[] listDic = _dic.Keys.ToArray();
            List<string> finder = new List<string>();
            foreach (string item in listDic)
            {
                if (item.Contains(refer.ToUpper()))
                {
                    string maker = item.Split('_', '(')[1].Trim();
                    string tape = item.Split('_', '(', ')')[2].Trim();

                    string key = $"{maker} supplier ({tape})";
                    finder.Add(key);
                }
            }
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(ws, finder.Concat(new[] { "Air bubble", "area:" }).ToArray());

            int r = dataComment.Rows.Count;
            int i = 0;
            foreach (string key in _dicTONG.Keys)
            {
                if (key.Contains(refer.ToUpper()) && !key.Contains("Refer"))
                {
                    if (dic.TryGetValue("area:", out string add))
                    {
                        string address = add.Split('-')[i];
                        i++;
                        string[] value = ws.Cells[address].Value.ToString().Split(':');
                        double fo = _dicTONG[key].Key;
                        ws.Cells[address].Value = $"{value[0]}: {fo} {value[1]}";
                        address = ExportProcess.AddRow(address, 1);
                        ExportProcess.InsertImageToCell(ws, ws.Cells[address], _dicTONG[key].Value, $"{Guid.NewGuid()}");
                    }
                }
            }

            foreach (string key in _dic.Keys)
            {
                string maker = key.Split('_', '(')[1].Trim();
                string tape = key.Split('_', '(', ')')[2].Trim();
                string keyX = $"{maker} supplier ({tape})";
                string comment1 = "";
                string comment2 = "";
                bool primeCmt = false;
                if (dic.TryGetValue(keyX, out string address))
                {
                    DataTable dataTable = _dic[key];
                    string itemCode = _itemCode, lotNo = _lotNo;
                    if (key.Contains("Refer"))
                    {
                        itemCode = key.Split(':')[1].Split('-')[0].Trim();
                        lotNo = key.Split(':')[1].Split('-')[1].Trim();
                    }
                    else
                    {
                        itemCode = _itemCode;
                        lotNo = _lotNo;
                    }
                    foreach (DataRow row in dataComment.Rows)
                    {
                        if (row["TapeName"].ToString().Equals(tape))
                        {

                            if (row["ItemCode"].ToString().Trim().Equals(itemCode))
                            {
                                comment1 = row["Comment1"].ToString();
                                comment2 = row["Comment2"].ToString();
                                primeCmt = true;
                                break;
                            }
                            //Debugger.Break();
                        }
                    }

                    for (int iz = 0; iz < 32; iz++)
                    {
                        string addressSample = ExportProcess.AddColumn(address, iz + 2);
                        if (iz < _beforeImage[$"{itemCode} - {lotNo}"].Rows.Count)
                        {
                            addressSample = ExportProcess.AddRow(addressSample, -1);
                            ws.Cells[addressSample].Value = _beforeImage[$"{itemCode} - {lotNo}"].Rows[iz]["ProductID"];

                            addressSample = ExportProcess.AddRow(addressSample, 2);
                            ExportProcess.InsertImageToCell(ws, ws.Cells[addressSample], (byte[])_beforeImage[$"{itemCode} - {lotNo}"].Rows[iz]["Image Before$CONVERTER"], $"{Guid.NewGuid()}");

                        }
                        if (iz < dataTable.Rows.Count)
                        {

                            addressSample = ExportProcess.AddRow(addressSample, 1);
                            ExportProcess.InsertImageToCell(ws, ws.Cells[addressSample], (byte[])dataTable.Rows[iz]["TONG"], $"{Guid.NewGuid()}");

                            addressSample = ExportProcess.AddRow(addressSample, 1);
                            if (double.TryParse(dataTable.Rows[iz]["Measure TONG (mm2)"].ToString(), out double res))
                            {
                                ws.Cells[addressSample].Value = res;
                            }
                            addressSample = ExportProcess.AddRow(addressSample, 1);
                            if (double.TryParse(dataTable.Rows[iz]["% Area TONG"].ToString(), out res))
                            {
                                if (_dicTONG.TryGetValue(key, out KeyValuePair<double, byte[]> pair))
                                {
                                    ws.Cells[addressSample].FormulaR1C1 = $"=R[-1]C / {pair.Key}";
                                }
                                else
                                {
                                    ws.Cells[addressSample].Value = res;

                                }
                                ws.Cells[addressSample].Style.Numberformat.Format = "0.00%";
                            }
                            if (primeCmt)
                            {
                                addressSample = ExportProcess.AddRow(addressSample, 1);
                                try
                                {
                                    addressSample = ExportProcess.AddRow(addressSample, 6);
                                    try
                                    {

                                        ExportProcess.InsertImageToCell(ws, ws.Cells[addressSample], (byte[])dataTable.Rows[iz]["TRU"], $"{Guid.NewGuid()}");
                                    }
                                    catch
                                    {
                                        ws.Cells[addressSample].Value = "No Air bubble";
                                    }
                                    addressSample = ExportProcess.AddRow(addressSample, 1);
                                    if (double.TryParse(dataTable.Rows[iz]["Measure TRU (mm2)"].ToString(), out res))
                                    {
                                        ws.Cells[addressSample].Value = res;
                                    }
                                    addressSample = ExportProcess.AddRow(addressSample, 1);
                                    if (double.TryParse(dataTable.Rows[iz]["% Area TRU"].ToString(), out res))
                                    {
                                        if (_dicTONG.TryGetValue(key, out KeyValuePair<double, byte[]> pair))
                                        {
                                            ws.Cells[addressSample].FormulaR1C1 = $"=R[-1]C/{pair.Key}";
                                        }
                                        else
                                        {
                                            ws.Cells[addressSample].Value = res;
                                        }
                                        ws.Cells[addressSample].Style.Numberformat.Format = "0.00%";
                                    }
                                }
                                catch
                                {
                                    addressSample = ExportProcess.AddRow(addressSample, 2);
                                }

                            }
                            addressSample = ExportProcess.AddRow(addressSample, 1);
                            ws.Cells[addressSample].Value = dataTable.Rows[iz]["Adhesive"];
                            addressSample = ExportProcess.AddRow(addressSample, 1);
                            try
                            {

                                if (double.TryParse(dataTable.Rows[iz]["Measure(mm)"].ToString(), out double result))
                                {
                                    ws.Cells[addressSample].Value = result;
                                }
                            }
                            catch
                            {

                            }
                            addressSample = ExportProcess.AddRow(addressSample, 1);
                            ws.Cells[addressSample].Value = dataTable.Rows[iz]["Judgement"];
                        }
                    }
                }
                if (!primeCmt)
                {
                    string adss = ExportProcess.AddRow(ExportProcess.AddColumn(address, 2), 11);
                    ws.Cells[adss].FormulaR1C1 = $"MIN(R[-8]C:R[-8]C[31])";
                    adss = ExportProcess.AddColumn(adss, 1);
                    ws.Cells[adss].FormulaR1C1 = $"MIN(R[-7]C[-1]:R[-7]C[30])";
                    adss = ExportProcess.AddColumn(adss, 1);
                    ws.Cells[adss].FormulaR1C1 = $"MIN(R[-5]C[-2]:R[-5]C[29])";
                    adss = ExportProcess.AddRow(adss, 1);
                    ws.Cells[adss].FormulaR1C1 = $"MAX(R[-6]C[-2]:R[-6]C[29])";
                    adss = ExportProcess.AddColumn(adss, -1);
                    ws.Cells[adss].FormulaR1C1 = $"MAX(R[-8]C[-1]:R[-8]C[30])";
                    adss = ExportProcess.AddColumn(adss, -1);
                    ws.Cells[adss].FormulaR1C1 = $"MAX(R[-9]C:R[-9]C[31])";
                    adss = ExportProcess.AddRow(adss, 1);
                    ws.Cells[adss].FormulaR1C1 = $"Average(R[-10]C:R[-10]C[31])";
                    adss = ExportProcess.AddColumn(adss, 1);
                    ws.Cells[adss].FormulaR1C1 = $"Average(R[-9]C[-1]:R[-9]C[30])";
                    adss = ExportProcess.AddColumn(adss, 1);
                    ws.Cells[adss].FormulaR1C1 = $"Average(R[-7]C[-2]:R[-7]C[29])";
                    //Debugger.Break();
                }
                else
                {
                    string adss = ExportProcess.AddRow(ExportProcess.AddColumn(address, 2), 8);
                    ws.Cells[adss].FormulaR1C1 = $"MIN(R[-5]C:R[-5]C[31])";
                    adss = ExportProcess.AddColumn(adss, 1);
                    ws.Cells[adss].FormulaR1C1 = $"MIN(R[-4]C[-1]:R[-4]C[30])";
                    adss = ExportProcess.AddColumn(adss, 1);
                    //ws.Cells[adss].FormulaR1C1 = $"MIN(R[-5]C[-2]:R[-5]C[29])";
                    adss = ExportProcess.AddRow(adss, 1);
                    //ws.Cells[adss].FormulaR1C1 = $"MAX(R[-6]C[-2]:R[-6]C[29])";
                    adss = ExportProcess.AddColumn(adss, -1);
                    ws.Cells[adss].FormulaR1C1 = $"MAX(R[-5]C[-1]:R[-5]C[30])";
                    adss = ExportProcess.AddColumn(adss, -1);
                    ws.Cells[adss].FormulaR1C1 = $"MAX(R[-6]C:R[-6]C[31])";
                    adss = ExportProcess.AddRow(adss, 1);
                    ws.Cells[adss].FormulaR1C1 = $"Average(R[-7]C:R[-7]C[31])";
                    adss = ExportProcess.AddColumn(adss, 1);
                    ws.Cells[adss].FormulaR1C1 = $"Average(R[-6]C[-1]:R[-6]C[30])";
                    adss = ExportProcess.AddColumn(adss, 1);
                    //ws.Cells[adss].FormulaR1C1 = $"Average(R[-7]C[-2]:R[-7]C[29])";
                    string adsz = ExportProcess.AddRow(ExportProcess.AddColumn(address, 2), 20);
                    ws.Cells[adsz].FormulaR1C1 = $"MIN(R[-8]C:R[-8]C[31])";
                    adsz = ExportProcess.AddColumn(adsz, 1);
                    ws.Cells[adsz].FormulaR1C1 = $"MIN(R[-7]C[-1]:R[-7]C[30])";
                    adsz = ExportProcess.AddColumn(adsz, 1);
                    ws.Cells[adsz].FormulaR1C1 = $"MIN(R[-5]C[-2]:R[-5]C[29])";
                    adsz = ExportProcess.AddRow(adsz, 1);
                    ws.Cells[adsz].FormulaR1C1 = $"MAX(R[-6]C[-2]:R[-6]C[29])";
                    adsz = ExportProcess.AddColumn(adsz, -1);
                    ws.Cells[adsz].FormulaR1C1 = $"MAX(R[-8]C[-1]:R[-8]C[30])";
                    adsz = ExportProcess.AddColumn(adsz, -1);
                    ws.Cells[adsz].FormulaR1C1 = $"MAX(R[-9]C:R[-9]C[31])";
                    adsz = ExportProcess.AddRow(adsz, 1);
                    ws.Cells[adsz].FormulaR1C1 = $"Average(R[-10]C:R[-10]C[31])";
                    adsz = ExportProcess.AddColumn(adsz, 1);
                    ws.Cells[adsz].FormulaR1C1 = $"Average(R[-9]C[-1]:R[-9]C[30])";
                    adsz = ExportProcess.AddColumn(adsz, 1);
                    ws.Cells[adsz].FormulaR1C1 = $"Average(R[-7]C[-2]:R[-7]C[29])";
                }
            }

        }


        public void Export()
        {

            LoadDataRefer();

            string location;
            ExportProcess exportProcess = new ExportProcess();

            using (ExcelPackage package = exportProcess.FindFormatProcess("AIR", _itemCode, _lotNo))
            {
                using (ExcelWorksheet ws = exportProcess.FindSheet(package, "Air bubble btw Liner-PSA"))
                {
                    //set up liner
                    ExportSetupLiner(ws, "Liner");
                    //Export liner
                    ExportLinerSheet(ws, "Liner");
                    using (ExcelWorksheet ws1 = exportProcess.FindSheet(package, "Air bubble btw PSA-FPC"))
                    {

                        ExportSetupLiner(ws1, "PSA");
                        ExportPSASheet(ws1, "PSA");
                        exportProcess.SaveExcelWorksheet(package, "Air bubble btw Liner-PSA:Air bubble btw PSA-FPC", $"{_itemCode.Trim()}-{_lotNo.Trim()}");




                        //save and open
                    }
                }
            }
        }

        private void ExportPSASheet(ExcelWorksheet ws, string type)
        {
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(ws, new[] { "supplier", "Air bubble", "area:" });
            List<string> listZ = new List<string>();
            int i = 0;
            foreach (string key in _dicTONG.Keys)
            {
                if (key.Contains(type.ToUpper()) && !key.Contains("Refer"))
                {
                    if (dic.TryGetValue("area:", out string add))
                    {
                        string address = add.Split('-')[i];
                        i++;
                        string[] value = ws.Cells[address].Value.ToString().Split(':');
                        double fo = _dicTONG[key].Key;
                        ws.Cells[address].Value = $"{value[0]}: {fo} {value[1]}";
                        address = ExportProcess.AddRow(address, 1);
                        ExportProcess.InsertImageToCell(ws, ws.Cells[address], _dicTONG[key].Value, $"{Guid.NewGuid()}");
                    }
                }
            }
            int rAir = ws.Cells[dic["Air bubble"].Split('-')[0]].End.Row;
            foreach (string add in dic["supplier"].Split('-'))
            {
                int row = ws.Cells[add].End.Row;
                if (row > rAir)
                {
                    listZ.Add(add);
                }
            }
            dic["supplier"] = string.Join("-", listZ);
            int a = dataComment.Rows.Count;
            string[] list = _dic.Keys.ToArray();
            foreach (string addressRoot in dic["supplier"].Split('-'))
            {
                string valueRoot = ws.Cells[addressRoot].Value.ToString();
                foreach (string item in list)
                {
                    string typeZ = item.Split('_')[0].Trim();
                    if (typeZ.ToUpper().Equals(type.ToUpper()))
                    {
                        string[] listRoot = valueRoot.Replace(")", "").Replace("supplier", "").Split('_', '(');
                        string[] listItem = item.Replace(")", "").Replace("supplier", "").Split('_', '(');
                        if (listRoot[0].Contains(listItem[0]) && listRoot[1].Trim().Contains(listItem[1].Trim()) && listRoot[2].Trim().Contains(listItem[2].Trim()))
                        {
                            //Debugger.Break();
                            bool primeComment = false;
                            string itemCode = _dic[item].Rows[0]["ItemCode"].ToString().Trim();
                            string lotno = _dic[item].Rows[0]["LotNo"].ToString().Trim();
                            foreach (DataRow row in dataComment.Rows)
                            {

                                if (listRoot[2].Trim().Contains(row["TapeName"].ToString()) && itemCode.Contains(row["ItemCode"].ToString().Trim()))
                                {
                                    primeComment = true;
                                    break;
                                }
                            }
                            ExportAPart(ws, _dic[item], primeComment, _beforeImage[$"{itemCode} - {lotno}"], addressRoot);
                            break;
                        }
                    }
                }
            }
        }
        private void ExportAPart(ExcelWorksheet ws, DataTable dataTable, bool commentPrime, DataTable before, string address)
        {
            int i = 0;
            address = ExportProcess.AddColumn(address, 2);
            foreach (DataRow row in dataTable.Rows)
            {
                string addressZ = ExportProcess.AddColumn(address, i++);
                if (before.Rows.Count >= i)
                {
                    ws.Cells[ExportProcess.AddRow(addressZ, -1)].Value = before.Rows[i - 1]["ProductID"];
                    ExportProcess.InsertImageToCell(ws, ws.Cells[ExportProcess.AddRow(addressZ, 1)], (byte[])before.Rows[i - 1]["Image Before$CONVERTER"], $"{Guid.NewGuid()}");
                    ExportProcess.InsertImageToCell(ws, ws.Cells[ExportProcess.AddRow(addressZ, 2)], (byte[])before.Rows[i - 1]["Image Before+$CONVERTER"], $"{Guid.NewGuid()}");
                }
                ExportProcess.InsertImageToCell(ws, ws.Cells[ExportProcess.AddRow(addressZ, 3)], (byte[])row["TONG"], $"{Guid.NewGuid()}");
                if (double.TryParse(row["Measure TONG (mm2)"].ToString(), out double num))
                {
                    ws.Cells[ExportProcess.AddRow(addressZ, 4)].Value = num;
                }
                string addressTong = ExportProcess.AddRow(addressZ, 5);
                if (double.TryParse(row["% Area TONG"].ToString(), out num))
                {
                    ws.Cells[ExportProcess.AddRow(addressZ, 5)].Value = num / 100;
                    ws.Cells[ExportProcess.AddRow(addressZ, 5)].Style.Numberformat.Format = "0.00%";
                }
                if (commentPrime)
                {
                    addressZ = ExportProcess.AddRow(addressZ, 5);
                    try
                    {
                        ExportProcess.InsertImageToCell(ws, ws.Cells[ExportProcess.AddRow(addressZ, 6)], (byte[])row["TRU"], $"{Guid.NewGuid()}");
                        if (double.TryParse(row["Measure TRU (mm2)"].ToString(), out num))
                        {
                            ws.Cells[ExportProcess.AddRow(addressZ, 7)].Value = num;
                        }
                        if (double.TryParse(row["% Area TRU"].ToString(), out num))
                        {
                            ws.Cells[ExportProcess.AddRow(addressZ, 8)].Value = num / 100;
                        }
                    }
                    catch
                    {

                    }
                    addressZ = ExportProcess.AddRow(addressZ, 3);
                }
                int rDistance = ws.Cells[ExportProcess.AddRow(addressZ, 6)].End.Row - ws.Cells[addressTong].End.Row;
                ws.Cells[ExportProcess.AddRow(addressZ, 6)].FormulaR1C1 = $"=if(R[{rDistance}]C>10, \"NG\", \"OK\")";
            }
            if (commentPrime)
            {
                ws.Cells[ExportProcess.AddRow(address, 8)].FormulaR1C1 = $"=MIN(R[-4]C:R[-4]C[31])";
                ws.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 8), 1)].FormulaR1C1 = $"=MIN(R[-3]C[-1]:R[-3]C[30])";

                ws.Cells[ExportProcess.AddRow(address, 9)].FormulaR1C1 = $"=MAX(R[-5]C:R[-5]C[31])";
                ws.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 9), 1)].FormulaR1C1 = $"=MAX(R[-4]C[-1]:R[-4]C[30])";

                ws.Cells[ExportProcess.AddRow(address, 10)].FormulaR1C1 = $"=AVERAGE(R[-6]C:R[-6]C[31])";
                ws.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 10), 1)].FormulaR1C1 = $"=AVERAGE(R[-5]C[-1]:R[-5]C[30])";
                address = ExportProcess.AddRow(address, 8);
            }
            ws.Cells[ExportProcess.AddRow(address, 9)].FormulaR1C1 = $"=MIN(R[-5]C:R[-5]C[31])";
            ws.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 9), 1)].FormulaR1C1 = $"=MIN(R[-4]C[-1]:R[-4]C[30])";

            ws.Cells[ExportProcess.AddRow(address, 10)].FormulaR1C1 = $"=MAX(R[-6]C:R[-6]C[31])";
            ws.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 10), 1)].FormulaR1C1 = $"=MAX(R[-5]C[-1]:R[-5]C[30])";

            ws.Cells[ExportProcess.AddRow(address, 11)].FormulaR1C1 = $"=AVERAGE(R[-7]C:R[-7]C[31])";
            ws.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 11), 1)].FormulaR1C1 = $"=AVERAGE(R[-6]C[-1]:R[-6]C[30])";
        }
        private void ExportSetupLiner(ExcelWorksheet ws, string type)
        {
            DataTable dt = _dBContext.LoadDataTable("AIR_BUBBLE_COMMENT", new[] { "ItemCode" }, new[] { _itemCode });
            foreach (KeyValuePair<string, string> item in listRefer)
            {
                DataTable dz = _dBContext.LoadDataTable("AIR_BUBBLE_COMMENT", new[] { "ItemCode" }, new[] { item.Key });
                if (dz.Rows.Count > 0)
                {
                    DataRow row = dt.NewRow();
                    row["Comment1"] = dz.Rows[0]["Comment1"];
                    row["Comment2"] = dz.Rows[0]["Comment2"];
                    row["ItemCode"] = dz.Rows[0]["ItemCode"];
                    row["TapeName"] = dz.Rows[0]["TapeName"];
                    dt.Rows.Add(row);
                }
            }
            IDictionary<string, string> dicZ = ExportProcess.FindAddressByText(ws, new[] { "Average", "Flex SN", "Air bubble", "supplier", "Sample 32" });
            IDictionary<string, string> dic = new Dictionary<string, string>();
            int rCompare = ws.Cells[dicZ["Air bubble"].Split('-')[0]].End.Row;
            foreach (string item in dicZ.Keys)
            {
                if (!item.Equals("Air bubble"))
                {
                    string[] jtem = dicZ[item].Split('-');
                    List<string> newValue = new List<string>();
                    foreach (string s in jtem)
                    {
                        int r = ws.Cells[s].End.Row;
                        if (rCompare < r)
                        {
                            newValue.Add(s);
                        }
                    }
                    dic.Add(item, string.Join("-", newValue));
                }
            }
            //<[Tape name], [Supplier] - [FlexSn] - [Sample] - [Average]>
            Dictionary<string, string> dicDuplicateList = new Dictionary<string, string>();


            foreach (string item in dic["Flex SN"].Split('-'))
            {
                string value = "";
                int rFlexSN = ws.Cells[item].End.Row;

                string addressPeek = "";

                string[] sup = dic["supplier"].Split('-');
                int i = 0;
                int point = int.MaxValue;
                while (i < sup.Count())
                {
                    int rSup = ws.Cells[sup[i]].End.Row;
                    if (rSup > rFlexSN)
                    {
                        if (point > rSup)
                        {
                            point = rSup;
                            addressPeek = sup[i];
                        }
                    }
                    i++;
                }

                string addressSample = "";
                point = int.MaxValue;
                i = 0;
                string[] supZ = dic["Sample 32"].Split('-');
                while (i < supZ.Count())
                {

                    int rSup = ws.Cells[supZ[i]].End.Row;
                    if (rSup > rFlexSN)
                    {
                        if (point > rSup)
                        {
                            point = rSup;
                            addressSample = supZ[i];
                        }
                    }
                    i++;
                }


                string addressAverage = "";
                point = int.MaxValue;
                i = 0;
                string[] supA = dic["Average"].Split('-');
                while (i < supA.Count())
                {
                    int rSup = ws.Cells[supA[i]].End.Row;
                    if (rSup > rFlexSN)
                    {
                        if (point > rSup)
                        {
                            point = rSup;
                            addressAverage = supA[i];
                        }
                    }
                    i++;
                }
                string tapeName = ws.Cells[addressPeek].Value.ToString().Split('(')[1].TrimEnd(')');

                value = $"{addressPeek} - {item} - {addressSample} - {addressAverage}";
                dicDuplicateList.Add(tapeName, value);
            }
            Dictionary<string, string> dicz = new Dictionary<string, string>();
            foreach (string item in _dic.Keys)
            {
                if (item.ToUpper().Contains(type.ToUpper()) && item.ToUpper().Contains("Refer".ToUpper()))
                {
                    foreach (string tapeName in dicDuplicateList.Keys)
                    {
                        if (item.ToUpper().Contains(tapeName.ToUpper()))
                        {
                            if (dicz.TryGetValue(tapeName, out string count))
                            {
                                dicz[tapeName] = count + $",{item}";
                            }
                            else
                            {
                                dicz.Add(tapeName, $"{item}");
                            }
                        }
                    }
                }
            }
            ExportProcess ep = new ExportProcess();

            List<string> listKey = dicDuplicateList.Keys.ToList();
            listKey.Reverse();
            // Duplicatate
            foreach (string item in listKey)
            {
                if (dicz.TryGetValue(item, out string count))
                {
                    foreach (string jtem in count.Split(','))
                    {
                        string[] value = dicDuplicateList[item].Split('-');

                        int rS = ws.Cells[value[0].Trim()].Start.Row - ws.Cells[value[1].Trim()].Start.Row;
                        int cS = ws.Cells[value[0].Trim()].Start.Column - ws.Cells[value[1].Trim()].Start.Column;

                        string name = $"{jtem.Split('_')[0]}_{jtem.Split('_')[1]}";
                        //Debugger.Break();
                        string range = $"{value[1].Trim()}:{ws.Cells[ws.Cells[value[3].Trim()].End.Row, ws.Cells[value[2].Trim()].End.Column + 3].Address}".Trim();
                        string dist = ExportProcess.AddColumn(ExportProcess.AddRow(value[3].Trim(), 1), -1);
                        string addressDist = ExportProcess.AddRow(dist, 2);
                        ep.CopyAndInsert(ws, range, ref dist, true);
                        if (ws.Cells[addressDist].Value != null)
                        {
                            string valueZ = ws.Cells[addressDist].Value.ToString();
                            string vali = $"{valueZ.Split('_')[0]}_ {name.Split('_')[1].Split('(')[0].Trim()} {valueZ.Split('_')[1]}";
                            ws.Cells[addressDist].Value = vali;
                        }
                    }
                }
            }
            LoadComment();
            IDictionary<string, string> dicZZ = ExportProcess.FindAddressByText(ws, new[] { "supplier" }, false, dicZ["Air bubble"].Split('-')[0], "CC200");
            listKey = dicZZ["supplier"].Split('-').ToList();
            listKey.Reverse();
            foreach (var item in listKey)
            {
                string valuZ = ws.Cells[item].Value.ToString();
                string tape = valuZ.Split('(')[1].TrimEnd(')');
                string maker = valuZ.Replace("supplier", "").Replace(" ", "").Split('(', '_')[1];
                bool refer = !string.IsNullOrEmpty(maker);

                foreach (var z in _dic.Keys)
                {
                    if (z.Contains(tape) && z.Contains(type.ToUpper()) && z.Contains(maker.ToUpper()) && refer == z.Contains("Refer"))
                    {
                        foreach (DataRow rowItem in dataComment.Rows)
                        {
                            string itemCode;
                            if (refer)
                            {
                                itemCode = z.Split(':')[1].Split('-')[0].Trim();
                            }
                            else
                            {
                                itemCode = _itemCode;
                                if (!refer)
                                {
                                    string makerNew = z.Split('_', '(')[1].Trim();
                                    ws.Cells[item].Value = $"{valuZ.Split('_')[0]}_{makerNew} {valuZ.Split('_')[1]}";
                                }
                            }
                            if (rowItem["ItemCode"].ToString().Trim().Equals(itemCode) && rowItem["TapeName"].ToString().Trim().Equals(tape))
                            {
                                if (type.Equals("PSA"))
                                {
                                    string address = ExportProcess.AddColumn(ExportProcess.AddRow(item, 6), 1);
                                    //Converter AirBubble
                                    ws.InsertRow(ws.Cells[address].End.Row, 8);
                                    string range = $"{ExportProcess.AddRow(address, 0)}:{ExportProcess.AddRow(ExportProcess.AddColumn(address, 40), 4)}";
                                    string rangeC = $"{ExportProcess.AddRow(address, 9)}:{ExportProcess.AddRow(ExportProcess.AddColumn(address, 40), 14)}";
                                    CopyRow(ws, rangeC, range);
                                    ////vhx
                                    range = $"{ExportProcess.AddRow(address, 5)}:{ExportProcess.AddRow(ExportProcess.AddColumn(address, 40), 5)}";
                                    rangeC = $"{ExportProcess.AddRow(address, -5)}:{ExportProcess.AddRow(ExportProcess.AddColumn(address, 40), -5)}";
                                    CopyRow(ws, rangeC, range);

                                    ////meansure - area
                                    range = $"{ExportProcess.AddRow(address, 6)}:{ExportProcess.AddRow(ExportProcess.AddColumn(address, 40), 7)}";
                                    rangeC = $"{ExportProcess.AddRow(address, -2)}:{ExportProcess.AddRow(ExportProcess.AddColumn(address, 40), -1)}";
                                    CopyRow(ws, rangeC, range);


                                    for (int i = 0; i < 10; i++)
                                    {
                                        string addressZ = ExportProcess.AddRow(address, i);
                                        ws.Cells[addressZ].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                        ws.Cells[addressZ].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    }
                                }
                                if (type.Equals("Liner"))
                                {

                                    //LINER
                                    string address = ExportProcess.AddColumn(ExportProcess.AddRow(item, 5), 1);
                                    //Converter AirBubble
                                    ws.InsertRow(ws.Cells[address].End.Row, 9);
                                    string range = $"{address}:{ExportProcess.AddRow(ExportProcess.AddColumn(address, 40), 5)}";
                                    string rangeC = $"{ExportProcess.AddRow(address, 12)}:{ExportProcess.AddRow(ExportProcess.AddColumn(address, 40), 17)}";
                                    CopyRow(ws, rangeC, range);

                                    range = $"{ExportProcess.AddRow(address, 6)}:{ExportProcess.AddRow(ExportProcess.AddColumn(address, 40), 9)}";
                                    rangeC = $"{ExportProcess.AddRow(address, -3)}:{ExportProcess.AddRow(ExportProcess.AddColumn(address, 40), -1)}";
                                    CopyRow(ws, rangeC, range);

                                    for (int i = 0; i < 10; i++)
                                    {
                                        string addressZ = ExportProcess.AddRow(address, i);
                                        ws.Cells[addressZ].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                        ws.Cells[addressZ].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    }
                                }
                                break;
                            }
                        }

                    }
                }
            }
        }
        private void CopyRow(ExcelWorksheet workSheet, string rangeForm, string rangeTo)
        {
            workSheet.Cells[rangeForm].Copy(workSheet.Cells[rangeTo]);
            workSheet.Cells[rangeForm].CopyStyles(workSheet.Cells[rangeTo]);

            int rForm = workSheet.Cells[rangeForm.Split(':')[0]].End.Row;
            int rTo = workSheet.Cells[rangeForm.Split(':')[1]].End.Row;
            int rForm1 = workSheet.Cells[rangeTo.Split(':')[0]].End.Row;
            int z = rTo - rForm;
            for (int i = 0; i <= z; i++)
            {
                ExportProcess.CopyRowStyle(workSheet, rForm + i, rForm1 + i);
            }
        }
        DataTable dataComment;
        private void LoadComment()
        {
            DataTable dataTable = _dBContext.LoadDataTable("AIR_BUBBLE_COMMENT", new[] { "ItemCode" }, new[] { _itemCode });
            foreach (string itemCode in ListComment)
            {
                DataTable dataTableDT = _dBContext.LoadDataTable("AIR_BUBBLE_COMMENT", new[] { "ItemCode" }, new[] { itemCode });
                foreach (DataRow row in dataTableDT.Rows)
                {
                    DataRow rowNew = dataTable.NewRow();
                    foreach (DataColumn col in dataTable.Columns)
                    {
                        rowNew[col.ColumnName] = row[col.ColumnName];
                    }
                    dataTable.Rows.Add(rowNew);
                }
            }
            dataComment = dataTable;
        }
        public KeyValuePair<string, string> LoadComment(string itemCode, string tapeName)
        {
            DataTable dataTable = _dBContext.LoadDataTable("AIR_BUBBLE_COMMENT", new[] { "ItemCode", "TapeName" }, new[] { itemCode, tapeName });
            if (dataTable.Rows.Count < 0)
            {
                throw new Exception("ItemCode Tape have not comment");
            }
            return new KeyValuePair<string, string>(dataTable.Rows[0]["Comment1"].ToString(), dataTable.Rows[0]["Comment2"].ToString());
        }

        public void SaveComment(string itemCode, string tape, string cmt1, string cmt2, bool v = false)
        {
            DataTable dataTable = null;
            if (!v)
            {
                dataTable = _dBContext.LoadDataTable("AIR_BUBBLE_COMMENT", new[] { "ItemCode", "TapeName" }, new[] { itemCode, tape });
                if (dataTable.Rows.Count > 0)
                {
                    throw new Exception("1234-Đã tồn tại comment! Bạn có muốn lưu đè?");
                }
            }
            else
            {
                dataTable = _dBContext.GetTableStructure("AIR_BUBBLE_COMMENT");
            }
            DataRow row = dataTable.NewRow();
            row["ItemCode"] = itemCode;
            row["TapeName"] = tape;
            row["Comment1"] = cmt1;
            row["Comment2"] = cmt2;
            dataTable.Rows.Add(row);
            _dBContext.BuckDataTable(dataTable, "AIR_BUBBLE_COMMENT", new[] { "ItemCode", "TapeName" }, null, "Id");
            throw new Exception("Lưu thành công!");
        }
        private List<string> ListComment = new List<string>();
        public void CalculateArea(DataTable dataTable, KeyValuePair<double, byte[]> pair)
        {
            double con = pair.Key;
            foreach (DataRow row in dataTable.Rows)
            {
                if (double.TryParse(row["Measure TONG (mm2)"].ToString(), out double res))
                {
                    row["% Area TONG"] = Math.Round((res / con) * 100, 2).ToString();
                }
                if (double.TryParse(row["Measure TRU (mm2)"].ToString(), out res))
                {
                    row["% Area TRU"] = Math.Round((res / con) * 100, 2).ToString();
                }
                CheckDataRow(row);
            }
        }

        public void CheckDataRow(DataRow row)
        {

            if (double.TryParse(row["% Area TONG"].ToString(), out double num) && num < 30 && double.TryParse(row["Measure(mm)"].ToString(), out double num2) && num2 < 0.1)
            {
                row["Judgement"] = "OK";
            }
            else
            {
                row["Judgement"] = "NG";
            }
        }
        public void LoadDataRefer()
        {
            LoadData();

            DataTable referList = _dBContext.LoadDataTable("AIR_BUBBLE_REFER", new[] { "ItemMain", "LotMain" }, new[] { _itemCode, _lotNo });
            if (referList.Rows.Count > 0)
            {
                foreach (DataRow row in referList.Rows)
                {
                    string itemCode = row["ItemCodeRefer"].ToString();
                    string lotNo = row["LotNoRefer"].ToString();
                    LoadData(itemCode, lotNo, true);
                    ListComment.Add(itemCode);

                }
            }
            foreach (string key in _dic.Keys)
            {
                if (_dicTONG.TryGetValue(key, out KeyValuePair<double, byte[]> pair))
                {
                    CalculateArea(_dic[key], pair);
                }
                else
                {
                    string type = key.Split('_')[0].Trim();
                    string tape = key.Split('(', ')')[1].Trim();
                    foreach (string keyNew in _dicTONG.Keys.ToArray())
                    {
                        if (keyNew.Contains(type) && keyNew.Contains(tape))
                        {
                            if (_dicTONG.TryGetValue(key, out KeyValuePair<double, byte[]> pairZ))
                            {
                                _dicTONG[key] = _dicTONG[keyNew];
                            }
                            else
                            {
                                _dicTONG.Add(key, _dicTONG[keyNew]);
                            }
                            CalculateArea(_dic[key], _dicTONG[keyNew]);
                        }
                    }
                }
            }

        }

        public void Remove(string itemCode, string tape)
        {
            _dBContext.DeleteData("AIR_BUBBLE_COMMENT", new[] { "ItemCode", "TapeName" }, new[] { itemCode, tape });
        }

        public void RemoveRefer(string itemCode, string lotNo)
        {
            int res = _dBContext.DeleteData("AIR_BUBBLE_REFER", new[] { "ItemMain", "LotMain", "ItemCodeRefer", "LotNoRefer" }, new[] { _itemCode, _lotNo, itemCode, lotNo });
            if (res > 0)
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("Id", typeof(int));
                dt.Columns.Add("ItemCode");
                dt.Columns.Add("LotNo");
                dt.Columns.Add("Status");
                DataRow row = dt.NewRow();
                row["ItemCode"] = itemCode;
                row["LotNo"] = lotNo;
                row["Status"] = "1";
                dt.Rows.Add(row);
                res = _dBContext.BuckDataTable(dt, "AIR_BUBBLE_AVAILABLE", new[] { "ItemCode", "LotNo" }, null, "Id");
            }
        }
    }
}
