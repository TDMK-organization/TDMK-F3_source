using Export_FPCA_OK2ship_Auto_System.Repositories;
using Export_FPCA_OK2ship_Auto_System.Services.TDMK_services;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Export_FPCA_OK2ship_Auto_System.Services
{
    public class PackagingService
    {
        private DBContext _dbContext = new DBContext();
        private string _NAMETABLE = "PACKAGING";

        public DataTable Load(string itemCode)
        {
            itemCode = itemCode.Trim();
            if (string.IsNullOrEmpty(itemCode))
            {
                throw new ArgumentException("Item code cannot be null or empty.");
            }

            DataTable dataTable = _dbContext.LoadDataTable("PACKAGING_LOGFILE", new[] { "ItemCode" }, new[] { itemCode });
            if (dataTable.Rows.Count <= 0)
            {
                return dataTable;
            }
            NasRepository nas = new NasRepository();
            nas.MergeDataTable(dataTable, $"{_NAMETABLE}", itemCode, "", dataTable.Rows[0]["LocationIMG"].ToString());
            return dataTable;
        }



        public string Export(string itemCode, string lotNo)
        {
            try
            {
                ExportToExcel(itemCode, lotNo, true);
                return "OK";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        public void ExportToExcel(string itemCode, string lotNo, bool prime)
        {
            DataTable dataTable = Load(itemCode);
            if (dataTable.Rows.Count == 0)
            {
                throw new Exception($"Item code {itemCode} not found in the database.");
            }
            ExportProcess exportProcess = new ExportProcess();
            using (ExcelPackage package = exportProcess.FindFormatProcess("PAKAGING", itemCode, ""))
            {
                using (ExcelWorksheet worksheet = exportProcess.FindSheet(package, "Packaging"))
                {
                    string[] name = new[] { "Packing Ship", "Picture", "6. Any liner drop" };
                    IDictionary<string, string> dic = ExportProcess.FindAddressByText(worksheet, name);

                    dic["Picture"] = dic["Picture"].Split('-').FirstOrDefault();
                    string startAddress = $"{dic["Packing Ship"]}:{worksheet.Cells[worksheet.Cells[dic["6. Any liner drop"]].End.Row, worksheet.Cells[dic["Picture"]].End.Column + 1]}";
                    string PasteAddress = dic["Packing Ship"];
                    for (int i = 0; i < dataTable.Rows.Count - 1; i++)
                    {
                        PasteAddress = ExportProcess.AddColumn(PasteAddress, 3);
                        ExportProcess.CopyColumn(worksheet, worksheet.Cells[startAddress], PasteAddress);
                    }
                    name = new[] { "Packing Ship", "Picture", "1. Any Tray deformation" };
                    dic = ExportProcess.FindAddressByText(worksheet, name, true);
                    dic.Add("Packing Ship", ExportProcess.FindAddressByText(worksheet, new[] { "Packing Ship" }).First().Value);
                    for (int i = 0; i < dic["Picture"].Split('-').Count(); i++)
                    {

                        DataRow row = dataTable.Rows[i];
                        string address = dic["Packing Ship"].Split('-')[i];
                        string value = worksheet.Cells[address].Text;
                        value = value.Replace("{tray code}", row["TrayCode"].ToString().Trim());
                        value = value.Replace("{packing ship}", row["ShippingTo"].ToString().Trim());
                        worksheet.Cells[address].Value = value;
                        address = dic["Picture"].Split('-')[i];
                        while (true)
                        {
                            address = ExportProcess.AddColumn(ExportProcess.AddRow(address, 1), -1);
                            value = worksheet.Cells[address].Text;
                            if (string.IsNullOrEmpty(value))
                            {
                                break;
                            }

                            address = ExportProcess.AddColumn(address, 1);
                            if (value.Contains("Flex") && value.Contains("Top") && value.Contains("side"))
                            {
                                try
                                {

                                    ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[address], (byte[])row["FlexTop"], $"FlexTopImage{i}");
                                }
                                catch
                                {

                                }
                            }
                            else if (value.Contains("Flex") && value.Contains("Bottom") && value.Contains("side"))
                            {
                                try
                                {

                                    ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[address], (byte[])row["FlexBottom"], $"FlexBottomImage{i}");
                                }
                                catch
                                {

                                }
                            }
                            else if (value.Contains("Tray") && !value.Contains("AL"))
                            {
                                try
                                {

                                    ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[address], (byte[])row["Tray"], $"Tray{i}");
                                }
                                catch
                                {

                                }
                            }
                            else if (value.Contains("Tray") && value.Contains("AL") && value.Contains("bag"))
                            {
                                try
                                {

                                    ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[address], (byte[])row["TrayAL"], $"TrayAL{i}");
                                    ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[address], (byte[])row["ALBag"], $"ALBAG{i}");
                                }
                                catch
                                {

                                }
                            }
                            else if (value.Contains("Carton Box"))
                            {
                                try
                                {

                                    ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[address], (byte[])row["CartonBox"], $"CartonBox{i}");
                                }
                                catch { }
                            }
                            else { break; }


                        }
                        DataTable JsonZ = TDMK_ConverterService.JsonToDataTable(row["Data"].ToString());
                        int iz = 0;
                        address = ExportProcess.AddColumn(dic["1. Any Tray deformation"].Split('-')[i], 1);
                        while (true)
                        {
                            var zs = worksheet.Cells[address].Value;
                            if (zs == null || string.IsNullOrEmpty(zs.ToString()))
                            {
                                break;
                            }
                            worksheet.Cells[address].Value = JsonZ.Rows[iz]["Result"];
                            address = ExportProcess.AddRow(address, 1);
                        }

                    }
                    exportProcess.SaveExcelWorksheet(package, "Packaging", $"{itemCode}_{lotNo}", "NPI", false);
                }
            }
        }
    }
}

