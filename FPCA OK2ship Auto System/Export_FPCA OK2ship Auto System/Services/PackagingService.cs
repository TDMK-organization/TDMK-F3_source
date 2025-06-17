using Export_FPCA_OK2ship_Auto_System.Repositories;
using Export_FPCA_OK2ship_Auto_System.Services.TDMK_services;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Export_FPCA_OK2ship_Auto_System.Services
{
    public class PackagingService
    {
        private DBContext _dbContext = new DBContext();
        public DataTable Load(string itemCode)
        {
            itemCode = itemCode.Trim();
            if (string.IsNullOrEmpty(itemCode))
            {
                throw new ArgumentException("Item code cannot be null or empty.");
            }

            DataTable dataTable = _dbContext.LoadDataTable("PackagingLogWithImages", new[] { "ItemCode" }, new[] { itemCode });
            return dataTable;
        }
        public void ExportToExcel(ExcelWorksheet worksheet, string itemCode)
        {
            DataTable dataTable = Load(itemCode);
            if (dataTable.Rows.Count == 0)
            {
                throw new Exception($"Item code {itemCode} not found in the database.");
            }
            ExportProcess exportProcess = new ExportProcess();

            string[] name = new[] { "Packing Ship", "Picture", "6. Any liner drop" };
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(worksheet, name);
            string startAddress = $"{dic["Packing Ship"]}:{worksheet.Cells[worksheet.Cells[dic["6. Any liner drop"]].End.Row, worksheet.Cells[dic["Picture"]].End.Column]}";
            string PasteAddress = dic["Packing Ship"];
            for (int i = 0; i < dataTable.Rows.Count - 1; i++)
            {
                PasteAddress = ExportProcess.AddColumn(PasteAddress, 3);
                exportProcess.CopyColumn(worksheet, worksheet.Cells[startAddress], PasteAddress);
            }
            name = new[] { "Packing Ship", "Picture", "1. Any Tray deformation" };
            dic = ExportProcess.FindAddressByText(worksheet, name);

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
                        ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[address], (byte[])row["FlexTopImage"], $"FlexTopImage{i}");
                    }
                    else if (value.Contains("Flex") && value.Contains("Bottom") && value.Contains("side"))
                    {
                        ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[address], (byte[])row["FlexBottomImage"], $"FlexBottomImage{i}");
                    }
                    else if (value.Contains("Tray") && !value.Contains("AL"))
                    {
                        ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[address], (byte[])row["TrayImage"], $"Tray{i}");
                    }
                    else if (value.Contains("Tray") && value.Contains("AL") && value.Contains("bag"))
                    {
                        ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[address], (byte[])row["TrayALImage"], $"TrayAL{i}");
                        ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[address], (byte[])row["TrayALImage"], $"ALBAG{i}");
                    }
                    else if (value.Contains("Carton Box"))
                    {
                        ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[address], (byte[])row["CartonBoxImage"], $"CartonBox{i}");
                    }
                    else { break; }


                }
                DataTable JsonZ = TDMK_ConverterService.JsonToDataTable(row["LogData"].ToString());
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
        }
    }
}

