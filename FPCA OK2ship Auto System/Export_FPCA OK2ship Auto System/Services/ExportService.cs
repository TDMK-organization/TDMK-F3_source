using Export_FPCA_OK2ship_Auto_System.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Export_FPCA_OK2ship_Auto_System.Services
{

    public class ExportService
    {
        private DBContext _db = null;
        private string[] categories = 
        {
            "Coverpage",
            "Rev History",
            "User Guideline",
            "Low CPK Action",
            "Declaration",
            "Table of Contents",
            "Deviation summary",
            "Assy Yield",
            "FAI",
            "OQC Test",
            "Cross section",
            "GAP Connector",
            "Peel Test",
            "(Mating) Pull Test",
            "IQC Liner peeling (Coupon)",
            "IQC PSA peeling (Coupon)",
            "Shear test",
            "Liner peel test (On product)",
            "Air bubble btw Liner-PSA",
            "PSA peel test (On product)",
            "Air bubble btw PSA-FPC",
            "(IQC Unmating) Pull Test",
            "OQC B2B Mating-Unmating",
            "ORT-Assy",
            "Flex bending",
            "Thermal Cycling & bending",
            "Heat Soak & bending",
            "X-Ray picture",
            "Heat Soak and Recovery",
            "Thermal Cycling",
            "Thermal Shock",
            "Environment en-durance",
            "Impedance",
            "Switch Quality",
            "ACF",
            "SEM BSE & Binarization",
            "Bar Code Verification",
            "Packaging",
            "Mishandling test",
            "Process flow",
            "Process Comparison",
        };
        public Dictionary<string, int> _dicCategory = new Dictionary<string, int>();
        public DataTable setupDGV()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Category", typeof(string));
            dt.Columns.Add("StatusDB", typeof(string));
            dt.Columns.Add("Export status", typeof(bool));
            dt.Columns.Add("Export form DB", typeof(bool));
            foreach (string item in categories)
            {
                _dicCategory.Add(item, dt.Rows.Count + 1);
                dt.Rows.Add(dt.Rows.Count + 1, item, "None", false, false);
            }
            return dt;

        }

        
    }
}
