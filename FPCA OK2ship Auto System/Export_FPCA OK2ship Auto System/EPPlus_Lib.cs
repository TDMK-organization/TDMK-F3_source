using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;

namespace OK2SHIP_SMT
{
    public class EPPlus_Lib
    {
        public ExcelWorkbook open_excel_file(string file_name)
        {
            ExcelWorkbook result = null;
            FileInfo excel_file = new FileInfo(file_name);
            if (File.Exists(file_name))
            {
                ExcelPackage myexcel = new ExcelPackage(excel_file);
                result = myexcel.Workbook;
            }
            return result;
        }

        public static bool IsValidName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }
            var fc = name[0];
            if (!(char.IsLetter(fc) || fc == '_' || (fc == '\\' && name.Length > 2)))
            {
                return false;
            }

            //if (name.IndexOfAny(NameInvalidChars, 1) > 0)
            //{
            //    return false;
            //}

            if (ExcelCellBase.IsValidAddress(name))
            {
                return false;
            }

             
            return true;
        }
        //public Image get_pic(ExcelWorksheet wrksht, string pic_name)
        //{
        //    Image result = null;
        //    List<string> img_lst = new List<string>();
        //    foreach (ExcelPicture img in wrksht.Drawings)
        //    {
        //        img_lst.Add(img.Name);
        //    } 
        //    int inx = img_lst.IndexOf(pic_name);
        //    if (inx != -1)
        //    {
        //        ExcelPicture cur_img = wrksht.Drawings[img_lst[inx]] as ExcelPicture;
        //        result = cur_img.Image;
        //    }
        //    return result;
        //}
    }
}
