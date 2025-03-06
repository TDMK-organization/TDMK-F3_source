using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SQL;
using TDMK_SEEV_DLL;
using System.Drawing;
using myExcel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace ACF_Process
{
    public class ACF_Process_Lib
    {
        SEI_Lib myCode = new SEI_Lib();
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public struct ACF_Tape_Test
        {
            public byte[] picture { get; set; }
            public string force { get; set; }
            public ACF_Tape_Test(byte[] in_pic, string in_force)
            {
                picture = in_pic;
                force = in_force;
            }
        }
        public Dictionary<int, byte[]> Get_Image_LogFile_Data(string in_src_file)
        {
            Dictionary<int, byte[]> result_lst = new Dictionary<int, byte[]>();
            string[] files = Directory.GetFiles(in_src_file, "*.png");
            foreach (var f in files)
            {
                string f_name = Path.GetFileNameWithoutExtension(f);
                Image myImg = Bitmap.FromFile(f);
                ImageConverter imgCon = new ImageConverter();
                byte[] img_data = (byte[])imgCon.ConvertTo(myImg, typeof(byte[]));
                result_lst.Add(Convert.ToInt32( f_name), img_data);
        }
            return result_lst;
        }
        public string Get_Number_String(string src_str, char split_chr)
        {
            string _result = "";
            string[] temp = src_str.Split(split_chr);
            foreach (var t in temp)
            {
                if (myCode.IsNumeric(t))
                {
                    _result = t;
                    break;
                }
            }
            return _result;
        }
        
        public void Get_Image_logfile(string in_src, ref Dictionary<string, Dictionary<int, byte[]>> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            FileInfo[] temp_lst = tar_d.GetFiles("*.png");
            if (temp_lst.Length > 0)
            {
                Dictionary<int, byte[]> temp = Get_Image_LogFile_Data(tar_d.FullName);
                string f_na = tar_d.Name;
                lst_result.Add(f_na, temp);
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_Image_logfile(inter_lst.FullName, ref lst_result);
                    }
                }
            }
        } 

        public Dictionary<int, ACF_Tape_Test> Get_TapeTest_LogFile_Data(string in_src_file)
        {
            Dictionary<int, ACF_Tape_Test> result_lst = new Dictionary<int, ACF_Tape_Test>();
            string[] files = Directory.GetFiles(in_src_file, "*.xlsx");
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            foreach (var f in files)
            {
                string f_name = Path.GetFileNameWithoutExtension(f);
                myExcel.Workbook wrkbk = TDMK_Code.open_excel_file(f, "", "");
                myExcel.Worksheet wrksht = wrkbk.Sheets[1];
                myExcel.Range data_rgn = wrksht.Range["E5"];
                myExcel.Shape cur_image;
                cur_image = wrksht.Shapes.Item("Picture 1");
                cur_image.Copy();
                Image myImg = Clipboard.GetImage();
                ImageConverter imgCon = new ImageConverter();
                byte[] img_data = (byte[])imgCon.ConvertTo(myImg, typeof(byte[]));
               // int r_inx = 0;
                //while(myCode.checkDBNull(data_rgn.Offset[r_inx,0].Value)!="")
                for(int c_inx = 0; c_inx < 10; c_inx++)
                {
                    for (int r_inx = 0; r_inx < 20; r_inx++)
                    {
                        if (myCode.checkDBNull(data_rgn.Offset[r_inx, c_inx].Value) == "Max")
                        {

                            for (int i = 1; i < 10; i++)
                            {
                                if (myCode.checkDBNull(data_rgn.Offset[r_inx, c_inx + i].Value) != "")
                                {
                                    string force_val = myCode.checkDBNull(data_rgn.Offset[r_inx, c_inx + i].Value);
                                    result_lst.Add(Convert.ToInt32(f_name), new ACF_Tape_Test(img_data, force_val));
                                    goto lbl_close;
                                }
                            }
                        }

                    }
                }
                 
                lbl_close: 
                wrkbk.Close();
            }
            xlsApp.Quit();
            Marshal.ReleaseComObject(xlsApp);
            return result_lst;
        }
        public void Get_TapeTest_logfile(string in_src, ref Dictionary<string, Dictionary<int, ACF_Tape_Test>> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            FileInfo[] temp_lst = tar_d.GetFiles("*.xlsx");
            if (temp_lst.Length > 0)
            {
                Dictionary<int, ACF_Tape_Test> temp = Get_TapeTest_LogFile_Data(tar_d.FullName);
                string f_na = tar_d.Name;
                lst_result.Add(f_na, temp);
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_TapeTest_logfile(inter_lst.FullName, ref lst_result);
                    }
                }
            }

        }
    }
}
