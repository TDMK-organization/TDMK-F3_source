using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;

namespace OK2SHIP_SMT.Services
{
    static class ValidateService
    {
        public static bool isDigitAndChar(string str)
        {
            if (string.IsNullOrEmpty(str))
            {
                return false; // Chuỗi rỗng hoặc null không có ký tự đặc biệt
            }

            // Sử dụng Regular Expression để kiểm tra
            Regex regex = new Regex("[^a-zA-Z0-9]"); // Tìm ký tự không phải chữ cái hoặc số
            return !regex.IsMatch(str);
        }
        public static int isDigit(string number)
        {
            if (int.TryParse(number, out int total))
            {
                return total;
            }
            throw new Exception("Please enter a number!");
        }
        public static string PageSize(string pageSize)
        {
            return isDigit(pageSize).ToString();
        }
        public static string PageNumber(string pageNumber, int max, int min = 0)
        {
            int number = isDigit(pageNumber);
            if (number < min)
            {
                number = min;
            }
            if (number > max)
            {
                number = max;
            }
            return number.ToString();
        }

        /// <summary>
        /// Handle lotno
        /// </summary>
        /// <param name="lotno"></param>
        /// <returns></returns>
        public static string lotNoHandle(string lotno)
        {
            string[] str = lotno.Split('-');
            if (str.Length < 0 && str.Length > 2)
            {
                return lotno;
            }
            string lot = str[0];
            while (lot.Length < 5)
            {
                lot = "0" + lot;
            }

            if (str.Length > 1)
            {
                string no = str[1];
                while (no.Length < 2)
                {
                    no = "0" + no;
                }
                lot = lot + "-" + no;
            }
            return lot;
        }
    }
}
