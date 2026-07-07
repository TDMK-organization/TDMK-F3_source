using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OK2SHIP_SMT.Services
{
    public class AccountService
    {
        DBContext _dbContext = new DBContext();

        public string UpdateAccount(string username, string password)
        {
            try
            {
                DataTable dt = _dbContext.GetTableStructure("ACCOUNT");
                DataRow row = dt.NewRow();
                row["User_ID"] = UserSession.Instance.User_ID;
                row["Username"] = username;
                row["Password"] = password;
                dt.Rows.Add(row);
                int res = _dbContext.Update("ACCOUNT", dt, new[] { "User_ID" });
                return  res > 0 ? "Cập nhật tài khoản thành công" : "Cập nhật thất bại";
            }
            catch (Exception ex)
            {
                return $"Error {ex.Message}";
            }
        }
        public string UpdateAccount(DataTable dt)
        {
            try
            {

                int res = _dbContext.Update("ACCOUNT", dt, new[] { "User_ID" });
                if (res > 0)
                {
                    return $"Cập nhật {res} tài khoản thành công";
                }
                return "Cập nhật thất bại";
            }
            catch (Exception ex)
            {
                return $"Error {ex.Message}";
            }
        }

        public string DeleteAccount(string[] userId)
        {
            try
            {
                if (userId.Length != 0)
                {
                    int res = _dbContext.DeleteData("ACCOUNT", new[] { "User_ID" }, userId);
                    if (res != 0)
                    {
                        return "Account deleted successfully.";
                    }
                }
                return "Account not found.";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

    }
}
