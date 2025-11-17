using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Authentication;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace OK2SHIP_SMT.Services
{
    public sealed class UserSession
    {
        private const string _NAMETABLE = "ACCOUNT";
        private DBContext _dbContext = new DBContext();
        private static UserSession instance = null;
        private static readonly object padlock = new object();
        public string Role { get; private set; } = "";
        public string Username { get; private set; }
        public string User_ID { get; private set; }
        public bool IsLoggedIn { get; private set; }
        private const string ADMIN_USERNAME = "TDMK_ADMIN";
        private const string ADMIN_PASSWORD = "qwer";

        private UserSession()
        {
            // Constructor private để ngăn việc tạo instance từ bên ngoài
            IsLoggedIn = false;
//#if DEBUG
//            Role = "admin";
//            Username = "TDMK_ADMIN";
//            IsLoggedIn = true;
//            User_ID = "ADMIN";
//#endif

        }

        public static UserSession Instance
        {
            get
            {
                lock (padlock)
                {
                    if (instance == null)
                    {
                        instance = new UserSession();
                    }
                    return instance;
                }
            }
        }

        public bool Login(string username, string password)
        {
            if (username == ADMIN_USERNAME && password == ADMIN_PASSWORD)
            {
                Role = "admin";
                Username = username;
                IsLoggedIn = true;
                User_ID = "ADMIN";
            }
            else if (username == ADMIN_USERNAME)
            {
                throw new Exception("Sai mật khẩu");
            }
            else
            {
                DataTable dt = _dbContext.LoadDataTable(_NAMETABLE, new[] { "UserName", "Password" }, new[] { username, password }, new[] { "User_ID", "UserName", "Role", "IsActive" });
                if (dt.Rows.Count > 0)
                {
                    DataRow row = dt.Rows[0];
                    Role = row["Role"].ToString();
                    Username = row["UserName"].ToString();
                    IsLoggedIn = row["IsActive"].ToString() == "1";
                    User_ID = row["User_ID"].ToString();
                }
                else
                {
                    throw new Exception("Sai mật khẩu, tài khoản");
                }
            }
            return IsLoggedIn;
        }
        public int EditingUser(string userName, string password, string user_ID, string role, string active = "ACTIVE")
        {
            //int numChange = 0;
            throw new Exception("Chưa triển khai");
        }
        public int CreateUser(string userName, string password, string user_ID, string role, int active = 1)
        {
            int numChange = 0;
            if (string.IsNullOrEmpty(userName) || string.IsNullOrEmpty(password) || string.IsNullOrEmpty(user_ID))
            {
                throw new Exception("Tên đăng nhập, mật khẩu và ID người dùng không được để trống.");
            }
            DataTable z = _dbContext.GetTableStructure(_NAMETABLE);
            DataRow row = z.NewRow();
            row["User_ID"] = user_ID;
            row["Username"] = userName;
            row["Password"] = password;
            switch (role.ToString())
            {
                case "Admin":
                    role = "0";
                    break;
                case "QA":
                    role = "1";
                    break;
                default:
                    throw new Exception("Nhập lại role");
            }
            row["Role"] = int.Parse(role);
            row["IsActive"] = active;
            z.Rows.Add(row);
            numChange = _dbContext.SaveDataTable(z, _NAMETABLE);
            return numChange;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        /// <exception cref="AuthenticationException"></exception>
        /// <exception cref="Exception"></exception>
        public DataTable ListOfUser()
        {
            if (!this.Role.Equals("admin"))
            {
                throw new AuthenticationException("Chỉ có admin mới có quyền xem danh sách người dùng.");
            }
            DataTable dt = _dbContext.LoadDataTable(_NAMETABLE, null, null);

            if (dt == null)
            {
                throw new Exception("Không thể tải danh sách người dùng.");
            }
            if (dt.Rows.Count == 0)
            {
                throw new Exception("Không có người dùng nào trong hệ thống.");
            }
            return dt;
        }
        public void Logout()
        {
            Role = null;
            Username = null;
            IsLoggedIn = false;
        }
    }
}
