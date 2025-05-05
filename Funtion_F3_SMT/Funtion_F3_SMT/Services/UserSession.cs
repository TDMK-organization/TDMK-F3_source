using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        private UserSession()
        {
            // Constructor private để ngăn việc tạo instance từ bên ngoài
            IsLoggedIn = false;
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
            if (username == "Admin" && password == "TDMK")
            {
                Role = "admin";
                Username = username;
                IsLoggedIn = true;

            }
            else
            {
                DataTable dt = _dbContext.LoadDataTable(_NAMETABLE, new[] { "UserName", "Password" }, new[] { username, password }, new[] { "UserName", "Role", "IsActive" });
                if (dt.Rows.Count > 0)
                {
                    DataRow row = dt.Rows[0];
                    Role = row["Role"].ToString();
                    Username = row["UserName"].ToString();
                    IsLoggedIn = row["IsActive"].ToString().Trim().Equals("ACTIVE");
                }
            }
            return IsLoggedIn;
        }
        public int EditingUser(string userName, string password, string user_ID, string role, string active = "ACTIVE")
        {
            //int numChange = 0;
            throw new Exception("Chưa triển khai");
        }
        public int CreateUser(string userName, string password, string user_ID, string role = "STAFF", string active = "ACTIVE")
        {
            int numChange = 0;
            if (string.IsNullOrEmpty(userName) || string.IsNullOrEmpty(password) || string.IsNullOrEmpty(user_ID))
            {
                throw new Exception("Tên đăng nhập, mật khẩu và ID người dùng không được để trống.");
            }
            return numChange;
        }
        public DataTable ListOfUser()
        {
            if (!this.Role.Equals("admin"))
            {
                throw new Exception("Chỉ có admin mới có quyền xem danh sách người dùng.");
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
