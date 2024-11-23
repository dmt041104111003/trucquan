using CSharpCounterFinalProject.Classes;
using CSharpCounterFinalProject.ViewNguoiMua;
using System;
using System.Data;
using System.Windows.Forms;

namespace CSharpCounterFinalProject.Sign
{
    public partial class SignIn : Form
    {
        // Declare and constructor global variable in class to use DataBaseProcess
        Classes.DataBaseProcess dtBase = new Classes.DataBaseProcess();
        public SignIn()
        {
            InitializeComponent();
            CenterToScreen();
        }

        private void BtnHuyClick(object sender, EventArgs e)
        {
            Close();
        }

        private void BtnDangNhapCkick(object sender, EventArgs e)
        {
            string username = txtName.Text;
            string password = txtPassword.Text;
            if(username==""|| password == "")
            {
                MessageBox.Show("Không được để trống","Thông báo",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                return;
            }
            
            bool isNguoiDung = checkNguoiDung(username, password);
            bool isNhanVien = checkNhanVien(username, password);
            if (isNhanVien)
            {
                MessageBox.Show("Đăng nhập thành công!");
                var childView = new HomeFrm();
                this.Hide();
                
                childView.Show();
            }else if (isNguoiDung)
            {
                MessageBox.Show("Đăng nhập thành công!");
                HomeView home = new HomeView(Program.currentUser.TenUser);
                this.Hide();
                home.ShowDialog();
            }
            else
            {
                MessageBox.Show("Thông tin đăng nhập sai!");
                return;
            }
        }

        private bool checkNhanVien(string username, string password)
        {
            bool check = false;
            string query = "select * from NhanVien where MaNV = '" + username + "' and MatKhau ='" + password + "'";
            DataTable dt = new DataTable();
            dt = dtBase.DataReader(query);
            if(dt.Rows.Count > 0)
            {
                check = true;
                
            }
            return check;
        }

        private bool checkNguoiDung(string username, string password)
        {
            bool check = false;
            string query = "select * from KhachHang where MaKH = '" + username + "' and MatKhau ='" + password + "'";
            DataTable dt = new DataTable();
            dt = dtBase.DataReader(query);
            if (dt.Rows.Count > 0)
            {
                check = true;
                Program.currentUser.TenUser = dt.Rows[0]["TenKhachHang"].ToString();
                Program.currentUser.MaUser = dt.Rows[0]["MaKH"].ToString();
            }
            return check;
        }

        private void BtnSignUpClick(object sender, EventArgs e)
        {
            var childView = new SignUp();
            childView.Show();
        }

        private void SignIn_Load(object sender, EventArgs e)
        {

        }
    }
}
