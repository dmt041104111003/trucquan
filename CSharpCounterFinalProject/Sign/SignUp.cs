using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSharpCounterFinalProject.Sign
{
    public partial class SignUp : Form
    {
        // Declare and constructor global variable in class to use DataBaseProcess
        Classes.DataBaseProcess dtBase = new Classes.DataBaseProcess();
        public SignUp()
        {
            InitializeComponent();
            CenterToScreen();
        }

        private void BtnSignUpClick(object sender, EventArgs e)
        {
            string sql = "";
            var id = txtCMND.Text;
            var name = txtFullName.Text;
            var address = txtAddress.Text;
            var email = txtEmail.Text;
            var phone = txtPhoneNumber.Text;
            var birthDate = dtpDate.Value;
            var username = txtUsername.Text;
            var password = txtPassword.Text;
            //Insert vào CSDL
            if(username == "admin" && password == "admin")
            {
                MessageBox.Show("Tài khoản admin không thể tạo!");
                return;
            }
            DataTable data = new DataTable();
            data = dtBase.DataReader("select * from Customer where Customer_ID = '"+ txtCMND.Text + "' or Username = '"+txtUsername.Text+"'");
            if(data.Rows.Count > 0 ) 
            {
                MessageBox.Show("Đã tồn tại TK này!");
                return; 
            }

            sql = "INSERT INTO Customer (Customer_ID, BirthDate, Address, PhoneNumber, CustomerType, Point, CreateTime, Email, FullName, Username, Password) VALUES (";
            sql += "N'" + id + "',N'" + birthDate + "',N'" +
                   address + "',N'" + phone + "',N'" + "Khách mua lẻ" + "',N'" + 0 + "',N'" + DateTime.Now.ToString("MM/dd/yyyy") + "',N'" + email + "',N'" + name + "',N'" + username + "',N'" + password +"')";
            dtBase.DataChange(sql);
            MessageBox.Show("Thêm thành công!", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            Close();
            var childview = new SignIn();
            childview.Show();
        }
        private void txtUsername_Validating(object sender, CancelEventArgs e)
        {
            if (string.IsNullOrEmpty(txtUsername.Text))
            {
                this.errorProvider.SetError(this.txtUsername, "Hãy nhập tên đăng nhập.");
                return;
            }
            DataTable dt = new DataTable();
            dt = this.dtBase.DataReader("select Customer_ID from Customer where Username = '" + txtUsername.Text + "'");
            if (dt.Rows.Count > 0)
            {
                this.errorProvider.SetError(txtUsername, "Tên đăng nhập đã tồn tại");
                return;
            }

        }

        private void txtPassword_Validating(object sender, CancelEventArgs e)
        {
            if (string.IsNullOrEmpty(txtPassword.Text))
            {
                this.errorProvider.SetError(txtPassword, "Hãy nhập mật khẩu.");
                return;
            }
        }

        private void txtPasswordComform_Validating(object sender, CancelEventArgs e)
        {
            if (string.IsNullOrEmpty(txtPasswordComform.Text))
            {
                this.errorProvider.SetError(txtPasswordComform, "Hãy nhập lại mật khẩu.");
                return;
            }
            if (txtPasswordComform.Text != txtPassword.Text)
            {
                this.errorProvider.SetError(txtPasswordComform, "Hãy nhập lại mật khẩu đúng với mật khẩu đăng kí.");
                return;
            }
        }

        private void txtCMND_Validating(object sender, CancelEventArgs e)
        {
            if (string.IsNullOrEmpty(txtCMND.Text))
            {
                this.errorProvider.SetError(txtCMND, "Hãy nhập CMND.");
                return;
            }
            if (txtCMND.Text.Length != 8 && int.TryParse(txtCMND.Text, out int r))
            {
                this.errorProvider.SetError(txtCMND, "Hãy nhập 8 số CMND.");
                return;
            }
        }

        private void txtEmail_Validating(object sender, CancelEventArgs e)
        {
            if (string.IsNullOrEmpty(txtEmail.Text))
            {
                e.Cancel = true;
                errorProvider.SetError(txtEmail, "Hãy nhập Email!");
            }
            else if (!IsValidEmail(txtEmail.Text))
            {
                e.Cancel = true;
                errorProvider.SetError(txtEmail, "Hãy nhập đúng định dạng Email!");
            }
        }
        private static bool IsValidEmail(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
                return false;

            try
            {
                // Normalize the domain
                email = Regex.Replace(email, @"(@)(.+)$", DomainMapper,
                                      RegexOptions.None, TimeSpan.FromMilliseconds(200));

                // Examines the domain part of the email and normalizes it.
                string DomainMapper(Match match)
                {
                    // Use IdnMapping class to convert Unicode domain names.
                    var idn = new IdnMapping();

                    // Pull out and process domain name (throws ArgumentException on invalid)
                    string domainName = idn.GetAscii(match.Groups[2].Value);

                    return match.Groups[1].Value + domainName;
                }
            }
            catch (RegexMatchTimeoutException e)
            {
                return false;
            }
            catch (ArgumentException e)
            {
                return false;
            }

            try
            {
                return Regex.IsMatch(email,
                    @"^[^@\s]+@[^@\s]+\.[^@\s]+$",
                    RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250));
            }
            catch (RegexMatchTimeoutException)
            {
                return false;
            }
        }

        private void txtPhoneNumber_Validating(object sender, CancelEventArgs e)
        {
            if (string.IsNullOrEmpty(txtPhoneNumber.Text))
            {
                e.Cancel = true;
                errorProvider.SetError(txtPhoneNumber, "Hãy nhập SĐT!");
            }
            if (txtPhoneNumber.Text.Length != 10 && long.TryParse(txtPhoneNumber.Text, out long r))
            {
                e.Cancel = true;
                errorProvider.SetError(txtPhoneNumber, "SĐT là 10 số.");
            }
        }

        private void SignUp_Load(object sender, EventArgs e)
        {

        }

        private void btnHuy_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
