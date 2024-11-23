using System;
using System.Drawing;
using System.Windows.Forms;

namespace CSharpCounterFinalProject
{
    public partial class AddEditCustomerFrm : Form
    {
        // Declare and constructor global variable in class to use DataBaseProcess
        Classes.DataBaseProcess dtBase = new Classes.DataBaseProcess();
        public AddEditCustomerFrm()
        {
            InitializeComponent();
            CenterToScreen();
        }

        public AddEditCustomerFrm(string str, string customer_ID, string fullName, DateTime birthDate, string address,
                                  string phone, string customerType, int point, DateTime createTime, string email)
        {
            InitializeComponent();
            CenterToScreen();
            if(btnAddNew != null)
            {
                txtCustomerId.Text = "";
                btnAddNew.Text = str;
                txtCustomerId.Text = customer_ID;
                txtFullName.Text = fullName;
                txtAddress.Text = address;
                txtPhoneNumber.Text = phone;
                txtEmail.Text = email;
                datePickerBirthDate.Value = birthDate;
                datePickerCreatedAcc.Value = createTime;
                numericPoint.Value = point;
                comboCustomerType.Text = customerType;
            }
        }

        private void TxtIdTextFocus(object sender, EventArgs e)
        {
            //txtCustomerId.Text = "";
            txtCustomerId.ForeColor = Color.Black;
        }

        private void TxtIdLeaveFocus(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtCustomerId.Text))
            {
                txtCustomerId.Text = "CMND/CCCD";
                txtCustomerId.ForeColor = Color.Gray;
            }
            else
            {
                txtCustomerId.ForeColor = Color.Black;
            }
        }

        private void BtnCancelClick(object sender, EventArgs e)
        {
            var title = "Xác nhận";
            var message = "Bạn có chắc muốn hủy bỏ";
            var ans = ShowConfirmMessage(title, message);
            if (ans == DialogResult.Yes)
            {
                Dispose();
            }
        }
        private DialogResult ShowConfirmMessage(string title, string message)
        {
            return MessageBox.Show(message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        }

        private void BtnAddUpdateClick(object sender, EventArgs e)
        {
            if (btnAddNew.Text.CompareTo("Thêm mới") == 0)
            {
                string sql = "";
                var id = txtCustomerId.Text;
                var name = txtFullName.Text;
                var address = txtAddress.Text;
                var email = txtEmail.Text;
                var phone = txtPhoneNumber.Text;
                var birthDate = datePickerBirthDate.Value;
                var createdAccTime = datePickerCreatedAcc.Value;
                var point = (int)numericPoint.Value;
                var customerType = comboCustomerType.Text;

                //Insert vào CSDL
                sql = "INSERT INTO Customer (Customer_ID, BirthDate, Address, PhoneNumber, CustomerType, Point, CreateTime, Email, FullName) VALUES (";
                sql += "N'" + id + "',N'" + birthDate + "',N'" +
                       address + "',N'" + phone + "',N'" + customerType + "',N'" + point + "',N'" + createdAccTime + "',N'" + email + "',N'" + name + "')";
                dtBase.DataChange(sql);
                MessageBox.Show("Thêm thành công!", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                Close();
            }
            else if (btnAddNew.Text == "Cập Nhật")
            {
                string sql = "";
                var id = txtCustomerId.Text;
                var name = txtFullName.Text;
                var address = txtAddress.Text;
                var email = txtEmail.Text;
                var phone = txtPhoneNumber.Text;
                var birthDate = datePickerBirthDate.Value;
                var createdAccTime = datePickerCreatedAcc.Value;
                var point = (int)numericPoint.Value;
                var customerType = comboCustomerType.Text;

                sql = "UPDATE Customer SET ";
                sql += "Customer_ID = N'" + id + "',";
                sql += "FullName = N'" + name + "',";
                sql += "Address = N'" + address + "',";
                sql += "Email = N'" + email + "',";
                sql += "PhoneNumber = N'" + phone + "',";
                sql += "BirthDate = N'" + birthDate + "', ";
                sql += "CreateTime = N'" + createdAccTime + "', ";
                sql += "CustomerType = N'" + customerType + "', ";
                sql += "Point = N'" + point + "' ";
                sql += "WHERE Customer_ID = " + id;
                dtBase.DataChange(sql);
                MessageBox.Show("Cập nhật thành công!", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                Close();
            }

        }

        private void ShowErrorMessage(string msg)
        {
            var title = "Lỗi dữ liệu";
            MessageBox.Show(msg, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void AddEditCustomerFrm_Load(object sender, EventArgs e)
        {

        }
    }
}
