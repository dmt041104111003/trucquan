using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSharpCounterFinalProject.ViewCustomer
{
    public partial class CustomerInformationView : Form
    {
        Classes.DataBaseProcess data = new Classes.DataBaseProcess();
        string cus_id;
        string cus_name;
        public CustomerInformationView(string username)
        {
            InitializeComponent();
            CenterToScreen();
            txtInfoUserNameCus.Text = username;
        }

        private void CustomerInformationView_Load(object sender, EventArgs e)
        {
            DisplayInfoUser();
        }

        private void DisplayInfoUser()
        {
            // Thực hiện truy vấn SQL và lấy dữ liệu
            string query = "select c.Customer_ID, c.BirthDate, c.Address, c.PhoneNumber, c.CustomerType, c.Point, c.CreateTime, c.Email, c.FullName, c.Username, c.Password " +
                            "from[dbo].[Customer] as c " +
                            "where c.Username = N'" + txtInfoUserNameCus.Text + "' ";
                            //"and c.Password = N'" + txtInfoPassCus.Text + "'";

            // dtBase.DataReader là một đối tượng chưa kết nối và thực hiện truy vấn

            string strConnect = Classes.DataBaseProcess.strConnect;
            using (SqlConnection connection = new SqlConnection(strConnect)) // Kết nối cơ sở dữ liệu
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(query, connection)) // Thực hiện truy vấn
                {
                    using (SqlDataReader reader = command.ExecuteReader()) // Đọc dữ liệu
                    {
                        if (reader.Read())
                        {
                            // Gán dữ liệu từ truy vấn lên các thẻ Label
                            txtInfoIDCus.Text = reader["Customer_ID"].ToString();
                            dateTimePicker1.Text = reader["BirthDate"].ToString();
                            txtInfoAddressCus.Text = reader["Address"].ToString();
                            txtInfoPhoneCus.Text = reader["PhoneNumber"].ToString();
                            txtInfoTypeCus.Text = reader["CustomerType"].ToString();
                            txtInfoPointCus.Text = reader["Point"].ToString();
                            dateTimePicker2.Text = reader["CreateTime"].ToString();
                            txtInfoEmailCus.Text = reader["Email"].ToString();
                            txtInfoNameCus.Text = reader["FullName"].ToString();
                            txtInfoUserNameCus.Text = reader["Username"].ToString();
                            txtInfoPassCus.Text = reader["Password"].ToString();
                        }
                        else
                        {
                            // Xử lý trường hợp không có dữ liệu
                            // Ví dụ: labelCart_ID.Text = "Không có dữ liệu";
                        }
                    }
                }
            }
            cus_id = this.txtInfoIDCus.Text;
            cus_name = this.txtInfoUserNameCus.Text;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
            var children = new HomeCustomerView(cus_name);
            children.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string sql = "";
            var id = txtInfoIDCus.Text;
            var name = txtInfoNameCus.Text;
            var address = txtInfoAddressCus.Text;
            var email = txtInfoEmailCus.Text;
            var phone = txtInfoPhoneCus.Text;
            var birthDate = dateTimePicker1.Value.ToString();
            var createdAccTime = dateTimePicker2.Value.ToString();
            var point = int.Parse(txtInfoPointCus.Text);
            var customerType = txtInfoTypeCus.Text;

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
            sql += "WHERE Customer_ID = " + cus_id;
            data.DataChange(sql);
            MessageBox.Show("Cập nhật thành công!", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            
        }
    }
}
