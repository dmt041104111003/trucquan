using CSharpCounterFinalProject.Classes;
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

namespace CSharpCounterFinalProject
{
    public partial class PaymentFrm : Form
    {
        // Declare and constructor global variable in class to use DataBaseProcess
        Classes.DataBaseProcess dtBase = new Classes.DataBaseProcess();

        private string statusBill;

        public PaymentFrm()
        {
            InitializeComponent();
            CenterToScreen();
        }

        public PaymentFrm(string lableCus_id, string labelName)
        {
            InitializeComponent();
            CenterToScreen();
            txtCus_id.Text = lableCus_id;
            txtCustomerName.Text = labelName;
        }

        public PaymentFrm(int cart_ID, string nameCustomer, int totalItem, double totalDiscount, double totalAmount, string status)
        {
            InitializeComponent();
            CenterToScreen();
            txtCart_ID.Text = cart_ID.ToString();
            txtCustomerName.Text = nameCustomer;
            txtTotalItem.Text = totalItem.ToString();
            txtTotalDiscount.Text = totalDiscount.ToString();
            txtTotalAmount.Text = totalAmount.ToString();
            statusBill = status;

        }

        private void PaymentClick(object sender, EventArgs e)
        {
            if (btnFinish.Text.CompareTo("Hoàn tất") == 0)
            {
                string sql = "";
                var cart_ID = txtCart_ID.Text;
                var createTime = dtpCreatedTime.Value;
                var staffName = comboStaffName.Text;
                var paymentMethod = comboPaymentMethod.Text;
                var totalAmount = txtTotalAmount.Text;
                var totalDiscount = txtTotalDiscount.Text;
                var totalItem = txtTotalItem.Text;
                statusBill = "Đã thanh toán";
				//Insert vào CSDL
				sql = "INSERT INTO Bill (CreateTime, SubTotal, Status, PaymentMethod, StaffName, Cart_ID, TotalDiscountAmount, TotalAmount, TotalItem) VALUES (";
                sql += "N'" + createTime + "', " + totalAmount + ",N' " + statusBill + "',N'" + paymentMethod + "',N'" + staffName + "'," + cart_ID + "," + totalDiscount + "," + totalAmount + "," + totalItem + ") ";
                sql += "exec [dbo].[UpdateItemInventory]";
                dtBase.DataChange(sql);
                if(statusBill.CompareTo("Đã thanh toán") == 0)
                    MessageBox.Show("Thanh toán thành công!", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                else if (statusBill.CompareTo("Đang chờ thanh toán") == 0)
                    MessageBox.Show("Lưu thành công!", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                //lay lít san pham de in excel
                List<string> list = new List<string>();
                list.Add(txtCus_id.Text);
                list.Add(txtCart_ID.Text);
                list.Add(txtCustomerName.Text);
                list.Add(comboStaffName.Text);
                list.Add(dtpCreatedTime.Text);
                list.Add(txtTotalItem.Text);
                list.Add(txtTotalDiscount.Text);
                list.Add(txtTotalAmount.Text);


              
                


                string sql1 = "";
                sql1 = "delete from [dbo].[SelectedItem] where SelectedItem_ID > 33";
                dtBase.DataChange(sql1);
                Close();

               // var child = new ReceiptCusCus(list);
               // child.Show();
            }
        }

        private void BtnCancelPayment(object sender, EventArgs e)
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


		private void DisplayCart()
		{
			// Thực hiện truy vấn SQL và lấy dữ liệu
			string query = "SELECT TOP 1 c.Cart_ID, cus.Customer_ID, cus.FullName, c.TotalItem, c.TotalDiscountAmount, c.TotalAmount " +
						   "FROM [dbo].[Cart] AS c " +
						   "JOIN [dbo].[Customer] AS cus ON cus.Customer_ID = c.Customer_ID " +
						   "ORDER BY c.Cart_ID DESC";

            // dtBase.DataReader là một đối tượng chưa kết nối và thực hiện truy vấn

            string strConnect = "Data Source=DESKTOP-JSDGA71\\SQLEXPRESS;Initial Catalog=BTL_LTTQ;Integrated Security=True";

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
							txtCart_ID.Text = reader["Cart_ID"].ToString();
                            txtTotalAmount.Text = reader["TotalAmount"].ToString();
                            txtTotalDiscount.Text = reader["TotalDiscountAmount"].ToString();
                            txtTotalItem.Text = reader["TotalItem"].ToString();
						}
						else
						{
							// Xử lý trường hợp không có dữ liệu
							// Ví dụ: labelCart_ID.Text = "Không có dữ liệu";
						}
					}
				}
			}
		}

		private void PaymentFrm_Load(object sender, EventArgs e)
		{
            DisplayCart();
		}

        private void comboStaffName_Validating(object sender, CancelEventArgs e)
        {
            if (string.IsNullOrEmpty(comboStaffName.Text))
            {
                MessageBox.Show("Không được để trống");
                return;
            }
        }

        private void comboPaymentMethod_Validating(object sender, CancelEventArgs e)
        {
            
        }

        private void comboStaffName_Validated(object sender, EventArgs e)
        {
            
        }

        private void comboPaymentMethod_Validated(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(comboPaymentMethod.Text))
            {
                MessageBox.Show("Không được để trống");
                return;
            }
        }
    }
}
