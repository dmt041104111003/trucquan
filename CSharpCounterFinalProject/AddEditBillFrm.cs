using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using System.Data.SqlClient;

namespace CSharpCounterFinalProject
{
    public partial class AddEditBillFrm : Form
    {
        // Declare and constructor global variable in class to use DataBaseProcess
        Classes.DataBaseProcess dtBase = new Classes.DataBaseProcess();
        public AddEditBillFrm()
        {
            InitializeComponent();
            CenterToScreen();
        }

        // payment
        private void BtnPayClick(object sender, EventArgs e)
        {
            int cart_ID = int.Parse(labelCart_ID.Text);
            string nameCustomer = labelCustomerNameBill.Text;
            int totalItem = int.Parse(labelTotalItem.Text);
            double totalDiscount = double.Parse(labelTotalDiscount.Text);
            double totalAmount = double.Parse(labelTotalAmount.Text);
            if(cart_ID == 0 || string.IsNullOrEmpty(nameCustomer)  == true || totalItem == 0 || totalAmount == 0)
            {
                MessageBox.Show("Dữ liệu không hợp lệ, Vui lòng hoàn thành các mục được yêu cầu trên!");
                return;
            }
            if(sender.Equals(btnPay))
            {
                string statusBill = "Đã thanh toán";

                var paymentView = new PaymentFrm(cart_ID, nameCustomer, totalItem, totalDiscount, totalAmount,  statusBill);
                paymentView.Show();
            }
            else if(sender.Equals(btnSave))
            {
                string statusBill = "Mua hàng trả sau";

                var paymentView = new PaymentFrm(cart_ID, nameCustomer, totalItem, totalDiscount, totalAmount, statusBill);
                paymentView.Show();
            }


        }

        private void BtnSearchCustomerBillClick(object sender, EventArgs e)
        {
            tblSearchedCustomer.Columns["Customer_ID"].DataPropertyName = "Customer_ID";
            tblSearchedCustomer.Columns["FullName"].DataPropertyName = "FullName";

            var name = txtSearchCustomer.Text;
            string sql = "SELECT Customer_ID, FullName " +
                         "FROM Customer  " +
                         "WHERE FullName LIKE (N'%" + name + "%')";

            tblSearchedCustomer.DataSource = dtBase.DataReader(sql);
        }

        private void BtnSearchItemBillClick(object sender, EventArgs e)
        {
            string sql = "";
            tblSearchedItem.Columns["Item_ID"].DataPropertyName = "Item_ID";
            tblSearchedItem.Columns["ItemName"].DataPropertyName = "ItemName";
            tblSearchedItem.Columns["Quantity"].DataPropertyName = "Quantity";

            var name = txtSearchItem.Text;
            var quantity = (int)numericSelectedQuantity.Value;
            if (quantity > 0)
            {
                sql = "SELECT Item_ID, ItemName, Quantity " +
                             "FROM Item  " +
                             "WHERE ItemName LIKE (N'%" + name + "%') AND Quantity > "+ quantity;
            }else
            {
                sql = "SELECT Item_ID, ItemName, Quantity " +
                            "FROM Item  " +
                            "WHERE ItemName LIKE (N'%" + name + "%')";
            }

            tblSearchedItem.DataSource = dtBase.DataReader(sql);
        }

        private void TblCustomerCellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == tblSearchedCustomer.Columns["tblCustomerColSelect"].Index)
            {
                DataGridViewRow row = tblSearchedCustomer.Rows[e.RowIndex];
                string name = row.Cells["FullName"].Value.ToString();
                int customer_ID = int.Parse(row.Cells["Customer_ID"].Value.ToString());
                labelCustomerNameBill.Text = name;
                labelCustomer_IDBill.Text = customer_ID.ToString();
                //DisplayCart();
            }
        }

        private void TblItemCellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == tblSearchedItem.Columns["tblItemColSelect"].Index)
            {
                DataGridViewRow row = tblSearchedItem.Rows[e.RowIndex];
                int item_ID = int.Parse(row.Cells["Item_ID"].Value.ToString());

                int numofSelectItem = (int)numericSelectedQuantity.Value;

                int priceAfterDiscount = 0;

                string sql = "";
                

                //Insert vào CSDL
                sql = "INSERT INTO SelectedItem (NumberOfSelectedItem, PriceAfterDiscount, Item_ID) VALUES (";
                sql += "N'" + numofSelectItem + "',N'" + priceAfterDiscount + "',N'" + item_ID + "')";
                dtBase.DataChange(sql);

                DisplayBillDetail();

                //string customer_id = labelCustomer_IDBill.Text;
                //if (!string.IsNullOrEmpty(labelCustomerNameBill.Text))
                //{
                //    // cart
                //    //Insert vào CSDL
                //    sql = "INSERT INTO Cart (TotalItem, Customer_ID, SelectedItem_ID, TotalDiscountAmount, TotalAmount) VALUES (";
                //    sql += "0 ,N'" + customer_id + "', 33, 0, 0)";
                //    dtBase.DataChange(sql);
                //    MessageBox.Show("Tạo thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //    DisplayCart();
                //}

            }
        }

        // create cart
        private void BtnCreateCart(object sender, EventArgs e)
        {
            string sql = "";
            string customer_id = labelCustomer_IDBill.Text;
            if (!string.IsNullOrEmpty(labelCustomerNameBill.Text))
            {
                // cart
                //Insert vào CSDL
                sql = "INSERT INTO Cart (TotalItem, Customer_ID, SelectedItem_ID, TotalDiscountAmount, TotalAmount) VALUES (";
                sql += "0 ,N'" + customer_id + "', 33, 0, 0)";
                dtBase.DataChange(sql);
                MessageBox.Show("Tạo thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                DisplayCart();
                btnCreateCartBill.Enabled = false;
            }
        }
        

        private void TblBillCellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == tblBillDetail.Columns["btnEditBill"].Index)
            {

                DataGridViewRow row = tblBillDetail.Rows[e.RowIndex];

                string txtName = row.Cells["ItemNameBill"].Value.ToString();
                int selectedItem_ID = int.Parse(row.Cells["SelectedItem_ID"].Value.ToString());
                int numberOfItemChoice = int.Parse(row.Cells["NumberOfSelectedItem"].Value.ToString());

                numericSelectedQuantity.Focus();

                txtSearchItem.Text = txtName;
                txtSearchItem.Enabled = false;
                numericSelectedQuantity.Value = numberOfItemChoice;
                txtBillID.Text = selectedItem_ID.ToString();
                btnEditBillItem.Enabled = true;
            }
            else if (e.RowIndex >= 0 && e.ColumnIndex == tblBillDetail.Columns["btnRemoveBill"].Index)
            {
                var title = "Xác nhận xóa";
                var msg = "Bạn có chắc chắn muốn xóa bản ghi này hay không?";
                var ans = ShowComfirmDialog(msg, title);

                DataGridViewRow row = tblBillDetail.Rows[e.RowIndex];
                int selectedItem_ID = int.Parse(row.Cells["SelectedItem_ID"].Value.ToString());

                if (ans == DialogResult.Yes)
                {
                    string sql = "";
                    sql = "DELETE FROM SelectedItem WHERE SelectedItem_ID = N'" + selectedItem_ID + "'";
                    dtBase.DataChange(sql);
                    MessageBox.Show("Xóa thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    DisplayBillDetail();
                }
                DisplayCart();
            }
        }

        private DialogResult ShowComfirmDialog(string msg, string title)
        {
            return MessageBox.Show(msg, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        }

        // edit bill
        private void BtnEditBillClick(object sender, EventArgs e)
        {
            int numofSelectItem = (int)numericSelectedQuantity.Value;
            int selectedItem_ID = int.Parse(txtBillID.Text);

            string sql = "";


            sql = "UPDATE SelectedItem SET ";
            sql += "NumberOfSelectedItem = N'" + numofSelectItem + "' ";
            sql += "WHERE SelectedItem_ID = " + selectedItem_ID;
            dtBase.DataChange(sql);
            MessageBox.Show("Cập nhật thành công!", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

            DisplayBillDetail();

            btnEditBillItem.Enabled = false;
            txtSearchItem.Enabled = true;

            DisplayCart();
        }

        // Form load
        private void AddEditBillFrm_Load(object sender, EventArgs e)
        {
            // Display Bill
            DisplayBillDetail();
            // Display cart
            //DisplayCart();
        }

        // display detail
        private void DisplayBillDetail()
        {
            tblBillDetail.Columns["SelectedItem_ID"].DataPropertyName = "SelectedItem_ID";
            tblBillDetail.Columns["Item_IDBill"].DataPropertyName = "Item_ID";
            tblBillDetail.Columns["ItemNameBill"].DataPropertyName = "ItemName";
            tblBillDetail.Columns["NumberOfSelectedItem"].DataPropertyName = "NumberOfSelectedItem";
            tblBillDetail.Columns["PriceBill"].DataPropertyName = "Price";
            tblBillDetail.Columns["PriceAfterDiscount"].DataPropertyName = "PriceAfterDiscount";
            tblBillDetail.Columns["TotalAmount"].DataPropertyName = "TotalAmount";
            // Display customer
            tblBillDetail.DataSource = dtBase.DataReader("SELECT s.SelectedItem_ID, it.Item_ID, it.ItemName, s.NumberOfSelectedItem, it.Price, s.PriceAfterDiscount, s.TotalAmount " +
                                                       "from [dbo].[SelectedItem] as s " +
                                                       "join [dbo].[Item] as it on it.Item_ID = s.Item_ID " +
                                                       "WHERE s.SelectedItem_ID NOT IN ( " +
                                                       "SELECT TOP 1 SelectedItem_ID " +
                                                       "FROM [dbo].[SelectedItem]" +
                                                       ");");
        }

        // display cart
        private void DisplayCart()
        {
            // Thực hiện truy vấn SQL và lấy dữ liệu
            string query = "SELECT TOP 1 c.Cart_ID, cus.Customer_ID, cus.FullName, c.TotalItem, c.TotalDiscountAmount, c.TotalAmount " +
                           "FROM [dbo].[Cart] AS c " +
                           "JOIN [dbo].[Customer] AS cus ON cus.Customer_ID = c.Customer_ID " +
                           "ORDER BY c.Cart_ID DESC";

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
                            labelCart_ID.Text = reader["Cart_ID"].ToString();
                            //labelCustomer_IDBill.Text = reader["Customer_ID"].ToString();
                            //labelCustomerNameBill.Text = reader["FullName"].ToString();
                            labelTotalItem.Text =reader["TotalItem"].ToString();
                            labelTotalDiscount.Text =reader["TotalDiscountAmount"].ToString();
                            labelTotalAmount.Text =reader["TotalAmount"].ToString();
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

        // refresh cart
        private void btnRefreshCart_Click(object sender, EventArgs e)
        {
            DisplayCart();
        }

        private void BtnRemoveCart(object sender, EventArgs e)
        {
            labelCart_ID.Text = "0";
            labelCustomer_IDBill.Text = "Code";
            labelCustomerNameBill.Text = "Code";
            labelTotalItem.Text = "0";
            labelTotalDiscount.Text = "0";
            labelTotalAmount.Text = "0";
        }

        private void BtnCancelCart(object sender, EventArgs e)
        {
            Close();
        }

		private void labelCustomer_IDBill_Click(object sender, EventArgs e)
		{

		}
	}
}
