using CSharpCounterFinalProject.Classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSharpCounterFinalProject.ViewCustomer
{
    public partial class ListItemCus : Form
    {

        // Declare and constructor global variable in class to use DataBaseProcess
        Classes.DataBaseProcess dtBase = new Classes.DataBaseProcess();

        List<SelectedItem> shoppingCart;
        public ListItemCus(string labelNameCart)
        {
            InitializeComponent();
            CenterToScreen();
            labelName.Text = labelNameCart;
        }
        List<SelectedItem> GetProductsFromDatabase()
        {
            List<SelectedItem> items = new List<SelectedItem>();
            string connectionString = Classes.DataBaseProcess.strConnect;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = "select it.ItemName, sel.SelectedItem_ID, sel.NumberOfSelectedItem, sel.PriceAfterDiscount, it.Image, sel.TotalAmount,it.Price " +
                                "from [dbo].[SelectedItem] as sel " +
                                "join [dbo].[Item] as it on it.Item_ID = sel.Item_ID " +
                                "WHERE sel.SelectedItem_ID NOT IN ( " +
                                                       "SELECT TOP 1 SelectedItem_ID " +
                                                       "FROM [dbo].[SelectedItem]" +
                                                       ");";
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            SelectedItem item = new SelectedItem();
                            item.SelectedItem_ID = reader["SelectedItem_ID"].ToString();
                            item.NumberOfSelectedItem = (int)reader["NumberOfSelectedItem"];
                            item.Price = (int)reader["Price"];
                            item.PriceAfterDiscount = reader.GetDecimal(reader.GetOrdinal("PriceAfterDiscount"));
                            item.Image = reader["Image"].ToString();
                            item.Itemname = reader["Itemname"].ToString();
                            item.TotalAmount = reader.GetDecimal(reader.GetOrdinal("TotalAmount"));
                            items.Add(item);
                        }
                    }
                }
            }
            return items;
        }

        public void ListItemCus_Load(object sender, EventArgs e)
        {
            shoppingCart = GetProductsFromDatabase();
            UpdateTableLayoutPanel(shoppingCart);

            tableLayoutPanel.AutoScroll = true;
            // load id
            DisplayInfoUserCus_id();

        }
        void UpdateTableLayoutPanel(List<SelectedItem> products)
        {

            FlowLayoutPanel flowLayoutPanel1 = new FlowLayoutPanel();
            flowLayoutPanel1.Dock = DockStyle.Fill;
            flowLayoutPanel1.AutoScroll = true;
            tableLayoutPanel.Controls.Add(flowLayoutPanel1);
            flowLayoutPanel1.Padding = new Padding(10);


            //tableLayoutPanel1.Controls.Clear();
            tableLayoutPanel.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F)); // 3 cột
            tableLayoutPanel.AutoScroll = true;
            tableLayoutPanel.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;
            tableLayoutPanel.RowCount = 2; // Khởi tạo số hàng là 0
            tableLayoutPanel.ColumnCount = 0;
            tableLayoutPanel.Cursor = Cursors.Hand;

            //tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 50F)); // 2 hàng
            // Lặp qua danh sách sản phẩm


            foreach (var product in products)
            {
                // Tạo `TableLayoutPanel` con với 3 hàng và 1 cột
                TableLayoutPanel productLayoutPanel = new TableLayoutPanel();
                productLayoutPanel.RowCount = 3;
                productLayoutPanel.ColumnCount = 1;
                productLayoutPanel.Tag = product;
                productLayoutPanel.Height = 300;
                productLayoutPanel.Width = 500;
                productLayoutPanel.Click += ProductPanel_Click;
                productLayoutPanel.Margin = new Padding(30);
                productLayoutPanel.Anchor = AnchorStyles.None;

                int productIndex = 0;
                int columnsPerRow = 2;

                // PictureBox
                PictureBox pictureBox = new PictureBox();
                string fileImage = product.Image.ToString();
                string imagePath = Application.StartupPath + "\\Items\\" + fileImage;

                if (File.Exists(imagePath))
                {
                    pictureBox.Image = System.Drawing.Image.FromFile(imagePath);
                }
                else
                {
                    // Xử lý lỗi, ví dụ: hiển thị thông báo hoặc thực hiện một hành động khác.
                }
                pictureBox.SizeMode = PictureBoxSizeMode.Zoom;
                pictureBox.Size = new Size(130, 130);
                //pictureBox.Dock = DockStyle.Top;
                pictureBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

                // nameLabel
                Label nameLabel = new Label();
                nameLabel.Text = $"Tên sp: {product.Itemname}";
                nameLabel.TextAlign = ContentAlignment.MiddleCenter; // Đặt vị trí giữa cho Label
                nameLabel.ForeColor = ColorTranslator.FromHtml("#609966");
                nameLabel.Dock = DockStyle.Bottom;

                // priceAfterDiscount
                Label priceAfterDiscount = new Label();
                priceAfterDiscount.Text = $"Giá KM: {product.PriceAfterDiscount} vnđ";
                priceAfterDiscount.TextAlign = ContentAlignment.MiddleCenter; // Đặt vị trí giữa cho Label
                priceAfterDiscount.Dock = DockStyle.Bottom;
                priceAfterDiscount.ForeColor = ColorTranslator.FromHtml("#609966");

                // numberOfSelectedItem
                Label numberOfSelectedItem = new Label();
                numberOfSelectedItem.Text = $"Số sp: {product.NumberOfSelectedItem}";
                numberOfSelectedItem.TextAlign = ContentAlignment.MiddleCenter; // Đặt vị trí giữa cho Label
                numberOfSelectedItem.Dock = DockStyle.Bottom;
                numberOfSelectedItem.ForeColor = ColorTranslator.FromHtml("#609966");

                // totalAmount
                Label totalAmount = new Label();
                totalAmount.Text = $"=>Tổng tiền: {product.TotalAmount}";
                totalAmount.TextAlign = ContentAlignment.MiddleCenter; // Đặt vị trí giữa cho Label
                totalAmount.Dock = DockStyle.Bottom;
                totalAmount.ForeColor = ColorTranslator.FromHtml("#609966");

                // price
                Label price = new Label();
                price.Text = $"Giá gốc: {product.Price}";
                price.TextAlign = ContentAlignment.MiddleCenter; // Đặt vị trí giữa cho Label
                price.Dock = DockStyle.Bottom;
                price.ForeColor = ColorTranslator.FromHtml("#609966");
                price.Font = new Font(price.Font, FontStyle.Strikeout);

                // Thêm các điều khiển vào `TableLayoutPanel` con
                productLayoutPanel.Controls.Add(pictureBox, 0, 0);
                productLayoutPanel.Controls.Add(nameLabel, 0, 1);
                productLayoutPanel.Controls.Add(price, 0, 2);
                productLayoutPanel.Controls.Add(priceAfterDiscount, 0, 3);
                productLayoutPanel.Controls.Add(numberOfSelectedItem, 0, 4);
                productLayoutPanel.Controls.Add(totalAmount, 0, 5);

                //// Thêm `TableLayoutPanel` con vào `TableLayoutPanel` chính
                //tableLayoutPanel1.Controls.Add(productLayoutPanel, productIndex % 2, productIndex / 2);
                //productIndex++;

                // Thêm sản phẩm vào `FlowLayoutPanel`
                flowLayoutPanel1.Controls.Add(productLayoutPanel);

                productLayoutPanel.Location = new Point((flowLayoutPanel1.ClientSize.Width - productLayoutPanel.Width) / 2, productLayoutPanel.Location.Y);

                // Tính toán vị trí hàng và cột
                int row = productIndex / columnsPerRow;
                int column = productIndex % columnsPerRow;

                // Đặt kích thước của `TableLayoutPanel` con để có 2 cột
                productLayoutPanel.Width = tableLayoutPanel.ClientSize.Width / columnsPerRow;

                // Đặt vị trí cho `TableLayoutPanel` con
                flowLayoutPanel1.SetFlowBreak(productLayoutPanel, column == 0);

                productIndex++;
            }
        }

        private void ProductPanel_Click(object sender, EventArgs e)
        {


            // Ép kiểu sender thành TableLayoutPanel để lấy dữ liệu từ bên trong
            TableLayoutPanel productPanel = (TableLayoutPanel)sender;
            SelectedItem product = (SelectedItem)productPanel.Tag;

            var children = new EditNumberItems(product,this);
            children.Show();


            

        }

        private void btbPaymentCus_Click_1(object sender, EventArgs e)
        {
            string sql = "";
            string customer_id = labelCus_IDCus.Text;
            if (!string.IsNullOrEmpty(labelName.Text))
            {
                // cart
                //Insert vào CSDL
                sql = "INSERT INTO Cart (TotalItem, Customer_ID, SelectedItem_ID, TotalDiscountAmount, TotalAmount) VALUES (";
                sql += "0 ,N'" + customer_id + "', 33, 0, 0)";
                dtBase.DataChange(sql);
                MessageBox.Show("Tạo thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                var child = new PaymentFrm(customer_id, labelName.Text);
                child.Show();
            }
            Close();
        }

        private void DisplayInfoUserCus_id()
        {
            // Thực hiện truy vấn SQL và lấy dữ liệu
            string query = "select c.Customer_ID " +
                            "from[dbo].[Customer] as c " +
                            "where c.Username = N'" + labelName.Text + "' ";
            //"and c.Password = N'" + txtInfoPassCus.Text + "'";

            // dtBase.DataReader là một đối tượng chưa kết nối và thực hiện truy vấn

            string strConnect = "Data Source=VanHai\\SQLEXPRESS;" +
                            "DataBase=BTL_LTTQ;User ID=sa;" +
                            "Password=123;Integrated Security=false";
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
                            labelCus_IDCus.Text = reader["Customer_ID"].ToString();
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

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
