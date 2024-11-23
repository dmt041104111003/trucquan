using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSharpCounterFinalProject.ViewCustomer
{
    public partial class DetailItemCus : Form
    {
        // Declare and constructor global variable in class to use DataBaseProcess
        Classes.DataBaseProcess dtBase = new Classes.DataBaseProcess();
        private string fileImage;
        public DetailItemCus(string itemID,string productName,string price,string imagePath)
        {
            InitializeComponent();
            CenterToScreen();
            txtItemIDDetail.Text = itemID;
            txtItemNameDetail.Text = productName;
            txtPriceDetail.Text = price;
            fileImage = imagePath;
            string image = System.Windows.Forms.Application.StartupPath + "\\Items\\" + fileImage;

            if (File.Exists(imagePath))
            {
                pictureBox1.Image = System.Drawing.Image.FromFile(imagePath);
            }
            else
            {
                // Xử lý lỗi, ví dụ: hiển thị thông báo hoặc thực hiện một hành động khác.
            }
            pictureBox1.Size = new Size(157, 192);
            pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            //pictureBox.Dock = DockStyle.Top;
            pictureBox1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        }

        private void btnAddDetail_Click(object sender, EventArgs e)
        {
            //check số lượng trong kho
            int _number = 0;
            string connectionString = Classes.DataBaseProcess.strConnect;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = "select Quantity from Item where Item_ID = "+txtItemIDDetail.Text+""; 
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            _number = reader.GetInt32(0);
                        }
                    }
                }
            }

            if((int)nudQuantityDetail.Value > _number)
            {
                MessageBox.Show("Sản phẩm trong kho không đủ số lượng");
                nudQuantityDetail.Value = 0;
                return;
            }

            int item_ID = int.Parse(txtItemIDDetail.Text);

            int numofSelectItem = (int)nudQuantityDetail.Value;

            int priceAfterDiscount = 0;

            string sql = "";


            //Insert vào CSDL
            sql = "INSERT INTO SelectedItem (NumberOfSelectedItem, PriceAfterDiscount, Item_ID) VALUES (";
            sql += "N'" + numofSelectItem + "',N'" + priceAfterDiscount + "',N'" + item_ID + "')";
            dtBase.DataChange(sql);
            MessageBox.Show("Thêm thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Close();
        }

        private void btnCancelDetail_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void DetailItemCus_Load(object sender, EventArgs e)
        {

        }

        private void nudQuantityDetail_Validating(object sender, CancelEventArgs e)
        {
            if(nudQuantityDetail.Value < 1)
            {
                MessageBox.Show("Số lượng sản phẩm ít nhất là 1.");
                return;
            }
        }
    }
}
