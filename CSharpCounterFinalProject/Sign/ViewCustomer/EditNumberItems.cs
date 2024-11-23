using CSharpCounterFinalProject.Classes;
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
    public partial class EditNumberItems : Form
    {
        ListItemCus a;
        SelectedItem _item = new SelectedItem();
        Classes.DataBaseProcess dtbase = new DataBaseProcess();
        private string fileImage;
        public EditNumberItems()
        {
            InitializeComponent();
            CenterToScreen();
        }
        public EditNumberItems(string itemID, string productName, string price, string imagePath)
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
        public EditNumberItems(SelectedItem item,ListItemCus b)
        {
            InitializeComponent();
            CenterToScreen();
            this._item.SelectedItem_ID = item.SelectedItem_ID;
            this.a = b;
            string image = "";
            string connectionString = Classes.DataBaseProcess.strConnect;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = "select Item_ID,ItemName,Price,Image from Item where Item_ID in (select Item_ID from SelectedItem where SelectedItem_ID = "+this._item.SelectedItem_ID+")";
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            
                            this.txtItemIDDetail.Text = reader["Item_ID"].ToString();
                            this.txtItemNameDetail.Text = reader["ItemName"].ToString();
                            this.txtPriceDetail.Text = reader["Price"].ToString();
                            image = reader["image"].ToString();
                            
                        }
                    }
                }
            }
            string imagePath = System.Windows.Forms.Application.StartupPath + "\\Items\\" + image;

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
            //numericUpDown1.Value
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //check số lượng trong kho
            int _number = 0;
            string connectionString = Classes.DataBaseProcess.strConnect;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = "select Quantity from Item where Item_ID = " + txtItemIDDetail.Text + "";
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

            if ((int)numericUpDown1.Value > _number)
            {
                MessageBox.Show("Sản phẩm trong kho không đủ số lượng");
                numericUpDown1.Value = 0;
                return;
            }
            if (this.numericUpDown1.Value == 0)
            {
                dtbase.DataChange("delete from SelectedItem\r\nwhere SelectedItem_ID = " + _item.SelectedItem_ID + "");
            }
            else
            {
                dtbase.DataChange("update SelectedItem\r\nset NumberOfSelectedItem = "+ this.numericUpDown1.Value + " where SelectedItem_ID = " + _item.SelectedItem_ID +"");
            }
            MessageBox.Show("Thành công!");
            ////this.Close();
            this.a.tableLayoutPanel.Controls.Clear();
            this.a.ListItemCus_Load(sender, e);
            this.Close();
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void EditNumberItems_Load(object sender, EventArgs e)
        {

        }
    }
}
