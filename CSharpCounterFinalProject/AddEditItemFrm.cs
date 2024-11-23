using System;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;



namespace CSharpCounterFinalProject
{
    public partial class AddEditItemFrm : Form
    {
        // Declare and constructor global variable in class to use DataBaseProcess
        Classes.DataBaseProcess dtBase = new Classes.DataBaseProcess();
        private string fileImage;
        public AddEditItemFrm()
        {
            InitializeComponent();
            CenterToParent();
        }


        public AddEditItemFrm(string strBtn,int item_id, string itemName, string itemType,int quantity,
                              string brand,DateTime releaseDate,int price,string discount, string image)
        {
            InitializeComponent(); 
            CenterToParent();
            if(strBtn == "Xem")
            {
                btnAddNew.Enabled = false;
                btnChooseImage.Enabled = false;
                btnChooseImage.Text = "Hình ảnh";
            }
            if (btnAddNew != null)
            {
                txtId.Text = $"{item_id}";
                btnAddNew.Text = strBtn;
                txtItemName.Text = itemName;
                txtBrand.Text = brand;
                comboItemType.Text = itemType;
                comboDiscount.Text = discount;
                datePickerReleaseDate.Value = releaseDate;
                numericPrice.Value = price;
                numericQuantity.Value = quantity;
                fileImage = image;
                pictureImage.Image = Image.FromFile(Application.StartupPath + "\\Items\\" + image);
            }
        }

        private void BtnAddNewItemClick(object sender, EventArgs e)
        {
            if(btnAddNew.Text == "Thêm Mới")
            {
                string sql = "";
                var ItemName = txtItemName.Text;
                var ItemType = comboItemType.Text;
                var Quantity = (int)numericQuantity.Value;
                var Brand = txtBrand.Text;
                var ReleaseDate = datePickerReleaseDate.Text;
                var Price = (int)numericPrice.Value;
                var Discount = comboDiscount.Text;
                var Image = fileImage;

                int Discount_ID = 0; // Giá trị mặc định là 0
                if (!string.IsNullOrEmpty(Discount))
                {
                    Discount_ID = GetDiscountIDByName(Discount);
                }

                //Insert vào CSDL
                sql = "INSERT INTO Item (ItemName, ItemType, Quantity, Brand, ReleaseDate, Price, Discount_ID, Image) VALUES (";
                sql += "N'" + ItemName + "',N'" + ItemType + "','" +
                       Quantity + "',N'" + Brand + "',N'" + ReleaseDate + "',N'" + Price + "',N'" + Discount_ID + "',N'" + Image + "')";
                dtBase.DataChange(sql);
                MessageBox.Show("Thêm thành công!", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                Close();
            }else if(btnAddNew.Text == "Cập Nhật")
            {
                string sql = "";
                var ItemName = txtItemName.Text;
                var ItemType = comboItemType.Text;
                var Quantity = (int)numericQuantity.Value;
                var Brand = txtBrand.Text;
                var ReleaseDate = datePickerReleaseDate.Text;
                var Price = (int)numericPrice.Value;
                var Discount = comboDiscount.Text;
                var Item_ID = txtId.Text;
                int Discount_ID = 0; // Giá trị mặc định là 0
                string Image = fileImage;
                if (!string.IsNullOrEmpty(Discount))
                {
                    Discount_ID = GetDiscountIDByName(Discount);
                }
                sql = "UPDATE Item SET ";
                sql += "ItemName = N'" + ItemName + "',";
                sql += "ItemType = N'" + ItemType + "',";
                sql += "Quantity = N'" + Quantity + "',";
                sql += "Brand = N'" + Brand + "',";
                sql += "ReleaseDate = N'" + ReleaseDate + "',";
                sql += "Price = N'" + Price + "', ";
                sql += "Image = N'" + Image + "', ";
                sql += "Discount_ID = N'" + Discount_ID + "' ";
                sql += "WHERE Item_ID = " + Item_ID;
                dtBase.DataChange(sql);
                MessageBox.Show("Cập nhật thành công!", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                Close();
            }
            else if(btnAddNew.Text == "Xem")
            {
                Close();
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

        // Hàm này sử dụng để lấy Discount_ID từ bảng Discount dựa trên tên Discount
        private int GetDiscountIDByName(string discountName)
        {
            string sql = "SELECT Discount_ID FROM Discount WHERE Name = N'" + discountName + "'";
            DataTable dataTable = dtBase.DataReader(sql);
            if (dataTable != null && dataTable.Rows.Count > 0)
            {
                return Convert.ToInt32(dataTable.Rows[0]["Discount_ID"]);
            }
            return -1; // Trả về -1 nếu không tìm thấy
        }

        private void BtnChooseImageClick(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "JPEG Images|*.jpg|PNG Images|*.png|All files|*.*";
            openFileDialog.FilterIndex = 1;

            openFileDialog.InitialDirectory = Application.StartupPath;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                pictureImage.Image = Image.FromFile(openFileDialog.FileName);

                fileImage = Path.GetFileName(openFileDialog.FileName);
                MessageBox.Show(openFileDialog.FileName);
            }
        }

        private void AddEditItemFrm_Load(object sender, EventArgs e)
        {

        }

        private void comboItemType_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
