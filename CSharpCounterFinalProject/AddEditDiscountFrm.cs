using System;
using System.Windows.Forms;

namespace CSharpCounterFinalProject
{
    public partial class AddEditDiscountFrm : Form
    {
        // Declare and constructor global variable in class to use DataBaseProcess
        Classes.DataBaseProcess dtBase = new Classes.DataBaseProcess();
        public AddEditDiscountFrm()
        {
            InitializeComponent();
            CenterToScreen();
        }

        public AddEditDiscountFrm(string str,string discount_ID, string name, DateTime startTime, DateTime endTime, 
                                  string type,int percent, int priceAmount, int minPrice)
        {
            InitializeComponent();
            CenterToScreen();
            if (btnAddNew != null)
            {
                btnAddNew.Text = str;
                txtDiscountId.Text = discount_ID;
                txtDiscountName.Text = name;
                dtPickerStart.Value = startTime;
                dtPickerEnd.Value = endTime;
                comboDiscountType.Text = type;
                numericDiscountPercent.Value = percent;
                numericDiscountAmount.Value = priceAmount;
                numericMinPriceToUseDiscount.Value = minPrice;
            }
        }

        private void BtnAddNewDiscountClick(object sender, EventArgs e)
        {
            if (btnAddNew.Text == "Thêm mới")
            {
                string sql = "";
                
                var name = txtDiscountName.Text;
                var startTime = dtPickerStart.Value;
                var endTime = dtPickerEnd.Value;
                var discountType = comboDiscountType.Text;
                var discountPercent = (int)numericDiscountPercent.Value;
                var discountPriceAmount = (int)numericDiscountAmount.Value;
                var discountMinPrice = (int)numericMinPriceToUseDiscount.Value;

                //Insert vào CSDL
                sql = "INSERT INTO Discount (Name, StartTime, EndTime, DiscountType, DiscountPercent, DiscountPriceAmount, MinPriceToUseDiscount) VALUES (";
                sql += "N'" + name + "',N'" + startTime + "','" +
                       endTime + "',N'" + discountType + "',N'" + discountPercent + "',N'" + discountPriceAmount + "',N'" + discountMinPrice + "')";
                dtBase.DataChange(sql);
                MessageBox.Show("Thêm thành công!", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                Close();
            }else if(btnAddNew.Text == "Cập Nhật")
            {
                string sql = "";

                var id = txtDiscountId.Text;
                var name = txtDiscountName.Text;
                var startTime = dtPickerStart.Value;
                var endTime = dtPickerEnd.Value;
                var discountType = comboDiscountType.Text;
                var discountPercent = numericDiscountPercent.Value;
                var discountPriceAmount = numericDiscountAmount.Value;
                var discountMinPrice = (int)numericMinPriceToUseDiscount.Value;

                //Update vào CSDL
                sql = "UPDATE Discount SET ";
                sql += "Name = N'" + name + "',";
                sql += "StartTime = N'" + startTime + "',";
                sql += "EndTime = N'" + endTime + "',";
                sql += "DiscountType = N'" + discountType + "', ";
                sql += "DiscountPercent = N'" + discountPercent + "', ";
                sql += "MinPriceToUseDiscount = N'" + discountMinPrice + "', ";
                sql += "DiscountPriceAmount = N'" + discountPriceAmount + "' ";
                sql += "WHERE Discount_ID = " + id;
                dtBase.DataChange(sql);
                MessageBox.Show("Cập nhật thành công!", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                Close();
            }
        }

        private void AddEditDiscountFrm_Load(object sender, EventArgs e)
        {

        }
    }
}
