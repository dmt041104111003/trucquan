using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace CSharpCounterFinalProject.ViewCustomer
{
	public partial class ReceiptCusCus : Form
	{
        List<string> receiptList = new List<string>();
		public ReceiptCusCus()
		{
			InitializeComponent();
			CenterToScreen();
		}
        public ReceiptCusCus(List<string> data)
        {
            InitializeComponent();
            CenterToScreen();
            this.receiptList = data;
        }

        private void ReceiptCusCus_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            Excel.Application exApp = new Excel.Application();
            Excel.Workbook exBook = exApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet exSheet = (Excel.Worksheet)exBook.Worksheets[1];
            Excel.Range tenTruong = (Excel.Range)exSheet.Cells[1, 1]; //Đưa con trỏ vào ô A1
            tenTruong.Range["A1:D1"].MergeCells = true;
            tenTruong.Range["A1"].Value = "CỬA HÀNG TẠP HÓA VĂN HẢI";
            tenTruong.Range["A2"].Value = "Địa chỉ: số 3 - Cầu Giấy - Hà Nội";
            tenTruong.Range["A3"].Value = "Điện thoại: 0329990540";
            tenTruong.Range["c5:f5"].MergeCells = true;
            tenTruong.Range["C5:F5"].Font.Size = 18;
            tenTruong.Range["C5:F5"].Font.Color = System.Drawing.Color.Red;
            tenTruong.Range["C5"].Value = "HÓA ĐƠN BÁN";
            tenTruong.Range["A7"].Value = "Mã HĐ: " + receiptList[0];
            tenTruong.Range["A8"].Value = "Khách hàng: " + receiptList[2];
            tenTruong.Range["A9"].Value = "Thời gian: " + receiptList[4];
            tenTruong.Range["A10"].Value = "Tổng số sản phẩm: " + receiptList[5];
            tenTruong.Range["A11"].Value = "Tổng khuyến mãi: " + receiptList[6];
            tenTruong.Range["A12"].Value = "Tổng tiền: " + receiptList[7];
            exSheet.Name = "HoaDonBan";
            exBook.Activate();
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                exBook.SaveAs(saveFileDialog.FileName.ToString());
            exApp.Quit();
            this.Close();
        }
    }
}
