using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace CSharpCounterFinalProject
{
    public partial class ChartTable : Form
    {
        private DataGridView tblItem;
        private DataGridView tblCustomer;
        public ChartTable(DataGridView tblStatItem, DataGridView tblStatCustomer)
        {
            InitializeComponent();
            CenterToScreen();
            tblItem = tblStatItem;
            tblCustomer = tblStatCustomer;
        }

        private void ChartTable_Load(object sender, EventArgs e)
        {

            foreach (DataGridViewRow row in tblItem.Rows)
            {
                if (!row.IsNewRow)
                {
                    // Trích xuất dữ liệu từ các cột
                    string itemName = row.Cells["inven_name"].Value.ToString();
                    int itemQuantity = Convert.ToInt32(row.Cells["inven_quantity"].Value);
                    int itemRevenue = (Convert.ToInt32(row.Cells["inven_revenue"].Value)) / 1000;

                    // Thêm dữ liệu vào biểu đồ
                    chart1.Series.Add(new Series(itemName));
                    chart1.Series[itemName].Points.AddXY("Số lượng", itemQuantity);
                    chart1.Series[itemName].Points.AddXY("Doanh thu x 1000", itemRevenue);
                }
            }


            foreach (DataGridViewRow row in tblCustomer.Rows)
            {
                if (!row.IsNewRow)
                {
                    string cusName = row.Cells["stat_customer_name"].Value.ToString();
                    int cusTimes = Convert.ToInt32(row.Cells["stat_customer_times"].Value);
                    int cusQuantity = Convert.ToInt32(row.Cells["stat_customer_quantity"].Value);
                    int cusTotal = (Convert.ToInt32(row.Cells["stat_customer_total"].Value)) / 1000;

                    // Thêm dữ liệu vào biểu đồ
                    chart2.Series.Add(new Series(cusName));
                    chart2.Series[cusName].Points.AddXY("Số lần mua hàng", cusTimes);
                    chart2.Series[cusName].Points.AddXY("Số mặt hàng đã mua", cusQuantity);
                    chart2.Series[cusName].Points.AddXY("Tổng tiền x 1000", cusTotal);
                }
            }
            // Thiết lập tiêu đề và các giá trị trên trục
            chart1.ChartAreas[0].AxisX.Title = "Tên cột";
            chart1.ChartAreas[0].AxisY.Title = "Số liệu";

            chart2.ChartAreas[0].AxisX.Title = "Tên cột";
            chart2.ChartAreas[0].AxisY.Title = "Số liệu";

            // Hiển thị biểu đồ
            chart1.Invalidate();
            chart2.Invalidate();
        }
    }
}
