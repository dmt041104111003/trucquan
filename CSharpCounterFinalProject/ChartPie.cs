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
    public partial class ChartPie : Form
    {
        private DataGridView tblItem;
        private DataGridView tblCustomer;
        public ChartPie(DataGridView tblStatItem, DataGridView tblStatCustomer)
        {
            InitializeComponent();
            CenterToScreen();
            tblItem = tblStatItem;
            tblCustomer = tblStatCustomer;
        }

        private void ChartPie_Load(object sender, EventArgs e)
        {
            // item
            chart1.Size = new Size(400, 600);
            ChartArea chartArea = new ChartArea();
            chart1.ChartAreas.Add(chartArea);
            Series series = new Series("Item");
            chart1.Series.Add(series);

            //// customer
            chart2.Size = new Size(400, 600);
            ChartArea chartAreaCustomer = new ChartArea();
            chart2.ChartAreas.Add(chartAreaCustomer);
            Series seriesCustomer = new Series("Customer");
            chart2.Series.Add(seriesCustomer);

            // customer
            foreach (DataGridViewRow row in tblCustomer.Rows)
            {
                if (!row.IsNewRow)
                {
                    string[] name_array = row.Cells["stat_customer_name"].Value.ToString().Split(' ');
                    string cusName = name_array[name_array.Length - 1];    
                    int cusTimes = Convert.ToInt32(row.Cells["stat_customer_times"].Value);
                    int cusQuantity = Convert.ToInt32(row.Cells["stat_customer_quantity"].Value);
                    int cusTotal = (Convert.ToInt32(row.Cells["stat_customer_total"].Value));

                    // Thêm dữ liệu vào biểu đồ tròn
                    DataPoint dataPoint = new DataPoint();
                    //dataPoint.SetValueY(cusTimes);
                    //dataPoint.SetValueY(cusQuantity);
                    dataPoint.SetValueY(cusTotal);
                    dataPoint.Label = cusName;
                    seriesCustomer.Points.Add(dataPoint);
                }
            }
            // item
            foreach (DataGridViewRow row in tblItem.Rows)
            {
                if (!row.IsNewRow)
                {
                    string itemName = row.Cells["inven_name"].Value.ToString();
                    int itemQuantity = Convert.ToInt32(row.Cells["inven_quantity"].Value);

                    // Thêm dữ liệu vào biểu đồ tròn
                    DataPoint dataPoint = new DataPoint();
                    dataPoint.SetValueY(itemQuantity);
                    dataPoint.Label = itemName;
                    series.Points.Add(dataPoint);
                }
            }

            // Đặt tên cho loạt dữ liệu và hiển thị biểu đồ tròn
            series.ChartType = SeriesChartType.Pie;
            seriesCustomer.ChartType = SeriesChartType.Pie;
            chart1.Invalidate();
            chart2.Invalidate();

            Controls.Add(chart1);
            Controls.Add(chart2);


        }
    }
}
