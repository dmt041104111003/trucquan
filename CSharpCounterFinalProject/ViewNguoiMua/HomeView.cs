using CSharpCounterFinalProject.Classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSharpCounterFinalProject.ViewNguoiMua
{
    public partial class HomeView : Form
    {
        DataBaseProcess db = new DataBaseProcess(); private bool isTheLoaiLoaded = false;
        private TabPage currentTab; 

        public HomeView(string name)
        {
            InitializeComponent();
            label1.Text += " " + name;
            tableSanPhams.ColumnCount = 5;
            tableSanPhams.RowCount = 2;

            for (int i = 0; i < tableSanPhams.ColumnCount; i++) tableSanPhams.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 161));
            
            for (int i = 0; i < tableSanPhams.RowCount; i++) tableSanPhams.RowStyles.Add(new RowStyle(SizeType.Absolute, 231));

            tableSanPhams.Size = new Size(161 * tableSanPhams.ColumnCount, 231 * tableSanPhams.RowCount);
            cbTheloai.DropDown += CbTheloai_DropDown;
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void HomeView_Load(object sender, EventArgs e)
        {
            this.Size = new Size(1000, 600);
            splitContainer2.SplitterDistance = 100;
            tabControl1.SelectedTab = tabSanPham;
            currentTab = tabSanPham;
            LoadCacSanPham();
            LoadTheLoai();
            CenterToScreen();


        }
        private void CbTheloai_DropDown(object sender, EventArgs e)
        {
            if (!isTheLoaiLoaded)
            {
                LoadTheLoai();
                isTheLoaiLoaded = true; 
            }
        }

    
        private void LoadTheLoai()
        {
            try
            {
                string query = "SELECT TenTheLoai FROM TheLoai"; 
                DataTable dt = db.DataReader(query);

           
                cbTheloai.Items.Clear();

       
                foreach (DataRow row in dt.Rows)
                {
                    cbTheloai.Items.Add(row["TenTheLoai"].ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tải thể loại: " + ex.Message);
            }
        }
        private void LoadCacSanPham()
        {
            string query = @"
SELECT TOP 5 
    s.MaSach,
    s.TenSach,
    s.DonGia,
    ISNULL(km.KM, 0) AS GiamGia,
    CASE 
        WHEN km.MaSach IS NOT NULL 
        AND GETDATE() BETWEEN km.ThoigianBatDau AND km.ThoiGianKetThuc 
        THEN (s.DonGia - (s.DonGia * km.KM / 100))
        ELSE s.DonGia 
    END AS GiaMoi,
    s.Anh,
    CASE 
        WHEN km.MaSach IS NOT NULL 
        AND GETDATE() BETWEEN km.ThoigianBatDau AND km.ThoiGianKetThuc 
        THEN 1 
        ELSE 0 
    END AS DangGiamGia
FROM Sach s 
LEFT JOIN KhuyenMai km ON s.MaSach = km.MaSach";

            DataTable dt = db.DataReader(query);
            foreach (DataRow dr in dt.Rows)
            {
                string linksp = dr["Anh"].ToString();
                string tensp = dr["TenSach"].ToString();
                string gia = dr["GiaMoi"].ToString();
                string masp = dr["MaSach"].ToString();
                bool dangGiamGia = Convert.ToInt32(dr["DangGiamGia"]) == 1;
                decimal giaGoc = Convert.ToDecimal(dr["DonGia"]);
                AddProductPanel(tableSanPhams, linksp, tensp, gia, masp, dangGiamGia, giaGoc);
            }
        }

     
        public void AddProductPanel(TableLayoutPanel tableSanPhams, string tenfileanh, string label1Text,
     string label2Text, string masp, bool dangGiamGia, decimal giaGoc)
        {
            TableLayoutPanel innerTable = new TableLayoutPanel
            {
                ColumnCount = 2, 
                RowCount = 3,
                Size = new Size(161, 231)
            };

            innerTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60)); 
            innerTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40)); 

            innerTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 161));
            innerTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));   
            innerTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));  

            PictureBox pictureBox = new PictureBox
            {
                Dock = DockStyle.Fill,
                BackColor = Color.LightGray,
                SizeMode = PictureBoxSizeMode.StretchImage
            };

            string imagePath = System.Windows.Forms.Application.StartupPath + "\\AnhSP\\" + tenfileanh;

            try
            {
                pictureBox.Image = Image.FromFile(imagePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi tải ảnh: {ex.Message}\nĐường dẫn: {imagePath}");
            }
            innerTable.Controls.Add(pictureBox, 0, 0);
            innerTable.SetColumnSpan(pictureBox, 2); 


            Label label1 = new Label
            {
                Text = label1Text,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                BackColor = Color.LightYellow
            };
            innerTable.Controls.Add(label1, 0, 1);
            innerTable.SetColumnSpan(label1, 2); 



            Label label2 = new Label  
            {
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                BackColor = Color.LightYellow
            };

            if (dangGiamGia)
            {
                label2.Text = $"{label2Text} VND\n(Giá gốc: {giaGoc} VND)";
            }
            else
            {
                label2.Text = $"{giaGoc} VND";
            }

            innerTable.Controls.Add(label2, 0, 2);

            Label buyNowLabel = new Label
            {
                Text = "Mua ngay",
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                ForeColor = Color.Red,
                BackColor = Color.LightYellow
            };
                 buyNowLabel.Click += (sender, e) => BuyNowLabel_Click(sender, e, label1Text, label2Text,tenfileanh,masp);

            innerTable.Controls.Add(buyNowLabel, 1, 2);



            tableSanPhams.Controls.Add(innerTable);
        }


        private void BuyNowLabel_Click(object sender, EventArgs e, string label1Text, string label2Text, string anh,string masp)
        {

            ChiTietSanPhamView chitietsp = new ChiTietSanPhamView(label1Text, label2Text,anh,masp);
            this.Close();
            chitietsp.ShowDialog();

        }

    



        private void btnGioHang_Click(object sender, EventArgs e)
        {
            GioHangView gioHangView = new GioHangView();
            gioHangView.Show();
        }

        private void btnDangXuat_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSanPham_Click(object sender, EventArgs e)
        {
            LoadCacSanPham();
        }

        private void tabSanPham_Click(object sender, EventArgs e)
        {
            tabSanPham.AutoScroll = true;
            this.AutoScroll = true;

        }




        //private void btTimkiem_Click(object sender, EventArgs e)
        //{
        //    string tenSach = txtTensanpham.Text.Trim();
        //    string tacGia = txtTentacgia.Text.Trim();
        //    string theLoai = cbTheloai.SelectedItem?.ToString();
        //    string giaTu = txtGiatu.Text.Trim();
        //    string giaDen = txtGiaden.Text.Trim();

        //    StringBuilder queryBuilder = new StringBuilder(@"
        //SELECT TOP 5 
        //    s.MaSach,
        //    s.TenSach,
        //    s.DonGia,
        //    ISNULL(km.KM, 0) AS GiamGia,
        //    CASE 
        //        WHEN km.MaSach IS NOT NULL 
        //        AND GETDATE() BETWEEN km.ThoigianBatDau AND km.ThoiGianKetThuc 
        //        THEN (s.DonGia - (s.DonGia * km.KM / 100))
        //        ELSE s.DonGia 
        //    END AS GiaMoi,
        //    s.Anh,
        //    CASE 
        //        WHEN km.MaSach IS NOT NULL 
        //        AND GETDATE() BETWEEN km.ThoigianBatDau AND km.ThoiGianKetThuc 
        //        THEN 1 
        //        ELSE 0 
        //    END AS DangGiamGia
        //FROM Sach s 
        //LEFT JOIN KhuyenMai km ON s.MaSach = km.MaSach 
        //WHERE 1=1");


        //        if (!string.IsNullOrEmpty(tenSach))
        //        queryBuilder.Append($" AND s.TenSach LIKE N'%{tenSach}%'");

        //    if (!string.IsNullOrEmpty(tacGia))
        //        queryBuilder.Append($" AND s.TacGia LIKE N'%{tacGia}%'");

        //    if (!string.IsNullOrEmpty(theLoai))
        //        queryBuilder.Append($" AND s.MaTheLoai IN (SELECT MaTheLoai FROM TheLoai WHERE TenTheLoai = N'{theLoai}')");


        //        if (!string.IsNullOrEmpty(giaTu) && decimal.TryParse(giaTu, out decimal giaMin))
        //        queryBuilder.Append(@" AND (
        //        CASE 
        //            WHEN km.MaSach IS NOT NULL 
        //            AND GETDATE() BETWEEN km.ThoigianBatDau AND km.ThoiGianKetThuc 
        //            THEN (s.DonGia - (s.DonGia * km.KM / 100))
        //            ELSE s.DonGia 
        //        END) >= " + giaMin);

        //    if (!string.IsNullOrEmpty(giaDen) && decimal.TryParse(giaDen, out decimal giaMax))
        //        queryBuilder.Append(@" AND (
        //        CASE 
        //            WHEN km.MaSach IS NOT NULL 
        //            AND GETDATE() BETWEEN km.ThoigianBatDau AND km.ThoiGianKetThuc 
        //            THEN (s.DonGia - (s.DonGia * km.KM / 100))
        //            ELSE s.DonGia 
        //        END) <= " + giaMax);


        //        LoadKetQuaTimKiem(queryBuilder.ToString());
        //}

        private void btTimkiem_Click(object sender, EventArgs e)
        {
            string tenSach = txtTensanpham.Text.Trim();
            string tacGia = txtTentacgia.Text.Trim();
            string theLoai = cbTheloai.SelectedItem?.ToString();
            string giaTu = txtGiatu.Text.Trim();
            string giaDen = txtGiaden.Text.Trim();

            StringBuilder queryBuilder = new StringBuilder();

            if (currentTab == tabSanPham)
            {
                queryBuilder.Append(@"
    SELECT TOP 5 
        s.MaSach,
        s.TenSach,
        s.DonGia,
        ISNULL(km.KM, 0) AS GiamGia,
        CASE 
            WHEN km.MaSach IS NOT NULL 
            AND GETDATE() BETWEEN km.ThoigianBatDau AND km.ThoiGianKetThuc 
            THEN (s.DonGia - (s.DonGia * km.KM / 100))
            ELSE s.DonGia 
        END AS GiaMoi,
        s.Anh,
        CASE 
            WHEN km.MaSach IS NOT NULL 
            AND GETDATE() BETWEEN km.ThoigianBatDau AND km.ThoiGianKetThuc 
            THEN 1 
            ELSE 0 
        END AS DangGiamGia
    FROM Sach s 
    LEFT JOIN KhuyenMai km ON s.MaSach = km.MaSach 
    WHERE 1=1");

                if (!string.IsNullOrEmpty(tenSach))
                    queryBuilder.Append($" AND s.TenSach LIKE N'%{tenSach}%'");

                if (!string.IsNullOrEmpty(tacGia))
                    queryBuilder.Append($" AND s.TacGia LIKE N'%{tacGia}%'");

                if (!string.IsNullOrEmpty(theLoai))
                    queryBuilder.Append($" AND s.MaTheLoai IN (SELECT MaTheLoai FROM TheLoai WHERE TenTheLoai = N'{theLoai}')");

                if (!string.IsNullOrEmpty(giaTu) && decimal.TryParse(giaTu, out decimal giaMin))
                    queryBuilder.Append(@" AND (
            CASE 
                WHEN km.MaSach IS NOT NULL 
                AND GETDATE() BETWEEN km.ThoigianBatDau AND km.ThoiGianKetThuc 
                THEN (s.DonGia - (s.DonGia * km.KM / 100))
                ELSE s.DonGia 
            END) >= " + giaMin);

                if (!string.IsNullOrEmpty(giaDen) && decimal.TryParse(giaDen, out decimal giaMax))
                    queryBuilder.Append(@" AND (
            CASE 
                WHEN km.MaSach IS NOT NULL 
                AND GETDATE() BETWEEN km.ThoigianBatDau AND km.ThoiGianKetThuc 
                THEN (s.DonGia - (s.DonGia * km.KM / 100))
                ELSE s.DonGia 
            END) <= " + giaMax);

            }
            else if (currentTab == tabSachmienphi)
            {
                queryBuilder.Append(@"
            SELECT 
                s.MaSach,
                s.TenSach,
                s.DonGia,
                s.Anh
            FROM Sach s
            WHERE s.DonGia = 0");

                if (!string.IsNullOrEmpty(tenSach))
                    queryBuilder.Append($" AND s.TenSach LIKE N'%{tenSach}%'");

                if (!string.IsNullOrEmpty(tacGia))
                    queryBuilder.Append($" AND s.TacGia LIKE N'%{tacGia}%'");

                if (!string.IsNullOrEmpty(theLoai))
                    queryBuilder.Append($" AND s.MaTheLoai IN (SELECT MaTheLoai FROM TheLoai WHERE TenTheLoai = N'{theLoai}')");
            }
           

            if (currentTab == tabSanPham)
            {
                tableSanPhams.Controls.Clear();
                LoadKetQuaTimKiem(queryBuilder.ToString(), tableSanPhams);
            }
            else if (currentTab == tabSachmienphi)
            {
                tableSanPhamsSachmienphi.Controls.Clear();
                LoadKetQuaTimKiem(queryBuilder.ToString(), tableSanPhamsSachmienphi);
            }
        
        }



        private void LoadKetQuaTimKiem(string query, TableLayoutPanel tableToLoad)
        {
            try
            {
                DataTable dt = db.DataReader(query);

      
                tableToLoad.Controls.Clear();

                // Thêm các sản phẩm vào bảng
                foreach (DataRow dr in dt.Rows)
                {
                    string linksp = dr["Anh"].ToString(); 
                    string tensp = dr["TenSach"].ToString(); 
                    string masp = dr["MaSach"].ToString(); 
                    string gia = string.Empty; // Giá hiển thị

                    if (currentTab == tabSanPham)
                    {
                        // Tab Sản phẩm: Hiển thị giá đã giảm (GiaMoi) và trạng thái giảm giá (DangGiamGia)
                        gia = dr["GiaMoi"].ToString();
                        bool dangGiamGia = Convert.ToInt32(dr["DangGiamGia"]) == 1; // Kiểm tra có giảm giá không
                        decimal giaGoc = Convert.ToDecimal(dr["DonGia"]);
                        AddProductPanel(tableToLoad, linksp, tensp, gia, masp, dangGiamGia, giaGoc);
                    }
                    else if (currentTab == tabSachmienphi)
                    {
                        // Tab Sách Miễn Phí: Hiển thị giá gốc (DonGia)
                        gia = dr["DonGia"].ToString(); 
                        decimal giaGoc = Convert.ToDecimal(dr["DonGia"]);
                        AddProductPanel(tableToLoad, linksp, tensp, gia, masp, false, giaGoc); // Không có giảm giá
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi tải kết quả tìm kiếm: {ex.Message}");
            }
        }


        private void tabSachmienphi_Click(object sender, EventArgs e)
        {
        
        }
   
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            currentTab = tabControl1.SelectedTab;
            if (currentTab == tabSachmienphi)
            {
                txtGiatu.Enabled = false;
                txtGiaden.Enabled = false;
         
            }
            else if (currentTab == tabSanPham)
            {
                txtGiatu.Enabled = true;
                txtGiaden.Enabled = true;
            }

          
            if (currentTab == tabSanPham)
            {
                tableSanPhams.Controls.Clear(); 
                LoadCacSanPham(); 
            }
            else if (currentTab == tabSachmienphi)
            {
                tableSanPhamsSachmienphi.Controls.Clear(); 
                LoadSachMienPhi();
            }
            //else if (currentTab == tabMoiphathanh)
            //{
            //    tableSanPhamsPhathanh.Controls.Clear();
            //    LoadSachMoiPhatHanh(); 
            //}
            //else if (currentTab == tabSapphathanh)
            //{
            //    tableSanPhamsSapphathanh.Controls.Clear(); 
            //    LoadSachSapPhatHanh(); 
            //}
        }

        private void LoadSachMienPhi()
        {
            try
            {
      
                string query = @"
        SELECT 
            s.MaSach,
            s.TenSach,
            s.DonGia,
            s.Anh
        FROM Sach s
        WHERE s.DonGia = 0";

                DataTable dt = db.DataReader(query);

          

                // Thêm các sản phẩm miễn phí vào bảng
                foreach (DataRow dr in dt.Rows)
                {
                    string linksp = dr["Anh"].ToString();
                    string tensp = dr["TenSach"].ToString();
                    string masp = dr["MaSach"].ToString();
                    decimal giaGoc = Convert.ToDecimal(dr["DonGia"]);
                    AddProductPanel(tableSanPhamsSachmienphi, linksp, tensp, "0", masp, false, giaGoc);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi tải sách miễn phí: {ex.Message}");
            }
        }

    }
}
