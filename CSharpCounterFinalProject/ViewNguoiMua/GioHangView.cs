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
    public partial class GioHangView : Form
    {
        DataBaseProcess db = new DataBaseProcess();
        private bool isUserChanging = true; // Biến kiểm soát

        public GioHangView()
        {
            InitializeComponent();

            panel1.AutoScroll = true;
            panel1.Size = new Size(panel1.Width, 350);
            tableSanPhams.Size = new Size(panel1.Width, panel1.Height);
            splitContainer2.SplitterDistance = 350;

            // Đặt splitContainer3 ở dưới cùng của form
            splitContainer3.Dock = DockStyle.Bottom;

            // Cố định splitter để ngăn người dùng thay đổi vị trí
            splitContainer3.IsSplitterFixed = false;

            // Thiết lập minWidth cho hai panel bên trong
            splitContainer3.Panel1MinSize = 100;
            splitContainer3.Panel2MinSize = 50;

        }

        private void splitContainer2_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void GioHangView_Load(object sender, EventArgs e)
        {
            CenterToScreen();
            this.Size = new Size(600, 600);
            boxPTThanhToan.SelectedItem = boxPTThanhToan.Items[0];
            DataTable dt = new DataTable();
            string query = "select * from giohang gh join GioSanPham gsp on gh.MaGH = gsp.MaGH\r\njoin SanPham sp on sp.MaSanPham = gsp.MaSanPham \r\nwhere gh.MaNguoiDung = '" + Program.currentUser.MaUser+"'";
            dt = db.DataReader(query);
            if(dt.Rows.Count > 0)
            {
                NhapDanhSachSP(dt);
                HienThiSanPham(tableSanPhams);
            }
            else
            {
                MessageBox.Show("Loi DB", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void NhapDanhSachSP(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                string tensp = row["TenSanPham"].ToString();
                string masp = row["MaSanPham"].ToString();
                string phanloai = row["PhanLoai"].ToString();
                string hangsx = row["HangSX"].ToString();
                string anh = row["Anh"].ToString();
                int soLuong = Convert.ToInt32(row["SoLuong"]);
                decimal giaCa = Convert.ToDecimal(row["GiaCa"]);
                string mota = row["MoTa"].ToString();


                SanPham item = new SanPham(masp, tensp, phanloai, hangsx, soLuong, giaCa, anh, mota);
                Program.danhSachSP.ThemSanPham(item);
            }
        }

        // them sp vao tableSanPhams
        private void HienThiSanPham(TableLayoutPanel tableSanPhams)
        {
            // 1 cot va nhieu hang
            tableSanPhams.ColumnCount = 1;
            tableSanPhams.RowCount = Program.danhSachSP.TongSoSP(); //hang = so sp
            tableSanPhams.AutoSize = true;
            tableSanPhams.Controls.Clear(); // Xóa các sản phẩm cũ nếu có

            // Cài đặt chế độ tự động điều chỉnh chiều cao hàng
            tableSanPhams.RowStyles.Clear();
            for(int i = 0;i< Program.danhSachSP.TongSoSP();i++)
            {
                tableSanPhams.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            // Lặp qua mỗi sản phẩm trong DataTable
            //tong so
            int tongSoLuong = 0;
            decimal tong = 0;
            for (int i = 0; i < Program.danhSachSP.TongSoSP(); i++)
            {
                SanPham item = Program.danhSachSP.GetSanPham(i);
                // Tạo TableLayoutPanel nhỏ cho từng sản phẩm (2 cột, 1 hàng chính)
                TableLayoutPanel productPanel = new TableLayoutPanel
                {
                    ColumnCount = 2,
                    RowCount = 1,
                    Size = new Size(300, 120), // Chỉnh kích thước theo nhu cầu
                    Dock = DockStyle.Top,
                    CellBorderStyle = TableLayoutPanelCellBorderStyle.Single // Tùy chọn, nếu muốn có viền
                };

                // Đặt kích thước cột
                productPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F)); // Ảnh sản phẩm
                productPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));  // Thông tin sản phẩm

                // Tạo PictureBox cho ảnh sản phẩm 
                PictureBox pictureBox = new PictureBox
                {
                    Dock = DockStyle.Fill,
                    SizeMode = PictureBoxSizeMode.StretchImage,
                    Margin = new Padding(3)
                };

                string imagePath = System.Windows.Forms.Application.StartupPath + "\\AnhSP\\" + item.anhSp;
                try
                {
                    pictureBox.Image = Image.FromFile(imagePath);
                }
                catch
                {
                    pictureBox.BackColor = Color.LightGray; // Màu nền nếu ảnh không tồn tại
                }
                productPanel.Controls.Add(pictureBox, 0, 0);

                // Tạo TableLayoutPanel con cho thông tin sản phẩm
                TableLayoutPanel infoPanel = new TableLayoutPanel
                {
                    ColumnCount = 1,
                    RowCount = 3,
                };

                // Cài đặt tỷ lệ kích thước cho từng hàng trong infoPanel
                infoPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F)); // Tên sản phẩm
                infoPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F)); // Giá sản phẩm
                infoPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F)); // Số lượng

                
                // Tạo và thêm Label tên sản phẩm
                Label lblTenSp = new Label
                {
                    Text = item.tenSp,
                    //Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Font = new Font("Arial", 10, FontStyle.Bold)
                };
                infoPanel.Controls.Add(lblTenSp);

                // Tạo và thêm Label giá sản phẩm
                Label lblGiaSp = new Label
                {
                    Text = "Giá: " + item.giaSp + " VND",
                    TextAlign = ContentAlignment.MiddleLeft,
                    Font = new Font("Arial", 9)
                };
                infoPanel.Controls.Add(lblGiaSp);

                // Thay thế Label bằng NumericUpDown cho số lượng
                NumericUpDown numSoLuong = new NumericUpDown
                {
                    Value = item.soLuong, // Giá trị mặc định là số lượng tồn kho
                    Minimum = 0, // Số lượng không thể nhỏ hơn 0
                    //Maximum = Convert.ToDecimal(row["TonKho"]), // Tối đa là số lượng tồn kho
                    Tag = (item.maSp, item.tenSp),
                    TextAlign = HorizontalAlignment.Center,
                    Font = new Font("Arial", 9)
                };
                numSoLuong.Tag = numSoLuong.Value;
                
                numSoLuong.ValueChanged += NumSoLuong_ValueChanged;

                infoPanel.Controls.Add(numSoLuong);

                productPanel.Controls.Add(infoPanel, 1, 0);

                tableSanPhams.Controls.Add(productPanel);

               
            }

            
            txtTongSoSP.Text = Program.danhSachSP.TongSoLuong().ToString();
            txtTongSoMatHang.Text = Program.danhSachSP.TongSoSP().ToString();
            txtTongTien.Text = Program.danhSachSP.TongGiaTri().ToString();

        }

        private void NumSoLuong_ValueChanged(object sender, EventArgs e)
        {
            if (!isUserChanging)
                return;
            NumericUpDown numSoLuong = (NumericUpDown)sender;
            var tag = ((string maSp, string tenSp))numSoLuong.Tag;
           
            decimal oldValue = Convert.ToDecimal(numSoLuong.Tag);
            decimal newValue = numSoLuong.Value;

            // Hiển thị message box để hỏi người dùng
            DialogResult result = MessageBox.Show($"Bạn có muốn thay đổi số lượng hàng thành {newValue} không?",
                                                 "Thay đổi số lượng",
                                                 MessageBoxButtons.YesNo,
                                                 MessageBoxIcon.Question);

            // Xử lý kết quả từ message box
            if (result == DialogResult.Yes)
            {
                // Cập nhật giá trị mới
                if (newValue == 0)
                {
                    //xoa sp
                    DialogResult xoa =  MessageBox.Show("Bạn muốn xoá sản phẩm khỏi giỏ hàng?","Cẩn thận", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (xoa == DialogResult.OK)
                    {
                        XoaSanPhamKhoiGioHang(tag.maSp, tag.tenSp);
                    }else if(xoa == DialogResult.Cancel)
                    {
                        isUserChanging = false; 
                        numSoLuong.Value = oldValue; 
                        isUserChanging = true;
                    }
                }
                else
                {

                numSoLuong.Tag = newValue; // Cập nhật giá trị cũ
                    //thay doi trong csdl
                    
               MessageBox.Show("Đổi thành công","Thông báo",MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                // Giữ nguyên giá trị cũ
                isUserChanging = false; // Tạm tắt sự kiện
                numSoLuong.Value = oldValue; // Gán lại giá trị cũ
            }
        }

        private void XoaSanPhamKhoiGioHang(string masp, string tensp)
        {
           
            Program.danhSachSP.XoaSanPham(masp);


           MessageBox.Show("Đã xoá sản phẩm "+ tensp, "Thông báo",MessageBoxButtons.OK,MessageBoxIcon.Information);
            
        }

        private void splitContainer2_SplitterMoved(object sender, SplitterEventArgs e)
        {

        }

        private void btnQuayLai_Click(object sender, EventArgs e)
        {
            this.Close();
            HomeView mainFrm = new HomeView(Program.currentUser.TenUser);
        }
    }
}
