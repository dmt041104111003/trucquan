using CSharpCounterFinalProject.Classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSharpCounterFinalProject.ViewNguoiMua
{
    public partial class ChiTietSanPhamView : Form
    {
        private Timer blinkTimer; private bool isTextVisible = true;
        private string fileSachPath;
        private string fileXemThuPath; 
        private bool isFreeBook; 

        DataBaseProcess db = new DataBaseProcess();

        string tensp, gia, fileanh, masp;
        public ChiTietSanPhamView(string label1Text, string label2Text, string anh, string masp2)
        {
            InitializeComponent();
            tensp = label1Text;
            gia = label2Text;
            fileanh = anh;
            masp = masp2;
            labelTenSP.Text = tensp;
            Controls.Add(linklbDownload);
            linklbDownload.Visible = false;
            string imagePath = System.Windows.Forms.Application.StartupPath + "\\AnhSP\\" + fileanh;
            boxAnhSP.Image = Image.FromFile(imagePath);
            boxAnhSP.SizeMode = PictureBoxSizeMode.Zoom;

            string query = @"
            SELECT s.*, 
                   tl.TenTheLoai, 
                   nxb.TenNXB,
                   km.KM,
                   km.ThoigianBatDau,
                   km.ThoiGianKetThuc,
                   CASE 
                        WHEN km.MaSach IS NOT NULL 
                        AND GETDATE() BETWEEN km.ThoigianBatDau AND km.ThoiGianKetThuc 
                        THEN 1 
                        ELSE 0 
                   END AS DangGiamGia
            FROM Sach s
            LEFT JOIN TheLoai tl ON s.MaTheLoai = tl.MaTheLoai
            LEFT JOIN NhaXuatBan nxb ON s.MaNXB = nxb.MaNXB
            LEFT JOIN KhuyenMai km ON s.MaSach = km.MaSach
            WHERE s.MaSach = @MaSach";

            DataTable dt = db.DataReader(query.Replace("@MaSach", $"'{masp}'"));
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("Không tìm thấy sách", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DataRow dataRow = dt.Rows[0];
            bool dangGiamGia = Convert.ToInt32(dataRow["DangGiamGia"]) == 1;
            decimal donGia = Convert.ToDecimal(dataRow["DonGia"]);
            string tacGia = dataRow["TacGia"].ToString();
            string theLoai = dataRow["TenTheLoai"].ToString();
            string nhaXuatBan = dataRow["TenNXB"].ToString();
            string soTrang = dataRow["SoTrang"].ToString();
            decimal giamGia = dataRow["KM"] != DBNull.Value ? Convert.ToDecimal(dataRow["KM"]) : 0;
            if (dataRow["ThoigianBatDau"] != DBNull.Value && dataRow["ThoiGianKetThuc"] != DBNull.Value)
            {
                DateTime thoiGianBatDau = Convert.ToDateTime(dataRow["ThoigianBatDau"]);
                DateTime thoiGianKetThuc = Convert.ToDateTime(dataRow["ThoiGianKetThuc"]);
                DateTime now = DateTime.Now;

                // Kiểm tra nếu thời gian hiện tại nằm trong khoảng thời gian khuyến mãi
                if (now >= thoiGianBatDau && now <= thoiGianKetThuc)
                {

                    decimal giaMoi = donGia - (donGia * giamGia / 100);
                    labelGia.Text = $"{giaMoi:N0} VND";
                    lbGiagoc.Text = $"{donGia:N0} VND";
                    lbGiagoc.Visible = true; 

                    lbEnd.Text = $"Đang giảm giá {giamGia}%!";
                    lbEnd.Visible = true;


                    if (blinkTimer == null)
                    {
                        blinkTimer = new Timer
                        {
                            Interval = 400 
                        };
                        blinkTimer.Tick += (s, e) =>
                        {
                            isTextVisible = !isTextVisible;
                            lbEnd.Visible = isTextVisible;
                        };
                        blinkTimer.Start();
                    }
                }
                else
                {
                    // Không có khuyến mãi hoặc đã hết hạn
                    labelGia.Text = $"{donGia:N0} VND";
                    lbEnd.Visible = false; 
                    lbGiagoc.Visible = false;
                    if (blinkTimer != null)
                    {
                        blinkTimer.Stop();
                        blinkTimer.Dispose();
                        blinkTimer = null;
                    }
                }
            }
            else
            {
                // Trường hợp không có thông tin thời gian khuyến mãi
                labelGia.Text = $"{donGia:N0} VND";
                lbEnd.Visible = false; 
                lbGiagoc.Visible = false;

                if (blinkTimer != null)
                {
                    blinkTimer.Stop();
                    blinkTimer.Dispose();
                    blinkTimer = null;
                }
            }


            labelHangSP.Text = nhaXuatBan;
            labelLoaiSP.Text = theLoai;
            lbSotrang.Text = soTrang;
            lbTacgia.Text = tacGia;

            if (donGia == 0)
            {
                isFreeBook = true;
                if (dataRow["FileSach"] != DBNull.Value)
                {
                    fileSachPath = dataRow["FileSach"].ToString();
                    linklbDownload.Text = "Tải sách PDF miễn phí";
                    linklbDownload.Visible = true;
                }
            }
            else
            {
                isFreeBook = false;
                if (dataRow["FileXemThu"] != DBNull.Value)
                {
                    fileXemThuPath = dataRow["FileXemThu"].ToString();
                    linklbDownload.Text = "Tải bản xem thử";
                    linklbDownload.Visible = true;
                }
            }

            txtMoTaSP.ReadOnly = true;
            txtMoTaSP.TabStop = false;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void LinklbDownload_Click(object sender, MouseEventArgs e)
        {
            string fullPath;
            string baseFileName;

            if (isFreeBook)
            {
                fullPath = System.Windows.Forms.Application.StartupPath + "\\pdf\\" + fileSachPath;
                baseFileName = $"{tensp}_Full";
            }
            else
            {
                fullPath = System.Windows.Forms.Application.StartupPath + "\\pdf\\" + fileXemThuPath;
                baseFileName = $"{tensp}_Preview";
            }

            if (!Directory.Exists("F:"))
            {
                MessageBox.Show("Không tìm thấy ổ đĩa F", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string downloadFolder = "F:";
            if (!Directory.Exists(downloadFolder))
            {
                Directory.CreateDirectory(downloadFolder);
            }

            try
            {
                if (!File.Exists(fullPath))
                {
                    MessageBox.Show("Không tìm thấy file PDF!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string fileName = $"{baseFileName}.pdf";
                string destinationPath = Path.Combine(downloadFolder, fileName);

                if (File.Exists(destinationPath))
                {
                    DialogResult result = MessageBox.Show(
                        $"File {fileName} đã tồn tại trong ổ F. Bạn có muốn tải thêm bản mới không?",
                        "File đã tồn tại",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (result == DialogResult.No)
                    {
                        if (MessageBox.Show("Bạn có muốn mở file hiện có không?", "Mở file",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            Process.Start(destinationPath);
                        }
                        return;
                    }
                    else
                    {
                        int maxNumber = 0;
                        string searchPattern = $"{baseFileName}(*).pdf";
                        foreach (string file in Directory.GetFiles(downloadFolder, $"{baseFileName}*.pdf"))
                        {
                            string filename = Path.GetFileName(file);
                            if (filename == $"{baseFileName}.pdf") continue;

                            int startIndex = filename.LastIndexOf('(');
                            int endIndex = filename.LastIndexOf(')');
                            if (startIndex != -1 && endIndex != -1)
                            {
                                string numberStr = filename.Substring(startIndex + 1, endIndex - startIndex - 1);
                                if (int.TryParse(numberStr, out int number))
                                {
                                    maxNumber = Math.Max(maxNumber, number);
                                }
                            }
                        }

                        fileName = $"{baseFileName} ({maxNumber + 1}).pdf";
                        destinationPath = Path.Combine(downloadFolder, fileName);
                    }
                }

                File.Copy(fullPath, destinationPath, true);
                if (MessageBox.Show($"Đã tải thành công file {fileName}! Bạn có muốn mở file không?", "Thông báo",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    Process.Start(destinationPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi tải file: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ChiTietSanPhamView_Load(object sender, EventArgs e)
        {
            this.Size = new Size(650, 400);
            CenterToScreen();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           MessageBox.Show($"Bạn đã thêm sản phẩm {tensp}  vào giỏ hàng !\n Hãy kiểm tra giỏ hàng nhé !", "Thông báo", MessageBoxButtons.OK);
        }
    }
}
