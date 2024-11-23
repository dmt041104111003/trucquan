using CSharpCounterFinalProject.Sign;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WMPLib;

namespace CSharpCounterFinalProject.ViewCustomer
{
    public partial class HomeCustomerView : Form
    {
        public HomeCustomerView(string username)
        {
            InitializeComponent();
            CenterToScreen();
            labelName.Text = username;
        }

        private void btnBuyItem_Click(object sender, EventArgs e)
        {
            var children = new CartCus(labelName.Text);
            children.Show();
            Close();
        }

        private void btnInfor_Click(object sender, EventArgs e)
        {
            var children = new CustomerInformationView(labelName.Text);
            children.Show();
            Close();
        }

        private void btnSignOut_Click(object sender, EventArgs e)
        {
            Close();
            var children = new SignIn();
            children.Show();
        }

        private void HomeCustomerView_Load(object sender, EventArgs e)
        {
            // Tạo danh sách phát
            IWMPPlaylist playlist = WMPlayer.newPlaylist("MyPlaylist", "");


            string[] videoFileNames = new string[]
            {
                "1.mp4",
                "2.mp4",
                "3.mp4"
            };

            string appDirectory = Path.GetDirectoryName(Application.ExecutablePath); // Lấy thư mục gốc của ứng dụng

            string videoDirectory = Path.Combine(appDirectory, "Assets\\Video"); // Tạo đường dẫn đầy đủ đến thư mục chứa video

            // Thêm các video vào danh sách phát
            foreach (string fileName in videoFileNames)
            {
                string videoPath = Path.Combine(videoDirectory, fileName); // Tạo đường dẫn đầy đủ
                IWMPMedia media = WMPlayer.newMedia(videoPath);
                playlist.appendItem(media);
            }

            // Đặt danh sách phát cho Windows Media Player Control
            WMPlayer.currentPlaylist = playlist;

            // Ẩn các nút điều khiển
            WMPlayer.uiMode = "none";

            // Bắt đầu phát danh sách phát
            WMPlayer.Ctlcontrols.play();

            WMPlayer.settings.setMode("loop", true);

        }
    }
}
