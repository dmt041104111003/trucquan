using CSharpCounterFinalProject.Classes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSharpCounterFinalProject
{
    internal static class Program
    {

        public static UserInfo currentUser = new UserInfo();
        public static DanhSachSanPham danhSachSP = new DanhSachSanPham();
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Sign.SignIn());

        }
    }
}
