using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Models
{
    public class BillDetail : Bill
    {
        public string PaymentMethod { get; set; }
        public string StaffName { get; set; }

        public BillDetail()
        {

        }
        public BillDetail(string paymentMethod, string staffName)
        {
            PaymentMethod = paymentMethod;
            StaffName = staffName;
        }
        public BillDetail(int billId, Cart cart, DateTime createTime, int totalItem, long subTotal, long totalDiscountAmount, long totalAmount, string status,
            string paymentMethod, string staffName) : base(billId, cart, createTime, totalItem, subTotal, totalDiscountAmount, totalAmount, status)
        {
            PaymentMethod = paymentMethod;
            StaffName = staffName;
        }
    }
}
