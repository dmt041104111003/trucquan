using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Models
{
    public class Discount
    {
        private static int s_autoId = 1000000;
        public int DiscountID { get; set; }
        public string Name { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public string DiscountType { get; set; }
        public int DiscountPercent { get; set; }
        public int DiscountPriceAmount { get; set; }

        public Discount()
        {
            
        }

        public Discount(int discountID, string name, DateTime startTime, DateTime endTime, string discountType, 
            int discountPercent, int discountPriceAmount)
        {
            DiscountID = discountID;
            Name = name;
            StartTime = startTime;
            EndTime = endTime;
            DiscountType = discountType;
            DiscountPercent = discountPercent;
            DiscountPriceAmount = discountPriceAmount;
        }

        public override bool Equals(object obj)
        {
            return obj is Discount discount &&
                   DiscountID == discount.DiscountID;
        }

        public override int GetHashCode()
        {
            return -1825600901 + DiscountID.GetHashCode();
        }
    }
}
