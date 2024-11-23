using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Models
{
    public class Cart : IComparable<Cart>
    {
        private static int s_autoId = 10000000;
        public int CartId { get; set; }
        public Customer Customer { get; set; }
        public int TotalItem { get; set; }
        List<SelectedItem> SelectedItems { get; set; } = new List<SelectedItem>();

        public Cart()
        {
            
        }
        public Cart(int id)
        {
            CartId = id > 0 ? id : s_autoId++;
        }
        public Cart(int cartId, Customer customer, int totalItem, List<SelectedItem> selectedItems) : this(cartId)
        {
            Customer = customer;
            TotalItem = totalItem;
            SelectedItems = selectedItems;
        }

        public override bool Equals(object obj)
        {
            return obj is Cart cart &&
                   CartId == cart.CartId;
        }

        public override int GetHashCode()
        {
            return -1568810734 + CartId.GetHashCode();
        }

        public int CompareTo(Cart other)
        {
            return CartId - other.CartId;
        }
        public static void UpdateAutoId(int v)
        {
            s_autoId = v;
        }
    }
}
