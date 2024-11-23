using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Models
{
    public class Fullname
    {
        public string FirstName { get; set; }
        public string MidName { get; set; }
        public string LastName { get; set; }

        public Fullname()
        {

        }

        public Fullname(string fullName)
        {
            var fullname = fullName?.Split(' ');
            var mid = "";
            for (int i = 1; i < fullname.Length - 1; i++)
            {
                mid = fullname[i] + " ";
            }
            MidName = mid.TrimEnd();
            FirstName = fullname[0];
            LastName = fullname[fullname.Length - 1];
        }

        public Fullname(string firstName, string midName, string lastName)
        {
            FirstName = firstName;
            MidName = midName;
            LastName = lastName;
        }
        public override string ToString()
        {
            return $"{LastName} {MidName} {FirstName}";
        }
    }
}
