using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Individual
{
    internal class Employee
    {
        public int ID { get; set; }
        public string FullName { get; set; }
        public string Post { get; set; }

        public Employee(int ID, string fullName, string post)
        {
            this.ID = ID;
            this.FullName = fullName;
            this.Post = post;
        }

    }
}
