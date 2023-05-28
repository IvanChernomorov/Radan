using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Individual
{
    internal class Report
    {
        public int ID { get; set; }
        public int EmpID { get; set; }
        public DateTime ExpenceDate { get; set; }
        public string ExpenceItem { get; set; }
        public double Total { get; set; }

        public Report(int iD, int empID, DateTime expenceDate, string expenceItem, double total)
        {
            ID = iD;
            EmpID = empID;
            ExpenceDate = expenceDate;
            ExpenceItem = expenceItem;
            Total = total;
        }
    }
}
