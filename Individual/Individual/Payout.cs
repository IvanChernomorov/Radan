using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Individual
{
    internal class Payout
    {
        public int EmpID { get; set; }
        public DateTime IssueDate { get; set; }
        public double Total { get; set; }

        public Payout(int empID, DateTime issueDate, double total)
        {
            EmpID = empID;
            IssueDate = issueDate;
            Total = total;
        }
    }
}
