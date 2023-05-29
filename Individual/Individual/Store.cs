using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Individual
{
    public class Store
    {
        public List<Employee> Employees { get; private set; }
        public List<Report> Reports { get; private set; }
        public List<Payout> Payouts { get; private set; }

        public Store()
        {
            Employees = new List<Employee>();
            Reports = new List<Report>();
            Payouts = new List<Payout>();
        }

        public void deleteEmployee(int id)
        {
            foreach (var employee in Employees)
            {
                if (employee.ID == id)
                {
                    Employees.Remove(employee);
                    break;
                }
            }
            for(int i = 0; i < Reports.Count; i++)
            {
                if (Reports[i].EmpID == id)
                {
                    Reports.RemoveAt(i);
                    i--;
                }
            }
            for (int i = 0; i < Payouts.Count; i++)
            {
                if (Payouts[i].EmpID == id)
                {
                    Payouts.RemoveAt(i);
                    i--;
                }
            }
        }

        public void deleteReport(int id)
        {
            foreach (var rep in Reports)
            {
                if (rep.ID == id)
                {
                    Reports.Remove(rep);
                    break;
                }
            }
        }

        public void deletePayout(int empID, DateTime date)
        {
            foreach (var payout in Payouts)
            {
                if (payout.EmpID == empID && payout.IssueDate == date)
                {
                    Payouts.Remove(payout);
                    break;
                }
            }
        }
    }
}
