using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Individual
{
    public partial class Form1 : Form
    {
        private NpgsqlConnection con;
        private string connString = "Host=127.0.0.1;Username=postgres;Password=123;Database=Reports";
        private Store store;
        private List<string> empIds;


        public Form1()
        {
            InitializeComponent();
            con = new NpgsqlConnection(connString);
            con.Open();
            store = new Store();
            empIds = new List<string>();
            LoadEmployees();
            LoadReports();
            LoadPayouts();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var startDate = dateTimePicker1.Value;
            var endDate = dateTimePicker2.Value;

            if (startDate.ToShortDateString().CompareTo(endDate.ToShortDateString()) >= 0)
            {
                label2.Text = "Неверная дата";
                label2.Visible = true;
                return;
            }
            if (startDate.Day != endDate.Day)
            {
                label2.Text = "Период должен быть кратен месяцу!";
                label2.Visible = true;
                return;
            }
            label2.Visible = false;

            Excel.Application exApp = new Excel.Application();
            Excel.Workbook workbook = exApp.Workbooks.Add();
            Excel.Worksheet wsh = (Excel.Worksheet)workbook.ActiveSheet;
            int cellIndex = 1;
            int repIndex = 0;
            int payIndex = 0;
            foreach (var emp in store.Employees)
            {
                double total = 0;
                wsh.Cells[cellIndex, 1] = emp.FullName;
                wsh.Cells[cellIndex + 1, 1] = "Дата платежа";
                wsh.Cells[cellIndex + 1, 2] = "Статья расхода";
                wsh.Cells[cellIndex + 1, 3] = "Сумма";
                cellIndex+=2;
                for(; repIndex < store.Reports.Count; repIndex++ )
                {
                    if (store.Reports[repIndex].EmpID != emp.ID)
                        break;
                    if (store.Reports[repIndex].ExpenceDate <= endDate && store.Reports[repIndex].ExpenceDate >= startDate)
                    {
                        wsh.Cells[cellIndex, 1] = store.Reports[repIndex].ExpenceDate.ToShortDateString();
                        wsh.Cells[cellIndex, 2] = store.Reports[repIndex].ExpenceItem;
                        wsh.Cells[cellIndex, 3] = store.Reports[repIndex].Total;
                        total += store.Reports[repIndex].Total;
                        cellIndex++;
                    }
                }
                double deposits = 0;
                for (; payIndex < store.Payouts.Count; payIndex++)
                {
                    if (store.Payouts[payIndex].EmpID != emp.ID)
                        break;
                    if (store.Payouts[payIndex].IssueDate <= endDate && store.Payouts[payIndex].IssueDate >= startDate)
                        deposits += store.Payouts[payIndex].Total;
                }
                wsh.Cells[cellIndex, 1] = $"Остаток по счёту = {deposits - total}";
                cellIndex+=2;
            }
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                workbook.SaveAs(saveFileDialog1.FileName, Excel.XlFileFormat.xlWorkbookDefault,
                    Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            workbook.Close();
            exApp.Quit();
        }

        private void LoadEmployees()
        {
            string sql = "SELECT * FROM employee ORDER BY emp_id";
            DataTable dt = new DataTable();
            NpgsqlDataAdapter adap = new NpgsqlDataAdapter(sql, con);
            adap.Fill(dt);
            dataGridView1.DataSource = dt;
            foreach (DataRow row in dt.Rows)
            {
                string fullName = row[2].ToString() + " " + row[1].ToString() + " " + row[3].ToString();
                comboBox1.Items.Add(row[0] + " " + fullName);
                Employee emp = new Employee((int)row[0], fullName, row[4].ToString());
                store.Employees.Add(emp);
            }
            adap.Dispose();
        }

        private void LoadReports()
        {
            string sql = "SELECT * FROM advance_reports ORDER BY emp_id, expence_date";
            DataTable dt = new DataTable();
            NpgsqlDataAdapter adap = new NpgsqlDataAdapter(sql, con);
            adap.Fill(dt);
            dataGridView3.DataSource = dt;
            foreach (DataRow row in dt.Rows)
            {
                Report rep = new Report((int)row[0], (int)row[1], (DateTime)row[2], row[3].ToString(), double.Parse(row[4].ToString()));
                store.Reports.Add(rep);
            }
            adap.Dispose();
        }

        private void LoadPayouts()
        {
            string sql = "SELECT * FROM advance_ammounts ORDER BY emp_id, issue_date";
            DataTable dt = new DataTable();
            NpgsqlDataAdapter adap = new NpgsqlDataAdapter(sql, con);
            adap.Fill(dt);
            dataGridView2.DataSource = dt;
            foreach (DataRow row in dt.Rows)
            {
                Payout payout = new Payout((int)row[0], (DateTime)row[1], double.Parse(row[2].ToString()));
                store.Payouts.Add(payout);
            }
            adap.Dispose();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1)
                return;
            string id = comboBox1.SelectedItem.ToString().Split()[0];
            if (!empIds.Contains(id))
            {
                empIds.Add(id);
                Button button = new Button();
                button.Text = id;
                button.Click += (sender1, e1) =>
                {
                    flowLayoutPanel1.Controls.Remove(button);
                    empIds.Remove(button.Text);
                };
                button.Height = 30;
                flowLayoutPanel1.Controls.Add(button);
            }            
            comboBox1.SelectedIndex = -1;
            button2.Focus();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var startDate = dateTimePicker3.Value;
            var endDate = dateTimePicker4.Value;
            var startMounthDate = startDate.AddDays(-startDate.Day + 1);

            if (startDate.ToShortDateString().CompareTo(endDate.ToShortDateString()) >= 0)
            {
                label5.Text = "Неверная дата";
                label5.Visible = true;
                return;
            }

            if(empIds.Count == 0)
            {
                label5.Text = "Выберите сотрудников";
                label5.Visible = true;
                return;
            }

            label5.Visible = false;

            Excel.Application exApp = new Excel.Application();
            Excel.Workbook workbook = exApp.Workbooks.Add();
            Excel.Worksheet wsh = (Excel.Worksheet)workbook.ActiveSheet;
            int repIndex = 0;
            int payIndex = 0;
            double total = 0;
            double deposits = 0;
            foreach (var emp in store.Employees)
            {
                if (!empIds.Contains(emp.ID.ToString()))
                {
                    continue;
                }
                for (; payIndex < store.Payouts.Count; payIndex++)
                {
                    if (store.Payouts[payIndex].EmpID != emp.ID)
                        break;
                    if (store.Payouts[payIndex].IssueDate >= startMounthDate && store.Payouts[payIndex].IssueDate <= endDate)
                        deposits += store.Payouts[payIndex].Total;
                }
                for (; repIndex < store.Reports.Count; repIndex++)
                {
                    if (store.Reports[repIndex].EmpID != emp.ID)
                        break;
                    if (store.Reports[repIndex].ExpenceDate <= endDate && store.Reports[repIndex].ExpenceDate >= startDate)
                    {
                        total += store.Reports[repIndex].Total;
                    }
                }
            }

            wsh.Cells[1, 1] = $"Общая сумма денег без отчёта для работников = {deposits - total}";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                workbook.SaveAs(saveFileDialog1.FileName, Excel.XlFileFormat.xlWorkbookDefault,
                    Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            workbook.Close();
            exApp.Quit();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            con?.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DeleteForm delete = new DeleteForm(tabControl1.SelectedTab.Text, (DataGridView)tabControl1.SelectedTab.Controls[0], store);
            delete.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            UpdateForm update = new UpdateForm(tabControl1.SelectedTab.Text, (DataGridView)tabControl1.SelectedTab.Controls[0], store);
            update.ShowDialog();
        }
    }
}
