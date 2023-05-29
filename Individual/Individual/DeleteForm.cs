using Microsoft.Office.Interop.Excel;
using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Individual
{
    public partial class DeleteForm : Form
    {
        private Store store;
        private DataGridViewRow row;
        private string schema;
        public DeleteForm()
        {
            InitializeComponent();
        }

        public DeleteForm(string title, DataGridView dataGridView, Store store) : this()
        {
            schema = title;
            this.store = store;
            label1.Text = title;
            dataGridView1.DataSource = dataGridView.DataSource;
        }

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            row = dataGridView1.CurrentRow;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (row != null)
            {
                string connString = "Host=127.0.0.1;Username=postgres;Password=123;Database=Reports";
                NpgsqlConnection con = new NpgsqlConnection(connString);
                con.Open();
                int id = (int)dataGridView1[0, row.Index].Value;
                string sql = "";
                if (schema == "Работники")
                {
                    sql = $"DELETE FROM employee " +
                        $"WHERE emp_id = @id";
                    NpgsqlCommand com = new NpgsqlCommand(sql, con);
                    com.Parameters.AddWithValue("@id", id);
                    com.ExecuteNonQuery();
                    store.deleteEmployee(id);
                } 
                else if(schema == "Траты")
                {
                    sql = $"DELETE FROM advance_reports " +
                        $"WHERE rep_id = @id";
                    NpgsqlCommand com = new NpgsqlCommand(sql, con);
                    com.Parameters.AddWithValue("@id", id);
                    com.ExecuteNonQuery();
                    store.deleteEmployee(id);
                }
                else if(schema == "Выдачи")
                {
                    string date = dataGridView1[1, row.Index].Value.ToString();
                    date = date.Split('.')[1] + "." + date.Split('.')[0] + "." + date.Split('.')[2];
                    sql = $"DELETE FROM advance_ammounts " +
                       $"WHERE emp_id = @id and issue_date = '@date'";
                    NpgsqlCommand com = new NpgsqlCommand(sql, con);
                    com.Parameters.AddWithValue("@id", id);
                    com.Parameters.AddWithValue("@date", date);
                    com.ExecuteNonQuery();
                    store.deletePayout(id, (DateTime)dataGridView1[1, row.Index].Value);
                }    
                dataGridView1.Rows.Remove(row);
                con.Close();
            }
        }
    }
}
