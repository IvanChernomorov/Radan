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
    public partial class UpdateForm : Form
    {
        private DataGridViewCell cell;
        private Store store;
        private string schema;

        public UpdateForm()
        {
            InitializeComponent();
        }

        public UpdateForm(string title, DataGridView dataGridView, Store store) : this()
        {
            schema = title;
            label1.Text = title;
            dataGridView1.DataSource = dataGridView.DataSource;
            this.store = store;
        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            var source = sender as DataGridView;
            cell = source.SelectedCells[0];
            textBox1.Text = cell.Value.ToString();
            textBox2.Text = "";
            button2.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (cell != null)
            {
                string connString = "Host=127.0.0.1;Username=postgres;Password=123;Database=Reports";
                NpgsqlConnection con = new NpgsqlConnection(connString);
                con.Open();
                string columnName = dataGridView1.Columns[cell.ColumnIndex].HeaderText;
                string value = textBox2.Text;
                var id = (int)dataGridView1[0, cell.RowIndex].Value;
                string sql = "";
                if (schema == "Работники")
                {
                    sql = $"UPDATE employee SET {columnName} = '@value'" +
                        $"WHERE emp_id = @id";
                    NpgsqlCommand com = new NpgsqlCommand(sql, con);
                    com.Parameters.AddWithValue("@id", id);
                    com.ExecuteNonQuery();
                }
                else if (schema == "Траты")
                {
                    sql = $"UPDATE advance_reports SET {columnName} = @value  " +
                        $"WHERE rep_id = @id";
                    NpgsqlCommand com = new NpgsqlCommand(sql, con);
                    com.Parameters.AddWithValue("@value", value);
                    com.Parameters.AddWithValue("@id", id);
                    com.ExecuteNonQuery();
                }
                else if (schema == "Выдачи")
                {
                    DateTime date = (DateTime)dataGridView1[1, cell.RowIndex].Value;

                    sql = $"UPDATE advance_ammounts SET {columnName} = @value" +
                       $"WHERE emp_id = @id and issue_date = @date";
                    NpgsqlCommand com = new NpgsqlCommand(sql, con);
                    com.Parameters.AddWithValue("@id", id);
                    com.Parameters.AddWithValue("@date", date);
                    com.ExecuteNonQuery();
                }
                cell.Value = value;
                button2.Enabled = false;
                con.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
