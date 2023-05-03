using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace FigureDraw
{
    public partial class FigureForm : Form
    {
        public FigureForm()
        {
            InitializeComponent();
            comboBox1.SelectedIndex = 0;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox1.SelectedIndex == 1)
            {
                label1.Text = "Точка 1 ( x, y )";
                label2.Text = "Точка 2 ( x, y )";
                textBox4.Visible = true;
            }
            else if(comboBox1.SelectedIndex == 0)
            {
                label1.Text = "Центр ( x, y )";
                label2.Text = "Радиус";
                textBox4.Visible = false;
            }

            tabControl1.SelectedIndex = comboBox1.SelectedIndex;
            textBox1.Focus();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                float x = float.Parse(textBox1.Text);
                float y = float.Parse(textBox2.Text);
                if (comboBox1.SelectedIndex == 0)
                {
                    float r = float.Parse(textBox3.Text);

                    drawCircle(x, y, r);

                    dataGridView1.Rows.Add(x, y, r, Math.Round(2 * r * Math.PI, 5), Math.Round(r * r * Math.PI, 5));
                    label3.Text = "Длина : " + Math.Round(2 * r * Math.PI, 5);
                    label4.Text = "Площадь: " + Math.Round(r * r * Math.PI, 5);
                }
                else if (comboBox1.SelectedIndex == 1)
                {
                    float x2 = float.Parse(textBox3.Text);
                    float y2 = float.Parse(textBox4.Text);
                    float h = Math.Abs(y - y2);
                    float w = Math.Abs(x - x2);

                    drawRectangle(x, y, x2, y2);

                    dataGridView2.Rows.Add(x, y, x2, y2, w, h, 2 * (w + h), w * h);
                    label3.Text = "Периметр: " + Math.Round(2 * (w + h), 5);
                    label4.Text = "Площадь: " + Math.Round(w * h, 5);

                }
            }
            catch (FormatException)
            {
                MessageBox.Show(this, "Введите корректные данные", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application();
            Excel.Workbook workbook = exApp.Workbooks.Add();
            saveDataToExcel(workbook);
            saveExcelFile(workbook);
            workbook.Close();
            exApp.Quit();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                Excel.Workbook workbook = exApp.Workbooks.Open(openFileDialog1.FileName);
                Excel.Worksheet wsh = workbook.Worksheets[2];
                Range range = wsh.UsedRange;
                dataGridView1.Rows.Clear();

                for (int i = 2; i <= range.Rows.Count; i++)
                {
                    dataGridView1.Rows.Add();
                    for (int j = 1; j <= range.Columns.Count; j++)
                        dataGridView1[j - 1, i - 2].Value = range.Cells[i, j].Value;
                    drawCircle(float.Parse(dataGridView1[0,i-2].Value.ToString()), 
                        float.Parse(dataGridView1[1, i - 2].Value.ToString()), 
                        float.Parse(dataGridView1[2, i - 2].Value.ToString()));
                }

                wsh = workbook.Worksheets[1];
                range = wsh.UsedRange;
                dataGridView2.Rows.Clear();
                for (int i = 2; i <= range.Rows.Count; i++)
                {
                    dataGridView2.Rows.Add();
                    for (int j = 1; j <= range.Columns.Count; j++)
                        dataGridView2[j - 1, i - 2].Value = range.Cells[i, j].Value;
                    drawRectangle(float.Parse(dataGridView2[0, i - 2].Value.ToString()),
                       float.Parse(dataGridView2[1, i - 2].Value.ToString()),
                       float.Parse(dataGridView2[2, i - 2].Value.ToString()),
                       float.Parse(dataGridView2[3, i - 2].Value.ToString()));
                }
                workbook.Close();
            }
            exApp.Quit();
        }
        private void saveDataToExcel(Excel.Workbook workbook)
        {
            Excel.Worksheet wsh = (Excel.Worksheet)workbook.ActiveSheet;
            wsh.Name = "Круг";

            for (int i = 0; i < dataGridView1.ColumnCount; i++)
                wsh.Cells[1, i+1] = dataGridView1.Columns[i].HeaderText;
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    wsh.Cells[i + 2, j + 1] = dataGridView1[j, i].Value;

            wsh = workbook.Worksheets.Add();
            wsh.Name = "Прямоугольник";

            for (int i = 0; i < dataGridView2.ColumnCount; i++)
                wsh.Cells[1, i + 1] = dataGridView2.Columns[i].HeaderText;
            for (int i = 0; i <= dataGridView2.RowCount - 2; i++)
                for (int j = 0; j <= dataGridView2.ColumnCount - 1; j++)
                    wsh.Cells[i + 2, j + 1] = dataGridView2[j, i].Value;
        }

        private void saveExcelFile(Excel.Workbook workbook)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                workbook.SaveAs(saveFileDialog1.FileName, Excel.XlFileFormat.xlWorkbookDefault,
                    Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
        }

        private void drawRectangle(float x1, float y1, float x2, float y2)
        {
            Graphics g = pictureBox1.CreateGraphics();
            int centerX = pictureBox1.Width / 2;
            int centerY = pictureBox1.Height / 2;
            Pen pen = new Pen(Color.Black, 2);

            g.TranslateTransform(centerX + Math.Min(x1, x2), centerY - Math.Max(y1, y2));
            g.DrawRectangle(pen, 0, 0, Math.Abs(x1- x2), Math.Abs(y1-y2));
            g.TranslateTransform(centerX, centerY);
        }

        private void drawCircle(float x, float y, float r)
        {
            Graphics g = pictureBox1.CreateGraphics();
            int centerX = pictureBox1.Width / 2;
            int centerY = pictureBox1.Height / 2;
            Pen pen = new Pen(Color.Black, 2);

            g.TranslateTransform(centerX + x - r, centerY - r - y);
            g.DrawEllipse(pen, 0, 0, 2 * r, 2 * r);
            g.TranslateTransform(centerX, centerY);
        }


    }
}
