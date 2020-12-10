using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;


namespace Exel
{
    public partial class Form1 : Form
    {
        DataTable table = new DataTable();
        public Form1()
        {
            InitializeComponent();
            table.Columns.Add("ID", typeof(string));
            table.Columns.Add("NAME", typeof(string));
            table.Columns.Add("AGE", typeof(string));
            table.Columns.Add("MOBAIL", typeof(string));
            table.Columns.Add("EMAIL", typeof(string));

            dataGridView1.DataSource = table;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            saveFileDialog1.InitialDirectory = "C:";
            saveFileDialog1.Title = "Save as Exel File";
            saveFileDialog1.FileName = "";
            saveFileDialog1.Filter = "Exel Files(2019)|*.xlsx";
            if (saveFileDialog1.ShowDialog() !=DialogResult.Cancel)
            {
                Microsoft.Office.Interop.Excel.Application ExelApp = new Microsoft.Office.Interop.Excel.Application();
                ExelApp.Application.Workbooks.Add(Type.Missing);

                ExelApp.Columns.ColumnWidth = 20;
                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    ExelApp.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < dataGridView1.Columns.Count-1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        ExelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }
                ExelApp.ActiveWorkbook.SaveCopyAs(saveFileDialog1.FileName.ToString());
                ExelApp.ActiveWorkbook.Saved = true;
                ExelApp.Quit();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string[] linia = File.ReadAllLines(@"D:\Prejects\Exel\Exel\123.txt");
            string[] znacenie;

            table.Clear();

            for (int i = 0; i < linia.Length; i++)
            {
                znacenie = linia[i].ToString().Split('/');
                string[] stroka = new string[znacenie.Length];

                for (int j = 0; j < znacenie.Length; j++)
                {
                    stroka[j] = znacenie[j].Trim();
                }
                table.Rows.Add(stroka);
            }
        }
    }
}
