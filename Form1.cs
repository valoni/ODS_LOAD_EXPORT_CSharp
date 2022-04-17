using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using WindowsFormsAppODF.Properties;

namespace WindowsFormsAppODF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

         private void Form1_Load(object sender, EventArgs e)
         {
                
         }


        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog theDialog = new OpenFileDialog();
            theDialog.Title = "Open Document Spreadsheet";
            theDialog.Filter = "Open Document Spreadsheet (.ods)|*.ods";
          

            if (theDialog.ShowDialog() == DialogResult.OK)
            {
                this.textBox1.Text = theDialog.FileName.ToString();

                button2_Click(null, null);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            ds = new OdsReaderWriter().ReadOdsFile(this.textBox1.Text);

            dataGridView1.DataSource = ds.Tables[0];
        }

        private DataTable ToDataTable(DataGridView dataGridView)
        {
            var dt = new DataTable();
            foreach (DataGridViewColumn dataGridViewColumn in dataGridView.Columns)
            {
                if (dataGridViewColumn.Visible)
                {
                    dt.Columns.Add();
                }
            }
            var cell = new object[dataGridView.Columns.Count];
            foreach (DataGridViewRow dataGridViewRow in dataGridView.Rows)
            {
                for (int i = 0; i < dataGridViewRow.Cells.Count; i++)
                {
                    cell[i] = dataGridViewRow.Cells[i].Value;
                }
                dt.Rows.Add(cell);
            }
            return dt;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();

            DataTable dt = ToDataTable(this.dataGridView1);
            dt.TableName = "Sheet1";

            ds.Tables.Add(dt);

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "Save Open Document Spreadsheet Files";
            saveFileDialog1.Filter = "Open Document Spreadsheet (.ods)|*.ods";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = saveFileDialog1.FileName;
            }

            if (this.textBox2.Text.Length>3)
            {
                new OdsReaderWriter().WriteOdsFile(ds, this.textBox2.Text);

            }
        }
    }
}
