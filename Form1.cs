using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;


namespace Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)

        {
            try
            {
                OpenFileDialog file = new OpenFileDialog();
                file.Filter = "Excel Dosyası |*.xlsx";
                file.ShowDialog();

                string ExcelYolu = file.FileName;
                string baglantiAdresi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelYolu + ";Extended Properties='Excel 12.0;IMEX=1;'";

                OleDbConnection baglanti = new OleDbConnection(baglantiAdresi);
                OleDbCommand komut = new OleDbCommand("Select * From[" + "welcome" + "$]", baglanti);
                baglanti.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(komut);
                DataTable data = new DataTable();
                da.Fill(data);
                dataGridView1.DataSource = data;
                this.dataGridView1.Sort(dataGridView1.Columns["Tarih"], ListSortDirection.Descending);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
               
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Stream myStream;
            SaveFileDialog file = new SaveFileDialog();
            string filePath = "";
            file.Filter = "txt Dosyaları (*.txt)|*.txt|Tüm Dosyalar(*.*)|*.*";
            file.FilterIndex = 2;
            file.RestoreDirectory = true;
            if (file.ShowDialog() == DialogResult.OK)
            {
                if ((myStream = file.OpenFile()) != null)
                {
                    myStream.Close();
                }
                filePath = file.FileName;

                TextWriter dosya = new StreamWriter(filePath);
                string header = "";
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    if (i == (dataGridView1.Columns.Count - 1))
                        header += dataGridView1.Columns[i].HeaderCell.Value.ToString() + "\n";
                    else
                        header += dataGridView1.Columns[i].HeaderCell.Value.ToString() + ";";
                }

                dosya.Write(header);

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    string rows = "";
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if (j == (dataGridView1.Columns.Count - 1))
                            rows += dataGridView1.Rows[i].Cells[j].Value.ToString() + "\n";
                        else
                            rows += dataGridView1.Rows[i].Cells[j].Value.ToString() + ";";
                    }
                    dosya.Write(rows);
                }
                dosya.Close();
                int txtRowsCount = File.ReadLines(filePath).Count();

                var groupByYears = (dataGridView1.DataSource as DataTable).AsEnumerable().Where(x => x.Field<DateTime?>("Tarih").HasValue).GroupBy(r => r.Field<DateTime?>("Tarih").Value.Year);

                string CountByYear = "";
                foreach (var item in groupByYears)
                {
                    CountByYear += item.Key.ToString() + " Yılındaki Kayıt Sayısı: " + item.Count() + "\n";
                }

                MessageBox.Show(CountByYear +
                    "Satır Sayısı: " + txtRowsCount);

            }
            else
            {
                MessageBox.Show("İptal Edildi");
            }


        }
    }
}
