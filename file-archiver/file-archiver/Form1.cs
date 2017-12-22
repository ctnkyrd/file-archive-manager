using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Threading;

namespace file_archiver
{
    public partial class Form1 : Form
    {
#pragma warning disable CS0618 // Type or member is obsolete
        public static string mainPath = ConfigurationSettings.AppSettings.Get("folderPath");
#pragma warning restore CS0618 // Type or member is obsolete
        string[] kurulNames = Directory.GetDirectories(mainPath);

        public Form1()
        {
            InitializeComponent();

            
            fillComboKurul();

            if (comboBox1.Items.Count > 0)
            {
                comboBox1.SelectedIndex = 0;
            }
            else
            {
                MessageBox.Show("Kurul Listesi Alınamadı! Lütfen Dosya Yolunu Düzeltiniz!");
            }
            changeColorComboKurul(comboBox1);
            changeColorTextboxExcel(textBox2);
            changeColorTextboxExcel(textBox1);


            groupBox1.Text = mainPath;
        }

        private void StartForm()
        {
            System.Windows.Forms.Application.Run(new frmSplash()); 
        }

         
        public void ililceGrid()
        {
            Thread t = new Thread(new ThreadStart(StartForm));

            dataGridView1.DataSource = null;
            label4.Text = "Yükleniyor...";
            label4.Visible = true;
            var excelName = textBox1.Text;
            var application = new Microsoft.Office.Interop.Excel.Application();
            var workbook = application.Workbooks.Open(excelName);
            var worksheet_1 = workbook.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;


            int rowCount = worksheet_1.UsedRange.Rows.Count;
            int columnCount = worksheet_1.UsedRange.Columns.Count;

            if (columnCount == 5)
            {
                t.Start();
                System.Data.DataTable ilIlce = new System.Data.DataTable("ilIlce");

                ilIlce.Columns.Add("ilce_id");
                ilIlce.Columns.Add("ilce_adi");
                ilIlce.Columns.Add("il_id");
                ilIlce.Columns.Add("il_adi");
                ilIlce.Columns.Add("dosya_desimal_kodu");

                DataRow row;

                int index = 0;
                object rowIndex = 2;
                while (((Microsoft.Office.Interop.Excel.Range)worksheet_1.Cells[rowIndex, 1]).Value2 != null)
                {
                    rowIndex = 2 + index;
                    row = ilIlce.NewRow();

                    row[0] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)worksheet_1.Cells[rowIndex, 1]).Value2);
                    row[1] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)worksheet_1.Cells[rowIndex, 2]).Value2);
                    row[2] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)worksheet_1.Cells[rowIndex, 3]).Value2);
                    row[3] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)worksheet_1.Cells[rowIndex, 4]).Value2);
                    row[4] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)worksheet_1.Cells[rowIndex, 5]).Value2);
                    index++;
                    ilIlce.Rows.Add(row);

                }
                application.Workbooks.Close();

                dataGridView1.DataSource = ilIlce;
                label4.Visible = false;
                t.Abort();
            }

            

            else
            {
                label4.Text = "Excel Dosyası Hatalı!";
            }
        }

        public void logging(string logText)
        {
            if (richTextBox1.Text.Length==0)
            {
                richTextBox1.Text = DateTime.Now.ToString("[dd-mm-yyy HH:mm:ss]") + "-" + logText;
            }
            else
            {
                richTextBox1.Text = richTextBox1.Text + Environment.NewLine + DateTime.Now.ToString("[dd-mm-yyy HH:mm:ss]") + "-" + logText;
            }
            
        }


        void fillComboKurul()
        {
            foreach (var item in kurulNames)
            {
                comboBox1.Items.Add(item.ToString().Split('\\')[item.ToString().Split('\\').Length-1]);
            }
        }

        public void changeColorComboKurul(ComboBox cb)
        {
            if (cb.SelectedItem.ToString().Length>0)
            {
                cb.BackColor = Color.LightGreen;
                cb.ForeColor = Color.DarkGreen;
            }
            else
            {
                cb.BackColor = Color.LightPink;
                cb.ForeColor = Color.DarkRed;
            }
        }

        public void changeColorTextboxExcel (System.Windows.Forms.TextBox tb)
        {
            string fileExtension = tb.Text.Split('.')[tb.Text.Split('.').Length - 1];

            if (fileExtension == "xls" || fileExtension == "xlsx")
            {
                tb.BackColor = Color.LightGreen;
                tb.ForeColor = Color.DarkGreen;
            }
            else
            {
                tb.BackColor = Color.LightPink;
                tb.ForeColor = Color.DarkRed;
            }
        }

        void getExcelSheets(string excelFile)
        {
            listBox1.Items.Clear();
            var application = new Microsoft.Office.Interop.Excel.Application();
            var workbook = application.Workbooks.Open(excelFile);
            var worksheet_1 = workbook.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
            var worksheet_2 = workbook.Worksheets[2] as Microsoft.Office.Interop.Excel.Worksheet;

            var sheetName_1 = worksheet_1.Name.ToString();
            var sheetName_2 = worksheet_2.Name.ToString();

            listBox1.Items.Add(sheetName_1);
            listBox1.Items.Add(sheetName_2);

            int rowCountTiff = worksheet_1.UsedRange.Rows.Count;
            int rowCountPdf = worksheet_2.UsedRange.Rows.Count;
            int columnCountTiff = worksheet_1.UsedRange.Columns.Count;
            int columnCountPdf = worksheet_2.UsedRange.Columns.Count;


            //if (columnCountTiff == 5)
            //{
            //    System.Data.DataTable ilIlce = new System.Data.DataTable("ilIlce");

            //    ilIlce.Columns.Add("ilce_id");
            //    ilIlce.Columns.Add("ilce_adi");
            //    ilIlce.Columns.Add("il_id");
            //    ilIlce.Columns.Add("il_adi");
            //    ilIlce.Columns.Add("dosya_desimal_kodu");

            //    DataRow row;

            //    int index = 0;
            //    object rowIndex = 2;
            //    while (((Microsoft.Office.Interop.Excel.Range)worksheet_1.Cells[rowIndex, 1]).Value2 != null)
            //    {
            //        rowIndex = 2 + index;
            //        row = ilIlce.NewRow();

            //        row[0] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)worksheet_1.Cells[rowIndex, 1]).Value2);
            //        row[1] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)worksheet_1.Cells[rowIndex, 2]).Value2);
            //        row[2] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)worksheet_1.Cells[rowIndex, 3]).Value2);
            //        row[3] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)worksheet_1.Cells[rowIndex, 4]).Value2);
            //        row[4] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)worksheet_1.Cells[rowIndex, 5]).Value2);
            //        index++;
            //        ilIlce.Rows.Add(row);

            //    }
            //    application.Workbooks.Close();

            //    dataGridView1.DataSource = ilIlce;
            //    label4.Visible = false;
            //}

            //else
            //{
            //    label4.Text = "Excel Dosyası Hatalı!";
            //}

        }


        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.BackColor == Color.LightGreen)
            {
                string path = mainPath + comboBox1.SelectedItem.ToString() + "\\";
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                if (Directory.Exists(path))
                {
                    openFileDialog1.InitialDirectory = path;
                }
                else
                {
                    openFileDialog1.InitialDirectory = mainPath;
                }
                openFileDialog1.Filter = "Excel Files |*.xls;*.xlsx";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    textBox2.Text = openFileDialog1.FileName;
                    changeColorTextboxExcel(textBox2);
                    getExcelSheets(textBox2.Text);
                }
            }
            else
            {
                MessageBox.Show("Önce İl-İlçe Kod Excelini Seçiniz");
            }
           
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            changeColorTextboxExcel(textBox2);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            changeColorComboKurul(comboBox1);
            logging(comboBox1.SelectedItem.ToString() + " Kurulu Seçildi!");

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = mainPath + comboBox1.SelectedItem.ToString() + "\\";
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            if (Directory.Exists(path))
            {
                openFileDialog1.InitialDirectory = path;
            }
            else
            {
                openFileDialog1.InitialDirectory = mainPath;
            }
            openFileDialog1.Filter = "Excel Files |*.xls;*.xlsx";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
                changeColorTextboxExcel(textBox1);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            changeColorTextboxExcel(textBox1);
            ililceGrid();
        }

    }
}