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
using System.Reflection;
using System.Data.SqlClient;

namespace file_archiver
{
    public partial class Form1 : Form
    {
#pragma warning disable CS0618 // Type or member is obsolete
        public static string mainPath = ConfigurationSettings.AppSettings.Get("folderPath");
        public static string dosya_arsiv_path = ConfigurationSettings.AppSettings.Get("dosya_arsiv_path");
#pragma warning restore CS0618 // Type or member is obsolete
        string[] kurulNames = Directory.GetDirectories(mainPath);
        public static string dosyalarPath;
        //The DataTables
        System.Data.DataTable tiffFilesDT = new System.Data.DataTable("tiffFilesDT");
        System.Data.DataTable pdfFilesDT = new System.Data.DataTable("pdfFilesDT");
        System.Data.DataTable ilIlce = new System.Data.DataTable("ilIlce");
        System.Data.DataTable dosyaArsiv = new System.Data.DataTable("dosyaArsiv");
        //db parameters

        public static string dbHost, dbUser, dbPass, dbName;

        public Form1()
        {
            InitializeComponent();

            createDosyaArsivUnionTableColumns();
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
            changeColorTextboxExcel(textBox3);
            if (Directory.Exists(textBox3.Text))
            {
                textBox3.BackColor = Color.LightGreen;
                textBox3.ForeColor = Color.DarkGreen;
                dosyalarPath = textBox3.Text;

            }
            else
            {
                textBox3.BackColor = Color.LightPink;
                textBox3.ForeColor = Color.DarkRed;
            }

            groupBox1.Text = mainPath;
        }

        void fillDosyaArsivUnionTableColumns(int idKayit, string desimalNo, string Aciklama, string onayNo, string onayTarihi, string turu, string barkod, int ilId, int ilceID)
        {
            DataRow row = dosyaArsiv.NewRow();
            row[0] = idKayit;
            row[1] = desimalNo;
            row[2] = Aciklama;
            row[3] = onayNo;
            row[4] = onayTarihi;
            row[5] = turu;
            row[6] = barkod;
            row[7] = ilId;
            row[8] = ilceID;
            dosyaArsiv.Rows.Add(row);
        }

        void createDosyaArsivUnionTableColumns()
        {
            dosyaArsiv.Columns.Add("id_kayit");
            dosyaArsiv.Columns.Add("desimal_no");
            dosyaArsiv.Columns.Add("aciklama");
            dosyaArsiv.Columns.Add("onay_no");
            dosyaArsiv.Columns.Add("onay_tarihi");
            dosyaArsiv.Columns.Add("turu");
            dosyaArsiv.Columns.Add("barkod_no");
            dosyaArsiv.Columns.Add("il_id");
            dosyaArsiv.Columns.Add("ilce_id");
            dosyaArsiv.Columns.Add("exists");
            dosyaArsiv.Columns.Add("dosya_yolu");

        }

        private void mergeTables()
        {
            int idKayit, ilId, ilceId;
            string desimalNo, Aciklama, onayNo, turu, barkod,desimalSelection,onayTarihi;

            foreach (DataRow dr in tiffFilesDT.Rows)
            {
                if (dr["ID_kayıt_no"] != DBNull.Value)
                {
                    idKayit = Convert.ToInt32(dr["ID_kayıt_no"]);
                    desimalNo = dr["desimal_no"].ToString();
                    desimalSelection = desimalNo.Split('.')[0] + "." + desimalNo.Split('.')[1];

                    //Select relational values from ililcekod excel
                    ilId = Convert.ToInt32(ilIlce.Select("dosya_desimal_kodu = '" + desimalSelection + "'")[0]["il_id"]);
                    ilceId = Convert.ToInt32(ilIlce.Select("dosya_desimal_kodu = '" + desimalSelection + "'")[0]["ilce_id"]);
                    if (dr["proje_açıklaması"] != DBNull.Value)
                    {
                        Aciklama = dr["proje_açıklaması"].ToString();
                    }
                    else
                    {
                        Aciklama = null;
                    }
                    if (dr["onay_no"] != DBNull.Value)
                    {
                        onayNo = dr["onay_no"].ToString();
                    }
                    else
                    {
                        onayNo = null;
                    }
                    turu = ".tif";
                    barkod = null;
                    if (dr["onay_tarihi"] != DBNull.Value)
                    {
                        onayTarihi = DateTime.FromOADate(Convert.ToDouble(dr["onay_tarihi"])).ToString("yyyy-MM-dd HH:mm:ss.fff");
                    }
                    else
                    {
                        onayTarihi = null;
                    }
                    fillDosyaArsivUnionTableColumns(idKayit, desimalNo, Aciklama, onayNo, onayTarihi, turu, barkod, ilId,ilceId);
                }
                dataGridView4.DataSource = dosyaArsiv;
                tabControl1.SelectedIndex = 5;
                
            }
            foreach (DataRow dr in pdfFilesDT.Rows)
            {
                if (dr["id_kayit_no"] != DBNull.Value)
                {
                    idKayit = Convert.ToInt32(dr["id_kayit_no"]);
                    desimalNo = dr["desimal_no"].ToString();
                    desimalSelection = desimalNo.Split('.')[0] + "." + desimalNo.Split('.')[1];

                    //Select relational values from ililcekod excel
                    ilId = Convert.ToInt32(ilIlce.Select("dosya_desimal_kodu = '" + desimalSelection + "'")[0]["il_id"]);
                    ilceId = Convert.ToInt32(ilIlce.Select("dosya_desimal_kodu = '" + desimalSelection + "'")[0]["ilce_id"]);
                    Aciklama = null;
                    onayNo = null;
                    turu = ".pdf";
                    if (dr["barkod_no"] != DBNull.Value)
                    {
                        barkod = dr["barkod_no"].ToString();
                    }
                    else
                    {
                        barkod = null;
                    }
                    onayTarihi = null;
                    fillDosyaArsivUnionTableColumns(idKayit, desimalNo, Aciklama, onayNo, onayTarihi, turu, barkod,ilId,ilceId);
                }
                
            }
        }

        private void updateDosyaArsivDT(string dosyaDesimal, string onayNumarasi, DateTime? onayTarihi)
        {   dataGridView4.DataSource = null;
            foreach (DataRow dr in dosyaArsiv.Select("desimal_no = '"+dosyaDesimal+"' and onay_no = '"+onayNumarasi+"' and onay_tarihi = '"+onayTarihi+"'"))
            {
                dr["exists"] = 1;
            }
            dataGridView4.DataSource = dosyaArsiv;
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

            logging(excelName + "-il-ilce Datatable Aktarım Başladı");

            int rowCount = worksheet_1.UsedRange.Rows.Count;
            int columnCount = worksheet_1.UsedRange.Columns.Count;

            if (columnCount == 5)
            {
                t.Start();


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
                //application.Workbooks.Close();
                workbook.Close(false, Missing.Value, Missing.Value);
                application.Quit();

                dataGridView1.DataSource = ilIlce;
                label4.Visible = false;
                t.Abort();
                logging(excelName + "-il-ilce Datatable Aktarım Tamamlandı");

            }



            else
            {
                label4.Text = "Excel Dosyası Hatalı!";
            }

        }

        public void logging(string logText)
        {
            if (richTextBox1.Text.Length == 0)
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
                comboBox1.Items.Add(item.ToString().Split('\\')[item.ToString().Split('\\').Length - 1]);
            }
        }

        private void dbTextbox_TextChanged(object sender, EventArgs e)
        {
            dbHost = tbHost.Text;
            dbUser = tbUser.Text;
            dbPass = tbPass.Text;
            System.Windows.Forms.TextBox txt = (System.Windows.Forms.TextBox)sender;
            getAllDatabeses(dbHost,dbUser,dbPass);
        }

        void getAllDatabeses(string serverName, string user, string passwd)
        {
            comboBox2.Items.Clear();
            string conString = "server=" + serverName + ";uid=" + user + ";pwd=" + passwd;
            try
            {
                using (SqlConnection con = new SqlConnection(conString))
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand("SELECT name from sys.databases", con))
                    {
                        label14.Text = "Connected!";
                        label14.ForeColor = Color.LightGreen;
                        using (IDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                string dbName = dr[0].ToString();
                                if (!comboBox2.Items.Contains(dbName))
                                {
                                    comboBox2.Items.Add(dbName);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                label14.Text = "Disconnected!";
                label14.ForeColor = Color.Red;
                //logging(ex.ToString());
            }
        }

        private bool isExists(string dosyaDesimal, string onayNo, string onayTarihi)
        {
            try
            {
                string conString = "server=" + dbHost + ";uid=" + dbUser+ ";pwd=" + dbPass+ ";database="+dbName;
                using (SqlConnection con = new SqlConnection(conString))
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand("SELECT * FROM DOSYA_ARSIV WHERE DOSYA_DESIMAL = '"+dosyaDesimal+ "' AND PROJE_ONAY_SAYISI = '" + onayNo + "' AND PROJE_ONAY_TARIHI ='"+ onayTarihi+"'", con))
                    {                       
                        using (IDataReader dr = cmd.ExecuteReader())
                        {
                            if (dr.Read())
                            {
                                return true;
                            }
                            else
                            {
                                return false;
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw;
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
            try
            {
                logging(excelFile + "-Veri DataTable Aktarım Başladı");
                Thread t = new Thread(new ThreadStart(StartForm));


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

                t.Start();

                tiffFilesDT.Columns.Add("ID_kayıt_no");
                tiffFilesDT.Columns.Add("desimal_no");
                tiffFilesDT.Columns.Add("proje_açıklaması");
                tiffFilesDT.Columns.Add("onay_no");
                tiffFilesDT.Columns.Add("onay_tarihi");

                pdfFilesDT.Columns.Add("id_kayit_no");
                pdfFilesDT.Columns.Add("barkod_no");
                pdfFilesDT.Columns.Add("desimal_no");

                DataRow row;

                int index = 0;
                object rowIndex = 2;
                while (((Range)worksheet_1.Cells[rowIndex, 1]).Value2 != null)
                {
                    rowIndex = 2 + index;
                    row = tiffFilesDT.NewRow();

                    if (Convert.ToInt32(((Range)worksheet_1.Cells[rowIndex, 1]).Value2)>=1)
                    {
                        row[0] = Convert.ToString(((Range)worksheet_1.Cells[rowIndex, 1]).Value2);
                        row[1] = Convert.ToString(((Range)worksheet_1.Cells[rowIndex, 2]).Value2);
                        row[2] = Convert.ToString(((Range)worksheet_1.Cells[rowIndex, 3]).Value2);
                        row[3] = Convert.ToString(((Range)worksheet_1.Cells[rowIndex, 4]).Value2);
                        row[4] = Convert.ToString(((Range)worksheet_1.Cells[rowIndex, 5]).Value2);
                        index++;
                        tiffFilesDT.Rows.Add(row);
                    }
                    else
                    {
                        Console.WriteLine("dosya_sonu");
                    }
                }
                logging(excelFile + "-tiff Dosyalar Listelenmesi Tamamlandı");
                index = 0;
                rowIndex = 2;
                while (((Range)worksheet_2.Cells[rowIndex, 1]).Value2 != null)
                {
                    rowIndex = 2 + index;
                    row = pdfFilesDT.NewRow();
                    if (Convert.ToInt32(((Range)worksheet_2.Cells[rowIndex, 1]).Value2>=1))
                    {
                        row[0] = Convert.ToString(((Range)worksheet_2.Cells[rowIndex, 1]).Value2);
                        row[1] = Convert.ToString(((Range)worksheet_2.Cells[rowIndex, 2]).Value2);
                        row[2] = Convert.ToString(((Range)worksheet_2.Cells[rowIndex, 3]).Value2);
                        index++;
                        pdfFilesDT.Rows.Add(row);
                    }
                    else
                    {
                        Console.WriteLine("dosya_sonu");
                    }

                }
                logging(excelFile + "-pdf Dosyalar Listelenmesi Tamamlandı");

                workbook.Close(false, Missing.Value, Missing.Value);
                application.Quit();
                FormCollection fc = System.Windows.Forms.Application.OpenForms;
                t.Abort();
                logging(excelFile + "-Listelenme Tamamlandı");

            }
            catch (Exception ex)
            {

                logging(excelFile+"-"+ex.ToString());
            }
            
        }

        private string dosyaArsiv_path (string dosyaKodu, string extension)
        {
            try
            {
                string fullFileName = dosyaKodu + extension;
                string[] allFiles = Directory.GetFiles(dosyalarPath, fullFileName, SearchOption.AllDirectories);
                if (allFiles.Length>0)
                {
                    return allFiles[0];
                }
                else
                {
                    return fullFileName + "Dosya Bulunamadı!";
                }
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                return "Error! Check Logging!";
            }
            
        }

        public int insertIntoDosyaArsiv(int kurulId, int ilId, int ilceId, int dosyaNo, string desimalNo, string onayKodu, string onayTarihi)
        {
            try
            {
                string conString = "server=" + dbHost + ";uid=" + dbUser + ";pwd=" + dbPass + ";database=" + dbName;
                using (SqlConnection con = new SqlConnection(conString))
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand("INSERT INTO DOSYA_ARSIV (KURUL_ID, IL_ID, ILCE_ID, DOSYA_DESIMAL, PROJE_ONAY_SAYISI,PROJE_ONAY_TARIHI, KAYIT_NO) VALUES (" + kurulId + ", " + ilId + "," + ilceId + ",'" + desimalNo + "','" + onayKodu + "','" + onayTarihi + "'," + dosyaNo + ")"))
                    {
                        int oid = cmd.ExecuteNonQuery();
                        return oid;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw;
            }
        }
        //Not used for speedUp reasons
        public int checkNInsert(int kurulId, int ilId, int ilceId,int dosyaNo,string desimalNo, string onayKodu, string onayTarihi)
        {
            try
            {
                string conString = "server=" + dbHost + ";uid=" + dbUser + ";pwd=" + dbPass + ";database=" + dbName;
                using (SqlConnection con = new SqlConnection(conString))
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand("SELECT * FROM DOSYA_ARSIV WHERE DOSYA_DESIMAL = '" + desimalNo+"'", con))
                    {
                        using (IDataReader dr = cmd.ExecuteReader())
                        {
                            if (dr.Read())
                            {
                                return 0;
                            }
                            else
                            {
                                using (SqlCommand icmd = new SqlCommand("INSERT INTO DOSYA_ARSIV (KURUL_ID, IL_ID, ILCE_ID, DOSYA_DESIMAL, PROJE_ONAY_SAYISI,PROJE_ONAY_TARIHI, KAYIT_NO) VALUES ("+kurulId+", "+ilId+","+ilceId+",'"+desimalNo+"','"+onayKodu+"','"+onayTarihi+"',"+dosyaNo+")"))
                                {
                                    int oid = icmd.ExecuteNonQuery();
                                    return oid;
                                }
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw;
            }
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
                    //getExcelSheets(textBox2.Text);
                    backgroundWorker2.RunWorkerAsync();
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
            textBox3.Text = mainPath + comboBox1.SelectedItem.ToString() + "\\" + comboBox1.SelectedItem.ToString() + "\\";
            if (Directory.Exists(textBox3.Text))
            {
                textBox3.BackColor = Color.LightGreen;
                textBox3.ForeColor = Color.DarkGreen;
                dosyalarPath = textBox3.Text;

            }
            else
            {
                textBox3.BackColor = Color.LightPink;
                textBox3.ForeColor = Color.DarkRed;
            }
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

        private void buttonQDT_Click(object sender, EventArgs e)
        {
            mergeTables();

            if (label14.Text == "Disconnected!")
            {
                MessageBox.Show("Lütfen Veritabanı Bilgilerini Giriniz!");
                tabControl1.SelectedIndex = 4;
            }
            else if (label14.Text == "Connected!")
            {
                progressBar1.Value = 0;
                buttonQDT.Enabled = false;
                backgroundWorker1.RunWorkerAsync();
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            int rowNumber = 0;
            foreach (DataRow row in dosyaArsiv.Rows)
            {
                
                int? idKayit, onayNo, ilId, ilceId;
                string desimalNo, filePath, turu,barkodNo;
                DateTime? onayTarihi;

                //Read From Dosyalar Excel
                if (row[0] != DBNull.Value)
                {
                    idKayit = Convert.ToInt32(row[0]);
                }
                else
                {
                    idKayit = null;
                }
                if (row[1] != DBNull.Value)
                {
                    desimalNo = row[1].ToString();
                }
                else
                {
                    desimalNo = null;
                }
                if (row[3] != DBNull.Value)
                {
                    onayNo = Convert.ToInt32(row[3]);
                }
                else
                {
                    onayNo = null;
                }
                string onayTarihiSQL;
                if (row[4] != DBNull.Value)
                {
                    //onayTarihi = DateTime.FromOADate(Convert.ToDouble(row[4]));
                    onayTarihiSQL = Convert.ToString(row[4]);
                }
                else
                {
                    onayTarihi = null;
                    onayTarihiSQL = onayTarihi.ToString();
                }

                turu = Convert.ToString(row[5]);
                barkodNo = Convert.ToString(row[6]);
                ilId = Convert.ToInt32(row[7]);
                ilceId = Convert.ToInt32(row[8]);


                if (turu == ".tif")
                {
                    if (isExists(desimalNo, onayNo.ToString(), onayTarihiSQL))
                    {
                        row[9] = 1;//existence!
                        filePath = dosyaArsiv_path(idKayit.ToString(), turu);
                        row[10] = filePath;//filepath
                    }
                    else
                    {
                        //Search for tiff file
                        filePath = dosyaArsiv_path(idKayit.ToString(), turu);
                        row[9] = 0;
                        checkNInsert(1, Convert.ToInt32(ilId), Convert.ToInt32(ilceId), Convert.ToInt32(idKayit), desimalNo, onayNo.ToString(), onayTarihiSQL);
                    }
                }
                else
                {
                    //something for .pdf files
                }
                


                //progressbar precentage increment
                rowNumber++;
                int percentage = (rowNumber * 100) / dosyaArsiv.Rows.Count;
                backgroundWorker1.ReportProgress(percentage);
            }
        }

        void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            buttonQDT.Enabled = true;
            dataGridView4.DataSource = null;
            dataGridView4.DataSource = dosyaArsiv;
        }

        //second worker for excel progress watch
       
        private void button3_Click(object sender, EventArgs e)
        {
            string path = mainPath + comboBox1.SelectedItem.ToString() + "\\";
            FolderBrowserDialog folderBrowserdialog1 = new FolderBrowserDialog();
            if (Directory.Exists(path))
            {
                folderBrowserdialog1.SelectedPath = path;
            }
            else
            {
                folderBrowserdialog1.SelectedPath = mainPath;
            }
            if (folderBrowserdialog1.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = folderBrowserdialog1.SelectedPath;
                if (Directory.Exists(textBox3.Text))
                {
                    textBox3.BackColor = Color.LightGreen;
                    textBox3.ForeColor = Color.DarkGreen;
                }
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (Directory.Exists(textBox3.Text))
            {
                textBox3.BackColor = Color.LightGreen;
                textBox3.ForeColor = Color.DarkGreen;
                dosyalarPath = textBox3.Text;
            }
            else
            {
                textBox3.BackColor = Color.LightPink;
                textBox3.ForeColor = Color.DarkRed;
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            dbName = comboBox2.SelectedItem.ToString();
            buttonDbConnect.Enabled = true;
        }

        //2nd bg worker
        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            string excelFile = textBox2.Text;
            try
            {
                logging(excelFile + "-Veri DataTable Aktarım Başladı");

                //listBox1.Items.Clear();
                var application = new Microsoft.Office.Interop.Excel.Application();
                var workbook = application.Workbooks.Open(excelFile);
                var worksheet_1 = workbook.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
                var worksheet_2 = workbook.Worksheets[2] as Microsoft.Office.Interop.Excel.Worksheet;

                var sheetName_1 = worksheet_1.Name.ToString();
                var sheetName_2 = worksheet_2.Name.ToString();

                //listBox1.Items.Add(sheetName_1);
                //listBox1.Items.Add(sheetName_2);

                int rowCountTiff = worksheet_1.UsedRange.Rows.Count;
                int rowCountPdf = worksheet_2.UsedRange.Rows.Count;
                int totalRowCount = rowCountPdf + rowCountTiff;
                int currentRow = 1;


                tiffFilesDT.Columns.Add("ID_kayıt_no");
                tiffFilesDT.Columns.Add("desimal_no");
                tiffFilesDT.Columns.Add("proje_açıklaması");
                tiffFilesDT.Columns.Add("onay_no");
                tiffFilesDT.Columns.Add("onay_tarihi");

                pdfFilesDT.Columns.Add("id_kayit_no");
                pdfFilesDT.Columns.Add("barkod_no");
                pdfFilesDT.Columns.Add("desimal_no");

                DataRow row;

                int index = 0;
                object rowIndex = 2;
                while (((Microsoft.Office.Interop.Excel.Range)worksheet_1.Cells[rowIndex, 1]).Value2 != null)
                {
                    rowIndex = 2 + index;
                    row = tiffFilesDT.NewRow();

                    row[0] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)worksheet_1.Cells[rowIndex, 1]).Value2);
                    row[1] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)worksheet_1.Cells[rowIndex, 2]).Value2);
                    row[2] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)worksheet_1.Cells[rowIndex, 3]).Value2);
                    row[3] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)worksheet_1.Cells[rowIndex, 4]).Value2);
                    row[4] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)worksheet_1.Cells[rowIndex, 5]).Value2);
                    index++;
                    tiffFilesDT.Rows.Add(row);
                    currentRow++;
                    backgroundWorker2.ReportProgress(currentRow);
                }
                logging(excelFile + "-tiff Dosyalar Listelenmesi Tamamlandı");
                index = 0;
                rowIndex = 2;
                while (((Microsoft.Office.Interop.Excel.Range)worksheet_2.Cells[rowIndex, 1]).Value2 != null)
                {
                    rowIndex = 2 + index;
                    row = pdfFilesDT.NewRow();

                    row[0] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)worksheet_2.Cells[rowIndex, 1]).Value2);
                    row[1] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)worksheet_2.Cells[rowIndex, 2]).Value2);
                    row[2] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)worksheet_2.Cells[rowIndex, 3]).Value2);
                    index++;
                    pdfFilesDT.Rows.Add(row);
                    currentRow++;
                    backgroundWorker2.ReportProgress(currentRow);
                }
                logging(excelFile + "-pdf Dosyalar Listelenmesi Tamamlandı");

                workbook.Close(false, Missing.Value, Missing.Value);
                application.Quit();


                logging(excelFile + "-Listelenme Tamamlandı");
            }
            catch (Exception ex)
            {


                logging(excelFile + "-" + ex.ToString());
            }
        }
        void backgroundWorker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            lbExcelRowNu.Text = e.ProgressPercentage.ToString();
        }
        void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            lbExcelRowNu.Text = "Completed with "+lbExcelRowNu.Text;
            //completed..
        }
    }
}