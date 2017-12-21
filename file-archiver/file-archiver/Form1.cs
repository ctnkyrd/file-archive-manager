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

            groupBox1.Text = mainPath;
        }

        //Arşiv folder path

         


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


        private void button2_Click(object sender, EventArgs e)
        {
            string path = mainPath + comboBox1.SelectedItem.ToString()+"\\";
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
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            changeColorTextboxExcel(textBox2);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            changeColorComboKurul(comboBox1);
        }
    }
}