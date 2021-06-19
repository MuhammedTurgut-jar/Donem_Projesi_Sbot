using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;

namespace Sbot
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
      
        List<string> ad;
        List<string> soyad;
        List<string> eposta;
        List<string> kullaniciAdi;
        List<string> pozisyon;
        List<string> yetki;
        List<string> basvuru;

        private void Form1_Load(object sender, EventArgs e) { }
        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "Kullanıcı İşlemleri" && comboBox2.Text == "Ekle")
            {
                
                Boolean status = Sbot.driver.KullaniciEkle(textBox2.Text, textBox3.Text, ad, soyad, eposta, kullaniciAdi);

                if (status == false)
                {
                    MessageBox.Show("Kullanıcı adı ve/veya parola yanlış!");
                }
                else
                {
                    MessageBox.Show("İşlem tamamlandı!");
                }

                ad = null;
                soyad = null;
                eposta = null;
                kullaniciAdi = null;
                yetki = null;
                pozisyon = null;
                basvuru = null;


            }

            if (comboBox1.Text == "Kullanıcı İşlemleri" && comboBox2.Text == "Sil")
            {

                Boolean status = Sbot.driver.KullaniciSil(textBox2.Text, textBox3.Text, kullaniciAdi);

                if (status == false)
                {
                    MessageBox.Show("Kullanıcı adı ve/veya parola yanlış!");
                }
                else
                {
                    MessageBox.Show("İşlem tamamlandı!");
                }

                ad = null;
                soyad = null;
                eposta = null;
                kullaniciAdi = null;
                yetki = null;
                pozisyon = null;
                basvuru = null;

            }

            if (comboBox1.Text == "Pozisyon İşlemleri" && comboBox2.Text == "Ekle")
            {

                Boolean status = Sbot.driver.PozisyonEkle(textBox2.Text, textBox3.Text, kullaniciAdi,pozisyon);

                if (status == false)
                {
                    MessageBox.Show("Kullanıcı adı ve/veya parola yanlış!");
                }
                else
                {
                    MessageBox.Show("İşlem tamamlandı!");
                }

                ad = null;
                soyad = null;
                eposta = null;
                kullaniciAdi = null;
                yetki = null;
                pozisyon = null;
                basvuru = null;

            }
           
            if (comboBox1.Text == "Yetki İşlemleri" && comboBox2.Text == "Ekle")
            {

                Boolean status = Sbot.driver.YetkiEkle(textBox2.Text, textBox3.Text, kullaniciAdi, yetki);

                if (status == false)
                {
                    MessageBox.Show("Kullanıcı adı ve/veya parola yanlış!");
                }
                else
                {
                    MessageBox.Show("İşlem tamamlandı!");
                }

                ad = null;
                soyad = null;
                eposta = null;
                kullaniciAdi = null;
                yetki = null;
                pozisyon = null;
                basvuru = null;

            }


            if (comboBox1.Text == "Başvuru İşlemleri" && comboBox2.Text == "Transfer")
            {
                Boolean status = Sbot.driver.BasvuruTansfer(textBox2.Text, textBox3.Text, kullaniciAdi,basvuru);

                if (status == false)
                {
                    MessageBox.Show("Kullanıcı adı ve/veya parola yanlış!");
                }
                else
                {
                    MessageBox.Show("İşlem tamamlandı!");
                }

                ad = null;
                soyad = null;
                eposta = null;
                kullaniciAdi = null;
                yetki = null;
                pozisyon = null;
                basvuru = null;

            }



        }
        private void button3_Click(object sender, EventArgs e)
        {
            

            textBox1.Text = excelData.DosyaYoluSec();
            dataGridView1.DataSource = excelData.DosyaOku();

        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        private void Form1_FormClosed(object sender, FormClosedEventArgs e) { }
        private void label1_Click(object sender, EventArgs e) { }
        private void textBox1_TextChanged(object sender, EventArgs e) { }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            ad = new List<string>();
            soyad = new List<string>();
            eposta = new List<string>();
            kullaniciAdi = new List<string>();
            yetki = new List<string>();
            pozisyon = new List<string>();
            basvuru = new List<string>();
            

            if (comboBox1.Text == "Kullanıcı İşlemleri" && comboBox2.Text == "Ekle")
            {
                
                dataGridView1.Columns.Add("Id", "Kullanıcı Adı");

                for (int i=0; i < dataGridView1.RowCount - 1; i++)
                {

                    eposta.Add(dataGridView1.Rows[i].Cells[0].Value.ToString().ToLower(new CultureInfo("en-US", false)));
                    ad.Add(dataGridView1.Rows[i].Cells[1].Value.ToString().ToUpper(new CultureInfo("en-US", false)));
                    soyad.Add(dataGridView1.Rows[i].Cells[2].Value.ToString().ToUpper(new CultureInfo("en-US", false)));
                    dataGridView1.Rows[i].Cells[3].Value = dataGridView1.Rows[i].Cells[0].Value.ToString().ToUpper(new CultureInfo("en-US", false)).Split('@')[0];
                    kullaniciAdi.Add(dataGridView1.Rows[i].Cells[3].Value.ToString());

                   
                }
                dataGridView1.Columns.Remove("Id");
            }

            if (comboBox1.Text == "Kullanıcı İşlemleri" && comboBox2.Text == "Sil")
            {

                dataGridView1.Columns.Add("Id", "Kullanıcı Adı");

                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    dataGridView1.Rows[i].Cells[1].Value = dataGridView1.Rows[i].Cells[0].Value.ToString().ToUpper(new CultureInfo("en-US", false)).Split('@')[0];
                    kullaniciAdi.Add(dataGridView1.Rows[i].Cells[1].Value.ToString());

                }
                dataGridView1.Columns.Remove("Id");
            }

            if (comboBox1.Text == "Pozisyon İşlemleri" && comboBox2.Text == "Ekle")
            {

                dataGridView1.Columns.Add("Id", "Kullanıcı Adı");

                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    pozisyon.Add(dataGridView1.Rows[i].Cells[1].Value.ToString());
                    dataGridView1.Rows[i].Cells[2].Value = dataGridView1.Rows[i].Cells[0].Value.ToString().ToUpper(new CultureInfo("en-US", false)).Split('@')[0];
                    kullaniciAdi.Add(dataGridView1.Rows[i].Cells[2].Value.ToString());

                }
                dataGridView1.Columns.Remove("Id");
            }

            if (comboBox1.Text == "Pozisyon İşlemleri" && comboBox2.Text == "Sil")
            {

                dataGridView1.Columns.Add("Id", "Kullanıcı Adı");

                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    pozisyon.Add(dataGridView1.Rows[i].Cells[1].Value.ToString());
                    dataGridView1.Rows[i].Cells[2].Value = dataGridView1.Rows[i].Cells[0].Value.ToString().ToUpper(new CultureInfo("en-US", false)).Split('@')[0];
                    kullaniciAdi.Add(dataGridView1.Rows[i].Cells[2].Value.ToString());

                }
                dataGridView1.Columns.Remove("Id");
            }

            if (comboBox1.Text == "Yetki İşlemleri" && comboBox2.Text == "Ekle")
            {

                dataGridView1.Columns.Add("Id", "Kullanıcı Adı");

                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    yetki.Add(dataGridView1.Rows[i].Cells[1].Value.ToString());
                    dataGridView1.Rows[i].Cells[2].Value = dataGridView1.Rows[i].Cells[0].Value.ToString().ToUpper(new CultureInfo("en-US", false)).Split('@')[0];
                    kullaniciAdi.Add(dataGridView1.Rows[i].Cells[2].Value.ToString());

                }
                dataGridView1.Columns.Remove("Id");
            }

            if (comboBox1.Text == "Yetki İşlemleri" && comboBox2.Text == "Sil")
            {

                dataGridView1.Columns.Add("Id", "Kullanıcı Adı");

                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    yetki.Add(dataGridView1.Rows[i].Cells[1].Value.ToString());
                    dataGridView1.Rows[i].Cells[2].Value = dataGridView1.Rows[i].Cells[0].Value.ToString().ToUpper(new CultureInfo("en-US", false)).Split('@')[0];
                    kullaniciAdi.Add(dataGridView1.Rows[i].Cells[2].Value.ToString());

                }
                dataGridView1.Columns.Remove("Id");
            }

            if (comboBox1.Text == "Başvuru İşlemleri" && comboBox2.Text == "Transfer")
            {
             
                dataGridView1.Columns.Add("Id", "Kullanıcı Adı");

                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    basvuru.Add(dataGridView1.Rows[i].Cells[0].Value.ToString());
                    dataGridView1.Rows[i].Cells[2].Value = dataGridView1.Rows[i].Cells[1].Value.ToString().ToUpper(new CultureInfo("en-US", false)).Split('@')[0];
                    kullaniciAdi.Add(dataGridView1.Rows[i].Cells[2].Value.ToString());
                  
                }
                dataGridView1.Columns.Remove("Id");
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
