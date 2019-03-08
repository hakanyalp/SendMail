using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace mail_gonder
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            MaximizeBox = false;
            numericUpDown1.Value = wait_time / 1000;    // milisaniyeyi saniye cinsinden elde edebilmek için 1000e bölünür

            try
            {
                StreamReader reader = new StreamReader("./mail_list.txt");
                int repeat_count = 0;
                List<string> err_mail = new List<string>();     //txt'deki hatalı mail adresleri burada tutulup log'lanacak
                while (!reader.EndOfStream)
                {
                    string text = reader.ReadLine();

                    bool repeat_control = false;
                    bool err_control = false;
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        repeat_control = false;

                        // tekrar kontrolü yapılıyor
                        if (text == checkedListBox1.Items[i].ToString())    // eğer eklenecek eleman listede varsa eklenmeyecek
                        {
                            repeat_control = true;
                            repeat_count++;
                            break;
                        }
                        //mail formatı uygun mu kontrolü yapılıyor
                        else if (string.IsNullOrEmpty(text) || !Regex.IsMatch(text, pattern))   // boş değilse veya format uygun değilse kontrolü
                        {
                            err_control = true;
                        }
                    }
                    if (!repeat_control && !err_control)
                        checkedListBox1.Items.Add(text);

                    if (err_control)    // gelen satır mail formatına uygun değilse log'lamak için listeye ekleniyor
                        err_mail.Add(text);
                }
                reader.Close();
            }
            catch
            {
                // hiçbir şey yapma
            }
        }

        int wait_time = 1000;   // milisecond
        private void button1_Click(object sender, EventArgs e)
        {
            //mail gönder

            DialogResult dialogResult = MessageBox.Show("Seçili kişilere mail gönderilecek, emin misiniz?", "Mail Gönderilecek", MessageBoxButtons.YesNo);

            bool mailhtml_control = true;       //mail.html dosyası var mı diye kontrol ediliyor
            if (dialogResult == DialogResult.Yes)
            {
                for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
                {
                    try
                    {
                        string senderEmail = "kodyazari0@gmail.com";  //mail adresi kullanıcı mail adresi girilecek
                        string senderPassword = ".Q1w2e3r4";  //mail adresi kullanıcı şifresi girilecek

                        SmtpClient client = new SmtpClient("smtp.gmail.com", 587);
                        client.Timeout = 100000;
                        client.DeliveryMethod = SmtpDeliveryMethod.Network;
                        client.UseDefaultCredentials = false;
                        client.Credentials = new NetworkCredential(senderEmail, senderPassword);
                        client.EnableSsl = true;

                        string toEmail;
                        toEmail = checkedListBox1.CheckedItems[i].ToString();
                        string subject = "Otomatik Mail";
                        string emailBody = "";

                        // mail gönderilecek döküman
                        string fileName = "mail" + ".html";
                        // Read using File.OpenText
                        if (System.IO.File.Exists(fileName))
                        {
                            using (System.IO.StreamReader sr = System.IO.File.OpenText(fileName))
                            {
                                String input;
                                while ((input = sr.ReadLine()) != null)
                                {
                                    emailBody += input + "\r\n";
                                }
                            }
                        }
                        if (emailBody == "")    // mail içeriği boşsa mail gönderme olmayacak
                        {
                            mailhtml_control = false;
                            break;
                        }

                        MailMessage mailMessage = new MailMessage(senderEmail, toEmail, subject, emailBody);
                        mailMessage.IsBodyHtml = true;

                        mailMessage.BodyEncoding = UTF8Encoding.UTF8;

                        client.Send(mailMessage);
                        System.Threading.Thread.Sleep(wait_time);

                        lblSonMailTarihi.Text = DateTime.Now.ToString();
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show(checkedListBox1.CheckedItems[i].ToString() + " kişisine mail gönderme başarısız oldu!");
                    }
                }
            }
            else if (dialogResult == DialogResult.No)
            {
                MessageBox.Show("Mail gönderme iptal edildi");
            }

            if (!mailhtml_control)
            {
                MessageBox.Show("mail.html dosyası bulunamadı!");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //içeri aktar

            OpenFileDialog ofd = new OpenFileDialog { Filter = "Metin Dosyası | *.txt;*.html" };
            ofd.ShowDialog();
            if (!string.IsNullOrEmpty(ofd.FileName))
            {
                StreamReader reader = new StreamReader(ofd.FileName);
                int repeat_count = 0;
                List<string> err_mail = new List<string>();     //txt'deki hatalı mail adresleri burada tutulup log'lanacak
                while (!reader.EndOfStream)
                {
                    string text = reader.ReadLine();

                    bool repeat_control = false;
                    bool err_control = false;
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        repeat_control = false;

                        // tekrar kontrolü yapılıyor
                        if (text == checkedListBox1.Items[i].ToString())    // eğer eklenecek eleman listede varsa eklenmeyecek
                        {
                            repeat_control = true;
                            repeat_count++;
                            break;
                        }
                        //mail formatı uygun mu kontrolü yapılıyor
                        else if (string.IsNullOrEmpty(text) || !Regex.IsMatch(text, pattern))   // boş değilse veya format uygun değilse kontrolü
                        {
                            err_control = true;
                        }
                    }
                    if (!repeat_control && !err_control)
                        checkedListBox1.Items.Add(text);

                    if (err_control)    // gelen satır mail formatına uygun değilse log'lamak için listeye ekleniyor
                        err_mail.Add(text);
                }
                reader.Close();

                //hata mesajı verilecek kısım
                if (repeat_count > 0 || err_mail.Count > 0)
                {
                    string err_message = "";
                    if (repeat_count > 0 && err_mail.Count > 0)     //hem listede aynı kişi var, hem de mail formatı uygun olmayanlar var
                    {
                        err_message += repeat_count + " kişi listede kayıtlı olduğu için, " + err_mail.Count() + " mail adresi formatı uygun olmadığı için eklenemedi. Ayrıca hatalı mail adreslerinin log kaydı yapıldı.";

                        //burada ayrıca eklenemeyenlerin logu tutulacak
                        for (int i = 0; i < err_mail.Count; i++)
                        {
                            File.AppendAllText("log.txt", err_mail[i] + Environment.NewLine);
                        }
                    }
                    else if (repeat_count > 0 && err_mail.Count == 0)      //listede aynı kişi var, bütün maillerin formatı doğru
                    {
                        err_message += repeat_count + " kişi listede kayıtlı olduğu için eklenemedi.";
                    }
                    else if (repeat_count == 0 && err_mail.Count > 0)        //tekrar eden yok, mail formatı yanlış olan var
                    {
                        err_message += err_mail.Count() + " mail adresi formatı uygun olmadığı için eklenemedi. Hatalı mail adreslerinin log kaydı yapıldı";

                        //burada ayrıca eklenemeyenlerin logu tutulacak
                        for (int i = 0; i < err_mail.Count; i++)
                        {
                            File.AppendAllText("log.txt", err_mail[i] + Environment.NewLine);
                        }
                    }
                    MessageBox.Show(err_message);
                }
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //seçilenleri sil
            int count = checkedListBox1.CheckedItems.Count;
            List<string> deleting = new List<string>();
            for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
            {
                deleting.Add(checkedListBox1.CheckedItems[i].ToString());
            }
            bool delete_control = false;
            for (int i = 0; i < deleting.Count; i++)
            {
                checkedListBox1.Items.Remove(deleting[i]);
                delete_control = true;
            }
            if (!delete_control)
                MessageBox.Show("Silinecek eleman seçmelisiniz");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //tümünü seç
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, true);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //tümünü kaldır
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, false);
            }
        }
        //mail pattern

        string pattern = @"^(([\w-]+\.)+[\w-]+|([a-zA-Z]{1}|[\w-]{2,}))@"
   + @"((([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\."
   + @"([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\."
   + @"([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\."
   + @"([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])){1}|"
   + @"([a-zA-Z]+[\w-]+\.)+[a-zA-Z]{2,4})$";

        private void button6_Click(object sender, EventArgs e)
        {
            //kullanıcının girdiği mail adresini ekle

            string data = textBox1.Text;

            if (!string.IsNullOrEmpty(data) && Regex.IsMatch(data, pattern))
            {
                bool repeat_control = false;
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    if (data == checkedListBox1.Items[i].ToString())    // eğer eklenecek eleman listede varsa eklenmeyecek
                    {
                        repeat_control = true;
                        break;
                    }
                }
                if (!repeat_control)
                {
                    checkedListBox1.Items.Add(data);
                    textBox1.Clear();
                }
                else
                    MessageBox.Show("Kişi zaten listede kayıtlı!");
            }

            else
            {
                MessageBox.Show("Mail adresi geçerli değildir!");
            }

        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                button6.PerformClick();
            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            wait_time = Convert.ToInt32(numericUpDown1.Value) * 1000;
        }
    }
}
