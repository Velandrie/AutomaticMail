using System;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;

namespace AutomaticMail
{
    public partial class MainForm : MetroFramework.Forms.MetroForm
    {
        //Excel'den mail adresi, şifre, body, subject, receiver mail almak için kullanılacak string değişkenler.
        string bodyWomen = "";
        string bodyMen = "";
        string subject = "";
        string mail = "";
        string password = "";
        //Excel Dosyasına konulan şifre (güçlü değil, denemelik :) )
        string excelPassword = "nexthorizons112";

        //Program açıldığında Excel dosyasının path'ini otomatikman bulması için verdiğim path.(Excel dosyası, uygulamanın dosyalarının içinde, bin/debug içerisinde.)
        string path = Environment.CurrentDirectory + "\\Excel File\\Excel_MailList.xlsx";

        public MainForm()
        {
            InitializeComponent();

        }

        private void sendBtn_Click(object sender, EventArgs e)
        {
            //Gönder Butonuna basıldığında ReadExcel metodu tetikleniyor.
            ReadExcel();

        }

        //Excel dosyasından ilgili verilen okunup sonra maillere mesajların gönderilmesini sağlayan metod
        public void ReadExcel()
        {
            //Excel Application, dosya ve uzunluk tanımlama işlemleri
             Excel.Application xlApp = new Excel.Application();
            //Şifreli Excel dosyası açılırken şifresi girilerek açılıyor.
             Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Filename: path, Password: excelPassword);
             Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
             Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
                colCount = 2;

             //Kadın ve erkeğe gönderilecek sabit mesajlar, konu, kendimize ait mail ve şifresinin sabit bulunduğu konumları tanımladı.
             bodyWomen = (string)(xlWorksheet.Cells[2, 5] as Excel.Range).Value;
             bodyMen = (string)(xlWorksheet.Cells[3, 5] as Excel.Range).Value;
             subject = (string)(xlWorksheet.Cells[2, 6] as Excel.Range).Value;
             mail = (string)(xlWorksheet.Cells[2, 7] as Excel.Range).Value;
             password = (string)(xlWorksheet.Cells[2, 8] as Excel.Range).Value;

            //Excel sütununda içerisinde veri oldukça okumaya devam edecek.
                for (int i = 2; i <= rowCount; i++)
                {
                        //Cinsiyet kısmında erkek mi ya da kadın ve boş mu, dolu mu olduğuna bakan karar mekanizması.                    
                        if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, colCount].Value2 != null && xlRange.Cells[i,3].Value2.ToString() == "Erkek")
                        {
                            //Gönderilen kişi erkekse, ilgili kişiye ona göre mesaj gönderiliyor.
                            SendMessage(xlRange.Cells[i, 2].Value2.ToString(), subject, bodyMen, mail, password);
                            //ListBox'a, gönderilen kişinin mail'i yazdırılıyor.
                            listBox1.Items.Add("'" + ((string)(xlWorksheet.Cells[i, colCount] as Excel.Range).Value) + "'" + " adresine otomatik mesajınız gönderildi");
                        }
                        else if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, colCount].Value2 != null && xlRange.Cells[i, 3].Value2.ToString() == "Kadın")
                        {
                            //Gönderilen kişi kadınsa, ilgili kişiye ona göre mesaj gönderiliyor.
                            SendMessage(xlRange.Cells[i, 2].Value2.ToString(), subject, bodyWomen, mail, password);
                            //ListBox'a, gönderilen kişinin mail'i yazdırılıyor.
                            listBox1.Items.Add("'" + ((string)(xlWorksheet.Cells[i, colCount] as Excel.Range).Value) + "'" + " adresine otomatik mesajınız gönderildi");
                        }

                }
            //Excel dosyasını bırakma ve kapatma işlemleri. (bazı durumlarda kapatmıyor.)
             GC.Collect();
             GC.WaitForPendingFinalizers();
             Marshal.ReleaseComObject(xlRange);
             Marshal.ReleaseComObject(xlWorksheet);
             xlWorkbook.Close();
             Marshal.ReleaseComObject(xlWorkbook);
             xlApp.Quit();
             Marshal.ReleaseComObject(xlApp);

            //Kapattığından emin olmak için, Excel işlemlerini kapatan metodu burada da çağırdım.
            processKill();
            
            /*catch
            {
                browseBtn.PerformClick();
            }*/

        }

        // Mesajın ilgili kişinin mail adresine gönderilmesi işlemi. Metod, alıcı maili, mesajı, mesaj başlığını, mailimizin adresini ve şifresini alıyor.
        public void SendMessage(string receiverMail, string subject, string body, string mail, string password)
        {
            try
            {
                MailAddress fromAddress = new MailAddress(mail);
                MailAddress toAddress = new MailAddress(receiverMail);
                string fromPassword = password;
                SmtpClient smtp = new SmtpClient
                {
                    // SMTP üzerinden mesajın gönderilmesi işlemi. Gmail ve Hotmail'e gönderebiliyor, diğerleri için denenmedi.
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(fromAddress.Address, fromPassword)
                };
                using (var message = new MailMessage(fromAddress, toAddress)
                {
                    Subject = subject,
                    Body = body
                })
                {
                    smtp.Send(message);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //Dosya Seç butonuna tıklama işlemi.
        private void browseBtn_Click(object sender, EventArgs e)
        {
            //Dosya Seç butonuna tıklanınca bir Open File Dialog açılıyor
            OpenFileDialog folder = new OpenFileDialog();

            //Sadece Excel dosyaları görünecek şekilde açılıyor.
            folder.Title = "Excel dosyanızı seçin";
            folder.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";


            if (folder.ShowDialog() == DialogResult.OK)
            {
                //Seçilen dosyanın Path'ı alınarak genel path string'ine yazdırılıyor.
                string sFileName = folder.FileName;
                string[] arrAllFiles = folder.FileNames;
                //Seçilen Excel'den veri çekildikten sonra isimleri gösteren ListView de çağrılıyor.
                path = sFileName;   

            }

        }

        //İsimlerin ekranda gösterilmesini sağlayacak ListView metodu, Excel'in path'ini parametre olarak alıyor.
        public void nameListview(string path)
        {
            //Excel işlemleri
            Excel.Application xlApp = new Excel.Application();
            //Şifreli Excel dosyası açılırken şifresi girilerek açılıyor.
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Filename : path, Password : excelPassword);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            colCount = 1;
            for (int i = 2; i <= rowCount; i++)
            {

                if (xlRange.Cells[i, colCount] != null && xlRange.Cells[i, colCount].Value2 != null)
                {
                    // Excel'deki isim soyisim sütununda ne kadar veri varsa ListView'e yazdırılıyor
                    listView_name.Items.Add(xlRange.Cells[i, 1].Value2.ToString());
                }

            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
          
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            //Açılan Excel dosyalarını tam olarak kapatmadığını gözlemlediğim için processKill metodunu burada da çağırdım.
            processKill();
        }

        //Form load olduğunda otomatik olarak isimleri gösteren ListView'in çağrılması için kullanılan event.
        private void MainForm_Load(object sender, EventArgs e)
        {
            //Excel dosyasının path'i bulunamazsa kullanıcıdan "Dosya Seç" butonuna basmasını istiyor.
            try
            {
                nameListview(path);
            }
            catch(Exception)
            {
                MessageBox.Show("Excel Dosyasını bulamadık, lütfen 'Dosya Seç' butonuna basarak seçin.");
            }
        }

        //ListView'de bulunan isimlere tıklandığında, ona karşılık gelen notun açılmasını sağlayan Event.
        private void listView_name_DoubleClick(object sender, EventArgs e)
        {
            //Excel işlemleri
            Excel.Application xlApp = new Excel.Application();
            //Şifreli Excel dosyası açılırken şifresi girilerek açılıyor.
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Filename: path, Password: excelPassword);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            //Tıklanan ismin Index numarasının alınması
            int var = listView_name.FocusedItem.Index;

            //Alınan index ile beraber, o isme karşılık gelen notun Excel dosyasından çekilmesi
            richTextBoxNotes.Text = (xlRange.Cells[var + 2, 4].Value2.ToString());

            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            //Açılan Excel dosyalarını tam olarak kapatmadığını gözlemlediğim için processKill metodunu burada da çağırdım.
            processKill();



        }

        //Eklenen notun Excel dosyasına yazdırılma işlemini yapan metod.
        public void writeExcel(int i, string note)
        {
            //Excel işlemleri.
            Excel.Application xlApp = new Excel.Application();
            //Şifreli Excel dosyası açılırken şifresi girilerek açılıyor.
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Filename: path, Password: excelPassword);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            // TextView'e eklenen notun ilgili excel dosyasındaki bölüme yazdırılması
            xlRange.Cells[i + 2, 4].Value2 = note.ToString();

            GC.Collect();
            GC.WaitForPendingFinalizers();            
            
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            //processKill metodunu burada da çağırdım.
            processKill();
        }

        //TextBox'a yazılan notun önce bir string değişkene aktarılması ve bu değişken yardımıyla writeExcel metoduna aktarılması.
        private void button2_Click(object sender, EventArgs e)
        {
           int var = listView_name.FocusedItem.Index;
           string a = richTextBoxNotes.Text;
        
           writeExcel(var, a);
           
        }

        //Arka planda çalışan process'lere erişerek, Excel'in bulunup, process'in kapatılmasına yarayan metod.
        public void processKill()
        {
            Process[] _proceses = null;
            _proceses = Process.GetProcessesByName("Excel");
            foreach (Process proces in _proceses)
            {
                proces.Kill();
            }
        }

        //Uygulama içerisinden Excel Dosyasına erişmek için kullanılan button
        private void OpenExcelFileBtn_Click(object sender, EventArgs e)
        {
            //Excel dosyasının path'ı kullanılıyor
            FileInfo fi = new FileInfo(path);
            if (fi.Exists)
            {
                
                Process.Start(new ProcessStartInfo(path));
            }
            else
            {
                MessageBox.Show("Excel Dosyasını bulamadık, lütfen 'Dosya Seç' butonuna basarak seçin.");
            }
        }

        
    }
}
