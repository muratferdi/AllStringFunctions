using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Device.Location;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media.Effects;

namespace ListeAraclar
{
    public partial class ListeAraclari : UserControl
    {
        Stack<string> liste1Islemler = new Stack<string>();
        Stack<string> liste2Islemler = new Stack<string>(); 

        public ListeAraclari()
        {
            InitializeComponent();
            if (DesignerProperties.GetIsInDesignMode(this))
            {
                return;
            }
            Loaded += ListeAraclari_Loaded;
        }

        async void ListeAraclari_Loaded(object sender, RoutedEventArgs e)
        {
            Loaded -= ListeAraclari_Loaded;
            liste1Islemler.Push(txtListe1.Text);
            liste2Islemler.Push(txtListe2.Text);
        }
        

        private string SonIslemGetir1(string metin)
        {
            var result = metin;
            if (liste1Islemler.Count > 0)
            {
                result = liste1Islemler.Pop();
            }
            return result;
        }

        private string SonIslemGetir2(string metin)
        {
            var result = metin;
            if (liste2Islemler.Count > 0)
            {
                result = liste2Islemler.Pop();
            }
            return result;
        }

        string TekrarTemizle(string metin)
        {
            var result = metin;
            if (!string.IsNullOrWhiteSpace(metin))
            {
                var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                var temiz = liste.Distinct().ToList();
                MesajWindow.GosterUyari((liste.Count() - temiz.Count()) + " Adet tekrar eden kayıt silindi.");
                result = string.Join(Environment.NewLine, temiz.ToArray());
            }
            return result.Trim();
        }

        string DosyaYukle()
        {
            var dlg = new OpenFileDialog();
            dlg.Filter = "*.TXT|*.TXT";
            dlg.Multiselect = true;
            if (dlg.ShowDialog() == true)
            {
                var sb = new StringBuilder();
                foreach (var item in dlg.FileNames)
                {
                    sb.AppendLine(File.ReadAllText(item, Encoding.UTF8).Trim());
                }
                return sb.ToString();
            }
            else
            {
                return null;
            }
        }

        void DosyaKaydet(string metin)
        {
            var dlg = new SaveFileDialog();
            dlg.FileName = DateTime.Now.ToString("yyyyMMdd_HHmmSS") + "_DosyaKaydet.TXT";
            dlg.Filter = "Metin Dosyası (*.TXT)|*.TXT|CSV Dosyası (*.CSV)|*.CSV";
            if (dlg.ShowDialog() == true)
            {
                File.WriteAllText(dlg.FileName, metin, Encoding.UTF8);
            }
        }

        string TCKNDogrula(string metin)
        {
            var result = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(metin))
            {
                var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                for (int i = 0; i < liste.Count(); i++)
                {
                    var satir = liste[i].Split(' ');
                    for (int j = 0; j < satir.Length; j++)
                    {
                        var s = satir[j].Replace("-", "").Replace("/", "").Replace("(", "").Replace(")", "").Trim();
                        if (DogrulamaYardimcisi.TcKimlikDogrumu(s))
                        {
                            result.AppendLine(s);
                        }
                    }
                }
            }
            return result.ToString();
        }

        string VKKNDogrula(string metin)
        {
            var result = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(metin))
            {
                var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                for (int i = 0; i < liste.Count(); i++)
                {
                    var satir = liste[i].Split(' ');
                    for (int j = 0; j < satir.Length; j++)
                    {
                        var s = satir[j].Replace("-", "").Replace("/", "").Replace("(", "").Replace(")", "").Trim();
                        if (DogrulamaYardimcisi.VergiNoDogruMu(s))
                        {
                            result.AppendLine(s);
                        }
                    }
                }
            }
            return result.ToString();
        }

        void SonucPenceresiOlustur(string baslik, string metin)
        {
            var w = new Window();
            var effect = new DropShadowEffect();
            effect.BlurRadius = 15;
            effect.Opacity = 0.4f;
            effect.BlurRadius = 15f;
            effect.Direction = 270f;
            w.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            w.Width = 800;
            w.Width = 600;
            var txtResult = new TextBox() { Margin = new Thickness(10), Effect = effect, AcceptsReturn = true, HorizontalAlignment = HorizontalAlignment.Stretch, VerticalAlignment = VerticalAlignment.Stretch, VerticalScrollBarVisibility = ScrollBarVisibility.Visible, VerticalContentAlignment = VerticalAlignment.Top };
            txtResult.Text = metin;
            w.Content = txtResult;
            w.Title = baslik + " (" + metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None).Where(x => !string.IsNullOrWhiteSpace(x)).Count() + ") Adet Satır";
            w.ShowDialog();
        }

        private void SadeceListe1(string metin1, string metin2)
        {
            var liste1 = metin1.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            var liste2 = metin2.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            var sonuc = liste1.Except(liste2);
            SonucPenceresiOlustur("Sadece Liste 1 içerisinde olanlar", string.Join(Environment.NewLine, sonuc.ToArray()));
        }

        private void SadeceListe2(string metin1, string metin2)
        {
            var liste1 = metin1.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            var liste2 = metin2.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            var sonuc = liste2.Except(liste1);
            SonucPenceresiOlustur("Sadece Liste 2 içerisinde olanlar", string.Join(Environment.NewLine, sonuc.ToArray()));
        }

        private void HemListe1HemListe2(string metin1, string metin2)
        {
            var liste1 = metin1.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            var liste2 = metin2.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            var sonuc = liste2.Intersect(liste1);
            SonucPenceresiOlustur("Hem Liste 1 Hemde Liste 2 içerisinde olanlar", string.Join(Environment.NewLine, sonuc.ToArray()));
        }

        private bool ListeParcala(string metin)
        {
            bool result = false;
            {
                var bilgi = BilgiGirisWindow.Show("1000", "Bölümlenmesini İstediğiniz Satır Sayısı Giriniz");
                if (bilgi.MessageBoxResult == MessageBoxResult.OK)
                {
                    int parca = 0;
                    try
                    {
                        parca = Convert.ToInt32(bilgi.Answer);
                        if (parca < 1)
                        {
                            MesajWindow.GosterHata("Bölümlenmesini İstediğiniz Satır Sayısını Yanlış Girdiniz");
                            return false;
                        }
                    }
                    catch (Exception)
                    {
                        MesajWindow.GosterHata("Bölümlenmesini İstediğiniz Satır Sayısını Yanlış Girdiniz");
                        return false;
                    }

                    MesajWindow.GosterBilgi("Sonuç Dosyalarının Kaydedileceği Klasörü Seçiniz");

                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "Metin Dosyaları|*.txt";
                    save.FileName = "ListeParca";
                    if (save.ShowDialog() == true)
                    {
                        var liste = metin.Trim().Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                        if (liste.Length > 0)
                        {
                            int dongu = liste.Length / parca;
                            for (int i = 0; i < dongu + 1; i++)
                            {
                                var dest = liste.Skip(i * parca).Take(parca).ToArray();
                                var path = System.IO.Path.GetDirectoryName(save.FileName);
                                var prefix = System.IO.Path.GetFileNameWithoutExtension(save.FileName);
                                string fileName = prefix + "_" + ((i * parca) + 1).ToString() + "-" + ((i + 1) * parca) + ".txt";
                                if (!string.IsNullOrEmpty(dest.ToString().Trim()))
                                {
                                    System.IO.File.WriteAllLines(path + "\\" + fileName, dest);
                                }
                            }
                            result = true;
                        }
                    }
                }
            }

            return result;
        }



        private string ExcelKolonOku()
        {
            var result = string.Empty;
            var dlg = new OpenFileDialog();
            dlg.Filter = "Excel Dosyaları|*.xls;*.xlsx";
            dlg.Multiselect = false;
            MesajWindow.GosterBilgi("Belirlenen kolon verilerinin alınacağı Excel dosyasını seçiniz.");

            if (dlg.ShowDialog() == true)
            {
                var bilgi = BilgiGirisWindow.Show("1", "Veri Alınacak Kolon Sırasını Giriniz.\nörn: A kolonu için 1 giriniz.\nBirden fazla kolon almak için 1,2,3 şeklide giriniz.");
                if (bilgi.MessageBoxResult == MessageBoxResult.OK)
                {
                    var secilenKolonlar = new List<int>();
                    try
                    {
                        Convert.ToInt32(bilgi.Answer.Replace(",", "").Trim());
                        var splitted = bilgi.Answer.Split(',');
                        foreach (var itemKolon in splitted)
                        {
                            secilenKolonlar.Add(Convert.ToInt32(itemKolon));
                        }
                    }
                    catch (Exception ex)
                    {
                        MesajWindow.GosterHata("Girilen bilgi istenilen formatta değil !" + Environment.NewLine + ex.Message);
                        return null;
                    }

                    try
                    {
                        string ayirici = "";
                        if (bilgi.Answer.Contains(","))
                        {
                            var bilgiAyirici = BilgiGirisWindow.Show("", "Birden fazla kolon seçtiğinizden dolayı\nAraya konulacak ayırıcı karakteri giriniz.\nÖrneğin ; yada |");
                            if (bilgiAyirici.MessageBoxResult == MessageBoxResult.OK)
                            {
                                ayirici = bilgiAyirici.Answer;
                            }
                        }

                        var dt = ExcelYardimcisi.ExcelToDataTable(dlg.FileName, 1, true);
                        if (dt == null || dt.Columns.Count < 1)
                        {
                            MesajWindow.GosterHata("Excel dosyasının üst bilgileri okunamadı.\nExcel dosyası bozuk olabilir. Yeniden oluşturunuz.");
                            return null;
                        }

                        var all = dt.Rows.Cast<DataRow>().Where(x => x[0] != null).Select(x => x[0].ToString()).ToList();
                        result = string.Join(Environment.NewLine, all);

                    }
                    catch (Exception ex)
                    {
                        MesajWindow.GosterHata("Kolon numarası için sadece rakam giriniz.\n" + ex.Message);
                        return null;
                    }
                }
            }

            return result;
        }


        private bool ExcelKolonEslestir()
        {
            var result = false;
            MesajWindow.GosterBilgi("Exceldeki düşey ara gibi ortak değer üzerinden iki excel dosyasını birleştirir.\nÖncelikle Eşleştirme yapılacak 2 excel dosyasında da 1. kolona ortak değeri taşıyınız.");

            MesajWindow.GosterBilgi("Şimdi 1. Ana  dosyayı seçiniz.");
            var dlg1 = new OpenFileDialog();
            dlg1.Filter = "Excel Dosyaları|*.xls;*.xlsx";
            dlg1.Multiselect = false;
            if (dlg1.ShowDialog() != true)
            {
                return false;
            }

            MesajWindow.GosterBilgi("Şimdi 2. Verilerin olduğu dosyayı seçiniz.");
            var dlg2 = new OpenFileDialog();
            dlg2.Filter = "Excel Dosyaları|*.xls;*.xlsx";
            dlg2.Multiselect = false;
            if (dlg2.ShowDialog() != true)
            {
                return false;
            }

            var dtMain1 = ExcelYardimcisi.ExcelToDataTable(dlg1.FileName);
            var dtMain2 = ExcelYardimcisi.ExcelToDataTable(dlg2.FileName);

            foreach (DataColumn item in dtMain2.Columns)
            {
                item.ColumnName = item.ColumnName + ">(Yeni2)";
            }

            foreach (DataColumn item in dtMain2.Columns)
            {
                dtMain1.Columns.Add(item.ColumnName, typeof(string));
            }

            foreach (DataRow item1 in dtMain1.Rows)
            {
                if (item1[0] != DBNull.Value)
                {
                    string ara = item1[0].ToString().Trim();
                    var bulunan = dtMain2.Rows.Cast<DataRow>().FirstOrDefault(x => x[0] != DBNull.Value && x[0].ToString().Trim() == ara);
                    if (bulunan != null)
                    {
                        foreach (DataColumn item2 in dtMain2.Columns)
                        {
                            item1[item2.ColumnName] = bulunan[item2.ColumnName];
                        }
                    }
                }
            }

            MesajWindow.GosterBilgi("Sonuçların kaydedileceği dosyayı seçiniz.");
            var dlgSave = new SaveFileDialog();
            dlgSave.Filter = "Excel Dosyası (*.xlsx)|*.xlsx";
            dlgSave.FileName = "SONUC_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            if (dlgSave.ShowDialog() == true)
            {
                ExcelYardimcisi.DataTableToExcel(dtMain1, dlgSave.FileName);
                result = true;
            }
            return result;
        }

        private bool ExcelDuseyAraBirlestir()
        {
            var result = false;
            MesajWindow.GosterBilgi("Düşey ara ile listeler çakıştırılıp istenilen Kolon içeriği tek hücrede birleştirilir.\nSorgulanacak değerler 1. excel dosyasına kaydedilir. 2. excel dosyasında ise aranacak anahtar(TCKN) A1 Kolon taşınır, düşey ara ile 2. excel aranır ve istenilen sütun içeriği tek bir hücre içerisinde birleştirilir.");

            MesajWindow.GosterBilgi("Sorgu Listesi Değerleri 1. dosyayı seçiniz.");
            var dlg1 = new OpenFileDialog();
            dlg1.Filter = "Excel Dosyaları|*.xls;*.xlsx";
            dlg1.Multiselect = false;
            if (dlg1.ShowDialog() != true)
            {
                return false;
            }

            MesajWindow.GosterBilgi("Şimdi Verilerin olduğu 2. dosyayı seçiniz.\nDikkat !! Sorgulanacak anahtar değeri(TCKN) A1 Kolonuna taşıyınız.");
            var dlg2 = new OpenFileDialog();
            dlg2.Filter = "Excel Dosyaları|*.xls;*.xlsx";
            dlg2.Multiselect = false;
            if (dlg2.ShowDialog() != true)
            {
                return false;
            }

            var dtMain1 = ExcelYardimcisi.ExcelToDataTable(dlg1.FileName);
            var dtMain2 = ExcelYardimcisi.ExcelToDataTable(dlg2.FileName);

            if (dtMain1.Columns.Count != 1)
            {
                MesajWindow.GosterUyari("1. dosyada sadece A1 sütununa sorgulanacak değerleri yazınız.");
                return false;
            }

            dtMain1.Columns.Add("BULUNDU_BİRLEŞTİRİLDİ", typeof(string));
            int birlestirColNum = -1;
            bool distinct = MesajWindow.GosterOnay("Bulanan Değerler Tekilleştirilsin mi ?") == MessageBoxResult.Yes;

            var bilgi = BilgiGirisWindow.Show("2", "Bulunup birleştirilecek Kolon Sırasını Giriniz.\nörn: A kolonu için 1 giriniz. B için 2 giriniz.");
            if (bilgi.MessageBoxResult == MessageBoxResult.OK)
            {
                try
                {
                    birlestirColNum = Convert.ToInt32(bilgi.Answer);
                }
                catch (Exception)
                {
                }
            }

            if (birlestirColNum != -1 && birlestirColNum <= dtMain2.Columns.Count)
            {
                //NOTHING
            }
            else
            {
                MesajWindow.GosterUyari("2. dosyada kolon sıra numarasında hata var. Lütfen Tekrar Deneyin.");
                return false;
            }

            foreach (DataRow rows1 in dtMain1.Rows)
            {
                var degerler = new List<string>();
                if (distinct)
                {
                    degerler = dtMain2.Rows.Cast<DataRow>().Where(x => x[0].ToString() == rows1[0].ToString()).Select(x => x[birlestirColNum - 1].ToString()).Distinct().OrderBy(x => x).ToList();
                }
                else
                {
                    degerler = dtMain2.Rows.Cast<DataRow>().Where(x => x[0].ToString() == rows1[0].ToString()).Select(x => x[birlestirColNum - 1].ToString()).ToList();
                }

                rows1[1] = string.Join("|", degerler);
            }

            MesajWindow.GosterBilgi("Sonuçların kaydedileceği dosyayı seçiniz.");
            var dlgSave = new SaveFileDialog();
            dlgSave.Filter = "Excel Dosyası (*.xlsx)|*.xlsx";
            dlgSave.FileName = "SONUC_BİRLEŞTİRİLDİ_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            if (dlgSave.ShowDialog() == true)
            {
                ExcelYardimcisi.DataTableToExcel(dtMain1, dlgSave.FileName);
                result = true;
            }
            return result;
        }

        private bool ExcelDuseyAraFiltrele()
        {
            var result = false;
            MesajWindow.GosterBilgi("Düşey ara ile listeler çakıştırılıp Sadece filtreye uygun satırlar bırakılır.\nSorgulanacak değerler 1. excel dosyasına kaydedilir. 2. excel dosyasında ise aranacak anahtar(TCKN) A1 Kolonuna taşınır, düşey ara ile 2. excel aranır ve uygun satırlar filtrelenmiş olur.");

            MesajWindow.GosterBilgi("1. dosyayı seçiniz. Sorgu Listesi Değerleri");
            var dlg1 = new OpenFileDialog();
            dlg1.Filter = "Excel Dosyaları|*.xls;*.xlsx";
            dlg1.Multiselect = false;
            if (dlg1.ShowDialog() != true)
            {
                return false;
            }

            MesajWindow.GosterBilgi("Şimdi Verilerin olduğu 2. dosyayı seçiniz.\nDikkat !! Sorgulanacak anahtar değeri(TCKN) A1 Kolonuna taşıyınız.");
            var dlg2 = new OpenFileDialog();
            dlg2.Filter = "Excel Dosyaları|*.xls;*.xlsx";
            dlg2.Multiselect = false;
            if (dlg2.ShowDialog() != true)
            {
                return false;
            }

            var dtMain1 = ExcelYardimcisi.ExcelToDataTable(dlg1.FileName);
            var dtMain2 = ExcelYardimcisi.ExcelToDataTable(dlg2.FileName);

            if (dtMain1.Columns.Count != 1)
            {
                MesajWindow.GosterUyari("1. dosyada sadece A1 sütununa sorgulanacak değerleri yazınız.");
                return false;
            }

            var sorguDegerleri = dtMain1.Rows.Cast<DataRow>().Select(x => x[0].ToString()).Distinct().ToList();
            var sonuclar = dtMain2.Rows.Cast<DataRow>().Where(x => sorguDegerleri.Contains(x[0].ToString())).ToList();


            MesajWindow.GosterBilgi("Sonuçların kaydedileceği dosyayı seçiniz.");
            var dlgSave = new SaveFileDialog();
            dlgSave.Filter = "Excel Dosyası (*.xlsx)|*.xlsx";
            dlgSave.FileName = "SONUC_BİRLEŞTİRİLDİ_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            if (dlgSave.ShowDialog() == true)
            {
                ExcelYardimcisi.DataTableToExcel(sonuclar.CopyToDataTable(), dlgSave.FileName);
                result = true;
            }
            return result;
        }
        private bool ExcelGruplaParcala()
        {
            var result = false;
            MesajWindow.GosterBilgi("Verilen Listeler çakıştırılıp Sadece filtreye uygun satırlar bırakılır ve ayrı ayrı dosyalara parçalanarak kaydedilir.\n1. excel dosyasına Sorgulanacak değerler kaydedilir. 2. excel dosyasında ise aranacak anahtar(TCKN) A Kolonuna taşınır, düşey ara ile 2. excel aranır ve uygun satırlar filtrelenmiş ve ayrı dosyalara parçalanmış olur.");

            MesajWindow.GosterBilgi("1. dosyayı seçiniz. Sorgu Listesi Değerleri");
            var dlg1 = new OpenFileDialog();
            dlg1.Filter = "Excel Dosyaları|*.xls;*.xlsx";
            dlg1.Multiselect = false;
            if (dlg1.ShowDialog() != true)
            {
                return false;
            }

            MesajWindow.GosterBilgi("Şimdi Verilerin olduğu 2. dosyayı seçiniz.\nDikkat !! Sorgulanacak anahtar değerini(TCKN) A Kolonuna taşıyınız.");
            var dlg2 = new OpenFileDialog();
            dlg2.Filter = "Excel Dosyaları|*.xls;*.xlsx";
            dlg2.Multiselect = false;
            if (dlg2.ShowDialog() != true)
            {
                return false;
            }

            var dtMain1 = ExcelYardimcisi.ExcelToDataTable(dlg1.FileName);
            var dtMain2 = ExcelYardimcisi.ExcelToDataTable(dlg2.FileName);

            if (dtMain1.Columns.Count != 1)
            {
                MesajWindow.GosterUyari("1. dosyada sadece A sütununa sorgulanacak değerleri yazınız.");
                return false;
            }
            try
            {
                var folder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\PARCALANANLAR\\";
                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }
                var sorguDegerleri = dtMain1.Rows.Cast<DataRow>().Select(x => x[0].ToString()).Distinct().ToList();
                foreach (var deger in sorguDegerleri)
                {
                    var sonuclar = dtMain2.Rows.Cast<DataRow>().Where(x => x[0].ToString() == deger);
                    var filename = "SONUC_PARCALANDI_" + deger + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

                    ExcelYardimcisi.DataTableToExcel(sonuclar.CopyToDataTable(), (Path.Combine(folder, filename)));
                }
                result = true;
                Process.Start(folder);

            }
            catch (Exception ex)
            {
                MesajWindow.GosterUyari(ex.Message);
                result = false;
            }
            return result;
        }


         
        private string BosluklariTemizle(string metin)
        {
            var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            for (int i = 0; i < liste.Count(); i++)
            {
                liste[i] = liste[i].Trim();
            }
            var temiz = liste.Where(x => !string.IsNullOrWhiteSpace(x.Trim())).ToList();
            return string.Join(Environment.NewLine, temiz.ToArray());
        }
        private string ilkBosluklariTemizle(string metin)
        {
            var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            for (int i = 0; i < liste.Count(); i++)
            {
                var s = liste[i].Split(' ');
                if (s.Length > 0)
                {
                    liste[i] = s[0];
                }
            }
            var temiz = liste.Where(x => !string.IsNullOrWhiteSpace(x.Trim())).ToList();
            return string.Join(Environment.NewLine, temiz.ToArray());
        }

        private string SonBosluklariTemizle(string metin)
        {
            var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            for (int i = 0; i < liste.Count(); i++)
            {
                var s = liste[i].Split(' ');
                if (s.Length > 0)
                {
                    liste[i] = s[s.Length - 1];
                }
            }
            var temiz = liste.Where(x => !string.IsNullOrWhiteSpace(x.Trim())).ToList();
            return string.Join(Environment.NewLine, temiz.ToArray());
        }

        private string BuyukKucukHarfCevir(string metin, string islem)
        {
            var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            for (int i = 0; i < liste.Count(); i++)
            {
                if (islem == "BÜYÜK")
                {
                    liste[i] = CultureInfo.GetCultureInfo("tr-TR").TextInfo.ToUpper(liste[i]);
                }
                else
                if (islem == "KÜÇÜK")
                {
                    liste[i] = CultureInfo.GetCultureInfo("tr-TR").TextInfo.ToLower(liste[i]);
                }
                else
                {
                    liste[i] = CultureInfo.GetCultureInfo("tr-TR").TextInfo.ToTitleCase(liste[i]);
                }
            }
            var temiz = liste.Where(x => !string.IsNullOrWhiteSpace(x.Trim())).ToList();
            return string.Join(Environment.NewLine, temiz.ToArray());
        }

        private string TarihSirala(string metin)
        {
            string format = "MM-dd-yyyy";
            string tur = "yıl";

            var w = new Window();
            w.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            w.Width = 320;
            w.Height = 280;
            w.SizeToContent = SizeToContent.Manual;
            w.WindowStyle = WindowStyle.ToolWindow;
            w.Title = "Parametreleri Seçiniz";
            var comboTarih = new ComboBox();
            comboTarih.Margin = new Thickness(5);
            comboTarih.Items.Add("Tarih Değerinin Formatını Seçiniz...");
            comboTarih.Items.Add("gün-ay-yıl Örn: 30-01-2022");
            comboTarih.Items.Add("yıl-ay-gün Örn: 2022-01-30");
            comboTarih.SelectedIndex = 0;

            var comboTip = new ComboBox();
            comboTip.Margin = new Thickness(5);
            comboTip.Items.Add("Sıralanacak Alanı Seçiniz...");
            comboTip.Items.Add("gün 'e göre sırala");
            comboTip.Items.Add("ay 'a göre sırala");
            comboTip.Items.Add("yıl 'a göre sırala");
            comboTip.SelectedIndex = 0;

            var comboAyirici = new ComboBox();
            comboAyirici.Margin = new Thickness(5);
            comboAyirici.Items.Add("yıl ay gün Arası Ayırıcı Karakteri Seçiniz...");
            comboAyirici.Items.Add("-");
            comboAyirici.Items.Add(".");
            comboAyirici.Items.Add("/");
            comboAyirici.SelectedIndex = 0;

            var comboSirala = new ComboBox();
            comboSirala.Margin = new Thickness(5);
            comboSirala.Items.Add("Sıralama Türünü Seçiniz...");
            comboSirala.Items.Add("Artan Sıralama");
            comboSirala.Items.Add("Azalan Sıralama");
            comboSirala.SelectedIndex = 0;

            var button = new Button();
            button.Content = "TAMAM";
            button.Height = 40;
            button.Background = System.Windows.Media.Brushes.ForestGreen;
            button.Foreground = System.Windows.Media.Brushes.White;
            button.Click += (sender, args) => { w.DialogResult = true; };
            button.Margin = new Thickness(5);
            var stack = new StackPanel();
            stack.VerticalAlignment = VerticalAlignment.Center;
            stack.Margin = new Thickness(10);
            stack.Children.Add(comboTarih);
            stack.Children.Add(comboTip);
            stack.Children.Add(comboAyirici);
            stack.Children.Add(comboSirala);
            stack.Children.Add(button);
            w.Content = stack;
            var sonuc = w.ShowDialog();
            if (sonuc == true && comboTarih.SelectedIndex > 0 && comboTip.SelectedIndex > 0 && comboAyirici.SelectedIndex > 0 && comboSirala.SelectedIndex > 0)
            {
                if (comboTarih.SelectedIndex == 1)
                {
                    format = "dd" + comboAyirici.SelectedValue + "MM" + comboAyirici.SelectedValue + "yyyy";
                }
                else
                {
                    format = "yyyy" + comboAyirici.SelectedValue + "MM" + comboAyirici.SelectedValue + "dd";
                }
                if (comboTip.SelectedIndex == 1)
                {
                    tur = "gün";
                }
                else
                if (comboTip.SelectedIndex == 2)
                {
                    tur = "ay";
                }
                else
                {
                    tur = "yıl";
                }
                var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                var convertList = new List<DateTime>();
                for (int i = 0; i < liste.Count(); i++)
                {
                    string item = liste[i].Trim();
                    if (!string.IsNullOrEmpty(item))
                    {
                        try
                        {
                            var dt = DateTime.ParseExact(liste[i].Trim(), format, CultureInfo.InvariantCulture);
                            convertList.Add(dt);
                        }
                        catch (Exception ex)
                        {
                            MesajWindow.GosterUyari("Listedeki veride hata var !\nTarih formatı uygun olmadığından tarihe çevrilemedi\nLütfen kontrol edip tekrar deneyiniz.\nHata olan satır no: " + (i + 1).ToString());
                            return null;
                        }
                    }
                }

                if (tur == "yıl")
                {
                    if (comboSirala.SelectedIndex == 1)
                    {
                        convertList = convertList.OrderBy(x => x.Year).ThenBy(x => x.Month).ThenBy(x => x.Day).ToList();
                    }
                    else
                    {
                        convertList = convertList.OrderByDescending(x => x.Year).ThenByDescending(x => x.Month).ThenByDescending(x => x.Day).ToList();
                    }
                }
                else
                if (tur == "ay")
                {
                    if (comboSirala.SelectedIndex == 1)
                    {
                        convertList = convertList.OrderBy(x => x.Month).ThenBy(x => x.Day).ThenBy(x => x.Year).ToList();
                    }
                    else
                    {
                        convertList = convertList.OrderByDescending(x => x.Month).ThenByDescending(x => x.Day).ThenByDescending(x => x.Year).ToList();
                    }
                }
                else
                if (tur == "gün")
                {
                    if (comboSirala.SelectedIndex == 1)
                    {
                        convertList = convertList.OrderBy(x => x.Day).ThenBy(x => x.Month).ThenBy(x => x.Year).ToList();
                    }
                    else
                    {
                        convertList = convertList.OrderByDescending(x => x.Day).ThenByDescending(x => x.Month).ThenByDescending(x => x.Year).ToList();
                    }
                }

                return string.Join(Environment.NewLine, convertList.Select(x => x.ToString(format)));
            }

            return null;
        }


        private string SatirDegistir(string metin)
        {
            var onay = MesajWindow.GosterOnay("Her bir satırda belirlenen karakterleri yenisi ile değiştirmek istiyormusunuz?");
            if (onay != MessageBoxResult.Yes)
            {
                return null;
            }

            var eski = BilgiGirisWindow.Show("", "Değiştirilecek eski değeri giriniz.");
            if (eski.MessageBoxResult != MessageBoxResult.OK)
            {
                return null;
            }
            if (string.IsNullOrEmpty(eski.Answer))
            {
                MesajWindow.GosterHata("Değiştirilecek eski değeri giriniz.");
                return null;
            }

            var yeni = BilgiGirisWindow.Show("", "Yeni değeri giriniz.");
            if (yeni.MessageBoxResult != MessageBoxResult.OK)
            {
                return null;
            }
            var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            for (int i = 0; i < liste.Count(); i++)
            {
                liste[i] = liste[i].Replace(eski.Answer, yeni.Answer);
            }
            return string.Join(Environment.NewLine, liste.ToArray());
        }

        private string SatirDegistirKonum(string metin)
        {
            var onay = MesajWindow.GosterOnay("Her bir satırda belirlenen konuma yeni bir karakter eklemek istiyormusunuz?");
            if (onay != MessageBoxResult.Yes)
            {
                return null;
            }

            var konum = BilgiGirisWindow.Show("", "Veri eklenecek konumun harf sıra numarasını giriniz.");
            if (konum.MessageBoxResult != MessageBoxResult.OK)
            {
                return null;
            }
            if (string.IsNullOrWhiteSpace(konum.Answer))
            {
                MesajWindow.GosterHata("Veri eklenecek konumun harf sıra numarasını giriniz.");
                return null;
            }
            try
            {
                Convert.ToInt32(konum.Answer);
            }
            catch
            {
                MesajWindow.GosterHata("Veri eklenecek konumun harf sıra numarasını giriniz.");
                return null;
            }

            var yeni = BilgiGirisWindow.Show("", "Yeni değeri giriniz.");
            if (yeni.MessageBoxResult != MessageBoxResult.OK)
            {
                return null;
            }
            var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            for (int i = 0; i < liste.Count(); i++)
            {
                try
                {
                    liste[i] = liste[i].Insert(Convert.ToInt32(konum.Answer), yeni.Answer);
                }
                catch
                {
                }

            }
            return string.Join(Environment.NewLine, liste.ToArray());
        }

        private string SatirSilKonum(string metin)
        {
            var onay = MesajWindow.GosterOnay("Her bir satırda belirlenen konumdaki verinin silinmesini istiyormusunuz?");
            if (onay != MessageBoxResult.Yes)
            {
                return null;
            }

            var konum = BilgiGirisWindow.Show("", "Veri silinecek konumun harf sıra numarasını giriniz.\nKaçıncı harf silinecek?");
            if (konum.MessageBoxResult != MessageBoxResult.OK)
            {
                return null;
            }
            if (string.IsNullOrWhiteSpace(konum.Answer))
            {
                MesajWindow.GosterHata("Veri silinecek konumun harf sıra numarasını giriniz.");
                return null;
            }
            try
            {
                Convert.ToInt32(konum.Answer);
            }
            catch
            {
                MesajWindow.GosterHata("Veri silinecek konumun harf sıra numarasını giriniz.");
                return null;
            }

            var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            for (int i = 0; i < liste.Count(); i++)
            {
                try
                {
                    liste[i] = liste[i].Remove(Convert.ToInt32(konum.Answer) - 1, 1);
                }
                catch
                {
                }

            }
            return string.Join(Environment.NewLine, liste.ToArray());
        }

        private string SatirSilKarakterSonrasi(string metin)
        {
            var onay = MesajWindow.GosterOnay("Her bir satırda belirlenen karakterin öncesindeki yada sonrasındaki metin silinir.\nÖncesindeki metnin silinmesi için EVET'e sonrasındaki verilerin silinmesi için HAYIR'a basınız.");

            var karakter = BilgiGirisWindow.Show("", "Hangi karakter sonrasındaki / öncesindeki\nveri silinecek?");
            if (karakter.MessageBoxResult != MessageBoxResult.OK)
            {
                return null;
            }
            if (string.IsNullOrEmpty(karakter.Answer))
            {
                MesajWindow.GosterHata("Sonrasındaki / Öncesindeki Veri silinecek\niçin karakter belirleyiniz. Örn: -");
                return null;
            }

            var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            for (int i = 0; i < liste.Count(); i++)
            {
                try
                {
                    if (onay == MessageBoxResult.Yes)
                    {
                        liste[i] = liste[i].Remove(0, liste[i].IndexOf(karakter.Answer) + karakter.Answer.Length);
                    }
                    else
                    {
                        liste[i] = liste[i].Remove(liste[i].IndexOf(karakter.Answer), liste[i].Length - liste[i].IndexOf(karakter.Answer));
                    }
                }
                catch
                {
                }
            }
            return string.Join(Environment.NewLine, liste.ToArray());
        }


        private string BelliUzunlukOlmayanlariTemizle(string metin)
        {
            var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            var bilgi = BilgiGirisWindow.Show("10", "Satır harf sayısını giriniz.");
            if (bilgi.MessageBoxResult != MessageBoxResult.OK)
            {
                return null;
            }
            if (string.IsNullOrWhiteSpace(bilgi.Answer))
            {
                MesajWindow.GosterHata("Satır harf sayısını giriniz.");
                return null;
            }
            int adet = 0;
            try
            {
                adet = Convert.ToInt32(bilgi.Answer);
                var temiz = liste.Where(x => x.Length == adet).ToList();
                return string.Join(Environment.NewLine, temiz.ToArray());
            }
            catch
            {
                MesajWindow.GosterHata("Satır harf sayısını RAKAM olarak giriniz.");
                return null;
            }
        }

        private string SatirdakiSayilariGetir(string metin)
        {
            MesajWindow.GosterBilgi("Metin önünde-arkasında BOŞLUK veya TAB karakteri olanlar ayrıştırılacak ve sayı ise listeye eklenecektir.");
            return ListeYardimcisi.SatirdakiSayilariGetir(metin);
        }

        private string ListeOrtakKesisimler()
        {
            MesajWindow.GosterBilgi("Toplu halde seçilen TEXT listelerinin\ntamamında geçen ortak değerleri bulur.\nLütfen karşılaştıracağınız listeleri seçiniz.");

            //Toplu metin listesi içinde ortakları bul            
            var allListe = new List<Tuple<string, string>>();
            var dlg = new OpenFileDialog();
            dlg.Multiselect = true;
            dlg.Filter = "Metin Dosyaları|*.txt";
            if (dlg.ShowDialog() == true)
            {
                var bilgi = BilgiGirisWindow.Show(dlg.FileNames.Count().ToString(), "Toplamda " + dlg.FileNames.Count() + " adet liste seçtiniz.\nOrtak değerler en az kaç liste içerisinde geçsin?");
                if (bilgi.MessageBoxResult != MessageBoxResult.OK)
                {
                    return null;
                }
                if (string.IsNullOrWhiteSpace(bilgi.Answer))
                {
                    MesajWindow.GosterHata("Ortak değerler en az kaç liste içerisinde geçsin değerini girmediniz.");
                    return null;
                }
                try
                {
                    var dosyaAdiGosterilsin = MesajWindow.GosterOnay("Bulunan sonuçların yanına hangi dosyada olduğu yazılsın mı?");
                    foreach (var itemFileName in dlg.FileNames)
                    {
                        var l = File.ReadAllLines(itemFileName).ToList();
                        for (int i = 0; i < l.Count; i++)
                        {
                            l[i] = l[i].Trim();
                        }
                        var dist = l.Distinct().ToList();
                        string item1 = Path.GetFileName(itemFileName);
                        foreach (var itemdist in dist)
                        {
                            allListe.Add(Tuple.Create<string, string>(item1, itemdist));
                        }
                    }
                    int adet = 0;
                    adet = Convert.ToInt32(bilgi.Answer);
                    if (adet > 0)
                    {
                        var sonuc = allListe.Select(x => x.Item2).GroupBy(x => x).Select(y => new { y.Key, Count = y.Count() }).Where(z => z.Count >= adet).Select(l => l.Key).OrderBy(k => k).ToList();

                        if (sonuc != null && sonuc.Count > 0)
                        {
                            if (dosyaAdiGosterilsin == MessageBoxResult.Yes)
                            {
                                for (int i = 0; i < sonuc.Count; i++)
                                {
                                    var dosyalar = allListe.Where(x => x.Item2 == sonuc[i]).Select(x => x.Item1);
                                    var d = sonuc[i] + " > (" + string.Join(",", dosyalar.ToArray()) + ")";
                                    sonuc[i] = d;
                                }
                            }
                            return string.Join(Environment.NewLine, sonuc.ToArray());
                        }
                    }
                    else
                    {
                        MesajWindow.GosterHata("Ortak değerler en az kaç liste içerisinde geçsin değeri sıfırdan büyük olmalıdır.");
                        return null;
                    }

                }
                catch
                {
                    MesajWindow.GosterHata("Ortak değerler en az kaç liste içerisinde geçsin değeri sıfırdan büyük olmalıdır.");
                    return null;
                }

            }
            return "";

        }

        private string ListeOrtakKesisimlerXLS()
        {
            MesajWindow.GosterBilgi("Toplu halde seçilen EXCEL dosyalarının belirlenen kolonundaki veriler alınır seçilen dosyaların tamamında geçen ortak değerleri bulur.\nLütfen karşılaştıracağınız EXCEL dosyalarını seçiniz.");

            //var listAll = new List<string>();
            var listAll = new List<Tuple<string, string>>();
            int adetGecsin = 0;
            var dlg = new OpenFileDialog();
            dlg.Filter = "Excel Dosyaları|*.xls;*.xlsx";
            dlg.Multiselect = true;

            if (dlg.ShowDialog() == true)
            {
                var bilgiAdet = BilgiGirisWindow.Show(dlg.FileNames.Count().ToString(), "Toplamda " + dlg.FileNames.Count() + " adet liste seçtiniz.\nOrtak değerler en az kaç dosya içerisinde geçsin?");
                if (bilgiAdet.MessageBoxResult != MessageBoxResult.OK)
                {
                    return null;
                }
                if (string.IsNullOrWhiteSpace(bilgiAdet.Answer))
                {
                    MesajWindow.GosterHata("Ortak değerler en az kaç liste içerisinde geçsin değerini girmediniz.");
                    return null;
                }

                adetGecsin = Convert.ToInt32(bilgiAdet.Answer);
                if (adetGecsin < 1)
                {
                    MesajWindow.GosterHata("Ortak değerler en az kaç liste içerisinde geçsin değeri sıfırdan büyük olmalıdır.");
                    return null;
                }

                var bilgi = BilgiGirisWindow.Show("1", "Veri Alınacak Kolon Sırasını Giriniz.\nörn: A kolonu için 1 giriniz.");
                if (bilgi.MessageBoxResult == MessageBoxResult.OK)
                {
                    var secilenKolon = 1;
                    try
                    {
                        secilenKolon = Convert.ToInt32(bilgi.Answer.Trim());
                    }
                    catch (Exception ex)
                    {
                        MesajWindow.GosterHata("Girilen bilgi istenilen formatta değil !" + Environment.NewLine + ex.Message);
                        return null;
                    }

                    var dosyaAdiGosterilsin = MesajWindow.GosterOnay("Bulunan sonuçların yanına hangi dosyada olduğu yazılsın mı?");

                    foreach (var itemFileName in dlg.FileNames)
                    {
                        try
                        {
                            var dt = ExcelYardimcisi.ExcelToDataTable(itemFileName, columnNo: secilenKolon);
                            if (dt == null || dt.Columns.Count < 1)
                            {
                                MesajWindow.GosterHata(itemFileName + "\nExcel dosyasının üst bilgileri okunamadı.\nExcel dosyası bozuk olabilir. Yeniden oluşturunuz.");
                                return null;
                            }

                            var dist = dt.Rows.Cast<DataRow>().Where(x => x[0] != null).Select(x => x[0].ToString().Trim()).Distinct().ToList();
                            string item1 = Path.GetFileName(itemFileName);
                            foreach (var itemdist in dist)
                            {
                                listAll.Add(Tuple.Create<string, string>(item1, itemdist));
                            }

                        }
                        catch (Exception ex)
                        {
                            MesajWindow.GosterHata("Kolon numarası için sadece rakam giriniz.\n" + ex.Message);
                            return null;
                        }
                    }
                    if (listAll != null && listAll.Count > 0)
                    {
                        var sonuc = listAll.Select(x => x.Item2).GroupBy(x => x).Select(y => new { y.Key, Count = y.Count() }).Where(z => z.Count >= adetGecsin).Select(l => l.Key).OrderBy(k => k).ToList();
                        if (sonuc != null)
                        {
                            if (dosyaAdiGosterilsin == MessageBoxResult.Yes)
                            {
                                for (int i = 0; i < sonuc.Count; i++)
                                {
                                    var dosyalar = listAll.Where(x => x.Item2 == sonuc[i]).Select(x => x.Item1);
                                    var d = sonuc[i] + " > (" + string.Join(",", dosyalar.ToArray()) + ")";
                                    sonuc[i] = d;
                                }
                            }

                            return string.Join<string>(Environment.NewLine, sonuc.ToArray());
                        }
                    }

                }
            }

            return null;
        }


        private string ListeDosyaBirlestir()
        {
            MesajWindow.GosterBilgi("Toplu halde seçilen TEXT listelerini birleştirir\nLütfen birleştireceğiniz dosyaları seçiniz.");

            var SonListe = new List<string>();
            var dlg = new OpenFileDialog();
            dlg.Multiselect = true;
            dlg.Filter = "Metin Dosyaları|*.txt";
            if (dlg.ShowDialog() == true)
            {
                foreach (var dosya in dlg.FileNames)
                {
                    var l = File.ReadAllLines(dosya).ToList();
                    SonListe.AddRange(l);
                }
                return string.Join(Environment.NewLine, SonListe.ToArray());
            }
            return "";
        }

        private string ListeBirlestir(string ilkListe, string ikinciListe)
        {
            var liste1 = ilkListe.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            var liste2 = ikinciListe.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            if (liste1.Count() != liste2.Count())
            {
                MesajWindow.GosterUyari("Seçilen ilk ve ikinci listenin satır sayıları farklıdır\nİlk listenin sonlarında eksiklik olabilir.");
            }
            var SonListe = new List<string>();


            for (int i = 0; i < liste1.Count(); i++)
            {
                if (i < liste2.Count())
                {
                    SonListe.Add(liste1[i] + liste2[i]);
                }
                else
                {
                    SonListe.Add(liste1[i]);
                }
            }

            return string.Join(Environment.NewLine, SonListe.ToArray());
        }

        private string KlsordekiDosyalar()
        {
            MesajWindow.GosterBilgi("Seçilen klasördeki Tüm dosyaların adlarını Liste olarak getirir");
            var SonListe = new List<string>();
            var dlg = new System.Windows.Forms.FolderBrowserDialog();

            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(dlg.SelectedPath))
            {
                var cevap = MesajWindow.GosterOnay("Sadece dosya adı mı getirilsin?");
                if (cevap == MessageBoxResult.Yes)
                {
                    foreach (var item in Directory.GetFiles(dlg.SelectedPath))
                    {
                        SonListe.Add(Path.GetFileName(item));
                    }
                }
                else
                {
                    foreach (var item in Directory.GetFiles(dlg.SelectedPath))
                    {
                        SonListe.Add(item);
                    }
                }

                return string.Join(Environment.NewLine, SonListe);
            }
            else
            {
                return null;
            }
        }

        public async void DosyaHASHListesiGetir()
        {
            var dlg = new OpenFileDialog();
            dlg.Filter = "Tüm Dosyalar|*.*";
            dlg.Multiselect = true;

            if (dlg.ShowDialog() == true)
            {
                busyIndicator.IsBusy = true;
                busyIndicator.BusyContent = "Lütfen bekleyin...";

                var liste = new DataTable("liste");
                liste.Columns.Add("Dosya Adı");
                liste.Columns.Add("HASH Değeri (MD5)");
                liste.Columns.Add("HASH Değeri (SHA1)");
                liste.Columns.Add("Dosya Boyutu (byte)");
                liste.Columns.Add("Değiştirme Tarihi");
                liste.Columns.Add("Dosya Yolu");

                var dosyalar = dlg.FileNames;
                var cevap = MesajWindow.GosterOnay("Alt klasörlerdeki tüm dosyalar taransın mı?");
                if (cevap == MessageBoxResult.Yes)
                {
                    dosyalar = Directory.GetFiles(Path.GetDirectoryName(dosyalar.First()), "*.*", SearchOption.AllDirectories);
                }

                if (dosyalar.Length > 100)
                {
                    var onay = MesajWindow.GosterOnay(dosyalar.Length + " Adet dosyanın HASH değeri hesaplanacak devam etmek istiyor musunuz?");
                    if (onay != MessageBoxResult.Yes)
                    {
                        busyIndicator.IsBusy = false;
                        return;
                    }
                }

                int i = 0;
                foreach (var item in dosyalar)
                {
                    i++;
                    await busyIndicator.Dispatcher.BeginInvoke(new Action(() => { busyIndicator.BusyContent = "Lütfen bekleyin..." + Environment.NewLine + i.ToString() + "/" + dosyalar.Length.ToString(); }));
                    var info = new FileInfo(item);

                    try
                    {
                        var dosyaByte = File.ReadAllBytes(item);
                        var hash = HashYardimcisi.DosyaHashAl(dosyaByte);
                        var sha1 = HashYardimcisi.DosyaSHA1HashAl(dosyaByte);
                        if (string.IsNullOrEmpty(hash))
                        {
                            hash = "HASH Değeri Hesaplanamadı.";
                        }
                        if (string.IsNullOrEmpty(sha1))
                        {
                            sha1 = "SHA1 Değeri Hesaplanamadı.";
                        }
                        liste.Rows.Add(Path.GetFileName(item), hash, sha1, info.Length, info.LastWriteTime, item);
                    }
                    catch (Exception)
                    {
                        liste.Rows.Add(Path.GetFileName(item), "DOSYA KULLANIMDA", "DOSYA KULLANIMDA", info.Length, info.LastWriteTime, item);
                    }
                    await Task.Delay(1);
                }


                busyIndicator.BusyContent = "Lütfen bekleyin...";
                busyIndicator.IsBusy = false;

                MesajWindow.GosterListePencere(liste.Rows.Cast<DataRow>().ToList(), "Dosyaların HASH Bilgileri", true);
            }

        }

        private string GunleriGetir()
        {
            var sonListe = new StringBuilder();
            sonListe.AppendLine("Pazartesi");
            sonListe.AppendLine("Salı");
            sonListe.AppendLine("Çarşamba");
            sonListe.AppendLine("Perşembe");
            sonListe.AppendLine("Cuma");
            sonListe.AppendLine("Cumartesi");
            sonListe.AppendLine("Pazar");
            return sonListe.ToString();
        }
        private string AylariGetir()
        {
            var sonListe = new StringBuilder();
            sonListe.AppendLine("Ocak");
            sonListe.AppendLine("Şubat");
            sonListe.AppendLine("Mart");
            sonListe.AppendLine("Nisan");
            sonListe.AppendLine("Mayıs");
            sonListe.AppendLine("Haziran");
            sonListe.AppendLine("Temmuz");
            sonListe.AppendLine("Ağustos");
            sonListe.AppendLine("Eylül");
            sonListe.AppendLine("Ekim");
            sonListe.AppendLine("Kasım");
            sonListe.AppendLine("Aralık");
            return sonListe.ToString();
        }

        private string OtomatikSayiOlustur()
        {
            var sonListe = new StringBuilder();
            int sayi1 = 0;
            int sayi2 = 0;
            string yil1 = new DateTime(DateTime.Now.Year, 1, 1).ToString("yyyyMM");
            string yil2 = new DateTime(DateTime.Now.Year, 12, 31).ToString("yyyyMM");
            var bilgi1 = BilgiGirisWindow.Show(yil1, "Başlangıç Değerini Sayı Olarak Giriniz?");
            if (bilgi1.MessageBoxResult != MessageBoxResult.OK)
            {
                return null;
            }
            if (string.IsNullOrEmpty(bilgi1.Answer.Trim()))
            {
                MesajWindow.GosterHata("Lütfen Başlangıç Değerini Sayı Olarak Giriniz");
                return null;
            }
            else
            {
                try
                {
                    sayi1 = Convert.ToInt32(bilgi1.Answer.Trim());
                }
                catch (Exception)
                {
                    MesajWindow.GosterHata("Lütfen Başlangıç Değerini Sayı Olarak Giriniz");
                    return null;
                }
            }

            var bilgi2 = BilgiGirisWindow.Show(yil2, "Bitiş Değerini Sayı Olarak Giriniz");
            if (bilgi2.MessageBoxResult != MessageBoxResult.OK)
            {
                return null;
            }
            if (string.IsNullOrEmpty(bilgi2.Answer.Trim()))
            {
                MesajWindow.GosterHata("Lütfen Bitiş Değerini Sayı Olarak Giriniz");
                return null;
            }
            else
            {
                try
                {
                    sayi2 = Convert.ToInt32(bilgi2.Answer.Trim());
                }
                catch (Exception)
                {
                    MesajWindow.GosterHata("Lütfen Bitiş Değerini Sayı Olarak Giriniz");
                    return null;
                }
            }

            if (sayi2 > sayi1)
            {
                for (int i = sayi1; i <= sayi2; i++)
                {
                    sonListe.AppendLine(i.ToString());
                }
            }
            else
            {
                MesajWindow.GosterHata("Lütfen Başlangıç ve Bitiş Değerlerini Kontrol Ediniz.\nBitiş Değeri Başlangıçtan büyük olmalıdır.");
                return null;
            }
            return sonListe.ToString();
        }
        private string SayiUzunluguTamamla(string metin)
        {
            var sonListe = new StringBuilder();
            int sayi1 = 0;
            var bilgi1 = BilgiGirisWindow.Show("3", "Bu değerden kısa olan sayıların başına\n(0) SIFIR koyarak belli uzunluğa tamamla");
            if (bilgi1.MessageBoxResult != MessageBoxResult.OK)
            {
                return null;
            }
            if (string.IsNullOrEmpty(bilgi1.Answer.Trim()))
            {
                MesajWindow.GosterHata("Lütfen Değeri Sayı Olarak Giriniz");
                return null;
            }
            else
            {
                try
                {
                    sayi1 = Convert.ToInt32(bilgi1.Answer.Trim());
                }
                catch (Exception)
                {
                    MesajWindow.GosterHata("Lütfen Değeri Sayı Olarak Giriniz");
                    return null;
                }
            }
            var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            foreach (var item in liste)
            {
                sonListe.AppendLine(item.Trim().PadLeft(sayi1, '0'));
            }

            return sonListe.ToString();
        }

        private string ListeSablonEkle(string metin)
        {
            var bilgi = BilgiGirisWindow.Show("Sorgu Sonucunda {x} TC Kimlik Numaralı", "Şablonu düzeltiniz\nDikkat ! {x} bölümünü silmeyiniz.");
            if (bilgi.MessageBoxResult != MessageBoxResult.OK)
            {
                return null;
            }
            if (string.IsNullOrEmpty(bilgi.Answer) || !bilgi.Answer.Contains("{x}"))
            {
                MesajWindow.GosterHata("Dikkat ! {x} bölümünü silmeyiniz.");
                return null;
            }

            var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            var SonListe = new List<string>();
            foreach (var item in liste)
            {
                string ekle = bilgi.Answer.Replace("{x}", item);
                SonListe.Add(ekle);
            }
            return string.Join(Environment.NewLine, SonListe.ToArray());
        }

        private string SatiraMetinEkle(string metin)
        {
            var bilgi = BilgiGirisWindow.Show("", "Ekleme yapmak istediğiniz metini giriniz.");
            if (bilgi.MessageBoxResult != MessageBoxResult.OK)
            {
                return null;
            }
            if (string.IsNullOrEmpty(bilgi.Answer))
            {
                MesajWindow.GosterHata("Eklenecek metin girmediniz.");
                return null;
            }

            var soru = MessageBox.Show("Metni satırın önüne eklemek için EVET'e\nMetni satırın sonuna eklemek için HAYIR'a basınız", "Metin Önüne mi Yoksa Sonuna Mı Eklensin?", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
            if (soru == MessageBoxResult.Cancel)
            {
                return null;
            }

            var result = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(metin))
            {
                var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                for (int i = 0; i < liste.Count(); i++)
                {
                    var s = liste[i].Trim();
                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        if (soru == MessageBoxResult.Yes)
                        {
                            result.AppendLine(bilgi.Answer + s);
                        }
                        else
                        if (soru == MessageBoxResult.No)
                        {
                            result.AppendLine(s + bilgi.Answer);
                        }
                    }
                }
            }
            return result.ToString();
        }

        private string SatirdakiMetinBaslarBiterIcerir(string metin, int BaslarBiterIcerir)
        {
            var bilgi = BilgiGirisWindow.Show("", "Kontrol edilecek metni giriniz.");
            if (bilgi.MessageBoxResult != MessageBoxResult.OK)
            {
                return null;
            }
            if (string.IsNullOrWhiteSpace(bilgi.Answer))
            {
                MesajWindow.GosterHata("Kontrol edilecek metin girmediniz.");
                return null;
            }

            if (!string.IsNullOrWhiteSpace(metin))
            {
                var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                if (BaslarBiterIcerir == 0)
                {
                    var temiz = liste.Where(x => x.StartsWith(bilgi.Answer)).ToList();
                    return string.Join(Environment.NewLine, temiz.ToArray());
                }
                else
                if (BaslarBiterIcerir == 1)
                {
                    var temiz = liste.Where(x => x.EndsWith(bilgi.Answer)).ToList();
                    return string.Join(Environment.NewLine, temiz.ToArray());
                }
                else
                {
                    var temiz = liste.Where(x => x.Contains(bilgi.Answer)).ToList();
                    return string.Join(Environment.NewLine, temiz.ToArray());
                }
            }
            else
            {
                return null;
            }
        }

        private string ListelerdeAramaYap()
        {
            var sb = new StringBuilder();
            MesajWindow.GosterBilgi("İçerisinde arama yapacağınız metin listelerini şeçiniz.");
            var dlg = new OpenFileDialog();
            dlg.Multiselect = true;
            dlg.Filter = "*.TXT|*.TXT";
            if (dlg.ShowDialog() == true)
            {
                var bilgi = BilgiGirisWindow.Show("", "Aranacak değeri giriniz.");
                if (bilgi.MessageBoxResult != MessageBoxResult.OK)
                {
                    return null;
                }
                if (string.IsNullOrWhiteSpace(bilgi.Answer))
                {
                    MesajWindow.GosterHata("Kontrol edilecek metin girmediniz.");
                    return null;
                }

                foreach (var item in dlg.FileNames)
                {
                    var gelen = File.ReadAllLines(item, Encoding.UTF8);
                    for (int i = 0; i < gelen.Length; i++)
                    {
                        if (gelen[i].Contains(bilgi.Answer))
                        {
                            sb.AppendLine(Path.GetFileName(item) + " > SatırNo: " + (i + 1));
                        }
                    }
                }

                return sb.ToString();
            }
            else
            {
                return null;
            }

        }


        private string ListedenBenzerleriniGetir(string metin)
        {
            var bilgiAranacak = BilgiGirisWindow.Show("", "Aranacak metni giriniz !");
            if (bilgiAranacak.MessageBoxResult != MessageBoxResult.OK)
            {
                return null;
            }
            if (string.IsNullOrWhiteSpace(bilgiAranacak.Answer))
            {
                MesajWindow.GosterHata("Aranacak metni girmediniz");
                return null;
            }

            var bilgiFark = BilgiGirisWindow.Show("1", "En fazla kaç fark olabilir");
            if (bilgiFark.MessageBoxResult != MessageBoxResult.OK)
            {
                return null;
            }
            if (string.IsNullOrWhiteSpace(bilgiFark.Answer))
            {
                MesajWindow.GosterHata("Fark girmediniz");
                return null;
            }

            if (!string.IsNullOrWhiteSpace(metin))
            {
                var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                var temiz = ListeYardimcisi.BenzerleriniGetir(liste, bilgiAranacak.Answer, Convert.ToInt32(bilgiFark.Answer));
                return temiz.Trim();
            }
            else
            {
                return null;
            }
        }


        private string SatirGrupla(string metin)
        {
            var bilgi = BilgiGirisWindow.Show("3", "Kaçar adet satır tek bir satıra gruplanacak !\nÖrn 3 satır adet 1 satır olsun");
            if (bilgi.MessageBoxResult != MessageBoxResult.OK)
            {
                return null;
            }
            if (string.IsNullOrWhiteSpace(bilgi.Answer))
            {
                MesajWindow.GosterHata("Gruplanacak satır adeti girmediniz.");
                return null;
            }

            int adet = 0;
            try
            {
                adet = Convert.ToInt32(bilgi.Answer);
            }
            catch (Exception)
            {
            }

            if (!string.IsNullOrWhiteSpace(metin))
            {
                var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                string ekle = "";
                var sbSonuc = new StringBuilder();
                for (int i = 0; i < liste.Count(); i++)
                {
                    if (((i + 1) % adet) != 0)
                    {
                        ekle += liste[i];
                    }
                    else
                    {
                        ekle += liste[i];
                        sbSonuc.AppendLine(ekle);
                        ekle = "";
                    }
                }
                return sbSonuc.ToString();
            }
            else
            {
                return null;
            }
        }

        private void KlasorDosyaGrupla()
        {
            MesajWindow.GosterBilgi("Dosyaların bulunduğu klasörü seçiniz.\nKlasör içerisindeki dosyalar;\ndosya adındaki belli karakterden öncesine göre\nklasörlere gruplanacak.");

            var dlg = new System.Windows.Forms.FolderBrowserDialog();

            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(dlg.SelectedPath))
            {
                var bilgi = BilgiGirisWindow.Show("11", "Dosya adındaki;\nHarf sayısına göre gruplamak için sayı giriniz!\n" +
                    "Örn: TCKimlik No için 11 giriniz\n" +
                    "Veya dosya adında geçen bir sembol girebilirsiniz!\n" +
                    "Örn: alt tire _ işareti veya parantez ( işareti gibi)");
                if (bilgi.MessageBoxResult != MessageBoxResult.OK)
                {
                    return;
                }
                if (string.IsNullOrWhiteSpace(bilgi.Answer))
                {
                    MesajWindow.GosterHata("Gruplamak için bir değer girmediniz.");
                    return;
                }



                int harfSayisi = 11;
                try
                {
                    harfSayisi = Convert.ToInt32(bilgi.Answer);
                }
                catch
                {
                    harfSayisi = 0;
                }

                var files = Directory.GetFiles(dlg.SelectedPath, "*.*");
                foreach (var element in files)
                {
                    var fileName = Path.GetFileName(element);
                    try
                    {
                        string klasorAdi = string.Empty;
                        if (harfSayisi != 0)
                        {
                            klasorAdi = fileName.Substring(0, harfSayisi).Trim();
                        }
                        else
                        {
                            klasorAdi = fileName.Substring(0, fileName.IndexOf(bilgi.Answer)).Trim();
                        }
                        if (!string.IsNullOrEmpty(klasorAdi))
                        {
                            var klasorYolu = Path.Combine(dlg.SelectedPath, klasorAdi);
                            if (!Directory.Exists(klasorYolu))
                            {
                                Directory.CreateDirectory(klasorYolu);
                            }
                            File.Move(element, Path.Combine(klasorYolu, fileName));
                        }

                    }
                    catch (Exception ex)
                    {
                    }
                }
            }

        }
         
        private string ListeEnUzunSatirlariBul(string metin, int adet)
        {
            var result = string.Empty;
            var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            if (liste.Count() > 0)
            {
                var bulunan = liste.Select(x => new { satir = x, boyut = x.Length }).OrderByDescending(y => y.boyut).Take(3).Select(y => (y.satir + " (" + y.boyut + ")"));
                result = string.Join("\n", bulunan);
            }
            return result;
        }

        private string ListeEnKisaSatirlariBul(string metin, int adet)
        {
            var result = string.Empty;
            var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            if (liste.Count() > 0)
            {
                var bulunan = liste.Select(x => new { satir = x, boyut = x.Length }).OrderBy(y => y.boyut).Take(3).Select(y => (y.satir + " (" + y.boyut + ")"));
                result = string.Join("\n", bulunan);
            }
            return result;
        }

        private string TextSHA1HashListesiGetir(string metin)
        {
            var result = string.Empty;
            var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            if (liste.Count() > 0)
            {
                foreach (var item in liste)
                {
                    if (!string.IsNullOrEmpty(item))
                    {
                        result += HashYardimcisi.GenerateTextSHA1(item) + Environment.NewLine;
                    }
                }
            }
            return result;
        }

        private string TextMD5HashListesiGetir(string metin)
        {
            var result = string.Empty;
            var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            if (liste.Count() > 0)
            {
                foreach (var item in liste)
                {
                    if (!string.IsNullOrEmpty(item))
                    {
                        result += HashYardimcisi.GenerateTextMD5(item) + Environment.NewLine;
                    }
                }
            }
            return result;
        }

        private string ExcelDosyasiAc(string mesaj, List<string> excelSutunlari = null)
        {
            string result = "";
            var dlg1 = new OpenFileDialog();
            dlg1.Filter = "Excel Dosyaları|*.xls;*.xlsx";
            dlg1.Multiselect = false;
            if (excelSutunlari == null || excelSutunlari.Count <= 0)
            {
                MesajWindow.GosterUyari(mesaj);
                if (dlg1.ShowDialog() == true)
                {
                    result = dlg1.FileName;
                }
            }
            else
            {
                var cevap = MesajWindow.GosterOnay(mesaj);
                if (cevap == MessageBoxResult.Yes)
                {
                    if (dlg1.ShowDialog() == true)
                    {
                        result = dlg1.FileName;
                    }
                }
                else
                {
                    string fileName = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\ÖrnekExcelDosyası_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
                    using (ExcelPackage xp = new ExcelPackage(new System.IO.FileInfo(fileName)))
                    {
                        ExcelWorksheet ws = xp.Workbook.Worksheets.Add("Sayfa1");
                        var table = new DataTable("table");
                        foreach (var item in excelSutunlari)
                        {
                            table.Columns.Add(item);
                        }
                        ws.Cells["A1"].LoadFromDataTable(table, true);
                        xp.Save();
                        Process.Start(new ProcessStartInfo(fileName) { UseShellExecute = true });
                    }
                }
            }
            return result;
        }


        public void ExcelKoordinatMesafeHesapla()
        {
            try
            {
                MesajWindow.GosterBilgi("Seçilen iki adet excel dosyasındaki koordinat bilgisinden (enlem,boylam) aralarındaki mesafeleri hesaplar.");
                var dosya1 = ExcelDosyasiAc("A1 sütununda Koordinat Adı,\nB1 sütununda Enlem\nC1 sütununda Boylam\nbilgilerinin olduğu 1. Excel dosyanız hazır mı?", new List<string>() { "KOORDİNAT_ADI", "ENLEM", "BOYLAM" });
                var dosya2 = ExcelDosyasiAc("Şimdi yine aynı şekilde\nA1 sütununda KoordinatAdı\nB1 sütununda Enlem\nC1 sütununda Boylam\nbilgilerinin olduğu 2. Excel dosyanız hazır mı?", new List<string>() { "KOORDİNAT_ADI", "ENLEM", "BOYLAM" });

                var satirlar1 = ListeYardimcisi.ExcelSatirlariniGetir(dosya1);
                var satirlar2 = ListeYardimcisi.ExcelSatirlariniGetir(dosya2);

                if (satirlar1 != null && satirlar1.Count > 0 && satirlar2 != null && satirlar2.Count > 0)
                {

                    var sonuc = new List<KoordinatAnalizMesafeSonuc>();
                    Task.Run(() => { sonuc = ListeYardimcisi.KoordinatListesiMesafeBul(satirlar1, satirlar2); }).Wait();

                    MesajWindow.GosterListePencere(sonuc, "Kordinatlar Arası Mesafeler >>> " + dosya1 + " --- " + dosya2, true);
                }
                else
                {
                    MesajWindow.GosterUyari("HATA ! Dosyalarınızı kontrol ediniz.");
                }

            }
            catch (Exception)
            {
            }
        }

        public void FotoKoordinatListele()
        {
            MesajWindow.GosterBilgi("Seçilen fotoğraflardaki EXIF GPS Koordinat verisini bularak\nliste şeklinde getirir.");
            try
            {
                var dlg = new OpenFileDialog();
                dlg.Filter = "JPG Resim Dosyaları|*.jpg";
                dlg.Multiselect = true;

                if (dlg.ShowDialog() == true)
                {
                    var liste = new DataTable("liste");
                    liste.Columns.Add("Dosya Adı");
                    liste.Columns.Add("X (Enlem)");
                    liste.Columns.Add("Y (Boylam)");
                    liste.Columns.Add("Dosya Yolu");

                    var dosyalar = dlg.FileNames;
                    var cevap = MesajWindow.GosterOnay("Alt klasörlerdeki tüm dosyalar da taransın mı?");
                    if (cevap == MessageBoxResult.Yes)
                    {
                        dosyalar = Directory.GetFiles(Path.GetDirectoryName(dosyalar.First()), "*.jpg", SearchOption.AllDirectories);
                    }

                    foreach (var item in dosyalar)
                    {
                        var exif = new ExifData(item);
                        GeoCoordinateExif lat;
                        GeoCoordinateExif lon;
                        decimal yukseklik;
                        exif.GetGpsAltitude(out yukseklik);
                        exif.GetGpsLatitude(out lat);
                        exif.GetGpsLongitude(out lon);
                        if (lat.Degree > 0 && lon.Degree > 0)
                        {
                            liste.Rows.Add(Path.GetFileName(item), GeoCoordinateExif.ToDecimal(lat).ToString("F6"), GeoCoordinateExif.ToDecimal(lon).ToString("F6"), item);
                        }
                        else
                        {
                            liste.Rows.Add(Path.GetFileName(item), "GPS YOK", "GPS YOK", item);
                        }
                    }

                    MesajWindow.GosterListePencere(liste.Rows.Cast<DataRow>().ToList(), "Fotoğraf GPS Koordinat Verileri", true);

                }

            }
            catch (Exception)
            {
            }

        }

         


        private bool ExcelSatirBirlestir()
        {
            var result = false;
            MesajWindow.GosterBilgi("Seçilen bir excel dosyasının belirtilen ID alanına göre\naynı dosyanın diğer kolonu altındaki satırları tek bir hücrede gruplar.");

            MesajWindow.GosterBilgi("Şimdi verilerin olduğu excel dosyasını seçiniz.");
            var dlg1 = new OpenFileDialog();
            dlg1.Filter = "Excel Dosyaları|*.xls;*.xlsx";
            dlg1.Multiselect = false;
            if (dlg1.ShowDialog() != true)
            {
                return false;
            }

            MesajWindow.GosterBilgi("Şimdi sonuçların kaydedileceği dosyayı seçiniz.");
            var dlg2 = new SaveFileDialog();
            dlg2.Filter = "Excel Dosyası (*.xlsx)|*.xlsx";
            dlg2.FileName = "SONUC_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            if (dlg2.ShowDialog() != true)
            {
                return false;
            }

            int idSira = -1;
            var bilgi = BilgiGirisWindow.Show("1", "ID (Anahtar) Verinin Olduğu Kolon Sırasını Giriniz.\nörn: A kolonu için 1 giriniz.");
            if (bilgi.MessageBoxResult == MessageBoxResult.OK)
            {
                try
                {
                    idSira = Convert.ToInt32(bilgi.Answer.Trim());
                }
                catch (Exception ex)
                {
                    MesajWindow.GosterHata("Girilen bilgi istenilen formatta değil !" + Environment.NewLine + ex.Message);
                    return false;
                }
            }
            else
            {
                return false;
            }

            int icerikSira = -1;
            var bilgi2 = BilgiGirisWindow.Show("2", "Birleştirilecek Verinin Kolon Sırasını Giriniz.\nörn: B kolonu için 2 giriniz.");
            if (bilgi2.MessageBoxResult == MessageBoxResult.OK)
            {
                try
                {
                    icerikSira = Convert.ToInt32(bilgi2.Answer.Trim());
                }
                catch (Exception ex)
                {
                    MesajWindow.GosterHata("Girilen bilgi istenilen formatta değil !" + Environment.NewLine + ex.Message);
                    return false;
                }
            }
            else
            {
                return false;
            }

            var birlesimTuru = MesajWindow.GosterOnay("Hücrede birleştirilecek verilerin yan yana yazılması için EVET'e, alt alta yazılması için HAYIR'a basınız.");
            {
                var dtMain = ExcelYardimcisi.ExcelToDataTable(dlg1.FileName);
                var dtSonuc = new DataTable("SONUC");

                if (dtMain != null && dtMain.Rows.Count > 0)
                {
                    dtSonuc.Columns.Add("Anahtar");
                    dtSonuc.Columns.Add("Degerler");
                    var idList = dtMain.Rows.Cast<DataRow>().Where(x => x[idSira - 1] != null).Select(x => x[idSira - 1].ToString().Trim()).Distinct().ToList();
                    foreach (var element in idList)
                    {
                        var all = dtMain.Rows.Cast<DataRow>().Where(x => x[idSira - 1].ToString() == element).Select(x => x[icerikSira - 1].ToString().Trim()).Distinct().ToList();
                        if (all != null && all.Count > 0)
                        {
                            var veri = string.Join((birlesimTuru == MessageBoxResult.Yes ? ", " : Environment.NewLine), all);
                            dtSonuc.Rows.Add(element, veri);
                        }
                    }

                    ExcelYardimcisi.DataTableToExcel(dtSonuc, dlg2.FileName);
                }
            }
            return result;
        }


        private bool KlasorOlarakKaydet(string metin)
        {
            var liste = metin.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            if (liste.Count() < 1)
            {
                MesajWindow.GosterUyari("Listenizde Veri Yok");
                return false;
            }
            MesajWindow.GosterBilgi("Klasorlerin Oluşturulacağı Konumu Seçiniz");
            bool hata = false;
            var dlg = new System.Windows.Forms.FolderBrowserDialog();
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(dlg.SelectedPath))
            {
                Task.Run(() =>
                {
                    try
                    {
                        foreach (var item in liste)
                        {
                            string adi = DilYardimcisi.UyumluDosyaAdiYap(item.ToString());
                            if (!string.IsNullOrEmpty(adi.Trim()))
                            {
                                Directory.CreateDirectory(dlg.SelectedPath + "\\" + adi.Trim());
                            }
                        }

                    }
                    catch
                    {
                        hata = true;
                    }
                });
            }
            if (hata)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public string TumSatirlariTekSatirYap(string metin)
        {
            return metin.Replace(Environment.NewLine, string.Empty);
        }



        private void IslemFiltrele(string filtre)
        {
            var butonlar = PanelIslemlerListesi.Children.OfType<Button>();
            foreach (var item in butonlar)
            {
                if (item.Content.ToString().ToLower().Contains(filtre))
                {
                    item.Visibility = Visibility.Visible;
                }
                else
                {
                    item.Visibility = Visibility.Collapsed;
                }
            }
        }


        private void TxtIslemFiltre_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            IslemFiltrele(txtIslemFiltre.Text.ToLower());
        }
    }
}

public static class ExcelYardimcisi
{
    public static DataSet ExcelToDataSetOleDb(string importFile, int sheetNo)
    {
        var resultDataSet = new DataSet("DataSet");

        try
        {
            if (System.IO.File.Exists(importFile))
            {
                #region Kolon Kontrol İçin İlk Hesap
                var providerName = new System.Data.OleDb.OleDbEnumerator().GetElements().AsEnumerable().OrderByDescending(x => x.Field<string>("SOURCES_NAME")).FirstOrDefault(x => x.Field<string>("SOURCES_NAME").StartsWith("Microsoft.ACE.OLEDB.1"))?.Field<string>("SOURCES_NAME");
                if (providerName == null)
                {
                    throw new Exception("Bilgisayarınızda Microsoft Database Provider yaması yüklü olmayabilir.\nÖncelikle AccessDatabaseEngine2010.exe yamalarını yükleyiniz.");
                }
                using (var cnn = new OleDbConnection("Provider=" + providerName + ";Data Source=" + importFile + ";Extended Properties=\"Excel 12.0 Xml;HDR=NO;IMEX=1;TypeGuessRows=0;ImportMixedTypes=Text;\""))
                {
                    try
                    {
                        cnn.Open();
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Bilgisayarınızda uygun Microsoft Database Provider yaması yüklü olmayabilir.\nÖncelikle AccessDatabaseEngine2010.exe yamalarını yükleyiniz." + Environment.NewLine + ex.Message);
                    }

                    //DEĞERLERİ OKU
                    var dtScheme = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    int sayfaN = 0;
                    var sheetNames = dtScheme.Rows.Cast<DataRow>().Select(x => x.Field<string>("TABLE_NAME").Replace("'", "")).ToList();
                    foreach (var sheetName in sheetNames.Where(x => x.EndsWith("$")))
                    {
                        var cmdVerileriGetir = new OleDbCommand("SELECT * FROM [" + sheetName + "] ", cnn) { CommandTimeout = 600 };

                        var dtFullSonuc = new DataTable(sheetName.Replace("#", "_").Replace("'", "").Replace(".", "").Replace("]", "").Replace("[", "").Trim());
                        dtFullSonuc.Load(cmdVerileriGetir.ExecuteReader());

                        if (dtFullSonuc != null && dtFullSonuc.Columns.Count > 0 && dtFullSonuc.Rows.Count > 0)
                        {
                            var IlkKolonlar = dtFullSonuc.Rows[0].ItemArray.OfType<object>().Select(x => (object)x).ToList();
                            for (int i = 0; i < IlkKolonlar.Count; i++)
                            {
                                if (IlkKolonlar[i] == DBNull.Value || IlkKolonlar[i] == null)
                                {
                                    IlkKolonlar[i] = "#Col__" + (i + 1);
                                }
                            }

                            var yeniKolonlar = new List<string>();
                            for (int i = 0; i < IlkKolonlar.Count; i++)
                            {
                                string colName = IlkKolonlar[i].ToString().Trim();
                                //TODO:
                                colName = colName.Length > 120 ? colName.Substring(0, 125) : colName;
                                colName = colName.Replace(".", "_").Replace("?", "_").Replace("*", "_").Replace("#", "_").Replace("/", "_").Replace("'", "_").Replace("\n", "_").Trim();
                                while (yeniKolonlar.FirstOrDefault(x => ConvertToEN(x) == ConvertToEN(colName)) != null)
                                {
                                    colName = colName + "_";
                                }

                                yeniKolonlar.Add(colName);
                            }

                            dtFullSonuc.Rows[0].Delete();
                            for (int i = 0; i < dtFullSonuc.Columns.Count; i++)
                            {
                                dtFullSonuc.Columns[i].ColumnName = yeniKolonlar[i];
                            }
                            resultDataSet.Tables.Add(dtFullSonuc);
                            dtFullSonuc.AcceptChanges();
                            sayfaN++;
                            if (sayfaN != 0 && sayfaN == sheetNo)
                            {
                                break;
                            }
                        }
                    }

                    cnn.Close();

                }
                #endregion

            }
            else
            {
                throw new Exception("ERROR 404 Dosya yok");
            }
        }
        catch (Exception ex)
        {
            throw ex;
        }
        return resultDataSet;
    }

    public static string ConvertToEN(string text)
    {
        return string.Join("", text.Normalize(NormalizationForm.FormD).Where(c => char.GetUnicodeCategory(c) != System.Globalization.UnicodeCategory.NonSpacingMark)).ToUpper(System.Globalization.CultureInfo.GetCultureInfo("EN"));
    }

    public static DataTable ExcelToDataTable(string excelFile, int sheetNo = 1, bool hasHeader = true, int columnNo = 0)
    {
        if (!System.IO.File.Exists(excelFile))
        {
            return null;
        }
        else
        {
            try
            {
                using (ExcelPackage pack = new ExcelPackage())
                {
                    using (var stream = File.OpenRead(excelFile))
                    {
                        pack.Load(stream);
                    }
                    ExcelWorksheet ws = pack.Workbook.Worksheets[sheetNo];
                    var dataTable = new DataTable("DataTable");

                    var all_columns = (columnNo > 0 ? (ws.Cells[1, columnNo, 1, columnNo]) : (ws.Cells[1, 1, 1, ws.Dimension.End.Column])).ToArray();
                    for (int i = 0; i < all_columns.Length; i++)
                    {
                        string colName = all_columns[i].Value.ToString();
                        if (dataTable.Columns[colName] != null || !hasHeader)
                        {
                            colName = "F" + i + 1;
                        }
                        dataTable.Columns.Add(colName);
                    }

                    var ilkRow = hasHeader ? 2 : 1;
                    for (int i = ilkRow; i <= ws.Dimension.End.Row; i++)
                    {
                        var erow = columnNo > 0 ? (ws.Cells[i, columnNo, i, columnNo]) : (ws.Cells[i, 1, i, all_columns.Length]);
                        DataRow row = dataTable.Rows.Add();
                        foreach (var cell in erow)
                        {
                            row[columnNo > 0 ? 0 : (cell.Start.Column - 1)] = cell.Value;
                        }
                    }

                    return dataTable;
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }
    }
    public static string DataTableToExcel(DataTable dataTable, string fileName = "")
    {
        if (dataTable != null && dataTable.Rows.Count > 0)
        {
            try
            {
                if (string.IsNullOrEmpty(fileName))
                {
                    fileName = System.IO.Path.GetTempPath() + "\\" + Guid.NewGuid().ToString() + ".xlsx";
                }

                using (ExcelPackage xp = new ExcelPackage(new System.IO.FileInfo(fileName)))
                {
                    ExcelWorksheet ws = xp.Workbook.Worksheets.Add("Sayfa1");
                    ws.DefaultColWidth = 15D;
                    ws.Cells["A1"].LoadFromDataTable(dataTable, true);
                    xp.Save();
                    return fileName;
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        else
        {
            return null;
        }
    }

    public static void ToCSV(DataTable dtDataTable, string strFilePath)
    {
        StreamWriter sw = new StreamWriter(strFilePath, false, Encoding.UTF8);
        //headers    
        for (int i = 0; i < dtDataTable.Columns.Count; i++)
        {
            sw.Write(dtDataTable.Columns[i]);
            if (i < dtDataTable.Columns.Count - 1)
            {
                sw.Write(";");
            }
        }
        sw.Write(sw.NewLine);
        foreach (DataRow dr in dtDataTable.Rows)
        {
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                if (!Convert.IsDBNull(dr[i]))
                {
                    string value = dr[i].ToString().Replace("\"", "");
                    if (value.Contains(';'))
                    {
                        value = String.Format("\"{0}\"", value);
                        sw.Write(value);
                    }
                    else
                    {
                        sw.Write("\"" + value + "\"");
                    }
                }
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(";");
                }
            }
            sw.Write(sw.NewLine);
        }
        sw.Close();
    }

    public static string ExcelCreateWithColumns(List<string> columns, string fileName, bool autoOpen = true)
    {
        string result = string.Empty;
        if (columns != null && columns.Count > 0)
        {
            try
            {
                using (ExcelPackage xp = new ExcelPackage(new System.IO.FileInfo(fileName)))
                {
                    ExcelWorksheet ws = xp.Workbook.Worksheets.Add("Sayfa1");
                    ws.DefaultColWidth = 15D;
                    var table = new DataTable("table");
                    foreach (var item in columns)
                    {
                        table.Columns.Add(item);
                    }
                    ws.Cells["A1"].LoadFromDataTable(table, true);
                    xp.Save();
                    result = fileName;
                }
            }
            catch (Exception ex)
            {
            }
        }
        if (autoOpen)
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(fileName) { UseShellExecute = true }).WaitForExit();
        }
        return result;
    }

}
