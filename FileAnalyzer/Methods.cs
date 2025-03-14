using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using Aspose.Pdf.Text;
using Aspose.Words;

namespace FileAnalyzer
{
    public class Methods
    {
        public string girisEkran()
        {
            Console.WriteLine("Hangi dosya türünü analiz etmek istersiniz?");
            Console.WriteLine("-------------");
            Console.WriteLine("| 1- Text   |");
            Console.WriteLine("| 2- Pdf    |");
            Console.WriteLine("| 3- Word   |");
            Console.WriteLine("-------------");
            Console.WriteLine("Seçiminizi yapınız: ");
            return null;
        }
        public string DosyaTercih(int tercih)
        {
            OpenFileDialog of = new OpenFileDialog();

            switch (tercih)
            {
                case 1:
                    of.Filter = "Text Dosyaları|*.txt";
                    Console.WriteLine("Text Dosyası Seçildi");
                    break;
                case 2:
                    of.Filter = "Pdf Dosyaları|*.pdf";
                    Console.WriteLine("Pdf Dosyası Seçildi");
                    break;
                case 3:
                    of.Filter = "Word Dosyaları|*.docx";
                    Console.WriteLine("Word Dosyası Seçildi");
                    break;
                default:
                    Console.WriteLine("Geçersiz Seçim");
                    return null;
            }
            if (of.ShowDialog() == DialogResult.OK)
            {
                Console.WriteLine("Seçilen Dosya: " + of.FileName);
                return of.FileName;
            }
            
            
            Console.WriteLine("Dosya Seçilmedi");
            return null;
            
            


        }
        public int baglacSay(string file, int tercih)
        {
            string[] baglaclar = { "ve", "veya", "ama", "fakat", "ancak", "çünkü", "ile", "zira" };
            if (tercih == 1)
            {
                
                int sayac = 0;
                foreach (var item in File.ReadAllLines(file))
                {
                    foreach (var baglac in baglaclar)
                    {
                        if (item.Contains(baglac))
                        {
                            sayac++;
                        }
                    }
                }
                return sayac;
            }
            else if (tercih == 2)
            {
                int sayac = 0;
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(file);
                TextAbsorber textAbsorber = new TextAbsorber();
                for(int i =1; i <= pdfDocument.Pages.Count; i++)
                {
                    pdfDocument.Pages[i].Accept(textAbsorber);
                    string pageText = textAbsorber.Text;
                    foreach (var baglac in baglaclar)
                    {
                        int index = 0;
                        while ((index = pageText.IndexOf(baglac, index, StringComparison.OrdinalIgnoreCase)) != -1)
                        {
                            sayac++;
                            index += baglac.Length;  // Sonraki arama için indeksi güncelliyoruz
                        }

                    }
                }
                return sayac;

            }
            else if (tercih == 3)
            {
                int sayac = 0;
                Aspose.Words.Document wordDocument = new Aspose.Words.Document(file); // Aspose.Words.Document kullanıyoruz
                foreach (Aspose.Words.Paragraph paragraph in wordDocument.GetChildNodes(NodeType.Paragraph, true))
                {
                    string paragraphText = paragraph.Range.Text; // Paragrafları metne dönüştürüyoruz
                    foreach (var baglac in baglaclar)
                    {
                        int index = 0;
                        while ((index = paragraphText.IndexOf(baglac, index, StringComparison.OrdinalIgnoreCase)) != -1)
                        {
                            sayac++;
                            index += baglac.Length;  // Sonraki arama için indeksi güncelliyoruz
                        }
                    }
                }
                return sayac;
            }
                return 0;
        }

        public int sayiSay(string file, int tercih)
        {
            string[] sayi = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" };
            if (tercih == 1)
            {
                int sayac = 0;
                foreach (var item in File.ReadAllLines(file))
                {
                    string[] words = item.Split(' ');
                    foreach (string s in words)
                    {
                        if (s.All(char.IsDigit))
                        {
                            sayac++;
                        }
                    }
                }
                return sayac;
            }
            else if (tercih == 2)
            {
                int sayac = 0;
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(file);
                TextAbsorber textAbsorber = new TextAbsorber();
                for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                {
                    pdfDocument.Pages[i].Accept(textAbsorber);
                    string pageText = textAbsorber.Text;
                    string[] words = pageText.Split(' ');
                    foreach (var s in sayi)
                    {
                        if (s.All(char.IsDigit))
                        {
                            sayac++;
                        }
                    }
                }
                return sayac;
            }
            else if (tercih == 3) // Word dosyası için
            {
                int sayac = 0;
                Aspose.Words.Document wordDocument = new Aspose.Words.Document(file);

                // Word dosyasındaki tüm metni paragraf paragraf okuyalım
                foreach (Aspose.Words.Paragraph paragraph in wordDocument.GetChildNodes(Aspose.Words.NodeType.Paragraph, true))
                {
                    string paragraphText = paragraph.Range.Text;

                    // Kelimeleri ayırırken boşluk ve diğer özel karakterleri dikkate alalım
                    string[] words = paragraphText.Split(new char[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string s in words)
                    {
                        if (s.All(char.IsDigit))
                        {
                            sayac++;
                        }
                    }
                }

                return sayac;
            }
            return 0;
        }

        public void toplamKelimeSay(string file, int tercih)
        {
            int baglacSayisi = baglacSay(file, tercih);
            int sayiSayisi = sayiSay(file, tercih);
            int noktalamaSayisi = noktalamaSay(file, tercih);
            int cumleSayisi = cumleSay(file, tercih);
            HashSet<string> kelimelerSeti = new HashSet<string>();
            
            if (tercih == 1)
            {
                int toplamKelimeSayisi = 0;
                foreach (var item in File.ReadAllLines(file))
                {
                    string[] words = item.Split(new char[] { ' ', '.', ',', '?', '!', ':', ';', '-', '_', '(', ')', '[', ']', '{', '}', '<', '>', '/', '\\', '|', '*', '+', '=', '&', '%', '$', '#', '@', '^', '~', '`' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var word in words)
                    {
                        kelimelerSeti.Add(word.ToLower());
                    }
                }
                toplamKelimeSayisi = kelimelerSeti.Count;
                
                
                Console.WriteLine("------------TOPLAM------------");
                    Console.WriteLine((toplamKelimeSayisi - baglacSayisi - sayiSayisi)+" benzersiz kelime,");
                Console.WriteLine(cumleSayisi + " cümle,");
                Console.WriteLine(sayiSayisi + " sayı,");
                Console.WriteLine(noktalamaSayisi + " noktalama işareti,");
                Console.WriteLine(baglacSayisi + " bağlaç,");
                Console.WriteLine("kullanılmıştır.");
                Console.WriteLine("------------------------------");
                Console.WriteLine(" Benzersiz kelime sayısına, bağlaçlar ve sayılar dahil edilmemiştir.");
                Console.WriteLine("------------------------------");

            }
            else if (tercih == 2) // PDF dosyası için
            {
                int toplamKelimeSayisi = 0;
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(file);
                TextAbsorber textAbsorber = new TextAbsorber();

                for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                {
                    pdfDocument.Pages[i].Accept(textAbsorber);
                    string pageText = textAbsorber.Text;
                    string[] words = pageText.Split(new char[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var word in words)
                    {
                        kelimelerSeti.Add(word.ToLower());
                    }
                }
                toplamKelimeSayisi = kelimelerSeti.Count;
                Console.WriteLine("------------TOPLAM------------");
                Console.WriteLine((toplamKelimeSayisi - baglacSayisi - sayiSayisi) + " benzersiz kelime,");
                Console.WriteLine(cumleSayisi + " cümle,");
                Console.WriteLine(sayiSayisi + " sayı,");
                Console.WriteLine(noktalamaSayisi + " noktalama işareti,");
                Console.WriteLine(baglacSayisi + " bağlaç,");
                Console.WriteLine("kullanılmıştır.");
                Console.WriteLine("------------------------------");
                Console.WriteLine(" Benzersiz kelime sayısına, bağlaçlar ve sayılar dahil edilmemiştir.");
                Console.WriteLine("------------------------------");
            }
            else if (tercih == 3) // Word dosyası için
            {
                int toplamKelimeSayisi = 0;
                Aspose.Words.Document wordDocument = new Aspose.Words.Document(file);

                // Word dosyasındaki tüm metni paragraf paragraf okuyalım
                foreach (Aspose.Words.Paragraph paragraph in wordDocument.GetChildNodes(Aspose.Words.NodeType.Paragraph, true))
                {
                    string paragraphText = paragraph.Range.Text;

                    // Kelimeleri ayırırken boşluk ve diğer özel karakterleri dikkate alalım
                    string[] words = paragraphText.Split(new char[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var word in words)
                    {
                        kelimelerSeti.Add(word.ToLower());
                    }
                }
                toplamKelimeSayisi = kelimelerSeti.Count;
                Console.WriteLine("------------TOPLAM------------");
                Console.WriteLine((toplamKelimeSayisi - baglacSayisi - sayiSayisi) + " benzersiz kelime,");
                Console.WriteLine(cumleSayisi + " cümle,");
                Console.WriteLine(sayiSayisi + " sayı,");
                Console.WriteLine(noktalamaSayisi + " noktalama işareti,");
                Console.WriteLine(baglacSayisi + " bağlaç,");
                Console.WriteLine("kullanılmıştır.");
                Console.WriteLine("------------------------------");
                Console.WriteLine(" Benzersiz kelime sayısına, bağlaçlar ve sayılar dahil edilmemiştir.");
                Console.WriteLine("------------------------------");
            }

        }

        public int cumleSay(string file, int tercih)
        {

            if (tercih == 1)
            {
                int sayac = 0;
                foreach (var item in File.ReadAllLines(file))
                {
                    string[] cumleler = item.Split(new char[] { '.', '!', '?' }, StringSplitOptions.RemoveEmptyEntries);
                    sayac += cumleler.Length;
                }
                return sayac;
            }
            else if (tercih == 2) // PDF dosyası için
            {
                int sayac = 0;
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(file);
                TextAbsorber textAbsorber = new TextAbsorber();
                for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                {
                    pdfDocument.Pages[i].Accept(textAbsorber);
                    string pageText = textAbsorber.Text;
                    // Cümleleri ayırmak için nokta, ünlem, soru işareti kullanıyoruz
                    string[] cumleler = pageText.Split(new char[] { '.', '!', '?' }, StringSplitOptions.RemoveEmptyEntries);
                    sayac += cumleler.Length;
                }
                return sayac;
            }
            else if (tercih == 3) // Word dosyası için
            {
                int sayac = 0;
                Aspose.Words.Document wordDocument = new Aspose.Words.Document(file);

                // Word dosyasındaki tüm metni paragraf paragraf okuyalım
                foreach (Aspose.Words.Paragraph paragraph in wordDocument.GetChildNodes(Aspose.Words.NodeType.Paragraph, true))
                {
                    string paragraphText = paragraph.Range.Text;
                    // Cümleleri ayırmak için nokta, ünlem, soru işareti kullanıyoruz
                    string[] cumleler = paragraphText.Split(new char[] { '.', '!', '?' }, StringSplitOptions.RemoveEmptyEntries);
                    sayac += cumleler.Length;
                }

                return sayac;
            }
            else
            {
                return 0;
            }
        }
        public int noktalamaSay(string file, int tercih)
        {
            char[] noktalama = new char[] { '.', ',', '?', '!', ':', ';', '-', '_', '(', ')', '[', ']', '{', '}', '<', '>', '/', '\\', '|', '*', '+', '=', '&', '%', '$', '#', '@', '^', '~', '`' };
            if (tercih == 1)
            {
                int sayac = 0;
                foreach (var item in File.ReadAllLines(file))
                {
                    foreach (var c in item)
                    {
                        if (noktalama.Contains(c))
                        {
                            sayac++;
                        }
                    }
                }
                return sayac;
            }
            else if (tercih == 2) // PDF dosyası için
            {
                int sayac = 0;
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(file);
                TextAbsorber textAbsorber = new TextAbsorber();
                for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                {
                    pdfDocument.Pages[i].Accept(textAbsorber);
                    string pageText = textAbsorber.Text;
                    foreach (var c in pageText)
                    {
                        if (noktalama.Contains(c))
                        {
                            sayac++;
                        }
                    }
                }
                return sayac;
            }
            else if (tercih == 3) // Word dosyası için
            {
                int sayac = 0;
                Aspose.Words.Document wordDocument = new Aspose.Words.Document(file);
                // Word dosyasındaki tüm metni paragraf paragraf okuyalım
                foreach (Aspose.Words.Paragraph paragraph in wordDocument.GetChildNodes(Aspose.Words.NodeType.Paragraph, true))
                {
                    string paragraphText = paragraph.Range.Text;
                    foreach (var c in paragraphText)
                    {
                        if (noktalama.Contains(c))
                        {
                            sayac++;
                        }
                    }
                }
                return sayac;
            }
            else
            {
                return 0;
            }
        }

        public void kelimeSay(string file, int tercih)
        {
            char[] noktalama = new char[] { ' ', '.', ',', '?', '!', ':', ';', '-', '_', '(', ')', '[', ']', '{', '}', '<', '>', '/', '\\', '|', '*', '+', '=', '&', '%', '$', '#', '@', '^', '~', '`' };

            if (tercih == 1)
            {
                string metin = File.ReadAllText(file);
                string temizle = new string(metin.Select(c => noktalama.Contains(c) && c != ' ' ? ' ' : c).ToArray());

                string[] words = temizle.Split(new char[] { ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                Dictionary<string, int> kelimeSay = new Dictionary<string, int>();
                foreach (string kelimeler in words)
                {
                    if (kelimeler.All(c => Char.IsLetter(c)))
                    {
                        if (kelimeSay.ContainsKey(kelimeler))
                        {
                            kelimeSay[kelimeler]++;
                        }
                        else
                        {
                            kelimeSay.Add(kelimeler, 1);
                        }
                    }
                }
                var sortedKelimeSay = kelimeSay.OrderByDescending(kv => kv.Value).ToList();
                foreach (var item in sortedKelimeSay)
                {
                    if (item.Value > 1)
                    {
                        Console.Write(" " + item.Key + "=" + item.Value + ", ");
                    }
                }
                Console.WriteLine("defa kullanılmıştır.");
            }
            else if (tercih == 2) // PDF dosyası için
            {
                string text = string.Empty;
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(file);
                TextAbsorber textAbsorber = new TextAbsorber();
                for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                {
                    pdfDocument.Pages[i].Accept(textAbsorber);
                    text += textAbsorber.Text;
                }

                string temizle = new string(text.Select(c => noktalama.Contains(c) && c != ' ' ? ' ' : c).ToArray());
                string[] words = temizle.Split(new char[] { ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                Dictionary<string, int> kelimeSay = new Dictionary<string, int>();
                foreach (string kelimeler in words)
                {
                    if (kelimeler.All(c => Char.IsLetter(c)))
                    {
                        if (kelimeSay.ContainsKey(kelimeler))
                        {
                            kelimeSay[kelimeler]++;
                        }
                        else
                        {
                            kelimeSay.Add(kelimeler, 1);
                        }
                    }
                }
                var sortedKelimeSay = kelimeSay.OrderByDescending(kv => kv.Value).ToList();
                foreach (var item in sortedKelimeSay)
                {
                    if (item.Value > 1)
                    {
                        Console.Write(" " + item.Key + "=" + item.Value + ", ");
                    }
                }
                Console.WriteLine("defa kullanılmıştır.");
            }
            else if (tercih == 3) // Word dosyası için
            {
                string text = string.Empty;
                Aspose.Words.Document wordDocument = new Aspose.Words.Document(file);

                // Word dosyasındaki tüm metni paragraf paragraf okuyalım
                foreach (Aspose.Words.Paragraph paragraph in wordDocument.GetChildNodes(Aspose.Words.NodeType.Paragraph, true))
                {
                    text += paragraph.Range.Text;
                }

                string temizle = new string(text.Select(c => noktalama.Contains(c) && c != ' ' ? ' ' : c).ToArray());
                string[] words = temizle.Split(new char[] { ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                Dictionary<string, int> kelimeSay = new Dictionary<string, int>();
                foreach (string kelimeler in words)
                {
                    if (kelimeSay.ContainsKey(kelimeler))
                    {
                        kelimeSay[kelimeler]++;
                    }
                    else
                    {
                        kelimeSay.Add(kelimeler, 1);
                    }
                }
                var sortedKelimeSay = kelimeSay.OrderByDescending(kv => kv.Value).ToList();
                foreach (var item in sortedKelimeSay)
                {
                    if (item.Value > 1)
                    {
                        Console.Write(" " + item.Key + "=" + item.Value + ", ");
                    }
                }
                Console.WriteLine("defa kullanılmıştır.");
            }
        }
    }
}
