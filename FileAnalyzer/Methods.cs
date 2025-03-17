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

        public string loginMenu()
        {
            Console.WriteLine("Which file type?");
            Console.WriteLine("-------------");
            Console.WriteLine("| 1- Text   |");
            Console.WriteLine("| 2- Pdf    |");
            Console.WriteLine("| 3- Word   |");
            Console.WriteLine("-------------");
            Console.WriteLine("Choose: ");
            return null;
        }
        public string filePrefer(int prefer)
        {
            OpenFileDialog of = new OpenFileDialog();

            switch (prefer)
            {
                case 1:
                    of.Filter = "Text Files|*.txt";
                    Console.WriteLine("txt file selected.");
                    break;
                case 2:
                    of.Filter = "Pdf Files|*.pdf";
                    Console.WriteLine("pdf file selected.");
                    break;
                case 3:
                    of.Filter = "Word Files|*.docx";
                    Console.WriteLine("docx file selected");
                    break;
                default:
                    Console.WriteLine("Wrong Choice");
                    return null;
            }
            if (of.ShowDialog() == DialogResult.OK)
            {
                Console.WriteLine("Selected File: " + of.FileName);
                return of.FileName;
            }
            
            
            Console.WriteLine("Not Selected File");
            return null;
            
            


        }
        public int conjunctionSay(string file, int prefer)
        {
            string[] conjunctions = { "ve", "veya", "ama", "fakat", "ancak", "çünkü", "ile", "zira" };
            if (prefer == 1)
            {
                
                int counter = 0;
                foreach (var item in File.ReadAllLines(file))
                {
                    foreach (var conjunction in conjunctions)
                    {
                        if (item.Contains(conjunction))
                        {
                            counter++;
                        }
                    }
                }
                return counter;
            }
            else if (prefer == 2)
            {
                int counter = 0;
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(file);
                TextAbsorber textAbsorber = new TextAbsorber();
                for(int i =1; i <= pdfDocument.Pages.Count; i++)
                {
                    pdfDocument.Pages[i].Accept(textAbsorber);
                    string pageText = textAbsorber.Text;
                    foreach (var conjunction in conjunctions)
                    {
                        int index = 0;
                        while ((index = pageText.IndexOf(conjunction, index, StringComparison.OrdinalIgnoreCase)) != -1)
                        {
                            counter++;
                            index += conjunction.Length;  // Sonraki arama için indeksi güncelliyoruz
                        }

                    }
                }
                return counter;

            }
            else if (prefer == 3)
            {
                int counter = 0;
                Aspose.Words.Document wordDocument = new Aspose.Words.Document(file); // Aspose.Words.Document kullanıyoruz
                foreach (Aspose.Words.Paragraph paragraph in wordDocument.GetChildNodes(NodeType.Paragraph, true))
                {
                    string paragraphText = paragraph.Range.Text; // Paragrafları metne dönüştürüyoruz
                    foreach (var conjunction in conjunctions)
                    {
                        int index = 0;
                        while ((index = paragraphText.IndexOf(conjunction, index, StringComparison.OrdinalIgnoreCase)) != -1)
                        {
                            counter++;
                            index += conjunction.Length;  // Sonraki arama için indeksi güncelliyoruz
                        }
                    }
                }
                return counter;
            }
                return 0;
        }

        public int numberSay(string file, int prefer)
        {
            string[] number = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" };
            if (prefer == 1)
            {
                int counter = 0;
                foreach (var item in File.ReadAllLines(file))
                {
                    string[] words = item.Split(' ');
                    foreach (string s in words)
                    {
                        if (s.All(char.IsDigit))
                        {
                            counter++;
                        }
                    }
                }
                return counter;
            }
            else if (prefer == 2)
            {
                int counter = 0;
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(file);
                TextAbsorber textAbsorber = new TextAbsorber();
                for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                {
                    pdfDocument.Pages[i].Accept(textAbsorber);
                    string pageText = textAbsorber.Text;
                    string[] words = pageText.Split(' ');
                    foreach (var s in number)
                    {
                        if (s.All(char.IsDigit))
                        {
                            counter++;
                        }
                    }
                }
                return counter;
            }
            else if (prefer == 3) // Word dosyası için
            {
                int counter = 0;
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
                            counter++;
                        }
                    }
                }

                return counter;
            }
            return 0;
        }

        public void totalWordSay(string file, int prefer)
        {
            int consuctionTotal = conjunctionSay(file, prefer);
            int numberTotal = numberSay(file, prefer);
            int punctuationTotal = punctuationSay(file, prefer);
            int sentenceTotal = sentenceSay(file, prefer);
            HashSet<string> wordSet = new HashSet<string>();
            
            if (prefer == 1)
            {
                int totalWordSay = 0;
                foreach (var item in File.ReadAllLines(file))
                {
                    string[] words = item.Split(new char[] { ' ', '.', ',', '?', '!', ':', ';', '-', '_', '(', ')', '[', ']', '{', '}', '<', '>', '/', '\\', '|', '*', '+', '=', '&', '%', '$', '#', '@', '^', '~', '`' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var word in words)
                    {
                        wordSet.Add(word.ToLower());
                    }
                }
                totalWordSay = wordSet.Count;
                
                
                Console.WriteLine("------------TOTAL------------");
                    Console.WriteLine((totalWordSay - consuctionTotal - numberTotal)+" benzersiz kelime,");
                Console.WriteLine(sentenceTotal + " cümle,");
                Console.WriteLine(numberTotal + " sayı,");
                Console.WriteLine(punctuationTotal + " noktalama işareti,");
                Console.WriteLine(consuctionTotal + " bağlaç,");
                Console.WriteLine("kullanılmıştır.");
                Console.WriteLine("------------------------------");
                Console.WriteLine(" Benzersiz kelime sayısına, bağlaçlar ve sayılar dahil edilmemiştir.");
                Console.WriteLine("------------------------------");

            }
            else if (prefer == 2) // PDF dosyası için
            {
                int totalWordSay = 0;
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(file);
                TextAbsorber textAbsorber = new TextAbsorber();

                for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                {
                    pdfDocument.Pages[i].Accept(textAbsorber);
                    string pageText = textAbsorber.Text;
                    string[] words = pageText.Split(new char[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var word in words)
                    {
                        wordSet.Add(word.ToLower());
                    }
                }
                totalWordSay = wordSet.Count;
                Console.WriteLine("------------TOTAL------------");
                Console.WriteLine((totalWordSay - consuctionTotal - numberTotal) + " benzersiz kelime,");
                Console.WriteLine(sentenceTotal + " cümle,");
                Console.WriteLine(numberTotal + " sayı,");
                Console.WriteLine(punctuationTotal + " noktalama işareti,");
                Console.WriteLine(consuctionTotal + " bağlaç,");
                Console.WriteLine("kullanılmıştır.");
                Console.WriteLine("------------------------------");
                Console.WriteLine(" Benzersiz kelime sayısına, bağlaçlar ve sayılar dahil edilmemiştir.");
                Console.WriteLine("------------------------------");
            }
            else if (prefer == 3) // Word dosyası için
            {
                int totalWordSay = 0;
                Aspose.Words.Document wordDocument = new Aspose.Words.Document(file);

                // Word dosyasındaki tüm metni paragraf paragraf okuyalım
                foreach (Aspose.Words.Paragraph paragraph in wordDocument.GetChildNodes(Aspose.Words.NodeType.Paragraph, true))
                {
                    string paragraphText = paragraph.Range.Text;

                    // Kelimeleri ayırırken boşluk ve diğer özel karakterleri dikkate alalım
                    string[] words = paragraphText.Split(new char[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var word in words)
                    {
                        wordSet.Add(word.ToLower());
                    }
                }
                totalWordSay = wordSet.Count;
                Console.WriteLine("------------TOTAL------------");
                Console.WriteLine((totalWordSay - consuctionTotal - numberTotal) + " benzersiz kelime,");
                Console.WriteLine(sentenceTotal + " cümle,");
                Console.WriteLine(numberTotal + " sayı,");
                Console.WriteLine(punctuationTotal + " noktalama işareti,");
                Console.WriteLine(consuctionTotal + " bağlaç,");
                Console.WriteLine("kullanılmıştır.");
                Console.WriteLine("------------------------------");
                Console.WriteLine(" Benzersiz kelime sayısına, bağlaçlar ve sayılar dahil edilmemiştir.");
                Console.WriteLine("------------------------------");
            }

        }

        public int sentenceSay(string file, int prefer)
        {

            if (prefer == 1)
            {
                int counter = 0;
                foreach (var item in File.ReadAllLines(file))
                {
                    string[] sentences = item.Split(new char[] { '.', '!', '?' }, StringSplitOptions.RemoveEmptyEntries);
                    counter += sentences.Length;
                }
                return counter;
            }
            else if (prefer == 2) // PDF dosyası için
            {
                int counter = 0;
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(file);
                TextAbsorber textAbsorber = new TextAbsorber();
                for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                {
                    pdfDocument.Pages[i].Accept(textAbsorber);
                    string pageText = textAbsorber.Text;
                    // Cümleleri ayırmak için nokta, ünlem, soru işareti kullanıyoruz
                    string[] sentences = pageText.Split(new char[] { '.', '!', '?' }, StringSplitOptions.RemoveEmptyEntries);
                    counter += sentences.Length;
                }
                return counter;
            }
            else if (prefer == 3) // Word dosyası için
            {
                int counter = 0;
                Aspose.Words.Document wordDocument = new Aspose.Words.Document(file);

                // Word dosyasındaki tüm metni paragraf paragraf okuyalım
                foreach (Aspose.Words.Paragraph paragraph in wordDocument.GetChildNodes(Aspose.Words.NodeType.Paragraph, true))
                {
                    string paragraphText = paragraph.Range.Text;
                    // Cümleleri ayırmak için nokta, ünlem, soru işareti kullanıyoruz
                    string[] sentences = paragraphText.Split(new char[] { '.', '!', '?' }, StringSplitOptions.RemoveEmptyEntries);
                    counter += sentences.Length;
                }

                return counter;
            }
            else
            {
                return 0;
            }
        }
        public int punctuationSay(string file, int prefer)
        {
            char[] punctuations = new char[] { '.', ',', '?', '!', ':', ';', '-', '_', '(', ')', '[', ']', '{', '}', '<', '>', '/', '\\', '|', '*', '+', '=', '&', '%', '$', '#', '@', '^', '~', '`' };
            if (prefer == 1)
            {
                int counter = 0;
                foreach (var item in File.ReadAllLines(file))
                {
                    foreach (var c in item)
                    {
                        if (punctuations.Contains(c))
                        {
                            counter++;
                        }
                    }
                }
                return counter;
            }
            else if (prefer == 2) // PDF dosyası için
            {
                int counter = 0;
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(file);
                TextAbsorber textAbsorber = new TextAbsorber();
                for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                {
                    pdfDocument.Pages[i].Accept(textAbsorber);
                    string pageText = textAbsorber.Text;
                    foreach (var c in pageText)
                    {
                        if (punctuations.Contains(c))
                        {
                            counter++;
                        }
                    }
                }
                return counter;
            }
            else if (prefer == 3) // Word dosyası için
            {
                int counter = 0;
                Aspose.Words.Document wordDocument = new Aspose.Words.Document(file);
                // Word dosyasındaki tüm metni paragraf paragraf okuyalım
                foreach (Aspose.Words.Paragraph paragraph in wordDocument.GetChildNodes(Aspose.Words.NodeType.Paragraph, true))
                {
                    string paragraphText = paragraph.Range.Text;
                    foreach (var c in paragraphText)
                    {
                        if (punctuations.Contains(c))
                        {
                            counter++;
                        }
                    }
                }
                return counter;
            }
            else
            {
                return 0;
            }
        }

        public void wordSay(string file, int prefer)
        {
            char[] punctuation = new char[] { ' ', '.', ',', '?', '!', ':', ';', '-', '_', '(', ')', '[', ']', '{', '}', '<', '>', '/', '\\', '|', '*', '+', '=', '&', '%', '$', '#', '@', '^', '~', '`' };

            if (prefer == 1)
            {
                string text = File.ReadAllText(file);
                string clear = new string(text.Select(c => punctuation.Contains(c) && c != ' ' ? ' ' : c).ToArray());

                string[] word = clear.Split(new char[] { ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                Dictionary<string, int> wordSay = new Dictionary<string, int>();
                foreach (string words in word)
                {
                    if (words.All(c => Char.IsLetter(c)))
                    {
                        if (wordSay.ContainsKey(words))
                        {
                            wordSay[words]++;
                        }
                        else
                        {
                            wordSay.Add(words, 1);
                        }
                    }
                }
                var sortedWordSay = wordSay.OrderByDescending(kv => kv.Value).ToList();
                foreach (var item in sortedWordSay)
                {
                    if (item.Value > 1)
                    {
                        Console.Write(" " + item.Key + "=" + item.Value + ", ");
                    }
                }
                Console.WriteLine("defa kullanılmıştır.");
            }
            else if (prefer == 2) // PDF dosyası için
            {
                string text = string.Empty;
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(file);
                TextAbsorber textAbsorber = new TextAbsorber();
                for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                {
                    pdfDocument.Pages[i].Accept(textAbsorber);
                    text += textAbsorber.Text;
                }

                string clear = new string(text.Select(c => punctuation.Contains(c) && c != ' ' ? ' ' : c).ToArray());
                string[] words = clear.Split(new char[] { ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                Dictionary<string, int> wordSay = new Dictionary<string, int>();
                foreach (string word in words)
                {
                    if (word.All(c => Char.IsLetter(c)))
                    {
                        if (wordSay.ContainsKey(word))
                        {
                            wordSay[word]++;
                        }
                        else
                        {
                            wordSay.Add(word, 1);
                        }
                    }
                }
                var sortedWordSay = wordSay.OrderByDescending(kv => kv.Value).ToList();
                foreach (var item in sortedWordSay)
                {
                    if (item.Value > 1)
                    {
                        Console.Write(" " + item.Key + "=" + item.Value + ", ");
                    }
                }
                Console.WriteLine("defa kullanılmıştır.");
            }
            else if (prefer == 3) // Word dosyası için
            {
                string text = string.Empty;
                Aspose.Words.Document wordDocument = new Aspose.Words.Document(file);

                // Word dosyasındaki tüm metni paragraf paragraf okuyalım
                foreach (Aspose.Words.Paragraph paragraph in wordDocument.GetChildNodes(Aspose.Words.NodeType.Paragraph, true))
                {
                    text += paragraph.Range.Text;
                }

                string clear = new string(text.Select(c => punctuation.Contains(c) && c != ' ' ? ' ' : c).ToArray());
                string[] words = clear.Split(new char[] { ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                Dictionary<string, int> wordSay = new Dictionary<string, int>();
                foreach (string word in words)
                {
                    if (wordSay.ContainsKey(word))
                    {
                        wordSay[word]++;
                    }
                    else
                    {
                        wordSay.Add(word, 1);
                    }
                }
                var sortedWordSay = wordSay.OrderByDescending(kv => kv.Value).ToList();
                foreach (var item in sortedWordSay)
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
