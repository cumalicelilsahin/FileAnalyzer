using System;
using System.Windows.Forms;

namespace FileAnalyzer
{
    class Program
    {
        [STAThread]        
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Methods m = new Methods();
            m.girisEkran();
            int secim = int.Parse(Console.ReadLine());
            Console.WriteLine("------------------------------");

            string dosyaYolu = m.DosyaTercih(secim);
            if (!string.IsNullOrEmpty(dosyaYolu))
            {
                Console.WriteLine("------------------------------");
                m.toplamKelimeSay(dosyaYolu, secim);
                Console.WriteLine("------------------------------");
                m.kelimeSay(dosyaYolu, secim);
                Console.WriteLine("------------------------------");

            }
        }
    }
}
