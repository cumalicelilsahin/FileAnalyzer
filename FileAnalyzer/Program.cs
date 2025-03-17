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
            m.loginMenu();
            int prefer = int.Parse(Console.ReadLine());
            Console.WriteLine("------------------------------");

            string fileRoad = m.filePrefer(prefer);
            if (!string.IsNullOrEmpty(fileRoad))
            {
                Console.WriteLine("------------------------------");
                m.totalWordSay(fileRoad, prefer);
                Console.WriteLine("------------------------------");
                m.wordSay(fileRoad, prefer);
                Console.WriteLine("------------------------------");

            }
        }
    }
}
