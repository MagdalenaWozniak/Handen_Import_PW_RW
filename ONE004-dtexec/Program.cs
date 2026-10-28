using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ONE032
{
    static class Program
    {
        public static bool zalogowano = false;
        public static string args0 ="";
        public static string args1 ="";
        /// <summary>
        /// Główny punkt wejścia dla aplikacji.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        { 
            if (args.Count() == 2)
            {
                args0 = args[0];
                args1 = args[1];
            }

            //MessageBox.Show(args0+" "+ args1);
            
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Logowanie());
            //if(zalogowano)
            //MessageBox.Show(args0 +" i "+ args1);
            Application.Run(new ONE032WIN());
        }
    }
}
