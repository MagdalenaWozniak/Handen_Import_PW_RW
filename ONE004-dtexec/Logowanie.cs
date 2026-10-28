using CDNBase;

//using OfficeOpenXml;
//using OfficeOpenXml.Table;
using ONE012.Properties;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ONE012
{
    public partial class Logowanie : Form
    {
        protected static ILogin Login = null;

        public Logowanie()
        {
            ///Zamiast usuwać stare metody komentuje je, by przy tesotwaniu można było dojść do tego jak rzecz działała wcześniej



            //MessageBox.Show(Program.args0 + " aTeraz sprawdzam czy w środku sie to zmienia " + Program.args1);
            InitializeComponent();
            LogowanieBezPopup();
            //if (Program.args0 != null)
            //{
            //    textBox1.Text = Program.args0;
            //}
            ////else
            ////{
            ////    textBox1.Text = Properties.Settings.userName;
            ////}
            //textBox3.Text = Program.args1;
        }


        protected static CDNBase.Application Application = null;

        private void Button1_Click(object sender, EventArgs e)
        {
            ONE012WIN.Operator = "ONE_DOD";
            //ONE012WIN.Haslo = "NE!w0rk12";
            ONE012WIN.Firma = "FIRMA_DEMO";

            try
            {
                LogowanieAutomatyczne();
                Program.zalogowano = true;
            }
            catch (COMException eo)
            {
                if (eo.InnerException != null)
                {
                    MessageBox.Show("Błąd logowania: " + eo.InnerException.Message, "Komunikat1",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show("Błąd logowania: " + eo.Message, "Komunikat2",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ei)
            {
                if (ei.InnerException != null)
                {
                    MessageBox.Show("Błąd logowania: " + ei.InnerException.Message, "Komunikat3",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show("Błąd logowania: " + ei.Message, "Komunikat4",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            if (!Program.zalogowano)
            {
                System.Windows.Forms.Application.Exit();
            }

            Wylogowanie();

            Close();
        }

        private void LogowanieBezPopup()
        {

            ///Tu  podane są "na sztywno" dane logowania.
            ONE012WIN.Operator = "ONE_DOD";
            //ONE012WIN.Haslo = "NE!w0rk12";
            ONE012WIN.Firma = "FIRMA_DEMO";







            //ONE012WIN.Operator = Program.args0;
            //ONE012WIN.Haslo = "";
            //ONE012WIN.Firma = Program.args1;
            //if (checkBox1.Checked)
            //{
            //    Settings.Default.userName = textBox1.Text;
            //    Settings.Default.firmaName = textBox3.Text;
            //    Properties.Settings.Default.Save();
            //}
            //else if (!checkBox1.Checked)
            //{
            //    Settings.Default.userName = null;
            //    Settings.Default.firmaName = null;
            //    Properties.Settings.Default.Save();
            //}


            try
            {
                LogowanieAutomatyczne();
                Program.zalogowano = true;
            }
            catch (COMException eo)
            {
                if (eo.InnerException != null)
                {
                    MessageBox.Show("Błąd logowania: " + eo.InnerException.Message, "Komunikat1",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show("Błąd logowania: " + eo.Message, "Komunikat2",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ei)
            {
                if (ei.InnerException != null)
                {
                    MessageBox.Show("Błąd logowania: " + ei.InnerException.Message, "Komunikat3",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show("Błąd logowania: " + ei.Message, "Komunikat4",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            if (!Program.zalogowano)
            {
                System.Windows.Forms.Application.Exit();
            }

            Wylogowanie();

            //Close();
        }

        //private void Form1_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        //{
        //    // Determine if text has changed in the textbox by comparing to original text.
        //    if (System.Windows.Forms.Application.MessageLoop)
        //    {
        //        // WinForms app
        //        System.Windows.Forms.Application.Exit();
        //    }
        //    else
        //    {
        //        // Console app
        //        System.Environment.Exit(0);
        //    }
        //}

        static protected void LogowanieAutomatyczne()//do walniecia logowanie
        {
            object[] hPar = new object[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0 };    // do jakich modułów się logujemy
                                                                                                   /* Kolejno: KP, KH, KHP, ST, FA, MAG, PK, PKXL, CRM, ANL, DET, BIU, SRW, ODB, KB, KBP, HAP
                                                                                                    */

            // katalog, gdzie jest zainstalowana Optima (bez ustawienia tej zmiennej nie zadziała, chyba że program odpalimy z katalogu O!)
            System.Environment.CurrentDirectory = @"C:\Program Files (x86)\Comarch ERP Optima";
            // tworzymy nowy obiekt apliakcji
            Application = new CDNBase.Application();
            // blokujemy
            Application.LockApp(513, 5000, null, null, null, null);
            // logujemy się do podanej Firmy, na danego operatora, do podanych modułów
            Login = Application.Login(vUser: ONE012WIN.Operator, vChk: ""/*ONE012WIN.Haslo*/, vFirm: ONE012WIN.Firma, _ModulKP: hPar[0], _ModulKH: hPar[1], _ModulKHP: hPar[2], _ModulST: hPar[3], _ModulFA: hPar[4], _ModulMAG: hPar[5], _ModulPK: hPar[6], _ModulPKXL: hPar[7], _ModulCRM: hPar[8], _ModulANL: hPar[9], _ModulDET: hPar[10], _ModulBIU: hPar[11], _ModulSRW: hPar[12], _ModulOBD: hPar[13], _ModulKB: hPar[14], _ModulKBP: hPar[15], _ModulHAP: hPar[16]);

            //problem jest taki, że nei ma w tej metodzie dostępu do KBP do którego chc się zalogować

            // tu jesteśmy zalogowani do O!
            //MessageBox.Show("Jesteśmy zalogowani do O!");
        }

        protected static void Wylogowanie()
        {
            // niszczymy Login
            Login = null;
            // odblokowanie (wylogowanie) O!
            Application.UnlockApp();
            // niszczymy obiekt Aplikacji
            Application = null;
        }

        private void Logowanie_Load(object sender, EventArgs e)
        {
            textBox1.Text = Settings.Default.userName;
            textBox3.Text = Settings.Default.firmaName;
            ToolTip ToolTip1 = new ToolTip();
            ToolTip1.SetToolTip(this.checkBox1, "Zapamiętane zostaną \"Operator\" i \"Firma\".");
        }
    }
}