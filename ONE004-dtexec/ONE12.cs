using System;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using CDNBase;
using ExcelDataReader;

using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using CDNHlmn;

namespace ONE032
{
    public partial class ONE032WIN : Form
    {
        public static string sKnt_Kod = "!NIEOKREŚLONY!";
        public static string Operator = "ADMIN";
        public static string Haslo = "NE!w0rk12";
        public static string Firma = "Handen";
        public static string Baza = "[CDN_Handen].[CDN].";
        string connectionString = @"Provider=SQLNCLI10; DataTypeCompatibility=80; Data Source=s354base30;Initial Catalog=" + Baza.Replace(".[CDN].","")+ "; User ID=onesupport;Password=j75kj#jt{493!@ma3";

       
        protected static CDNBase.Application Application = null;
        protected static ILogin Login = null;
        CDNBase.AdoSession Sesja;
        public ONE032WIN()
        {
            InitializeComponent();
           
        }

        #region funkcje optimy


        static protected void LogowanieAutomatyczne()
        {
            object[] hPar = new object[] {
                         0,  0,   0,  0,  0,   0,  0,    0,   0,   0,   0,   0,   0,   0,  0,   1,  1, 1 };    // do jakich modułów się logujemy
            /* Kolejno: KP, KH, KHP, ST, FA, MAG, PK, PKXL, CRM, ANL, DET, BIU, SRW, ODB, KB, KBP, HAP*/
            // katalog, gdzie jest zainstalowana Optima (bez ustawienia tej zmiennej nie zadziała, chyba że program odpalimy z katalogu O!)
            System.Environment.CurrentDirectory = @"C:\Program Files (x86)\Comarch ERP Optima";
            // tworzymy nowy obiekt apliakcji
            Application = new CDNBase.Application();
            // blokujemy
            Application.LockApp(1);//(513, 5000, null, null, null, null);
            // logujemy się do podanej Firmy, na danego operatora, do podanych modułów
            Login = Application.Login(vUser: Operator, vChk: Haslo, vFirm: Firma, _ModulKP: hPar[0], _ModulKH: hPar[1], _ModulKHP: hPar[2], _ModulST: hPar[3], _ModulFA: hPar[4], _ModulMAG: hPar[5], _ModulPK: hPar[6], _ModulPKXL: hPar[7], _ModulCRM: hPar[8], _ModulANL: hPar[9], _ModulDET: hPar[10], _ModulBIU: hPar[11], _ModulSRW: hPar[12], _ModulOBD: hPar[13], _ModulKB: hPar[14], _ModulKBP: hPar[15], _ModulHAP: hPar[16]);

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
        #endregion

        void WyjscieZApki(string a)
        {
            Cursor.Current = Cursors.Default;
            Wylogowanie();
            if (a == "done")
            {
                MessageBox.Show("Zakończono przetwarzanie pliku.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Generowanie dokumentów zakończyło się niepowodzeniem.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            
            if(Login==null)
            {
                try
                {
                    LogowanieAutomatyczne();
                    Sesja = Login.CreateSession();
                }
                catch (COMException eo)
                {
                    if (eo.InnerException != null)
                        MessageBox.Show("Błąd logowania: " + eo.InnerException.Message, "Komunikat",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                        MessageBox.Show("Błąd logowania: " + eo.Message, "Komunikat",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception ei)
                {
                    if (ei.InnerException != null)
                        MessageBox.Show("Błąd logowania: " + ei.InnerException.Message, "Komunikat",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                        MessageBox.Show("Błąd logowania: " + ei.Message, "Komunikat",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                
            }

            

            #region excel5
            Cursor.Current = Cursors.WaitCursor;

            if (tB1.Text == "")
            {
                MessageBox.Show("Nie wybrano pliku arkusza.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Wylogowanie();
                return;
            }
            string filePath = tB1.Text;
            DataTable dataTable;
            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    IExcelDataReader reader;
                    reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

                    var conf = new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = true
                        }
                    };

                    var dataSet = reader.AsDataSet(conf);
                    dataTable = dataSet.Tables[0];
                }
            }
            catch(Exception en)
            {

                MessageBox.Show("Błąd pliku arkusza: "+en.Message, "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Wylogowanie();
                return;
            }

            


            //Przypisywanie danych z datasetu do listy przyszłych pozycji.
            //List<ONE032Data> listaPozycjiPz = new List<ONE032Data>();
            if (dataTable.Rows.Count == 0)
            {
                MessageBox.Show("Pusty arkusz", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Error);//błąd
                return;
            }




            #endregion

            #region dzialanie na optimie
            #region logowanie
            //Logowanie do O! bez wyświetlania okienka logowania
            Cursor.Current = Cursors.WaitCursor;

           
            #endregion




            #endregion

            dodawanieDok(dataTable, Sesja);

            // dodawanieRW(dataTable, Sesja);



            //zapisujemy
            Sesja.Save();

            WyjscieZApki("done");


        }

        #region wybranie pliku
        void WybraniePliku()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog()
            {
                InitialDirectory = Directory.GetCurrentDirectory(),
                DefaultExt = "xlsx",
                Filter = "Pliki xlsx (*.xlsx)|*.xlsx|Pliki xls (*.xls)|*.xls|Wszystkie pliki (*.*)|*.*",
            };
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tB1.Text = openFileDialog1.FileName;
                //MessageBox.Show(path);
            }
            return;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            WybraniePliku();
        }
        private void TextBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            WybraniePliku();
        }
        #endregion



        void dodawanieDok(DataTable dataTable, CDNBase.AdoSession Sesja)
        {
            try
            {

                List<DateTime> daty = new List<DateTime>();
                DateTime data1 = DateTime.MinValue;
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    //MessageBox.Show(Convert.ToString(dataTable.Rows[i][7]));
                    //if (String.IsNullOrEmpty(Convert.ToString(dataTable.Rows[i][7])) == false && Convert.ToDouble(Convert.ToString(dataTable.Rows[i][7])) > 0)
                        data1 = Convert.ToDateTime(dataTable.Rows[i][0].ToString());
                    if (!daty.Contains(data1) && data1 != DateTime.MinValue)
                    {
                        daty.Add(data1);
                    }

                }

                int k = 0;
                pBar1.Visible = true;
                // Set Minimum to 1 to represent the first file being copied.
                pBar1.Minimum = 1;
                // Set Maximum to the total number of files to copy.
                pBar1.Maximum = (daty.Count*2) + 1;
                // Set the initial value of the ProgressBar.
                pBar1.Value = 1;
                // Set the Step property to a value of 1 to represent each file being copied.
                pBar1.Step = 1;

                daty.Sort();

                bool czyBlad = false;
                foreach (DateTime data in daty)
                {
                    try
                    {
                        pBar1.PerformStep();
                        Sesja = Login.CreateSession();
                        List<string> magazyny = new List<string>();
                        List<string> serie = new List<string>();

                        List<string> magazynyRW = new List<string>();
                        List<string> serieRW = new List<string>();
                        object dummy = Type.Missing;





                        string magazyn = "";
                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            if (Convert.ToDateTime(dataTable.Rows[i][0].ToString()) == data && String.IsNullOrEmpty(Convert.ToString(dataTable.Rows[i][5])) == false && Convert.ToDouble(Convert.ToString(dataTable.Rows[i][5])) > 0)
                            {
                                magazyn = Convert.ToString(dataTable.Rows[i][2]);
                                if (!magazyny.Contains(magazyn))
                                {
                                    magazyny.Add(magazyn);
                                    serie.Add(Convert.ToString(dataTable.Rows[i][1]));
                                    //MessageBox.Show("Dodawanie kodu-> Kod: " + kod, "Komunikat",
                                    // MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                        }


                        string zly_magazyn = "";
                        bool dobre_mag = true;
                        foreach (string mag in magazyny)
                        {

                            var sqlQuery2 = "SELECT TOP (1) Mag_Symbol FROM [CDN].[Magazyny] WHERE Mag_Symbol = '" + mag + "'";
                            var wynikQuery2 = Sesja.Connection.Execute(sqlQuery2, out dummy, 0);
                            //MessageBox.Show("Kod: " + aKody , "Komunikat",
                            //  MessageBoxButtons.OK, MessageBoxIcon.Information);
                            if (wynikQuery2.EOF)
                            {
                                dobre_mag = false;
                                zly_magazyn = mag;
                                //MessageBox.Show("Czy zły kod? "+ Convert.ToString(zly_kod), "Komunikat",
                                //MessageBoxButtons.OK, MessageBoxIcon.Information);
                                break;
                            }

                        }



                        string magazynRW = "";
                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            if (Convert.ToDateTime(dataTable.Rows[i][0].ToString()) == data && String.IsNullOrEmpty(Convert.ToString(dataTable.Rows[i][4])) == false && Convert.ToDouble(Convert.ToString(dataTable.Rows[i][4])) < 0)
                            {
                                magazynRW = Convert.ToString(dataTable.Rows[i][2]);
                                if (!magazynyRW.Contains(magazynRW))
                                {
                                    magazynyRW.Add(magazynRW);
                                    serieRW.Add(Convert.ToString(dataTable.Rows[i][1]));
                                    //MessageBox.Show("Dodawanie kodu-> Kod: " + kod, "Komunikat",
                                    // MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                        }


                        string zly_magazynRW = "";
                        bool dobre_magRW = true;
                        foreach (string mag in magazynyRW)
                        {

                            var sqlQuery2 = "SELECT TOP (1) Mag_Symbol FROM [CDN].[Magazyny] WHERE Mag_Symbol = '" + mag + "'";
                            var wynikQuery2 = Sesja.Connection.Execute(sqlQuery2, out dummy, 0);
                            //MessageBox.Show("Kod: " + aKody , "Komunikat",
                            //  MessageBoxButtons.OK, MessageBoxIcon.Information);
                            if (wynikQuery2.EOF)
                            {
                                dobre_magRW = false;
                                zly_magazynRW = mag;
                                //MessageBox.Show("Czy zły kod? "+ Convert.ToString(zly_kod), "Komunikat",
                                //MessageBoxButtons.OK, MessageBoxIcon.Information);
                                break;
                            }

                        }



                        if (dobre_mag == true)
                        {
                            int s = 0;
                            foreach (string mag_symbol in magazyny)
                            {

                                czyBlad = dodawaniePW(dataTable, Sesja, mag_symbol, data, serie[s]);

                                
                                s++;
                            }

                            
                        }
                        else
                        {
                            MessageBox.Show("Magazyn " + zly_magazyn + " nie istnieje w bazie. Nie importowano dokumentu dla daty " + data.ToShortDateString(), "Komunikat",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            czyBlad = true;

                        }


                        if (czyBlad)
                        {
                            pBar1.Visible = false;
                            return;
                        }

                        pBar1.PerformStep();

                        Sesja = Login.CreateSession();


                        if (dobre_magRW == true)
                        {
                            int s = 0;
                            foreach (string mag_symbol in magazynyRW)
                            {

                                czyBlad=dodawanieRW(dataTable, Sesja, mag_symbol, data, serieRW[s]);


                                s++;
                            }


                        }
                        else
                        {
                            MessageBox.Show("Magazyn " + zly_magazynRW + " nie istnieje w bazie. Nie importowano dokumentu dla daty " + data.ToShortDateString(), "Komunikat",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            czyBlad = true;

                        }

                       // dodawanieAsysta(Sesja);


                        k++;
                        Sesja.Save();
                        Sesja.Dispose();



                    }

                    catch (Exception e)
                    {

                        MessageBox.Show("Błąd: " + e.Message + "\nNie dodano dokumentów dla daty: " + data.ToShortDateString(), "Komunikat",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        czyBlad = true;

                    }

                    if (czyBlad)
                    {
                        pBar1.Visible = false;
                        return;
                    }
                }



                pBar1.Value = 1;
                pBar1.Visible = false;
               // MessageBox.Show("Zakończono generowanie PW.", "Komunikat",
                //   MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (AccessViolationException ex)
            {
                MessageBox.Show("Błąd Access Violation: " + ex.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }




        }

        void dodawanieAsysta(CDNBase.AdoSession Sesja)
        {
            CDNHlmn.DokumentyHaMag dokumentyRW = (CDNHlmn.DokumentyHaMag)Sesja.CreateObject("CDN.DokumentyHaMag", null);
            var dokumentRW = (CDNHlmn.IDokumentHaMag)dokumentyRW.AddNew(null);



            dokumentRW.Rodzaj = 304000;

            dokumentRW.TypDokumentu = 304;


            var rMagazyny = (CDNBase.ICollection)(Sesja.CreateObject("CDN.MAGAZYNY", null));
            var rMagazyn = (CDNHlmn.Magazyn)rMagazyny["Mag_Symbol='PMG Wierz.Handlowy'"];


           // dokumentRW.MagazynDocelowy = rMagazyn;
            dokumentRW.MagazynZrodlowy = rMagazyn;


            var rNumerator = dokumentRW.Numerator;


            rNumerator.Rejestr = "MAG";

            var kontrahenci = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kontrahenci", null));
            var rKontrahent = (CDNHeal.IKontrahent)kontrahenci["Knt_Kod='!NIEOKREŚLONY!'"];

            dokumentRW.Podmiot = rKontrahent;



            dokumentRW.Bufor = 0;



            CDNBase.ICollection pozycje = dokumentRW.Elementy;
            var rPozycja = (CDNHlmn.IElementHaMag)pozycje.AddNew(null);

            var zTowary = (CDNTwrb1.Towary)(Sesja.CreateObject("CDN.Towary", null));
            CDNTwrb1.ITowar zTowar = (CDNTwrb1.ITowar)zTowary["Twr_Kod='G001_PMGW-H'"];
            //MessageBox.Show(zTowar.Kod, "Komunikat",
            //MessageBoxButtons.OK, MessageBoxIcon.Error);
            rPozycja.TowarKod = zTowar.Kod;


            rPozycja.Ilosc = 1;



            dokumentRW.Regeneracja = 1;



            Sesja.Save();
        }

        bool dodawaniePW(DataTable dataTable, CDNBase.AdoSession Sesja, string mag_symbol, DateTime data, string seria)
        {
            object dummy = Type.Missing;
            List<string> aKody = new List<string>();
            string kody = "";
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                if (Convert.ToDateTime(dataTable.Rows[i][0].ToString()) == data && Convert.ToString(dataTable.Rows[i][2]) == mag_symbol && String.IsNullOrEmpty(Convert.ToString(dataTable.Rows[i][5])) == false && Convert.ToDouble(Convert.ToString(dataTable.Rows[i][5])) > 0)
                {
                    kody = Convert.ToString(dataTable.Rows[i][3]);
                    if (!aKody.Contains(kody))
                    {
                        aKody.Add(kody);
                        //MessageBox.Show("Dodawanie kodu-> Kod: " + kod, "Komunikat",
                        // MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }


            string zly_kod = "";
            bool dobre_kody = true;
            foreach (string kod in aKody)
            {

                var sqlQuery2 = "SELECT TOP (1) Twr_Kod FROM [CDN].[Towary] WHERE Twr_Kod = '" + kod + "'";
                var wynikQuery2 = Sesja.Connection.Execute(sqlQuery2, out dummy, 0);
                //MessageBox.Show("Kod: " + aKody , "Komunikat",
                //  MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (wynikQuery2.EOF)
                {
                    dobre_kody = false;
                    zly_kod = kod;
                    //MessageBox.Show("Czy zły kod? "+ Convert.ToString(zly_kod), "Komunikat",
                    //MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
                }

            }
            if (dobre_kody == true)
            {

                CDNHlmn.DokumentyHaMag PeZety = (CDNHlmn.DokumentyHaMag)Sesja.CreateObject("CDN.DokumentyHaMag", null);
                var PeZet = (CDNHlmn.IDokumentHaMag)PeZety.AddNew(null);
                var zTowary = (CDNTwrb1.Towary)(Sesja.CreateObject("CDN.Towary", null));

                //sKnt_Kod=knt_kod[k] -> odkomentować w głównym programie
                var kontrahenci = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kontrahenci", null));
                var kontrahent = (CDNHeal.IKontrahent)kontrahenci["Knt_Kod='!NIEOKREŚLONY!'"];   //Tutaj kod ostrawki

                var rMagazyny = (CDNBase.ICollection)(Sesja.CreateObject("CDN.MAGAZYNY", null));
                var rMagazyn = (CDNHlmn.Magazyn)rMagazyny["Mag_Symbol='" + mag_symbol + "'"];

                //MessageBox.Show("Dodawanie PW dla magazynu: " + rMagazyn.Nazwa + " w dniu " + data.ToShortDateString());

                //dodac typ PW
                PeZet.Rodzaj = 303000;
                PeZet.TypDokumentu = 303;
                PeZet.Bufor = 0;

                var rNumerator = PeZet.Numerator;
                rNumerator.Rejestr = seria;




                // PeZet.NumerObcy = "";

                //Ustawiamy date
                PeZet.DataDok = data;
                PeZet.DataWys = data;
                PeZet.Podmiot = (CDNHeal.IPodmiot)kontrahent;


                PeZet.MagazynDocelowy = rMagazyn;
                PeZet.MagazynZrodlowy = rMagazyn;


                //Sesja.Save();

              //  var id_PZ = PeZet.ID;

                

                for (int a = 0; a < dataTable.Rows.Count; a++)
                {
                    try
                    {
                       // MessageBox.Show("Sprawdzanie pozycji mag: " + Convert.ToString(dataTable.Rows[a][2]) + " = " + mag_symbol + "\nData: " + Convert.ToDateTime(dataTable.Rows[a][0].ToString()) + " = " +
                        //    data + "\nIlość: " + dataTable.Rows[a][4].ToString());
                        if (Convert.ToDateTime(dataTable.Rows[a][0].ToString()) == data && Convert.ToString(dataTable.Rows[a][2]) == mag_symbol && String.IsNullOrEmpty(Convert.ToString(dataTable.Rows[a][5])) == false && Convert.ToDouble(Convert.ToString(dataTable.Rows[a][5])) > 0)
                        {
                            try
                            {
                              //  MessageBox.Show("Dodawanie pozycji mag: " + Convert.ToString(dataTable.Rows[a][2]) + " = " + mag_symbol + "\nData: " + Convert.ToDateTime(dataTable.Rows[a][0].ToString()) + " = " +
                          // data + "\nIlość: " + dataTable.Rows[a][4].ToString());

                                //Sesja = Login.CreateSession();
                                //PeZet = Sesja.CreateObject("CDN.DokumentyHaMag").Item("TrN_TrNId=" + id_PZ);
                                CDNBase.ICollection pozycje = PeZet.Elementy;
                                    var pozycja = (CDNHlmn.IElementHaMag)pozycje.AddNew(null);

                                    //przypisywanie wartości pozycji
                                    CDNTwrb1.ITowar zTowar = (CDNTwrb1.ITowar)zTowary["Twr_Kod='" + Convert.ToString(dataTable.Rows[a][3]) + "'"];
                                    //MessageBox.Show(zTowar.Kod, "Komunikat",
                                    //MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    pozycja.TowarKod = zTowar.Kod;
                                    pozycja.Ilosc = Convert.ToDouble(Convert.ToString(dataTable.Rows[a][4]));
                                //pozycja.Cena0 = Convert.ToDecimal(Convert.ToString(dataTable.Rows[a][7]));
                                //pozycja.Cena0WD = Convert.ToDecimal(Convert.ToString(dataTable.Rows[a][7]));
                                pozycja.WartoscNetto = Convert.ToDecimal(Convert.ToString(dataTable.Rows[a][5]));
                                pozycja.WartoscBrutto = Convert.ToDecimal(Convert.ToString(dataTable.Rows[a][5]));

                                string dea_Kod = "NR FAKTURY ZAKUPU";

                                    var ZasPozycji = pozycja.Atrybuty;
                                    var rAtrybuty = (CDNBase.ICollection)Sesja.CreateObject("CDN.DefAtrybuty", null);
                                    var Atr2 = ZasPozycji.AddNew(null);

                                    Atr2.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybuty["dea_Kod = '" + dea_Kod + "'"];
                                    Atr2.Wartosc = Convert.ToString(dataTable.Rows[a][6]);


                                   // Sesja.Save();
                                   // Sesja.Dispose();


                                


                            }
                            catch (Exception se)
                            {

                                MessageBox.Show("Błąd przy dodawaniu pozycji: " + Convert.ToString(se), "Komunikat",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return true;

                            }


                        }
                    }
                    catch (AccessViolationException ae)
                    {
                        MessageBox.Show("Błąd Access Violation dla pozycji: " + ae.Message, "Komunikat",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return true;
                    }


                }

                Sesja.Save();
                return false;


            }
            else
            {
                MessageBox.Show("Kod " + zly_kod + " nie istnieje w bazie. Nie importowano dokumentu dla daty" + data.ToShortDateString() + " i magazynu " + mag_symbol, "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return true;

            }
        }




        bool dodawanieRW(DataTable dataTable, CDNBase.AdoSession Sesja, string mag_symbol, DateTime data, string seria)
        {
            object dummy = Type.Missing;
            List<string> aKody = new List<string>();
            string kody = "";
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                if (Convert.ToDateTime(dataTable.Rows[i][0].ToString()) == data && Convert.ToString(dataTable.Rows[i][2]) == mag_symbol && String.IsNullOrEmpty(Convert.ToString(dataTable.Rows[i][4])) == false && Convert.ToDouble(Convert.ToString(dataTable.Rows[i][4])) < 0)
                {
                    kody = Convert.ToString(dataTable.Rows[i][3]);
                    if (!aKody.Contains(kody))
                    {
                        aKody.Add(kody);
                        //MessageBox.Show("Dodawanie kodu-> Kod: " + kod, "Komunikat",
                        // MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }


            string zly_kod = "";
            bool dobre_kody = true;
            foreach (string kod in aKody)
            {

                var sqlQuery2 = "SELECT TOP (1) Twr_Kod FROM [CDN].[Towary] WHERE Twr_Kod = '" + kod + "'";
                var wynikQuery2 = Sesja.Connection.Execute(sqlQuery2, out dummy, 0);
                //MessageBox.Show("Kod: " + aKody , "Komunikat",
                //  MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (wynikQuery2.EOF)
                {
                    dobre_kody = false;
                    zly_kod = kod;
                    //MessageBox.Show("Czy zły kod? "+ Convert.ToString(zly_kod), "Komunikat",
                    //MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
                }

            }
            if (dobre_kody == true)
            {
                try
                {
                    CDNHlmn.DokumentyHaMag PeZety = (CDNHlmn.DokumentyHaMag)Sesja.CreateObject("CDN.DokumentyHaMag", null);
                    var PeZet = (CDNHlmn.IDokumentHaMag)PeZety.AddNew(null);
                    var zTowary = (CDNTwrb1.Towary)(Sesja.CreateObject("CDN.Towary", null));

                    //sKnt_Kod=knt_kod[k] -> odkomentować w głównym programie
                    var kontrahenci = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kontrahenci", null));
                    var kontrahent = (CDNHeal.IKontrahent)kontrahenci["Knt_Kod='!NIEOKREŚLONY!'"];   //Tutaj kod ostrawki

                    var rMagazyny = (CDNBase.ICollection)(Sesja.CreateObject("CDN.MAGAZYNY", null));
                    var rMagazyn = (CDNHlmn.Magazyn)rMagazyny["Mag_Symbol='" + mag_symbol + "'"];

                    //MessageBox.Show("Dodawanie PW dla magazynu: " + rMagazyn.Nazwa + " w dniu " + data.ToShortDateString());

                    //dodac typ RW
                    PeZet.Rodzaj = 304000;
                    PeZet.TypDokumentu = 304;



                    var rNumerator = PeZet.Numerator;
                    rNumerator.Rejestr = seria;



                    //PeZet.MagazynDocelowy = rMagazyn;
                    PeZet.MagazynZrodlowy = rMagazyn;

                    // PeZet.NumerObcy = "";

                    //Ustawiamy date
                    PeZet.DataDok = data;
                    PeZet.DataWys = data;
                    PeZet.Podmiot = (CDNHeal.IPodmiot)kontrahent;




                    PeZet.Bufor = 0;



                    // PeZet.Regeneracja = 1;


                    // Sesja.Save();

                    //  var id_PZ = PeZet.ID;

                    for (int a = 0; a < dataTable.Rows.Count; a++)
                    {
                        try
                        {
                            if (Convert.ToDateTime(dataTable.Rows[a][0].ToString()) == data && Convert.ToString(dataTable.Rows[a][2]) == mag_symbol && String.IsNullOrEmpty(Convert.ToString(dataTable.Rows[a][4])) == false && Convert.ToDouble(Convert.ToString(dataTable.Rows[a][4])) < 0)
                            {
                                try
                                {

                                    //Sesja = Login.CreateSession();
                                    //PeZet = Sesja.CreateObject("CDN.DokumentyHaMag").Item("TrN_TrNId=" + id_PZ);
                                    CDNBase.ICollection pozycje = PeZet.Elementy;
                                    var pozycja = (CDNHlmn.IElementHaMag)pozycje.AddNew(null);

                                    //przypisywanie wartości pozycji
                                    CDNTwrb1.ITowar zTowar = (CDNTwrb1.ITowar)zTowary["Twr_Kod='" + Convert.ToString(dataTable.Rows[a][3]) + "'"];
                                    //MessageBox.Show(zTowar.Kod, "Komunikat",
                                    //MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    pozycja.TowarKod = zTowar.Kod;
                                    pozycja.Ilosc = -Convert.ToDouble(Convert.ToString(dataTable.Rows[a][4]));
                                    //pozycja.Cena0 = Convert.ToDecimal(Convert.ToString(dataTable.Rows[a][7]));
                                    //pozycja.Cena0WD = Convert.ToDecimal(Convert.ToString(dataTable.Rows[a][7]));
                                    //pozycja.WartoscNetto = 0;
                                    //pozycja.WartoscBrutto = 0;






                                }
                                catch (Exception se)
                                {

                                    MessageBox.Show("Błąd przy dodawaniu pozycji: " + Convert.ToString(se), "Komunikat",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return true;

                                }


                            }
                        }
                        catch (AccessViolationException ae)
                        {
                            MessageBox.Show("Błąd Access Violation dla pozycji: " + ae.Message, "Komunikat",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return true;
                        }


                    }
                    PeZet.Regeneracja = 1;



                    Sesja.Save();
                    return false;
                }
                catch(Exception e)
                {
                    MessageBox.Show("Błąd dodawania RW dla magazynu: " + mag_symbol + " w dniu " + data.ToShortDateString()+" \n"+e.Message , "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return true;
                }

            }
            else
            {
                MessageBox.Show("Kod " + zly_kod + " nie istnieje w bazie. Nie importowano dokumentu RW dla daty" + data.ToShortDateString() + " i magazynu " + mag_symbol, "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return true;

            }
        }



        
    }
}
    