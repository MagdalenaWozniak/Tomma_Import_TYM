using cdn_api;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ImportTYM
{
    public partial class Form1 : Form
    {
        private int XLWersjaAPI = 20220;
        private int XLTrybWsadowy = 1;
        private int XLWinieta = -1;
        private int XLUtworzWlasnaSesje = 1;
        private int sesja = -1;
        private string XLSQLServer = "10.33.21.10";
        private string XLSQLDatabase = "TOMMA-ERP";//"ERPXL_PoloMarket"
        private string XLUser = "MACIEJ"; //
        private string XLPass = "NE!w0rk12";//"Dupa2003!";
        string connectionString = @"Data Source = COMARCHBD; Initial Catalog = ERPXL_TOMMA-ERP; User ID = sa; Password=fr$g4i9sSa032V,_";
        string date1 = "1800-12-28";

        //@"Data Source = SQLSERVER; Initial Catalog = DomarXL; User ID=AdminS;Password=6,eF#W451,t6,D@";

        public Form1()
        {
            //this.connectionString = GetConnectionString();
            InitializeComponent();
        }


        void Button1_Click(object sender, EventArgs e)
        {
            #region excel5
            #region wybranie pliku
            Cursor.Current = Cursors.WaitCursor;
            if (tB1.Text == "nie wybrano pliku")
            {
                MessageBox.Show("Nie wybrano pliku arkusza.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string filePath = tB1.Text;
            DataTable dataTable;
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader;
                reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = false//true
                    }
                };

                var dataSet = reader.AsDataSet(conf);
                dataTable = dataSet.Tables[0];
            }
            #endregion

            #region Przypisywanie danych z datasetu do listy przyszłych pozycji.

            SprawdzamCzyArkuszNieJestPusty(dataTable);
            dodajDokumenty(dataTable);

            MessageBox.Show("Zakończono dodawanie faktur zakupu TYM", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Application.Exit();
            #endregion

            #endregion


        }
        #region Dodawanie dokumetów
        void dodajDokumenty(DataTable dataTable)
        {


            XLLoginInfo_20220 nowyUzytkownik = new XLLoginInfo_20220();
            nowyUzytkownik.Wersja = XLWersjaAPI;
            nowyUzytkownik.Winieta = XLWinieta;
            nowyUzytkownik.TrybWsadowy = XLTrybWsadowy;
            nowyUzytkownik.UtworzWlasnaSesje = XLUtworzWlasnaSesje;
            nowyUzytkownik.SerwerKlucza = XLSQLServer;
            nowyUzytkownik.Baza = XLSQLDatabase;
            nowyUzytkownik.OpeIdent = XLUser;
            nowyUzytkownik.OpeHaslo = XLPass;
            nowyUzytkownik.PortHTTPSerweraKlucza = 1;
            nowyUzytkownik.RodzajSerweraKlucza = 0;
            nowyUzytkownik.SerwerKlucza = "COMARCHBD::5000024502";
            nowyUzytkownik.ProgramID = "ONE-023 Import TYM";

            int blad = cdn_api.cdn_api.XLLogin(nowyUzytkownik, ref sesja);
            if (blad != 0)
            {
                MessageBox.Show("Błąd logowania: " + blad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            
            pBar1.Visible = true;
            // Set Minimum to 1 to represent the first file being copied.
            pBar1.Minimum = 1;
            // Set Maximum to the total number of files to copy.
            pBar1.Maximum = dataTable.Rows.Count;
            // Set the initial value of the ProgressBar.
            pBar1.Value = 1;
            // Set the Step property to a value of 1 to represent each file being copied.
            pBar1.Step = 1;


            for (int i = 1; i < dataTable.Rows.Count; i++)
            {
                try
                {
                    pBar1.PerformStep();

                    string date2 = Convert.ToString(dataTable.Rows[i][5]).Substring(6, 4) + "-" + Convert.ToString(dataTable.Rows[i][5]).Substring(3, 2) + "-" + Convert.ToString(dataTable.Rows[i][5]).Substring(0, 2);
                   // MessageBox.Show("Data w Excel: "+Convert.ToString(dataTable.Rows[i][5]));
                    //MessageBox.Show("Data po convert: " + date2);
                    //MessageBox.Show("Rok: " + (dataTable.Rows[i][5]).ToString().Substring(3,2));
                   // MessageBox.Show("Miesiąc: " + dataTable.Rows[i][5].ToString().Substring(0, 2));
                    string query5 = "SELECT DATEDIFF(day, '" + date1 + "','" + date2 + "'); ";


                    XLDokumentNagInfo_20220 nowyDokument = new XLDokumentNagInfo_20220();
                    nowyDokument.Wersja = XLWersjaAPI;
                    nowyDokument.GIDFirma = 0;
                    nowyDokument.GIDLp = 0;
                    nowyDokument.GIDNumer = 0;
                    nowyDokument.GIDTyp = 0;
                    nowyDokument.Akronim = dataTable.Rows[i][1].ToString();
                    //nowyDokument.KntFirma = Convert.ToInt32(ConnectToSQL(connectionString, query));
                    //nowyDokument.KntLp = Convert.ToInt32(ConnectToSQL(connectionString, query2));
                    //nowyDokument.KntNumer = Convert.ToInt32(ConnectToSQL(connectionString, query3));
                    //nowyDokument.KntTyp = Convert.ToInt32(ConnectToSQL(connectionString, query4));
                    nowyDokument.Typ = 1; //FZ
                    nowyDokument.Avista = 1;
                    nowyDokument.Tryb = 2;
                    nowyDokument.Korekta = 0;
                    nowyDokument.Rok = Convert.ToInt32((dataTable.Rows[i][5]).ToString().Substring(6, 4));
                    nowyDokument.Miesiac = Convert.ToInt32((dataTable.Rows[i][5]).ToString().Substring(3, 2));
                    nowyDokument.Seria = "TYM";
                    nowyDokument.Rejestr = "ZAK";
                    //nowyDokument.IDOddzial = Convert.ToInt32(dataTable.Rows[i][2]);
                    nowyDokument.DataSpr = Convert.ToInt32(ConnectToSQL(connectionString, query5));
                    nowyDokument.Data= Convert.ToInt32(ConnectToSQL(connectionString, query5)); 
                    nowyDokument.DataMag= Convert.ToInt32(ConnectToSQL(connectionString, query5)); 
                    nowyDokument.DataVat= Convert.ToInt32(ConnectToSQL(connectionString, query5)); 
                    int dokumentId = 0;
                    blad=cdn_api.cdn_api.XLNowyDokument(sesja, ref dokumentId, nowyDokument);
                    if(blad!=0)
                    {
                        if(blad==8)
                        {
                            MessageBox.Show("Brak kontrahenta: "+ dataTable.Rows[i][1].ToString() , "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        else
                            MessageBox.Show("Błąd XLNowuDokument w wierszu " + i + ": " + blad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        
                    }
                    

                    blad = dodajPozycje(dataTable, i, dokumentId);
                    if (blad != 0)
                    {
                        MessageBox.Show("Błąd Dodawania pozycji w wierszu " + i + ": " + blad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    }


                    XLZamkniecieDokumentuInfo_20220 zamkniecie = new XLZamkniecieDokumentuInfo_20220();
                    zamkniecie.Wersja = XLWersjaAPI;
                    zamkniecie.Tryb = 0;
                    cdn_api.cdn_api.XLZamknijDokument(dokumentId, zamkniecie);


                    if (dataTable.Rows[i][10].ToString() != "")
                    {
                        int opisID = 0;
                        blad = dodajOpis(ref opisID, nowyDokument);
                        if (blad != 0)
                        {
                            MessageBox.Show("Błąd Dodawania opisu w wierszu " + i + ":" + blad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        }


                        blad = dodajLinie(dataTable, i, opisID);
                        if (blad != 0)
                        {
                            MessageBox.Show("Błąd Dodawania linii w wierszu " + i + ": " + blad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        }


                        while (i + 1 < dataTable.Rows.Count && dataTable.Rows[i][0].ToString() == dataTable.Rows[i + 1][0].ToString())
                        {
                            i++;
                            pBar1.PerformStep();
                            blad = dodajLinie(dataTable, i, opisID);
                            if (blad != 0)
                            {
                                MessageBox.Show("Błąd Dodawania linii w wierszu " + i + ": " + blad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                            }

                        }

                        XLZamknijOpisInfo_20220 zamOpis = new XLZamknijOpisInfo_20220();
                        zamOpis.Wersja = XLWersjaAPI;
                        cdn_api.cdn_api.XLZamknijOpis(opisID, zamOpis);
                    }


                    XLAtrybutInfo_20220 atr = new XLAtrybutInfo_20220();
                    atr.Wersja = XLWersjaAPI;
                    atr.GIDFirma = nowyDokument.GIDFirma;
                    atr.GIDLp = nowyDokument.GIDLp;
                    atr.GIDNumer = nowyDokument.GIDNumer;
                    atr.GIDSubLp = 0;
                    atr.GIDTyp = nowyDokument.GIDTyp;
                    atr.Klasa = "Rezerwa rozliczona";
                    atr.Wartosc = "NIE";
                    blad = cdn_api.cdn_api.XLDodajAtrybut(sesja, atr);
                    if (blad != 0)
                    {
                        MessageBox.Show("Błąd atrybutu w wierszu " + i + ": " + blad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    }

                }
                catch(Exception e)
                {
                    MessageBox.Show("Błąd dodawania dokumentu: " + e.Message, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }


            }


            cdn_api.cdn_api.XLLogout(sesja);


        }

        int dodajPozycje(DataTable dataTable, int i, int dokID)
        {
            XLVatInfo_20220 vat = new XLVatInfo_20220();

            vat.Wersja = XLWersjaAPI;
            vat.StawkaPod = Convert.ToInt32(Convert.ToDouble(dataTable.Rows[i][4].ToString()) * 100);
            vat.Wartosc = Math.Round(Convert.ToDouble(dataTable.Rows[i][2].ToString()), 2, MidpointRounding.AwayFromZero).ToString();
            vat.WartoscVat = Math.Round(Convert.ToDouble(dataTable.Rows[i][3].ToString()), 2, MidpointRounding.AwayFromZero).ToString();
            vat.DeklRok = Convert.ToInt32((dataTable.Rows[i][5]).ToString().Substring(6, 4));
            vat.DeklMiesiac = Convert.ToInt32((dataTable.Rows[i][5]).ToString().Substring(3, 2));
            vat.IgnorujVat = 1;
            vat.OdliczeniaVat = 2;
            return cdn_api.cdn_api.XLDodajVAT(dokID, vat);
           

        }


        int dodajOpis(ref int opisID, XLDokumentNagInfo_20220 dok)
        {
            XLNowyOpisInfo_20220 opis = new XLNowyOpisInfo_20220();
            opis.GIDFirma = dok.GIDFirma;
            opis.GIDLp = dok.GIDLp;
            opis.GIDNumer = dok.GIDNumer;
            opis.GIDTyp = dok.GIDTyp;
            opis.Wersja = XLWersjaAPI;
            opis.Tryb = 2;
            return cdn_api.cdn_api.XLNowyOpis(sesja, ref opisID, opis);
            

        }

        int dodajLinie(DataTable dataTable, int i, int opisID)
        {
            XLDodajLinieOpisuInfo_20220 linia = new XLDodajLinieOpisuInfo_20220();
            linia.Wersja = XLWersjaAPI;
            linia.Tryb = 2;
            linia.Lp = 0;
            linia.Wartosc = Math.Round(Convert.ToDouble(dataTable.Rows[i][10].ToString()), 2, MidpointRounding.AwayFromZero).ToString();
            int blad=cdn_api.cdn_api.XLDodajLinieOpisu(opisID, linia);

            if(dataTable.Rows[i][7].ToString() != "")
            {
                XLDodajWymiarOpisuInfo_20220 kat = new XLDodajWymiarOpisuInfo_20220();
                kat.Wersja = XLWersjaAPI;
                kat.Tryb = 2;
                kat.Typ = 0;
                kat.Wymiar = 23;
                kat.Element = dataTable.Rows[i][7].ToString();
                int bladWym = cdn_api.cdn_api.XLDodajWymiarOpisu(opisID, kat);
                if(bladWym!=0)
                {
                    MessageBox.Show("Błąd Dodawania wymiaru kategoria kosztu w wierszu " + i + ": " + bladWym, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            if (dataTable.Rows[i][8].ToString() != "")
            {
                try
                {
                    string rodzajQuery = "select WMR_ID from cdn.Wymiary where WMR_ParID = 12 and WMR_Nazwa = '"+ dataTable.Rows[i][8].ToString()+"'";
                    string rodzaj = ConnectToSQL(connectionString, rodzajQuery);
                    if (rodzaj != "")
                    {
                        XLDodajWymiarOpisuInfo_20220 bad = new XLDodajWymiarOpisuInfo_20220();
                        bad.Wersja = XLWersjaAPI;
                        bad.Tryb = 2;
                        bad.Typ = 0;
                        bad.Wymiar = Convert.ToInt32(rodzaj);
                        //bad.Element = dataTable.Rows[i][8].ToString();
                        int bladWym = cdn_api.cdn_api.XLDodajWymiarOpisu(opisID, bad);
                        if (bladWym != 0)
                        {
                            MessageBox.Show("Błąd Dodawania wymiaru rodzaj badania w wierszu " + i + ": " + bladWym, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Brak definicji wymiaru dla badania: "+ dataTable.Rows[i][8].ToString(), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch(Exception e)
                {
                    MessageBox.Show("Błąd Dodawania wymiaru rodzaj badania w wierszu " + i + ": " + e.Message, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

                if (dataTable.Rows[i][9].ToString() != "")
                {
                    try { 
                        string rodzajQuery = "select WMR_ID from cdn.Wymiary where WMR_ParID = 6 and WMR_Nazwa = '" + dataTable.Rows[i][9].ToString() + "'";
                        string rodzaj = ConnectToSQL(connectionString, rodzajQuery);
                        if (rodzaj != "")
                        {
                            XLDodajWymiarOpisuInfo_20220 cecha = new XLDodajWymiarOpisuInfo_20220();
                            cecha.Wersja = XLWersjaAPI;
                            cecha.Tryb = 2;
                            cecha.Typ = 0;
                            cecha.Wymiar = Convert.ToInt32(rodzaj);
                            //cecha.Element = dataTable.Rows[i][9].ToString();
                            int bladWym = cdn_api.cdn_api.XLDodajWymiarOpisu(opisID, cecha);
                            if (bladWym != 0)
                            {
                                MessageBox.Show("Błąd Dodawania wymiaru cecha kosztów w wierszu " + i + ": " + bladWym, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Brak definicji wymiaru dla cechy: " + dataTable.Rows[i][9].ToString(), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("Błąd Dodawania wymiaru cecha kosztu w wierszu " + i + ": " + e.Message, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }


            if (dataTable.Rows[i][6].ToString() != "")
            {
                try
                {
                    string rodzajQuery = "SELECT [FRS_ID] FROM [CDN].[FrmStruktura] WHERE FRS_Oddzial = 0 and FRS_Nazwa like '%"
                                        + dataTable.Rows[i][6].ToString().Substring(dataTable.Rows[i][6].ToString().IndexOf(".")+1).Trim() + "%'";
                    string rodzaj = ConnectToSQL(connectionString, rodzajQuery);
                    if (rodzaj != "")
                    {

                 //  MessageBox.Show("Rodzaj: " +rodzaj);
                        XLDodajWymiarOpisuInfo_20220 oddzial = new XLDodajWymiarOpisuInfo_20220();
                        oddzial.Wersja = XLWersjaAPI;
                        oddzial.Tryb = 2;
                        oddzial.Typ = 1;
                        oddzial.Wymiar = Convert.ToInt32(rodzaj);//1600;
                        //oddzial.Sciezka = "Tomma Diagnostyka Obrazowa.ODDZIAŁY.227-BIELSKO-BIAŁA";//+dataTable.Rows[i][6].ToString().Trim();
                        int bladWym = cdn_api.cdn_api.XLDodajWymiarOpisu(opisID, oddzial);
                        if (bladWym != 0)
                        {
                            MessageBox.Show("Błąd Dodawania wymiaru oddział w wierszu " + i + ": " + bladWym, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Brak definicji wymiaru dla oddziału: " + dataTable.Rows[i][6].ToString(), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Błąd Dodawania wymiaru Oddział w wierszu " + i + ": " + e.Message, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }


            




            XLZamknijLinieOpisuInfo_20220 zamLin = new XLZamknijLinieOpisuInfo_20220();
            zamLin.Wersja = XLWersjaAPI;
            cdn_api.cdn_api.XLZamknijLinieOpisu(opisID, zamLin);

            return blad;
            
        }



        public string ConnectToSQL(string connection_string, string query)
        {
            string str = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, "Query");
                    if (dataSet.Tables.Count == 0)
                    {
                        return "";
                    }
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                    {
                        str = row[0].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                str = str + "Bład SQL:" + ex.ToString() + "\n";
            }
            return str;
        }
        #endregion

        #region funkcje wybranie pliku
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

        #region funkcje pobranie danych z pliku
        bool SprawdzamCzyArkuszNieJestPusty(DataTable _dataTable)
        {
            if (_dataTable.Rows.Count == 0)
            {
                MessageBox.Show("Wybrany arkusz pusty.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Error);//błąd
                return true;
            }
            return false;
        }
        #endregion

        private void title1_Click(object sender, EventArgs e)
        {

        }

    }
}

