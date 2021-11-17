using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Org.BouncyCastle.Crypto;
using System.IO;
using iText.Signatures;
using iText.Kernel.Pdf;
using iText.Forms.Fields;
using iText.Forms;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography;
using iText.Kernel.Pdf.Annot;
using iText.Kernel.Pdf.Action;
using iText.Layout;
using iText.Layout.Element;
using iText.Kernel.Font;
using iText.IO.Font.Constants;
using iText.IO.Font;
using iText.StyledXmlParser.Resolver.Font;
using iText.Kernel.Colors;
using iText.Kernel.Pdf.Filespec;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Geom;
using System.Reflection;
using System.Linq;

namespace OknoStartowe
{
    static class Program
    {
        /// <summary>
        /// Główny punkt wejścia dla aplikacji.
        /// Parametry wejściowe:
        /// 0. Zapis / odczyt
        /// 1. Sciezka do pliku
        /// (2). hasło do certyfikatu ---- WYCOFANE
        /// 2(3). rodzaj podpisu: konstr/spr/zatw
        /// 3(4) wysokosc polozenia pola podpisu
        /// </summary>
        [STAThread]
        static void Main(string[] Arg)
        {
            //foreach (string item in Arg)
            //{
            //    File.AppendAllText(@"C:\TMP\log.txt", Environment.NewLine + item); 
            //}
            if (Arg.Length > 0 && Arg[0].EndsWith(".pdf") && File.Exists(Arg[0]))
            {
                UsuwaniePodpisow(Arg);
                return;
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            ZaczytajUstawienia();
            //START_LINKA_TEST();
            //TrescPDF(@"E:\PDFOutput-lstoklosa.pdf");
            //Statystyka.ZPliku();
            //return;
            if (Arg.Length == 1 && Arg[0] == "Serwer")
            {
                Application.Run(new FormCzuwania());
                //uruchom jako serwer i nasluchuj
            }
            else if (Arg.Length < 2)
            {
                MessageBox.Show("Zbyt mało parametrów", "Podpisywanie PDF", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            else if (Arg[0] == "read")
            {
                Odczyt(Arg);
            }
            else if (Arg[0] == "write" && Arg.Length < 4)
            {
                MessageBox.Show("Zbyt mało parametrów do wykonania podpisu", "Podpisywanie PDF", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            else if (Arg[0] == "write")
            {
                Podpisywanie(Arg[1], Arg[2], float.Parse(Arg[3]));
            }
            else if (Arg[0] == "PDFlink")
            {
                dodajLinkaDoPDF(Arg[1], Arg[2]);
            }
            else if (Arg[0] == "test")
            {
                MessageBox.Show("Udany");
            }
            else if (Arg[0] == "split")
            {
                TrescPDF(Arg[1]);
            }
            else if (Arg[0] == "create")
            {
                UtworzPDF.NowyPDF(Arg[1], Arg[2]);
            }
            else if (Arg[0] == "print")
            {
                Drukowanie.DrukujPlik(Arg[1], Arg[2]);
            }
            else if (Arg[0] == "stay" && Arg[1] == "awake")
            {
                FormCzuwania.Czuwanie();
            }
            else if (Arg[0] == "ReadSign" && File.Exists(Arg[1]))
            {
                OdczytPodpisow.PodpisanePrzez(Arg[1]);
            }
            //MessageBox.Show("Koniec");
            return;
            //Application.Run(new Form1());
        }
        public enum Podpis { Projektowal, Sprawdzil, Zatwierdzil }
        public static string PlikUstawien = Application.StartupPath + @"\UstawieniaPodpisow.ini";
        public static string NazwaSzablonu = "";
        public static string LogPDF { get; set; } = "";
        public static string Sciezka(this string FullFileName)
        {
            int i = FullFileName.LastIndexOf(@"\");
            if (i == -1)
            {
                throw new Exception("To nie jest sciezka");
            }
            FullFileName = FullFileName.Substring(0, i + 1);
            return FullFileName;
        }
        public static string NazwaPliku(this string FullFileName, bool ZRozszerzeniem)
        {
            int i = FullFileName.LastIndexOf(@"\");
            FullFileName = FullFileName.Substring(i + 1);
            if (ZRozszerzeniem == true)
            {
                return FullFileName;
            }
            i = FullFileName.LastIndexOf(".");
            if (i == -1)
            {
                return FullFileName;
            }
            FullFileName = FullFileName.Substring(0, i);
            return FullFileName;
        }
        public static void ZaczytajUstawienia()
        {
            List<string> Ustawienia = File.ReadAllLines(PlikUstawien).ToList();
            LogPDF = Ustawienia[0].Split('|')[1];
            NazwaSzablonu = Ustawienia[1].Split('|')[1];
        }
        public static void Podpisywanie(string SciezkaDoPDF, string PowodPodpisania, float Wysokosc)
        {
            X509Store store = null;
            RSA rSA = null;
            try
            {
                PdfReader reader = new PdfReader(SciezkaDoPDF);

                store = new X509Store(StoreName.My);
                store.Open(OpenFlags.ReadOnly);
                X509Certificate2 x509Certificate = null;

                X509Certificate2Collection ZbiorCert = (X509Certificate2Collection)store.Certificates;
                X509Certificate2Collection Pofiltrowane = ZbiorCert.Find(X509FindType.FindByTemplateName, NazwaSzablonu, true);
                if (Pofiltrowane.Count > 0)
                {
                    x509Certificate = Pofiltrowane[0];
                }
                else
                {
                    throw new Exception("Nie znaleziono odpowiedniego certyfikatu.");
                }

                //foreach (var certificate in store.Certificates)
                //{
                //    TimeSpan timeSpan = certificate.NotAfter - certificate.NotBefore;
                //    //X509ExtensionCollection collection = (X509ExtensionCollection)certificate.Extensions.SyncRoot;
                //    //string TMP = collection[0].Oid.Value;
                //    //MessageBox.Show(TMP);

                //    if (timeSpan.Days == 365 && certificate.Issuer == "CN=WISS-ROOT-CA-ETA1, DC=wiss, DC=com, DC=pl")
                //    {
                //        x509Certificate = certificate;
                //        break;
                //    }
                //    if (store.Certificates.IndexOf(certificate) == store.Certificates.Count - 1) { throw new Exception("Nie znaleziono odpowiedniego certyfikatu."); }
                //}

                rSA = x509Certificate.GetRSAPrivateKey();
                AsymmetricCipherKeyPair akp = Org.BouncyCastle.Security.DotNetUtilities.GetKeyPair(rSA);

                AsymmetricKeyParameter pk_zProvidera = akp.Private;

                Org.BouncyCastle.X509.X509Certificate x509Przekonwertowany = Org.BouncyCastle.Security.DotNetUtilities.FromX509Certificate(x509Certificate);

                Org.BouncyCastle.X509.X509Certificate[] chain = new Org.BouncyCastle.X509.X509Certificate[] { x509Przekonwertowany };

                string DEST = "H:\\TMP.pdf";

                PdfDocument pdfDocument = new PdfDocument(reader);
                float PolozenieX = 0;
                PdfPage pdfPage = pdfDocument.GetFirstPage();
                iText.Kernel.Geom.Rectangle rectangle = pdfPage.GetPageSize();
                if (rectangle.GetRight() == 1191)
                {
                    PolozenieX = 835 - 40;
                }
                else
                {
                    PolozenieX = 240 - 40;
                }
                pdfDocument.Close();
                reader = new PdfReader(SciezkaDoPDF);

                //PdfSigner signer = new PdfSigner(reader, new FileStream(DEST, FileMode.Create), new StampingProperties() );
                //StampingProperties stampingProperties = new StampingProperties();
                PdfSigner signer = new PdfSigner(reader, new FileStream(DEST, FileMode.Create), true);
                //signer = new PdfSigner(reader, new FileStream(DEST, FileMode.Create),);
                //pole 1: .SetPageRect(new iText.Kernel.Geom.Rectangle(240, 142, 120, 30))
                //pole 2: .SetPageRect(new iText.Kernel.Geom.Rectangle(240, 142-15, 120, 30))
                //pole 3: .SetPageRect(new iText.Kernel.Geom.Rectangle(240, 142-30, 120, 30))
                PdfSignatureAppearance appearance = signer.GetSignatureAppearance();
                string Podpisujacy = chain[0].SubjectDN.ToString();
                string[] RozbityWiersz = Podpisujacy.Split(',');
                if (RozbityWiersz.Length == 8)
                {
                    Podpisujacy = RozbityWiersz[6].Replace("CN=", "");
                }

                appearance.SetReason(PowodPodpisania)
                    .SetPageNumber(1)
                    .SetPageRect(new iText.Kernel.Geom.Rectangle(PolozenieX, Wysokosc, 200 + 40, 30))
                    .SetLocation("Bielsko-Biała");
                signer.SetFieldName(PowodPodpisania);

                appearance.SetLayer2Text($"{PowodPodpisania}: {Podpisujacy} Data: {signer.GetSignDate()}");
                appearance.SetLayer2FontSize(8);
                IExternalSignature pks = new PrivateKeySignature(pk_zProvidera, DigestAlgorithms.SHA512);

                //ICollection<ICrlClient> crlList = new List<ICrlClient> { new CrlClientOnline("") };
                //ICrlClient crl = new CrlClientOnline(@"");

                //IOcspClient ocsp = new OcspClientBouncyCastle();

                //ICollection<byte[]> TabelaCLR = crlClient.GetEncoded(x509Przekonwertowany, "");

                signer.SignDetached(pks, chain, null, null, null, 0, PdfSigner.CryptoStandard.CMS);
                reader.Close();
                File.Move(SciezkaDoPDF, SciezkaDoPDF + "1");
                File.Move(DEST, SciezkaDoPDF);
                ZmianaNaReadOnly(SciezkaDoPDF + "1", false);
                File.Delete(SciezkaDoPDF + "1");

                string ZebranePodpisy = OdczytPojedynczegoPDF(SciezkaDoPDF);
                ZapisLogu(SciezkaDoPDF.ToUpper(), ZebranePodpisy);
                Console.Write(ZebranePodpisy);
            }
            catch (Exception ex)
            {
                if (ex.Message == "Field has been already signed.")
                {
                    MessageBox.Show($"Dla pliku:\n{SciezkaDoPDF}\nw polu {PowodPodpisania} widnieje już podpis", "Podpis PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(ex.Message, "Podpis PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            finally
            {
                try
                {
                    rSA.Clear();
                    store.Close();
                }
                catch { }
            }
        }
        private static void ZapisLogu(string SciezkaPDF, string Wpis)
        {
            string NazwaPlikuTXT = SciezkaPDF.NazwaPliku(true).Substring(0, 7) + ".txt";

            List<string> TrescPliku = new List<string>();
            if (File.Exists(LogPDF + NazwaPlikuTXT))
            {
                TrescPliku = File.ReadAllLines(LogPDF + NazwaPlikuTXT).ToList();
            }
            List<string> Znalezione = TrescPliku.Where(poz => poz.ToLower().Contains(SciezkaPDF.ToLower())).ToList();
            foreach(string E in Znalezione)
            {
                TrescPliku.Remove(E);
            }
            string NowyWpis = $"{File.GetLastWriteTime(SciezkaPDF)};{Wpis}";
            TrescPliku.Add(NowyWpis);
            File.WriteAllLines(LogPDF + NazwaPlikuTXT, TrescPliku);
        }
        public static bool ZmianaNaReadOnly(string SciezkaDoPliku, bool SetReadOnly)
        {
            try
            {
                FileAttributes Atrybuty = System.IO.File.GetAttributes(SciezkaDoPliku);
                if ((Atrybuty & FileAttributes.ReadOnly) != FileAttributes.ReadOnly && SetReadOnly)
                {
                    File.SetAttributes(SciezkaDoPliku, File.GetAttributes(SciezkaDoPliku) | FileAttributes.ReadOnly);
                    return true;
                }
                else if (SetReadOnly == false)
                {
                    Atrybuty = RemoveAttribute(Atrybuty, FileAttributes.ReadOnly);
                    File.SetAttributes(SciezkaDoPliku, Atrybuty);
                    return true;
                }
            }
            catch { return false; }
            return false;
        }
        private static FileAttributes RemoveAttribute(FileAttributes attributes, FileAttributes attributesToRemove)
        {
            return attributes & ~attributesToRemove;
        }
        /// <summary>
        /// Odczytuje jakie rodzaje podpisów zostały wykonane na pliku PDF
        /// </summary>
        /// <param name="ListaSciezek"></param>
        public static void Odczyt(string[] ListaSciezek)
        {
            string DoZwrotu = "";
            for (int i = 1; i < ListaSciezek.Length; i++)
            {
                string TMP = OdczytPojedynczegoPDF(ListaSciezek[i]);
                DoZwrotu += TMP;
            }
            Console.Write(DoZwrotu);

        }
        public static string OdczytPojedynczegoPDF(string SciezkaPDF)
        {
            string DoZwrotu = "";
            PdfReader reader = null;
            PdfDocument pdfDocument = null;
            try
            {
                DoZwrotu = SciezkaPDF;
                reader = new PdfReader(SciezkaPDF);
                pdfDocument = new PdfDocument(reader);
                PdfPage pdfPage = pdfDocument.GetPage(1);
                PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDocument, false);
                IDictionary<string, PdfFormField> PolaNaPDF = form.GetFormFields();
                foreach (KeyValuePair<string, PdfFormField> Element in PolaNaPDF)
                {
                    if (Element.Value is PdfSignatureFormField)
                    {
                        DoZwrotu += "|" + Element.Value.GetFieldName().ToString();
                        //PdfObject TMP = Element.Value.();
                        string aaa = Element.Value.GetValueAsString();
                    }
                }
                DoZwrotu += ";";
                reader.Close();
            }
            catch { DoZwrotu = ""; }
            finally
            {
                try
                {
                    reader.Close();
                    pdfDocument.Close();
                }
                catch { }
            }
            return DoZwrotu;
        }
        private static void dodajLinkaDoPDF(string SciezkaPDF, string SciezkaDoLinka)
        {
            try
            {
                //string TMPdoc = @"H:\TMP.pdf";
                string TMPdoc = SciezkaPDF + "-";
                PdfReader reader = new PdfReader(SciezkaPDF);
                PdfDocument pdfDocSource = new PdfDocument(reader);

                PdfDocument pdfDocument = new PdfDocument(new PdfWriter(TMPdoc));
                Document layoutDocument = new Document(pdfDocument);

                pdfDocSource.CopyPagesTo(1, pdfDocSource.GetNumberOfPages(), pdfDocument);

                PdfFileSpec spec = PdfFileSpec.CreateExternalFileSpec(pdfDocument, SciezkaDoLinka);
                PdfAction action = PdfAction.CreateLaunch(spec);


                //Paragraph portfolioText = new Paragraph("Kliknij by otworzyć model 3D");
                Paragraph portfolioText = new Paragraph(new Link("Kliknij by otworzyc model 3D", action));
                portfolioText.SetFont(PdfFontFactory.CreateFont());
                portfolioText.SetFontColor(Color.ConvertRgbToCmyk(new DeviceRgb(System.Drawing.Color.Black)));
                portfolioText.SetFixedLeading(-60);//(100.1f);
                portfolioText.SetFirstLineIndent(1f);

                //------------------------------------------------------------
                //string KK = @"C:\Windows\Fonts\Code39.ttf";
                ////FontProgram fontProgram = FontProgramFactory.CreateFont(KK);
                //PdfFont KodKreskowy = PdfFontFactory.CreateFont(KK, true);

                //layoutDocument.Add(new Paragraph(DateTime.Now.ToString())
                //.SetFontSize(25)
                //.SetFont(KodKreskowy)
                //.SetFixedPosition(200, 200, 250))
                //.SetFontColor(Color.ConvertRgbToCmyk(new DeviceRgb(System.Drawing.Color.Black)));



                //------------------------------------------------------------

                //portfolioText.SetAction(PdfAction.CreateURI(SciezkaDoLinka));
                layoutDocument.Add(portfolioText);
                layoutDocument.Close();

                reader.Close();
                pdfDocSource.Close();
                pdfDocument.Close();

                try
                {
                    File.Delete(SciezkaPDF);
                }
                catch (Exception ex)
                {
                    File.Move(SciezkaPDF, SciezkaPDF + "_");
                }
                File.Move(TMPdoc, SciezkaPDF);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Dodawanie linka", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        private static void TrescPDF(string SciezkaPDF)
        {
            PdfReader reader = new PdfReader(SciezkaPDF);
            PdfDocument pdfDocSource = new PdfDocument(reader);
            int IlStr = pdfDocSource.GetNumberOfPages();
            List<string> OdczytaneStrony = new List<string>();
            List<PdfPage> StrPDF = new List<PdfPage>();

            for (int i = 1; i <= IlStr; i++)
            {
                PdfPage page = pdfDocSource.GetPage(i);
                Rectangle rectangle = page.GetPageSize();
                string Tresc = page.ExtractText(rectangle)[0];
                OdczytaneStrony.Add(Tresc);
                StrPDF.Add(page);
            }
            ObrobkaTresci(OdczytaneStrony, StrPDF, pdfDocSource, SciezkaPDF);
            pdfDocSource.Close();
            //MessageBox.Show("Koniec");
        }
        /// <summary>
        /// Zadaniem funkcji jest zapisanie plików pdf zakończonych znacznikiem "koniec"
        /// pod taką samą nazwą należy zapisać też plik TXT zawierający tekst z tych stron
        /// </summary>
        /// <param name="OdczytaneStrony"></param>
        /// <param name="StrPDF"></param>
        private static void ObrobkaTresci (List<string> OdczytaneStrony, List<PdfPage> StrPDF, PdfDocument pdfDocSource, string FFNpdf)
        {
            List<string> PolaczoneStrony = new List<string>();
            List<int> StronyDoPliku = new List<int>();
            string NowyKatalog = FFNpdf.Sciezka() + FFNpdf.NazwaPliku(false) + @"\";
            if (Directory.Exists(NowyKatalog) == false)
            {
                Directory.CreateDirectory(NowyKatalog);
            }
            for (int i = 0; i <= StrPDF.Count-1; i++)
            {
                string TrescStrony = OdczytaneStrony[i];
                List<string> Rozbity = TrescStrony.Split('\n').ToList();
                string TMP = Rozbity[Rozbity.Count - 1].Replace("*","").Replace(" ","").ToLower();
                PolaczoneStrony.AddRange(Rozbity);
                StronyDoPliku.Add(i);
                if (TMP.Contains("koniec"))
                {
                    //string SciezkaTMP = @"H:\Zlecenie\Skladany_" + i.ToString("000") ;
                    string SciezkaTMP = NowyKatalog + @"Skladany_" + i.ToString("000");
                    PdfDocument pdfDocument = new PdfDocument(new PdfWriter(SciezkaTMP + ".pdf"));
                    pdfDocSource.CopyPagesTo(StronyDoPliku[0]+1, StronyDoPliku[StronyDoPliku.Count - 1]+1, pdfDocument);
                    pdfDocument.Close();
                    File.WriteAllLines(SciezkaTMP + ".txt", PolaczoneStrony);

                    PolaczoneStrony.Clear();
                    StronyDoPliku.Clear();
                    //utwórz nowy dokument pdf i wkopiuj do niego wszystkie strony z kolekcji StronyDoPliku
                    //utwórz pod taką samą nazwą plik TXT i zapisz do niego treść PolaczoneStrony
                    //wyczyść powyższe kolekcje
                }
                else
                {
                }
            }
        }
        private static void UsuwaniePodpisow(string[] dest)
        {
            foreach(string E in dest)
            {
                string PlikDocelowy = E.Sciezka() + E.NazwaPliku(false) + "_BP.pdf";
                PdfWriter pdfWriter = new PdfWriter(PlikDocelowy);
                PdfDocument pdfDoc = new PdfDocument(new PdfReader(E), pdfWriter);
                Document document = new Document(pdfDoc);
                
                PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, true);

                form.FlattenFields();
                
                document.Add(new Paragraph("KOPIA NIENADZOROWANA")
                    .SetFixedPosition(40, 15, 200));
                document.Close();
                pdfDoc.Close();
            }
        }
    }
    /// <summary>
    /// Do wyciągania txt z pdf
    /// </summary>
    public static class ReaderExtensions
    {
        public static string[] ExtractText(this PdfPage page, params Rectangle[] rects)
        {
            var textEventListener = new LocationTextExtractionStrategy();
            PdfTextExtractor.GetTextFromPage(page, textEventListener);
            string[] result = new string[rects.Length];
            for (int i = 0; i < result.Length; i++)
            {
                result[i] = textEventListener.GetResultantText(rects[i]);
            }
            return result;
        }
        public static string GetResultantText(this LocationTextExtractionStrategy strategy, Rectangle rect)
        {
            IList<TextChunk> locationalResult = (IList<TextChunk>)locationalResultField.GetValue(strategy);
            List<TextChunk> nonMatching = new List<TextChunk>();
            foreach (TextChunk chunk in locationalResult)
            {
                ITextChunkLocation location = chunk.GetLocation();
                Vector start = location.GetStartLocation();
                Vector end = location.GetEndLocation();
                if (!rect.IntersectsLine(start.Get(Vector.I1), start.Get(Vector.I2), end.Get(Vector.I1), end.Get(Vector.I2)))
                {
                    nonMatching.Add(chunk);
                }
            }
            nonMatching.ForEach(c => locationalResult.Remove(c));
            try
            {
                return strategy.GetResultantText();
            }
            finally
            {
                nonMatching.ForEach(c => locationalResult.Add(c));
            }
        }
        static FieldInfo locationalResultField = typeof(LocationTextExtractionStrategy).GetField("locationalResult", BindingFlags.NonPublic | BindingFlags.Instance);
    }
}
