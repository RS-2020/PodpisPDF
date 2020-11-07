using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Pkcs;
using Org.BouncyCastle.X509;
using System.IO;
using iText.Signatures;
using iText.Kernel.Pdf;
using iText.IO.Image;
using iText.Forms.Fields;
using com.itextpdf;
using iText.Forms;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography;
using System.Security.Permissions;
using Org.BouncyCastle.Crypto.Parameters;

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
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            if (Arg.Length < 2)
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
            //MessageBox.Show("Koniec");
            return;
            //Application.Run(new Form1());
        }
        public enum Podpis { Projektowal, Sprawdzil, Zatwierdzil }
        public static void Podpisywanie(string SciezkaDoPDF, string PowodPodpisania, float Wysokosc)
        {
            X509Store store = null;
            try
            {
                PdfReader reader = new PdfReader(SciezkaDoPDF);
                
                store = new X509Store(StoreName.My);
                store.Open(OpenFlags.ReadOnly);
                X509Certificate2 x509Certificate = null;
                foreach (var certificate in store.Certificates)
                {
                    TimeSpan timeSpan = certificate.NotAfter - certificate.NotBefore;
                    if (timeSpan.Days == 365)
                    {
                        x509Certificate = certificate;
                        break;
                    }
                    if (store.Certificates.IndexOf(certificate) == store.Certificates.Count -1 ) { throw new Exception("Nie znaleziono odpowiedniego certyfikatu."); }
                }

                RSA rSA = x509Certificate.GetRSAPrivateKey();
                AsymmetricCipherKeyPair akp = Org.BouncyCastle.Security.DotNetUtilities.GetKeyPair(rSA);

                AsymmetricKeyParameter pk_zProvidera = akp.Private;

                Org.BouncyCastle.X509.X509Certificate x509Przekonwertowany = Org.BouncyCastle.Security.DotNetUtilities.FromX509Certificate(x509Certificate);

                Org.BouncyCastle.X509.X509Certificate[] chain = new Org.BouncyCastle.X509.X509Certificate[] {x509Przekonwertowany };
                
                string DEST = "H:\\TMP.pdf";

                PdfDocument pdfDocument = new PdfDocument(reader);
                float PolozenieX = 0;
                PdfPage pdfPage = pdfDocument.GetFirstPage();
                iText.Kernel.Geom.Rectangle rectangle = pdfPage.GetPageSize();
                if (rectangle.GetRight() == 1191)
                {
                    PolozenieX = 835;
                }
                else
                {
                    PolozenieX = 240;
                }
                pdfDocument.Close();
                reader = new PdfReader(SciezkaDoPDF);
                
                //PdfSigner signer = new PdfSigner(reader, new FileStream(DEST, FileMode.Create), new StampingProperties() );
                PdfSigner signer = new PdfSigner(reader, new FileStream(DEST, FileMode.Create), true);

                //pole 1: .SetPageRect(new iText.Kernel.Geom.Rectangle(240, 142, 120, 30))
                //pole 2: .SetPageRect(new iText.Kernel.Geom.Rectangle(240, 142-15, 120, 30))
                //pole 3: .SetPageRect(new iText.Kernel.Geom.Rectangle(240, 142-30, 120, 30))
                PdfSignatureAppearance appearance = signer.GetSignatureAppearance();
                string Podpisujacy = chain[0].SubjectDN.ToString();
                string[] RozbityWiersz = Podpisujacy.Split(',');
                if (RozbityWiersz.Length == 8)
                {
                    Podpisujacy = RozbityWiersz[6].Replace("CN=","");
                }

                appearance.SetReason(PowodPodpisania)
                    .SetPageNumber(1)
                    .SetPageRect(new iText.Kernel.Geom.Rectangle(PolozenieX, Wysokosc, 200, 30))
                    .SetLocation("Bielsko-Biała");
                signer.SetFieldName(PowodPodpisania);

                appearance.SetLayer2Text($"Data: {signer.GetSignDate()} {PowodPodpisania}: {Podpisujacy} ");
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
                    store.Close();
                }
                catch { }
            }
        }
        public static bool ZmianaNaReadOnly(string SciezkaDoPliku, bool SetReadOnly)
        {
            try
            {
                FileAttributes Atrybuty = System.IO.File.GetAttributes(SciezkaDoPliku);
                if ((Atrybuty & FileAttributes.ReadOnly) != FileAttributes.ReadOnly && SetReadOnly)
                {
                    System.IO.File.SetAttributes(SciezkaDoPliku, System.IO.File.GetAttributes(SciezkaDoPliku) | FileAttributes.ReadOnly);
                    return true;
                }
                else if (SetReadOnly == false)
                {
                    Atrybuty = RemoveAttribute(Atrybuty, FileAttributes.ReadOnly);
                    System.IO.File.SetAttributes(SciezkaDoPliku, Atrybuty);
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
        private static string OdczytPojedynczegoPDF(string SciezkaPDF)
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
    }
}
