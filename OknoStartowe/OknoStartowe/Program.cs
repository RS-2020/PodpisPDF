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

namespace OknoStartowe
{
    static class Program
    {
        /// <summary>
        /// Główny punkt wejścia dla aplikacji.
        /// Parametry wejściowe:
        /// 0. Zapis / odczyt
        /// 1. Sciezka do pliku
        /// 2. hasło do certyfikatu
        /// 3. rodzaj podpisu: konstr/spr/zatw
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
                Podpisywanie(Arg[1], Arg[2], Arg[3], float.Parse(Arg[4]));
            }
            //MessageBox.Show("Koniec");
            return;
            //Application.Run(new Form1());
        }
        public enum Podpis { Projektowal, Sprawdzil, Zatwierdzil }
        public static void Podpisywanie(string SciezkaDoPDF, string Haslo, string PowodPodpisania, float Wysokosc)
        {
            try
            {
                PdfReader reader = new PdfReader(SciezkaDoPDF);
                string Certyfikat = File.ReadAllText(@"H:\LokalizacjaCertyfikatu.txt");
                char[] Pass = Haslo.ToCharArray();
                Pkcs12Store pk12 = new Pkcs12Store(new FileStream(Certyfikat, FileMode.Open, FileAccess.Read), Pass);
                string alias = "";
                foreach (object a in pk12.Aliases)
                {
                    alias = (string)a;
                    if (pk12.IsKeyEntry(alias))
                    {
                        break;
                    }
                }
                ICipherParameters pk = pk12.GetKey(alias).Key;

                X509CertificateEntry[] ce = pk12.GetCertificateChain(alias);
                X509Certificate[] chain = new X509Certificate[ce.Length];
                for (int k = 0; k < ce.Length; ++k)
                {
                    chain[k] = ce[k].Certificate;
                }

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

                PdfSigner signer = new PdfSigner(reader,
                new FileStream(DEST, FileMode.Create),
                new StampingProperties());
                //pole 1: .SetPageRect(new iText.Kernel.Geom.Rectangle(240, 142, 120, 30))
                //pole 2: .SetPageRect(new iText.Kernel.Geom.Rectangle(240, 142-15, 120, 30))
                //pole 3: .SetPageRect(new iText.Kernel.Geom.Rectangle(240, 142-30, 120, 30))
                PdfSignatureAppearance appearance = signer.GetSignatureAppearance();
                
                appearance.SetReason(PowodPodpisania)
                    .SetPageNumber(1)
                    .SetPageRect(new iText.Kernel.Geom.Rectangle(PolozenieX, Wysokosc, 200, 30))
                    .SetLocation("Bielsko-Biała");
                signer.SetFieldName(PowodPodpisania);

                appearance.SetLayer2Text($"Data: {signer.GetSignDate()} {PowodPodpisania}: {signer.GetSignatureDictionary()} ");
                IExternalSignature pks = new PrivateKeySignature(pk, DigestAlgorithms.SHA256);
                signer.SignDetached(pks, chain, null, null, null, 0, PdfSigner.CryptoStandard.CMS);

                reader.Close();
                File.Move(SciezkaDoPDF, SciezkaDoPDF + "1");
                File.Move(DEST, SciezkaDoPDF);
                ZmianaNaReadOnly(SciezkaDoPDF + "1", false);
                File.Delete(SciezkaDoPDF + "1");
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
            //OdczytPojedynczegoPDF(ListaSciezek[1]);
            //string PelnaListaPol = "";

            for (int i = 1; i < ListaSciezek.Length; i++)
            {
                string TMP = OdczytPojedynczegoPDF(ListaSciezek[i]);
                Console.Write(TMP);
                //PelnaListaPol += TMP + System.Environment.NewLine;
            }
            //StreamWriter sw = new StreamWriter(Console.OpenStandardOutput());
            
        }
        private static string OdczytPojedynczegoPDF(string SciezkaPDF)
        {
            PdfReader reader = new PdfReader(SciezkaPDF);
            PdfDocument pdfDocument = new PdfDocument(reader);
            PdfPage pdfPage = pdfDocument.GetPage(1);
            iText.Kernel.Geom.Rectangle rectangle = pdfPage.GetPageSize();
            PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDocument, false);
            IDictionary<string, PdfFormField> PolaNaPDF = form.GetFormFields();
            string DoZwrotu = SciezkaPDF + ";";
            foreach(KeyValuePair<string, PdfFormField> Element in PolaNaPDF)
            {
                if (Element.Value is PdfSignatureFormField)
                {
                    DoZwrotu += Element.Value.GetFieldName().ToString() + ";";
                }
            }
            reader.Close();
            return DoZwrotu;
        }
    }
}
