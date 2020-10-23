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
            MessageBox.Show("Koniec");
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

                PdfSigner signer = new PdfSigner(reader,
                new FileStream(DEST, FileMode.Create),
                new StampingProperties());
                //pole 1: .SetPageRect(new iText.Kernel.Geom.Rectangle(240, 142, 120, 30))
                //pole 2: .SetPageRect(new iText.Kernel.Geom.Rectangle(240, 142-15, 120, 30))
                //pole 3: .SetPageRect(new iText.Kernel.Geom.Rectangle(240, 142-30, 120, 30))
                PdfSignatureAppearance appearance = signer.GetSignatureAppearance();
                appearance.SetReason(PowodPodpisania)
                    .SetPageNumber(1)
                    .SetPageRect(new iText.Kernel.Geom.Rectangle(240, Wysokosc, 120, 30))
                    .SetLocation("Bielsko-Biała");
                signer.SetFieldName(PowodPodpisania);

                appearance.SetLayer2Text($"Data: {signer.GetSignDate()} {PowodPodpisania}: {signer.GetSignatureDictionary()} ");
                IExternalSignature pks = new PrivateKeySignature(pk, DigestAlgorithms.SHA256);
                signer.SignDetached(pks, chain, null, null, null, 0, PdfSigner.CryptoStandard.CMS);

                reader.Close();
                File.Move(SciezkaDoPDF, SciezkaDoPDF + "1");
                File.Move(DEST, SciezkaDoPDF);
                File.Delete(SciezkaDoPDF + "1");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public static void Odczyt(string[] ListaSciezek)
        {
            for (int i = 0; i < ListaSciezek.Length; i++)
            {
                PdfReader reader = new PdfReader(ListaSciezek[i]);
                

                reader.Close();
            }
        }
    }
}
