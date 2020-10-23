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

namespace OknoStartowe
{
    static class Program
    {
        /// <summary>
        /// Główny punkt wejścia dla aplikacji.
        /// Parametry wejściowe:
        /// 1. Sciezka do pliku
        /// 2. hasło do certyfikatu
        /// 3. rodzaj podpisu: konstr/spr/zatw
        /// </summary>
        [STAThread]
        static void Main(string[] Arg)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            if (Arg.Length< 3)
            {
                MessageBox.Show("Zbyt mało parametrów", "Podpisywanie PDF", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            else
            {
                Podpisywanie(Arg[0], Arg[1], Arg[2]);
            }
            Application.Run(new Form1());
        }
        public enum Podpis { Projektowal, Sprawdzil, Zatwierdzil }
        public static void Podpisywanie(string SciezkaDoPDF, string Haslo, string PowodPodpisania)
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

                PdfSignatureAppearance appearance = signer.GetSignatureAppearance();
                appearance.SetReason(PowodPodpisania)
                    .SetPageRect(new iText.Kernel.Geom.Rectangle(36, 648, 200, 100))
                    .SetPageNumber(1);
                signer.SetFieldName("Podpis elektroniczny");
                IExternalSignature pks = new PrivateKeySignature(pk, DigestAlgorithms.SHA256);
                signer.SignDetached(pks, chain, null, null, null, 0, PdfSigner.CryptoStandard.CMS);

                File.Move(SciezkaDoPDF, SciezkaDoPDF + "1");
                File.Move(DEST, SciezkaDoPDF);
                File.Delete(SciezkaDoPDF + "1");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
