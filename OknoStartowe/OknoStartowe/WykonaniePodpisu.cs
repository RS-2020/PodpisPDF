using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Signatures;
using Org.BouncyCastle.Crypto;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OknoStartowe
{
    public static class WykonaniePodpisu
    {
        public static void Podpisywanie(string SciezkaDoPDF, string PowodPodpisania, float Wysokosc)
        {
            string ProtokolZmian = "Zatwierdzenie protokołu";
            X509Store store = null;
            RSA rSA = null;
            try
            {
                if (PowodPodpisania != ProtokolZmian && ZweryfikujCzySaPoprzedniePodpisy(SciezkaDoPDF, PowodPodpisania) == false)
                {
                    MessageBox.Show($"Brak wcześniejszych podpisów dla {SciezkaDoPDF}", "Podpis PDF (zew)", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

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

                pdfDocument.Close();
                reader = new PdfReader(SciezkaDoPDF);
                string Calibri = @"C:\Windows\Fonts\Calibri.ttf";
                PdfFont MyFont = PdfFontFactory.CreateFont(Calibri, "Identity-H", true);
                PdfSigner signer = new PdfSigner(reader, new FileStream(DEST, FileMode.Create), true);
                //signer = new PdfSigner(reader, new FileStream(DEST, FileMode.Create),);
                //pole 1: .SetPageRect(new iText.Kernel.Geom.Rectangle(240, 142, 120, 30))
                //pole 2: .SetPageRect(new iText.Kernel.Geom.Rectangle(240, 142-15, 120, 30))
                //pole 3: .SetPageRect(new iText.Kernel.Geom.Rectangle(240, 142-30, 120, 30))
                PdfSignatureAppearance appearance = signer.GetSignatureAppearance();
                float WlkFonta = 8;
                if (PowodPodpisania == ProtokolZmian)
                {
                    PolozenieX = 185;
                    WlkFonta = 10;
                    appearance.SetLayer2Font(MyFont);
                }
                else if (rectangle.GetRight() == 1191)
                {
                    PolozenieX = 835 - 40;
                }
                else
                {
                    PolozenieX = 240 - 40;
                }
                string Podpisujacy = chain[0].SubjectDN.ToString();
                string[] RozbityWiersz = Podpisujacy.Split(',');
                string UserName = Environment.UserName.ToLower();
                foreach (string El in RozbityWiersz)
                {
                    if (El.ToLower().Contains(UserName))
                    {
                        Podpisujacy = El.Replace("CN=", "");
                        break;
                    }
                }

                appearance.SetReason(PowodPodpisania)
                    .SetPageNumber(1)
                    .SetPageRect(new iText.Kernel.Geom.Rectangle(PolozenieX, Wysokosc, 200 + 40, 30))
                    .SetLocation("Bielsko-Biała");
                signer.SetFieldName(PowodPodpisania);

                appearance.SetLayer2Text($"{PowodPodpisania}: {Podpisujacy} Data: {signer.GetSignDate()}");
                appearance.SetLayer2FontSize(WlkFonta);

                IExternalSignature pks = new PrivateKeySignature(pk_zProvidera, DigestAlgorithms.SHA512);

                signer.SignDetached(pks, chain, null, null, null, 0, PdfSigner.CryptoStandard.CMS);
                reader.Close();
                File.Move(SciezkaDoPDF, SciezkaDoPDF + "1");
                File.Move(DEST, SciezkaDoPDF);
                Program.ZmianaNaReadOnly(SciezkaDoPDF + "1", false);
                File.Delete(SciezkaDoPDF + "1");

                string ZebranePodpisy = OdczytPojedynczegoPDF(SciezkaDoPDF);
                //ZapisLogu(SciezkaDoPDF.ToUpper(), ZebranePodpisy);
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
    }
}
