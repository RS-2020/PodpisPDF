using System;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Signatures;
using Org.BouncyCastle.Crypto;

namespace OknoStartowe
{
    public static class Podpis
    {
        public const string ProtokolZmian = "Zatwierdzenie protokołu";
        public static void Wejscie()
        {
            // Ścieżka do pliku PDF, który chcesz podpisać
            string inputPdfPath = @"C:\tmp\PEKS001750-TEST.pdf";

            // Ścieżka do zapisania pliku podpisanego
            string outputPdfPath = @"C:\tmp\PEKS001750-TEST_Podpisany.pdf";

            Program.ZaczytajUstawienia();
            // Odczytaj certyfikat z magazynu certyfikatów
            X509Certificate2 certificate = PobierzCertyfikat();

            // Podpisz plik PDF
            Podpisuj(inputPdfPath, outputPdfPath, certificate);

            Console.WriteLine("Plik PDF został pomyślnie podpisany.");
        }
        private static void Podpisuj(string PlikZrodlowy, string PlikDocelowy, X509Certificate2 Certyfikat)
        {
            using (PdfReader reader = new PdfReader(PlikZrodlowy))
            {
                using (FileStream StrumienWyjscia = new FileStream(PlikDocelowy, FileMode.Create))
                {
                    
                    StampingProperties SP = new StampingProperties();
                    SP.UseAppendMode();
                    PdfSigner signer = new PdfSigner(reader, StrumienWyjscia, SP);

                    using (PdfReader reader2 = new PdfReader(PlikZrodlowy))
                    {
                        using (PdfDocument pdfDocument = new PdfDocument(reader2))
                        {
                            UstawWlasciwosciPodpisu(signer, Certyfikat, "PowodPodpisu", 123, pdfDocument);

                            RSA rSA = Certyfikat.GetRSAPrivateKey();
                            AsymmetricCipherKeyPair akp = Org.BouncyCastle.Security.DotNetUtilities.GetKeyPair(rSA);

                            AsymmetricKeyParameter pk_zProvidera = akp.Private;

                            Org.BouncyCastle.X509.X509Certificate x509Przekonwertowany = Org.BouncyCastle.Security.DotNetUtilities.FromX509Certificate(Certyfikat);
                            Org.BouncyCastle.X509.X509Certificate[] chain = new Org.BouncyCastle.X509.X509Certificate[] { x509Przekonwertowany };

                            IExternalSignature pks = new PrivateKeySignature(pk_zProvidera, DigestAlgorithms.SHA512);

                            signer.SignDetached(pks, chain, null, null, null, 0, PdfSigner.CryptoStandard.CMS);

                            //PdfSignatureAppearance appearance = signer.GetSignatureAppearance();
                            //iText.Kernel.Pdf.Xobject.PdfFormXObject XObj = appearance.GetLayer0();
                        } 
                    }
                    using (PdfReader reader3 = new PdfReader(PlikDocelowy))
                    {
                        using (PdfDocument Odczytywany = new PdfDocument(new PdfWriter(StrumienWyjscia)))
                        {
                            //using (PdfDocument Zapisywany = new PdfDocument(new PdfWriter(@"C:\tmp\PEKS001750-TEST_Podpisany_AllPages.pdf")))
                            //{
                                for (int i = 2; i <= Odczytywany.GetNumberOfPages(); i++)
                                {
                                    try
                                    {
                                        PdfPage page = Odczytywany.GetPage(i);

                                        PdfCanvas pageCanvas = new PdfCanvas(page.NewContentStreamAfter(), page.GetResources(), Odczytywany);
                                        pageCanvas.AddXObjectAt(signer.GetSignatureAppearance().GetLayer0(), 0, 0);
                                    }
                                    catch { }
                                }
                            //}
                        }
                    }
                }
            }
            
        }
        private static string LoginZCertyfikatu(X509Certificate2 Certyfikat)
        {
            string Temat = Certyfikat.Subject;
            string[] Rozbity = Temat.Split(',');
            foreach(string El in Rozbity)
            {
                if (El.Trim().StartsWith("CN="))
                {
                    string Login = El.Trim().Replace("CN=", "");
                    return Login;
                }
            }
            return "";
        }
        private static void UstawWlasciwosciPodpisu(PdfSigner signer, X509Certificate2 Certyfikat, string PowodPodpisania, float Wysokosc, PdfDocument pdfDocument)
        {
            string Calibri = @"C:\Windows\Fonts\Calibri.ttf";
            PdfFont MyFont = PdfFontFactory.CreateFont(Calibri, "Identity-H", true);
            PdfSignatureAppearance appearance = signer.GetSignatureAppearance();
            float WlkFonta = 8;
            float PolozenieX = 0;

            PdfPage pdfPage = pdfDocument.GetFirstPage(); // przerobić by podpis był na każdej stronie
            Rectangle rectangle = pdfPage.GetPageSize();

            string Podpisujacy = LoginZCertyfikatu(Certyfikat);
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
            appearance.SetReason(PowodPodpisania)
                    .SetPageNumber(1)
                    .SetPageRect(new iText.Kernel.Geom.Rectangle(PolozenieX, Wysokosc, 200 + 40, 30))
                    .SetLocation("Bielsko-Biała");
            signer.SetFieldName(PowodPodpisania);

            appearance.SetLayer2Text($"{PowodPodpisania}: {Podpisujacy} Data: {signer.GetSignDate().ToString("yyyy-MM-dd HH:mm:ss")}");
            appearance.SetLayer2FontSize(WlkFonta);


        }
        public static X509Certificate2 PobierzCertyfikat()
        {
            using (X509Store store = new X509Store(StoreName.My))
            {
                store.Open(OpenFlags.ReadOnly);
                X509Certificate2 x509Certificate = null;
                X509Certificate2Collection ZbiorCert = (X509Certificate2Collection)store.Certificates;
                X509Certificate2Collection Pofiltrowane = ZbiorCert.Find(X509FindType.FindByTemplateName, Program.NazwaSzablonu, true);
                if (Pofiltrowane.Count > 0)
                {
                    x509Certificate = Pofiltrowane[0];
                }
                else
                {
                    throw new Exception("Nie znaleziono odpowiedniego certyfikatu.");
                }
                return x509Certificate;
            }
        }
    }
}
