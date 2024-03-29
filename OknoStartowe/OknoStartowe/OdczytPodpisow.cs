using System;
using System.Collections.Generic;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Pkcs;
using Org.BouncyCastle.Security;
using iText.Kernel.Pdf;
using iText.Signatures;
using Org.BouncyCastle.X509;
using iText.Forms;

namespace OknoStartowe
{
    public static class OdczytPodpisow
    {
        /// <summary>
        /// Zadaniem funkcji jest odczytanie nazwisk osób, które podpisały PDF.
        /// </summary>
        /// <param name="FFN_PDF">Ścieżka do pliku PDF który chcemy odczytać.</param>
        public static List<Podpis> PodpisanePrzez(string FFN_PDF)
        {
            List<Podpis> LPodpisow = new List<Podpis>();
            using (PdfReader pdfReader = new PdfReader(FFN_PDF))
            using (PdfDocument pdfDocument = new PdfDocument(pdfReader))
            {
                SignatureUtil signatureUtil = new SignatureUtil(pdfDocument);
                var ListaCertyfikatow = signatureUtil.GetSignatureNames();
                int IloscPodpisow = ListaCertyfikatow.Count;
                int i = 1;
                foreach (string signatureName in ListaCertyfikatow)
                {
                    Podpis podpis = new Podpis();
                    podpis.TypPodpisu = signatureName;

                    PdfSignature signature = signatureUtil.GetSignature(signatureName);                  //acroForm.GetSignature(signatureName);
                    PdfString DataSTR = signature.GetDate(); 

                    //Console.WriteLine($"Informacje o podpisie: {signatureName}");

                    PdfPKCS7 pkcs7 = signatureUtil.ReadSignatureData(signatureName);
                    podpis.DataPodpisu = pkcs7.GetSignDate();

                    // Odczytaj certyfikat z podpisu
                    X509Certificate cert = pkcs7.GetSigningCertificate();

                    var TMP = cert.SubjectDN.GetValueList();
                    if (TMP.Count >= 8)
                    {
                        podpis.LoginPodpisujacego = TMP[6].ToString().ToLower(); // UWAGA: tutaj jest problem: login może być na innej pozycji
                    }
                    podpis.PozycjaPodpisu = i;
                    i++;
                    podpis.IloscPodpisowNaDokumencie = IloscPodpisow;
                    LPodpisow.Add(podpis);
                }
            }
            return LPodpisow;
        }
        public class Podpis
        {
            public string TypPodpisu { get; internal set; }
            public string LoginPodpisujacego { get; internal set; }
            public string SciezkaDoPDF { get; internal set; }
            public int IloscPodpisowNaDokumencie { get; internal set; }
            /// <summary>
            /// Numer kolejny podpisu - licząc od jedynki
            /// </summary>
            public int PozycjaPodpisu { get; internal set; }
            public DateTime DataPodpisu { get; internal set; }
            /// <summary>
            /// Zwraca zawartość klasy w postaci:<br>
            /// ŚciezkaDoPliku|CałkowitaIloscPodpisow|NumerKolejnyPodpisuLiczacOdJeden|RodzajPodpisu|LoginPodpisujacego|DataIGodzinaPodpisu</br>
            /// </summary>
            /// <returns></returns>
            public override string ToString()
            {
                //ŚciezkaDoPliku|CałkowitaIloscPodpisow|NumerKolejnyPodpisuLiczacOdJeden|RodzajPodpisu|LoginPodpisujacego|DataIGodzinaPodpisu
                return $"{SciezkaDoPDF}|{IloscPodpisowNaDokumencie}|{PozycjaPodpisu}|{TypPodpisu}|{LoginPodpisujacego}|{DataPodpisu}";
            }
            public static Podpis KonwertujStringNaKlasePodpis(string KlasaWStringu)
            {
                string[] Rozbita = KlasaWStringu.Split('|');
                Podpis podpis = new Podpis();
                podpis.SciezkaDoPDF = Rozbita[0];
                podpis.IloscPodpisowNaDokumencie = int.Parse(Rozbita[1]);
                podpis.PozycjaPodpisu = int.Parse(Rozbita[2]);
                podpis.TypPodpisu = Rozbita[3];
                podpis.LoginPodpisujacego = Rozbita[4];
                podpis.DataPodpisu = DateTime.Parse(Rozbita[5]);
                return podpis;
            }
        }
    }
}
