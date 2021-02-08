using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
using iText.Kernel.Pdf;
using System.IO;

namespace OknoStartowe
{
    public class Statystyka
    {
        public string SciezkaPDF { get; set; }
        public string NrArtykulu { get; set; }
        public int IloscArkuszy { get; set; }
        /// <summary>
        /// TRUE - składa się z arkuszy A4, FALSE - składa się z arkuszy A3
        /// </summary>
        public bool FormatA4 { get; set; }
        /// <summary>
        /// Może oznaczać datę wydruku lub ostatniej modyfikacji
        /// </summary>
        public DateTime DataWydruku { get; set; }
        public override string ToString()
        {
            return $"{NrArtykulu};{IloscArkuszy};{FormatA4.ToString()};{DataWydruku}";
        }
        public static void Testowy()
        {
            List<string> Lista = new List<string>
            {
                
            };
            ZbadajPDF(Lista);
        }
        public static void ZPliku()
        {
            //string Dane = @"H:\StatystykaPDF.txt";
            string Dane = @"H:\PDFZdata.csv";
            List<string> Plik = File.ReadAllLines(Dane).ToList();
            List<Statystyka> Wynik = ZbadajPDF(Plik);
            List<string> WynikiSTR = new List<string>();
            foreach (Statystyka _stat in Wynik)
            {
                WynikiSTR.Add(_stat.ToString());
            }
            File.WriteAllLines(@"H:\PDFWydrukowane.txt", WynikiSTR);
        }
        public static List<Statystyka> ZbadajPDF(List<string> ListaPDF)
        {
            List<Statystyka> ListaWynikow = new List<Statystyka>();
            foreach (string SciezkaPDF in ListaPDF)
            {
                try
                {
                    //string[] rozbity = SciezkaPDF.Split(';');
                    if (File.Exists(SciezkaPDF) == false) { continue; }
                    Statystyka _stat = new Statystyka();
                    PdfReader reader = new PdfReader(SciezkaPDF);
                    PdfDocument pdfDocument = new PdfDocument(reader);
                    PdfPage pdfPage = pdfDocument.GetFirstPage();
                    Rectangle rectangle = pdfPage.GetPageSize();
                    _stat.SciezkaPDF = SciezkaPDF;
                    _stat.NrArtykulu = SciezkaPDF.NazwaPliku(false);
                    _stat.IloscArkuszy = pdfDocument.GetNumberOfPages();
                    _stat.DataWydruku = File.GetLastWriteTime(SciezkaPDF);
                    _stat.FormatA4 = rectangle.GetWidth() == 595 || rectangle.GetHeight() == 595;

                    ListaWynikow.Add(_stat);

                    pdfDocument.Close();
                    reader.Close();
                }
                catch { }
            }
            return ListaWynikow;
        }
    }
}
