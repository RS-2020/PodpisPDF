using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OknoStartowe
{
    public static class OdczytPodpisow
    {
        /// <summary>
        /// FUNKCJA NIE DZIAŁA - do zrobienia kiedyś tam...
        /// Zadaniem funkcji jest odczytanie nazwisk osób, które podpisały PDF.
        /// Niestety pola z podpisem są pomijane podczas zbierania tekstu z pdf
        /// </summary>
        /// <param name="FFN_PDF"></param>
        public static void PodpisanePrzez(string FFN_PDF)
        {
            PdfReader reader = new PdfReader(FFN_PDF);
            PdfDocument pdfDocSource = new PdfDocument(reader);
            PdfPage Page = pdfDocSource.GetFirstPage();

            Rectangle Prostokat = Page.GetPageSize();

            string[] Tresc = Page.ExtractText(Prostokat);
            reader.Close();
        }
    }
}
