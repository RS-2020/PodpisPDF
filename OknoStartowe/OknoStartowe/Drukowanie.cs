using System.Drawing.Printing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using PdfiumViewer;
using iText.Kernel.Pdf;
using iText.Layout.Element;

namespace OknoStartowe
{
    class Drukowanie
    {
        /// <summary>
        /// plik wejściowy dla drukowania PDF, obsługuje format A3
        /// </summary>
        /// <param name="SciezkaDoPDF">Ścieżka w postaci C:\TMP\plik.pdf</param>
        /// <param name="DodajDate">1 - jeśli login drukującego i data mają być dodane, 0 - jeśli nie</param>
        public static void DrukujPlik(string SciezkaDoPDF, string DodajDate)
        {
            File.AppendAllText(@"C:\TMP\log.txt", Environment.NewLine + $"Start drukowania: {SciezkaDoPDF.NazwaPliku(true)}");
            try
            {
                PrinterSettings printerSettings = new PrinterSettings();
                string NazwaDrukarki = printerSettings.PrinterName;
                string Papier = RozmiarPapieru(SciezkaDoPDF); 
                if (DodajDate == "1")
                {
                    SciezkaDoPDF = DodawanieDaty(SciezkaDoPDF);
                }
                PrintPDF(NazwaDrukarki, Papier, SciezkaDoPDF, 1);
                if (DodajDate == "1" && SciezkaDoPDF.StartsWith(@"C:\TMP\") && File.Exists(SciezkaDoPDF))
                {
                    try { File.Delete(SciezkaDoPDF); } catch { }
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(@"C:\TMP\log.txt", Environment.NewLine + ex.Message);
            }
        }
        private static string DodawanieDaty(string PlikPoprawiony)
        {
            string PlikDocelowy = @"C:\TMP\" + PlikPoprawiony.NazwaPliku(false) + "_D.pdf";
            PdfWriter pdfWriter = new PdfWriter(PlikDocelowy);
            iText.Kernel.Pdf.PdfDocument pdfDoc = new iText.Kernel.Pdf.PdfDocument(new PdfReader(PlikPoprawiony), pdfWriter);
            iText.Layout.Document document = new iText.Layout.Document(pdfDoc);

            iText.Forms.PdfAcroForm form = iText.Forms.PdfAcroForm.GetAcroForm(pdfDoc, true);

            //form.FlattenFields();

            document.Add(new Paragraph($"Data wydruku: {DateTime.Now} przez: {Environment.UserName}")
                .SetFixedPosition(19, 17, 250)
                .SetFontSize(6));
            document.Close();
            pdfDoc.Close();
            return PlikDocelowy;
        }
        private static string RozmiarPapieru(string SciezkaDoPDF)
        {
            PdfReader reader = null;
            iText.Kernel.Pdf.PdfDocument pdfDocument = null;
            try
            {
                reader = new PdfReader(SciezkaDoPDF);
                pdfDocument = new iText.Kernel.Pdf.PdfDocument(reader);
                PdfPage pdfPage = pdfDocument.GetFirstPage();
                iText.Kernel.Geom.Rectangle rectangle = pdfPage.GetPageSize();
                float Wysokosc = rectangle.GetHeight();
                float Szerokosc = rectangle.GetWidth();
                float MAX = Math.Max(Wysokosc, Szerokosc);
                if (MAX == 1191)
                {
                    return "A3";
                }
                else
                {
                    return "A4";
                }
            }
            catch(Exception ex)
            {
                throw new Exception("Błąd przy ustalaniu rozmiaru papieru:\n" + ex.Message);
            }
            finally
            {
                reader.Close();
            }
        }
        public static bool PrintPDF(string printer, string paperName, string filename, int copies)
        {
            try
            {
                // Create the printer settings for our printer
                var printerSettings = new PrinterSettings
                {
                    PrinterName = printer,
                    Copies = (short)copies,
                };

                // Create our page settings for the paper size selected
                var pageSettings = new PageSettings(printerSettings)
                {
                    Margins = new Margins(0, 0, 0, 0),
                };
                foreach (PaperSize paperSize in printerSettings.PaperSizes)
                {
                    if (paperSize.PaperName == paperName)
                    {
                        pageSettings.PaperSize = paperSize;
                        break;
                    }
                }

                // Now print the PDF document
                using (var document = PdfiumViewer.PdfDocument.Load(filename))
                {
                    using (var printDocument = document.CreatePrintDocument())
                    {
                        printDocument.PrinterSettings = printerSettings;
                        printDocument.DefaultPageSettings = pageSettings;
                        printDocument.PrintController = new StandardPrintController();
                        printDocument.Print();
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}