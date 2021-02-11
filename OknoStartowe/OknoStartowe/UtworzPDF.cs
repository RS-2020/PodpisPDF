using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iText.Forms.Fields;
using iText.IO.Font;
using iText.IO.Font.Constants;
using iText.IO.Image;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Draw;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;

namespace OknoStartowe
{
    public class UtworzPDF
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="PlikDanych"></param>
        /// <param name="OtworzNaKoniec">0 - nie otwiera, 1 - otwiera</param>
        public static void NowyPDF(string PlikDanych, string OtworzNaKoniec)
        {
            string SciezkaZapisu = PlikDanych.Sciezka();
            string NazwaPliku = PlikDanych.NazwaPliku(false);
            PdfWriter writer = new PdfWriter(SciezkaZapisu + NazwaPliku + ".pdf" );
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);
            List<Paragraph> LPar = new List<Paragraph>();
            document.Add(new Paragraph("PRZEWODNIK PRODUKCYJNY")
                .SetTextAlignment(TextAlignment.LEFT)
                .SetFontSize(16));
            //PdfFont MyFont = PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN, "UTF-8" ,true);
            PdfFont bold = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            string TimesNewRoman = @"C:\Windows\Fonts\Times.ttf";
            PdfFont MyFont = PdfFontFactory.CreateFont(TimesNewRoman, "Identity-H", true);

            string KK = @"C:\Windows\Fonts\Code39.ttf";
            //FontProgram fontProgram = FontProgramFactory.CreateFont(KK);
            PdfFont KodKreskowy = PdfFontFactory.CreateFont(KK, true);
            List<string> TrescPliku = File.ReadAllLines(PlikDanych).ToList();
            document.Add(new Paragraph($"Nr zlecenia: {TrescPliku[0].ToUpper()} \t\t\t Nr zgłoszenia: {TrescPliku[1]}")
                .SetTextAlignment(TextAlignment.LEFT)
                .SetFontSize(12)
                .SetFont(MyFont));
            document.Add(new Paragraph(TrescPliku[4].ToUpper())
                .SetFontSize(16)
                .SetFont(bold)
                .SetFixedPosition(395+20, 642 - 25-30, 150));
            document.Add(new Paragraph(TrescPliku[1].ToUpper())
                .SetFontSize(25)
                .SetFont(KodKreskowy)
                .SetFixedPosition(395 + 25, 842 - 70, 150));
            //PdfTextFormField pdfText = PdfTextFormField.CreateText(pdf, new Rectangle(395 - 20, 642 - 20, 150, 20), "name");
            if (File.Exists(TrescPliku[11]))
            {
                ImageData imageData = ImageDataFactory.Create(TrescPliku[11]);
                Image image = new Image(imageData);
                image.ScaleToFit(150, 150);
                image.SetFixedPosition(595-150-50, 842-150-50-30);
                document.Add(new Paragraph()
                    .SetTextAlignment(TextAlignment.RIGHT)
                    .Add(image));
            }
            document.Add(new Paragraph($"Termin rozpoczęcia:\t\t  {TrescPliku[2]}\nTermin zakończenia:\t\t {TrescPliku[3]}\n" +
                $"______________________________________________________________________________")
                .SetFontSize(8)
                .SetFont(MyFont)
                );
            document.Add(new Paragraph("Produkt zlecenia:")
                .SetFont(MyFont)
                .SetFixedLeading(10)
                .SetFontSize(10));
            document.Add(new Paragraph($"{TrescPliku[4]}")
                .SetFontSize(12)
                .SetFixedLeading(10)
                .SetFont(bold));
            document.Add(new Paragraph($"{TrescPliku[5]}\n{TrescPliku[6]}\n{TrescPliku[7]}\n{TrescPliku[8]}")
                .SetFont(MyFont)
                .SetFixedLeading(10)
                .SetFontSize(10));
            document.Add(new Paragraph($"Wielkośc produkcji: {TrescPliku[9]} {TrescPliku[10]}")
                .SetFont(MyFont)
                .SetFixedLeading(10)
                .SetFontSize(10));

            //LPar.Add(new Paragraph($"{TrescPliku[4]}")
            //    .SetFont(bold)
            //    .SetFontSize(14)
            //    .SetTextAlignment(TextAlignment.RIGHT));
            //document.Add(new Paragraph($"_____________________________________________________________________________"));
            document.Add(new LineSeparator(new SolidLine()));
            document.Add(new Paragraph($"PRZEZNACZENIE")
                .SetFont(bold)
                .SetFontSize(12));
            for (int i = 12; i < TrescPliku.Count; i++)
            {
                string[] rozbity = TrescPliku[i].Split('|');
                if (rozbity[0] == "WYK")
                {
                    document.Add(new Paragraph($"{rozbity[1]}\t\t\t{rozbity[2]}\t\t\t{rozbity[3]}")
                        .SetFont(MyFont)
                        .SetFontSize(10));
                }
                else
                {
                    break;
                }
            }

            
            List<Cell> Komorki = new List<Cell>(); //UWAGA!!! Ważna kolejność dodawania

            document.Add(new Paragraph("OPERACJE")
                .SetFont(bold)
                .SetFontSize(12));
            //foreach (Paragraph par in LPar)
            //{
            //    document.Add(par);
            //}
            float szerokoscTabeli = pdf.GetFirstPage().GetPageSize().GetWidth() - 60;
            for (int i = 12; i < TrescPliku.Count; i++)
            {
                string[] rozbity = TrescPliku[i].Split('|');
                if (rozbity[0] == "OPE")
                {
                    float[] columnWidths = { 10, 5 };
                    Table table = new Table(UnitValue.CreatePercentArray(columnWidths));
                    //PdfPage pdfPage = pdf.GetFirstPage();
                    //Rectangle rectangle = pdfPage.GetPageSize();
                    //float Width = rectangle.GetWidth();
                    table.SetWidth(szerokoscTabeli);
                    Paragraph t1 = new Paragraph($"----------------------------------------------------------------------------------------------------\n" +
                        $"{rozbity[1]}\t\t\t{rozbity[2]}\t\t\t{rozbity[3]}\t\t\t{rozbity[4]}\n" +
                        $"----------------------------------------------------------------------------------------------------")
                        .SetFontSize(10)
                        .SetFont(MyFont)
                        .SetTextAlignment(TextAlignment.LEFT);
                    Paragraph t2 = new Paragraph(rozbity[5])
                        .SetFontSize(20)
                        .SetFont(KodKreskowy)
                        .SetTextAlignment(TextAlignment.RIGHT);
                    Cell a = new Cell(1, 1)
                        .Add(t1)
                        .SetBorder(iText.Layout.Borders.Border.NO_BORDER);
                    Cell b = new Cell(1, 1)
                        .Add(t2)
                        .SetBorder(iText.Layout.Borders.Border.NO_BORDER);
                    table.AddCell(a);
                    table.AddCell(b);
                    document.Add(table);
                    //LPar.Add(new Paragraph()
                    //    .Add(t1)
                    //    .Add(t2));
                }
                else if (rozbity[0] == "MAT")
                {
                    float[] columnWidths = { 10, 1 };
                    Table table = new Table(UnitValue.CreatePercentArray(columnWidths));
                    table.SetWidth(szerokoscTabeli);
                    //Table table = new Table(2);
                    Paragraph t1 = new Paragraph($"         \t\t{rozbity[1]}\t\t\t{rozbity[2]}\t\t\t{rozbity[3]}")
                        .SetTextAlignment(TextAlignment.LEFT)
                        .SetFont(MyFont)
                        .SetFixedLeading(8)
                        .SetFontSize(8);
                    Paragraph t2 = new Paragraph($"{rozbity[4]}{rozbity[5]}")
                        .SetFontSize(8)
                        .SetFixedLeading(8)
                        .SetFont(MyFont)
                        .SetTextAlignment(TextAlignment.RIGHT);
                    Cell a = new Cell(1, 1)
                        .Add(t1)
                        .SetBorder(iText.Layout.Borders.Border.NO_BORDER);
                    if (rozbity.Length == 7 && rozbity[6] != "")
                    {
                        Paragraph t1prim = new Paragraph("Uwagi: " + rozbity[6].Replace(";", " "))
                            .SetFontSize(7);
                        a.Add(t1prim);
                    }
                    Cell b = new Cell(1, 1)
                        .Add(t2)
                        .SetBorder(iText.Layout.Borders.Border.NO_BORDER);
                    table.AddCell(a);
                    table.AddCell(b);
                    document.Add(table);
                }
            }

            
            document.Close();
            if (OtworzNaKoniec == "1")
            {
                Process.Start(SciezkaZapisu + NazwaPliku + ".pdf"); 
            }

        }
    }
}
