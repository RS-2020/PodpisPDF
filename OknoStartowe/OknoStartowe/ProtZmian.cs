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
    class ProtZmian
    {
        public static void UtworzPDF(string SciezkaZrodlowa, string SciezkaDocelowa, int OtworzNaKoniec)
        {
            List<string> Plik = File.ReadAllLines(SciezkaZrodlowa).ToList();
            
            PdfWriter writer = new PdfWriter(SciezkaDocelowa);
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);
            string Calibri = @"C:\Windows\Fonts\Calibri.ttf";
            //string Arial = @"C:\Windows\Fonts\Arial.ttf";
            PdfFont MyFont = PdfFontFactory.CreateFont(Calibri, "Identity-H", true);
            //PdfFont MyFont2 = PdfFontFactory.CreateFont(Arial, "UTF-8", true);

            float szerokoscTabeli = 595-60; //pdf.GetFirstPage().GetPageSize().GetWidth() - 60;
            float[] columnWidths = { 40, 34, 60, 47 };

            ///------------------------------------------------------------------------------------
            // ------------------------------Tabela górna - nagłówek-------------------------------
            ///------------------------------------------------------------------------------------

            Table tabelaGorna = new Table(UnitValue.CreatePercentArray(columnWidths));
            tabelaGorna.SetWidth(szerokoscTabeli);


            string Image_FFN = Program.Logo;
            ImageData data = ImageDataFactory.Create(Image_FFN);
            Image img = new Image(data);
            img.ScaleToFit(100, 100);
            //img.SetFixedPosition(50, 720);

            //Paragraph p1 = new Paragraph("AAAA");
            Cell kom1 = new Cell(3,1).Add(img);
            kom1.SetVerticalAlignment(VerticalAlignment.MIDDLE);
            kom1.SetHorizontalAlignment(HorizontalAlignment.CENTER);
            kom1.SetHeight(90);
            
            //Paragraph p2 = new Paragraph("BBBB");
            //Cell kom2 = new Cell().Add(p2);
            
            Paragraph p3 = new Paragraph("Protokół zmian").SetFont(MyFont).SetFontSize(16).SetBold().SetTextAlignment(TextAlignment.CENTER);
            Cell kom3 = new Cell(1,2).Add(p3).SetVerticalAlignment(VerticalAlignment.MIDDLE).SetHorizontalAlignment(HorizontalAlignment.CENTER);
            kom3.SetHeight(45);
            kom3.SetHorizontalAlignment(HorizontalAlignment.CENTER);
            kom3.SetVerticalAlignment(VerticalAlignment.MIDDLE);

            Paragraph p4 = new Paragraph("F04-PRQ-03").SetFont(MyFont).SetFontSize(16).SetBold().SetTextAlignment(TextAlignment.CENTER);
            Cell kom4 = new Cell().Add(p4);
            kom4.SetHorizontalAlignment(HorizontalAlignment.CENTER).SetVerticalAlignment(VerticalAlignment.MIDDLE);

            Paragraph p5 = new Paragraph("Nr artykułu:").SetTextAlignment(TextAlignment.CENTER).SetFont(MyFont);
            Cell kom5 = new Cell().Add(p5);
            kom5.SetHorizontalAlignment(HorizontalAlignment.CENTER).SetVerticalAlignment(VerticalAlignment.MIDDLE);

            Paragraph p6 = new Paragraph(Plik[0].ToUpper()).SetTextAlignment(TextAlignment.CENTER).SetFont(MyFont).SetFontSize(16).SetBold(); // <-------- Dana1 NrArt
            Cell kom6 = new Cell().Add(p6);
            kom6.SetHorizontalAlignment(HorizontalAlignment.CENTER).SetVerticalAlignment(VerticalAlignment.MIDDLE);

            Paragraph p7 = new Paragraph("Edycja 3\nData wydania:\n2023-01-03").SetTextAlignment(TextAlignment.CENTER).SetFont(MyFont);
            Cell kom7 = new Cell(2,1).Add(p7);
            kom7.SetHorizontalAlignment(HorizontalAlignment.CENTER).SetVerticalAlignment(VerticalAlignment.MIDDLE);

            Paragraph p8 = new Paragraph("Wersja po zmianie:").SetTextAlignment(TextAlignment.CENTER).SetFont(MyFont);
            Cell kom8 = new Cell().Add(p8);
            kom8.SetHorizontalAlignment(HorizontalAlignment.CENTER).SetVerticalAlignment(VerticalAlignment.MIDDLE);

            Paragraph p9 = new Paragraph(Plik[1]).SetTextAlignment(TextAlignment.CENTER).SetFont(MyFont).SetFontSize(16).SetBold(); // <----------- Dana2 Wersja
            Cell kom9 = new Cell().Add(p9);
            kom9.SetHorizontalAlignment(HorizontalAlignment.CENTER).SetVerticalAlignment(VerticalAlignment.MIDDLE);

            tabelaGorna.AddCell(kom1);

            tabelaGorna.AddCell(kom3);
            tabelaGorna.AddCell(kom4);
            tabelaGorna.AddCell(kom5);
            tabelaGorna.AddCell(kom6);
            tabelaGorna.AddCell(kom7);
            tabelaGorna.AddCell(kom8);
            tabelaGorna.AddCell(kom9);

            document.Add(tabelaGorna);

            document.Add(new Paragraph("\n")); // dystans między tabelami

            ///------------------------------------------------------------------------------------
            // -----------------------------Tabela pod nagłówkiem----------------------------------
            ///------------------------------------------------------------------------------------

            Table tabela2 = new Table(UnitValue.CreatePercentArray(columnWidths));
            tabela2.SetWidth(szerokoscTabeli);

            Paragraph p10 = new Paragraph("Zmiana zgloszona przez:").SetFont(MyFont).SetTextAlignment(TextAlignment.CENTER).SetFontSize(10);//.SetFont(MyFont).SetFontSize(16).SetBold();
            Cell kom10 = new Cell().Add(p10);
            kom10.SetHorizontalAlignment(HorizontalAlignment.CENTER).SetVerticalAlignment(VerticalAlignment.MIDDLE);

            Paragraph p11 = new Paragraph("Stanowisko:").SetTextAlignment(TextAlignment.LEFT).SetFontSize(8).SetFont(MyFont);
            Paragraph p13 = new Paragraph(Plik[2]).SetTextAlignment(TextAlignment.CENTER).SetFont(MyFont); // <-------------- Dana3 Stanowisko
            Cell kom11 = new Cell().Add(p11);
            kom11.Add(p13);
            kom10.SetHorizontalAlignment(HorizontalAlignment.CENTER).SetVerticalAlignment(VerticalAlignment.MIDDLE);

            Paragraph p14 = new Paragraph("Imię i nazwisko:").SetTextAlignment(TextAlignment.LEFT).SetFontSize(8).SetFont(MyFont);
            Paragraph p12 = new Paragraph(Plik[3]).SetTextAlignment(TextAlignment.CENTER).SetFont(MyFont); // <------- Dana4 Imię i nazwisko
            
            Cell kom12 = new Cell().Add(p14).Add(p12);
            kom10.SetHorizontalAlignment(HorizontalAlignment.CENTER).SetVerticalAlignment(VerticalAlignment.MIDDLE);

            Paragraph p15 = new Paragraph("Data zgłoszenia zmiany:").SetTextAlignment(TextAlignment.LEFT).SetFontSize(8).SetFont(MyFont);
            Paragraph p16 = new Paragraph(Plik[4]).SetTextAlignment(TextAlignment.CENTER);//<------------ Dana5 Data zgłoszenia

            Cell kom13 = new Cell().Add(p15).Add(p16);
            kom13.SetHorizontalAlignment(HorizontalAlignment.CENTER).SetVerticalAlignment(VerticalAlignment.MIDDLE);

            Paragraph p17 = new Paragraph("Zmiana zatwierdzona \nprzez DP:").SetTextAlignment(TextAlignment.CENTER).SetFontSize(10).SetFont(MyFont);//.SetFont(MyFont).SetFontSize(16).SetBold();
            Cell kom14 = new Cell().Add(p17);
            kom13.SetHorizontalAlignment(HorizontalAlignment.CENTER).SetVerticalAlignment(VerticalAlignment.MIDDLE);

            Cell kom15 = new Cell(1, 2);
            Cell kom16 = new Cell();

            tabela2.AddCell(kom10);
            tabela2.AddCell(kom11);
            tabela2.AddCell(kom12);
            tabela2.AddCell(kom13);
            tabela2.AddCell(kom14);
            tabela2.AddCell(kom15);
            tabela2.AddCell(kom16);

            document.Add(tabela2);

            document.Add(new Paragraph("\n"));

            ///------------------------------------------------------------------------------------
            // ------------------------------Tabela opis zmiany------------------------------------
            ///------------------------------------------------------------------------------------


            Table OpisZmiany = new Table(UnitValue.CreatePercentArray(new float[] { 1}));
            OpisZmiany.SetWidth(szerokoscTabeli);

            Paragraph p18 = new Paragraph("Opis zmiany").SetTextAlignment(TextAlignment.CENTER);
            Cell kom17 = new Cell().Add(p18);

            Paragraph p19 = new Paragraph(Plik[5].Replace("<CCC>","").Replace("<EEE>",Environment.NewLine)).SetFontSize(10).SetFont(MyFont); // <-------------------- Dana6 Opis zmiany
            Cell kom18 = new Cell().Add(p19);

            OpisZmiany.AddCell(kom17);
            OpisZmiany.AddCell(kom18);

            document.Add(OpisZmiany);

            document.Add(new Paragraph("\n"));

            ///------------------------------------------------------------------------------------
            // ----------------------------Tabela zleceń produkcyjnych-----------------------------
            ///------------------------------------------------------------------------------------


            float[] Szer2 = {23, 23, 16, 34, 85 }; //{ 23, 16, 34, 108};
            Table ZlecProd = new Table(UnitValue.CreatePercentArray(Szer2));
            ZlecProd.SetWidth(szerokoscTabeli);

            Paragraph P4_1 = new Paragraph($"Zlecenia produkcyjne. Stan na {DateTime.Now.ToShortDateString()} {DateTime.Now.ToShortTimeString()}").SetTextAlignment(TextAlignment.CENTER).SetFontSize(10).SetFont(MyFont).SetBold();
            Cell Kom4_1 = new Cell(1, 5).Add(P4_1).SetHorizontalAlignment(HorizontalAlignment.CENTER).SetVerticalAlignment(VerticalAlignment.MIDDLE);
            ZlecProd.AddCell(Kom4_1);

            Paragraph P4_2 = new Paragraph("Nr zlecenia").SetTextAlignment(TextAlignment.CENTER).SetFontSize(10).SetFont(MyFont).SetBold();
            Cell Kom4_2 = new Cell().Add(P4_2).SetHorizontalAlignment(HorizontalAlignment.CENTER).SetVerticalAlignment(VerticalAlignment.MIDDLE);
            ZlecProd.AddCell(Kom4_2);

            Paragraph P4_2_1 = new Paragraph("Nr artykułu").SetTextAlignment(TextAlignment.CENTER).SetFontSize(10).SetFont(MyFont).SetBold();
            Cell Kom4_2_1 = new Cell().Add(P4_2_1).SetHorizontalAlignment(HorizontalAlignment.CENTER).SetVerticalAlignment(VerticalAlignment.MIDDLE);
            ZlecProd.AddCell(Kom4_2_1);

            Paragraph P4_3 = new Paragraph("Nr\nzgłoszenia").SetTextAlignment(TextAlignment.CENTER).SetFontSize(10).SetFont(MyFont).SetBold();
            Cell Kom4_3 = new Cell().Add(P4_3).SetHorizontalAlignment(HorizontalAlignment.CENTER).SetVerticalAlignment(VerticalAlignment.MIDDLE);
            ZlecProd.AddCell(Kom4_3);

            Paragraph P4_4 = new Paragraph("Status").SetTextAlignment(TextAlignment.CENTER).SetFontSize(10).SetFont(MyFont).SetBold();
            Cell Kom4_4 = new Cell().Add(P4_4).SetHorizontalAlignment(HorizontalAlignment.CENTER).SetVerticalAlignment(VerticalAlignment.MIDDLE);
            ZlecProd.AddCell(Kom4_4);

            Paragraph P4_5 = new Paragraph("Sposób naprawy").SetTextAlignment(TextAlignment.CENTER).SetFontSize(10).SetFont(MyFont).SetBold();
            Cell Kom4_5 = new Cell().Add(P4_5).SetHorizontalAlignment(HorizontalAlignment.CENTER).SetVerticalAlignment(VerticalAlignment.MIDDLE);
            ZlecProd.AddCell(Kom4_5);

            for (int i = 6; i < Plik.Count; i++)
            {
                string[] Rozbity = Plik[i].Split(';');

                Paragraph PX_2 = new Paragraph(Rozbity[0]).SetTextAlignment(TextAlignment.CENTER).SetFontSize(10).SetFont(MyFont);
                Cell KomX_2 = new Cell().Add(PX_2);
                ZlecProd.AddCell(KomX_2);

                Paragraph PX_2_1 = new Paragraph(Rozbity[1]).SetTextAlignment(TextAlignment.CENTER).SetFontSize(10).SetFont(MyFont);
                Cell KomX_2_1 = new Cell().Add(PX_2_1);
                ZlecProd.AddCell(KomX_2_1);

                Paragraph PX_3 = new Paragraph(Rozbity[2]).SetTextAlignment(TextAlignment.CENTER).SetFontSize(10).SetFont(MyFont);
                Cell KomX_3 = new Cell().Add(PX_3);
                ZlecProd.AddCell(KomX_3);

                Paragraph PX_4 = new Paragraph(Rozbity[3]).SetTextAlignment(TextAlignment.CENTER).SetFontSize(10).SetFont(MyFont);
                Cell KomX_4 = new Cell().Add(PX_4);
                ZlecProd.AddCell(KomX_4);

                Paragraph PX_5 = new Paragraph(Rozbity[4]).SetTextAlignment(TextAlignment.CENTER).SetFontSize(10).SetFont(MyFont);
                Cell KomX_5 = new Cell().Add(PX_5);
                ZlecProd.AddCell(KomX_5);
            }

            document.Add(ZlecProd);

            document.Close();

            if (File.Exists(SciezkaDocelowa) && (DateTime.Now - File.GetLastWriteTime(SciezkaDocelowa)).TotalSeconds < 60)
            {
                File.Delete(SciezkaZrodlowa);
            }
            if (OtworzNaKoniec == 1)
            {
                Process.Start(SciezkaDocelowa);
            }
        }
    }
}

