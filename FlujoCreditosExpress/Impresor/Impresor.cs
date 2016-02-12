using System;
using System.Collections.Generic;
using System.Linq;

using System.IO;
using iTextSharp.text.pdf;
using System.Data;
using System.Text;
using iTextSharp.text.pdf.parser;
using System.util.collections;
using iTextSharp.text;
using System.Net.Mail;

namespace FlujoCreditosExpress
{
    class Impresor
    {
        public string P_OutputStream = "Reporte de Flujo_" + 
            DateTime.Now.Year + DateTime.Now.Month + DateTime.Now.Day + "." +
            DateTime.Now.Hour + DateTime.Now.Minute + ".pdf";

        //Create a brand new PDF from scratch and without a template
        public void CreatePDFNoTemplate()
        {
            Document pdfDoc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(pdfDoc, new FileStream(P_OutputStream, FileMode.OpenOrCreate));
            Image image = null;

            using (var inputImageStream = new FileStream("logo.png", FileMode.Open))
            {
                image = Image.GetInstance(inputImageStream);
                image.ScalePercent(25);
                image.SetAbsolutePosition(10, 790);
            }

            pdfDoc.Open();
            pdfDoc.AddTitle("Reporte de Flujo");
            pdfDoc.AddSubject("Reporte de la generación del flujo de créditos");
            pdfDoc.AddKeywords("Metadata, iTextSharp 5.4.4, Reporte, Flujo");
            pdfDoc.AddCreator("iTextSharp 5.4.4");
            pdfDoc.AddAuthor("Financiera Zafy");
            pdfDoc.AddHeader("Ninguno", "Sin cabecera");

            PdfPTable table = new PdfPTable(3);
            PdfPCell cell = new PdfPCell(new Phrase("Flujo generado"));
            cell.Colspan = 3;
            cell.HorizontalAlignment = 2; //0=Left, 1=Centre, 2=Right
            cell.BorderWidth = 0;
            table.AddCell(cell);
            table.AddCell("Col 1 Row 1");
            table.AddCell("Col 2 Row 1");
            table.AddCell("Col 3 Row 1");
            table.AddCell("Col 1 Row 2");
            table.AddCell("Col 2 Row 2");
            table.AddCell("Col 3 Row 2");
            pdfDoc.Add(table);

            PdfContentByte cb = writer.DirectContent;
            cb.AddImage(image);
            cb.MoveTo(10, 690);
            //cb.LineTo(pdfDoc.PageSize.Width / 2, pdfDoc.PageSize.Height);
            cb.Stroke();

            pdfDoc.Close();
        }
    }
}
