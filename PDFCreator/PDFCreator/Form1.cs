using Aspose.Pdf;
using Aspose.Pdf.InteractiveFeatures.Annotations;
using Aspose.Pdf.Text;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PDFCreator
{
    public partial class Form1 : Form
    {

        string myDir = @"..\pdfpages\";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            CreatePDFWithTable();

            return;

            // ExStart:ApplyNumberStyle
            // The path to the documents directory.
            string dataDir = @"..\pdfpages\sample.pdf";

            Document pdfDoc = new Document(dataDir);
            pdfDoc.PageInfo.Width = 612.0;
            pdfDoc.PageInfo.Height = 792.0;
            pdfDoc.PageInfo.Margin = new Aspose.Pdf.MarginInfo();
            pdfDoc.PageInfo.Margin.Left = 72;
            pdfDoc.PageInfo.Margin.Right = 72;
            pdfDoc.PageInfo.Margin.Top = 72;
            pdfDoc.PageInfo.Margin.Bottom = 72;

            Aspose.Pdf.Page pdfPage = pdfDoc.Pages.Add();
            pdfPage.PageInfo.Width = 612.0;
            pdfPage.PageInfo.Height = 792.0;
            pdfPage.PageInfo.Margin = new Aspose.Pdf.MarginInfo();
            pdfPage.PageInfo.Margin.Left = 72;
            pdfPage.PageInfo.Margin.Right = 72;
            pdfPage.PageInfo.Margin.Top = 2;
            pdfPage.PageInfo.Margin.Bottom = 72;

            Aspose.Pdf.FloatingBox floatBox = new Aspose.Pdf.FloatingBox();
            floatBox.Margin = pdfPage.PageInfo.Margin;

            pdfPage.Paragraphs.Add(floatBox);

            TextFragment textFragment = new TextFragment();
            TextSegment segment = new TextSegment();

            Aspose.Pdf.Heading heading = new Aspose.Pdf.Heading(1);
            heading.IsInList = true;
            heading.StartNumber = 1;
            heading.Text = "List 1";
            heading.Style = NumberingStyle.NumeralsRomanLowercase;
            heading.IsAutoSequence = true;

            floatBox.Paragraphs.Add(heading);

            Aspose.Pdf.Heading heading2 = new Aspose.Pdf.Heading(1);
            heading2.IsInList = true;
            heading2.StartNumber = 13;
            heading2.Text = "List 2";
            heading2.Style = NumberingStyle.NumeralsRomanLowercase;
            heading2.IsAutoSequence = true;

            floatBox.Paragraphs.Add(heading2);

            Aspose.Pdf.Heading heading3 = new Aspose.Pdf.Heading(2);
            heading3.IsInList = true;
            heading3.StartNumber = 1;
            heading3.Text = "the value, as of the effective date of the plan, of property to be distributed under the plan onaccount of each allowed";
            heading3.Style = NumberingStyle.LettersLowercase;
            heading3.IsAutoSequence = true;

            floatBox.Paragraphs.Add(heading3);
            dataDir = dataDir + "9992.pdf";
            pdfDoc.Save(dataDir);
            // ExEnd:ApplyNumberStyle
            Console.WriteLine("\nNumber style applied successfully in headings.\nFile saved at " + dataDir);
        }

        private void CreatePDFWithTable()
        {
            String inFile = myDir + "sample.pdf";
            Random rand = new Random();
            String outFile = myDir + rand.Next(1000, 10000) + ".pdf";
            Document pdfDoc = new Document(inFile);
            RemoveHeader(pdfDoc);
            RemoveFooter(pdfDoc);


            Table table = new Table();
            table.Border = new BorderInfo(BorderSide.All, 0.1F);
            var margin = new MarginInfo { Top = 0f, Left = 0f, Right = 0f, Bottom = 0f };
            table.Margin = margin;
            double pageWidth = pdfDoc.Pages[1].PageInfo.Width;
            int columnWidth = ((int)pageWidth) / 2;
            table.ColumnWidths = columnWidth + " " + columnWidth;
            Row row1 = new Row();
            var cell1 = new Cell();
            cell1.ColSpan = 2;

            TextFragment text1 = new TextFragment("!!!!!!!! This is a sample footer overwritten");
            text1.Margin.Bottom = 2;
            text1.Margin.Top = 2;
            text1.Margin.Right = 2;
            text1.Margin.Left = 2;
            text1.TextState.Font = FontRepository.FindFont("Verdana");
            text1.TextState.FontSize = 12;
            cell1.Paragraphs.Add(text1);
            row1.Cells.Add(cell1);
            table.Rows.Add(row1);

            Row row2 = new Row();
            Cells cells2 = new Cells();
            //add text
            cells2.Add("Row 2");

            //add logo
            Aspose.Pdf.Image img = new Aspose.Pdf.Image();
            img.File = myDir + "Panda.png";
            // img2.File = myDir + "images.png";
            img.FixWidth = 78;
            img.FixHeight = 31;
            Cell cell2 = row2.Cells.Add();
            cell2.Paragraphs.Add(img);
            cells2.Add(cell2);
            row2.Cells = cells2;
            table.Rows.Add(row2);


            Table tableHeader = new Table();
            tableHeader.Border = new BorderInfo(BorderSide.All, 0.1F);
            var margin1 = new MarginInfo { Top = 0f, Left = 0f, Right = 0f, Bottom = 0f };
            tableHeader.Margin = margin1;
            double pageWidth1 = pdfDoc.Pages[1].PageInfo.Width;
            int columnWidth1 = ((int)pageWidth) / 2;
            tableHeader.ColumnWidths = columnWidth + " " + columnWidth;
            Row row3 = new Row();
            var cell3 = new Cell();
            cell1.ColSpan = 2;

            TextFragment text2 = new TextFragment(" This is a sample Header overwritten....!");
            text2.Margin.Bottom = 2;
            text2.Margin.Top = 2;
            text2.Margin.Right = 2;
            text2.Margin.Left = 2;
            text2.TextState.Font = FontRepository.FindFont("Verdana");
            text2.TextState.FontSize = 12;
            cell3.Paragraphs.Add(text2);
            row3.Cells.Add(cell3);
            tableHeader.Rows.Add(row3);

            Row row4 = new Row();
            Cells cells4 = new Cells();
            //add text
            cells4.Add("Row 1");

            //add logo
            Aspose.Pdf.Image img2 = new Aspose.Pdf.Image();
            //img.File = myDir + "Panda.png";
            img2.File = myDir + "images.png";
            img2.FixWidth = 78;
            img2.FixHeight = 31;
            Cell cells5 = row4.Cells.Add();
            cells5.Paragraphs.Add(img2);
            cells4.Add(cells5);
            row4.Cells = cells4;
            tableHeader.Rows.Add(row4);

            // creating a header
            HeaderFooter header = new HeaderFooter();
            var marginheader = new MarginInfo { Top = 0f, Left = 10f, Right = 10f, Bottom = 0f };
            header.Margin = marginheader;
            header.Paragraphs.Add(tableHeader);

            // creating a footer
            HeaderFooter footer = new HeaderFooter();
            var marginFooter = new MarginInfo { Top = 0f, Left = 10f, Right = 10f, Bottom = 0f };
            footer.Margin = marginFooter;
            footer.Paragraphs.Add(table);

            for (int i = 1; i <= pdfDoc.Pages.Count; i++)
            {
                pdfDoc.Pages[i].Header = header;
                pdfDoc.Pages[i].Footer = footer;
                pdfDoc.ProcessParagraphs();
            }

            pdfDoc.Save(outFile);
        }

        private void RemoveHeader(Aspose.Pdf.Document pdfDoc)
        {
            try
            {
                for (int i = 1; i <= pdfDoc.Pages.Count; i++)
                {
                    Page page = pdfDoc.Pages[i];
                    Aspose.Pdf.Rectangle rect = new Aspose.Pdf.Rectangle(0, page.Rect.Height * 0.95, page.Rect.Width, page.Rect.Height);
                    RedactionAnnotation annot = new RedactionAnnotation(page, rect);
                    annot.FillColor = Aspose.Pdf.Color.White;
                    annot.BorderColor = Aspose.Pdf.Color.Yellow;
                    annot.Color = Aspose.Pdf.Color.Blue;

                    annot.TextAlignment = Aspose.Pdf.HorizontalAlignment.Center;
                    page.Annotations.Add(annot);
                    annot.Redact();

                    TextAbsorber textAbsorber = new TextAbsorber();
                    pdfDoc.Pages[i].Accept(textAbsorber);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void RemoveFooter(Aspose.Pdf.Document pdfDoc)
        {
            try
            {
                for (int i = 1; i <= pdfDoc.Pages.Count; i++)
                {
                    Page page = pdfDoc.Pages[i];
                    Aspose.Pdf.Rectangle rect = new Aspose.Pdf.Rectangle(0, 75, page.Rect.Width, 1);
                    RedactionAnnotation annot = new RedactionAnnotation(page, rect);
                    annot.FillColor = Aspose.Pdf.Color.White;
                    annot.BorderColor = Aspose.Pdf.Color.Yellow;
                    annot.Color = Aspose.Pdf.Color.White;

                    annot.TextAlignment = Aspose.Pdf.HorizontalAlignment.Center;
                    page.Annotations.Add(annot);
                    annot.Redact();

                    TextAbsorber textAbsorber = new TextAbsorber();
                    pdfDoc.Pages[i].Accept(textAbsorber);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
    }
}
