using System;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Windows.Forms;
using Microsoft.VisualBasic.FileIO;
using System.Linq;

namespace REPLACE_PDF_CONTENT
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string csvFilePath = @"C:\Users\AL5441\Downloads\MOHSIN\Kotak1.csv";
            string logoPath = @"C:\Users\AL5441\Downloads\MOHSIN\KotakLogo.png";
            string pdfOutputPath = @"C:\Users\AL5441\Downloads\MOHSIN\Statement1.pdf";
            string fontPath = @"C:\Users\AL5441\Downloads\MOHSIN\Century Gothic\centurygothic.ttf";
            string fontPath1 = @"C:\Users\AL5441\Downloads\MOHSIN\Century Gothic\centurygothic_bold.ttf";
            string fontPathUni = @"C:\Users\AL5441\Downloads\MOHSIN\Century Gothic\arial-unicode-ms.ttf";

            BaseFont customBaseFont = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            BaseFont customBaseFontBold = BaseFont.CreateFont(fontPath1, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            BaseFont customBaseFontArial = BaseFont.CreateFont(fontPathUni, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

            Font customFont_20BB = new Font(customBaseFontBold, 24);
            Font customFont_9B = new Font(customBaseFontBold, 8);
            Font customFont_9 = new Font(customBaseFont, 9);
            Font customFont_8 = new Font(customBaseFont, 7.5F);
            Font customFont_8_Gray = new Font(customBaseFont, 8F, 1, BaseColor.GRAY);
            Font customFont_9Rupee = new Font(customBaseFontArial, 9);

            Document document = new Document(PageSize.A4, 50, 50, 50, 50);
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(pdfOutputPath, FileMode.Create));
            document.Open();

            PdfPTable headerTable = new PdfPTable(2);
            headerTable.WidthPercentage = 100;
            headerTable.SetWidths(new float[] { 50f, 50f });

            try
            {
                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(logoPath);
                logo.ScaleToFit(120f, 120f);
                PdfPCell logoCell = new PdfPCell(logo);
                logoCell.Border = Rectangle.NO_BORDER;
                headerTable.AddCell(logoCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading logo: {ex.Message}");
                PdfPCell errorCell = new PdfPCell(new Phrase("Logo not available", customFont_9));
                errorCell.Border = Rectangle.NO_BORDER;
                headerTable.AddCell(errorCell);
            }

            PdfPCell accountInfoCell = new PdfPCell();
            accountInfoCell.Border = Rectangle.NO_BORDER;
            accountInfoCell.AddElement(new Paragraph("", customFont_9));
            accountInfoCell.AddElement(new Paragraph("", customFont_9));
            headerTable.AddCell(accountInfoCell);

            document.Add(headerTable);

            PdfPTable accountDetailsTable = new PdfPTable(2);
            accountDetailsTable.WidthPercentage = 100;
            accountDetailsTable.SetWidths(new float[] { 60f, 40f });

            PdfPCell leftDetailsCell1 = new PdfPCell();
            leftDetailsCell1.Border = Rectangle.NO_BORDER;
            leftDetailsCell1.AddElement(new Paragraph("Account Statement", customFont_20BB));
            leftDetailsCell1.AddElement(new Paragraph("01 Jan 2022 - 31 Dec 2022", customFont_8_Gray));

            PdfPCell rightDetailsCell1 = new PdfPCell();
            rightDetailsCell1.Border = Rectangle.NO_BORDER;
            rightDetailsCell1.VerticalAlignment = Element.ALIGN_BOTTOM;
            rightDetailsCell1.AddElement(new Paragraph("", customFont_9));
            rightDetailsCell1.AddElement(new Paragraph("", customFont_9));
            rightDetailsCell1.AddElement(new Paragraph("Account # 6011680927 SAVINGS", customFont_9));
            rightDetailsCell1.AddElement(new Paragraph("Branch Mumbai - Andheri (East)", customFont_9));

            accountDetailsTable.AddCell(leftDetailsCell1);
            accountDetailsTable.AddCell(rightDetailsCell1);

            PdfPCell leftDetailsCell = new PdfPCell();
            leftDetailsCell.Border = Rectangle.NO_BORDER;
            leftDetailsCell.AddElement(new Paragraph("Ali Ahmed Khan", customFont_9));
            leftDetailsCell.AddElement(new Paragraph("CRN XXXXX520", customFont_9));
            leftDetailsCell.AddElement(new Paragraph("203 7B TWIN HEIGHTS RUPA CST", customFont_9));
            leftDetailsCell.AddElement(new Paragraph("ROAD KURLA WEST", customFont_9));
            leftDetailsCell.AddElement(new Paragraph("Mumbai - 400070", customFont_9));

            PdfPCell rightDetailsCell = new PdfPCell();
            rightDetailsCell.Border = Rectangle.NO_BORDER;
            rightDetailsCell.AddElement(new Paragraph("Nominee registered   Khalil Ahmed Khan", customFont_8_Gray));
            rightDetailsCell.AddElement(new Paragraph("IFSC  KKBK0000651", customFont_8_Gray));
            rightDetailsCell.AddElement(new Paragraph("MICR  400485003", customFont_8_Gray));

            accountDetailsTable.AddCell(leftDetailsCell);
            accountDetailsTable.AddCell(rightDetailsCell);

            document.Add(accountDetailsTable);

            document.Add(new Paragraph("\n"));

            PdfPTable table = new PdfPTable(7);
            table.WidthPercentage = 100;
            table.SetWidths(new float[] { 5f, 15f, 15f, 35f, 15f, 10f, 15f });

            table.AddCell(new PdfPCell(new Phrase("#", customFont_9B)) { BackgroundColor = BaseColor.LIGHT_GRAY });
            table.AddCell(new PdfPCell(new Phrase("Transaction Date", customFont_9B)) { BackgroundColor = BaseColor.LIGHT_GRAY });
            table.AddCell(new PdfPCell(new Phrase("Value Date", customFont_9B)) { BackgroundColor = BaseColor.LIGHT_GRAY });
            table.AddCell(new PdfPCell(new Phrase("Transaction Details", customFont_9B)) { BackgroundColor = BaseColor.LIGHT_GRAY });
            table.AddCell(new PdfPCell(new Phrase("Chq / Ref No.", customFont_9B)) { BackgroundColor = BaseColor.LIGHT_GRAY });
            table.AddCell(new PdfPCell(new Phrase("Debit/Credit", customFont_9Rupee)) { BackgroundColor = BaseColor.LIGHT_GRAY });
            table.AddCell(new PdfPCell(new Phrase("Balance", customFont_9B)) { BackgroundColor = BaseColor.LIGHT_GRAY });

            Random random = new Random();
            decimal previousBalance = 0m; // Initialize with zero or any appropriate value

            using (TextFieldParser parser = new TextFieldParser(csvFilePath))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");

                bool isFirstRow = true;
                while (!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields();

                    if (isFirstRow)
                    {
                        isFirstRow = false;
                        // Assuming the first row is header, skip it
                        try
                        {
                            previousBalance = decimal.Parse(fields[7].Replace(",", "")); // Initialize with the first balance
                        }
                        catch (Exception)
                        {
                        }
                        continue;
                    }

                    // Convert current balance to decimal
                    decimal currentBalance = decimal.Parse(fields[7].Replace(",", ""));

                    // Determine the sign for the debit/credit field based on the comparison
                    string sign = currentBalance < previousBalance ? "-" : "+";
                    previousBalance = currentBalance; // Update previous balance

                    for (int i = 0; i < fields.Length; i++)
                    {
                        if (i == 6 || i == 8)
                            continue;

                        if (fields[5] == "8,285.00")
                        {
                            if (i == 3) // Assuming index 3 is where the transaction details should be updated
                            {
                                table.AddCell(new PdfPCell(new Phrase("RUPEEK GOLD Private Limited", customFont_8)));
                            }
                            else if (i == 4) // Assuming index 4 is where the reference number should be
                            {
                                string randomRef = GenerateRandomReference(random);
                                table.AddCell(new PdfPCell(new Phrase(randomRef, customFont_8)));
                            }
                            else if (i == 5) // Assuming index 5 is where the debit/credit amount should be
                            {
                                var sign1 = fields[6] == "CR" ? "" : "-";
                                table.AddCell(new PdfPCell(new Phrase(sign1 + fields[i], customFont_8)));
                            }
                            else
                            {
                                table.AddCell(new PdfPCell(new Phrase(fields[i], customFont_8)));
                            }
                        }
                        else
                        {
                            if (i == 5) // Assuming index 5 is where the debit/credit amount should be
                            {
                                var sign1 = fields[6] == "CR" ? "" : "-";
                                table.AddCell(new PdfPCell(new Phrase(sign1 + fields[i], customFont_8)));
                            }
                            else
                                table.AddCell(new PdfPCell(new Phrase(fields[i], customFont_8)));
                        }
                    }
                }
            }

            document.Add(table);
            document.Close();

            MessageBox.Show("PDF generated successfully!");
        }

        private string GenerateRandomReference(Random random)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            var randomRef = "UPI-" + new string(Enumerable.Repeat(chars, 12)
                .Select(s => s[random.Next(s.Length)]).ToArray());
            return randomRef;
        }
    }
}
