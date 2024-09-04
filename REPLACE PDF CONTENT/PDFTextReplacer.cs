using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace REPLACE_PDF_CONTENT
{
    class PdfTextReplacer
    {
        public static void ReplaceTextInPdf(string inputFilePath, string outputFilePath, string searchText, string replaceText)
        {
            using (PdfReader pdfReader = new PdfReader(inputFilePath))
            using (PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(outputFilePath, FileMode.Create)))
            {
                for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                {
                    // Extract the text from the page
                    var content = PdfTextExtractor.GetTextFromPage(pdfReader, i);
                    if (content.Contains(searchText))
                    {
                        // Get the page content
                        PdfContentByte pdfContentByte = pdfStamper.GetOverContent(i);

                        // Replace the text by overlaying new text
                        ReplaceText(pdfContentByte, pdfReader, i, searchText, replaceText);
                    }
                }
            }
        }

        private static void ReplaceText(PdfContentByte pdfContentByte, PdfReader pdfReader, int pageNumber, string searchText, string replaceText)
        {
            // Create a content parser for the page
            PdfDictionary pageDict = pdfReader.GetPageN(pageNumber);
            PdfObject contentObject = pageDict.Get(PdfName.CONTENTS);
            PRStream stream = contentObject as PRStream;

            if (stream == null)
            {
                return;
            }

            // Extract the content bytes of the stream
            byte[] data = PdfReader.GetStreamBytes(stream);

            // Convert the content to a string
            string content = System.Text.Encoding.UTF8.GetString(data);

            // Replace the target text with the new text
            content = content.Replace(searchText, replaceText);

            // Convert the string back to bytes
            byte[] newData = System.Text.Encoding.UTF8.GetBytes(content);

            // Write the new content back to the PDF stream
            stream.SetData(newData);
        }
    }
}