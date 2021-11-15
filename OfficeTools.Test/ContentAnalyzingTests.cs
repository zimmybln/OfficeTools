using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace OfficeTools.Test
{
    public class ContentAnalyzingTests
    {
        [Test]
        public void RetrieveParagraphText()
        {
            const string fileName = "Samples//DocumentWithContent.docx";

            using (WordprocessingDocument output =
                WordprocessingDocument.Open(fileName, false))
            {
                Document document = output.MainDocumentPart?.Document;

                if (document == null)
                    return;

                Body body = document.Body;

                if (body == null)
                    return;

                foreach (Paragraph paragraph in body.Descendants<Paragraph>())
                {
                    var paragraphProperties = paragraph.Descendants<ParagraphProperties>().FirstOrDefault();

                    if (paragraphProperties != null)
                    {
                        Console.WriteLine($"Style gefunden {paragraphProperties.ParagraphStyleId.Val}");
                    }



                    var innerText = paragraph.InnerText;

                    if (!string.IsNullOrEmpty((innerText)))
                    {
                        Console.WriteLine((paragraph.InnerText));
                    }
                }
            }
        }

        [Test]
        public void RetrieveParagraphChildren()
        {
            const string fileName = "Samples//DocumentWithContent.docx";

            using (WordprocessingDocument output =
                WordprocessingDocument.Open(fileName, false))
            {
                Document document = output.MainDocumentPart?.Document;

                if (document == null)
                    return;

                Body body = document.Body;

                if (body == null)
                    return;

                foreach (Paragraph paragraph in body.Descendants<Paragraph>())
                {
                    var descendants = paragraph.Descendants().ToList();

                    if (descendants.Any())
                    {
                        Console.WriteLine("Absatzanfang");
                        foreach (OpenXmlElement child in paragraph.Descendants())
                        {
                            Console.WriteLine($"\t{child.LocalName}, {child.InnerText}");
                        }

                        Console.WriteLine("Absatzende");
                    }
                }
            }
        }

        [Test]
        public void RetrieveBoldFormattedText()
        {
            const string fileName = "Samples//DocumentWithContent.docx";

            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Open(fileName, false))
            {
                Document document = wordDocument.MainDocumentPart?.Document;

                if (document == null)
                    return;

                Body body = document.Body;

                if (body == null)
                    return;

                foreach (Paragraph paragraph in body.Descendants<Paragraph>())
                {
                    // Überprüfen, ob die Formatvorlage kursiv formattiert wurde
                    var paragraphProperties = paragraph.Descendants<ParagraphProperties>().FirstOrDefault();

                    if (paragraphProperties != null)
                    {
                        // Ermittle die Formatvorlage
                        var style = wordDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles
                            .Elements<Style>().FirstOrDefault(st => st.Type == StyleValues.Paragraph
                                                                    && st.StyleId.HasValue
                                                                    && st.StyleId.Value.Equals(paragraphProperties.ParagraphStyleId?.Val));

                        // Ermittle die Kursiv-Kennzeichnung
                        var boldNode = style?.StyleRunProperties?.OfType<Bold>().FirstOrDefault();

                        if (boldNode != null && boldNode.Val?.Value == true)
                        {
                            Console.WriteLine(paragraph.InnerText);
                            continue;
                        }
                    }

                    // Die einzelnen Bestandteile des Absatzes werden nur durchlaufen,
                    // wenn nicht der Absatz mit einer Formatvorlage gekennzeichnet ist,
                    // die kursiv formatiert wurde

                    foreach (Run run in paragraph.Descendants<Run>())
                    {
                        RunProperties runProperties = run.Descendants<RunProperties>().FirstOrDefault();

                        if (runProperties == null)
                            continue;

                        if (runProperties.Bold != null && runProperties.Bold.Val?.Value == true)
                        {
                            Console.WriteLine($"{run.InnerText}");
                        }
                    }

                }
            }
        }

        public void RetrieveParagraphStyleProperties()
        {

        }
    }
}
