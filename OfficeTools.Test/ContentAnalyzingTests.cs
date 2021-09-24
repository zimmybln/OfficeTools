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

            using (WordprocessingDocument output =
                WordprocessingDocument.Open(fileName, false))
            {
                Document document = output.MainDocumentPart?.Document;

                if (document == null)
                    return;

                Body body = document.Body;

                if (body == null)
                    return;

                foreach (Run run in body.Descendants<Run>())
                {
                    // <Run><RunProperties><Bold>...

                    string runDescriptor = String.Empty;

                    RunProperties runProperties = run.Descendants<RunProperties>().FirstOrDefault();

                    if (runProperties == null)
                        continue;
                    
                    if (runProperties.Bold != null && runProperties.Bold.Val?.Value != false)
                    {
                        Console.WriteLine($"{run.InnerText}");    
                    }
                }
            }
        }

        public void RetrieveParagraphStyleProperties()
        {

        }
    }
}
