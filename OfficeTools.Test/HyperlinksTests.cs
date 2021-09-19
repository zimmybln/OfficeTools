using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;
using OfficeTools.Extensions;
using OpenXmlPowerTools;

namespace OfficeTools.Test
{
    class HyperlinksTests
    {
        [Test]
        public void CollectHyperlinks()
        {
            using WordprocessingDocument document = WordprocessingDocument.Open("Samples//HyperlinkDocument.docx", false);

            var hyperLinkRelations = document.MainDocumentPart.HyperlinkRelationships.ToList();

            // HyperlinkRelationship 
            foreach (var hyperlink in document.MainDocumentPart.Document.Body.Descendants<Hyperlink>())
            {
                // ToDo: Ziel fehlt
                Console.WriteLine($"Hyperlink: Id: '{hyperlink.Id}', Text: '{hyperlink.InnerText}', Target: ");

                var hyperlinkRelation = hyperLinkRelations.FirstOrDefault(l => l.Id.Equals(hyperlink.Id));

                if (hyperlinkRelation != null)
                {
                    Console.WriteLine($"Target = {hyperlinkRelation.Uri}");
                }
            }
        }

        [Test]
        public void ChangeHyperlinks()
        {

        }

        [Test]
        public void RemoveHyperlink()
        {

        }

        [Test]
        public void CreateHyperlink()
        {

            const string fileName = "Files//createhyperlink.docx";

            if (File.Exists(fileName))
                File.Delete(fileName);

            using (OpenXmlMemoryStreamDocument streamDoc = OpenXmlMemoryStreamDocument.CreateWordprocessingDocument())
            {
                using (WordprocessingDocument output = streamDoc.GetWordprocessingDocument())
                {
                    var paragraph = new Paragraph(
                        new Run(
                            new Text($"Test {MethodBase.GetCurrentMethod().Name} {DateTime.Now.ToString()}")
                        )
                    );


                    output.MainDocumentPart.Document =
                        new Document(
                            new Body(paragraph)
                         );

                    paragraph.AppendChild(output.PrepareHyperlink("Hier geht es zu Microsoft", new Uri("http://www.microsoft.com"), "Das ist ein Tooltip für den Link"));

                    output.SaveAs(fileName);
                }
            }
        }

        //private static Hyperlink CreateDocumentHyperlink(MainDocumentPart mainPart, string url, string text)
        //{

        //    HyperlinkRelationship hr = mainPart.AddHyperlinkRelationship(new Uri(url), true);
        //    string hrContactId = hr.Id;
        //    return
        //        new Hyperlink(
        //            new ProofError() { Type = ProofingErrorValues.GrammarStart },
        //            new Run(
        //                new RunProperties(
        //                    new RunStyle() { Val = "Hyperlink" },
        //                    new Color { ThemeColor = ThemeColorValues.Hyperlink }),
        //                new Text(text) { Space = SpaceProcessingModeValues.Preserve }
        //            ))
        //        { History = OnOffValue.FromBoolean(true), Id = hrContactId };
        //}
    }
}
