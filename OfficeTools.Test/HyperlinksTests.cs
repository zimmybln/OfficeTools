using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
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

            foreach (HyperlinkRelationship relationShip in hyperLinkRelations)
            {
                Console.WriteLine($"Type {relationShip.RelationshipType}, Id {relationShip.Id}, Uri: {relationShip.Uri}, IsExternal {relationShip.IsExternal}");
            }

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

             string fileName = $"Samples//created_{MethodBase.GetCurrentMethod()?.Name}.docx";

            if (File.Exists(fileName))
                File.Delete(fileName);


            using (WordprocessingDocument document = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
            {
                document.AddMainDocumentPart();

                if (document.MainDocumentPart == null)
                    throw new InvalidOperationException();

                // Erstellen eines Textes und hinzufügen zu einem Absatz
                var text = new Text($"Test {MethodBase.GetCurrentMethod().Name} {DateTime.Now.ToString()}");

                var run = new Run(text);

                var paragraph = new Paragraph(run);


                HyperlinkRelationship hr = document.MainDocumentPart.AddHyperlinkRelationship(new Uri("http://www.microsoft.com"), true);
                var relationshipId = hr.Id;

                Hyperlink hyperlink = new Hyperlink()
                {
                    Id = relationshipId
                };

                hyperlink.AppendChild(new Run(
                        new RunProperties(
                                new RunStyle() { Val = "Hyperlink" },
                                                  new Color { ThemeColor = ThemeColorValues.Hyperlink }),
                                          new Text("Hier geht es zu Microsoft") { Space = SpaceProcessingModeValues.Preserve }));

                var paragraphWithHyperlink = new Paragraph();

                paragraphWithHyperlink.AppendChild(hyperlink);

                var body = new Body();
                body.AppendChild(paragraph);
                body.AppendChild(paragraphWithHyperlink);


                document.MainDocumentPart.Document = new Document(body);
                document.Save();
            }

        }
    }
}
