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
    class ContentCreationTests
    {
        [SetUp]
        public void Setup()
        {
            if (!Directory.Exists("Samples"))
            {
                Directory.CreateDirectory("Samples");
            }
        }

        [Test]
        public void CreateDocumentWithContent()
        {
            string fileName = $"Samples//created_{MethodBase.GetCurrentMethod()?.Name}.docx";

            if (File.Exists(fileName))
                File.Delete(fileName);


            using (WordprocessingDocument output = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
            {
                output.AddMainDocumentPart();

                                var text = new Text($"Test {MethodBase.GetCurrentMethod().Name} {DateTime.Now.ToString()}");

                var run = new Run(text);

                var paragraph = new Paragraph(run);

                var body = new Body(paragraph);

                output.MainDocumentPart.Document = new Document(body);
                
                output.Save();
            }


            Assert.IsTrue(File.Exists(fileName));
        }

        [Test]
        public void CreateDocumentWithContentAsBoldFormattedText()
        {
            string fileName = $"Samples//created_{MethodBase.GetCurrentMethod()?.Name}.docx";

            if (File.Exists(fileName))
                File.Delete(fileName);


            using (WordprocessingDocument output = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
            {
                output.AddMainDocumentPart();

                var text = new Text($"Test {MethodBase.GetCurrentMethod()?.Name} {DateTime.Now.ToString()}");

                var run = new Run(text);
                run.RunProperties = new RunProperties()
                {
                    Bold = new Bold()
                };


                var paragraph = new Paragraph(run);

                var body = new Body(paragraph);

                output.MainDocumentPart.Document = new Document(body);

                output.Save();
            }


            Assert.IsTrue(File.Exists(fileName));
        }

        [Test]
        public void CreateDocumentWithContentAsBoldFormattedStyle()
        {
            string fileName = $"Samples//created_{MethodBase.GetCurrentMethod()?.Name}.docx";

            if (File.Exists(fileName))
                File.Delete(fileName);


            using (WordprocessingDocument document = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
            {
                document.AddMainDocumentPart();

                if (document.MainDocumentPart == null)
                    throw new InvalidOperationException();

                // Hinzufügen der Stylesinformationen


                // Erstellen eines Styles
                Style styleWithBold = new Style()
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "AlsFettFormatierteVorlage",
                    CustomStyle = true,
                    Default = false,
                    StyleName = new StyleName()
                    {
                        Val = "Als Fett formatierte Vorlage"
                    },
                    StyleRunProperties = new StyleRunProperties()
                    {
                        Bold = new Bold()
                    }
                };

                // Erstellen der Dokumentinformationen für Styles
                StyleDefinitionsPart stylesPart = document.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();

                stylesPart.Styles = new Styles();

                stylesPart.Styles.Append(styleWithBold);

                // Erstellen des Textes
                var text = new Text($"Test {MethodBase.GetCurrentMethod()?.Name} {DateTime.Now.ToString()}");

                var run = new Run(text);

                // Erstellen des Absatzes mit dem Verweis auf die erstellte Formatvorlage
                var paragraph = new Paragraph(run)
                {
                    ParagraphProperties = new ParagraphProperties()
                    {
                        ParagraphStyleId = new ParagraphStyleId()
                        {
                            Val = "AlsFettFormatierteVorlage"
                        }
                    }
                };

                var body = new Body(paragraph);

                document.MainDocumentPart.Document = new Document(body);

                document.Save();
            }


            Assert.IsTrue(File.Exists(fileName));

        }

        [Test]
        public void CreateDocumentFailingByDoubleStyle()
        {
            string fileName = $"Samples//created_{MethodBase.GetCurrentMethod()?.Name}.docx";

            if (File.Exists(fileName))
                File.Delete(fileName);


            using (WordprocessingDocument document = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
            {
                document.AddMainDocumentPart();

                if (document.MainDocumentPart == null)
                    throw new InvalidOperationException();

                // Erstellen eines Styles
                Style styleWithBold_First = new Style()
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "AlsFettFormatierteVorlage",
                    CustomStyle = true,
                    Default = false,
                    StyleName = new StyleName()
                    {
                        Val = "Als Fett formatierte Vorlage"
                    },
                    StyleRunProperties = new StyleRunProperties()
                    {
                        Bold = new Bold()
                    }
                };

                Style styleWithBold_Second = new Style()
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "AlsFettFormatierteVorlage",
                    CustomStyle = true,
                    Default = false,
                    StyleName = new StyleName()
                    {
                        Val = "Als Fett formatierte Vorlage"
                    },
                    StyleRunProperties = new StyleRunProperties()
                    {
                        Bold = new Bold()
                    }
                };

                // Erstellen der Dokumentinformationen für Styles
                StyleDefinitionsPart stylesPart = document.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();

                stylesPart.Styles = new Styles();

                stylesPart.Styles.Append(styleWithBold_First);
                stylesPart.Styles.Append(styleWithBold_Second);

                // Erstellen des Textes
                var text = new Text($"Test {MethodBase.GetCurrentMethod()?.Name} {DateTime.Now.ToString()}");

                var run = new Run(text);

                // Erstellen des Absatzes mit dem Verweis auf die erstellte Formatvorlage
                var paragraph = new Paragraph(run)
                {
                    ParagraphProperties = new ParagraphProperties()
                    {
                        ParagraphStyleId = new ParagraphStyleId()
                        {
                            Val = "AlsFettFormatierteVorlage"
                        }
                    }
                };

                var body = new Body(paragraph);

                document.MainDocumentPart.Document = new Document(body);

                document.Save();
            }


            Assert.IsTrue(File.Exists(fileName));

        }

        [Test]
        public void CreateDocumentWithContent_PT()
        {
            string fileName = $"Samples//created_{MethodBase.GetCurrentMethod()?.Name}.docx";

            try
            {
                if (File.Exists(fileName))
                    File.Delete(fileName);

                using (OpenXmlMemoryStreamDocument streamDoc = OpenXmlMemoryStreamDocument.CreateWordprocessingDocument())
                {
                    using (WordprocessingDocument output = streamDoc.GetWordprocessingDocument())
                    {
                        output.MainDocumentPart.Document =
                            new Document(
                                new Body(
                                    new Paragraph(
                                        new Run(
                                            new Text($"Test {MethodBase.GetCurrentMethod().Name} {DateTime.Now.ToString()}")
                                        )
                                    )
                                 )
                             );

                        output.SaveAs(fileName);
                    }
                }

                Assert.IsTrue(File.Exists(fileName));
            }
            finally
            {

            }


        }


    }
}
