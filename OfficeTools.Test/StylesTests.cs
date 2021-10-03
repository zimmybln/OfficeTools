using System;
using System.Collections.Generic;
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
    class StylesTests
    {
        [Test]
        public void ListStyles()
        {
            using WordprocessingDocument document = WordprocessingDocument.Open("Samples//HeadStyles.docx", false);

            if (document.MainDocumentPart.StyleDefinitionsPart == null)
            {
                Console.WriteLine("Es sind keine Styles vorhanden");
            }
            else
            {

                Styles s = document.MainDocumentPart.StyleDefinitionsPart.Styles;

                foreach (Style style in s.Elements<Style>().Where(st => st.Type == StyleValues.Paragraph))
                {
                    Console.WriteLine($"Style: {style.StyleName.Val}");

                    
                    if (style.StyleRunProperties == null)
                    {
                        Console.WriteLine("\tKeine Eigenschaften vorhanden");
                        continue;
                    }

                    var bold = style.Descendants<Bold>().ToList();


                    Console.WriteLine($"Bold vorhanden: {bold.Any()}");


                    // Eigenschaften auflisten

                    Console.WriteLine($"\tBold {style.IsBold()}");
                    Console.WriteLine($"\tItalic {style.GetItalic()}");
                    Console.WriteLine($"\tFontsize {style.GetFontSize()}");
                  

                }
            }





            document.Close();


        }

        [Test]
        public void ApplyStyleToParagraph()
        {
            // Erstellen des Dokumentes
            using OpenXmlMemoryStreamDocument streamDoc = OpenXmlMemoryStreamDocument.CreateWordprocessingDocument();
            using WordprocessingDocument document = streamDoc.GetWordprocessingDocument();

            // Hinzufügen des Textes
            Paragraph p = document.AddCommonParagraph($"Test {MethodBase.GetCurrentMethod().Name} {DateTime.Now.ToString()}");

            // Eigenschaft erstellen
            if (p.Elements<ParagraphProperties>().Count() == 0)
            {
                p.PrependChild<ParagraphProperties>(new ParagraphProperties());
            }

            // Ermitteln des Bereichs für Styles
            StyleDefinitionsPart part = document.EnsureStylesPart();

            // Überprüfen, ob der Style existiert
            var styleId = "Frage";
            var styleName = "Frage";

            if (!document.IsStyleIdInDocument(styleId))
            {
                part.AddNewStyle(styleId, styleName, "Lucida Console", 24);
            }

            document.ApplyStyleToParagraph(styleId, styleName, p);

            document.SaveAs("Samples//created_testwithformat.docx");
        }
    }
}
