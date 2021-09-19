using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

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

                    var bold = style.Descendants<Bold>();





                    // Eigenschaften auflisten

                    //Console.WriteLine($"\tBold {style.IsBold()}");
                    //Console.WriteLine($"\tItalic {style.GetItalic()}");
                    //Console.WriteLine($"\tFontsize {style.GetFontSize()}");

                    //foreach (var property in style.StyleRunProperties.OfType<Bold>())
                    //{
                    //    Console.WriteLine($"\tTyp {property.LocalName}");
                    //}
                }
            }





            document.Close();


        }
    }
}
