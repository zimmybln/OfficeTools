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

namespace OfficeTools.Test
{
    public class TableTests
    {
        // https://docs.microsoft.com/en-us/office/open-xml/how-to-add-tables-to-word-processing-documents

        [SetUp]
        public void Setup()
        {
            if (!Directory.Exists("Samples"))
            {
                Directory.CreateDirectory("Samples");
            }
        }

        [Test]
        public void CreateSimpleTable()
        {
            string fileName = $"Samples//created_{MethodBase.GetCurrentMethod()?.Name}.docx";

            if (File.Exists(fileName))
                File.Delete(fileName);


            using (WordprocessingDocument output = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
            {
                output.AddMainDocumentPart();

                // create the table and it's properties
                Table table = new Table();

                TableProperties tableProperties = new TableProperties(
                    new TableBorders(
                        new TopBorder { Val = new (BorderValues.Single), Size = 12 },
                                        new BottomBorder { Val = new(BorderValues.Single), Size = 12 },
                                        new LeftBorder { Val = new (BorderValues.Single), Size = 12 },
                                        new RightBorder { Val = new (BorderValues.Single), Size = 12 },
                                        new InsideHorizontalBorder { Val = new (BorderValues.Single), Size = 12 },
                                        new InsideVerticalBorder { Val = new (BorderValues.Single), Size = 12 }));

                table.AppendChild<TableProperties>(tableProperties);

                // add rows and cells
                for (int i = 0; i < 5; i++)
                {
                    var row = new TableRow();

                    for (int j = 0; j < 5; j++)
                    {
                        var cell = new TableCell();

                        cell.Append(new Paragraph(
                            new Run(
                                new Text($"Zelle {i + 1}:{j + 1}"))));
                        row.Append(cell);
                    }

                    table.Append(row);
                }

                output.MainDocumentPart.Document = new Document(new Body(table));

                output.Save();
            }
        }

    }
}
