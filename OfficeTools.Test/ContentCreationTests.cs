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
            const string fileName = "Samples//created_content.docx";

            try
            {
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
            finally
            {

            }


        }


        [Test]
        public void CreateDocumentWithContent_PT()
        {
            const string fileName = "Samples//created_content_pt.docx";

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

        [Test]
        public void CreateDocumentWithProperties()
        {
            throw new NotImplementedException();
        }

        [Test]
        public void CreateDocumentWithStyles()
        {
            throw new NotImplementedException();
        }

        [Test]
        public void MergeDocuments()
        {
            throw new NotImplementedException();
        }
    }
}
