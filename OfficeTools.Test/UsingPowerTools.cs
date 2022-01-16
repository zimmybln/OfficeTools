using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using OpenXmlPowerTools;

namespace OfficeTools.Test
{
    [TestFixture]
    public class UsingPowerTools
    {
        [Test]
        public void CombineDocuments()
        {
            string firstDocument = "Samples//CombineDocument1.docx";
            string secondDocument = "Samples//CombineDocument2.docx";

            string targetDocument = "Samples//CombinedDocument.docx";

            List<Source> sourceDocuments = new ()
            {
                new Source(new WmlDocument(firstDocument)),
                new Source(new WmlDocument(secondDocument))
            };

            DocumentBuilder.BuildDocument(sourceDocuments, targetDocument);

        }
    }
}
