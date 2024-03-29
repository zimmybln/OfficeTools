﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;
using OfficeTools.Extensions;

namespace OfficeTools.Test
{
    class FieldsTests
    {
        [Test]
        public void CollectFields()
        {
            using WordprocessingDocument document = WordprocessingDocument.Open("Samples//DocumentWithFields.docx", false);
            
            foreach (FieldCode simpleField in document.MainDocumentPart.Document.Body.Descendants<FieldCode>())
            {
                Console.WriteLine($"Feld: {simpleField.InnerText}");
            }

        }
    }
}
