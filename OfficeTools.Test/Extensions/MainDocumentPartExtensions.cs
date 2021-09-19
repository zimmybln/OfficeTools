using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeTools.Extensions
{
    public delegate void ListHyperlinkDelegate(string id, string target, string text);


    public static class MainDocumentPartExtensions
    {
        public static int ListHyperlinks(this MainDocumentPart documentPart, ListHyperlinkDelegate funcListHyperlink)
        {
            int counter = 0;
            var hyperLinkRelations = documentPart.HyperlinkRelationships?.ToList();

            if (hyperLinkRelations == null)
                return 0;

            // HyperlinkRelationship 
            foreach (var hyperlink in documentPart.Document.Body.Descendants<Hyperlink>())
            {

                var hyperlinkRelation = hyperLinkRelations.FirstOrDefault(l => l.Id.Equals(hyperlink.Id));

                if (hyperlinkRelation != null)
                {
                    funcListHyperlink?.Invoke(hyperlink.Id, hyperlinkRelation.Uri.ToString(), hyperlink.InnerText);

                    counter++;
                }
            }

            return counter;
        }
    }
}
