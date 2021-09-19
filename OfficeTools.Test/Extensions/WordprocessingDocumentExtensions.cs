    using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;

namespace OfficeTools.Extensions
{
    public static class WordprocessingDocumentExtensions
    {
        public static Paragraph AddCommonParagraph(this WordprocessingDocument document, string text)
        {
            Paragraph paragraph = new Paragraph( new Run( new Text(text)));

            document.MainDocumentPart.Document =
                    new Document(new Body(paragraph));

            return paragraph;
        }

        public static StyleDefinitionsPart AddStylesPartToPackage(this WordprocessingDocument document)
        {
            StyleDefinitionsPart part;
            part = document.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            Styles root = new Styles();
            root.Save(part);

            return part;
        }

        /// <summary>
        /// Stellt sicher, dass es einen Abschnitt für Styles gibt.
        /// </summary>
        public static StyleDefinitionsPart EnsureStylesPart(this WordprocessingDocument document)
        {
            StyleDefinitionsPart part = document.MainDocumentPart.StyleDefinitionsPart ?? document.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();

            part.Styles ??= new Styles();

            return part;
        }

        public static bool IsStyleIdInDocument(this WordprocessingDocument document, string styleid)
        {
            // Get access to the Styles element for this document.
            Styles s = document.MainDocumentPart.StyleDefinitionsPart.Styles;

            // Check that there are styles and how many.
            int n = s.Elements<Style>().Count();
            if (n == 0)
                return false;

            // Look for a match on styleid.
            Style style = s.Elements<Style>()
                .Where(st => (st.StyleId == styleid) && (st.Type == StyleValues.Paragraph))
                .FirstOrDefault();

            if (style == null)
                return false;

            return true;
        }

        public static void AddNewStyle(this StyleDefinitionsPart styleDefinitionsPart, string styleid, string stylename, 
                                        string fontName, float fontSize, bool isDefault = false)
        {
            // Get access to the root element of the styles part.
            Styles styles = styleDefinitionsPart.Styles;

            // Create a new paragraph style and specify some of the properties.
            Style style = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = styleid,
                CustomStyle = true,
                Default = isDefault
            };

            StyleName styleName1 = new StyleName() { Val = stylename };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = stylename };

            style.Append(styleName1);
            style.Append(basedOn1);
            style.Append(nextParagraphStyle1);

            // Create the StyleRunProperties object and specify some of the run properties.
            StyleRunProperties styleRunProperties1 = new StyleRunProperties
            {
                FontSize = new FontSize() {Val = (fontSize * 2).ToString()},
                RunFonts = new RunFonts() {Ascii = fontName}
            };


            ////Color color1 = new Color() { ThemeColor = ThemeColorValues.Accent2 };
            //RunFonts font1 = 
            
            //// Specify a 12 point size.
            //FontSize fontSize1 = 
            //styleRunProperties1.Append(font1);
            //styleRunProperties1.Append(fontSize1);

            // Add the run properties to the style.
            style.Append(styleRunProperties1);

            // Add the style to the styles part.
            styles.Append(style);
        }

        public static void ApplyStyleToParagraph(this WordprocessingDocument doc, string styleid, string stylename, Paragraph p)
        {
            // If the paragraph has no ParagraphProperties object, create one.
            if (p.Elements<ParagraphProperties>().Count() == 0)
            {
                p.PrependChild<ParagraphProperties>(new ParagraphProperties());
            }
                        
            // Get the paragraph properties element of the paragraph.
            ParagraphProperties pPr = p.Elements<ParagraphProperties>().First();

            // Get the Styles part for this document.
            StyleDefinitionsPart part =
                doc.MainDocumentPart.StyleDefinitionsPart;

            // If the Styles part does not exist, add it and then add the style.
            if (part == null)
            {
                part = AddStylesPartToPackage(doc);
                AddNewStyle(part, styleid, stylename, "Lucida Console", 24);
            }
            else
            {
                // If the style is not in the document, add it.
                if (IsStyleIdInDocument(doc, styleid) != true)
                {
                    // No match on styleid, so let's try style name.
                    string styleidFromName = GetStyleIdFromStyleName(doc, stylename);
                    if (styleidFromName == null)
                    {
                        AddNewStyle(part, styleid, stylename, "Lucida Console", 24);
                    }
                    else
                        styleid = styleidFromName;
                }
            }

            // Set the style of the paragraph.
            pPr.ParagraphStyleId = new ParagraphStyleId() { Val = styleid };

            
        }

        // Return styleid that matches the styleName, or null when there's no match.
        public static string GetStyleIdFromStyleName(WordprocessingDocument doc, string styleName)
        {
            StyleDefinitionsPart stylePart = doc.MainDocumentPart.StyleDefinitionsPart;
            string styleId = stylePart.Styles.Descendants<StyleName>()
                .Where(s => s.Val.Value.Equals(styleName) &&
                    (((Style)s.Parent).Type == StyleValues.Paragraph))
                .Select(n => ((Style)n.Parent).StyleId).FirstOrDefault();
            return styleId;
        }

        public static Hyperlink PrepareHyperlink(this WordprocessingDocument document, string title, Uri target, string tooltip = null)
        {
            string relationshipId = null;


            foreach (HyperlinkRelationship hRel in document.MainDocumentPart.HyperlinkRelationships)
            {
                if (hRel.Uri.OriginalString == target.OriginalString)
                {
                    relationshipId = hRel.Id;
                    break;
                }
            }

            if (string.IsNullOrEmpty(relationshipId))
            {
                HyperlinkRelationship hr = document.MainDocumentPart.AddHyperlinkRelationship(target, true);
                relationshipId = hr.Id;
            }
           

            Hyperlink hyperlink = new Hyperlink(
                        new ProofError() { Type = ProofingErrorValues.GrammarStart }
                        )
            {
                Id = relationshipId
            };

            if (!string.IsNullOrEmpty(tooltip))
            {
                hyperlink.Tooltip = tooltip;
            }

            hyperlink.AppendChild(new Run(
                    new RunProperties(
                            new RunStyle() { Val = "Hyperlink" },
                                              new Color { ThemeColor = ThemeColorValues.Hyperlink }),
                                      new Text(title)  { Space = SpaceProcessingModeValues.Preserve }));

            return hyperlink;
        }

        public static byte[] ToByteArray(this WordprocessingDocument document, bool isEditable = false)
        {
            Byte[] buffer;

            using var memory = new MemoryStream();

            document.Clone(memory, isEditable);

            memory.Flush();
            memory.Position = 0;

            buffer = new byte[memory.Length];

            memory.Read(buffer, 0, (int)memory.Length);

            return buffer;
        }
    }
}
