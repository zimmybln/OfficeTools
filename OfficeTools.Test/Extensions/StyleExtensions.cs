using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeTools.Extensions
{

    // https://docs.microsoft.com/de-de/dotnet/api/documentformat.openxml.wordprocessing.style?view=openxml-2.8.1

    public static class StyleExtensions
    {
        public static bool IsBold(this Style style)
        {
            if (style == null)
                throw new ArgumentNullException();

            if (style.StyleRunProperties == null)
            {
                return false;
            }
            
            var boldNode = style.StyleRunProperties.OfType<Bold>()?.FirstOrDefault();
            
            if (boldNode == null || boldNode.Val == null || !boldNode.Val.HasValue)
                return false;

            return boldNode.Val.Value;
        }

        public static void SetBold(this Style style, bool isBold)
        {
            if (style == null)
                throw new ArgumentNullException();

            StyleRunProperties styleRunProperties = null;

            if (style.StyleParagraphProperties == null)
            {
                styleRunProperties = new StyleRunProperties();
                style.StyleRunProperties = styleRunProperties;
            }
            else
            {
                styleRunProperties = style.StyleRunProperties;
            }

            var boldNode = styleRunProperties.OfType<Bold>()?.FirstOrDefault();

            if (boldNode == null && !isBold)
            {
                return;
            }
            
            if (boldNode == null)
            {
                boldNode = new Bold();
                boldNode.Val = new OnOffValue(isBold);
                styleRunProperties.AppendChild(boldNode);
            }
            else
            {
                boldNode.Val.Value = isBold;    
            }
        }

        public static bool GetItalic(this Style style)
        {
            if (style == null)
                throw new ArgumentNullException();

            var italicNode = style.StyleRunProperties?.OfType<Italic>()?.FirstOrDefault();

            if (italicNode == null || italicNode.Val == null || !italicNode.Val.HasValue)
                return false;

            return italicNode.Val.Value;
        }

        public static void SetItalic(this Style style, bool isItalic)
        {
            if (style == null)
                throw new ArgumentNullException();

            StyleRunProperties styleRunProperties = null;

            if (style.StyleParagraphProperties == null)
            {
                styleRunProperties = new StyleRunProperties();
                style.StyleRunProperties = styleRunProperties;
            }
            else
            {
                styleRunProperties = style.StyleRunProperties;
            }

            var italicNode = styleRunProperties.OfType<Italic>()?.FirstOrDefault();

            if (italicNode == null && !isItalic)
            {
                return;
            }

            if (italicNode == null)
            {
                italicNode = new Italic();
                italicNode.Val = new OnOffValue(isItalic);
                styleRunProperties.AppendChild(italicNode);
            }
            else
            {
                italicNode.Val.Value = isItalic;
            }
        }

        public static void SetFontName(this Style style, string fontName)
        {
            if (style == null)
                throw new ArgumentNullException();

            StyleRunProperties styleRunProperties = null;

            if (style.StyleParagraphProperties == null)
            {
                styleRunProperties = new StyleRunProperties();
                style.StyleRunProperties = styleRunProperties;
            }
            else
            {
                styleRunProperties = style.StyleRunProperties;
            }

            //var fontNode = styleRunProperties.OfType<RunFonts>()?.FirstOrDefault();

            //if (fontNode == null)
            //{
            //    fontNode = new RunFonts() { Ascii = fontName};
            //    styleRunProperties.AppendChild(fontNode);
            //}
            //else
            //{
            //    fontNode.Ascii = fontName;
            //}

            styleRunProperties.RunFonts = new RunFonts() {Ascii = fontName};


        }

        public static float GetFontSize(this Style style)
        {
            if (style == null)
                throw new ArgumentNullException();

            StyleRunProperties styleRunProperties = null;

            if (style.StyleParagraphProperties == null)
                return Single.NaN;
                
            styleRunProperties = style.StyleRunProperties;
            

            var fontSizeNode = styleRunProperties.OfType<FontSize>()?.FirstOrDefault();

            if (fontSizeNode == null)
                return Single.NaN;

            var fontSizeValue = fontSizeNode.Val;

            if (float.TryParse(fontSizeValue, out float fontSizeResult))
            {
                return fontSizeResult;
            }

            return Single.NaN;

        }

        public static string GetFontName(this Style style)
        {
            if (style == null)
                throw new ArgumentNullException();

            StyleRunProperties styleRunProperties = null;

            if (style.StyleParagraphProperties == null)
                return null;

            styleRunProperties = style.StyleRunProperties;


            var fontNameNode = styleRunProperties.OfType<RunFonts>()?.FirstOrDefault();

            return fontNameNode?.Ascii;

        }

        public static void SetFontSize(this Style style, float fontSize)
        {
            if (style == null)
                throw new ArgumentNullException();

            StyleRunProperties styleRunProperties = null;

            if (style.StyleParagraphProperties == null)
            {
                styleRunProperties = new StyleRunProperties();
                style.StyleRunProperties = styleRunProperties;
            }
            else
            {
                styleRunProperties = style.StyleRunProperties;
            }

            styleRunProperties.FontSize = new FontSize() {Val = (fontSize * 2).ToString()};

            //var fontSizeNode = styleRunProperties.OfType<FontSize>()?.FirstOrDefault();

            //if (fontSizeNode == null)
            //{
            //    fontSizeNode = new FontSize() { Val = (fontSize * 2).ToString()};
            //    styleRunProperties.AppendChild(fontSizeNode);
            //}
            //else
            //{
            //    fontSizeNode.Val = (fontSize * 2).ToString();
            //}
        }

        public static void ListStyles(this WordprocessingDocument wordprocessingDocument)
        {
            if (wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart != null)
            {
                Styles s = wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles;

                foreach (Style style in s.Elements<Style>().Where(st => st.Type == StyleValues.Paragraph))
                {
                    Debug.WriteLine($"Style: {style.StyleName.Val}");

                    if (style.StyleRunProperties == null)
                    {
                        Console.WriteLine("\tKeine Eigenschaften vorhanden");
                        continue;
                    }

                    // Eigenschaften auflisten

                    Debug.WriteLine($"\tBold {style.IsBold()}");
                    Debug.WriteLine($"\tItalic {style.GetItalic()}");

                    //foreach (var property in style.StyleRunProperties.OfType<Bold>())
                    //{
                    //    Console.WriteLine($"\tTyp {property.LocalName}");
                    //}
                }
            }
        }

        /// <summary>
        /// Ermittelt einen Style anhand seines Namens
        /// </summary>
        public static DocumentFormat.OpenXml.Wordprocessing.Style GetStyle(
            this WordprocessingDocument wordprocessingDocument, string name)
        {
            return wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart?.Styles
                .Elements<Style>().FirstOrDefault(st => st.Type == StyleValues.Paragraph
                                                        && st.StyleName.Val.HasValue
                                                        && st.StyleName.Val.Value.Equals(name));
        }


        public static DocumentFormat.OpenXml.Wordprocessing.Style GetDefaultStyle(
            this WordprocessingDocument wordprocessingDocument)
        {
            return wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart?.Styles
                .Elements<Style>().FirstOrDefault(st => st.Type == StyleValues.Paragraph
                                                        && st.Default?.Value == true);
        }

    }
}
