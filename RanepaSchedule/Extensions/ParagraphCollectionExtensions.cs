using System;
using System.Text;
using Spire.Doc.Collections;
using Spire.Doc.Documents;

namespace RanepaSchedule.Extensions
{
    public static class ParagraphCollectionExtensions
    {
        public static string GetText(this ParagraphCollection paragraphs)
        {
            var text = new StringBuilder();
            foreach (Paragraph paragraph in paragraphs)
            {
                var paragraphText = paragraph.Text.Trim().Replace("\n", String.Empty).Replace("\t", String.Empty);
                if (!string.IsNullOrEmpty(paragraphText) && !string.IsNullOrWhiteSpace(paragraphText))
                    text.Append(paragraph.Text).Append(' ');
            }

            return text.ToString().Trim();
        }
    }
}