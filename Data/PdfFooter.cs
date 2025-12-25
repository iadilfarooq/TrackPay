using iTextSharp.text.pdf;
using System.Reflection.Metadata;

namespace TrackPay.Data
{
    using iTextSharp.text;
    using iTextSharp.text.pdf;

    public class PdfFooter : PdfPageEventHelper
    {
        private PdfTemplate template;
        private BaseFont baseFont;
        private string appName;
        private int pageCount; // Track page count ourselves

        public PdfFooter(string applicationName)
        {
            appName = applicationName;
            pageCount = 0;
        }

        public override void OnOpenDocument(PdfWriter writer, Document document)
        {
            template = writer.DirectContent.CreateTemplate(50, 50);
            baseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        }

        public override void OnEndPage(PdfWriter writer, Document document)
        {
            pageCount++;
            PdfContentByte cb = writer.DirectContent;

            // Configuration variables - adjust these values as needed
            float verticalPosition = 20f; // Distance from bottom (increase to move footer up)
            float leftMargin = document.LeftMargin; // Left side margin
            float rightMargin = document.RightMargin; // Right side margin
            float rightTextOffset = 8f; // Additional right padding for page numbers

            // Left-aligned app name
            cb.BeginText();
            cb.SetFontAndSize(baseFont, 8);
            cb.SetTextMatrix(leftMargin, verticalPosition); // X, Y coordinates
            cb.ShowText(appName);
            cb.EndText();

            // Right-aligned page number ("Page X of Y")
            string pageText = $"Page {writer.PageNumber} of ";
            float textWidth = baseFont.GetWidthPoint(pageText, 8);

            cb.BeginText();
            cb.SetFontAndSize(baseFont, 8);
            // Calculate X position: page width - right margin - text width - additional offset
            cb.SetTextMatrix(document.PageSize.Width - rightMargin - textWidth - rightTextOffset, verticalPosition);
            cb.ShowText(pageText);
            cb.EndText();

            // Add the total page count template
            cb.AddTemplate(template, document.PageSize.Width - rightMargin - rightTextOffset, verticalPosition);
        }

        public override void OnCloseDocument(PdfWriter writer, Document document)
        {
            // Use our tracked page count instead of writer.PageNumber
            template.BeginText();
            template.SetFontAndSize(baseFont, 8);
            template.SetTextMatrix(0, 0);
            template.ShowText(pageCount.ToString());
            template.EndText();
        }
    }
}
