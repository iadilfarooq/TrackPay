using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.EMMA;
using iTextSharp.text.pdf;
using iTextSharp.text;
using Microsoft.AspNetCore.Mvc;
using TrackPay.Data;
using TrackPay.Models;

namespace TrackPay.Controllers
{
    public class UploadFilesController : Controller
    {
        private readonly UploadFilesDAL uploadFilesDAL;

        public UploadFilesController(IConfiguration configuration)
        {
            uploadFilesDAL = new UploadFilesDAL(configuration);
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View(new UploadFilesViewModel());
        }

        [HttpPost]
        public IActionResult Index(UploadFilesViewModel model, IFormCollection form)
        {
            // Extract month and year from "MonthYear" input
            string monthYear = form["MonthYear"];
            if (!string.IsNullOrEmpty(monthYear) &&
                DateTime.TryParseExact(monthYear + "-01", "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
            {
                model.Month = parsedDate.Month;
                model.Year = parsedDate.Year;
            }

            if (model.Month > 0 && model.Year > 0 && !string.IsNullOrEmpty(model.DataType))
            {
                if (model.DataType == "TaskData")
                {
                    model.TaskData = uploadFilesDAL.GetTaskDataOfRider(model.Month, model.Year, model.CourierId);
                }
                else if (model.DataType == "TimeStamps")
                {
                    model.TimeStamps = uploadFilesDAL.GetTimeStampsOfRider(model.Month, model.Year, model.CourierId);
                }
            }

            return View(model);
        }

        public ActionResult ExportMonthlyTaskDataToExcel(int month, int year, int? courierId = null)
        {
            var data = uploadFilesDAL.GetTaskDataOfRider(month, year, courierId);
            var courierInfo = data.FirstOrDefault();
            var monthName = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month);

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Monthly Task Data");

                // Headers
                worksheet.Cell(1, 1).Value = "#";
                worksheet.Cell(1, 2).Value = "Courier ID";
                worksheet.Cell(1, 3).Value = "Name";
                worksheet.Cell(1, 4).Value = "City";
                worksheet.Cell(1, 5).Value = "Purchase ID";
                worksheet.Cell(1, 6).Value = "Delivered Date & Time";
                worksheet.Cell(1, 7).Value = "Distance KM";

                int row = 2;
                int count = 1;

                foreach (var item in data)
                {
                    worksheet.Cell(row, 1).Value = count++;
                    worksheet.Cell(row, 2).Value = item.CourierID;
                    worksheet.Cell(row, 3).Value = item.Name;
                    worksheet.Cell(row, 4).Value = item.City;
                    worksheet.Cell(row, 5).Value = item.PurchaseID;
                    worksheet.Cell(row, 6).Value = item.DeliveredDateTime.ToString("dd/MM/yyyy HH:mm:ss");
                    worksheet.Cell(row, 7).Value = item.DistanceKM;

                    // Format distance to 2 decimal places
                    worksheet.Cell(row, 7).Style.NumberFormat.Format = "0.00";

                    row++;
                }

                // Formatting
                worksheet.Range("A1:G1").Style.Font.Bold = true;
                worksheet.Columns().AdjustToContents();
                worksheet.SheetView.Freeze(1, 0); // Freeze header row

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var fileName = courierId.HasValue
                        ? $"TaskData_{courierInfo?.Name}_{courierId}_{monthName}_{year}.xlsx"
                        : $"TaskData_AllRiders_{monthName}_{year}.xlsx";

                    return File(stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        fileName);
                }
            }
        }

        public ActionResult ExportMonthlyTaskDataToPdf(int month, int year, int? courierId = null)
        {
            var data = uploadFilesDAL.GetTaskDataOfRider(month, year, courierId);
            var courierInfo = data.FirstOrDefault();
            var monthName = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month);

            using (var memoryStream = new MemoryStream())
            {
                // Document setup (A4 landscape for better fit)
                var document = new Document(PageSize.A4.Rotate(), 15f, 15f, 15f, 30f);
                var writer = PdfWriter.GetInstance(document, memoryStream);
                writer.PageEvent = new PdfFooter("TrackPay Application v.1.2.0");

                document.Open();

                // Title
                var titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 14, BaseColor.WHITE);
                var titleText = courierId.HasValue
                    ? $"Monthly Task Data for {monthName} {year} (Name: {courierInfo?.Name} - ID: {courierInfo?.CourierID})"
                    : $"Monthly Task Data for {monthName} {year} (All Riders)";

                var titleTable = new PdfPTable(1) { WidthPercentage = 100 };
                var titleCell = new PdfPCell(new Phrase(titleText, titleFont))
                {
                    BackgroundColor = new BaseColor(25, 135, 84),
                    Border = Rectangle.NO_BORDER,
                    Padding = 8,
                    HorizontalAlignment = Element.ALIGN_CENTER
                };
                titleTable.AddCell(titleCell);
                document.Add(titleTable);

                if (data.Any())
                {
                    // Create table with 7 columns
                    var table = new PdfPTable(7)
                    {
                        WidthPercentage = 100,
                        SpacingBefore = 10f,
                        SpacingAfter = 10f,
                        HorizontalAlignment = Element.ALIGN_CENTER
                    };

                    // Set column widths (adjusted for landscape)
                    float[] columnWidths = { 0.5f, 1f, 1.5f, 1f, 1.5f, 1.8f, 1f };
                    table.SetWidths(columnWidths);

                    // Header row
                    var headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);
                    var headerBackground = new BaseColor(209, 231, 221);

                    AddPdfCell(table, "#", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Courier ID", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Name", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "City", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Purchase ID", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Delivered Date & Time", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Distance KM", headerFont, headerBackground, Element.ALIGN_CENTER);

                    // Data rows
                    var dataFont = FontFactory.GetFont(FontFactory.HELVETICA, 8);
                    int count = 1;

                    foreach (var item in data)
                    {
                        AddPdfCell(table, count++.ToString(), dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, item.CourierID.ToString(), dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, item.Name, dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, item.City, dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, item.PurchaseID, dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, item.DeliveredDateTime.ToString("dd/MM/yyyy HH:mm:ss"), dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, item.DistanceKM.ToString("0.00"), dataFont, null, Element.ALIGN_CENTER);
                    }

                    document.Add(table);
                }
                else
                {
                    var noDataFont = FontFactory.GetFont(FontFactory.HELVETICA, 12);
                    var noDataParagraph = new Paragraph("No task data found for the selected period.", noDataFont)
                    {
                        Alignment = Element.ALIGN_CENTER,
                        SpacingBefore = 20f
                    };
                    document.Add(noDataParagraph);
                }

                document.Close();
                var fileName = courierId.HasValue
                    ? $"TaskData_{courierInfo?.Name}_{courierId}_{monthName}_{year}.pdf"
                    : $"TaskData_AllRiders_{monthName}_{year}.pdf";

                return File(memoryStream.ToArray(), "application/pdf", fileName);
            }
        }

        public ActionResult ExportMonthlyTimeStampsToExcel(int month, int year, int? courierId = null)
        {
            var data = uploadFilesDAL.GetTimeStampsOfRider(month, year, courierId);
            var courierInfo = data.FirstOrDefault();
            var monthName = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month);

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Monthly Time Stamps");

                // Headers
                worksheet.Cell(1, 1).Value = "#";
                worksheet.Cell(1, 2).Value = "Courier ID";
                worksheet.Cell(1, 3).Value = "Name";
                worksheet.Cell(1, 4).Value = "City";
                worksheet.Cell(1, 5).Value = "Start Date";
                worksheet.Cell(1, 6).Value = "Start Time";
                worksheet.Cell(1, 7).Value = "End Time";

                int row = 2;
                int count = 1;

                foreach (var item in data)
                {
                    worksheet.Cell(row, 1).Value = count++;
                    worksheet.Cell(row, 2).Value = item.CourierID;
                    worksheet.Cell(row, 3).Value = item.Name;
                    worksheet.Cell(row, 4).Value = item.City;
                    worksheet.Cell(row, 5).Value = item.StartDate.ToString("dd/MM/yyyy");
                    worksheet.Cell(row, 6).Value = item.StartTime.ToString("dd/MM/yyyy HH:mm:ss");
                    worksheet.Cell(row, 7).Value = item.EndTime.ToString("dd/MM/yyyy HH:mm:ss");

                    row++;
                }

                // Formatting
                worksheet.Range("A1:G1").Style.Font.Bold = true; // Bold headers
                worksheet.Columns().AdjustToContents(); // Auto-size columns
                worksheet.SheetView.Freeze(1, 0); // Freeze header row

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var fileName = courierId.HasValue
                        ? $"TimeStamps_{courierInfo?.Name}_{courierId}_{monthName}_{year}.xlsx"
                        : $"TimeStamps_AllRiders_{monthName}_{year}.xlsx";

                    return File(stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        fileName);
                }
            }
        }

        public ActionResult ExportMonthlyTimeStampsToPdf(int month, int year, int? courierId = null)
        {
            // Get data using your existing method
            var data = uploadFilesDAL.GetTimeStampsOfRider(month, year, courierId);
            var courierInfo = data.FirstOrDefault();
            var monthName = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month);

            using (var memoryStream = new MemoryStream())
            {
                // Document setup (A4 landscape for better fit)
                var document = new Document(PageSize.A4.Rotate(), 15f, 15f, 15f, 30f);
                var writer = PdfWriter.GetInstance(document, memoryStream);
                writer.PageEvent = new PdfFooter("TrackPay Application");

                document.Open();

                // Title
                var titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 14, BaseColor.WHITE);
                var titleText = courierId.HasValue
                    ? $"Monthly Time Stamps for {monthName} {year} (Name: {courierInfo?.Name} - ID: {courierInfo?.CourierID})"
                    : $"Monthly Time Stamps for {monthName} {year} (All Riders)";

                var titleTable = new PdfPTable(1) { WidthPercentage = 100 };
                var titleCell = new PdfPCell(new Phrase(titleText, titleFont))
                {
                    BackgroundColor = new BaseColor(25, 135, 84),
                    Border = Rectangle.NO_BORDER,
                    Padding = 8,
                    HorizontalAlignment = Element.ALIGN_CENTER
                };
                titleTable.AddCell(titleCell);
                document.Add(titleTable);

                if (data.Any())
                {
                    // Create table with 7 columns
                    var table = new PdfPTable(7)
                    {
                        WidthPercentage = 100,
                        SpacingBefore = 10f,
                        SpacingAfter = 10f,
                        HorizontalAlignment = Element.ALIGN_CENTER
                    };

                    // Set column widths (adjusted for landscape)
                    float[] columnWidths = { 0.5f, 1f, 1.5f, 1f, 1f, 1.5f, 1.5f };
                    table.SetWidths(columnWidths);

                    // Header row
                    var headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);
                    var headerBackground = new BaseColor(209, 231, 221);

                    AddPdfCell(table, "#", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Courier ID", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Name", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "City", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Start Date", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Start Time", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "End Time", headerFont, headerBackground, Element.ALIGN_CENTER);

                    // Data rows
                    var dataFont = FontFactory.GetFont(FontFactory.HELVETICA, 8);
                    int count = 1;

                    foreach (var item in data)
                    {
                        AddPdfCell(table, count++.ToString(), dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, item.CourierID.ToString(), dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, item.Name, dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, item.City, dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, item.StartDate.ToString("dd/MM/yyyy"), dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, item.StartTime.ToString("dd/MM/yyyy HH:mm:ss"), dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, item.EndTime.ToString("dd/MM/yyyy HH:mm:ss"), dataFont, null, Element.ALIGN_CENTER);
                    }

                    document.Add(table);
                }
                else
                {
                    var noDataFont = FontFactory.GetFont(FontFactory.HELVETICA, 12);
                    var noDataParagraph = new Paragraph("No time stamps data found for the selected period.", noDataFont)
                    {
                        Alignment = Element.ALIGN_CENTER,
                        SpacingBefore = 20f
                    };
                    document.Add(noDataParagraph);
                }

                document.Close();
                var fileName = courierId.HasValue
                    ? $"TimeStamps_{courierInfo?.Name}_{courierId}_{monthName}_{year}.pdf"
                    : $"TimeStamps_AllRiders_{monthName}_{year}.pdf";

                return File(memoryStream.ToArray(), "application/pdf", fileName);
            }
        }

        // Helper method for adding PDF cells
        private void AddPdfCell(PdfPTable table, string text, Font font, BaseColor backgroundColor, int alignment)
        {
            var cell = new PdfPCell(new Phrase(text, font))
            {
                BackgroundColor = backgroundColor,
                HorizontalAlignment = alignment,
                Padding = 5
            };
            table.AddCell(cell);
        }
    }
}
