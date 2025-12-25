using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Newtonsoft.Json;
using Rotativa;
using Rotativa.AspNetCore;
using System.Diagnostics;
using System.Globalization;
using TrackPay.Data;
using TrackPay.Models;

namespace TrackPay.Controllers
{
    public class DeliveryController : Controller
    {
        private readonly DeliveryDAL deliveryDAL;

        public DeliveryController(IConfiguration configuration)
        {
            deliveryDAL = new DeliveryDAL(configuration);
        }

        [HttpGet]
        public IActionResult Index()
        {
            var model = new CombinedSummaryViewModel();
            // No need to populate months/years in ViewBag anymore
            return View(model);
        }

        [HttpPost]
        public IActionResult Index(CombinedSummaryViewModel model)
        {
            // Convert SelectedMonthYear to Month/Year if provided
            if (model.SelectedMonthYear.HasValue)
            {
                model.Month = model.SelectedMonthYear.Value.Month;
                model.Year = model.SelectedMonthYear.Value.Year;
            }

            var fullData = deliveryDAL.GetCompleteHourWithDeliveriesCalc();

            // For Monthly Reports
            if (model.Month > 0 && model.Year > 0)
            {
                model.SelectedMonthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(model.Month);

                if (model.CourierID > 0)
                {
                    // Daily Summary
                    model.DailySummaries = deliveryDAL.GetDailyCourierSummaries(model.Month, model.Year, model.CourierID);
                    model.ShowDailySummary = true;

                    var selectedCourier = fullData.FirstOrDefault(x => x.CourierID == model.CourierID);
                    model.SelectedCourierNameForDaily = selectedCourier?.Name ?? $"ID: {model.CourierID}";
                }
                else
                {
                    // Monthly Summary
                    model.MonthlySummaries = deliveryDAL.GetMonthlyCourierSummaries(model.Month, model.Year);
                    model.ShowMonthlySummary = true;
                }
            }

            // For Monthly Hourly Report
            if (model.SelectedMonthYearForHourly.HasValue && model.CourierIdMonthlyHourly > 0)
            {
                model.MonthlyHourlyData = deliveryDAL.GetMonthlyHourlyDataAllCalc(
                    model.SelectedMonthYearForHourly.Value.Year,
                    model.SelectedMonthYearForHourly.Value.Month,
                    model.CourierIdMonthlyHourly);

                model.SelectedMonthNameHourly = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(model.SelectedMonthYearForHourly.Value.Month);
                model.YearHourly = model.SelectedMonthYearForHourly.Value.Year;
                model.ShowMonthlyHourlySummary = true;

                // Updated line to avoid blank names
                model.SelectedCourierNameMonthlyHourly = fullData
                    .Where(x => x.CourierID == model.CourierIdMonthlyHourly)
                    .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x.Name))?.Name ?? "";
            }


            // For Hourly Report
            if (model.SelectedDate.HasValue && model.CourierIdHourly.HasValue)
            {
                model.HourlyData = deliveryDAL.GetHourlyDataAllCalc(model.SelectedDate.Value, model.CourierIdHourly.Value);
                model.ShowHourlySummary = true;

                model.SelectedCourierName = fullData
                    .Where(x => x.CourierID == model.CourierIdHourly)
                    .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x.Name))?.Name ?? $"ID: {model.CourierIdHourly}";
            }

            return View(model);
        }

        private void PopulateViewBags(int? month = null, int? year = null)
        {
            var fullData = deliveryDAL.GetCompleteHourWithDeliveriesCalc();

            // Months dropdown (always 1-12)
            ViewBag.Months = Enumerable.Range(1, 12)
                .Select(m => new SelectListItem
                {
                    Text = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(m),
                    Value = m.ToString(),
                    Selected = m == month
                })
                .ToList();

            // Years dropdown (only years with data)
            ViewBag.Years = fullData
                .Select(x => x.StartDate.Year)
                .Distinct()
                .OrderByDescending(y => y)
                .Select(y => new SelectListItem
                {
                    Text = y.ToString(),
                    Value = y.ToString(),
                    Selected = y == year
                })
                .ToList();
        }

        public ActionResult ExportMonthlySummaryToExcel(int month, int year)
        {
            var data = deliveryDAL.GetMonthlyCourierSummaries(month, year);

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Monthly Summary");

                // Headers
                worksheet.Cell(1, 1).Value = "#";
                worksheet.Cell(1, 2).Value = "Courier ID";
                worksheet.Cell(1, 3).Value = "Name";
                worksheet.Cell(1, 4).Value = "Total Order Pay";
                worksheet.Cell(1, 5).Value = "Total Distance Pay";
                worksheet.Cell(1, 6).Value = "Total Pay";

                int row = 2;
                int count = 1;
                foreach (var item in data.OrderBy(x => x.CourierID))
                {
                    worksheet.Cell(row, 1).Value = count++;
                    worksheet.Cell(row, 2).Value = item.CourierID;
                    worksheet.Cell(row, 3).Value = item.Name;

                    worksheet.Cell(row, 4).Value = Math.Round(item.TotalOrderPay, 2);
                    worksheet.Cell(row, 5).Value = Math.Round(item.TotalDistancePay, 2);
                    worksheet.Cell(row, 6).Value = Math.Round(item.TotalPay, 2);

                    // Format these cells with number format and " kr" suffix
                    worksheet.Cell(row, 4).Style.NumberFormat.Format = "#,##0.00 \"kr\"";
                    worksheet.Cell(row, 5).Style.NumberFormat.Format = "#,##0.00 \"kr\"";
                    worksheet.Cell(row, 6).Style.NumberFormat.Format = "#,##0.00 \"kr\"";

                    row++;
                }

                // Optional: bold headers
                worksheet.Range("A1:F1").Style.Font.Bold = true;

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return File(stream.ToArray(),
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                $"MonthlySummary_{month}_{year}.xlsx");
                }
            }
        }

        public ActionResult ExportMonthlySummaryToPdf(int month, int year)
        {
            var data = deliveryDAL.GetMonthlyCourierSummaries(month, year).OrderBy(x => x.CourierID).ToList();
            var grandTotalPay = data.Sum(x => x.TotalPay);
            var selectedMonthName = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month);

            using (var memoryStream = new MemoryStream())
            {
                // Document setup (A4 landscape)
                var document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 20f, 20f, 20f, 40f);
                var writer = PdfWriter.GetInstance(document, memoryStream);
                writer.PageEvent = new PdfFooter("TrackPay Application (v.1.2.0)");

                document.Open();

                // Title with green background
                var titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 16, BaseColor.WHITE);
                var titleParagraph = new iTextSharp.text.Paragraph($"Monthly Summary - {selectedMonthName} {year}", titleFont)
                {
                    Alignment = Element.ALIGN_CENTER,
                    SpacingAfter = 20f
                };

                var titleTable = new PdfPTable(1)
                {
                    WidthPercentage = 100,
                    HorizontalAlignment = Element.ALIGN_CENTER
                };

                var titleCell = new PdfPCell(titleParagraph)
                {
                    BackgroundColor = new BaseColor(25, 135, 84),
                    Border = iTextSharp.text.Rectangle.NO_BORDER,
                    Padding = 10,
                    HorizontalAlignment = Element.ALIGN_CENTER
                };
                titleTable.AddCell(titleCell);

                document.Add(titleTable);

                if (data.Any())
                {
                    // Create table with 6 columns
                    var table = new PdfPTable(6)
                    {
                        WidthPercentage = 100,
                        SpacingBefore = 10f,
                        SpacingAfter = 10f
                    };

                    // Set column widths
                    float[] columnWidths = { 0.5f, 1.5f, 3f, 2f, 2f, 2f };
                    table.SetWidths(columnWidths);

                    // Header row
                    var headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10);
                    var headerBackground = new BaseColor(209, 231, 221);

                    AddPdfCell(table, "#", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Courier ID", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Name", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Total Order Pay", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Total Distance Pay", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Total Pay", headerFont, headerBackground, Element.ALIGN_CENTER);

                    // Data rows
                    var dataFont = FontFactory.GetFont(FontFactory.HELVETICA, 10);
                    var boldFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10);

                    for (int i = 0; i < data.Count; i++)
                    {
                        var item = data[i];

                        AddPdfCell(table, (i + 1).ToString(), dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, item.CourierID.ToString(), dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, item.Name, dataFont, null, Element.ALIGN_LEFT);
                        AddPdfCell(table, item.TotalOrderPay.ToString("N2") + " kr", dataFont, null, Element.ALIGN_RIGHT);
                        AddPdfCell(table, item.TotalDistancePay.ToString("N2") + " kr", dataFont, null, Element.ALIGN_RIGHT);
                        AddPdfCell(table, item.TotalPay.ToString("N2") + " kr", boldFont, null, Element.ALIGN_RIGHT);
                    }

                    // Footer row
                    var footerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10);
                    var footerBackground = new BaseColor(209, 231, 221);

                    AddPdfCell(table, "GRAND TOTAL PAY", footerFont, footerBackground, Element.ALIGN_CENTER, 5);
                    AddPdfCell(table, grandTotalPay.ToString("N2") + " kr", footerFont, footerBackground, Element.ALIGN_RIGHT);

                    document.Add(table);
                }
                else
                {
                    var noDataFont = FontFactory.GetFont(FontFactory.HELVETICA, 12);
                    var noDataParagraph = new iTextSharp.text.Paragraph("No monthly summary data found for the selected criteria.", noDataFont)
                    {
                        Alignment = Element.ALIGN_CENTER,
                        SpacingBefore = 20f
                    };
                    document.Add(noDataParagraph);
                }

                if (writer.PageNumber == 1)
                {
                    // Just add a small empty paragraph to ensure footer appears
                    document.Add(new iTextSharp.text.Paragraph(" "));
                    document.Add(Chunk.NEWLINE);
                }

                document.Close();
                return File(memoryStream.ToArray(), "application/pdf", $"MonthlySummary_{month}_{year}.pdf");
            }
        }

        public ActionResult ExportDailySummaryToExcel(int month, int year, int courierId)
        {
            var data = deliveryDAL.GetDailyCourierSummaries(month, year, courierId);
            var selectedMonthName = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month);

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Daily Summary");

                // Headers
                worksheet.Cell(1, 1).Value = "Date";
                worksheet.Cell(1, 2).Value = "Courier ID";
                worksheet.Cell(1, 3).Value = "Total Order Pay";
                worksheet.Cell(1, 4).Value = "Total Distance Pay";
                worksheet.Cell(1, 5).Value = "Total Pay";

                int row = 2;
                foreach (var item in data)
                {
                    worksheet.Cell(row, 1).Value = item.StartDate.ToString("dd-MM-yyyy");
                    worksheet.Cell(row, 2).Value = item.CourierID;
                    worksheet.Cell(row, 3).Value = item.TotalOrderPay;
                    worksheet.Cell(row, 4).Value = item.TotalDistancePay;
                    worksheet.Cell(row, 5).Value = item.TotalPay;

                    // Apply number formatting
                    worksheet.Cell(row, 3).Style.NumberFormat.Format = "#,##0 \"kr\"";
                    worksheet.Cell(row, 4).Style.NumberFormat.Format = "#,##0.00 \"kr\"";
                    worksheet.Cell(row, 5).Style.NumberFormat.Format = "#,##0.00 \"kr\"";

                    row++;
                }

                // Add totals row
                var totalPay = data.Sum(x => x.TotalPay);
                worksheet.Cell(row, 1).Value = "TOTAL";
                worksheet.Range(row, 1, row, 4).Merge().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(row, 5).Value = totalPay;
                worksheet.Cell(row, 5).Style.NumberFormat.Format = "#,##0.00 \"kr\"";

                // Formatting
                worksheet.Range("A1:E1").Style.Font.Bold = true; // Header row
                worksheet.Range($"A{row}:E{row}").Style.Font.Bold = true; // Total row
                worksheet.Columns().AdjustToContents();

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return File(stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        $"DailySummary_{courierId}_{month}_{year}.xlsx");
                }
            }
        }

        public ActionResult ExportDailySummaryToPdf(int month, int year, int courierId, string courierName)
        {
            var data = deliveryDAL.GetDailyCourierSummaries(month, year, courierId);
            var selectedMonthName = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month);
            var totalPay = data.Sum(x => x.TotalPay);

            using (var memoryStream = new MemoryStream())
            {
                var document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 20f, 20f, 20f, 35f);
                var writer = PdfWriter.GetInstance(document, memoryStream);
                writer.PageEvent = new PdfFooter("TrackPay Application (v.1.2.0)");

                document.Open();

                // Title with green background
                var titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 14, BaseColor.WHITE);
                var titleText = $"Daily Summary - {selectedMonthName} {year}";
                if (!string.IsNullOrEmpty(courierName))
                {
                    titleText += $" (Name: {courierName} - ID: {courierId})";
                }
                else
                {
                    titleText += $" (ID: {courierId})";
                }

                var titleTable = new PdfPTable(1) { WidthPercentage = 100 };
                var titleCell = new PdfPCell(new Phrase(titleText, titleFont))
                {
                    BackgroundColor = new BaseColor(25, 135, 84),
                    Border = iTextSharp.text.Rectangle.NO_BORDER,
                    Padding = 8,
                    HorizontalAlignment = Element.ALIGN_CENTER
                };
                titleTable.AddCell(titleCell);
                document.Add(titleTable);

                if (data.Any())
                {
                    // Create table with 5 columns
                    var table = new PdfPTable(5)
                    {
                        WidthPercentage = 100,
                        SpacingBefore = 10f,
                        SpacingAfter = 10f
                    };

                    // Set column widths
                    float[] columnWidths = { 1f, 1f, 1f, 1f, 1f };
                    table.SetWidths(columnWidths);

                    // Header row
                    var headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);
                    var headerBackground = new BaseColor(209, 231, 221);

                    AddPdfCell(table, "Date", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Courier ID", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Total Order Pay", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Total Distance Pay", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Total Pay", headerFont, headerBackground, Element.ALIGN_CENTER);

                    // Data rows
                    var dataFont = FontFactory.GetFont(FontFactory.HELVETICA, 8);
                    var boldFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);

                    foreach (var item in data)
                    {
                        AddPdfCell(table, item.StartDate.ToString("dd-MM-yyyy"), dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, item.CourierID.ToString(), dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, item.TotalOrderPay.ToString("N0") + " kr", dataFont, null, Element.ALIGN_RIGHT);
                        AddPdfCell(table, item.TotalDistancePay.ToString("N2") + " kr", dataFont, null, Element.ALIGN_RIGHT);
                        AddPdfCell(table, item.TotalPay.ToString("N2") + " kr", boldFont, null, Element.ALIGN_RIGHT);
                    }

                    // Footer row with totals
                    var footerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);
                    var footerBackground = new BaseColor(209, 231, 221);

                    AddPdfCell(table, "TOTAL", footerFont, footerBackground, Element.ALIGN_CENTER, 4);
                    AddPdfCell(table, totalPay.ToString("N2") + " kr", footerFont, footerBackground, Element.ALIGN_RIGHT);

                    document.Add(table);
                }
                else
                {
                    var noDataFont = FontFactory.GetFont(FontFactory.HELVETICA, 12);
                    var noDataParagraph = new iTextSharp.text.Paragraph("No daily summary data found for the selected courier.", noDataFont)
                    {
                        Alignment = Element.ALIGN_CENTER,
                        SpacingBefore = 20f
                    };
                    document.Add(noDataParagraph);
                }

                document.Close();
                return File(memoryStream.ToArray(), "application/pdf", $"DailySummary_{courierId}_{month}_{year}.pdf");
            }
        }

        public ActionResult ExportMonthlyHourlySummaryToExcel(int month, int year, int courierId)
        {
            var data = deliveryDAL.GetMonthlyHourlyDataAllCalc(year, month, courierId);

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Monthly Hourly Summary");

                // Headers
                worksheet.Cell(1, 1).Value = "#";
                worksheet.Cell(1, 2).Value = "Date";
                worksheet.Cell(1, 3).Value = "Time Period";
                worksheet.Cell(1, 4).Value = "Complete Hours";
                worksheet.Cell(1, 5).Value = "Time Duration";
                worksheet.Cell(1, 6).Value = "Orders";
                worksheet.Cell(1, 7).Value = "Distance (KM)";
                worksheet.Cell(1, 8).Value = "Hourly Pay";
                worksheet.Cell(1, 9).Value = "Order Pay";
                worksheet.Cell(1, 10).Value = "Distance Pay";
                worksheet.Cell(1, 11).Value = "Total Pay";

                int row = 2;
                int count = 1;

                foreach (var item in data.OrderBy(x => x.StartDate)
                                       .ThenBy(x => x.StartTime == DateTime.MinValue ? 1 : 0)
                                       .ThenBy(x => x.StartTime))
                {
                    worksheet.Cell(row, 1).Value = count++;
                    worksheet.Cell(row, 2).Value = item.StartDate.ToString("dd-MM-yyyy");

                    worksheet.Cell(row, 3).Value = (item.StartTime != DateTime.MinValue && item.EndTime != DateTime.MinValue)
                        ? $"{item.StartTime:HH:mm:ss} - {item.EndTime:HH:mm:ss}"
                        : "No Complete/Partial Hours";

                    worksheet.Cell(row, 4).Value = item.IsCompleteHour ? "Yes" : "No";
                    worksheet.Cell(row, 5).Value = item.TimeDuration ?? "-";
                    worksheet.Cell(row, 6).Value = item.OrderDelivered;
                    worksheet.Cell(row, 7).Value = item.DistanceKM ?? 0;
                    worksheet.Cell(row, 8).Value = item.HourlyPay ?? 0;
                    worksheet.Cell(row, 9).Value = item.OrderPay ?? 0;
                    worksheet.Cell(row, 10).Value = item.DistancePay ?? 0;
                    worksheet.Cell(row, 11).Value = item.TotalPay ?? 0;

                    // Apply number formatting
                    worksheet.Cell(row, 7).Style.NumberFormat.Format = "#,##0.00";
                    worksheet.Cell(row, 8).Style.NumberFormat.Format = "#,##0.00 \"kr\"";
                    worksheet.Cell(row, 9).Style.NumberFormat.Format = "#,##0 \"kr\"";
                    worksheet.Cell(row, 10).Style.NumberFormat.Format = "#,##0.00 \"kr\"";
                    worksheet.Cell(row, 11).Style.NumberFormat.Format = "#,##0.00 \"kr\"";

                    row++;
                }

                // Calculate total time duration
                TimeSpan totalTime = new TimeSpan();
                foreach (var item in data)
                {
                    if (item.TimeDuration != null && TimeSpan.TryParse(item.TimeDuration, out TimeSpan time))
                    {
                        totalTime = totalTime.Add(time);
                    }
                }
                string totalTimeFormatted = $"{(int)totalTime.TotalHours}:{totalTime.Minutes:00}:{totalTime.Seconds:00}";

                // Add totals row (matching your view's footer)
                worksheet.Cell(row, 1).Value = "TOTAL";
                worksheet.Range(row, 1, row, 4).Merge().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(row, 5).Value = totalTimeFormatted;
                worksheet.Cell(row, 6).Value = data.Sum(x => x.OrderDelivered);
                worksheet.Cell(row, 7).Value = data.Sum(x => x.DistanceKM ?? 0);
                worksheet.Cell(row, 8).Value = data.Sum(x => x.HourlyPay ?? 0);
                worksheet.Cell(row, 9).Value = data.Sum(x => x.OrderPay ?? 0);
                worksheet.Cell(row, 10).Value = data.Sum(x => x.DistancePay ?? 0);
                worksheet.Cell(row, 11).Value = data.Sum(x => x.TotalPay ?? 0);

                // Format totals row
                worksheet.Range($"A{row}:K{row}").Style.Font.Bold = true;
                worksheet.Cell(row, 7).Style.NumberFormat.Format = "#,##0.00";
                worksheet.Cell(row, 8).Style.NumberFormat.Format = "#,##0.00 \"kr\"";
                worksheet.Cell(row, 9).Style.NumberFormat.Format = "#,##0 \"kr\"";
                worksheet.Cell(row, 10).Style.NumberFormat.Format = "#,##0.00 \"kr\"";
                worksheet.Cell(row, 11).Style.NumberFormat.Format = "#,##0.00 \"kr\"";

                // Bold headers and auto-size columns
                worksheet.Range("A1:K1").Style.Font.Bold = true;
                worksheet.Columns().AdjustToContents();

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return File(stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        $"MonthlyHourlySummary_{courierId}_{month}_{year}.xlsx");
                }
            }
        }

        public ActionResult ExportMonthlyHourlySummaryToPdf(int month, int year, int courierId, string courierName)
        {
            var data = deliveryDAL.GetMonthlyHourlyDataAllCalc(year, month, courierId);
            var selectedMonthName = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month);
            var totalPay = data.Sum(x => x.TotalPay ?? 0);

            // Calculate total time duration
            TimeSpan totalTime = new TimeSpan();
            foreach (var item in data)
            {
                if (item.TimeDuration != null && TimeSpan.TryParse(item.TimeDuration, out TimeSpan time))
                {
                    totalTime = totalTime.Add(time);
                }
            }
            string totalTimeFormatted = $"{(int)totalTime.TotalHours}:{totalTime.Minutes:00}:{totalTime.Seconds:00}";

            using (var memoryStream = new MemoryStream())
            {
                // Set to landscape orientation
                var pageSize = iTextSharp.text.PageSize.A4.Rotate();
                var document = new iTextSharp.text.Document(pageSize, 20f, 20f, 20f, 35f);
                //var document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 20f, 20f, 20f, 35f);
                var writer = PdfWriter.GetInstance(document, memoryStream);
                writer.PageEvent = new PdfFooter("TrackPay Application (v.1.2.0)");

                document.Open();

                // Title with green background
                var titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 14, BaseColor.WHITE);
                var titleText = $"Monthly Hourly Summary - {selectedMonthName} {year}";
                if (!string.IsNullOrEmpty(courierName))
                {
                    titleText += $" (Name: {courierName} - ID: {courierId})";
                }
                else
                {
                    titleText += $" (ID: {courierId})";
                }

                var titleTable = new PdfPTable(1) { WidthPercentage = 100 };
                var titleCell = new PdfPCell(new Phrase(titleText, titleFont))
                {
                    BackgroundColor = new BaseColor(25, 135, 84),
                    Border = iTextSharp.text.Rectangle.NO_BORDER,
                    Padding = 8,
                    HorizontalAlignment = Element.ALIGN_CENTER
                };
                titleTable.AddCell(titleCell);
                document.Add(titleTable);

                if (data.Any())
                {
                    // Create table with 11 columns for landscape
                    var table = new PdfPTable(11)
                    {
                        WidthPercentage = 100,
                        SpacingBefore = 10f,
                        SpacingAfter = 10f
                    };

                    // Set column widths (adjusted for landscape)
                    float[] columnWidths = { 0.5f, 1.2f, 1.8f, 0.8f, 1f, 0.8f, 0.9f, 0.9f, 0.9f, 0.9f, 1f };
                    table.SetWidths(columnWidths);

                    // Header row
                    var headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);
                    var headerBackground = new BaseColor(209, 231, 221);

                    AddPdfCell(table, "#", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Date", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Time Period", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Complete Hours", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Time Duration", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Orders", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Distance (KM)", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Hourly Pay", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Order Pay", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Distance Pay", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Total Pay", headerFont, headerBackground, Element.ALIGN_CENTER);

                    // Data rows
                    var dataFont = FontFactory.GetFont(FontFactory.HELVETICA, 8);
                    var boldFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);

                    for (int i = 0; i < data.Count; i++)
                    {
                        var item = data[i];

                        AddPdfCell(table, (i + 1).ToString(), dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, item.StartDate.ToString("dd-MM-yyyy"), dataFont, null, Element.ALIGN_CENTER);

                        // Time period
                        if (item.StartTime != DateTime.MinValue && item.EndTime != DateTime.MinValue)
                        {
                            AddPdfCell(table, $"{item.StartTime:HH:mm:ss} - {item.EndTime:HH:mm:ss}", dataFont, null, Element.ALIGN_CENTER);
                        }
                        else
                        {
                            AddPdfCell(table, "No Complete/Partial Hours", dataFont, null, Element.ALIGN_CENTER);
                        }

                        // Complete hour indicator
                        var completeCell = new PdfPCell(new Phrase(item.IsCompleteHour ? "Yes" : "No", dataFont))
                        {
                            HorizontalAlignment = Element.ALIGN_CENTER,
                            BackgroundColor = item.IsCompleteHour ? new BaseColor(220, 255, 220) : new BaseColor(255, 220, 220)
                        };
                        table.AddCell(completeCell);

                        // Time Duration
                        AddPdfCell(table, item.TimeDuration ?? "-", dataFont, null, Element.ALIGN_CENTER);

                        // Numeric data
                        AddPdfCell(table, item.OrderDelivered.ToString(), dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, (item.DistanceKM?.ToString("F2") ?? "-"), dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, (item.HourlyPay?.ToString("F2") ?? "-") + " kr", dataFont, null, Element.ALIGN_RIGHT);
                        AddPdfCell(table, (item.OrderPay?.ToString("F0") ?? "-") + " kr", dataFont, null, Element.ALIGN_RIGHT);
                        AddPdfCell(table, (item.DistancePay?.ToString("F2") ?? "-") + " kr", dataFont, null, Element.ALIGN_RIGHT);
                        AddPdfCell(table, (item.TotalPay?.ToString("F2") ?? "-") + " kr", boldFont, null, Element.ALIGN_RIGHT);
                    }

                    // Footer row with totals
                    var footerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);
                    var footerBackground = new BaseColor(209, 231, 221);

                    AddPdfCell(table, "TOTAL", footerFont, footerBackground, Element.ALIGN_CENTER, 4);
                    AddPdfCell(table, totalTimeFormatted, footerFont, footerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, data.Sum(x => x.OrderDelivered).ToString(), footerFont, footerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, data.Sum(x => x.DistanceKM ?? 0).ToString("F2") + " KM", footerFont, footerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, data.Sum(x => x.HourlyPay ?? 0).ToString("F2") + " kr", footerFont, footerBackground, Element.ALIGN_RIGHT);
                    AddPdfCell(table, data.Sum(x => x.OrderPay ?? 0).ToString("F0") + " kr", footerFont, footerBackground, Element.ALIGN_RIGHT);
                    AddPdfCell(table, data.Sum(x => x.DistancePay ?? 0).ToString("F2") + " kr", footerFont, footerBackground, Element.ALIGN_RIGHT);
                    AddPdfCell(table, totalPay.ToString("F2") + " kr", footerFont, footerBackground, Element.ALIGN_RIGHT);

                    document.Add(table);
                }
                else
                {
                    var noDataFont = FontFactory.GetFont(FontFactory.HELVETICA, 12);
                    var noDataParagraph = new iTextSharp.text.Paragraph("No monthly hourly data found for the selected courier.", noDataFont)
                    {
                        Alignment = Element.ALIGN_CENTER,
                        SpacingBefore = 20f
                    };
                    document.Add(noDataParagraph);
                }

                document.Close();
                return File(memoryStream.ToArray(), "application/pdf", $"MonthlyHourlySummary_{courierId}_{month}_{year}.pdf");
            }
        }

        public ActionResult ExportHourlySummaryToExcel(string date, int courierId)
        {
            var selectedDate = DateTime.Parse(date);
            var data = deliveryDAL.GetHourlyDataAllCalc(selectedDate, courierId);

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Hourly Summary");

                // Headers
                worksheet.Cell(1, 1).Value = "#";
                worksheet.Cell(1, 2).Value = "Time Period";
                worksheet.Cell(1, 3).Value = "Complete Hours";
                worksheet.Cell(1, 4).Value = "Time Duration";
                worksheet.Cell(1, 5).Value = "Orders";
                worksheet.Cell(1, 6).Value = "Distance (KM)";
                worksheet.Cell(1, 7).Value = "Hourly Pay";
                worksheet.Cell(1, 8).Value = "Order Pay";
                worksheet.Cell(1, 9).Value = "Distance Pay";
                worksheet.Cell(1, 10).Value = "Total Pay";

                int row = 2;
                int count = 1;

                foreach (var item in data.OrderBy(x => x.StartTime == DateTime.MinValue ? 1 : 0)
                                       .ThenBy(x => x.StartTime))
                {
                    worksheet.Cell(row, 1).Value = count++;

                    worksheet.Cell(row, 2).Value = (item.StartTime != DateTime.MinValue && item.EndTime != DateTime.MinValue)
                        ? $"{item.StartTime:HH:mm:ss} - {item.EndTime:HH:mm:ss}"
                        : "No Complete/Partial Hours";

                    worksheet.Cell(row, 3).Value = item.IsCompleteHour ? "Yes" : "No";
                    worksheet.Cell(row, 4).Value = item.TimeDuration ?? "-";
                    worksheet.Cell(row, 5).Value = item.OrderDelivered;
                    worksheet.Cell(row, 6).Value = item.DistanceKM ?? 0;
                    worksheet.Cell(row, 7).Value = item.HourlyPay ?? 0;
                    worksheet.Cell(row, 8).Value = item.OrderPay ?? 0;
                    worksheet.Cell(row, 9).Value = item.DistancePay ?? 0;
                    worksheet.Cell(row, 10).Value = item.TotalPay ?? 0;

                    // Apply number formatting
                    worksheet.Cell(row, 6).Style.NumberFormat.Format = "#,##0.00";
                    worksheet.Cell(row, 7).Style.NumberFormat.Format = "#,##0.00 \"kr\"";
                    worksheet.Cell(row, 8).Style.NumberFormat.Format = "#,##0 \"kr\"";
                    worksheet.Cell(row, 9).Style.NumberFormat.Format = "#,##0.00 \"kr\"";
                    worksheet.Cell(row, 10).Style.NumberFormat.Format = "#,##0.00 \"kr\"";

                    row++;
                }

                // Calculate total time duration
                TimeSpan totalTime = new TimeSpan();
                foreach (var item in data)
                {
                    if (TimeSpan.TryParse(item.TimeDuration, out TimeSpan time))
                    {
                        totalTime = totalTime.Add(time);
                    }
                }
                string totalTimeFormatted = $"{(int)totalTime.TotalHours}:{totalTime.Minutes:00}:{totalTime.Seconds:00}";

                // Add totals row
                worksheet.Cell(row, 1).Value = "TOTAL";
                worksheet.Range(row, 1, row, 3).Merge().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(row, 4).Value = totalTimeFormatted;
                worksheet.Cell(row, 5).Value = data.Sum(x => x.OrderDelivered);
                worksheet.Cell(row, 6).Value = data.Sum(x => x.DistanceKM ?? 0);
                worksheet.Cell(row, 7).Value = data.Sum(x => x.HourlyPay ?? 0);
                worksheet.Cell(row, 8).Value = data.Sum(x => x.OrderPay ?? 0);
                worksheet.Cell(row, 9).Value = data.Sum(x => x.DistancePay ?? 0);
                worksheet.Cell(row, 10).Value = data.Sum(x => x.TotalPay ?? 0);

                // Format totals row
                worksheet.Range($"A{row}:J{row}").Style.Font.Bold = true;
                worksheet.Cell(row, 6).Style.NumberFormat.Format = "#,##0.00";
                worksheet.Cell(row, 7).Style.NumberFormat.Format = "#,##0.00 \"kr\"";
                worksheet.Cell(row, 8).Style.NumberFormat.Format = "#,##0 \"kr\"";
                worksheet.Cell(row, 9).Style.NumberFormat.Format = "#,##0.00 \"kr\"";
                worksheet.Cell(row, 10).Style.NumberFormat.Format = "#,##0.00 \"kr\"";

                // Bold headers and auto-size columns
                worksheet.Range("A1:J1").Style.Font.Bold = true;
                worksheet.Columns().AdjustToContents();

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return File(stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        $"HourlySummary_{courierId}_{selectedDate:yyyyMMdd}.xlsx");
                }
            }
        }

        public ActionResult ExportHourlySummaryToPdf(string date, int courierId, string courierName)
        {
            var selectedDate = DateTime.Parse(date);
            var data = deliveryDAL.GetHourlyDataAllCalc(selectedDate, courierId);
            var totalPay = data.Sum(x => x.TotalPay ?? 0);

            // Calculate total time duration
            TimeSpan totalTime = new TimeSpan();
            foreach (var item in data)
            {
                if (TimeSpan.TryParse(item.TimeDuration, out TimeSpan time))
                {
                    totalTime = totalTime.Add(time);
                }
            }
            string totalTimeFormatted = $"{(int)totalTime.TotalHours}:{totalTime.Minutes:00}:{totalTime.Seconds:00}";

            using (var memoryStream = new MemoryStream())
            {
                // Change to landscape by rotating the page size
                var pageSize = iTextSharp.text.PageSize.A4.Rotate();
                var document = new iTextSharp.text.Document(pageSize, 20f, 20f, 20f, 35f);

                var writer = PdfWriter.GetInstance(document, memoryStream);
                writer.PageEvent = new PdfFooter("TrackPay Application (v.1.2.0)");

                document.Open();

                // Title with green background
                var titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 14, BaseColor.WHITE);
                var titleText = $"Hourly Summary - {selectedDate:dd MMM yyyy}";
                if (!string.IsNullOrEmpty(courierName))
                {
                    titleText += $" (Name: {courierName} - ID: {courierId})";
                }
                else
                {
                    titleText += $" (ID: {courierId})";
                }

                var titleTable = new PdfPTable(1) { WidthPercentage = 100 };
                var titleCell = new PdfPCell(new Phrase(titleText, titleFont))
                {
                    BackgroundColor = new BaseColor(25, 135, 84),
                    Border = iTextSharp.text.Rectangle.NO_BORDER,
                    Padding = 8,
                    HorizontalAlignment = Element.ALIGN_CENTER
                };
                titleTable.AddCell(titleCell);
                document.Add(titleTable);

                if (data.Any())
                {
                    // Create table with 10 columns
                    var table = new PdfPTable(10)
                    {
                        WidthPercentage = 100,
                        SpacingBefore = 10f,
                        SpacingAfter = 10f
                    };

                    // Set column widths
                    float[] columnWidths = { 0.5f, 1.8f, 1f, 1f, 0.8f, 0.9f, 0.9f, 0.9f, 0.9f, 1f };
                    table.SetWidths(columnWidths);

                    // Header row
                    var headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);
                    var headerBackground = new BaseColor(209, 231, 221);

                    AddPdfCell(table, "#", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Time Period", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Complete Hours", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Time Duration", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Orders", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Distance (KM)", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Hourly Pay", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Order Pay", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Distance Pay", headerFont, headerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, "Total Pay", headerFont, headerBackground, Element.ALIGN_CENTER);

                    // Data rows
                    var dataFont = FontFactory.GetFont(FontFactory.HELVETICA, 8);
                    var boldFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);

                    for (int i = 0; i < data.Count; i++)
                    {
                        var item = data[i];

                        AddPdfCell(table, (i + 1).ToString(), dataFont, null, Element.ALIGN_CENTER);

                        // Time period
                        if (item.StartTime != DateTime.MinValue && item.EndTime != DateTime.MinValue)
                        {
                            AddPdfCell(table, $"{item.StartTime:HH:mm:ss} - {item.EndTime:HH:mm:ss}", dataFont, null, Element.ALIGN_CENTER);
                        }
                        else
                        {
                            AddPdfCell(table, "No Complete/Partial Hours", dataFont, null, Element.ALIGN_CENTER);
                        }

                        // Complete hour indicator
                        var completeCell = new PdfPCell(new Phrase(item.IsCompleteHour ? "Yes" : "No", dataFont))
                        {
                            HorizontalAlignment = Element.ALIGN_CENTER,
                            BackgroundColor = item.IsCompleteHour ? new BaseColor(220, 255, 220) : new BaseColor(255, 220, 220)
                        };
                        table.AddCell(completeCell);

                        // Time Duration
                        AddPdfCell(table, item.TimeDuration ?? "-", dataFont, null, Element.ALIGN_CENTER);

                        // Numeric data
                        AddPdfCell(table, item.OrderDelivered.ToString(), dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, (item.DistanceKM?.ToString("F2") ?? "-"), dataFont, null, Element.ALIGN_CENTER);
                        AddPdfCell(table, (item.HourlyPay?.ToString("F2") ?? "-") + " kr", dataFont, null, Element.ALIGN_RIGHT);
                        AddPdfCell(table, (item.OrderPay?.ToString("F0") ?? "-") + " kr", dataFont, null, Element.ALIGN_RIGHT);
                        AddPdfCell(table, (item.DistancePay?.ToString("F2") ?? "-") + " kr", dataFont, null, Element.ALIGN_RIGHT);
                        AddPdfCell(table, (item.TotalPay?.ToString("F2") ?? "-") + " kr", boldFont, null, Element.ALIGN_RIGHT);
                    }

                    // Footer row with totals
                    var footerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);
                    var footerBackground = new BaseColor(209, 231, 221);

                    AddPdfCell(table, "TOTAL", footerFont, footerBackground, Element.ALIGN_CENTER, 3);
                    AddPdfCell(table, totalTimeFormatted, footerFont, footerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, data.Sum(x => x.OrderDelivered).ToString(), footerFont, footerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, data.Sum(x => x.DistanceKM ?? 0).ToString("F2") + " KM", footerFont, footerBackground, Element.ALIGN_CENTER);
                    AddPdfCell(table, data.Sum(x => x.HourlyPay ?? 0).ToString("F2") + " kr", footerFont, footerBackground, Element.ALIGN_RIGHT);
                    AddPdfCell(table, data.Sum(x => x.OrderPay ?? 0).ToString("F0") + " kr", footerFont, footerBackground, Element.ALIGN_RIGHT);
                    AddPdfCell(table, data.Sum(x => x.DistancePay ?? 0).ToString("F2") + " kr", footerFont, footerBackground, Element.ALIGN_RIGHT);
                    AddPdfCell(table, totalPay.ToString("F2") + " kr", footerFont, footerBackground, Element.ALIGN_RIGHT);

                    document.Add(table);
                }
                else
                {
                    var noDataFont = FontFactory.GetFont(FontFactory.HELVETICA, 12);
                    var noDataParagraph = new iTextSharp.text.Paragraph("No hourly data found for the selected criteria.", noDataFont)
                    {
                        Alignment = Element.ALIGN_CENTER,
                        SpacingBefore = 20f
                    };
                    document.Add(noDataParagraph);
                }

                document.Close();
                return File(memoryStream.ToArray(), "application/pdf", $"HourlySummary_{courierId}_{selectedDate:yyyyMMdd}.pdf");
            }
        }

        private void AddPdfCell(PdfPTable table, string text, iTextSharp.text.Font font, BaseColor backgroundColor, int horizontalAlignment, int colSpan = 1)
        {
            var cell = new PdfPCell(new Phrase(text, font))
            {
                HorizontalAlignment = horizontalAlignment,
                VerticalAlignment = Element.ALIGN_MIDDLE, // Add this for vertical centering
                BackgroundColor = backgroundColor,
                Colspan = colSpan,
                Padding = 5 // Add some padding for better appearance
            };
            table.AddCell(cell);
        }

        [HttpGet]
        public IActionResult UploadExcelFiles()
        {
            // Task Data
            if (TempData["TaskDataMessage"] != null)
                ViewBag.TaskDataMessage = TempData["TaskDataMessage"];

            if (TempData["TaskDataErrors"] != null)
                ViewBag.TaskDataErrors = JsonConvert.DeserializeObject<List<string>>(TempData["TaskDataErrors"].ToString());

            // Time Stamps
            if (TempData["TimeStampsMessage"] != null)
                ViewBag.TimeStampsMessage = TempData["TimeStampsMessage"];

            if (TempData["TimeStampsErrors"] != null)
                ViewBag.TimeStampsErrors = JsonConvert.DeserializeObject<List<string>>(TempData["TimeStampsErrors"].ToString());

            return View();
        }

        [HttpPost]
        public async Task<IActionResult> UploadTaskDataExcelFile(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                TempData["TaskDataMessage"] = "Please select a valid Excel file.";
                return RedirectToAction("UploadExcelFiles");
            }

            var result = await deliveryDAL.UploadTaskDataFromExcel(file);

            if (result.ErrorMessages.Any())
            {
                TempData["TaskDataMessage"] = "Some errors occurred during upload.";
                TempData["TaskDataErrors"] = JsonConvert.SerializeObject(result.ErrorMessages);
            }
            else
            {
                TempData["TaskDataMessage"] = $"{result.InsertedCount} rows inserted. {result.DuplicateCount} duplicate rows were skipped.";
            }

            return RedirectToAction("UploadExcelFiles");
        }

        [HttpPost]
        public async Task<IActionResult> UploadTimeStampsExcelFile(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                TempData["TimeStampsMessage"] = "Please select a valid Excel file.";
                return RedirectToAction("UploadExcelFiles");
            }

            var result = await deliveryDAL.UploadTimeStampsFromExcel(file);

            if (result.ErrorMessages.Any())
            {
                TempData["TimeStampsMessage"] = "Some errors occurred during upload.";
                TempData["TimeStampsErrors"] = JsonConvert.SerializeObject(result.ErrorMessages);
            }
            else
            {
                TempData["TimeStampsMessage"] = $"{result.InsertedCount} rows inserted. {result.DuplicateCount} duplicate rows were skipped.";
            }

            return RedirectToAction("UploadExcelFiles");
        }
    }
}
