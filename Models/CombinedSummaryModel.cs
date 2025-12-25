namespace TrackPay.Models
{
    public class CombinedSummaryViewModel
    {
        public int Month { get; set; }
        public int Year { get; set; }
        public DateTime? SelectedMonthYear { get; set; }
        public int CourierID { get; set; }
        public DateTime? SelectedDate { get; set; }
        public int? CourierIdHourly { get; set; }

        public List<CompleteHourWithTaskData> MonthlyHourlyData { get; set; }
        public bool ShowMonthlyHourlySummary { get; set; }
        public DateTime? SelectedMonthYearForHourly { get; set; }

        public string SelectedMonthNameHourly { get; set; }
        public int MonthHourly { get; set; }
        public int YearHourly { get; set; }

        public int CourierIdMonthlyHourly { get; set; }
        public string SelectedCourierNameMonthlyHourly { get; set; }

        public string SelectedCourierName { get; set; }
        public string SelectedCourierNameForDaily { get; set; }
        public string SelectedMonthName { get; set; }

        public bool ShowMonthlySummary { get; set; }
        public bool ShowDailySummary { get; set; }
        public bool ShowHourlySummary { get; set; }

        public List<MonthlyCourierSummary> MonthlySummaries { get; set; }
        public List<DailyCourierSummary> DailySummaries { get; set; }
        public List<CompleteHourWithTaskData> HourlyData { get; set; }
    }
}
