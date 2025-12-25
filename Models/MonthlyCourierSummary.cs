namespace TrackPay.Models
{
    public class MonthlyCourierSummary
    {
        public string MonthYear { get; set; }
        public int CourierID { get; set; }
        public string Name { get; set; }
        public double TotalHoursValue { get; set; }
        public string TotalHoursDisplay { get; set; }
        public double TotalHourlyPay { get; set; }
        public double TotalOrderPay { get; set; }
        public double TotalDistancePay { get; set; }
        public double TotalPay { get; set; }
    }
}
