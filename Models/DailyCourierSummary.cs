namespace TrackPay.Models
{
    public class DailyCourierSummary
    {
        public DateTime StartDate { get; set; }
        public int CourierID { get; set; }
        public string Name { get; set; }

        public double TotalHourlyPay { get; set; }
        public double TotalOrderPay { get; set; }
        public double TotalDistancePay { get; set; }
        public double TotalPay { get; set; }
    }
}
