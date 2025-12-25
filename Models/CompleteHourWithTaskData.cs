namespace TrackPay.Models
{
    public class CompleteHourWithTaskData
    {
        public int CourierID { get; set; }
        public string Name { get; set; }
        public string City { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public TimeSpan DutyTime { get; set; }
        public string TimeDuration { get; set; }
        public bool IsCompleteHour { get; set; }
        public int OrderDelivered { get; set; }
        public double? DistanceKM { get; set; }
        public double? HourlyPay { get; set; }
        public double? OrderPay { get; set; }
        public double? DistancePay { get; set; }
        public double? TotalPay { get; set; }
    }
}
