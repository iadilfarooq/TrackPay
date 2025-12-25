namespace TrackPay.Models
{
    public class TaskData
    {
        public int Id { get; set; }
        public int CourierID { get; set; }
        public string City { get; set; }
        public string Name { get; set; }
        public string PurchaseID { get; set; }
        public DateTime DeliveredDateTime { get; set; }
        public double DistanceKM { get; set; }
    }
}
