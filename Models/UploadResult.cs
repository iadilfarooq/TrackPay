namespace TrackPay.Models
{
    public class UploadResult
    {
        public bool IsSuccess { get; set; }
        public int InsertedCount { get; set; }
        public int DuplicateCount { get; set; }
        public List<string> ErrorMessages { get; set; } = new();
    }
}
