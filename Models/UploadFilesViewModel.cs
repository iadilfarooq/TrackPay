namespace TrackPay.Models
{
    public class UploadFilesViewModel
    {
        public int Month { get; set; }
        public int Year { get; set; }
        public int? CourierId { get; set; }
        public string DataType { get; set; } // "TaskData" or "TimeStamps"

        public List<TaskData>? TaskData { get; set; }
        public List<TimeStamps>? TimeStamps { get; set; }
    }

}
