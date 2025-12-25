using ClosedXML.Excel;
using Microsoft.CodeAnalysis.Elfie.Diagnostics;
using Microsoft.Data.SqlClient;
using System.Data;
using TrackPay.Models;

namespace TrackPay.Data
{
    public class DeliveryDAL
    {
        private readonly string connectionString;

        // Constants
        const double HourlyRate = 160;
        const double OrderRate = 60;
        const double DistanceRate = 10;

        public DeliveryDAL(IConfiguration configuration)
        {
            connectionString = configuration.GetConnectionString("dbcs");
        }

        public List<MonthlyCourierSummary> GetMonthlyCourierSummaries(int month, int year)
        {
            List<CompleteHourWithTaskData> fullData = GetCompleteHourWithDeliveriesCalc();

            var monthlySummaries = fullData
                .Where(x => x.StartDate.Month == month && x.StartDate.Year == year) // Filter here
                .GroupBy(x => new { x.CourierID })
                .Select(g => new MonthlyCourierSummary
                {
                    MonthYear = new DateTime(year, month, 1).ToString("MMMM yyyy"), // Same for all
                    CourierID = g.Key.CourierID,
                    Name = g.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x.Name))?.Name ?? "",
                    TotalHourlyPay = g.Sum(x => x.HourlyPay ?? 0),
                    TotalOrderPay = g.Sum(x => x.OrderPay ?? 0),
                    TotalDistancePay = g.Sum(x => x.DistancePay ?? 0),
                    TotalPay = g.Sum(x => x.TotalPay ?? 0)
                })
                .OrderBy(x => x.CourierID)
                .ToList();

            return monthlySummaries;
        }

        public List<DailyCourierSummary> GetDailyCourierSummaries(int month, int year, int courierId)
        {
            List<CompleteHourWithTaskData> fullData = GetCompleteHourWithDeliveriesCalc();

            var filteredData = fullData
                .Where(x => x.CourierID == courierId
                            && x.StartDate.Month == month
                            && x.StartDate.Year == year)
                .ToList();

            var dailySummaries = filteredData
                .GroupBy(x => new { x.CourierID, StartDate = x.StartDate.Date })
                .Select(g => new DailyCourierSummary
                {
                    StartDate = g.Key.StartDate,
                    CourierID = g.Key.CourierID,
                    Name = g.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x.Name))?.Name ?? "",
                    TotalHourlyPay = g.Sum(x => x.HourlyPay ?? 0),
                    TotalOrderPay = g.Sum(x => x.OrderPay ?? 0),
                    TotalDistancePay = g.Sum(x => x.DistancePay ?? 0),
                    TotalPay = g.Sum(x => x.TotalPay ?? 0)
                })
                .OrderBy(x => x.StartDate)
                .ToList();

            return dailySummaries;
        }

        public List<CompleteHourWithTaskData> GetHourlyDataAllCalc(DateTime selectedDate, int courierId)
        {
            List<CompleteHourWithTaskData> results = new List<CompleteHourWithTaskData>();
            CompleteHourWithTaskData combinedRecord = null;

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                // 1. Get all time blocks (complete & partial) without duty hour check
                using (SqlCommand cmd = new SqlCommand("sp_SplitAllHoursRecords", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 120;

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var recordCourierId = Convert.ToInt32(reader["CourierID"]);
                            var recordStartDate = Convert.ToDateTime(reader["StartDate"]);

                            if (recordCourierId == courierId && recordStartDate.Date == selectedDate.Date)
                            {
                                CompleteHourWithTaskData record = new CompleteHourWithTaskData
                                {
                                    CourierID = recordCourierId,
                                    Name = reader["Name"].ToString(),
                                    StartDate = recordStartDate,
                                    StartTime = Convert.ToDateTime(reader["StartTime"]),
                                    EndTime = Convert.ToDateTime(reader["EndTime"]),
                                    IsCompleteHour = Convert.ToBoolean(reader["IsCompleteHour"]),
                                    TimeDuration = reader["timeDuration"].ToString(),
                                    HourlyPay = 0 // Will be calculated
                                };

                                results.Add(record);
                            }
                        }
                    }
                }

                // 2. Get Task Data for the specific courier and date
                List<(int CourierID, DateTime DeliveredDateTime, double DistanceKM)> taskList = new List<(int, DateTime, double)>();

                using (SqlCommand taskCmd = new SqlCommand("sp_GetAllTaskData", con))
                {
                    taskCmd.CommandType = CommandType.StoredProcedure;
                    taskCmd.CommandTimeout = 120;

                    using (SqlDataReader taskReader = taskCmd.ExecuteReader())
                    {
                        while (taskReader.Read())
                        {
                            int taskCourierId = Convert.ToInt32(taskReader["CourierID"]);
                            DateTime deliveredDateTime = Convert.ToDateTime(taskReader["DeliveredDateTime"]);

                            if (taskCourierId == courierId && deliveredDateTime.Date == selectedDate.Date)
                            {
                                taskList.Add((
                                    taskCourierId,
                                    deliveredDateTime,
                                    Convert.ToDouble(taskReader["DistanceKM"])
                                ));
                            }
                        }
                    }
                }

                // 3. Calculate Pay for Each Block (no duty time filtering now)
                foreach (var record in results)
                {
                    var matchedOrders = taskList.Where(x =>
                        x.CourierID == record.CourierID &&
                        x.DeliveredDateTime >= record.StartTime &&
                        x.DeliveredDateTime <= record.EndTime
                    ).ToList();

                    record.OrderDelivered = matchedOrders.Count;
                    record.DistanceKM = matchedOrders.Sum(x => x.DistanceKM);
                    record.OrderPay = record.OrderDelivered * OrderRate;
                    record.DistancePay = (record.DistanceKM ?? 0) * DistanceRate;

                    double totalOrderAndDistancePay = record.OrderPay.Value + record.DistancePay.Value;
                    double minutesWorked = (record.EndTime - record.StartTime).TotalMinutes;
                    double proratedHourlyPay = (minutesWorked / 60) * HourlyRate;

                    if (record.IsCompleteHour)
                    {
                        record.HourlyPay = HourlyRate;
                        record.TotalPay = Math.Max(HourlyRate, totalOrderAndDistancePay);
                    }
                    else
                    {
                        record.HourlyPay = proratedHourlyPay;
                        record.TotalPay = Math.Max(proratedHourlyPay, totalOrderAndDistancePay);
                    }
                }

                // 4. Handle extra deliveries outside any block
                var matchedCourierHours = results.Select(x => new { x.CourierID, x.StartTime, x.EndTime }).ToList();
                foreach (var task in taskList)
                {
                    bool isAlreadyMatched = matchedCourierHours.Any(x =>
                        x.CourierID == task.CourierID &&
                        task.DeliveredDateTime >= x.StartTime &&
                        task.DeliveredDateTime <= x.EndTime
                    );

                    if (!isAlreadyMatched)
                    {
                        if (combinedRecord == null)
                        {
                            combinedRecord = new CompleteHourWithTaskData
                            {
                                CourierID = task.CourierID,
                                Name = results.FirstOrDefault()?.Name ?? "",
                                StartDate = task.DeliveredDateTime.Date,
                                StartTime = DateTime.MinValue,
                                EndTime = DateTime.MinValue,
                                OrderDelivered = 0,
                                DistanceKM = 0,
                                HourlyPay = 0,
                                OrderPay = 0,
                                DistancePay = 0,
                                TotalPay = 0
                            };
                        }

                        combinedRecord.OrderDelivered += 1;
                        combinedRecord.DistanceKM += task.DistanceKM;
                        combinedRecord.OrderPay += OrderRate;
                        combinedRecord.DistancePay += task.DistanceKM * DistanceRate;
                        combinedRecord.TotalPay += OrderRate + (task.DistanceKM * DistanceRate);
                    }
                }

                if (combinedRecord != null)
                {
                    results.Add(combinedRecord);
                }
            }

            return results
                .OrderBy(x => x.StartTime == DateTime.MinValue ? 1 : 0)
                .ThenBy(x => x.StartTime)
                .ToList();
        }

        public List<CompleteHourWithTaskData> GetMonthlyHourlyDataAllCalc(int year, int month, int courierId)
        {
            List<CompleteHourWithTaskData> monthlyResults = new List<CompleteHourWithTaskData>();

            // Get all days in the specified month
            // and year
            int daysInMonth = DateTime.DaysInMonth(year, month);

            // Process each day of the month
            for (int day = 1; day <= daysInMonth; day++)
            {
                DateTime currentDate = new DateTime(year, month, day);

                // Get the daily data using existing method
                var dailyData = GetHourlyDataAllCalc(currentDate, courierId);

                // Add the daily data to the monthly results
                monthlyResults.AddRange(dailyData);
            }

            // Return all results ordered by date then by time
            return monthlyResults
                .OrderBy(x => x.StartDate)
                .ThenBy(x => x.StartTime == DateTime.MinValue ? 1 : 0) // Put invalid times last
                .ThenBy(x => x.StartTime) // Sort valid times
                .ToList();
        }

        public List<CompleteHourWithTaskData> GetCompleteHourWithDeliveriesCalc()
        {
            List<CompleteHourWithTaskData> results = new List<CompleteHourWithTaskData>();

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                // 1. Get all time blocks (complete & partial) from the stored procedure
                using (SqlCommand cmd = new SqlCommand("sp_SplitAllHoursRecords", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 120;

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            CompleteHourWithTaskData record = new CompleteHourWithTaskData
                            {
                                CourierID = Convert.ToInt32(reader["CourierID"]),
                                Name = reader["Name"].ToString(),
                                City = reader["City"].ToString(),
                                StartDate = Convert.ToDateTime(reader["StartDate"]),
                                StartTime = Convert.ToDateTime(reader["StartTime"]),
                                EndTime = Convert.ToDateTime(reader["EndTime"]),
                                TimeDuration = reader["TimeDuration"].ToString(),
                                IsCompleteHour = Convert.ToBoolean(reader["IsCompleteHour"]),
                                HourlyPay = 0
                            };

                            results.Add(record);
                        }
                    }
                }

                // 2. Get Task Data (Orders & Distance)
                List<(int CourierID, DateTime DeliveredDateTime, double DistanceKM)> taskList = new List<(int, DateTime, double)>();

                using (SqlCommand taskCmd = new SqlCommand("sp_GetAllTaskData", con))
                {
                    taskCmd.CommandType = CommandType.StoredProcedure;
                    taskCmd.CommandTimeout = 120;

                    using (SqlDataReader taskReader = taskCmd.ExecuteReader())
                    {
                        while (taskReader.Read())
                        {
                            taskList.Add((
                                Convert.ToInt32(taskReader["CourierID"]),
                                Convert.ToDateTime(taskReader["DeliveredDateTime"]),
                                Convert.ToDouble(taskReader["DistanceKM"])
                            ));
                        }
                    }
                }

                // 3. Calculate Pay for Each Block
                foreach (var record in results)
                {
                    bool isZeroDuration = record.TimeDuration == "00:00:00";

                    if (isZeroDuration)
                    {
                        record.OrderDelivered = 0;
                        record.DistanceKM = 0;
                        record.OrderPay = 0;
                        record.DistancePay = 0;
                        record.TotalPay = 0;
                        continue;
                    }

                    var matchedOrders = taskList.Where(x =>
                        x.CourierID == record.CourierID &&
                        x.DeliveredDateTime >= record.StartTime &&
                        x.DeliveredDateTime <= record.EndTime
                    ).ToList();

                    record.OrderDelivered = matchedOrders.Count;
                    record.DistanceKM = matchedOrders.Sum(x => x.DistanceKM);
                    record.OrderPay = record.OrderDelivered * OrderRate;
                    record.DistancePay = (record.DistanceKM ?? 0) * DistanceRate;
                    double totalOrderAndDistancePay = record.OrderPay.Value + record.DistancePay.Value;

                    TimeSpan duration = TimeSpan.Parse(record.TimeDuration);
                    double proratedHourlyPay = (duration.TotalHours) * HourlyRate;

                    if (record.IsCompleteHour)
                    {
                        record.HourlyPay = HourlyRate;
                        record.TotalPay = Math.Max(HourlyRate, totalOrderAndDistancePay);
                    }
                    else
                    {
                        record.HourlyPay = proratedHourlyPay;
                        record.TotalPay = Math.Max(proratedHourlyPay, totalOrderAndDistancePay);
                    }
                }

                // 4. Handle Extra Deliveries Outside Any Time Block — No duty hour filtering
                var matchedCourierHours = results.Select(x => new { x.CourierID, x.StartTime, x.EndTime }).ToList();
                foreach (var task in taskList)
                {
                    bool isAlreadyMatched = matchedCourierHours.Any(x =>
                        x.CourierID == task.CourierID &&
                        task.DeliveredDateTime >= x.StartTime &&
                        task.DeliveredDateTime <= x.EndTime
                    );

                    if (!isAlreadyMatched)
                    {
                        CompleteHourWithTaskData extraRecord = new CompleteHourWithTaskData
                        {
                            CourierID = task.CourierID,
                            Name = "",
                            StartDate = task.DeliveredDateTime.Date,
                            StartTime = DateTime.MinValue,
                            EndTime = DateTime.MinValue,
                            OrderDelivered = 1,
                            DistanceKM = task.DistanceKM,
                            HourlyPay = 0,
                            OrderPay = OrderRate,
                            DistancePay = task.DistanceKM * DistanceRate,
                            TotalPay = OrderRate + (task.DistanceKM * DistanceRate)
                        };

                        results.Add(extraRecord);
                    }
                }
            }

            return results.OrderBy(x => x.StartDate).ThenBy(x => x.CourierID).ThenBy(x => x.StartTime).ToList();
        }

        public async Task<UploadResult> UploadTaskDataFromExcel(IFormFile file)
        {
            var result = new UploadResult();
            try
            {
                using var stream = file.OpenReadStream();
                using var workbook = new XLWorkbook(stream);
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header

                using SqlConnection conn = new SqlConnection(connectionString);
                await conn.OpenAsync();

                foreach (var row in rows)
                {
                    var courierID = row.Cell(1).GetValue<string>()?.Trim();
                    var city = row.Cell(2).GetValue<string>()?.Trim();
                    var name = row.Cell(3).GetValue<string>()?.Trim();
                    var purchaseID = row.Cell(4).GetValue<string>()?.Trim();
                    var dateStr = row.Cell(5).GetValue<string>()?.Trim();
                    var distanceStr = row.Cell(6).GetValue<string>()?.Trim();

                    if (string.IsNullOrWhiteSpace(courierID) ||
                        string.IsNullOrWhiteSpace(city) ||
                        string.IsNullOrWhiteSpace(name) ||
                        string.IsNullOrWhiteSpace(purchaseID) ||
                        string.IsNullOrWhiteSpace(dateStr) ||
                        string.IsNullOrWhiteSpace(distanceStr))
                    {
                        result.ErrorMessages.Add($"Row {row.RowNumber()}: One or more fields are empty.");
                        continue;
                    }

                    if (!DateTime.TryParse(dateStr, out DateTime deliveredDateTime) ||
                        !double.TryParse(distanceStr, out double distanceKM))
                    {
                        result.ErrorMessages.Add($"Row {row.RowNumber()}: Invalid format for date or distance.");
                        continue;
                    }

                    // Check for duplicates (now includes PurchaseID)
                    var checkCmd = new SqlCommand(@"
                SELECT COUNT(*) FROM TaskData 
                WHERE CourierID = @CourierID AND DeliveredDateTime = @DeliveredDateTime AND PurchaseID = @PurchaseID", conn);
                    checkCmd.Parameters.AddWithValue("@CourierID", courierID);
                    checkCmd.Parameters.AddWithValue("@DeliveredDateTime", deliveredDateTime);
                    checkCmd.Parameters.AddWithValue("@PurchaseID", purchaseID);

                    var exists = (int)await checkCmd.ExecuteScalarAsync() > 0;
                    if (exists)
                    {
                        result.DuplicateCount++;
                        continue;
                    }

                    var insertCmd = new SqlCommand(@"
                INSERT INTO TaskData (CourierID, City, Name, PurchaseID, DeliveredDateTime, DistanceKM)
                VALUES (@CourierID, @City, @Name, @PurchaseID, @DeliveredDateTime, @DistanceKM)", conn);
                    insertCmd.Parameters.AddWithValue("@CourierID", courierID);
                    insertCmd.Parameters.AddWithValue("@City", city);
                    insertCmd.Parameters.AddWithValue("@Name", name);
                    insertCmd.Parameters.AddWithValue("@PurchaseID", purchaseID);
                    insertCmd.Parameters.AddWithValue("@DeliveredDateTime", deliveredDateTime);
                    insertCmd.Parameters.AddWithValue("@DistanceKM", distanceKM);

                    await insertCmd.ExecuteNonQueryAsync();
                    result.InsertedCount++;
                }

                result.IsSuccess = result.ErrorMessages.Count == 0;
            }
            catch (Exception ex)
            {
                result.ErrorMessages.Add("Unexpected error: " + ex.Message);
                result.IsSuccess = false;
            }

            return result;
        }

        public async Task<UploadResult> UploadTimeStampsFromExcel(IFormFile file)
        {
            var result = new UploadResult();
            try
            {
                using var stream = file.OpenReadStream();
                using var workbook = new XLWorkbook(stream);
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header

                using SqlConnection conn = new SqlConnection(connectionString);
                await conn.OpenAsync();

                foreach (var row in rows)
                {
                    var city = row.Cell(1).GetValue<string>()?.Trim();
                    var courierIdStr = row.Cell(2).GetValue<string>()?.Trim();
                    var name = row.Cell(3).GetValue<string>()?.Trim();
                    var startDateStr = row.Cell(4).GetValue<string>()?.Trim();
                    var startTimeStr = row.Cell(5).GetValue<string>()?.Trim();
                    var endTimeStr = row.Cell(6).GetValue<string>()?.Trim();

                    if (string.IsNullOrWhiteSpace(city) ||
                        string.IsNullOrWhiteSpace(courierIdStr) ||
                        string.IsNullOrWhiteSpace(name) ||
                        string.IsNullOrWhiteSpace(startDateStr) ||
                        string.IsNullOrWhiteSpace(startTimeStr) ||
                        string.IsNullOrWhiteSpace(endTimeStr))
                    {
                        result.ErrorMessages.Add($"Missing required field at row: {row.RowNumber()}");
                        continue;
                    }

                    if (!int.TryParse(courierIdStr, out int courierID) ||
                        !DateTime.TryParse(startDateStr, out DateTime startDate) ||
                        !DateTime.TryParse(startTimeStr, out DateTime startTime) ||
                        !DateTime.TryParse(endTimeStr, out DateTime endTime))
                    {
                        result.ErrorMessages.Add($"Invalid format at row: {row.RowNumber()}");
                        continue;
                    }

                    var checkCmd = new SqlCommand(@"
                SELECT COUNT(*) FROM TimeStamps 
                WHERE CourierID = @CourierID AND StartTime = @StartTime AND EndTime = @EndTime", conn);
                    checkCmd.Parameters.AddWithValue("@CourierID", courierID);
                    checkCmd.Parameters.AddWithValue("@StartTime", startTime);
                    checkCmd.Parameters.AddWithValue("@EndTime", endTime);

                    var exists = (int)await checkCmd.ExecuteScalarAsync() > 0;
                    if (exists)
                    {
                        result.DuplicateCount++;
                        continue;
                    }

                    var insertCmd = new SqlCommand(@"
                INSERT INTO TimeStamps (City, CourierID, Name, StartDate, StartTime, EndTime)
                VALUES (@City, @CourierID, @Name, @StartDate, @StartTime, @EndTime)", conn);
                    insertCmd.Parameters.AddWithValue("@City", city);
                    insertCmd.Parameters.AddWithValue("@CourierID", courierID);
                    insertCmd.Parameters.AddWithValue("@Name", name);
                    insertCmd.Parameters.AddWithValue("@StartDate", startDate);
                    insertCmd.Parameters.AddWithValue("@StartTime", startTime);
                    insertCmd.Parameters.AddWithValue("@EndTime", endTime);

                    await insertCmd.ExecuteNonQueryAsync();
                    result.InsertedCount++;
                }

                result.IsSuccess = result.ErrorMessages.Count == 0;
            }
            catch (Exception ex)
            {
                result.ErrorMessages.Add("Unexpected error: " + ex.Message);
                result.IsSuccess = false;
            }

            return result;
        }
    }
}
