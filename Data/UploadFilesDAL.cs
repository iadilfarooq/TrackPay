using Microsoft.Data.SqlClient;
using System.Data;
using TrackPay.Models;

namespace TrackPay.Data
{
    public class UploadFilesDAL
    {
        private readonly string connectionString;

        public UploadFilesDAL(IConfiguration configuration)
        {
            connectionString = configuration.GetConnectionString("dbcs");
        }

        public List<TaskData> GetTaskDataOfRider(int month, int year, int? id = null)
        {
            List<TaskData> list = new List<TaskData>();

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                string query = @"SELECT Id, CourierID, City, [Name], PurchaseID, DeliveredDateTime, DistanceKM 
                         FROM TaskData
                         WHERE MONTH(DeliveredDateTime) = @Month AND YEAR(DeliveredDateTime) = @Year";

                if (id.HasValue)
                {
                    query += " AND CourierID = @CourierID";
                }

                using (SqlCommand cmd = new SqlCommand(query, con))
                {
                    cmd.Parameters.AddWithValue("@Month", month);
                    cmd.Parameters.AddWithValue("@Year", year);
                    if (id.HasValue)
                        cmd.Parameters.AddWithValue("@CourierID", id.Value);

                    con.Open();
                    SqlDataReader reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        list.Add(new TaskData
                        {
                            Id = Convert.ToInt32(reader["Id"]),
                            CourierID = Convert.ToInt32(reader["CourierID"]),
                            City = reader["City"].ToString(),
                            Name = reader["Name"].ToString(),
                            PurchaseID = reader["PurchaseID"].ToString(),
                            DeliveredDateTime = Convert.ToDateTime(reader["DeliveredDateTime"]),
                            DistanceKM = Convert.ToDouble(reader["DistanceKM"])
                        });
                    }

                    con.Close();
                }
            }

            return list;
        }

        public List<TimeStamps> GetTimeStampsOfRider(int month, int year, int? id = null)
        {
            List<TimeStamps> list = new List<TimeStamps>();

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                string query = @"SELECT Id, City, CourierID, [Name], StartDate, StartTime, EndTime 
                         FROM TimeStamps
                         WHERE MONTH(StartDate) = @Month AND YEAR(StartDate) = @Year";

                if (id.HasValue)
                {
                    query += " AND CourierID = @CourierID";
                }

                using (SqlCommand cmd = new SqlCommand(query, con))
                {
                    cmd.Parameters.AddWithValue("@Month", month);
                    cmd.Parameters.AddWithValue("@Year", year);
                    if (id.HasValue)
                        cmd.Parameters.AddWithValue("@CourierID", id.Value);

                    con.Open();
                    SqlDataReader reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        list.Add(new TimeStamps
                        {
                            Id = Convert.ToInt32(reader["Id"]),
                            City = reader["City"].ToString(),
                            CourierID = Convert.ToInt32(reader["CourierID"]),
                            Name = reader["Name"].ToString(),
                            StartDate = Convert.ToDateTime(reader["StartDate"]),
                            StartTime = Convert.ToDateTime(reader["StartTime"]),
                            EndTime = Convert.ToDateTime(reader["EndTime"])
                        });
                    }

                    con.Close();
                }
            }

            return list;
        }




    }
}
