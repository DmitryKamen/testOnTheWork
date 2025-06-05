using ClosedXML.Excel;
using Microsoft.Data.SqlClient;


namespace testConsoleApp
{

    class Program
    {
        static string connectionString = @"Server=DESKTOP-PF5QIJD;Database=DatabaseTest;User Id=user1;Password=sa;";

        static void Main()
        {
            string filePath = "clients.xlsx";

            var clients = ReadClientsFromExcel(filePath);
            InsertClientsToDatabase(clients);

            Console.WriteLine("Импорт завершён.");
        }

        static List<Client> ReadClientsFromExcel(string path)
        {
            var clients = new List<Client>();

            var workbook = new XLWorkbook(path);

            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RangeUsed().RowsUsed().Skip(1);

            foreach (var row in rows)
            {
                try
                {
                    var client = new Client
                    {
                        CardCode = ParseInt(row.Cell("A").GetValue<string>()),
                        LastName = row.Cell("B").GetString(),
                        FirstName = row.Cell("C").GetString(),
                        SurName = row.Cell("D").GetString(),
                        PhoneMobile = row.Cell("E").GetString(),
                        Email = row.Cell("F").GetString(),
                        GenderId = row.Cell("G").GetString(),
                        Birthday = ParseDate(row.Cell("H").GetString()),
                        City = row.Cell("I").GetString(),
                        Pincode = row.Cell("J").GetString(),
                        Bonus = ParseInt(row.Cell("K").GetString()),
                        Turrover = ParseInt(row.Cell("L").GetString())
                    };
                    clients.Add(client);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка при чтении строки {row.RowNumber()}: {ex.Message}");

                }
            }


            return clients;
        }

        static int ParseInt(string s)
        {
            if (int.TryParse(s, out int val))
                return val;
            return 0;
        }

        static DateTime? ParseDate(string s)
        {
            if (DateTime.TryParse(s, out DateTime dt))
                return dt;
            return null;
        }

        static void InsertClientsToDatabase(List<Client> clients)
        {
            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();

                foreach (var client in clients)
                {
                    using (var command = new SqlCommand(@"
                    INSERT INTO Clients 
                    (CardCode, LastName, FirstName, SurName, PhoneMobile, Email, GenderId, Birthday, City, Pincode, Bonus, Turrover)
                    VALUES
                    (@CardCode, @LastName, @FirstName, @SurName, @PhoneMobile, @Email, @GenderId, @Birthday, @City, @Pincode, @Bonus, @Turrover)", connection))
                    {
                        command.Parameters.AddWithValue("@CardCode", client.CardCode);
                        command.Parameters.AddWithValue("@LastName", (object)client.LastName ?? DBNull.Value);
                        command.Parameters.AddWithValue("@FirstName", (object)client.FirstName ?? DBNull.Value);
                        command.Parameters.AddWithValue("@SurName", (object)client.SurName ?? DBNull.Value);
                        command.Parameters.AddWithValue("@PhoneMobile", (object)client.PhoneMobile ?? DBNull.Value);
                        command.Parameters.AddWithValue("@Email", (object)client.Email ?? DBNull.Value);
                        command.Parameters.AddWithValue("@GenderId", (object)client.GenderId ?? DBNull.Value);
                        command.Parameters.AddWithValue("@Birthday", (object)client.Birthday ?? DBNull.Value);
                        command.Parameters.AddWithValue("@City", (object)client.City ?? DBNull.Value);
                        command.Parameters.AddWithValue("@Pincode", (object)client.Pincode ?? DBNull.Value);
                        command.Parameters.AddWithValue("@Bonus", (object)client.Bonus ?? DBNull.Value);
                        command.Parameters.AddWithValue("@Turrover", (object)client.Turrover ?? DBNull.Value);

                        command.ExecuteNonQuery();
                    }
                }
            }
        }
    }
}