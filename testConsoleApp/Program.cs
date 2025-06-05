using ClosedXML.Excel;
using Microsoft.Data.SqlClient;


namespace testConsoleApp
{

    class Program
    {
        static string connectionString = @"Server=DESKTOP-PF5QIJD;Database=DatabaseTest;User Id=user1;
        Password=sa;TrustServerCertificate=True;";

        static void Main()
        {
            string filePath = "C:\\Users\\piton\\OneDrive\\Documents\\List.xlsx";

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
                    var checkCommand = new SqlCommand("SELECT COUNT(*) FROM Clients WHERE CardCode = @CardCode", connection);
                    checkCommand.Parameters.AddWithValue("@CardCode", client.CardCode);

                    int count = (int)checkCommand.ExecuteScalar();

                    if (count == 0)
                    {
                        var insertCommand = new SqlCommand(@"
                    INSERT INTO Clients 
                    (CardCode, LastName, FirstName, SurName, PhoneMobile, Email, GenderId, Birthday, City, Pincode, Bonus, Turrover)
                    VALUES
                    (@CardCode, @LastName, @FirstName, @SurName, @PhoneMobile, @Email, @GenderId, @Birthday, @City, @Pincode, @Bonus, @Turrover)", connection);

                        insertCommand.Parameters.AddWithValue("@CardCode", client.CardCode);
                        insertCommand.Parameters.AddWithValue("@LastName", (object)client.LastName ?? DBNull.Value);
                        insertCommand.Parameters.AddWithValue("@FirstName", (object)client.FirstName ?? DBNull.Value);
                        insertCommand.Parameters.AddWithValue("@SurName", (object)client.SurName ?? DBNull.Value);
                        insertCommand.Parameters.AddWithValue("@PhoneMobile", (object)client.PhoneMobile ?? DBNull.Value);
                        insertCommand.Parameters.AddWithValue("@Email", (object)client.Email ?? DBNull.Value);
                        insertCommand.Parameters.AddWithValue("@GenderId", (object)client.GenderId ?? DBNull.Value);
                        insertCommand.Parameters.AddWithValue("@Birthday", (object)client.Birthday ?? DBNull.Value);
                        insertCommand.Parameters.AddWithValue("@City", (object)client.City ?? DBNull.Value);
                        insertCommand.Parameters.AddWithValue("@Pincode", (object)client.Pincode ?? DBNull.Value);
                        insertCommand.Parameters.AddWithValue("@Bonus", (object)client.Bonus ?? DBNull.Value);
                        insertCommand.Parameters.AddWithValue("@Turrover", (object)client.Turrover ?? DBNull.Value);

                        insertCommand.ExecuteNonQuery();
                    }
                    else
                    {
                        var updateCommand = new SqlCommand(@"
                    UPDATE Clients SET 
                        LastName = @LastName,
                        FirstName = @FirstName,
                        SurName = @SurName,
                        PhoneMobile = @PhoneMobile,
                        Email = @Email,
                        GenderId = @GenderId,
                        Birthday = @Birthday,
                        City = @City,
                        Pincode = @Pincode,
                        Bonus = @Bonus,
                        Turrover = @Turrover
                    WHERE CardCode = @CardCode", connection);

                        updateCommand.Parameters.AddWithValue("@CardCode", client.CardCode);
                        updateCommand.Parameters.AddWithValue("@LastName", (object)client.LastName ?? DBNull.Value);
                        updateCommand.Parameters.AddWithValue("@FirstName", (object)client.FirstName ?? DBNull.Value);
                        updateCommand.Parameters.AddWithValue("@SurName", (object)client.SurName ?? DBNull.Value);
                        updateCommand.Parameters.AddWithValue("@PhoneMobile", (object)client.PhoneMobile ?? DBNull.Value);
                        updateCommand.Parameters.AddWithValue("@Email", (object)client.Email ?? DBNull.Value);
                        updateCommand.Parameters.AddWithValue("@GenderId", (object)client.GenderId ?? DBNull.Value);
                        updateCommand.Parameters.AddWithValue("@Birthday", (object)client.Birthday ?? DBNull.Value);
                        updateCommand.Parameters.AddWithValue("@City", (object)client.City ?? DBNull.Value);
                        updateCommand.Parameters.AddWithValue("@Pincode", (object)client.Pincode ?? DBNull.Value);
                        updateCommand.Parameters.AddWithValue("@Bonus", (object)client.Bonus ?? DBNull.Value);
                        updateCommand.Parameters.AddWithValue("@Turrover", (object)client.Turrover ?? DBNull.Value);

                        updateCommand.ExecuteNonQuery();
                    }
                }
            }
        }

    }
}
