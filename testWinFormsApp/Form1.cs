using Microsoft.Data.SqlClient;
using System.Data;

namespace testWinFormsApp
{
    public partial class Form1 : Form
    {
        string connectionString = @"Server=DESKTOP-PF5QIJD;Database=DatabaseTest;
        User Id=user1;Password=sa;TrustServerCertificate=True";
        DataTable clientsTable = new DataTable();

        public Form1()
        {
            InitializeComponent();

            dataGridView1.DataError += DataGridView1_DataError;
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            LoadClients();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveChanges();
        }

        private void LoadClients()
        {
            clientsTable.Clear();

            var connection = new SqlConnection(connectionString);
            connection.Open();
            var adapter = new SqlDataAdapter("SELECT * FROM Clients", connection);
            adapter.Fill(clientsTable);

            dataGridView1.DataSource = clientsTable;
        }

        private void SaveChanges()
        {
            var connection = new SqlConnection(connectionString);

            connection.Open();

            var adapter = new SqlDataAdapter("SELECT * FROM Clients", connection);
            var builder = new SqlCommandBuilder(adapter);

            adapter.Update(clientsTable);

            MessageBox.Show("Данные сохранены");

        }

        private void DataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("Ошибка данных: " + e.Exception.Message);
            e.ThrowException = false;
        }
    }
}
