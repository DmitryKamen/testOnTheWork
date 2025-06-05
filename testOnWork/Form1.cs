using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

public partial class Form1 : Form
{
    string connectionString = "your_connection_string_here";
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

        using (var connection = new SqlConnection(connectionString))
        {
            connection.Open();
            var adapter = new SqlDataAdapter("SELECT * FROM Clients", connection);
            adapter.Fill(clientsTable);
        }

        dataGridView.DataSource = clientsTable;
    }

    private void SaveChanges()
    {
        using (var connection = new SqlConnection(connectionString))
        {
            connection.Open();

            var adapter = new SqlDataAdapter("SELECT * FROM Clients", connection);
            var builder = new SqlCommandBuilder(adapter);

            adapter.Update(clientsTable);

            MessageBox.Show("Данные сохранены");
        }
    }

    private void DataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
    {
        MessageBox.Show("Ошибка данных: " + e.Exception.Message);
        e.ThrowException = false;
    }
}
