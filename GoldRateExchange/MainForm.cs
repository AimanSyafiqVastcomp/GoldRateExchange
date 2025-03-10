using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

// Add a specific using for HtmlAgilityPack to avoid ambiguity
using HtmlAgilityPack;

namespace GoldRatesExtractor
{
    public partial class MainForm : Form
    {
        // Connection string for LocalDB
        private string connectionString = "Data Source=LAPTOP-NBAO3H83;Initial Catalog=GoldRatesDB;User ID=sa;Password=1234;";

        public MainForm()
        {
            InitializeComponent();
        }

        // The designer file will handle the InitializeComponent method, don't define it here
        // We'll move all the controls setup to the designer file

        private void btnExtract_Click(object sender, EventArgs e)
        {
            try
            {
                statusTextBox.Text = "Starting extraction process...";

                // Instead of looking for a fixed file, let's use a dialog to select the file
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Title = "Select HTML Content File",
                    Filter = "Text Files (*.txt)|*.txt|HTML Files (*.html)|*.html|All Files (*.*)|*.*",
                    FilterIndex = 1
                };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string htmlFilePath = openFileDialog.FileName;
                    string htmlContent = File.ReadAllText(htmlFilePath);
                    ProcessHtmlContent(htmlContent);
                }
                else
                {
                    statusTextBox.Text = "File selection cancelled.";
                }
            }
            catch (Exception ex)
            {
                statusTextBox.Text = $"Error: {ex.Message}";
            }
        }

        private void ProcessHtmlContent(string htmlContent)
        {
            statusTextBox.Text = "Processing HTML content...";

            // Use fully qualified HtmlAgilityPack.HtmlDocument to avoid ambiguity
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(htmlContent);

            // Create DataTables to store our rates - one for each table
            DataTable ourRatesTable = new DataTable();
            ourRatesTable.Columns.Add("DetailName", typeof(string));
            ourRatesTable.Columns.Add("WeBuy", typeof(decimal));
            ourRatesTable.Columns.Add("WeSell", typeof(decimal));
            ourRatesTable.Columns.Add("ExtractedDate", typeof(DateTime));

            DataTable customerSellTable = new DataTable();
            customerSellTable.Columns.Add("DetailName", typeof(string));
            customerSellTable.Columns.Add("WeBuy", typeof(decimal));
            customerSellTable.Columns.Add("ExtractedDate", typeof(DateTime));

            // Extract Our Rates section
            statusTextBox.Text = "Extracting 'OUR RATES' data...";
            try
            {
                // Get all rows from the first table (OUR RATES), excluding the header row
                var ourRatesRows = doc.DocumentNode.SelectNodes("//div[contains(@class, 'col-lg-6')][1]//table//tbody//tr");

                if (ourRatesRows != null)
                {
                    foreach (var row in ourRatesRows)
                    {
                        var cells = row.SelectNodes("td");
                        if (cells != null && cells.Count >= 3)
                        {
                            string detail = Regex.Replace(cells[0].InnerText, @"\s+", " ").Trim();
                            string weBuyStr = cells[1].InnerText.Trim();
                            string weSellStr = cells[2].InnerText.Trim();

                            // Extract the numeric values using regex
                            var weBuyMatch = Regex.Match(weBuyStr, @"[\d,\.]+");
                            var weSellMatch = Regex.Match(weSellStr, @"[\d,\.]+");

                            if (weBuyMatch.Success && weSellMatch.Success)
                            {
                                decimal weBuy = decimal.Parse(weBuyMatch.Value, System.Globalization.NumberStyles.Any);
                                decimal weSell = decimal.Parse(weSellMatch.Value, System.Globalization.NumberStyles.Any);

                                ourRatesTable.Rows.Add(detail, weBuy, weSell, DateTime.Now);
                            }
                        }
                    }
                }

                // Extract Customer Sell section
                statusTextBox.Text = "Extracting 'CUSTOMER SELL' data...";
                var customerSellRows = doc.DocumentNode.SelectNodes("//div[contains(@class, 'col-lg-6')][2]//table//tbody//tr");

                if (customerSellRows != null)
                {
                    foreach (var row in customerSellRows)
                    {
                        var cells = row.SelectNodes("td");
                        if (cells != null && cells.Count >= 2)
                        {
                            string detail = Regex.Replace(cells[0].InnerText, @"\s+", " ").Trim();
                            string weBuyStr = cells[1].InnerText.Trim();

                            // Extract the numeric value using regex
                            var weBuyMatch = Regex.Match(weBuyStr, @"[\d,\.]+");

                            if (weBuyMatch.Success)
                            {
                                decimal weBuy = decimal.Parse(weBuyMatch.Value, System.Globalization.NumberStyles.Any);

                                // Customer Sell section only has WeBuy value, WeSell is null/0
                                customerSellTable.Rows.Add(detail, weBuy, DateTime.Now);
                            }
                        }
                    }
                }

                SaveToDatabase(ourRatesTable, customerSellTable);
            }
            catch (Exception ex)
            {
                statusTextBox.Text = $"Error during extraction: {ex.Message}";
            }
        }

        private void SaveToDatabase(DataTable ourRatesTable, DataTable customerSellTable)
        {
            statusTextBox.Text = "Saving data to SQL database...";

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    // Note: We're removing the call to sp_CreateGoldRatesTables since the tables should already exist
                    // The tables are now created/modified directly in SSMS

                    // Update OurRates table using stored procedure
                    int ourRatesUpdated = 0;
                    foreach (DataRow row in ourRatesTable.Rows)
                    {
                        string detailName = row["DetailName"].ToString();
                        decimal weBuy = (decimal)row["WeBuy"];
                        decimal weSell = (decimal)row["WeSell"];

                        using (SqlCommand command = new SqlCommand("sp_ourRate_upsert", connection))
                        {
                            command.CommandType = CommandType.StoredProcedure;
                            command.Parameters.AddWithValue("@DetailName", detailName);
                            command.Parameters.AddWithValue("@WeBuy", weBuy);
                            command.Parameters.AddWithValue("@WeSell", weSell);

                            // Execute and potentially get the ID
                            var result = command.ExecuteScalar();
                            ourRatesUpdated++;
                        }
                    }

                    // Update CustomerSell table using stored procedure
                    int customerSellUpdated = 0;
                    foreach (DataRow row in customerSellTable.Rows)
                    {
                        string detailName = row["DetailName"].ToString();
                        decimal weBuy = (decimal)row["WeBuy"];

                        using (SqlCommand command = new SqlCommand("sp_customerSell_upsert", connection))
                        {
                            command.CommandType = CommandType.StoredProcedure;
                            command.Parameters.AddWithValue("@DetailName", detailName);
                            command.Parameters.AddWithValue("@WeBuy", weBuy);

                            // Execute and potentially get the ID
                            var result = command.ExecuteScalar();
                            customerSellUpdated++;
                        }
                    }

                    statusTextBox.Text = $"Success! Updated {ourRatesUpdated} 'Our Rates' entries and {customerSellUpdated} 'Customer Sell' entries.";
                }
            }
            catch (Exception ex)
            {
                statusTextBox.Text = $"Database error: {ex.Message}";
            }
        }
    }
}