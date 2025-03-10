using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO; 
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace GoldRatesExtractor
{
    public partial class MainForm : Form
    {
        private string connectionString;
        private string currentCompanyName;
        private ChromeDriver driver;

        public MainForm()
        {
            InitializeComponent();
            LoadConfiguration();
            InitializeSelenium();
        }

        private void InitializeSelenium()
        {
            try
            {
                var options = new ChromeOptions();
                options.AddArgument("--headless"); // Run Chrome in headless mode (no UI)
                options.AddArgument("--disable-gpu");
                options.AddArgument("--window-size=1920,1080");

                driver = new ChromeDriver(options);
                statusTextBox.Text = "Selenium initialized successfully.";
            }
            catch (Exception ex)
            {
                statusTextBox.Text = $"Error initializing Selenium: {ex.Message}";
                MessageBox.Show($"Error initializing Selenium: {ex.Message}\n\nMake sure you have ChromeDriver installed and in your PATH.",
                    "Selenium Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);

            if (driver != null)
            {
                driver.Quit();
                driver.Dispose();
            }
        }

        private void LoadConfiguration()
        {
            try
            {
                // Load connection string from app.config
                connectionString = System.Configuration.ConfigurationManager.AppSettings["ConnectionString"];

                // Populate the website combo box
                cboWebsite.Items.Add($"1: {System.Configuration.ConfigurationManager.AppSettings["CompanyName1"]} - {System.Configuration.ConfigurationManager.AppSettings["UrlOption1"]}");
                cboWebsite.Items.Add($"2: {System.Configuration.ConfigurationManager.AppSettings["CompanyName2"]} - {System.Configuration.ConfigurationManager.AppSettings["UrlOption2"]}");
                cboWebsite.SelectedIndex = 0; // Default to first option
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading configuration: {ex.Message}", "Configuration Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void btnExtract_Click(object sender, EventArgs e)
        {
            if (driver == null)
            {
                MessageBox.Show("Selenium WebDriver is not initialized. Cannot proceed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // Disable the button during extraction
                btnExtract.Enabled = false;
                statusTextBox.Text = "Starting extraction process...";

                // Determine which website was selected
                int selectedIndex = cboWebsite.SelectedIndex;
                string url;

                if (selectedIndex == 0)
                {
                    url = System.Configuration.ConfigurationManager.AppSettings["UrlOption1"];
                    currentCompanyName = System.Configuration.ConfigurationManager.AppSettings["CompanyName1"];
                }
                else
                {
                    url = System.Configuration.ConfigurationManager.AppSettings["UrlOption2"];
                    currentCompanyName = System.Configuration.ConfigurationManager.AppSettings["CompanyName2"];
                }

                // Navigate to the URL and wait for JavaScript to load content
                await LoadWebpageAsync(url);

                // Process the page based on the selected website
                bool success;
                if (selectedIndex == 0)
                {
                    success = await ProcessTTTBullionPageAsync();
                }
                else
                {
                    success = await ProcessMSGoldPageAsync();
                }

                if (!success)
                {
                    MessageBox.Show($"Unable to extract data from {currentCompanyName}. No data was saved to the database.",
                        "Extraction Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                statusTextBox.Text += Environment.NewLine + $"Error: {ex.Message}";
                statusTextBox.Text += Environment.NewLine + $"Stack trace: {ex.StackTrace}";
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Re-enable the button
                btnExtract.Enabled = true;
            }
        }

        private async Task LoadWebpageAsync(string url)
        {
            statusTextBox.Text += Environment.NewLine + $"Loading webpage: {url}";

            await Task.Run(() => {
                driver.Navigate().GoToUrl(url);

                // Wait for page to load
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));

                // Additional wait for any AJAX calls to complete
                System.Threading.Thread.Sleep(2000);
            });

            statusTextBox.Text += Environment.NewLine + "Page loaded successfully";
        }

        private async Task<bool> ProcessTTTBullionPageAsync()
        {
            statusTextBox.Text += Environment.NewLine + "Processing TTT Bullion page...";

            try
            {
                // Create a collection to store the extracted data
                DataTable extractedData = new DataTable();
                extractedData.Columns.Add("DetailName", typeof(string));
                extractedData.Columns.Add("WeBuy", typeof(decimal));
                extractedData.Columns.Add("WeSell", typeof(decimal));

                // Wait a bit longer for dynamic content to load
                await Task.Delay(5000);

                // Find all tables on the page
                var tableElements = driver.FindElements(By.TagName("table"));
                int tableCount = tableElements.Count;
                statusTextBox.Text += Environment.NewLine + $"Found {tableCount} tables on the page";

                // We only want the first table (Gold rates), not the silver rates
                if (tableCount > 0)
                {
                    var goldTable = tableElements[0]; // First table contains gold rates
                    string tableText = goldTable.Text;

                    // Check if this is indeed the gold table
                    if (tableText.Contains("Gold") && !tableText.Contains("Silver"))
                    {
                        statusTextBox.Text += Environment.NewLine + "Processing Gold rates table";

                        // Get all rows in the table
                        var rows = goldTable.FindElements(By.TagName("tr"));
                        int rowCount = rows.Count;
                        statusTextBox.Text += Environment.NewLine + $"Gold table has {rowCount} rows";

                        // Skip the header row, process the data rows
                        for (int i = 1; i < rowCount; i++) // Start at 1 to skip header
                        {
                            var row = rows[i];
                            var cells = row.FindElements(By.TagName("td"));

                            if (cells.Count >= 3)
                            {
                                string detail = cells[0].Text.Trim();
                                string weBuyStr = cells[1].Text.Trim();
                                string weSellStr = cells[2].Text.Trim();

                                // Log what we found
                                statusTextBox.Text += Environment.NewLine +
                                    $"Gold data - Detail: '{detail}', Buy: '{weBuyStr}', Sell: '{weSellStr}'";

                                // Try to extract numeric values
                                if (!string.IsNullOrEmpty(weBuyStr) && !string.IsNullOrEmpty(weSellStr))
                                {
                                    var weBuyMatch = Regex.Match(weBuyStr, @"[\d,\.]+");
                                    var weSellMatch = Regex.Match(weSellStr, @"[\d,\.]+");

                                    if (weBuyMatch.Success && weSellMatch.Success)
                                    {
                                        try
                                        {
                                            decimal weBuy = decimal.Parse(weBuyMatch.Value, System.Globalization.NumberStyles.Any);
                                            decimal weSell = decimal.Parse(weSellMatch.Value, System.Globalization.NumberStyles.Any);

                                            // No need to normalize the name - it's already in the correct format
                                            extractedData.Rows.Add(detail, weBuy, weSell);
                                            statusTextBox.Text += Environment.NewLine +
                                                $"Extracted gold rate: {detail}, Buy: {weBuy}, Sell: {weSell}";
                                        }
                                        catch (Exception ex)
                                        {
                                            statusTextBox.Text += Environment.NewLine + $"Error parsing values: {ex.Message}";
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        statusTextBox.Text += Environment.NewLine + "First table doesn't appear to be the Gold rates table. Checking all tables...";

                        // If the first table wasn't the gold table, search through all tables
                        bool foundGoldTable = false;

                        for (int tableIndex = 0; tableIndex < tableCount; tableIndex++)
                        {
                            var table = tableElements[tableIndex];
                            string tblText = table.Text;

                            // Check if this table contains gold rates
                            if (tblText.Contains("Gold") && !tblText.Contains("Silver"))
                            {
                                statusTextBox.Text += Environment.NewLine + $"Found Gold rates in table {tableIndex}";

                                var rows = table.FindElements(By.TagName("tr"));
                                foreach (var row in rows)
                                {
                                    var cells = row.FindElements(By.TagName("td"));
                                    if (cells.Count >= 3)
                                    {
                                        string detail = cells[0].Text.Trim();

                                        // Skip header rows
                                        if (detail.Contains("Gold") || detail.Contains("DETAILS") || string.IsNullOrWhiteSpace(detail))
                                            continue;

                                        string weBuyStr = cells[1].Text.Trim();
                                        string weSellStr = cells[2].Text.Trim();

                                        statusTextBox.Text += Environment.NewLine +
                                            $"Gold data - Detail: '{detail}', Buy: '{weBuyStr}', Sell: '{weSellStr}'";

                                        if (!string.IsNullOrEmpty(weBuyStr) && !string.IsNullOrEmpty(weSellStr))
                                        {
                                            var weBuyMatch = Regex.Match(weBuyStr, @"[\d,\.]+");
                                            var weSellMatch = Regex.Match(weSellStr, @"[\d,\.]+");

                                            if (weBuyMatch.Success && weSellMatch.Success)
                                            {
                                                try
                                                {
                                                    decimal weBuy = decimal.Parse(weBuyMatch.Value, System.Globalization.NumberStyles.Any);
                                                    decimal weSell = decimal.Parse(weSellMatch.Value, System.Globalization.NumberStyles.Any);

                                                    extractedData.Rows.Add(detail, weBuy, weSell);
                                                    statusTextBox.Text += Environment.NewLine +
                                                        $"Extracted gold rate: {detail}, Buy: {weBuy}, Sell: {weSell}";
                                                    foundGoldTable = true;
                                                }
                                                catch (Exception ex)
                                                {
                                                    statusTextBox.Text += Environment.NewLine + $"Error parsing values: {ex.Message}";
                                                }
                                            }
                                        }
                                    }
                                }

                                // If we found and processed a gold table, we can stop looking
                                if (foundGoldTable)
                                    break;
                            }
                        }

                        if (!foundGoldTable)
                        {
                            statusTextBox.Text += Environment.NewLine + "Could not identify a Gold rates table on the page.";
                            return false;
                        }
                    }
                }
                else
                {
                    statusTextBox.Text += Environment.NewLine + "No tables found on the page.";
                    return false;
                }

                // Check if we found any data
                if (extractedData.Rows.Count == 0)
                {
                    statusTextBox.Text += Environment.NewLine + "No gold rate data extracted.";
                    return false;
                }

                statusTextBox.Text += Environment.NewLine + $"Extracted {extractedData.Rows.Count} gold rates.";

                // Clear existing data and save to database
                await ClearExistingCompanyDataAsync(currentCompanyName);
                await SaveTTTBullionDataAsync(extractedData);
                return true;
            }
            catch (Exception ex)
            {
                statusTextBox.Text += Environment.NewLine + $"Error processing TTT Bullion page: {ex.Message}";
                statusTextBox.Text += Environment.NewLine + $"Stack trace: {ex.StackTrace}";
                return false;
            }
        }

        

        private async Task<bool> ProcessMSGoldPageAsync()
        {
            statusTextBox.Text += Environment.NewLine + "Processing MS Gold page...";

            try
            {
                // Create collections to store the extracted data
                DataTable ourRatesData = new DataTable();
                ourRatesData.Columns.Add("DetailName", typeof(string));
                ourRatesData.Columns.Add("WeBuy", typeof(decimal));
                ourRatesData.Columns.Add("WeSell", typeof(decimal));
                ourRatesData.Columns.Add("Type", typeof(string));

                DataTable customerSellData = new DataTable();
                customerSellData.Columns.Add("DetailName", typeof(string));
                customerSellData.Columns.Add("WeBuy", typeof(decimal));
                customerSellData.Columns.Add("WeSell", typeof(decimal));
                customerSellData.Columns.Add("Type", typeof(string));

                // Find all tables on the page
                var tableElements = driver.FindElements(By.TagName("table"));
                statusTextBox.Text += Environment.NewLine + $"Found {tableElements.Count} tables on the page";

                bool foundOurRatesData = false;
                bool foundCustomerSellData = false;

                // Process each table
                for (int i = 0; i < tableElements.Count; i++)
                {
                    var table = tableElements[i];
                    string tableText = table.Text;
                    statusTextBox.Text += Environment.NewLine + $"Table {i} text: {tableText.Substring(0, Math.Min(100, tableText.Length))}...";

                    // Look for our rates table (usually has 3 columns including Buy and Sell)
                    if (tableText.Contains("WE BUY") && tableText.Contains("WE SELL"))
                    {
                        statusTextBox.Text += Environment.NewLine + $"Found OurRates table (index {i})";

                        var rows = table.FindElements(By.TagName("tr"));
                        statusTextBox.Text += Environment.NewLine + $"Table has {rows.Count} rows";

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count >= 3)
                            {
                                string detail = cells[0].Text.Trim();

                                // Skip header rows
                                if (detail.Contains("DETAILS") || detail.Contains("WE BUY") || string.IsNullOrWhiteSpace(detail))
                                    continue;

                                string weBuyStr = cells[1].Text.Trim();
                                string weSellStr = cells[2].Text.Trim();

                                // Log raw values
                                statusTextBox.Text += Environment.NewLine + $"Raw values - Detail: '{detail}', Buy: '{weBuyStr}', Sell: '{weSellStr}'";

                                // Extract numeric values
                                var weBuyMatch = Regex.Match(weBuyStr, @"[\d,\.]+");
                                var weSellMatch = Regex.Match(weSellStr, @"[\d,\.]+");

                                if (weBuyMatch.Success && weSellMatch.Success)
                                {
                                    try
                                    {
                                        decimal weBuy = decimal.Parse(weBuyMatch.Value, System.Globalization.NumberStyles.Any);
                                        decimal weSell = decimal.Parse(weSellMatch.Value, System.Globalization.NumberStyles.Any);

                                        // For MS Gold, use the longer format terms
                                        string normalizedDetail = GetMSGoldDetailName(detail);

                                        ourRatesData.Rows.Add(normalizedDetail, weBuy, weSell, "OurRates");
                                        statusTextBox.Text += Environment.NewLine + $"Extracted OurRates: {normalizedDetail}, Buy: {weBuy}, Sell: {weSell}";
                                        foundOurRatesData = true;
                                    }
                                    catch (Exception ex)
                                    {
                                        statusTextBox.Text += Environment.NewLine + $"Error parsing values: {ex.Message}";
                                    }
                                }
                                else
                                {
                                    statusTextBox.Text += Environment.NewLine + $"Failed to extract numeric values";
                                }
                            }
                        }
                    }
                    // Look for customer sell table (usually has 2 columns)
                    else if (tableText.Contains("WE BUY") && !tableText.Contains("WE SELL"))
                    {
                        statusTextBox.Text += Environment.NewLine + $"Found CustomerSell table (index {i})";

                        var rows = table.FindElements(By.TagName("tr"));
                        statusTextBox.Text += Environment.NewLine + $"Table has {rows.Count} rows";

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count >= 2)
                            {
                                string detail = cells[0].Text.Trim();

                                // Skip header rows
                                if (detail.Contains("DETAILS") || detail.Contains("WE BUY") || string.IsNullOrWhiteSpace(detail))
                                    continue;

                                string weBuyStr = cells[1].Text.Trim();

                                // Log raw values
                                statusTextBox.Text += Environment.NewLine + $"Raw CustomerSell values - Detail: '{detail}', Buy: '{weBuyStr}'";

                                // Extract numeric values
                                var weBuyMatch = Regex.Match(weBuyStr, @"[\d,\.]+");

                                if (weBuyMatch.Success)
                                {
                                    try
                                    {
                                        decimal weBuy = decimal.Parse(weBuyMatch.Value, System.Globalization.NumberStyles.Any);

                                        // Parse gold purity
                                        if (detail.Contains("999") || detail.Contains("916") ||
                                            detail.Contains("835") || detail.Contains("750") ||
                                            detail.Contains("375"))
                                        {
                                            string purity = "";
                                            if (detail.Contains("999.9"))
                                                purity = "999.9";
                                            else if (detail.Contains("999"))
                                                purity = "999";
                                            else if (detail.Contains("916"))
                                                purity = "916";
                                            else if (detail.Contains("835"))
                                                purity = "835";
                                            else if (detail.Contains("750"))
                                                purity = "750";
                                            else if (detail.Contains("375"))
                                                purity = "375";

                                            // Format: "999 MYR / Gram"
                                            string normalizedDetail = $"{purity} MYR / Gram";

                                            customerSellData.Rows.Add(normalizedDetail, weBuy, null, "CustomerSell");
                                            statusTextBox.Text += Environment.NewLine + $"Extracted CustomerSell: {normalizedDetail}, Buy: {weBuy}";
                                            foundCustomerSellData = true;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        statusTextBox.Text += Environment.NewLine + $"Error parsing values: {ex.Message}";
                                    }
                                }
                                else
                                {
                                    statusTextBox.Text += Environment.NewLine + $"Failed to extract numeric values";
                                }
                            }
                        }
                    }
                }

                // If we found at least one type of data, save it
                if (foundOurRatesData || foundCustomerSellData)
                {
                    await ClearExistingCompanyDataAsync(currentCompanyName);
                    await SaveMSGoldDataAsync(ourRatesData, customerSellData);
                    return true;
                }
                else
                {
                    statusTextBox.Text += Environment.NewLine + "No data was extracted.";
                    return false;
                }
            }
            catch (Exception ex)
            {
                statusTextBox.Text += Environment.NewLine + $"Error processing MS Gold page: {ex.Message}";
                return false;
            }
        }

        private string GetMSGoldDetailName(string detail)
        {
            // Convert shortened format to MS Gold's longer format
            if (detail.Contains("USD") && detail.Contains("oz"))
                return "999.9 Gold USD / Oz";
            else if (detail.Contains("MYR") && detail.Contains("kg"))
                return "999.9 Gold MYR / KG";
            else if (detail.Contains("MYR") && detail.Contains("tael"))
                return "999.9 Gold MYR / Tael";
            else if (detail.Contains("MYR") && detail.Contains("g"))
                return "999.9 Gold MYR / Gram";
            else if (detail.Contains("USD") && detail.Contains("MYR"))
                return "USD / MYR";
            else
                return detail; // Return as-is if no match
        }

        private async Task ClearExistingCompanyDataAsync(string companyName)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();

                    using (SqlCommand command = new SqlCommand())
                    {
                        command.Connection = connection;

                        if (companyName == "TTTBullion")
                        {
                            statusTextBox.Text += Environment.NewLine + "Clearing existing TTTBullion data...";
                            command.CommandText = "DELETE FROM TTTBullion_GoldRates";
                        }
                        else if (companyName == "MSGold")
                        {
                            statusTextBox.Text += Environment.NewLine + "Clearing existing MSGold data...";
                            command.CommandText = "DELETE FROM MSGold_OurRates; DELETE FROM MSGold_CustomerSell;";
                        }

                        await command.ExecuteNonQueryAsync();
                        statusTextBox.Text += Environment.NewLine + $"Existing {companyName} data cleared successfully.";
                    }
                }
            }
            catch (Exception ex)
            {
                statusTextBox.Text += Environment.NewLine + $"Error clearing existing data: {ex.Message}";
                throw;
            }
        }

        private async Task SaveTTTBullionDataAsync(DataTable data)
        {
            statusTextBox.Text += Environment.NewLine + "Saving TTT Bullion data to database...";

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();
                    statusTextBox.Text += Environment.NewLine + "Connected to database.";

                    int rowsSaved = 0;

                    foreach (DataRow row in data.Rows)
                    {
                        string detailName = row["DetailName"].ToString();
                        decimal weBuy = (decimal)row["WeBuy"];
                        decimal? weSell = row["WeSell"] != DBNull.Value ? (decimal?)row["WeSell"] : null;

                        
                        using (SqlCommand command = new SqlCommand(
                            "INSERT INTO TTTBullion_GoldRates (GoldRate_DetailName, GoldRate_WeBuy, GoldRate_WeSell, GoldRate_LastUpdated) " +
                            "VALUES (@DetailName, @WeBuy, @WeSell, GETDATE())", connection))
                        {
                            command.Parameters.AddWithValue("@DetailName", detailName);
                            command.Parameters.AddWithValue("@WeBuy", weBuy);
                            command.Parameters.AddWithValue("@WeSell", weSell.HasValue ? (object)weSell.Value : DBNull.Value);

                            await command.ExecuteNonQueryAsync();
                            rowsSaved++;
                            statusTextBox.Text += Environment.NewLine + $"Saved data for {detailName}";
                        }
                    }

                    statusTextBox.Text += Environment.NewLine + $"Successfully saved {rowsSaved} records for TTT Bullion.";
                }
            }
            catch (Exception ex)
            {
                statusTextBox.Text += Environment.NewLine + $"Database error: {ex.Message}";
                throw;
            }
        }

        private async Task SaveMSGoldDataAsync(DataTable ourRatesData, DataTable customerSellData)
        {
            statusTextBox.Text += Environment.NewLine + "Saving MS Gold data to database...";

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();
                    statusTextBox.Text += Environment.NewLine + "Connected to database.";

                    // Save OurRates data
                    int ourRatesSaved = 0;
                    foreach (DataRow row in ourRatesData.Rows)
                    {
                        string detailName = row["DetailName"].ToString();
                        decimal weBuy = (decimal)row["WeBuy"];
                        decimal? weSell = row["WeSell"] != DBNull.Value ? (decimal?)row["WeSell"] : null;

                        using (SqlCommand command = new SqlCommand("INSERT INTO MSGold_OurRates (Rate_DetailName, Rate_WeBuy, Rate_WeSell, Rate_LastUpdated) VALUES (@DetailName, @WeBuy, @WeSell, GETDATE())", connection))
                        {
                            command.Parameters.AddWithValue("@DetailName", detailName);
                            command.Parameters.AddWithValue("@WeBuy", weBuy);
                            command.Parameters.AddWithValue("@WeSell", weSell.HasValue ? (object)weSell.Value : DBNull.Value);

                            await command.ExecuteNonQueryAsync();
                            ourRatesSaved++;
                            statusTextBox.Text += Environment.NewLine + $"Saved OurRates data for {detailName}";
                        }
                    }

                    // Save CustomerSell data
                    int customerSellSaved = 0;
                    foreach (DataRow row in customerSellData.Rows)
                    {
                        string detailName = row["DetailName"].ToString();
                        decimal weBuy = (decimal)row["WeBuy"];

                        using (SqlCommand command = new SqlCommand("INSERT INTO MSGold_CustomerSell (Customer_DetailName, Customer_WeBuy, Customer_LastUpdated) VALUES (@DetailName, @WeBuy, GETDATE())", connection))
                        {
                            command.Parameters.AddWithValue("@DetailName", detailName);
                            command.Parameters.AddWithValue("@WeBuy", weBuy);

                            await command.ExecuteNonQueryAsync();
                            customerSellSaved++;
                            statusTextBox.Text += Environment.NewLine + $"Saved CustomerSell data for {detailName}";
                        }
                    }

                    statusTextBox.Text += Environment.NewLine + $"Successfully saved {ourRatesSaved} OurRates records and {customerSellSaved} CustomerSell records for MS Gold.";
                }
            }
            catch (Exception ex)
            {
                statusTextBox.Text += Environment.NewLine + $"Database error: {ex.Message}";
                throw;
            }
        }
    }
}