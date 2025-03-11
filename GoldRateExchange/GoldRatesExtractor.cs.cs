using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Configuration;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace GoldRatesExtractor
{
    public class GoldRatesExtractor
    {
        private string connectionString;
        private string currentCompanyName;
        private ChromeDriver driver;
        private string logFilePath;
        private string errorLogFilePath;
        private int websiteOption;

        public GoldRatesExtractor()
        {
            // Create log directory if it doesn't exist
            string logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
            if (!Directory.Exists(logDirectory))
            {
                Directory.CreateDirectory(logDirectory);
            }

            // Set log file paths
            logFilePath = Path.Combine(logDirectory, $"GoldRates_Log_{DateTime.Now:yyyyMMdd}.txt");
            errorLogFilePath = Path.Combine(logDirectory, $"GoldRates_Error_{DateTime.Now:yyyyMMdd}.txt");
        }

        public async Task StartAsync()
        {
            LogToFile("Starting Gold Rates Extractor");

            try
            {
                // Load configuration
                LoadConfiguration();
                InitializeSelenium();

                // Run extraction process
                await ExtractGoldRatesAsync();
            }
            catch (Exception ex)
            {
                LogError($"Error during startup: {ex.Message}");
                LogError($"Stack trace: {ex.StackTrace}");
            }
            finally
            {
                // Clean up
                if (driver != null)
                {
                    driver.Quit();
                    driver.Dispose();
                }

                LogToFile("Gold Rates Extractor completed");
            }
        }

        private void InitializeSelenium()
        {
            try
            {
                var options = new ChromeOptions();
                options.AddArgument("--headless"); // Run Chrome in headless mode (no UI)
                options.AddArgument("--disable-gpu");
                options.AddArgument("--log-level=3");
                ChromeDriverService service = ChromeDriverService.CreateDefaultService();
                service.HideCommandPromptWindow = true;

                driver = new ChromeDriver(service, options);
                LogToFile("Selenium initialized successfully");
            }
            catch (Exception ex)
            {
                LogError($"Error initializing Selenium: {ex.Message}");
                throw; // Re-throw to handle at higher level
            }
        }

        private void LoadConfiguration()
        {
            try
            {
                // Load connection string from app.config
                connectionString = ConfigurationManager.AppSettings["ConnectionString"];

                // Get the website option from app.config
                string websiteOptionStr = ConfigurationManager.AppSettings["WebsiteOption"];
                if (!int.TryParse(websiteOptionStr, out websiteOption))
                {
                    // Default to option 1 if parsing fails
                    websiteOption = 1;
                    LogToFile("WebsiteOption not specified or invalid in app.config. Defaulting to 1 (TTTBullion)");
                }
                else
                {
                    LogToFile($"Using WebsiteOption {websiteOption} from app.config");
                }

                LogToFile("Configuration loaded successfully");
            }
            catch (Exception ex)
            {
                LogError($"Error loading configuration: {ex.Message}");
                throw; // Re-throw to handle at higher level
            }
        }

        private async Task ExtractGoldRatesAsync()
        {
            LogToFile("Starting gold rates extraction process");

            try
            {
                if (driver == null)
                {
                    LogError("Selenium WebDriver is not initialized. Re-initializing...");
                    InitializeSelenium();
                }

                // Determine which website to use based on configuration
                string url;

                if (websiteOption == 1)
                {
                    url = ConfigurationManager.AppSettings["UrlOption1"];
                    currentCompanyName = ConfigurationManager.AppSettings["CompanyName1"];
                    LogToFile($"Using Option 1: {currentCompanyName} at {url}");

                    // Navigate to URL and process
                    await LoadWebpageAsync(url);
                    bool success = await ProcessTTTBullionPageAsync();

                    if (success)
                    {
                        LogToFile($"Successfully extracted and saved data from {currentCompanyName}");
                    }
                    else
                    {
                        LogError($"Failed to extract data from {currentCompanyName}");
                    }
                }
                else
                {
                    url = ConfigurationManager.AppSettings["UrlOption2"];
                    currentCompanyName = ConfigurationManager.AppSettings["CompanyName2"];
                    LogToFile($"Using Option 2: {currentCompanyName} at {url}");

                    // Navigate to URL and process
                    await LoadWebpageAsync(url);
                    bool success = await ProcessMSGoldPageAsync();

                    if (success)
                    {
                        LogToFile($"Successfully extracted and saved data from {currentCompanyName}");
                    }
                    else
                    {
                        LogError($"Failed to extract data from {currentCompanyName}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogError($"Critical error in extraction process: {ex.Message}");
                LogError($"Stack trace: {ex.StackTrace}");
            }
        }

        private async Task LoadWebpageAsync(string url)
        {
            LogToFile($"Loading webpage: {url}");

            await Task.Run(() => {
                try
                {
                    driver.Navigate().GoToUrl(url);

                    // Wait for page to load
                    var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
                    wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));

                    // Additional wait for any AJAX calls to complete
                    System.Threading.Thread.Sleep(3000);
                    LogToFile("Page loaded successfully");
                }
                catch (Exception ex)
                {
                    LogError($"Error loading webpage: {ex.Message}");
                    throw;
                }
            });
        }

        private async Task<bool> ProcessTTTBullionPageAsync()
        {
            LogToFile("Processing TTT Bullion page...");

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
                LogToFile($"Found {tableCount} tables on the page");

                // We only want the first table (Gold rates), not the silver rates
                if (tableCount > 0)
                {
                    var goldTable = tableElements[0]; // First table contains gold rates
                    string tableText = goldTable.Text;

                    // Check if this is indeed the gold table
                    if (tableText.Contains("Gold") && !tableText.Contains("Silver"))
                    {
                        LogToFile("Processing Gold rates table");

                        // Get all rows in the table
                        var rows = goldTable.FindElements(By.TagName("tr"));
                        int rowCount = rows.Count;
                        LogToFile($"Gold table has {rowCount} rows");

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
                                LogToFile($"Gold data - Detail: '{detail}', Buy: '{weBuyStr}', Sell: '{weSellStr}'");

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
                                            LogToFile($"Extracted gold rate: {detail}, Buy: {weBuy}, Sell: {weSell}");
                                        }
                                        catch (Exception ex)
                                        {
                                            LogError($"Error parsing values: {ex.Message}");
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        LogToFile("First table doesn't appear to be the Gold rates table. Checking all tables...");

                        // If the first table wasn't the gold table, search through all tables
                        bool foundGoldTable = false;

                        for (int tableIndex = 0; tableIndex < tableCount; tableIndex++)
                        {
                            var table = tableElements[tableIndex];
                            string tblText = table.Text;

                            // Check if this table contains gold rates
                            if (tblText.Contains("Gold") && !tblText.Contains("Silver"))
                            {
                                LogToFile($"Found Gold rates in table {tableIndex}");

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

                                        LogToFile($"Gold data - Detail: '{detail}', Buy: '{weBuyStr}', Sell: '{weSellStr}'");

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
                                                    LogToFile($"Extracted gold rate: {detail}, Buy: {weBuy}, Sell: {weSell}");
                                                    foundGoldTable = true;
                                                }
                                                catch (Exception ex)
                                                {
                                                    LogError($"Error parsing values: {ex.Message}");
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
                            LogToFile("Could not identify a Gold rates table on the page.");
                            return false;
                        }
                    }
                }
                else
                {
                    LogToFile("No tables found on the page.");
                    return false;
                }

                // Check if we found any data
                if (extractedData.Rows.Count == 0)
                {
                    LogToFile("No gold rate data extracted.");
                    return false;
                }

                LogToFile($"Extracted {extractedData.Rows.Count} gold rates.");

                // Clear existing data and save to database
                await ClearExistingCompanyDataAsync(currentCompanyName);
                await SaveTTTBullionDataAsync(extractedData);
                return true;
            }
            catch (Exception ex)
            {
                LogError($"Error processing TTT Bullion page: {ex.Message}");
                LogError($"Stack trace: {ex.StackTrace}");
                return false;
            }
        }

        private async Task<bool> ProcessMSGoldPageAsync()
        {
            LogToFile("Processing MS Gold page...");

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
                LogToFile($"Found {tableElements.Count} tables on the page");

                bool foundOurRatesData = false;
                bool foundCustomerSellData = false;

                // Process each table
                for (int i = 0; i < tableElements.Count; i++)
                {
                    var table = tableElements[i];
                    string tableText = table.Text;
                    LogToFile($"Table {i} text: {tableText.Substring(0, Math.Min(100, tableText.Length))}...");

                    // Look for our rates table (usually has 3 columns including Buy and Sell)
                    if (tableText.Contains("WE BUY") && tableText.Contains("WE SELL"))
                    {
                        LogToFile($"Found OurRates table (index {i})");

                        var rows = table.FindElements(By.TagName("tr"));
                        LogToFile($"Table has {rows.Count} rows");

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
                                LogToFile($"Raw values - Detail: '{detail}', Buy: '{weBuyStr}', Sell: '{weSellStr}'");

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
                                        LogToFile($"Extracted OurRates: {normalizedDetail}, Buy: {weBuy}, Sell: {weSell}");
                                        foundOurRatesData = true;
                                    }
                                    catch (Exception ex)
                                    {
                                        LogError($"Error parsing values: {ex.Message}");
                                    }
                                }
                                else
                                {
                                    LogToFile($"Failed to extract numeric values");
                                }
                            }
                        }
                    }
                    // Look for customer sell table (usually has 2 columns)
                    else if (tableText.Contains("WE BUY") && !tableText.Contains("WE SELL"))
                    {
                        LogToFile($"Found CustomerSell table (index {i})");

                        var rows = table.FindElements(By.TagName("tr"));
                        LogToFile($"Table has {rows.Count} rows");

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
                                LogToFile($"Raw CustomerSell values - Detail: '{detail}', Buy: '{weBuyStr}'");

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
                                            LogToFile($"Extracted CustomerSell: {normalizedDetail}, Buy: {weBuy}");
                                            foundCustomerSellData = true;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        LogError($"Error parsing values: {ex.Message}");
                                    }
                                }
                                else
                                {
                                    LogToFile($"Failed to extract numeric values");
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
                    LogToFile("No data was extracted.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                LogError($"Error processing MS Gold page: {ex.Message}");
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
                            LogToFile("Clearing existing TTTBullion data...");
                            command.CommandText = "DELETE FROM TTTBullion_GoldRates";
                        }
                        else if (companyName == "MSGold")
                        {
                            LogToFile("Clearing existing MSGold data...");
                            command.CommandText = "DELETE FROM MSGold_OurRates; DELETE FROM MSGold_CustomerSell;";
                        }

                        await command.ExecuteNonQueryAsync();
                        LogToFile($"Existing {companyName} data cleared successfully.");
                    }
                }
            }
            catch (Exception ex)
            {
                LogError($"Error clearing existing data: {ex.Message}");
                throw;
            }
        }

        private async Task SaveTTTBullionDataAsync(DataTable data)
        {
            LogToFile("Saving TTT Bullion data to database...");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();
                    LogToFile("Connected to database.");

                    int rowsSaved = 0;

                    foreach (DataRow row in data.Rows)
                    {
                        string detailName = row["DetailName"].ToString();
                        decimal weBuy = (decimal)row["WeBuy"];
                        decimal? weSell = row["WeSell"] != DBNull.Value ? (decimal?)row["WeSell"] : null;

                        // Use the stored procedure instead of direct SQL
                        using (SqlCommand command = new SqlCommand("sp_goldRates_upsert", connection))
                        {
                            command.CommandType = CommandType.StoredProcedure;

                            // Add parameters to call the stored procedure
                            command.Parameters.AddWithValue("@CompanyName", currentCompanyName);
                            command.Parameters.AddWithValue("@TableType", "OurRates"); // Always OurRates for gold rates
                            command.Parameters.AddWithValue("@DetailName", detailName);
                            command.Parameters.AddWithValue("@WeBuy", weBuy);
                            command.Parameters.AddWithValue("@WeSell", weSell.HasValue ? (object)weSell.Value : DBNull.Value);

                            await command.ExecuteNonQueryAsync();
                            rowsSaved++;
                            LogToFile($"Saved gold rate for {detailName}");
                        }
                    }

                    LogToFile($"Successfully saved {rowsSaved} gold rates records for TTT Bullion.");
                }
            }
            catch (Exception ex)
            {
                LogError($"Database error: {ex.Message}");
                throw;
            }
        }

        private async Task SaveMSGoldDataAsync(DataTable ourRatesData, DataTable customerSellData)
        {
            LogToFile("Saving MS Gold data to database...");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();
                    LogToFile("Connected to database.");

                    // Save OurRates data
                    int ourRatesSaved = 0;
                    foreach (DataRow row in ourRatesData.Rows)
                    {
                        string detailName = row["DetailName"].ToString();
                        decimal weBuy = (decimal)row["WeBuy"];
                        decimal? weSell = row["WeSell"] != DBNull.Value ? (decimal?)row["WeSell"] : null;

                        // Use the stored procedure instead of direct SQL
                        using (SqlCommand command = new SqlCommand("sp_goldRates_upsert", connection))
                        {
                            command.CommandType = CommandType.StoredProcedure;

                            // Add parameters
                            command.Parameters.AddWithValue("@CompanyName", currentCompanyName);
                            command.Parameters.AddWithValue("@TableType", "OurRates");
                            command.Parameters.AddWithValue("@DetailName", detailName);
                            command.Parameters.AddWithValue("@WeBuy", weBuy);
                            command.Parameters.AddWithValue("@WeSell", weSell.HasValue ? (object)weSell.Value : DBNull.Value);

                            await command.ExecuteNonQueryAsync();
                            ourRatesSaved++;
                            LogToFile($"Saved OurRates data for {detailName}");
                        }
                    }

                    // Save CustomerSell data
                    int customerSellSaved = 0;
                    foreach (DataRow row in customerSellData.Rows)
                    {
                        string detailName = row["DetailName"].ToString();
                        decimal weBuy = (decimal)row["WeBuy"];

                        // Use the stored procedure instead of direct SQL
                        using (SqlCommand command = new SqlCommand("sp_goldRates_upsert", connection))
                        {
                            command.CommandType = CommandType.StoredProcedure;

                            // Add parameters
                            command.Parameters.AddWithValue("@CompanyName", currentCompanyName);
                            command.Parameters.AddWithValue("@TableType", "CustomerSell");
                            command.Parameters.AddWithValue("@DetailName", detailName);
                            command.Parameters.AddWithValue("@WeBuy", weBuy);
                            command.Parameters.AddWithValue("@WeSell", DBNull.Value); // CustomerSell doesn't have WeSell values

                            await command.ExecuteNonQueryAsync();
                            customerSellSaved++;
                            LogToFile($"Saved CustomerSell data for {detailName}");
                        }
                    }

                    LogToFile($"Successfully saved {ourRatesSaved} OurRates records and {customerSellSaved} CustomerSell records for MS Gold.");
                }
            }
            catch (Exception ex)
            {
                LogError($"Database error: {ex.Message}");
                throw;
            }
        }

        private void LogToFile(string message)
        {
            try
            {
                string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
                File.AppendAllText(logFilePath, logMessage + Environment.NewLine);

                // Also output to console for visibility when running manually
                Console.WriteLine(logMessage);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write to log file: {ex.Message}. Original message: {message}");
            }
        }

        private void LogError(string errorMessage)
        {
            try
            {
                string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - ERROR: {errorMessage}";
                File.AppendAllText(errorLogFilePath, logMessage + Environment.NewLine);

                // Also log to the regular log file
                LogToFile($"ERROR: {errorMessage}");

                // Output to console in red
                ConsoleColor originalColor = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(logMessage);
                Console.ForegroundColor = originalColor;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write to error log file: {ex.Message}. Original error: {errorMessage}");
            }
        }
    }
}