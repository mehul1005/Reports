using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ComponentFactory.Krypton.Toolkit;
using System.Threading;
using System.Data.OleDb;
using ExcelDataReader;

namespace ReportsNewD
{
    public partial class Form1 : KryptonForm
    {
        // Declare class-level fields for excelWorkbook and excelApp
        private Microsoft.Office.Interop.Excel.Workbook excelWorkbook;
        private Microsoft.Office.Interop.Excel.Application excelApp;

        public Form1()
        {
            InitializeComponent();
            //this.MaximizeBox = false;
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtFolder.Text = dialog.SelectedPath;
                    // Set the value of txtExcelFilePath.Text with the same path as txtFolder.Text but with .xlsx extension
                    string excelFileName = Path.GetFileName(dialog.SelectedPath) + ".xlsx";
                    string excelFilePath = Path.Combine(dialog.SelectedPath, excelFileName);
                    txtExcelFilePath.Text = excelFilePath;
                }
            }
        }

        private void txtFolder_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFolder.Text))
            {
                // Disable the convert and convert-and-sum buttons
                btnConvert.Enabled = false;
                btnConvertAndSum.Enabled = false;
                // Clear the file path in txtExcelFilePath
                txtExcelFilePath.Text = string.Empty;
                // Display a message to prompt the user to select a CSV folder
                MessageBox.Show("Please select a CSV folder to continue.", "Folder Not Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                // Enable the convert and convert-and-sum buttons
                btnConvert.Enabled = true;
                btnConvertAndSum.Enabled = true;
                // Set the value of txtExcelFilePath.Text with the same path as txtFolder.Text but with .xlsx extension
                string folderPath = txtFolder.Text;
                string folderName = Path.GetFileName(folderPath);
                string excelFileName = folderName + ".xlsx";
                string excelFilePath = Path.Combine(folderPath, excelFileName);
                txtExcelFilePath.Text = excelFilePath;
            }
        }

        private void btnExcelBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                dialog.FilterIndex = 1;
                dialog.RestoreDirectory = true;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtExcelFilePath.Text = dialog.FileName;
                }
            }
        }

        // Write the data to the Excel worksheet
        Dictionary<string, double> uniqueNamesAndTotalYields = new Dictionary<string, double>();

        // Create a dictionary to store the uniqueRoughWeights for each Stone Name
        HashSet<string> uniqueRoughWeights = new HashSet<string>();

        // Create a dictionary to store the count of each unique stone name
        Dictionary<string, int> stoneNameCounts = new Dictionary<string, int>();

        // Map numerical values in "Part" column to letters
        Dictionary<int, string> partLetters = new Dictionary<int, string>
        {
            { 1, "A" }, { 2, "B" }, { 3, "C" }, { 4, "D" }, { 5, "E" }, { 6, "F" }, { 7, "G" }, { 8, "H" }, { 9, "I" }, { 10, "J" },
            { 11, "K" }, { 12, "L" }, { 13, "M" }, { 14, "N" }, { 15, "O" }, { 16, "P" }, { 17, "Q" }, { 18, "R" }, { 19, "S" },
            { 20, "T" }, { 21, "U" }, { 22, "V" }, { 23, "W" }, { 24, "X" }, { 25, "Y" },
        };

        private async void btnConvert_Click(object sender, EventArgs e)
        {
            await btnConvert_ClickAsync(sender, e);
        }

        private async Task btnConvert_ClickAsync(object sender, EventArgs e)
        {
            await Task.Run(() =>
            {
                Microsoft.Office.Interop.Excel.Application excelApp = null;

                try
                {
                    // Create a new Excel application instance for the background task
                    excelApp = new Microsoft.Office.Interop.Excel.Application() { Visible = false };

                    // Create a new Excel workbook
                    Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();

                    // Create a new Excel worksheet
                    Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = excelWorkbook.Sheets.Add();
                    excelWorksheet.Name = "Data";

                    // Get the path of the directory containing the CSV files
                    string csvFolderPath = txtFolder.Text;

                    // Check if csvFolderPath is blank
                    if (string.IsNullOrEmpty(csvFolderPath))
                    {
                        MessageBox.Show("Please Select CSV Folder to Continue.", "Folder Not Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Get the files with a .csv extension in the specified directory
                    var csvFiles = Directory.GetFiles(csvFolderPath, "*.csv");

                    // Write the column headers to the first row of the worksheet
                    List<string> columnNames = new List<string>
                    {
                        "Stone Name", "RoughWeight", "TotalYield(%)", "Part", "Shape", "PartWT", "Pw", "PartUsage",
                        "Sieve", "Cut", "Color", "Clarity", "Fls", "Dolar", "Discount", "W", "L", "Ratio",
                        "SawName", "SWmm", "SawWeightCarat", "GradingInstitute", "CUserName", "Day",
                        "Month", "Year", "Time"
                    };

                    for (int j = 0; j < columnNames.Count; j++)
                    {
                        excelWorksheet.Cells[1, j + 1] = columnNames[j];
                    }

                    // Specify the path for the "Missing Entry" folder
                    string missingEntryFolderPath = Path.Combine(csvFolderPath, "Missing Entry");
                    Directory.CreateDirectory(missingEntryFolderPath); // Create the folder if it doesn't exist

                    // Create a list to store the remaining CSV files that haven't been moved to the "Missing Entry" folder
                    var remainingCsvFiles = new List<string>();
                    int missingEntryCount = 0; // Counter for missing entries

                    // Loop through each CSV file
                    foreach (var csvFile in csvFiles)
                    {
                        // Read all lines from the current CSV file
                        string[] lines = File.ReadAllLines(csvFile);

                        // Check if the "Clarity" column is blank in any row
                        bool hasBlankClarity = lines.Skip(1).Any(line => line.Split(',')[35].Trim() == "");

                        if (hasBlankClarity)
                        {
                            // Move the CSV file to the "Missing Entry" folder
                            string fileName = Path.GetFileName(csvFile);
                            string newFilePath = Path.Combine(missingEntryFolderPath, fileName);
                            File.Move(csvFile, newFilePath);
                            missingEntryCount++;
                        }
                        else
                        {
                            // Add the CSV file to the remainingCsvFiles list
                            remainingCsvFiles.Add(csvFile);
                        }
                    }

                    // Calculate the total number of rows for non-missing CSV files
                    int totalRows = 0;
                    foreach (var csvFile in remainingCsvFiles)
                    {
                        int fileRowCount = File.ReadLines(csvFile).Count();
                        totalRows += fileRowCount;
                    }

                    // Loop through each CSV file
                    int currentRow = 2; // Start at row 2 to leave the first row for the headers

                    for (int i = 0; i < remainingCsvFiles.Count; i++)
                    {
                        string csvFile = remainingCsvFiles[i];

                        System.Data.DataTable csvData = new System.Data.DataTable();
                        using (var csvReader = new StreamReader(csvFile))
                        {
                            string[] headers = csvReader.ReadLine().Split(',');
                            foreach (string header in headers)
                            {
                                csvData.Columns.Add(header);
                            }

                            while (!csvReader.EndOfStream)
                            {
                                string[] rows = csvReader.ReadLine().Split(',');
                                DataRow dr = csvData.NewRow();
                                for (int j = 0; j < headers.Length; j++)
                                {
                                    dr[j] = rows[j];
                                }
                                csvData.Rows.Add(dr);
                            }
                        }

                        double totalPw = csvData.AsEnumerable().Sum(r => double.Parse(r.Field<string>("Pw")));
                        foreach (DataRow row in csvData.Rows)
                        {
                            string stoneName = row.Field<string>("Stone Name");
                            if (uniqueNamesAndTotalYields.ContainsKey(stoneName))
                            {
                                excelWorksheet.Cells[currentRow, 1].Value = "";
                                excelWorksheet.Cells[currentRow, columnNames.IndexOf("TotalYield(%)") + 1].Value = "";
                            }
                            else
                            {
                                double totalYield = 0;
                                double roughWeight;
                                if (double.TryParse(row["RoughWeight"].ToString(), out roughWeight) && roughWeight > 0)
                                {
                                    totalYield = totalPw / roughWeight * 100;
                                }

                                excelWorksheet.Cells[currentRow, 1].Value = stoneName;
                                excelWorksheet.Cells[currentRow, columnNames.IndexOf("TotalYield(%)") + 1].Value = totalYield.ToString("#0.00") + "%";
                                uniqueNamesAndTotalYields[stoneName] = totalYield;
                            }

                            int part;
                            string partLetter;
                            if (int.TryParse(row["Part"].ToString(), out part))
                            {
                                partLetter = partLetters.ContainsKey(part) ? partLetters[part] : "";
                            }
                            else
                            {
                                partLetter = "T"; // Set a different default value for partLetter
                            }

                            for (int j = 1; j < columnNames.Count; j++)
                            {
                                if (columnNames[j] == "PartUsage")
                                {
                                    double pw = 0, partWt = 0;
                                    double.TryParse(row["Pw"].ToString(), out pw);
                                    double.TryParse(row["PartWT"].ToString(), out partWt);
                                    double partUsage = pw / partWt * 100;
                                    excelWorksheet.Cells[currentRow, j + 1].Value = partUsage.ToString("#0.00") + "%";
                                }
                                else if (columnNames[j] == "TotalYield(%)")
                                {
                                    if (!uniqueNamesAndTotalYields.ContainsKey(stoneName))
                                    {
                                        double totalYield = 0;
                                        double.TryParse(row["TotalYield"].ToString(), out totalYield);
                                        excelWorksheet.Cells[currentRow, j + 1].Value = totalYield.ToString("#0.00") + "%";
                                    }
                                }
                                else if (columnNames[j] == "Part")
                                {
                                    excelWorksheet.Cells[currentRow, j + 1].Value = partLetter;
                                }
                                else if (columnNames[j] == "RoughWeight")
                                {
                                    if (uniqueRoughWeights.Contains(stoneName))
                                    {
                                        excelWorksheet.Cells[currentRow, j + 1].Value = "";
                                    }
                                    else
                                    {
                                        excelWorksheet.Cells[currentRow, j + 1].Value = row[columnNames[j]].ToString();
                                        uniqueRoughWeights.Add(stoneName);
                                    }
                                }
                                else
                                {
                                    excelWorksheet.Cells[currentRow, j + 1].Value = row[columnNames[j]].ToString();
                                }
                            }
                            currentRow++;
                        }

                        // Update the progress bar and label text on the UI thread
                        int progressPercentage = (int)Math.Round((double)(i + 1) / remainingCsvFiles.Count * 100);
                        Invoke(new System.Action(() =>
                        {
                            progressBar1.Value = progressPercentage;
                            lblProgress.Text = $"PROCESSING {i + 1} OF {remainingCsvFiles.Count} CSV FILES ({progressPercentage}%)";
                            lblMissingEntryCount.Text = $"Missing Entry : {missingEntryCount}";
                        }));
                    }

                    // Get the name of the current folder
                    string folderName = new DirectoryInfo(csvFolderPath).Name;

                    // Save the Excel workbook with the same name as the folder
                    string excelFilePath = Path.Combine(csvFolderPath, folderName + ".xlsx");
                    excelWorkbook.SaveAs(excelFilePath);
                    // Introduce a delay before closing the Excel workbook

                    Thread.Sleep(1000); // Adjust the delay as needed

                    // Close the Excel application and release resources
                    excelWorkbook.Close(SaveChanges: false);
                    excelApp.Quit();

                    // Create a new folder for the processed CSV files
                    string newFolderPath = Path.Combine(csvFolderPath, "Done Report");
                    Directory.CreateDirectory(newFolderPath);

                    // Move all the CSV files to the new folder
                    foreach (var csvFile in csvFiles)
                    {
                        string newFilePath = Path.Combine(newFolderPath, Path.GetFileName(csvFile));
                        if (File.Exists(csvFile))
                        {
                            File.Move(csvFile, newFilePath);
                        }
                    }
                }
                finally
                {
                    // Release COM objects and force garbage collection
                    if (excelWorkbook != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                        excelWorkbook = null;
                    }
                    if (excelApp != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                        excelApp = null;
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            });
        }

        #region Button Summary Convert Click

        private async void btnSumConvert_Click(object sender, EventArgs e)
        {
            await btnSumConvert_ClickAsync(sender, e);
        }

        private async Task btnSumConvert_ClickAsync(object sender, EventArgs e)
        {
            // Declare Excel objects
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbook wb = null;
            Microsoft.Office.Interop.Excel._Worksheet sheet = null;
            Microsoft.Office.Interop.Excel.Range range = null;

            await Task.Run(() =>
            {
                string filePath = txtExcelFilePath.Text;

                // Check if csvFolderPath is blank
                if (string.IsNullOrEmpty(filePath))
                {
                    MessageBox.Show("Please Select Excel File to Continue.", "Excel File Not Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                try
                {
                    // Create a new Excel application instance
                    excel = new Microsoft.Office.Interop.Excel.Application();
                    wb = excel.Workbooks.Open(filePath);
                    sheet = wb.Sheets[1];
                    range = sheet.UsedRange;

                    // Insert new columns
                    range.Columns[3].Insert();
                    range.Columns[4].Insert();
                    range.Columns[6].Insert();
                    range.Columns[7].Insert();

                    // Update headers of new columns
                    range.Cells[1, 3].Value2 = "TotalPartWeight";
                    range.Cells[1, 4].Value2 = "TotalPolishWeight";
                    range.Cells[1, 6].Value2 = "TotalValue($)";
                    range.Cells[1, 7].Value2 = "PolishSrNo";

                    Dictionary<string, int> clarityCounts = new Dictionary<string, int>();

                    int rowCount = range.Rows.Count;
                    int columnCount = range.Columns.Count;

                    string currentStoneName = "";
                    double currentTotalPartWeight = 0;
                    double currentTotalPolishWeight = 0;
                    double currentTotalValue = 0;

                    // initialize the last stone name row to 1
                    int lastStoneNameRow = 2;

                    // Exclude the header row
                    int totalRows = rowCount - 1;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        string stoneName = (range.Cells[i, 1].Value2 != null) ? range.Cells[i, 1].Value2.ToString() : "";

                        if (!string.IsNullOrEmpty(stoneName))
                        {
                            // If we encounter a new stone name, write the totals for the previous stone name
                            if (stoneName != currentStoneName && i > 2)
                            {
                                // Update the last stone name row data with the sum result
                                range.Cells[lastStoneNameRow, 3].Value2 = currentTotalPartWeight;
                                range.Cells[lastStoneNameRow, 4].Value2 = currentTotalPolishWeight;
                                range.Cells[lastStoneNameRow, 6].Value2 = currentTotalValue;
                                range.Cells[lastStoneNameRow, 6].NumberFormat = "0.00"; // set the number format to display integers only

                                // Update the current row's data with the new stone name and the updated values
                                currentStoneName = stoneName;
                                range.Cells[i, 1].Value2 = stoneName;
                                range.Cells[i, 3].Value2 = currentTotalPartWeight;
                                range.Cells[i, 4].Value2 = currentTotalPolishWeight;
                                range.Cells[i, 6].Value2 = currentTotalValue;
                                range.Cells[i, 6].NumberFormat = "0.00"; // set the number format to display integers only

                                // Update the last stone name row
                                lastStoneNameRow = i;
                                currentTotalPartWeight = 0;
                                currentTotalPolishWeight = 0;
                                currentTotalValue = 0;
                            }
                            else
                            {
                                // If the stone name is the same, do nothing
                                currentStoneName = stoneName;
                            }
                        }

                        double partWT = (range.Cells[i, 10].Value2 != null) ? Convert.ToDouble(range.Cells[i, 10].Value2) : 0;
                        currentTotalPartWeight += partWT;

                        double Pw = (range.Cells[i, 11].Value2 != null) ? Convert.ToDouble(range.Cells[i, 11].Value2) : 0;
                        currentTotalPolishWeight += Pw;

                        double Dolar = (range.Cells[i, 18].Value2 != null) ? Convert.ToDouble(range.Cells[i, 18].Value2) : 0;
                        currentTotalValue += Dolar;

                        // Update the progress bar and label text
                        int progressPercentage = (int)Math.Round((double)i / totalRows * 100);
                        progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                        Invoke(new System.Action(() =>
                        {
                            progressBar1.Value = progressPercentage;
                            string currentMethod = "UPDATING HEADERS OF NEW COLUMNS";
                            lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                        }));
                    }

                    // Update the last stone name row data with the sum result
                    range.Cells[lastStoneNameRow, 3].Value2 = currentTotalPartWeight;
                    range.Cells[lastStoneNameRow, 4].Value2 = currentTotalPolishWeight;
                    range.Cells[lastStoneNameRow, 6].Value2 = currentTotalValue;
                    range.Cells[lastStoneNameRow, 6].NumberFormat = "0.00"; // set the number format to display integers only

                    // Assign unique values to PolishSrNo column
                    int polishSrNo = 1;
                    for (int i = 2; i <= rowCount; i++)
                    {
                        range.Cells[i, 7].Value2 = polishSrNo;
                        range.Cells[i, 7].NumberFormat = "0"; // Set number format to display integers only
                        range.Cells[i, 7].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        polishSrNo++;

                        // Remove "x_" from the "Sieve" column data which is column "M"
                        string sieveValue = (range.Cells[i, 13].Value2 != null) ? range.Cells[i, 13].Value2.ToString() : "";
                        if (sieveValue.StartsWith("x_"))
                        {
                            range.Cells[i, 13].Value2 = sieveValue.Substring(2);
                        }

                        // Update the progress bar and label text
                        int progressPercentage = (int)Math.Round((double)i / totalRows * 100);
                        progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                        Invoke(new System.Action(() =>
                        {
                            progressBar1.Value = progressPercentage;
                            string currentMethod = "ASSIGNING UNIQUE VALUES TO POLISH SR NO. COLUMN";
                            lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                        }));
                    }

                    #region Initial SUMMARY Calculation

                    int stoneNameCount = 0;
                    double totalRoughtWeight = 0;
                    double partWTTotal = 0;
                    double pwTotal = 0.0;
                    double dolarTotal = 0;
                    int partCount = 0;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        if (range.Cells[i, 1].Value2 != null && range.Cells[i, 1].Value2.ToString() != "")
                        {
                            stoneNameCount++;
                        }

                        if (range.Cells[i, 2].Value2 != null && range.Cells[i, 2].Value2.ToString() != "")
                        {
                            totalRoughtWeight += Convert.ToDouble(range.Cells[i, 2].Value2);
                        }

                        if (range.Cells[i, 10].Value2 != null)
                        {
                            partWTTotal += double.Parse(range.Cells[i, 10].Value2.ToString());
                        }

                        if (range.Cells[i, 11].Value2 != null)
                        {
                            pwTotal += range.Cells[i, 11].Value2;
                        }

                        if (range.Cells[i, 18].Value2 != null && range.Cells[i, 18].Value2.ToString() != "")
                        {
                            dolarTotal += double.Parse(range.Cells[i, 18].Value2.ToString());
                        }

                        if (range.Cells[i, 8].Value2 != null && range.Cells[i, 8].Value2.ToString() != "")
                        {
                            partCount++;
                        }

                        // Update the progress bar and label text
                        int progressPercentage = (int)Math.Round((double)i / totalRows * 100);
                        progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                        Invoke(new System.Action(() =>
                        {
                            progressBar1.Value = progressPercentage;
                            string currentMethod = "COUNTING TOTAL OF STONE NAME, ROUGH WEIGHT, PART WEIGHT, POLISH WEIGHT, DOLAR AND PART";
                            lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                        }));
                    }

                    // Write the total in the end of column
                    range.Cells[rowCount + 4, 1].Value2 = "RoughPCs";
                    range.Cells[rowCount + 3, 1].Value2 = stoneNameCount;

                    // Write the total in the end of column
                    range.Cells[rowCount + 4, 2].Value2 = "RoughWeight";
                    range.Cells[rowCount + 3, 2].Value2 = Math.Round(totalRoughtWeight, 2);

                    // Write the sum at the end of column
                    range.Cells[rowCount + 4, 3].Value2 = "Total PartWeight";
                    range.Cells[rowCount + 3, 3].Value2 = Math.Round(partWTTotal, 2);

                    // Write the sum at the end of column
                    range.Cells[rowCount + 4, 4].Value2 = "Total PolishWeight";
                    range.Cells[rowCount + 3, 4].Value2 = Math.Round(pwTotal, 2);

                    // Write the Percentage at the end of column
                    double pwPercentage = (pwTotal / totalRoughtWeight) * 100;
                    range.Cells[rowCount + 4, 5].Value2 = "R2P(%)";
                    range.Cells[rowCount + 3, 5].Value2 = pwPercentage + "%";

                    // Write the dolarTotal in the end of column
                    range.Cells[rowCount + 4, 6].Value2 = "Total Value";
                    range.Cells[rowCount + 3, 6].Value2 = Math.Round(dolarTotal, 2);

                    // Calculate the Value/Rough
                    double valueRough = (dolarTotal / totalRoughtWeight);
                    range.Cells[rowCount + 4, 10].Value2 = "Value/Rough Crt";
                    range.Cells[rowCount + 3, 10].Value2 = valueRough.ToString("0.00");

                    // Calculate the Value/Polish
                    double valuePolish = (dolarTotal / pwTotal);
                    range.Cells[rowCount + 4, 11].Value2 = "Value/Polish Crt";
                    range.Cells[rowCount + 3, 11].Value2 = valuePolish.ToString("0.00");

                    // Write the total in the end of column
                    range.Cells[rowCount + 4, 7].Value2 = "PolishPCs";
                    range.Cells[rowCount + 3, 7].Value2 = partCount;

                    // Calculate the PolishPCs with totalRoughtWeight
                    double totalSize = (partCount / totalRoughtWeight);
                    range.Cells[rowCount + 4, 8].Value2 = "Craft Size";
                    range.Cells[rowCount + 3, 8].Value2 = totalSize.ToString("0.00");

                    // Calculate the PolishPCs with totalRoughtWeight
                    double polishSize = (partCount / pwTotal);
                    range.Cells[rowCount + 4, 9].Value2 = "Polish Size";
                    range.Cells[rowCount + 3, 9].Value2 = polishSize.ToString("0.00");

                    #endregion


                    #region Main SUMMARY Calculation

                    Dictionary<string, double> clarityPw = new Dictionary<string, double>();

                    // Define an array of clarity values in the given sequence
                    string[] clarityValues = { "VVS", "VVS1", "VVS2", "VS1", "VS2", "SI1+", "SI1", "SI2+",
                        "SI2", "SI3", "I1", "I2", "I3", "I1A", "I1B", "PK", "PK1", "PK2" };

                    // Filter unique clarity values and their counts
                    for (int i = 2; i <= rowCount; i++)
                    {
                        string clarity = (range.Cells[i, 16].Value2 != null) ? range.Cells[i, 16].Value2.ToString() : "";
                        double pw = (range.Cells[i, 11].Value2 != null) ? range.Cells[i, 11].Value2 : 0.0;

                        if (!string.IsNullOrEmpty(clarity) && clarity != "Clarity" && clarityValues.Contains(clarity))
                        {
                            if (clarityCounts.ContainsKey(clarity))
                            {
                                clarityCounts[clarity]++;
                                clarityPw[clarity] += pw;
                            }
                            else
                            {
                                clarityCounts.Add(clarity, 1);
                                clarityPw.Add(clarity, pw);
                            }
                        }

                        // Update the progress bar and label text
                        int progressPercentage = (int)Math.Round((double)i / totalRows * 100);
                        progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                        Invoke(new System.Action(() =>
                        {
                            progressBar1.Value = progressPercentage;
                            string currentMethod = "FILTERING UNIQUE CLARITY VALUES AND THEIR COUNTS";
                            lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                        }));
                    }

                    // Calculate total count and total Pw
                    int totalCount = clarityCounts.Values.Sum();
                    double totalPw = clarityPw.Values.Sum();

                    // Write the filtered data to the end of the original worksheet
                    int filterRow = rowCount + 1;

                    string[] headers = { "Sr No.", "Purity", "%of PolishWeight", "Rough CRT", "PolishPCs", "Craft Weight", "PolishWeight", "Rough Size", 
                        "Polish Size", "Craft To Polish %", "Rough To Polish %", "Polish Dollar", "Value/Rough Cts", "Value/Polish Cts" };

                    // Write header row
                    for (int col = 1; col <= headers.Length; col++)
                    {
                        sheet.Cells[filterRow + 5, col].Value2 = headers[col - 1];
                        sheet.Cells[filterRow + 5, col].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    }

                    // Apply formatting to header cells
                    Microsoft.Office.Interop.Excel.Range headerRange1 = sheet.Range[sheet.Cells[filterRow + 5, 1], sheet.Cells[filterRow + 5, 14]];
                    headerRange1.Interior.Color = System.Drawing.Color.LightGray;

                    // Define dictionary to store weight loss for each part
                    Dictionary<string, double> weightLossDict = new Dictionary<string, double>();

                    // Loop through rows in excel sheet
                    for (int i = 2; i <= rowCount; i++)
                    {
                        // Get values for stone name, rough weight, and total part weight
                        string stoneName = (range.Cells[i, 1].Value2 != null) ? range.Cells[i, 1].Value2.ToString() : " ";
                        double roughWeight = (range.Cells[i, 2].Value2 != null) ? range.Cells[i, 2].Value2 : 0;
                        double totalPartWeight = (range.Cells[i, 3].Value2 != null) ? range.Cells[i, 3].Value2 : 0;

                        // Subtract total part weight from rough weight to get weight loss
                        double weightLoss = roughWeight - totalPartWeight;

                        // Add weight loss to dictionary
                        if (!string.IsNullOrEmpty(stoneName))
                        {
                            if (weightLossDict.ContainsKey(stoneName))
                            {
                                weightLossDict[stoneName] += weightLoss;
                            }
                            else
                            {
                                weightLossDict.Add(stoneName, weightLoss);
                            }
                        }

                        // Update the progress bar and label text
                        int progressPercentage = (int)Math.Round((double)i / totalRows * 100);
                        progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                        Invoke(new System.Action(() =>
                        {
                            progressBar1.Value = progressPercentage;
                            string currentMethod = "DEFINING DICTIONARY TO STORE WEIGHT LOSS FOR EACH PART";
                            lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                        }));
                    }

                    // Define dictionary to store PartWT for each stone name
                    Dictionary<string, List<Tuple<double, double>>> partWtDict = new Dictionary<string, List<Tuple<double, double>>>();

                    string prevStoneeName = "";
                    double prevRoughhWeight = 0.0;
                    // Loop through rows in excel sheet again to count PartWT for each stone name
                    for (int t = 2; t <= rowCount; t++)
                    {
                        // Get values for stone name, PartWT, and Width
                        string stoneeName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStoneeName;
                        double roughhWeight = (range.Cells[t, 2].Value2 != null) ? range.Cells[t, 2].Value2 : prevRoughhWeight;
                        string partWT = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.0";
                        double width = (range.Cells[t, 20].Value2 != null) ? range.Cells[t, 20].Value2 : 0.0;
                        double polishWeight = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;

                        // Add PartWT and Width to the list for the stone name
                        if (!string.IsNullOrEmpty(stoneeName) && double.TryParse(partWT, out double partWtValue) && partWtValue > 0)
                        {
                            if (partWtDict.ContainsKey(stoneeName))
                            {
                                partWtDict[stoneeName].Add(Tuple.Create(partWtValue, width));
                            }
                            else
                            {
                                List<Tuple<double, double>> partWts = new List<Tuple<double, double>>();
                                partWts.Add(Tuple.Create(partWtValue, width));
                                partWtDict.Add(stoneeName, partWts);
                            }
                        }

                        prevStoneeName = stoneeName;
                        prevRoughhWeight = roughhWeight;

                        // Update the progress bar and label text
                        int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                        progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                        Invoke(new System.Action(() =>
                        {
                            progressBar1.Value = progressPercentage;
                            string currentMethod = "DEFINING DICTIONARY TO STORE PART WEIGHT FOR EACH STONE NAME";
                            lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                        }));
                    }

                    // Loop through weightLossDict and distribute weight loss to each part weight for the corresponding stone name
                    foreach (string stoneName in weightLossDict.Keys)
                    {
                        double weightLoss = weightLossDict[stoneName];
                        int numParts = partWtDict.ContainsKey(stoneName) ? partWtDict[stoneName].Count : 0;

                        if (numParts > 0)
                        {
                            double weightLossPerPart = weightLoss / numParts;

                            if (partWtDict.ContainsKey(stoneName))
                            {
                                List<Tuple<double, double>> partWts = partWtDict[stoneName];

                                for (int i = 0; i < partWts.Count; i++)
                                {
                                    Tuple<double, double> partWt = partWts[i];
                                    double updatedPartWt = partWt.Item1 + weightLossPerPart;
                                    partWts[i] = Tuple.Create(updatedPartWt, partWt.Item2);
                                }

                                partWtDict[stoneName] = partWts;
                            }
                        }
                    }

                    Dictionary<string, List<string>> clarityDict = new Dictionary<string, List<string>>();

                    string prevStoneeeName = "";
                    for (int t = 2; t <= rowCount; t++)
                    {
                        string stoneeName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStoneeeName;
                        string claritys = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";

                        // Check if claritys is blank and assign "Empty"
                        if (string.IsNullOrEmpty(claritys))
                        {
                            claritys = "Empty";

                            // Display error message and prompt user to continue or stop the report
                            var result = MessageBox.Show($"Clarity Value is Blank for Stone Name '{stoneeName}'. Continue the Report? By Default Value will be Assigned Empty", "Clarity Value Blank", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                            // If the user chooses to stop the report, break out of the loop
                            if (result == DialogResult.No)
                            {
                                break;
                            }
                        }

                        if (!string.IsNullOrEmpty(stoneeName))
                        {
                            if (clarityDict.ContainsKey(stoneeName))
                            {
                                clarityDict[stoneeName].Add(claritys);
                            }
                            else
                            {
                                clarityDict[stoneeName] = new List<string>() { claritys };
                            }
                        }
                        prevStoneeeName = stoneeName;

                        // Update the progress bar and label text
                        int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                        progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                        Invoke(new System.Action(() =>
                        {
                            progressBar1.Value = progressPercentage;
                            string currentMethod = "DEFINING DICTIONARY TO STORE CLARITY FOR EACH STONE NAME";
                            lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                        }));
                    }

                    Dictionary<string, double> clarityWeightDict = new Dictionary<string, double>();

                    foreach (string stoneName in partWtDict.Keys)
                    {
                        // Get the list of part weights for the current stone name
                        List<Tuple<double, double>> partWts = partWtDict[stoneName];

                        // Get the list of clarities for the current stone name
                        List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                        // Loop through each unique clarity for the current stone name
                        foreach (string clarity in clarities.Distinct())
                        {
                            // Sum the weights of the parts with the current clarity
                            double clarityWeight = partWts.Where((p, i) => clarities[i] == clarity).Sum(partWt => partWt.Item1);


                            // Add the clarity weight to the clarityWeightDict
                            if (clarityWeightDict.ContainsKey(clarity))
                            {
                                clarityWeightDict[clarity] += clarityWeight;
                            }
                            else
                            {
                                clarityWeightDict[clarity] = clarityWeight;
                            }
                        }
                    }

                    // Define dictionary to store PartWT for each stone name
                    Dictionary<string, List<double>> partWt2Dict = new Dictionary<string, List<double>>();

                    string prevStonee2Name = "";
                    // Loop through rows in excel sheet again to count PartWT for each stone name
                    for (int t = 2; t <= rowCount; t++)
                    {
                        // Get values for stone name and PartWT
                        string stoneeName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStonee2Name;
                        string partWT = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.0";

                        // Add PartWT to list for the stone name
                        if (!string.IsNullOrEmpty(stoneeName) && double.TryParse(partWT, out double partWtValue) && partWtValue > 0)
                        {
                            if (partWt2Dict.ContainsKey(stoneeName))
                            {
                                partWt2Dict[stoneeName].Add(partWtValue);
                            }
                            else
                            {
                                List<double> partWts = new List<double>();
                                partWts.Add(partWtValue);
                                partWt2Dict.Add(stoneeName, partWts);
                            }
                        }

                        prevStonee2Name = stoneeName;

                        // Update the progress bar and label text
                        int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                        progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                        Invoke(new System.Action(() =>
                        {
                            progressBar1.Value = progressPercentage;
                            string currentMethod = "DEFINING DICTIONARY TO STORE CLARITY WEIGHT FOR EACH STONE NAME";
                            lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                        }));
                    }

                    Dictionary<string, double> clarityWeight2Dict = new Dictionary<string, double>();

                    foreach (string stoneName in partWt2Dict.Keys)
                    {
                        // Get the list of part weights for the current stone name
                        List<double> partWts = partWt2Dict[stoneName];

                        // Get the list of clarities for the current stone name
                        List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                        // Loop through each unique clarity for the current stone name
                        foreach (string clarity in clarities.Distinct())
                        {
                            // Sum the weights of the parts with the current clarity
                            double clarityWeight = partWts.Where((p, i) => clarities[i] == clarity).Sum();

                            // Add the clarity weight to the clarityWeightDict
                            if (clarityWeight2Dict.ContainsKey(clarity))
                            {
                                clarityWeight2Dict[clarity] += clarityWeight;
                            }
                            else
                            {
                                clarityWeight2Dict[clarity] = clarityWeight;
                            }
                        }
                    }

                    Dictionary<string, List<double>> clarityDolar = new Dictionary<string, List<double>>();

                    string prevStonePoName = "";

                    // Loop through rows in excel sheet again to count Dolar for each stone name
                    for (int t = 2; t <= rowCount; t++)
                    {
                        // Get values for stone name and Dolar
                        string stoneeName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStonePoName;
                        string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.0";
                        if (!string.IsNullOrEmpty(stoneeName) && double.TryParse(PoDolar, out double partPoValue))
                        {
                            if (clarityDolar.ContainsKey(stoneeName))
                            {
                                clarityDolar[stoneeName].Add(partPoValue);
                            }
                            else
                            {
                                List<double> PoDolars = new List<double>();
                                PoDolars.Add(partPoValue);
                                clarityDolar.Add(stoneeName, PoDolars);
                            }
                        }
                        prevStonePoName = stoneeName;

                        // Update the progress bar and label text
                        int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                        progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                        Invoke(new System.Action(() =>
                        {
                            progressBar1.Value = progressPercentage;
                            string currentMethod = "DEFINING DICTIONARY TO STORE CLARITY DOLAR FOR EACH STONE NAME";
                            lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                        }));
                    }

                    Dictionary<string, double> clarityDolarDict = new Dictionary<string, double>();

                    foreach (string stoneName in clarityDolar.Keys)
                    {
                        // Get the list of dolar for the current stone name
                        List<double> PoDolars = clarityDolar[stoneName];

                        // Get the list of clarities for the current stone name
                        List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                        // Loop through each unique clarity for the current stone name
                        foreach (string clarity in clarities.Distinct())
                        {
                            // Sum the dolar of the parts with the current clarity
                            double clarityDolars = PoDolars.Where((p, i) => clarities[i] == clarity).Sum();

                            // Add the clarity weight to the clarityDolartDict
                            if (clarityDolarDict.ContainsKey(clarity))
                            {
                                clarityDolarDict[clarity] += clarityDolars;
                            }
                            else
                            {
                                clarityDolarDict[clarity] = clarityDolars;
                            }
                        }
                    }

                    int serialNumber = 1;

                    int totalClarityFilters = clarityValues.Count();
                    int currentClarityFilter = 0;

                    // Loop through each clarity filter and write the results to the worksheet
                    foreach (string clarity in clarityValues)
                    {
                        // Update progress
                        currentClarityFilter++;
                        int progressPercentage = (int)Math.Round((double)currentClarityFilter / totalClarityFilters * 100);
                        progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                        Invoke(new System.Action(() =>
                        {
                            progressBar1.Value = progressPercentage;
                            string currentMethod = "THE PROCESS OF FINALIZING THE 'SUMMARY' TOTAL VALUES IS UNDERWAY";
                            lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                        }));

                        if (clarityCounts.ContainsKey(clarity))
                        {
                            // Autofit columns
                            sheet.Columns.AutoFit();

                            // Serial number wise
                            sheet.Cells[filterRow + 6, 1].Value2 = serialNumber;
                            sheet.Cells[filterRow + 6, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                            // Get the clarity Filter
                            sheet.Cells[filterRow + 6, 2].Value2 = clarity;

                            // Get % of polish weight
                            double percentPw = clarityPw[clarity] / totalPw * 100.0;
                            sheet.Cells[filterRow + 6, 3].Value2 = percentPw.ToString("0.00") + "%";

                            // Get total count of roughPcs
                            sheet.Cells[filterRow + 6, 5].Value2 = clarityCounts[clarity];

                            // Get the Part Weight for the clarity
                            sheet.Cells[filterRow + 6, 7].Value2 = clarityPw[clarity];

                            // Get the rough Crt weight for this clarity
                            double clarityWeight = clarityWeightDict.ContainsKey(clarity) ? clarityWeightDict[clarity] : 0.0;
                            sheet.Cells[filterRow + 6, 4].Value2 = Math.Round(clarityWeight, 2);

                            // Get the craft weight for this clarity
                            double clarityWeight2 = clarityWeight2Dict.ContainsKey(clarity) ? clarityWeight2Dict[clarity] : 0.0;
                            sheet.Cells[filterRow + 6, 6].Value2 = Math.Round(clarityWeight2, 2);

                            // Calculate the division of clarityCounts by clarityWeightDict
                            double clarityCountDivided = clarityCounts[clarity] / clarityWeightDict[clarity];
                            sheet.Cells[filterRow + 6, 8].Value2 = Math.Round(clarityCountDivided, 2);

                            // Calculate the division of clarityCounts by clarityPolishWeight
                            double clarityPolishWtCount = clarityCounts[clarity] / clarityPw[clarity];
                            sheet.Cells[filterRow + 6, 9].Value2 = Math.Round(clarityPolishWtCount, 2);

                            // Calculate the % of polish weight with craft weight
                            double percentPwCr = clarityPw[clarity] / clarityWeight2Dict[clarity] * 100;
                            sheet.Cells[filterRow + 6, 10].Value2 = percentPwCr.ToString("0.00") + "%";

                            // Calculate the % rough weight with craft weight
                            double percentRoCr = clarityPw[clarity] / clarityWeightDict[clarity] * 100;
                            sheet.Cells[filterRow + 6, 11].Value2 = percentRoCr.ToString("0.00") + "%";

                            // Calculate dolar sum clarity wise
                            sheet.Cells[filterRow + 6, 12].Value2 = clarityDolarDict[clarity];

                            // Calculate division of dolar by rough crt clarity wise
                            double clarityRoughCrtDolar = clarityDolarDict[clarity] / clarityWeightDict[clarity];
                            sheet.Cells[filterRow + 6, 13].Value2 = Math.Round(clarityRoughCrtDolar, 2);

                            // Calculate division of dolar by rough crt clarity wise
                            double clarityPolishCrtDolar = clarityDolarDict[clarity] / clarityPw[clarity];
                            sheet.Cells[filterRow + 6, 14].Value2 = Math.Round(clarityPolishCrtDolar, 2);

                            serialNumber++;
                            filterRow++;
                        }
                    }

                    // Apply table formatting
                    Microsoft.Office.Interop.Excel.Range tableRange = sheet.Range[sheet.Cells[filterRow - serialNumber + 6, 1], sheet.Cells[filterRow + 5, 14]];
                    tableRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    tableRange.Font.Size = 11;
                    tableRange.Columns.AutoFit();

                    // Write the total count and total Pw to the worksheet
                    sheet.Cells[filterRow + 6, 1].Value2 = "Total";
                    sheet.Cells[filterRow + 6, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                    sheet.Cells[filterRow + 6, 4].Value2 = Math.Round(totalRoughtWeight, 3);
                    sheet.Cells[filterRow + 6, 6].Value2 = Math.Round(partWTTotal, 2);
                    sheet.Cells[filterRow + 6, 5].Value2 = totalCount;
                    sheet.Cells[filterRow + 6, 7].Value2 = Math.Round(totalPw, 2);
                    sheet.Cells[filterRow + 6, 8].Value2 = totalSize.ToString("0.00");
                    sheet.Cells[filterRow + 6, 9].Value2 = polishSize.ToString("0.00");

                    double crPwPercentage = (pwTotal / partWTTotal) * 100;
                    sheet.Cells[filterRow + 6, 10].Value2 = crPwPercentage.ToString("0.00") + "%";

                    sheet.Cells[filterRow + 6, 11].Value2 = pwPercentage + "%";
                    sheet.Cells[filterRow + 6, 12].Value2 = Math.Round(dolarTotal, 2);
                    sheet.Cells[filterRow + 6, 13].Value2 = valueRough.ToString("0.00");
                    sheet.Cells[filterRow + 6, 14].Value2 = valuePolish.ToString("0.00");

                    // Apply formatting to header cells
                    Microsoft.Office.Interop.Excel.Range headerRangee1 = sheet.Range[sheet.Cells[filterRow + 6, 1], sheet.Cells[filterRow + 6, 14]];
                    headerRangee1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    headerRangee1.Interior.Color = System.Drawing.Color.LightGray;


                    #endregion


                    #region Sieve Calculation

                    // Check if any of the checkboxes are checked
                    bool runSieveCalculations = checkBoxRunCode01.Checked || checkBoxRunCode02.Checked || checkBoxRunCode03.Checked ||
                    checkBoxRunCode04.Checked || checkBoxRunCode05.Checked || checkBoxRunCode06.Checked || checkBoxRunCode07.Checked ||
                    checkBoxRunCode001.Checked || checkBoxRunCode002.Checked || checkBoxRunCode003.Checked || checkBoxRunCode004.Checked ||
                    checkBoxRunCode005.Checked || checkBoxRunCode006.Checked || checkBoxRunCode007.Checked || checkBoxRunCode008.Checked ||
                    checkBoxRunCode009.Checked || checkBoxRunCode010.Checked || checkBoxRunCode011.Checked || checkBoxRunCode012.Checked ||
                    checkBoxRunCode013.Checked || checkBoxRunCode014.Checked || checkBoxRunCode015.Checked;

                    // Define dictionary to store PartWT for each stone name
                    Dictionary<string, List<Tuple<double, double, string, double>>> partWtForWdDict = new Dictionary<string, List<Tuple<double, double, string, double>>>();

                    if (runSieveCalculations)
                    {
                        string prevStoneeNameForWd = "";
                        double prevRoughhWeightForWd = 0.0;

                        // Loop through rows in excel sheet again to count PartWT for each stone name
                        for (int t = 2; t <= rowCount; t++)
                        {
                            // Get values for stone name, PartWT, and shape
                            string stoneeName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStoneeNameForWd;
                            double roughhWeight = (range.Cells[t, 2].Value2 != null) ? range.Cells[t, 2].Value2 : prevRoughhWeightForWd;
                            string partWT = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.0";
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            double width = (range.Cells[t, 20].Value2 != null) ? range.Cells[t, 20].Value2 : 0.000;
                            double polishWeight = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;

                            // Add PartWT and Width to the list for the stone name
                            if (!string.IsNullOrEmpty(stoneeName) && double.TryParse(partWT, out double partWtValue) && partWtValue > 0)
                            {
                                if (partWtForWdDict.ContainsKey(stoneeName))
                                {
                                    partWtForWdDict[stoneeName].Add(Tuple.Create(partWtValue, width, shape, polishWeight));
                                }
                                else
                                {
                                    List<Tuple<double, double, string, double>> partWts = new List<Tuple<double, double, string, double>>();
                                    partWts.Add(Tuple.Create(partWtValue, width, shape, polishWeight));
                                    partWtForWdDict.Add(stoneeName, partWts);
                                }
                            }

                            prevStoneeNameForWd = stoneeName;
                            prevRoughhWeightForWd = roughhWeight;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "STARTING SIEVE CALCULATIONS";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        // Loop through weightLossDict and distribute weight loss to each part weight for the corresponding stone name
                        foreach (string stoneName in weightLossDict.Keys)
                        {
                            double weightLoss = weightLossDict[stoneName];
                            int numParts = partWtForWdDict.ContainsKey(stoneName) ? partWtForWdDict[stoneName].Count : 0;

                            if (numParts > 0)
                            {
                                double weightLossPerPart = weightLoss / numParts;

                                if (partWtForWdDict.ContainsKey(stoneName))
                                {
                                    List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                                    for (int i = 0; i < partWts.Count; i++)
                                    {
                                        Tuple<double, double, string, double> partWt = partWts[i];
                                        double updatedPartWt = partWt.Item1 + weightLossPerPart;
                                        partWts[i] = Tuple.Create(updatedPartWt, partWt.Item2, partWt.Item3, partWt.Item4);
                                    }
                                    partWtForWdDict[stoneName] = partWts;
                                }
                            }
                        }
                    }

                    #endregion


                    #region 1st Sieve Round Method

                    if (checkBoxRunCode01.Checked)
                    {
                        sheet.Cells[filterRow + 9, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 9, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 9, 5].Value2 = "Diameter";

                        string updatedValue1 = textBox1.Text;
                        sheet.Cells[filterRow + 10, 3].Value2 = updatedValue1;

                        string updatedValue01 = txtWidthRange.Text;
                        sheet.Cells[filterRow + 10, 5].Value2 = updatedValue01;

                        txtWidthRange.Text = (string)sheet.Cells[filterRow + 10, 5].Value2;

                        sheet.Cells[filterRow + 10, 1].Value2 = "ROUND";

                        textBox1.Text = (string)sheet.Cells[filterRow + 10, 3].Value2;

                        sheet.Cells[filterRow + 11, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 11, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 11, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 11, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 11, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 11, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 11, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 11, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 11, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 11, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 11, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 11, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 11, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 11, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 11, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRange01 = sheet.Range[sheet.Cells[filterRow + 11, 1], sheet.Cells[filterRow + 11, 14]];
                        headerRange01.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the width range from the txtWidthRange TextBox
                        string widthRangeText = txtWidthRange.Text;
                        string[] widthRangeParts = widthRangeText.Split('-');
                        double minWidth = 0.9;
                        double maxWidth = 1.249;

                        // Parse the width range values if the input is valid
                        if (widthRangeParts.Length == 2 && double.TryParse(widthRangeParts[0], out double parsedMinWidth) && double.TryParse(widthRangeParts[1], out double parsedMaxWidth))
                        {
                            minWidth = parsedMinWidth;
                            maxWidth = parsedMaxWidth;
                        }

                        Dictionary<string, int> clarityCounts01 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPw01 = new Dictionary<string, double>();

                        Dictionary<string, List<(string shape, double width, string clarity, double partWT, string PoDolar)>> stoneWidths01 =
                            new Dictionary<string, List<(string shape, double width, string clarity, double partWT, string PoDolar)>>();

                        string prevStone01Name = "";
                        // Loop through rows in the excel sheet to read the "Width" column
                        for (int t = 2; t <= rowCount; t++)
                        {
                            // Get the stone name, clarity, polish weight, part weight, and PoDolar values
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStone01Name;
                            string widthString = (range.Cells[t, 20].Value2 != null) ? range.Cells[t, 20].Value2.ToString() : prevStone01Name;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            double pw = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "round" && double.TryParse(widthString, out double width))
                            {
                                // Check if the width falls within the specified range
                                if (width >= minWidth && width <= maxWidth && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCounts01.ContainsKey(clarity))
                                        {
                                            clarityCounts01[clarity]++;
                                            clarityPw01[clarity] += pw;
                                        }
                                        else
                                        {
                                            clarityCounts01.Add(clarity, 1);
                                            clarityPw01.Add(clarity, pw);
                                        }

                                        // Add the width, clarity, polish weight, and part weight to the stoneWidths01 dictionary for the current stone name
                                        if (stoneWidths01.ContainsKey(stoneName))
                                        {
                                            stoneWidths01[stoneName].Add((shape, width, clarity, double.Parse(partWTString), PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double width, string clarity, double partWT, string PoDolar)> data = new List<(string shape, double width, string clarity, double partWT, string PoDolar)>();
                                            data.Add((shape, width, clarity, double.Parse(partWTString), PoDolar));
                                            stoneWidths01.Add(stoneName, data);
                                        }
                                    }
                                }
                            }

                            prevStone01Name = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING FIRST SIEVE FOR SHAPE 'ROUND' CHECKING IF THE WIDTH FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        Dictionary<string, double> clarityWeightDict01 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> roundParts = partWts.Where(p => p.Item3.ToLower() == "round").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForWd01 = roundParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item2 >= minWidth && p.Item2 <= maxWidth)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDict01.ContainsKey(clarity))
                                {
                                    clarityWeightDict01[clarity] += clarityWeightForWd01;
                                }
                                else
                                {
                                    clarityWeightDict01[clarity] = clarityWeightForWd01;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDict02 = new Dictionary<string, double>();

                        // Iterate over the stone names in stoneWidths01
                        foreach (var stoneData in stoneWidths01.Values.SelectMany(data => data))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDict02.ContainsKey(clarity))
                            {
                                clarityWeightDict02[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDict02[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDict01 = new Dictionary<string, double>();

                        // Iterate over the stone names in stoneWidths01
                        foreach (var stoneData in stoneWidths01.Values.SelectMany(data => data))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDict01.ContainsKey(clarity))
                                {
                                    clarityDolarDict01[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDict01[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCount01 = clarityCounts01.Values.Sum();
                        double totalPw01 = clarityPw01.Values.Sum();

                        int serialNumber01 = 1;

                        int totalClarityFilters01 = clarityValues.Count();
                        int currentClarityFilter01 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter01++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilter01 / totalClarityFilters01 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING FIRST 'ROUND' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCounts01.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 12, 1].Value2 = serialNumber01;
                                sheet.Cells[filterRow + 12, 1].HorizontalAlignment =
                                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 12, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPw1 = clarityPw01[clarity] / totalPw01 * 100.0;
                                sheet.Cells[filterRow + 12, 3].Value2 = percentPw1.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 12, 5].Value2 = clarityCounts01[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 12, 7].Value2 = clarityPw01[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeight01 = clarityWeightDict01.ContainsKey(clarity) ? clarityWeightDict01[clarity] : 0.000;
                                sheet.Cells[filterRow + 12, 4].Value2 = Math.Round(clarityWeight01, 3);

                                // Get the craft weight for this clarity
                                double clarityWeight02 = clarityWeightDict02.ContainsKey(clarity) ? clarityWeightDict02[clarity] : 0.000;
                                sheet.Cells[filterRow + 12, 6].Value2 = Math.Round(clarityWeight02, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDivided01 = clarityCounts01[clarity] / clarityWeightDict01[clarity];
                                sheet.Cells[filterRow + 12, 8].Value2 = Math.Round(clarityCountDivided01, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCount01 = clarityCounts01[clarity] / clarityPw01[clarity];
                                sheet.Cells[filterRow + 12, 9].Value2 = Math.Round(clarityPolishWtCount01, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCr01 = clarityPw01[clarity] / clarityWeightDict02[clarity] * 100;
                                sheet.Cells[filterRow + 12, 10].Value2 = percentPwCr01.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCr01 = clarityPw01[clarity] / clarityWeightDict01[clarity] * 100;
                                sheet.Cells[filterRow + 12, 11].Value2 = percentRoCr01.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 12, 12].Value2 = clarityDolarDict01[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolar01 = clarityDolarDict01[clarity] / clarityWeightDict01[clarity];
                                sheet.Cells[filterRow + 12, 13].Value2 = Math.Round(clarityRoughCrtDolar01, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolar01 = clarityDolarDict01[clarity] / clarityPw01[clarity];
                                sheet.Cells[filterRow + 12, 14].Value2 = Math.Round(clarityPolishCrtDolar01, 2);

                                serialNumber01++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRange1 = sheet.Range[sheet.Cells[filterRow - serialNumber01 + 12, 1], sheet.Cells[filterRow + 11, 14]];
                        tableRange1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRange1.Font.Size = 11;
                        tableRange1.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 12, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 12, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeight01 = clarityWeightDict01.Values.Sum();
                        sheet.Cells[filterRow + 12, 4].Value2 = Math.Round(totalRoughtWeight01, 2);

                        sheet.Cells[filterRow + 12, 5].Value2 = totalCount01;

                        double partWTTotal01 = clarityWeightDict02.Values.Sum();
                        sheet.Cells[filterRow + 12, 6].Value2 = Math.Round(partWTTotal01, 2);

                        sheet.Cells[filterRow + 12, 7].Value2 = Math.Round(totalPw01, 3);

                        double totalSize01 = (totalCount01 / totalRoughtWeight01);
                        sheet.Cells[filterRow + 12, 8].Value2 = totalSize01.ToString("0.00");

                        double polishSize01 = (totalCount01 / totalPw01);
                        sheet.Cells[filterRow + 12, 9].Value2 = polishSize01.ToString("0.00");

                        double crPwPercentage01 = (totalPw01 / partWTTotal01) * 100;
                        sheet.Cells[filterRow + 12, 10].Value2 = crPwPercentage01.ToString("0.00") + "%";

                        double pwPercentage01 = (totalPw01 / totalRoughtWeight01) * 100;
                        sheet.Cells[filterRow + 12, 11].Value2 = pwPercentage01.ToString("0.00") + "%";

                        double dolarTotal01 = clarityDolarDict01.Values.Sum();
                        sheet.Cells[filterRow + 12, 12].Value2 = Math.Round(dolarTotal01, 2);

                        double valueRough01 = (dolarTotal01 / totalRoughtWeight01);
                        sheet.Cells[filterRow + 12, 13].Value2 = valueRough01.ToString("0.00");

                        double valuePolish01 = (dolarTotal01 / totalPw01);
                        sheet.Cells[filterRow + 12, 14].Value2 = valuePolish01.ToString("0.00");

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRange2 = sheet.Range[sheet.Cells[filterRow + 12, 1], sheet.Cells[filterRow + 12, 14]];
                        headerRange2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRange2.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion

                    #region 2nd Sieve Round Method

                    if (checkBoxRunCode02.Checked)
                    {
                        sheet.Cells[filterRow + 15, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 15, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 15, 5].Value2 = "Diameter";

                        string updatedValue2 = textBox2.Text;
                        sheet.Cells[filterRow + 16, 3].Value2 = updatedValue2;

                        string updatedValue02 = txtWidthRange01.Text;
                        sheet.Cells[filterRow + 16, 5].Value2 = updatedValue02;

                        //sheet.Cells[filterRow + 16, 5].Value2 = "1.250 - 1.499";
                        txtWidthRange01.Text = (string)sheet.Cells[filterRow + 16, 5].Value2;

                        sheet.Cells[filterRow + 16, 1].Value2 = "ROUND";

                        //sheet.Cells[filterRow + 16, 3].Value2 = "+2 -4.5";
                        textBox2.Text = (string)sheet.Cells[filterRow + 16, 3].Value2;

                        sheet.Cells[filterRow + 17, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 17, 1].HorizontalAlignment =
                            Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 17, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 17, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 17, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 17, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 17, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 17, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 17, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 17, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 17, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 17, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 17, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 17, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 17, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRange001 = sheet.Range[sheet.Cells[filterRow + 17, 1], sheet.Cells[filterRow + 17, 14]];
                        //headerRange001.Font.Bold = true;
                        headerRange001.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the width range from the txtWidthRange TextBox
                        string widthRangeText01 = txtWidthRange01.Text;
                        string[] widthRangeParts01 = widthRangeText01.Split('-');
                        double minWidth01 = 1.249;  // Default minimum width
                        double maxWidth01 = 1.499;  // Default maximum width

                        // Parse the width range values if the input is valid
                        if (widthRangeParts01.Length == 2 && double.TryParse(widthRangeParts01[0], out double parsedMinWidth01) && double.TryParse(widthRangeParts01[1], out double parsedMaxWidth01))
                        {
                            minWidth01 = parsedMinWidth01;
                            maxWidth01 = parsedMaxWidth01;
                        }

                        Dictionary<string, int> clarityCounts02 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPw02 = new Dictionary<string, double>();

                        Dictionary<string, List<(string shape, double width, string clarity, double partWT, string PoDolar)>> stoneWidths02 =
                            new Dictionary<string, List<(string shape, double width, string clarity, double partWT, string PoDolar)>>();

                        string prevStone02Name = "";
                        for (int t = 2; t <= rowCount; t++)
                        {
                            // Get the stone name, clarity, polish weight, part weight, and PoDolar values
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStone02Name;
                            string widthString = (range.Cells[t, 20].Value2 != null) ? range.Cells[t, 20].Value2.ToString() : prevStone02Name;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            double pw = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "round" && double.TryParse(widthString, out double width))
                            {
                                // Check if the width falls within the specified range
                                //if (width > 0.9 && width <= 1.249 && clarityValues.Contains(clarity))
                                if (width >= minWidth01 && width <= maxWidth01 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCounts02.ContainsKey(clarity))
                                        {
                                            clarityCounts02[clarity]++;
                                            clarityPw02[clarity] += pw;
                                        }
                                        else
                                        {
                                            clarityCounts02.Add(clarity, 1);
                                            clarityPw02.Add(clarity, pw);
                                        }

                                        // Add the width, clarity, polish weight, and part weight to the stoneWidths01 dictionary for the current stone name
                                        if (stoneWidths02.ContainsKey(stoneName))
                                        {
                                            stoneWidths02[stoneName].Add((shape, width, clarity, double.Parse(partWTString), PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double width, string clarity, double partWT, string PoDolar)> data = new List<(string shape, double width, string clarity, double partWT, string PoDolar)>();
                                            data.Add((shape, width, clarity, double.Parse(partWTString), PoDolar));
                                            //(shape, width, clarity, double.Parse(partWTString), PoDolar)                                
                                            stoneWidths02.Add(stoneName, data);
                                        }
                                    }
                                }
                            }

                            prevStone02Name = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING SECOND SIEVE FOR SHAPE 'ROUND' CHECKING IF THE WIDTH FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        Dictionary<string, double> clarityWeightDict001 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> roundParts = partWts.Where(p => p.Item3.ToLower() == "round").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForWd001 = roundParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item2 >= minWidth01 && p.Item2 <= maxWidth01)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDict001.ContainsKey(clarity))
                                {
                                    clarityWeightDict001[clarity] += clarityWeightForWd001;
                                }
                                else
                                {
                                    clarityWeightDict001[clarity] = clarityWeightForWd001;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDict002 = new Dictionary<string, double>();

                        // Iterate over the stone names in stoneWidths02
                        foreach (var stoneData in stoneWidths02.Values.SelectMany(data => data))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDict002.ContainsKey(clarity))
                            {
                                clarityWeightDict002[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDict002[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDict001 = new Dictionary<string, double>();

                        // Iterate over the stone names in stoneWidths02
                        foreach (var stoneData in stoneWidths02.Values.SelectMany(data => data))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDict001.ContainsKey(clarity))
                                {
                                    clarityDolarDict001[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDict001[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCount001 = clarityCounts02.Values.Sum();
                        double totalPw001 = clarityPw02.Values.Sum();

                        int serialNumber02 = 1;

                        int totalClarityFilters02 = clarityValues.Count();
                        int currentClarityFilter02 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter02++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilter02 / totalClarityFilters02 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING SECOND 'ROUND' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCounts02.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 18, 1].Value2 = serialNumber02;
                                sheet.Cells[filterRow + 18, 1].HorizontalAlignment =
                                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 18, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPw01 = clarityPw02[clarity] / totalPw001 * 100.0;
                                sheet.Cells[filterRow + 18, 3].Value2 = percentPw01.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 18, 5].Value2 = clarityCounts02[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 18, 7].Value2 = clarityPw02[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeight001 = clarityWeightDict001.ContainsKey(clarity) ? clarityWeightDict001[clarity] : 0.000;
                                sheet.Cells[filterRow + 18, 4].Value2 = Math.Round(clarityWeight001, 3);

                                // Get the craft weight for this clarity
                                double clarityWeight002 = clarityWeightDict002.ContainsKey(clarity) ? clarityWeightDict002[clarity] : 0.000;
                                sheet.Cells[filterRow + 18, 6].Value2 = Math.Round(clarityWeight002, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDivided001 = clarityCounts02[clarity] / clarityWeightDict001[clarity];
                                sheet.Cells[filterRow + 18, 8].Value2 = Math.Round(clarityCountDivided001, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCount001 = clarityCounts02[clarity] / clarityPw02[clarity];
                                sheet.Cells[filterRow + 18, 9].Value2 = Math.Round(clarityPolishWtCount001, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCr001 = clarityPw02[clarity] / clarityWeightDict002[clarity] * 100;
                                sheet.Cells[filterRow + 18, 10].Value2 = percentPwCr001.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCr001 = clarityPw02[clarity] / clarityWeightDict001[clarity] * 100;
                                sheet.Cells[filterRow + 18, 11].Value2 = percentRoCr001.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 18, 12].Value2 = clarityDolarDict001[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolar001 = clarityDolarDict001[clarity] / clarityWeightDict001[clarity];
                                sheet.Cells[filterRow + 18, 13].Value2 = Math.Round(clarityRoughCrtDolar001, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolar001 = clarityDolarDict001[clarity] / clarityPw02[clarity];
                                sheet.Cells[filterRow + 18, 14].Value2 = Math.Round(clarityPolishCrtDolar001, 2);

                                serialNumber02++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRange2 = sheet.Range[sheet.Cells[filterRow - serialNumber02 + 18, 1], sheet.Cells[filterRow + 17, 14]];
                        tableRange2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRange2.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRange2.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 18, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 18, 1].HorizontalAlignment =
                            Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeight001 = clarityWeightDict001.Values.Sum();
                        sheet.Cells[filterRow + 18, 4].Value2 = Math.Round(totalRoughtWeight001, 2);

                        sheet.Cells[filterRow + 18, 5].Value2 = totalCount001;

                        double partWTTotal001 = clarityWeightDict002.Values.Sum();
                        sheet.Cells[filterRow + 18, 6].Value2 = Math.Round(partWTTotal001, 2);

                        sheet.Cells[filterRow + 18, 7].Value2 = Math.Round(totalPw001, 3);

                        double totalSize001 = (totalCount001 / totalRoughtWeight001);
                        sheet.Cells[filterRow + 18, 8].Value2 = totalSize001.ToString("0.00");

                        double polishSize001 = (totalCount001 / totalPw001);
                        sheet.Cells[filterRow + 18, 9].Value2 = polishSize001.ToString("0.00");

                        double crPwPercentage001 = (totalPw001 / partWTTotal001) * 100;
                        sheet.Cells[filterRow + 18, 10].Value2 = crPwPercentage001.ToString("0.00") + "%";

                        double pwPercentage001 = (totalPw001 / totalRoughtWeight001) * 100;
                        sheet.Cells[filterRow + 18, 11].Value2 = pwPercentage001.ToString("0.00") + "%";

                        double dolarTotal001 = clarityDolarDict001.Values.Sum();
                        sheet.Cells[filterRow + 18, 12].Value2 = Math.Round(dolarTotal001, 2);

                        double valueRough001 = (dolarTotal001 / totalRoughtWeight001);
                        sheet.Cells[filterRow + 18, 13].Value2 = valueRough001.ToString("0.00");

                        double valuePolish001 = (dolarTotal001 / totalPw001);
                        sheet.Cells[filterRow + 18, 14].Value2 = valuePolish001.ToString("0.00");

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRange3 = sheet.Range[sheet.Cells[filterRow + 18, 1], sheet.Cells[filterRow + 18, 14]];
                        //headerRange3.Font.Bold = true;
                        headerRange3.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRange3.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion

                    #region 3rd Sieve Round Method

                    if (checkBoxRunCode03.Checked)
                    {
                        sheet.Cells[filterRow + 21, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 21, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 21, 5].Value2 = "Diameter";

                        string updatedValue3 = textBox3.Text;
                        sheet.Cells[filterRow + 22, 3].Value2 = updatedValue3;

                        string updatedValue03 = txtWidthRange02.Text;
                        sheet.Cells[filterRow + 22, 5].Value2 = updatedValue03;

                        //sheet.Cells[filterRow + 22, 5].Value2 = "1.500 - 1.799";
                        txtWidthRange02.Text = (string)sheet.Cells[filterRow + 22, 5].Value2;

                        sheet.Cells[filterRow + 22, 1].Value2 = "ROUND";

                        //sheet.Cells[filterRow + 22, 3].Value2 = "+4.5 -6.5";
                        textBox3.Text = (string)sheet.Cells[filterRow + 22, 3].Value2;

                        sheet.Cells[filterRow + 23, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 23, 1].HorizontalAlignment =
                            Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 23, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 23, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 23, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 23, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 23, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 23, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 23, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 23, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 23, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 23, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 23, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 23, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 23, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRange0001 = sheet.Range[sheet.Cells[filterRow + 23, 1], sheet.Cells[filterRow + 23, 14]];
                        //headerRange0001.Font.Bold = true;
                        headerRange0001.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the width range from the txtWidthRange TextBox
                        string widthRangeText02 = txtWidthRange02.Text;
                        string[] widthRangeParts02 = widthRangeText02.Split('-');
                        double minWidth02 = 1.499;  // Default minimum width
                        double maxWidth02 = 1.799;  // Default maximum width

                        // Parse the width range values if the input is valid
                        if (widthRangeParts02.Length == 2 && double.TryParse(widthRangeParts02[0], out double parsedMinWidth02) && double.TryParse(widthRangeParts02[1], out double parsedMaxWidth02))
                        {
                            minWidth02 = parsedMinWidth02;
                            maxWidth02 = parsedMaxWidth02;
                        }

                        Dictionary<string, int> clarityCounts03 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPw03 = new Dictionary<string, double>();

                        Dictionary<string, List<(string shape, double width, string clarity, double partWT, string PoDolar)>> stoneWidths03 =
                            new Dictionary<string, List<(string shape, double width, string clarity, double partWT, string PoDolar)>>();

                        string prevStone03Name = "";
                        for (int t = 2; t <= rowCount; t++)
                        {
                            // Get the stone name, clarity, polish weight, part weight, and PoDolar values
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStone03Name;
                            string widthString = (range.Cells[t, 20].Value2 != null) ? range.Cells[t, 20].Value2.ToString() : prevStone03Name;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            double pw = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "round" && double.TryParse(widthString, out double width))
                            {
                                // Check if the width falls within the specified range
                                //if (width > 0.9 && width <= 1.249 && clarityValues.Contains(clarity))
                                if (width >= minWidth02 && width <= maxWidth02 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCounts03.ContainsKey(clarity))
                                        {
                                            clarityCounts03[clarity]++;
                                            clarityPw03[clarity] += pw;
                                        }
                                        else
                                        {
                                            clarityCounts03.Add(clarity, 1);
                                            clarityPw03.Add(clarity, pw);
                                        }

                                        // Add the width, clarity, polish weight, and part weight to the stoneWidths01 dictionary for the current stone name
                                        if (stoneWidths03.ContainsKey(stoneName))
                                        {
                                            stoneWidths03[stoneName].Add((shape, width, clarity, double.Parse(partWTString), PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double width, string clarity, double partWT, string PoDolar)> data = new List<(string shape, double width, string clarity, double partWT, string PoDolar)>();
                                            data.Add((shape, width, clarity, double.Parse(partWTString), PoDolar));
                                            //(shape, width, clarity, double.Parse(partWTString), PoDolar)                                
                                            stoneWidths03.Add(stoneName, data);
                                        }
                                    }
                                }
                            }

                            prevStone03Name = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING THIRD SIEVE FOR SHAPE 'ROUND' CHECKING IF THE WIDTH FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        Dictionary<string, double> clarityWeightDict0001 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> roundParts = partWts.Where(p => p.Item3.ToLower() == "round").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForWd0001 = roundParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item2 >= minWidth02 && p.Item2 <= maxWidth02)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDict0001.ContainsKey(clarity))
                                {
                                    clarityWeightDict0001[clarity] += clarityWeightForWd0001;
                                }
                                else
                                {
                                    clarityWeightDict0001[clarity] = clarityWeightForWd0001;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDict0002 = new Dictionary<string, double>();

                        // Iterate over the stone names in stoneWidths03
                        foreach (var stoneData in stoneWidths03.Values.SelectMany(data => data))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDict0002.ContainsKey(clarity))
                            {
                                clarityWeightDict0002[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDict0002[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDict0001 = new Dictionary<string, double>();

                        // Iterate over the stone names in stoneWidths03
                        foreach (var stoneData in stoneWidths03.Values.SelectMany(data => data))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDict0001.ContainsKey(clarity))
                                {
                                    clarityDolarDict0001[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDict0001[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCount0001 = clarityCounts03.Values.Sum();
                        double totalPw0001 = clarityPw03.Values.Sum();

                        int serialNumber03 = 1;

                        int totalClarityFilters03 = clarityValues.Count();
                        int currentClarityFilter03 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter03++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilter03 / totalClarityFilters03 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING THIRD 'ROUND' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCounts03.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 24, 1].Value2 = serialNumber03;
                                sheet.Cells[filterRow + 24, 1].HorizontalAlignment =
                                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 24, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPw001 = clarityPw03[clarity] / totalPw0001 * 100.0;
                                sheet.Cells[filterRow + 24, 3].Value2 = percentPw001.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 24, 5].Value2 = clarityCounts03[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 24, 7].Value2 = clarityPw03[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeight0001 = clarityWeightDict0001.ContainsKey(clarity) ? clarityWeightDict0001[clarity] : 0.000;
                                sheet.Cells[filterRow + 24, 4].Value2 = Math.Round(clarityWeight0001, 3);

                                // Get the craft weight for this clarity
                                double clarityWeight0002 = clarityWeightDict0002.ContainsKey(clarity) ? clarityWeightDict0002[clarity] : 0.000;
                                sheet.Cells[filterRow + 24, 6].Value2 = Math.Round(clarityWeight0002, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDivided0001 = clarityCounts03[clarity] / clarityWeightDict0001[clarity];
                                sheet.Cells[filterRow + 24, 8].Value2 = Math.Round(clarityCountDivided0001, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCount0001 = clarityCounts03[clarity] / clarityPw03[clarity];
                                sheet.Cells[filterRow + 24, 9].Value2 = Math.Round(clarityPolishWtCount0001, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCr0001 = clarityPw03[clarity] / clarityWeightDict0002[clarity] * 100;
                                sheet.Cells[filterRow + 24, 10].Value2 = percentPwCr0001.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCr0001 = clarityPw03[clarity] / clarityWeightDict0001[clarity] * 100;
                                sheet.Cells[filterRow + 24, 11].Value2 = percentRoCr0001.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 24, 12].Value2 = clarityDolarDict0001[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolar0001 = clarityDolarDict0001[clarity] / clarityWeightDict0001[clarity];
                                sheet.Cells[filterRow + 24, 13].Value2 = Math.Round(clarityRoughCrtDolar0001, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolar0001 = clarityDolarDict0001[clarity] / clarityPw03[clarity];
                                sheet.Cells[filterRow + 24, 14].Value2 = Math.Round(clarityPolishCrtDolar0001, 2);

                                serialNumber03++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRange3 = sheet.Range[sheet.Cells[filterRow - serialNumber03 + 24, 1], sheet.Cells[filterRow + 23, 14]];
                        tableRange3.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRange3.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRange3.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 24, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 24, 1].HorizontalAlignment =
                            Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeight0001 = clarityWeightDict0001.Values.Sum();
                        sheet.Cells[filterRow + 24, 4].Value2 = Math.Round(totalRoughtWeight0001, 2);

                        sheet.Cells[filterRow + 24, 5].Value2 = totalCount0001;

                        double partWTTotal0001 = clarityWeightDict0002.Values.Sum();
                        sheet.Cells[filterRow + 24, 6].Value2 = Math.Round(partWTTotal0001, 2);

                        sheet.Cells[filterRow + 24, 7].Value2 = Math.Round(totalPw0001, 3);

                        double totalSize0001 = (totalCount0001 / totalRoughtWeight0001);
                        sheet.Cells[filterRow + 24, 8].Value2 = totalSize0001.ToString("0.00");

                        double polishSize0001 = (totalCount0001 / totalPw0001);
                        sheet.Cells[filterRow + 24, 9].Value2 = polishSize0001.ToString("0.00");

                        double crPwPercentage0001 = (totalPw0001 / partWTTotal0001) * 100;
                        sheet.Cells[filterRow + 24, 10].Value2 = crPwPercentage0001.ToString("0.00") + "%";

                        double pwPercentage0001 = (totalPw0001 / totalRoughtWeight0001) * 100;
                        sheet.Cells[filterRow + 24, 11].Value2 = pwPercentage0001.ToString("0.00") + "%";

                        double dolarTotal0001 = clarityDolarDict0001.Values.Sum();
                        sheet.Cells[filterRow + 24, 12].Value2 = Math.Round(dolarTotal0001, 2);

                        double valueRough0001 = (dolarTotal0001 / totalRoughtWeight0001);
                        sheet.Cells[filterRow + 24, 13].Value2 = valueRough0001.ToString("0.00");

                        double valuePolish0001 = (dolarTotal0001 / totalPw0001);
                        sheet.Cells[filterRow + 24, 14].Value2 = valuePolish0001.ToString("0.00");

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRange4 = sheet.Range[sheet.Cells[filterRow + 24, 1], sheet.Cells[filterRow + 24, 14]];
                        //headerRange4.Font.Bold = true;
                        headerRange4.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRange4.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion

                    #region 4th Sieve Round Method

                    if (checkBoxRunCode04.Checked)
                    {
                        sheet.Cells[filterRow + 27, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 27, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 27, 5].Value2 = "Diameter";

                        string updatedValue4 = textBox4.Text;
                        sheet.Cells[filterRow + 28, 3].Value2 = updatedValue4;

                        string updatedValue04 = txtWidthRange03.Text;
                        sheet.Cells[filterRow + 28, 5].Value2 = updatedValue04;

                        //sheet.Cells[filterRow + 28, 5].Value2 = "1.800 - 2.099";
                        txtWidthRange03.Text = (string)sheet.Cells[filterRow + 28, 5].Value2;

                        sheet.Cells[filterRow + 28, 1].Value2 = "ROUND";

                        //sheet.Cells[filterRow + 28, 3].Value2 = "+6.5 -8";
                        textBox4.Text = (string)sheet.Cells[filterRow + 28, 3].Value2;

                        sheet.Cells[filterRow + 29, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 29, 1].HorizontalAlignment =
                            Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 29, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 29, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 29, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 29, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 29, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 29, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 29, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 29, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 29, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 29, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 29, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 29, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 29, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRange00001 = sheet.Range[sheet.Cells[filterRow + 29, 1], sheet.Cells[filterRow + 29, 14]];
                        //headerRange00001.Font.Bold = true;
                        headerRange00001.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the width range from the txtWidthRange TextBox
                        string widthRangeText03 = txtWidthRange03.Text;
                        string[] widthRangeParts03 = widthRangeText03.Split('-');
                        double minWidth03 = 1.799;  // Default minimum width
                        double maxWidth03 = 2.099;  // Default maximum width

                        // Parse the width range values if the input is valid
                        if (widthRangeParts03.Length == 2 && double.TryParse(widthRangeParts03[0], out double parsedMinWidth03) && double.TryParse(widthRangeParts03[1], out double parsedMaxWidth03))
                        {
                            minWidth03 = parsedMinWidth03;
                            maxWidth03 = parsedMaxWidth03;
                        }

                        Dictionary<string, int> clarityCounts04 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPw04 = new Dictionary<string, double>();

                        Dictionary<string, List<(string shape, double width, string clarity, double partWT, string PoDolar)>> stoneWidths04 =
                            new Dictionary<string, List<(string shape, double width, string clarity, double partWT, string PoDolar)>>();

                        string prevStone04Name = "";
                        for (int t = 2; t <= rowCount; t++)
                        {
                            // Get the stone name, clarity, polish weight, part weight, and PoDolar values
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStone04Name;
                            string widthString = (range.Cells[t, 20].Value2 != null) ? range.Cells[t, 20].Value2.ToString() : prevStone04Name;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            double pw = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "round" && double.TryParse(widthString, out double width))
                            {
                                // Check if the width falls within the specified range
                                //if (width > 0.9 && width <= 1.249 && clarityValues.Contains(clarity))
                                if (width >= minWidth03 && width <= maxWidth03 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCounts04.ContainsKey(clarity))
                                        {
                                            clarityCounts04[clarity]++;
                                            clarityPw04[clarity] += pw;
                                        }
                                        else
                                        {
                                            clarityCounts04.Add(clarity, 1);
                                            clarityPw04.Add(clarity, pw);
                                        }

                                        // Add the width, clarity, polish weight, and part weight to the stoneWidths01 dictionary for the current stone name
                                        if (stoneWidths04.ContainsKey(stoneName))
                                        {
                                            stoneWidths04[stoneName].Add((shape, width, clarity, double.Parse(partWTString), PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double width, string clarity, double partWT, string PoDolar)> data = new List<(string shape, double width, string clarity, double partWT, string PoDolar)>();
                                            data.Add((shape, width, clarity, double.Parse(partWTString), PoDolar));
                                            //(shape, width, clarity, double.Parse(partWTString), PoDolar)                                
                                            stoneWidths04.Add(stoneName, data);
                                        }
                                    }
                                }
                            }

                            prevStone04Name = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING FOURTH SIEVE FOR SHAPE 'ROUND' CHECKING IF THE WIDTH FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        Dictionary<string, double> clarityWeightDict00001 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> roundParts = partWts.Where(p => p.Item3.ToLower() == "round").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForWd00001 = roundParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item2 >= minWidth03 && p.Item2 <= maxWidth03)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDict00001.ContainsKey(clarity))
                                {
                                    clarityWeightDict00001[clarity] += clarityWeightForWd00001;
                                }
                                else
                                {
                                    clarityWeightDict00001[clarity] = clarityWeightForWd00001;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDict00002 = new Dictionary<string, double>();

                        // Iterate over the stone names in stoneWidths04
                        foreach (var stoneData in stoneWidths04.Values.SelectMany(data => data))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDict00002.ContainsKey(clarity))
                            {
                                clarityWeightDict00002[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDict00002[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDict00001 = new Dictionary<string, double>();

                        // Iterate over the stone names in stoneWidths04
                        foreach (var stoneData in stoneWidths04.Values.SelectMany(data => data))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDict00001.ContainsKey(clarity))
                                {
                                    clarityDolarDict00001[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDict00001[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCount00001 = clarityCounts04.Values.Sum();
                        double totalPw00001 = clarityPw04.Values.Sum();

                        int serialNumber04 = 1;

                        int totalClarityFilters04 = clarityValues.Count();
                        int currentClarityFilter04 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter04++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilter04 / totalClarityFilters04 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING FOURTH 'ROUND' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCounts04.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 30, 1].Value2 = serialNumber04;
                                sheet.Cells[filterRow + 30, 1].HorizontalAlignment =
                                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 30, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPw0001 = clarityPw04[clarity] / totalPw00001 * 100.0;
                                sheet.Cells[filterRow + 30, 3].Value2 = percentPw0001.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 30, 5].Value2 = clarityCounts04[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 30, 7].Value2 = clarityPw04[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeight00001 = clarityWeightDict00001.ContainsKey(clarity) ? clarityWeightDict00001[clarity] : 0.000;
                                sheet.Cells[filterRow + 30, 4].Value2 = Math.Round(clarityWeight00001, 3);

                                // Get the craft weight for this clarity
                                double clarityWeight00002 = clarityWeightDict00002.ContainsKey(clarity) ? clarityWeightDict00002[clarity] : 0.000;
                                sheet.Cells[filterRow + 30, 6].Value2 = Math.Round(clarityWeight00002, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDivided00001 = clarityCounts04[clarity] / clarityWeightDict00001[clarity];
                                sheet.Cells[filterRow + 30, 8].Value2 = Math.Round(clarityCountDivided00001, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCount00001 = clarityCounts04[clarity] / clarityPw04[clarity];
                                sheet.Cells[filterRow + 30, 9].Value2 = Math.Round(clarityPolishWtCount00001, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCr00001 = clarityPw04[clarity] / clarityWeightDict00002[clarity] * 100;
                                sheet.Cells[filterRow + 30, 10].Value2 = percentPwCr00001.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCr00001 = clarityPw04[clarity] / clarityWeightDict00001[clarity] * 100;
                                sheet.Cells[filterRow + 30, 11].Value2 = percentRoCr00001.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 30, 12].Value2 = clarityDolarDict00001[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolar00001 = clarityDolarDict00001[clarity] / clarityWeightDict00001[clarity];
                                sheet.Cells[filterRow + 30, 13].Value2 = Math.Round(clarityRoughCrtDolar00001, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolar00001 = clarityDolarDict00001[clarity] / clarityPw04[clarity];
                                sheet.Cells[filterRow + 30, 14].Value2 = Math.Round(clarityPolishCrtDolar00001, 2);

                                serialNumber04++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRange4 = sheet.Range[sheet.Cells[filterRow - serialNumber04 + 30, 1], sheet.Cells[filterRow + 29, 14]];
                        tableRange4.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRange4.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRange4.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 30, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 30, 1].HorizontalAlignment =
                            Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeight00001 = clarityWeightDict00001.Values.Sum();
                        sheet.Cells[filterRow + 30, 4].Value2 = Math.Round(totalRoughtWeight00001, 2);

                        sheet.Cells[filterRow + 30, 5].Value2 = totalCount00001;

                        double partWTTotal00001 = clarityWeightDict00002.Values.Sum();
                        sheet.Cells[filterRow + 30, 6].Value2 = Math.Round(partWTTotal00001, 2);

                        sheet.Cells[filterRow + 30, 7].Value2 = Math.Round(totalPw00001, 3);

                        double totalSize00001 = (totalCount00001 / totalRoughtWeight00001);
                        sheet.Cells[filterRow + 30, 8].Value2 = totalSize00001.ToString("0.00");

                        double polishSize00001 = (totalCount00001 / totalPw00001);
                        sheet.Cells[filterRow + 30, 9].Value2 = polishSize00001.ToString("0.00");

                        double crPwPercentage00001 = (totalPw00001 / partWTTotal00001) * 100;
                        sheet.Cells[filterRow + 30, 10].Value2 = crPwPercentage00001.ToString("0.00") + "%";

                        double pwPercentage00001 = (totalPw00001 / totalRoughtWeight00001) * 100;
                        sheet.Cells[filterRow + 30, 11].Value2 = pwPercentage00001.ToString("0.00") + "%";

                        double dolarTotal00001 = clarityDolarDict00001.Values.Sum();
                        sheet.Cells[filterRow + 30, 12].Value2 = Math.Round(dolarTotal00001, 2);

                        double valueRough00001 = (dolarTotal00001 / totalRoughtWeight00001);
                        sheet.Cells[filterRow + 30, 13].Value2 = valueRough00001.ToString("0.00");

                        double valuePolish00001 = (dolarTotal00001 / totalPw00001);
                        sheet.Cells[filterRow + 30, 14].Value2 = valuePolish00001.ToString("0.00");

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRange5 = sheet.Range[sheet.Cells[filterRow + 30, 1], sheet.Cells[filterRow + 30, 14]];
                        //headerRange5.Font.Bold = true;
                        headerRange5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRange5.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion

                    #region 5th Sieve Round Method

                    if (checkBoxRunCode05.Checked)
                    {
                        sheet.Cells[filterRow + 33, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 33, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 33, 5].Value2 = "Diameter";

                        string updatedValue5 = textBox5.Text;
                        sheet.Cells[filterRow + 34, 3].Value2 = updatedValue5;

                        string updatedValue05 = txtWidthRange04.Text;
                        sheet.Cells[filterRow + 34, 5].Value2 = updatedValue05;

                        //sheet.Cells[filterRow + 34, 5].Value2 = "2.100 - 2.699";
                        txtWidthRange04.Text = (string)sheet.Cells[filterRow + 34, 5].Value2;

                        sheet.Cells[filterRow + 34, 1].Value2 = "ROUND";

                        //sheet.Cells[filterRow + 34, 3].Value2 = "+8 -11";
                        textBox5.Text = (string)sheet.Cells[filterRow + 34, 3].Value2;

                        sheet.Cells[filterRow + 35, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 35, 1].HorizontalAlignment =
                            Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 35, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 35, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 35, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 35, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 35, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 35, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 35, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 35, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 35, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 35, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 35, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 35, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 35, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRange000001 = sheet.Range[sheet.Cells[filterRow + 35, 1], sheet.Cells[filterRow + 35, 14]];
                        //headerRange000001.Font.Bold = true;
                        headerRange000001.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the width range from the txtWidthRange TextBox
                        string widthRangeText04 = txtWidthRange04.Text;
                        string[] widthRangeParts04 = widthRangeText04.Split('-');
                        double minWidth04 = 2.099;  // Default minimum width
                        double maxWidth04 = 2.699;  // Default maximum width

                        // Parse the width range values if the input is valid
                        if (widthRangeParts04.Length == 2 && double.TryParse(widthRangeParts04[0], out double parsedMinWidth04) && double.TryParse(widthRangeParts04[1], out double parsedMaxWidth04))
                        {
                            minWidth04 = parsedMinWidth04;
                            maxWidth04 = parsedMaxWidth04;
                        }

                        Dictionary<string, int> clarityCounts05 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPw05 = new Dictionary<string, double>();

                        Dictionary<string, List<(string shape, double width, string clarity, double partWT, string PoDolar)>> stoneWidths05 =
                            new Dictionary<string, List<(string shape, double width, string clarity, double partWT, string PoDolar)>>();

                        string prevStone05Name = "";
                        for (int t = 2; t <= rowCount; t++)
                        {
                            // Get the stone name, clarity, polish weight, part weight, and PoDolar values
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStone05Name;
                            string widthString = (range.Cells[t, 20].Value2 != null) ? range.Cells[t, 20].Value2.ToString() : prevStone05Name;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            double pw = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "round" && double.TryParse(widthString, out double width))
                            {
                                // Check if the width falls within the specified range
                                //if (width > 0.9 && width <= 1.249 && clarityValues.Contains(clarity))
                                if (width >= minWidth04 && width <= maxWidth04 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCounts05.ContainsKey(clarity))
                                        {
                                            clarityCounts05[clarity]++;
                                            clarityPw05[clarity] += pw;
                                        }
                                        else
                                        {
                                            clarityCounts05.Add(clarity, 1);
                                            clarityPw05.Add(clarity, pw);
                                        }

                                        // Add the width, clarity, polish weight, and part weight to the stoneWidths01 dictionary for the current stone name
                                        if (stoneWidths05.ContainsKey(stoneName))
                                        {
                                            stoneWidths05[stoneName].Add((shape, width, clarity, double.Parse(partWTString), PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double width, string clarity, double partWT, string PoDolar)> data = new List<(string shape, double width, string clarity, double partWT, string PoDolar)>();
                                            data.Add((shape, width, clarity, double.Parse(partWTString), PoDolar));
                                            //(shape, width, clarity, double.Parse(partWTString), PoDolar)                                
                                            stoneWidths05.Add(stoneName, data);
                                        }
                                    }
                                }
                            }

                            prevStone05Name = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING FIFTH SIEVE FOR SHAPE 'ROUND' CHECKING IF THE WIDTH FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        Dictionary<string, double> clarityWeightDict000001 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> roundParts = partWts.Where(p => p.Item3.ToLower() == "round").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForWd000001 = roundParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item2 >= minWidth04 && p.Item2 <= maxWidth04)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDict000001.ContainsKey(clarity))
                                {
                                    clarityWeightDict000001[clarity] += clarityWeightForWd000001;
                                }
                                else
                                {
                                    clarityWeightDict000001[clarity] = clarityWeightForWd000001;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDict000002 = new Dictionary<string, double>();

                        // Iterate over the stone names in stoneWidths05
                        foreach (var stoneData in stoneWidths05.Values.SelectMany(data => data))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDict000002.ContainsKey(clarity))
                            {
                                clarityWeightDict000002[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDict000002[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDict000001 = new Dictionary<string, double>();

                        // Iterate over the stone names in stoneWidths05
                        foreach (var stoneData in stoneWidths05.Values.SelectMany(data => data))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDict000001.ContainsKey(clarity))
                                {
                                    clarityDolarDict000001[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDict000001[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCount000001 = clarityCounts05.Values.Sum();
                        double totalPw000001 = clarityPw05.Values.Sum();

                        int serialNumber05 = 1;

                        int totalClarityFilters05 = clarityValues.Count();
                        int currentClarityFilter05 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter05++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilter05 / totalClarityFilters05 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING FIFTH 'ROUND' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCounts05.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 36, 1].Value2 = serialNumber05;
                                sheet.Cells[filterRow + 36, 1].HorizontalAlignment =
                                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 36, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPw00001 = clarityPw05[clarity] / totalPw000001 * 100.0;
                                sheet.Cells[filterRow + 36, 3].Value2 = percentPw00001.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 36, 5].Value2 = clarityCounts05[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 36, 7].Value2 = clarityPw05[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeight000001 = clarityWeightDict000001.ContainsKey(clarity) ? clarityWeightDict000001[clarity] : 0.000;
                                sheet.Cells[filterRow + 36, 4].Value2 = Math.Round(clarityWeight000001, 3);

                                // Get the craft weight for this clarity
                                double clarityWeight000002 = clarityWeightDict000002.ContainsKey(clarity) ? clarityWeightDict000002[clarity] : 0.000;
                                sheet.Cells[filterRow + 36, 6].Value2 = Math.Round(clarityWeight000002, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDivided000001 = clarityCounts05[clarity] / clarityWeightDict000001[clarity];
                                sheet.Cells[filterRow + 36, 8].Value2 = Math.Round(clarityCountDivided000001, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCount000001 = clarityCounts05[clarity] / clarityPw05[clarity];
                                sheet.Cells[filterRow + 36, 9].Value2 = Math.Round(clarityPolishWtCount000001, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCr000001 = clarityPw05[clarity] / clarityWeightDict000002[clarity] * 100;
                                sheet.Cells[filterRow + 36, 10].Value2 = percentPwCr000001.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCr000001 = clarityPw05[clarity] / clarityWeightDict000001[clarity] * 100;
                                sheet.Cells[filterRow + 36, 11].Value2 = percentRoCr000001.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 36, 12].Value2 = clarityDolarDict000001[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolar000001 = clarityDolarDict000001[clarity] / clarityWeightDict000001[clarity];
                                sheet.Cells[filterRow + 36, 13].Value2 = Math.Round(clarityRoughCrtDolar000001, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolar000001 = clarityDolarDict000001[clarity] / clarityPw05[clarity];
                                sheet.Cells[filterRow + 36, 14].Value2 = Math.Round(clarityPolishCrtDolar000001, 2);

                                serialNumber05++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRange5 = sheet.Range[sheet.Cells[filterRow - serialNumber05 + 36, 1], sheet.Cells[filterRow + 35, 14]];
                        tableRange5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRange5.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRange5.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 36, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 36, 1].HorizontalAlignment =
                            Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeight000001 = clarityWeightDict000001.Values.Sum();
                        sheet.Cells[filterRow + 36, 4].Value2 = Math.Round(totalRoughtWeight000001, 2);

                        sheet.Cells[filterRow + 36, 5].Value2 = totalCount000001;

                        double partWTTotal000001 = clarityWeightDict000002.Values.Sum();
                        sheet.Cells[filterRow + 36, 6].Value2 = Math.Round(partWTTotal000001, 2);

                        sheet.Cells[filterRow + 36, 7].Value2 = Math.Round(totalPw000001, 3);

                        double totalSize000001 = (totalCount000001 / totalRoughtWeight000001);
                        sheet.Cells[filterRow + 36, 8].Value2 = totalSize000001.ToString("0.00");

                        double polishSize000001 = (totalCount000001 / totalPw000001);
                        sheet.Cells[filterRow + 36, 9].Value2 = polishSize000001.ToString("0.00");

                        double crPwPercentage000001 = (totalPw000001 / partWTTotal000001) * 100;
                        sheet.Cells[filterRow + 36, 10].Value2 = crPwPercentage000001.ToString("0.00") + "%";

                        double pwPercentage000001 = (totalPw000001 / totalRoughtWeight000001) * 100;
                        sheet.Cells[filterRow + 36, 11].Value2 = pwPercentage000001.ToString("0.00") + "%";

                        double dolarTotal000001 = clarityDolarDict000001.Values.Sum();
                        sheet.Cells[filterRow + 36, 12].Value2 = Math.Round(dolarTotal000001, 2);

                        double valueRough000001 = (dolarTotal000001 / totalRoughtWeight000001);
                        sheet.Cells[filterRow + 36, 13].Value2 = valueRough000001.ToString("0.00");

                        double valuePolish000001 = (dolarTotal000001 / totalPw000001);
                        sheet.Cells[filterRow + 36, 14].Value2 = valuePolish000001.ToString("0.00");

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRange6 = sheet.Range[sheet.Cells[filterRow + 36, 1], sheet.Cells[filterRow + 36, 14]];
                        //headerRange6.Font.Bold = true;
                        headerRange6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRange6.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion

                    #region 6th Sieve Round Method

                    if (checkBoxRunCode06.Checked)
                    {
                        sheet.Cells[filterRow + 39, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 39, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 39, 5].Value2 = "Diameter";

                        string updatedValue6 = textBox6.Text;
                        sheet.Cells[filterRow + 40, 3].Value2 = updatedValue6;

                        string updatedValue06 = txtWidthRange05.Text;
                        sheet.Cells[filterRow + 40, 5].Value2 = updatedValue06;

                        //sheet.Cells[filterRow + 40, 5].Value2 = "2.700 - 3.299";
                        txtWidthRange05.Text = (string)sheet.Cells[filterRow + 40, 5].Value2;

                        sheet.Cells[filterRow + 40, 1].Value2 = "ROUND";

                        //sheet.Cells[filterRow + 40, 3].Value2 = "+11 -14";
                        textBox6.Text = (string)sheet.Cells[filterRow + 40, 3].Value2;

                        sheet.Cells[filterRow + 41, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 41, 1].HorizontalAlignment =
                            Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 41, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 41, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 41, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 41, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 41, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 41, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 41, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 41, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 41, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 41, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 41, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 41, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 41, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRange0000001 = sheet.Range[sheet.Cells[filterRow + 41, 1], sheet.Cells[filterRow + 41, 14]];
                        //headerRange0000001.Font.Bold = true;
                        headerRange0000001.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the width range from the txtWidthRange TextBox
                        string widthRangeText05 = txtWidthRange05.Text;
                        string[] widthRangeParts05 = widthRangeText05.Split('-');
                        double minWidth05 = 2.699;  // Default minimum width
                        double maxWidth05 = 3.299;  // Default maximum width

                        // Parse the width range values if the input is valid
                        if (widthRangeParts05.Length == 2 && double.TryParse(widthRangeParts05[0], out double parsedMinWidth05) && double.TryParse(widthRangeParts05[1], out double parsedMaxWidth05))
                        {
                            minWidth05 = parsedMinWidth05;
                            maxWidth05 = parsedMaxWidth05;
                        }

                        Dictionary<string, int> clarityCounts06 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPw06 = new Dictionary<string, double>();

                        Dictionary<string, List<(string shape, double width, string clarity, double partWT, string PoDolar)>> stoneWidths06 =
                            new Dictionary<string, List<(string shape, double width, string clarity, double partWT, string PoDolar)>>();

                        string prevStone06Name = "";
                        for (int t = 2; t <= rowCount; t++)
                        {
                            // Get the stone name, clarity, polish weight, part weight, and PoDolar values
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStone06Name;
                            string widthString = (range.Cells[t, 20].Value2 != null) ? range.Cells[t, 20].Value2.ToString() : prevStone06Name;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            double pw = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "round" && double.TryParse(widthString, out double width))
                            {
                                // Check if the width falls within the specified range
                                //if (width > 0.9 && width <= 1.249 && clarityValues.Contains(clarity))
                                if (width >= minWidth05 && width <= maxWidth05 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCounts06.ContainsKey(clarity))
                                        {
                                            clarityCounts06[clarity]++;
                                            clarityPw06[clarity] += pw;
                                        }
                                        else
                                        {
                                            clarityCounts06.Add(clarity, 1);
                                            clarityPw06.Add(clarity, pw);
                                        }

                                        // Add the width, clarity, polish weight, and part weight to the stoneWidths01 dictionary for the current stone name
                                        if (stoneWidths06.ContainsKey(stoneName))
                                        {
                                            stoneWidths06[stoneName].Add((shape, width, clarity, double.Parse(partWTString), PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double width, string clarity, double partWT, string PoDolar)> data = new List<(string shape, double width, string clarity, double partWT, string PoDolar)>();
                                            data.Add((shape, width, clarity, double.Parse(partWTString), PoDolar));
                                            //(shape, width, clarity, double.Parse(partWTString), PoDolar)                                
                                            stoneWidths06.Add(stoneName, data);
                                        }
                                    }
                                }
                            }

                            prevStone06Name = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING SIXTH SIEVE FOR SHAPE 'ROUND' CHECKING IF THE WIDTH FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        Dictionary<string, double> clarityWeightDict0000001 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> roundParts = partWts.Where(p => p.Item3.ToLower() == "round").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForWd0000001 = roundParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item2 >= minWidth05 && p.Item2 <= maxWidth05)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDict0000001.ContainsKey(clarity))
                                {
                                    clarityWeightDict0000001[clarity] += clarityWeightForWd0000001;
                                }
                                else
                                {
                                    clarityWeightDict0000001[clarity] = clarityWeightForWd0000001;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDict0000002 = new Dictionary<string, double>();

                        // Iterate over the stone names in stoneWidths05
                        foreach (var stoneData in stoneWidths06.Values.SelectMany(data => data))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDict0000002.ContainsKey(clarity))
                            {
                                clarityWeightDict0000002[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDict0000002[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDict0000001 = new Dictionary<string, double>();

                        // Iterate over the stone names in stoneWidths05
                        foreach (var stoneData in stoneWidths06.Values.SelectMany(data => data))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDict0000001.ContainsKey(clarity))
                                {
                                    clarityDolarDict0000001[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDict0000001[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCount0000001 = clarityCounts06.Values.Sum();
                        double totalPw0000001 = clarityPw06.Values.Sum();

                        int serialNumber06 = 1;

                        int totalClarityFilters06 = clarityValues.Count();
                        int currentClarityFilter06 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter06++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilter06 / totalClarityFilters06 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING SIXTH 'ROUND' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCounts06.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 42, 1].Value2 = serialNumber06;
                                sheet.Cells[filterRow + 42, 1].HorizontalAlignment =
                                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 42, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPw000001 = clarityPw06[clarity] / totalPw0000001 * 100.0;
                                sheet.Cells[filterRow + 42, 3].Value2 = percentPw000001.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 42, 5].Value2 = clarityCounts06[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 42, 7].Value2 = clarityPw06[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeight0000001 = clarityWeightDict0000001.ContainsKey(clarity) ? clarityWeightDict0000001[clarity] : 0.000;
                                sheet.Cells[filterRow + 42, 4].Value2 = Math.Round(clarityWeight0000001, 3);

                                // Get the craft weight for this clarity
                                double clarityWeight0000002 = clarityWeightDict0000002.ContainsKey(clarity) ? clarityWeightDict0000002[clarity] : 0.000;
                                sheet.Cells[filterRow + 42, 6].Value2 = Math.Round(clarityWeight0000002, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDivided0000001 = clarityCounts06[clarity] / clarityWeightDict0000001[clarity];
                                sheet.Cells[filterRow + 42, 8].Value2 = Math.Round(clarityCountDivided0000001, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCount0000001 = clarityCounts06[clarity] / clarityPw06[clarity];
                                sheet.Cells[filterRow + 42, 9].Value2 = Math.Round(clarityPolishWtCount0000001, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCr0000001 = clarityPw06[clarity] / clarityWeightDict0000002[clarity] * 100;
                                sheet.Cells[filterRow + 42, 10].Value2 = percentPwCr0000001.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCr0000001 = clarityPw06[clarity] / clarityWeightDict0000001[clarity] * 100;
                                sheet.Cells[filterRow + 42, 11].Value2 = percentRoCr0000001.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 42, 12].Value2 = clarityDolarDict0000001[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolar0000001 = clarityDolarDict0000001[clarity] / clarityWeightDict0000001[clarity];
                                sheet.Cells[filterRow + 42, 13].Value2 = Math.Round(clarityRoughCrtDolar0000001, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolar0000001 = clarityDolarDict0000001[clarity] / clarityPw06[clarity];
                                sheet.Cells[filterRow + 42, 14].Value2 = Math.Round(clarityPolishCrtDolar0000001, 2);

                                serialNumber06++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRange6 = sheet.Range[sheet.Cells[filterRow - serialNumber06 + 42, 1], sheet.Cells[filterRow + 41, 14]];
                        tableRange6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRange6.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRange6.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 42, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 42, 1].HorizontalAlignment =
                            Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeight0000001 = clarityWeightDict0000001.Values.Sum();
                        sheet.Cells[filterRow + 42, 4].Value2 = Math.Round(totalRoughtWeight0000001, 2);

                        sheet.Cells[filterRow + 42, 5].Value2 = totalCount0000001;

                        double partWTTotal0000001 = clarityWeightDict0000002.Values.Sum();
                        sheet.Cells[filterRow + 42, 6].Value2 = Math.Round(partWTTotal0000001, 2);

                        sheet.Cells[filterRow + 42, 7].Value2 = Math.Round(totalPw0000001, 3);

                        double totalSize0000001 = (totalCount0000001 / totalRoughtWeight0000001);
                        sheet.Cells[filterRow + 42, 8].Value2 = totalSize0000001.ToString("0.00");

                        double polishSize0000001 = (totalCount0000001 / totalPw0000001);
                        sheet.Cells[filterRow + 42, 9].Value2 = polishSize0000001.ToString("0.00");

                        double crPwPercentage0000001 = (totalPw0000001 / partWTTotal0000001) * 100;
                        sheet.Cells[filterRow + 42, 10].Value2 = crPwPercentage0000001.ToString("0.00") + "%";

                        double pwPercentage0000001 = (totalPw0000001 / totalRoughtWeight0000001) * 100;
                        sheet.Cells[filterRow + 42, 11].Value2 = pwPercentage0000001.ToString("0.00") + "%";

                        double dolarTotal0000001 = clarityDolarDict0000001.Values.Sum();
                        sheet.Cells[filterRow + 42, 12].Value2 = Math.Round(dolarTotal0000001, 2);

                        double valueRough0000001 = (dolarTotal0000001 / totalRoughtWeight0000001);
                        sheet.Cells[filterRow + 42, 13].Value2 = valueRough0000001.ToString("0.00");

                        double valuePolish0000001 = (dolarTotal0000001 / totalPw0000001);
                        sheet.Cells[filterRow + 42, 14].Value2 = valuePolish0000001.ToString("0.00");

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRange7 = sheet.Range[sheet.Cells[filterRow + 42, 1], sheet.Cells[filterRow + 42, 14]];
                        //headerRange7.Font.Bold = true;
                        headerRange7.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRange7.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion

                    #region 7th Sieve Round Method

                    if (checkBoxRunCode07.Checked)
                    {
                        sheet.Cells[filterRow + 45, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 45, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 45, 5].Value2 = "Diameter";

                        string updatedValue7 = textBox7.Text;
                        sheet.Cells[filterRow + 46, 3].Value2 = updatedValue7;

                        string updatedValue07 = txtWidthRange06.Text;
                        sheet.Cells[filterRow + 46, 5].Value2 = updatedValue07;

                        //sheet.Cells[filterRow + 46, 5].Value2 = "3.300 - 4.499";
                        txtWidthRange06.Text = (string)sheet.Cells[filterRow + 46, 5].Value2;

                        sheet.Cells[filterRow + 46, 1].Value2 = "ROUND";

                        //sheet.Cells[filterRow + 46, 3].Value2 = "+14 -20";
                        textBox7.Text = (string)sheet.Cells[filterRow + 46, 3].Value2;

                        sheet.Cells[filterRow + 47, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 47, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 47, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 47, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 47, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 47, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 47, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 47, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 47, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 47, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 47, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 47, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 47, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 47, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 47, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRange00000001 = sheet.Range[sheet.Cells[filterRow + 47, 1], sheet.Cells[filterRow + 47, 14]];
                        headerRange00000001.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the width range from the txtWidthRange TextBox
                        string widthRangeText06 = txtWidthRange06.Text;
                        string[] widthRangeParts06 = widthRangeText06.Split('-');
                        double minWidth06 = 3.299;  // Default minimum width
                        double maxWidth06 = 4.499;  // Default maximum width

                        // Parse the width range values if the input is valid
                        if (widthRangeParts06.Length == 2 && double.TryParse(widthRangeParts06[0], out double parsedMinWidth06) && double.TryParse(widthRangeParts06[1], out double parsedMaxWidth06))
                        {
                            minWidth06 = parsedMinWidth06;
                            maxWidth06 = parsedMaxWidth06;
                        }

                        Dictionary<string, int> clarityCounts07 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPw07 = new Dictionary<string, double>();

                        Dictionary<string, List<(string shape, double width, string clarity, double partWT, string PoDolar)>> stoneWidths07 =
                            new Dictionary<string, List<(string shape, double width, string clarity, double partWT, string PoDolar)>>();

                        string prevStone07Name = "";
                        for (int t = 2; t <= rowCount; t++)
                        {
                            // Get the stone name, clarity, polish weight, part weight, and PoDolar values
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStone07Name;
                            string widthString = (range.Cells[t, 20].Value2 != null) ? range.Cells[t, 20].Value2.ToString() : prevStone07Name;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            double pw = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "round" && double.TryParse(widthString, out double width))
                            {
                                // Check if the width falls within the specified range
                                //if (width > 0.9 && width <= 1.249 && clarityValues.Contains(clarity))
                                if (width >= minWidth06 && width <= maxWidth06 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCounts07.ContainsKey(clarity))
                                        {
                                            clarityCounts07[clarity]++;
                                            clarityPw07[clarity] += pw;
                                        }
                                        else
                                        {
                                            clarityCounts07.Add(clarity, 1);
                                            clarityPw07.Add(clarity, pw);
                                        }

                                        // Add the width, clarity, polish weight, and part weight to the stoneWidths01 dictionary for the current stone name
                                        if (stoneWidths07.ContainsKey(stoneName))
                                        {
                                            stoneWidths07[stoneName].Add((shape, width, clarity, double.Parse(partWTString), PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double width, string clarity, double partWT, string PoDolar)> data = new List<(string shape, double width, string clarity, double partWT, string PoDolar)>();
                                            data.Add((shape, width, clarity, double.Parse(partWTString), PoDolar));
                                            stoneWidths07.Add(stoneName, data);
                                        }
                                    }
                                }
                            }

                            prevStone07Name = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING SEVENTH SIEVE FOR SHAPE 'ROUND' CHECKING IF THE WIDTH FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        Dictionary<string, double> clarityWeightDict00000001 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> roundParts = partWts.Where(p => p.Item3.ToLower() == "round").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForWd00000001 = roundParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item2 >= minWidth06 && p.Item2 <= maxWidth06)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDict00000001.ContainsKey(clarity))
                                {
                                    clarityWeightDict00000001[clarity] += clarityWeightForWd00000001;
                                }
                                else
                                {
                                    clarityWeightDict00000001[clarity] = clarityWeightForWd00000001;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDict00000002 = new Dictionary<string, double>();

                        // Iterate over the stone names in stoneWidths05
                        foreach (var stoneData in stoneWidths07.Values.SelectMany(data => data))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDict00000002.ContainsKey(clarity))
                            {
                                clarityWeightDict00000002[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDict00000002[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDict00000001 = new Dictionary<string, double>();

                        // Iterate over the stone names in stoneWidths05
                        foreach (var stoneData in stoneWidths07.Values.SelectMany(data => data))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDict00000001.ContainsKey(clarity))
                                {
                                    clarityDolarDict00000001[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDict00000001[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCount00000001 = clarityCounts07.Values.Sum();
                        double totalPw00000001 = clarityPw07.Values.Sum();

                        int serialNumber07 = 1;

                        int totalClarityFilters07 = clarityValues.Count();
                        int currentClarityFilter07 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter07++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilter07 / totalClarityFilters07 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING SEVENTH 'ROUND' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCounts07.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 48, 1].Value2 = serialNumber07;
                                sheet.Cells[filterRow + 48, 1].HorizontalAlignment =
                                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 48, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPw0000001 = clarityPw07[clarity] / totalPw00000001 * 100.0;
                                sheet.Cells[filterRow + 48, 3].Value2 = percentPw0000001.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 48, 5].Value2 = clarityCounts07[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 48, 7].Value2 = clarityPw07[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeight00000001 = clarityWeightDict00000001.ContainsKey(clarity) ? clarityWeightDict00000001[clarity] : 0.000;
                                sheet.Cells[filterRow + 48, 4].Value2 = Math.Round(clarityWeight00000001, 3);

                                // Get the craft weight for this clarity
                                double clarityWeight00000002 = clarityWeightDict00000002.ContainsKey(clarity) ? clarityWeightDict00000002[clarity] : 0.000;
                                sheet.Cells[filterRow + 48, 6].Value2 = Math.Round(clarityWeight00000002, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDivided00000001 = clarityCounts07[clarity] / clarityWeightDict00000001[clarity];
                                sheet.Cells[filterRow + 48, 8].Value2 = Math.Round(clarityCountDivided00000001, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCount00000001 = clarityCounts07[clarity] / clarityPw07[clarity];
                                sheet.Cells[filterRow + 48, 9].Value2 = Math.Round(clarityPolishWtCount00000001, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCr00000001 = clarityPw07[clarity] / clarityWeightDict00000002[clarity] * 100;
                                sheet.Cells[filterRow + 48, 10].Value2 = percentPwCr00000001.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCr00000001 = clarityPw07[clarity] / clarityWeightDict00000001[clarity] * 100;
                                sheet.Cells[filterRow + 48, 11].Value2 = percentRoCr00000001.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 48, 12].Value2 = clarityDolarDict00000001[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolar00000001 = clarityDolarDict00000001[clarity] / clarityWeightDict00000001[clarity];
                                sheet.Cells[filterRow + 48, 13].Value2 = Math.Round(clarityRoughCrtDolar00000001, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolar00000001 = clarityDolarDict00000001[clarity] / clarityPw07[clarity];
                                sheet.Cells[filterRow + 48, 14].Value2 = Math.Round(clarityPolishCrtDolar00000001, 2);

                                serialNumber07++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRange7 = sheet.Range[sheet.Cells[filterRow - serialNumber07 + 48, 1], sheet.Cells[filterRow + 47, 14]];
                        tableRange7.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRange7.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRange7.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 48, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 48, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeight00000001 = clarityWeightDict00000001.Values.Sum();
                        sheet.Cells[filterRow + 48, 4].Value2 = Math.Round(totalRoughtWeight00000001, 2);

                        sheet.Cells[filterRow + 48, 5].Value2 = totalCount00000001;

                        double partWTTotal00000001 = clarityWeightDict00000002.Values.Sum();
                        sheet.Cells[filterRow + 48, 6].Value2 = Math.Round(partWTTotal00000001, 2);

                        sheet.Cells[filterRow + 48, 7].Value2 = Math.Round(totalPw00000001, 3);

                        double totalSize00000001 = (totalCount00000001 / totalRoughtWeight00000001);
                        sheet.Cells[filterRow + 48, 8].Value2 = totalSize00000001.ToString("0.00");

                        double polishSize00000001 = (totalCount00000001 / totalPw00000001);
                        sheet.Cells[filterRow + 48, 9].Value2 = polishSize00000001.ToString("0.00");

                        double crPwPercentage00000001 = (totalPw00000001 / partWTTotal00000001) * 100;
                        sheet.Cells[filterRow + 48, 10].Value2 = crPwPercentage00000001.ToString("0.00") + "%";

                        double pwPercentage00000001 = (totalPw00000001 / totalRoughtWeight00000001) * 100;
                        sheet.Cells[filterRow + 48, 11].Value2 = pwPercentage00000001.ToString("0.00") + "%";

                        double dolarTotal00000001 = clarityDolarDict00000001.Values.Sum();
                        sheet.Cells[filterRow + 48, 12].Value2 = Math.Round(dolarTotal00000001, 2);

                        double valueRough00000001 = (dolarTotal00000001 / totalRoughtWeight00000001);
                        sheet.Cells[filterRow + 48, 13].Value2 = valueRough00000001.ToString("0.00");

                        double valuePolish00000001 = (dolarTotal00000001 / totalPw00000001);
                        sheet.Cells[filterRow + 48, 14].Value2 = valuePolish00000001.ToString("0.00");

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRange8 = sheet.Range[sheet.Cells[filterRow + 48, 1], sheet.Cells[filterRow + 48, 14]];
                        //headerRange8.Font.Bold = true;
                        headerRange8.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRange8.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion


                    #region 1st Sieve PEAR Method

                    if (checkBoxRunCode001.Checked)
                    {
                        sheet.Cells[filterRow + 51, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 51, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 51, 5].Value2 = "Polish Weight";

                        string updatedValueForPw1 = textBox8.Text;
                        sheet.Cells[filterRow + 52, 3].Value2 = updatedValueForPw1;

                        string updatedValueForPw01 = txtPwRange01.Text;
                        sheet.Cells[filterRow + 52, 5].Value2 = updatedValueForPw01;

                        //sheet.Cells[filterRow + 10, 5].Value2 = "0.9 - 1.249";
                        txtPwRange01.Text = (string)sheet.Cells[filterRow + 52, 5].Value2;

                        sheet.Cells[filterRow + 52, 1].Value2 = "PEAR";

                        //sheet.Cells[filterRow + 10, 3].Value2 = "+000 -2";
                        textBox8.Text = (string)sheet.Cells[filterRow + 52, 3].Value2;

                        sheet.Cells[filterRow + 53, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 53, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 53, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 53, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 53, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 53, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 53, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 53, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 53, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 53, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 53, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 53, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 53, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 53, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 53, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeForPwp01 = sheet.Range[sheet.Cells[filterRow + 53, 1], sheet.Cells[filterRow + 53, 14]];
                        //headerRange1.Font.Bold = true;
                        headerRangeForPwp01.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the pw range from the txtPwRange01 TextBox
                        string pwRangeText01 = txtPwRange01.Text;
                        string[] pwRangeParts01 = pwRangeText01.Split('-');
                        double minPw01 = 0.00;  // Default minimum pw
                        double maxPw01 = 0.05;  // Default maximum pw

                        // Parse the width range values if the input is valid
                        if (pwRangeParts01.Length == 2 && double.TryParse(pwRangeParts01[0], out double parsedMinPw01) && double.TryParse(pwRangeParts01[1], out double parsedMaxPw01))
                        {
                            minPw01 = parsedMinPw01;
                            maxPw01 = parsedMaxPw01;
                        }

                        Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>> shapesDictionary =
                            new Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>>();

                        Dictionary<string, int> clarityCountss01 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPww01 = new Dictionary<string, double>();

                        string prevStoneShName = "";
                        // Loop through rows in excel sheet again to count Dolar for each stone name
                        for (int t = 2; t <= rowCount; t++)
                        {
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStoneShName;
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            double polishWeight = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "pear" && double.TryParse(polishWeight.ToString(), out double pwValue))
                            {
                                if (pwValue >= minPw01 && pwValue <= maxPw01 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCountss01.ContainsKey(clarity))
                                        {
                                            clarityCountss01[clarity]++;
                                            clarityPww01[clarity] += polishWeight;
                                        }
                                        else
                                        {
                                            clarityCountss01.Add(clarity, 1);
                                            clarityPww01.Add(clarity, polishWeight);
                                        }

                                        if (shapesDictionary.ContainsKey(stoneName))
                                        {
                                            shapesDictionary[stoneName].Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)> shapes = new List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>();
                                            shapes.Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                            shapesDictionary.Add(stoneName, shapes);
                                        }
                                    }
                                }
                            }
                            prevStoneShName = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING FIRST SIEVE FOR SHAPE 'PEAR' CHECKING IF THE PW FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        Dictionary<string, double> clarityWeightDictForPw01 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> pearParts = partWts.Where(p => p.Item3.ToLower() == "pear").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForPw01 = pearParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item4 >= minPw01 && p.Item4 <= maxPw01)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDictForPw01.ContainsKey(clarity))
                                {
                                    clarityWeightDictForPw01[clarity] += clarityWeightForPw01;
                                }
                                else
                                {
                                    clarityWeightDictForPw01[clarity] = clarityWeightForPw01;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDictForPw02 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesDictionary.Values.SelectMany(shapes => shapes))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDictForPw02.ContainsKey(clarity))
                            {
                                clarityWeightDictForPw02[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDictForPw02[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDictForPw01 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesDictionary.Values.SelectMany(shapes => shapes))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDictForPw01.ContainsKey(clarity))
                                {
                                    clarityDolarDictForPw01[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDictForPw01[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCountForPw01 = clarityCountss01.Values.Sum();
                        double totalPwForPw01 = clarityPww01.Values.Sum();

                        int serialNumberForPw01 = 1;

                        int totalClarityFiltersForPw01 = clarityValues.Count();
                        int currentClarityFilterForPw01 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilterForPw01 / totalClarityFiltersForPw01 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING FIRST 'PEAR' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCountss01.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 54, 1].Value2 = serialNumberForPw01;
                                sheet.Cells[filterRow + 54, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 54, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPwForPw1 = clarityPww01[clarity] / totalPwForPw01 * 100.0;
                                sheet.Cells[filterRow + 54, 3].Value2 = percentPwForPw1.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 54, 5].Value2 = clarityCountss01[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 54, 7].Value2 = clarityPww01[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeightForPw01 = clarityWeightDictForPw01.ContainsKey(clarity) ? clarityWeightDictForPw01[clarity] : 0.000;
                                sheet.Cells[filterRow + 54, 4].Value2 = Math.Round(clarityWeightForPw01, 3);

                                // Get the craft weight for this clarity
                                double clarityWeightForPw02 = clarityWeightDictForPw02.ContainsKey(clarity) ? clarityWeightDictForPw02[clarity] : 0.000;
                                sheet.Cells[filterRow + 54, 6].Value2 = Math.Round(clarityWeightForPw02, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDividedForPw01 = clarityCountss01[clarity] / clarityWeightDictForPw01[clarity];
                                sheet.Cells[filterRow + 54, 8].Value2 = Math.Round(clarityCountDividedForPw01, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCountForPw01 = clarityCountss01[clarity] / clarityPww01[clarity];
                                sheet.Cells[filterRow + 54, 9].Value2 = Math.Round(clarityPolishWtCountForPw01, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCrForPw01 = clarityPww01[clarity] / clarityWeightDictForPw02[clarity] * 100;
                                sheet.Cells[filterRow + 54, 10].Value2 = percentPwCrForPw01.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCrForPw01 = clarityPww01[clarity] / clarityWeightDictForPw01[clarity] * 100;
                                sheet.Cells[filterRow + 54, 11].Value2 = percentRoCrForPw01.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 54, 12].Value2 = clarityDolarDictForPw01[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolarForPw01 = clarityDolarDictForPw01[clarity] / clarityWeightDictForPw01[clarity];
                                sheet.Cells[filterRow + 54, 13].Value2 = Math.Round(clarityRoughCrtDolarForPw01, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolarForPw01 = clarityWeightDictForPw01[clarity] / clarityPww01[clarity];
                                sheet.Cells[filterRow + 54, 14].Value2 = Math.Round(clarityPolishCrtDolarForPw01, 2);

                                serialNumberForPw01++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRangeForPwp01 = sheet.Range[sheet.Cells[filterRow - serialNumberForPw01 + 54, 1], sheet.Cells[filterRow + 54, 14]];
                        tableRangeForPwp01.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRangeForPwp01.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRangeForPwp01.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 54, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 54, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeightForPw01 = clarityWeightDictForPw01.Values.Sum();
                        sheet.Cells[filterRow + 54, 4].Value2 = Math.Round(totalRoughtWeightForPw01, 2);

                        sheet.Cells[filterRow + 54, 5].Value2 = totalCountForPw01;

                        double partWTTotalForPw01 = clarityWeightDictForPw02.Values.Sum();
                        sheet.Cells[filterRow + 54, 6].Value2 = Math.Round(partWTTotalForPw01, 2);

                        sheet.Cells[filterRow + 54, 7].Value2 = Math.Round(totalPwForPw01, 3);

                        double totalSizeForPw01 = (totalCountForPw01 / totalRoughtWeightForPw01);
                        sheet.Cells[filterRow + 54, 8].Value2 = totalSizeForPw01.ToString("0.00");

                        double polishSizeForPw01 = (totalCountForPw01 / totalPwForPw01);
                        sheet.Cells[filterRow + 54, 9].Value2 = polishSizeForPw01.ToString("0.00");

                        double crPwPercentageForPw01 = (totalPwForPw01 / partWTTotalForPw01) * 100;
                        sheet.Cells[filterRow + 54, 10].Value2 = crPwPercentageForPw01.ToString("0.00") + "%";

                        double pwPercentageForPw01 = (totalPwForPw01 / totalRoughtWeightForPw01) * 100;
                        sheet.Cells[filterRow + 54, 11].Value2 = pwPercentageForPw01.ToString("0.00") + "%";

                        double dolarTotalForPw01 = clarityDolarDictForPw01.Values.Sum();
                        sheet.Cells[filterRow + 54, 12].Value2 = Math.Round(dolarTotalForPw01, 2);

                        double valueRoughForPw01 = (dolarTotalForPw01 / totalRoughtWeightForPw01);
                        sheet.Cells[filterRow + 54, 13].Value2 = valueRoughForPw01.ToString("0.00");

                        double valuePolishForPw01 = (dolarTotalForPw01 / totalPwForPw01);
                        sheet.Cells[filterRow + 54, 14].Value2 = valuePolishForPw01.ToString("0.00");

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeeForPwp01 = sheet.Range[sheet.Cells[filterRow + 54, 1], sheet.Cells[filterRow + 54, 14]];
                        //headerRangee1.Font.Bold = true;
                        headerRangeeForPwp01.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRangeeForPwp01.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion

                    #region 2nd Sieve PEAR Method

                    if (checkBoxRunCode002.Checked)
                    {
                        sheet.Cells[filterRow + 57, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 57, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 57, 5].Value2 = "Polish Weight";

                        string updatedValueForPw2 = textBox9.Text;
                        sheet.Cells[filterRow + 58, 3].Value2 = updatedValueForPw2;

                        string updatedValueForPw02 = txtPwRange02.Text;
                        sheet.Cells[filterRow + 58, 5].Value2 = updatedValueForPw02;

                        //sheet.Cells[filterRow + 10, 5].Value2 = "0.9 - 1.249";
                        txtPwRange02.Text = (string)sheet.Cells[filterRow + 58, 5].Value2;

                        sheet.Cells[filterRow + 58, 1].Value2 = "PEAR";

                        //sheet.Cells[filterRow + 10, 3].Value2 = "+000 -2";
                        textBox9.Text = (string)sheet.Cells[filterRow + 58, 3].Value2;

                        sheet.Cells[filterRow + 59, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 59, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 59, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 59, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 59, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 59, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 59, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 59, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 59, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 59, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 59, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 59, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 59, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 59, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 59, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeForPwp02 = sheet.Range[sheet.Cells[filterRow + 59, 1], sheet.Cells[filterRow + 59, 14]];
                        //headerRange1.Font.Bold = true;
                        headerRangeForPwp02.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the width range from the txtPwRange TextBox
                        string pwRangeText02 = txtPwRange02.Text;
                        string[] pwRangeParts02 = pwRangeText02.Split('-');
                        double minPw02 = 0.05;  // Default minimum pw
                        double maxPw02 = 0.1;  // Default maximum pw

                        // Parse the pw range values if the input is valid
                        if (pwRangeParts02.Length == 2 && double.TryParse(pwRangeParts02[0], out double parsedMinPw02) && double.TryParse(pwRangeParts02[1], out double parsedMaxPw02))
                        {
                            minPw02 = parsedMinPw02;
                            maxPw02 = parsedMaxPw02;
                        }

                        Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>> shapesDictionary02 =
                            new Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>>();

                        Dictionary<string, int> clarityCountss02 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPww02 = new Dictionary<string, double>();

                        string prevStoneShName02 = "";
                        // Loop through rows in excel sheet again to count for each stone name
                        for (int t = 2; t <= rowCount; t++)
                        {
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStoneShName02;
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            double polishWeight = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "pear" && double.TryParse(polishWeight.ToString(), out double pwValue))
                            {
                                if (pwValue >= minPw02 && pwValue <= maxPw02 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCountss02.ContainsKey(clarity))
                                        {
                                            clarityCountss02[clarity]++;
                                            clarityPww02[clarity] += polishWeight;
                                        }
                                        else
                                        {
                                            clarityCountss02.Add(clarity, 1);
                                            clarityPww02.Add(clarity, polishWeight);
                                        }

                                        if (shapesDictionary02.ContainsKey(stoneName))
                                        {
                                            shapesDictionary02[stoneName].Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)> shapes = new List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>();
                                            shapes.Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                            shapesDictionary02.Add(stoneName, shapes);
                                        }
                                    }
                                }
                            }
                            prevStoneShName02 = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING SECOND SIEVE FOR SHAPE 'PEAR' CHECKING IF THE PW FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        Dictionary<string, double> clarityWeightDictForPw001 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> pearParts = partWts.Where(p => p.Item3.ToLower() == "pear").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForPw001 = pearParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item4 >= minPw02 && p.Item4 <= maxPw02)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDictForPw001.ContainsKey(clarity))
                                {
                                    clarityWeightDictForPw001[clarity] += clarityWeightForPw001;
                                }
                                else
                                {
                                    clarityWeightDictForPw001[clarity] = clarityWeightForPw001;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDictForPw002 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesDictionary02.Values.SelectMany(shapes => shapes))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDictForPw002.ContainsKey(clarity))
                            {
                                clarityWeightDictForPw002[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDictForPw002[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDictForPw001 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesDictionary02.Values.SelectMany(shapes => shapes))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDictForPw001.ContainsKey(clarity))
                                {
                                    clarityDolarDictForPw001[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDictForPw001[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCountForPw001 = clarityCountss02.Values.Sum();
                        double totalPwForPw001 = clarityPww02.Values.Sum();


                        int serialNumberForPw02 = 1;

                        int totalClarityFiltersForPw02 = clarityValues.Count();
                        int currentClarityFilterForPw02 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilterForPw02 / totalClarityFiltersForPw02 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING SECOND 'PEAR' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCountss02.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 60, 1].Value2 = serialNumberForPw02;
                                sheet.Cells[filterRow + 60, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 60, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPwForPw2 = clarityPww02[clarity] / totalPwForPw001 * 100.0;
                                sheet.Cells[filterRow + 60, 3].Value2 = percentPwForPw2.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 60, 5].Value2 = clarityCountss02[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 60, 7].Value2 = clarityPww02[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeightForPw001 = clarityWeightDictForPw001.ContainsKey(clarity) ? clarityWeightDictForPw001[clarity] : 0.000;
                                sheet.Cells[filterRow + 60, 4].Value2 = Math.Round(clarityWeightForPw001, 3);

                                // Get the craft weight for this clarity
                                double clarityWeightForPw002 = clarityWeightDictForPw002.ContainsKey(clarity) ? clarityWeightDictForPw002[clarity] : 0.000;
                                sheet.Cells[filterRow + 60, 6].Value2 = Math.Round(clarityWeightForPw002, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDividedForPw001 = clarityCountss02[clarity] / clarityWeightDictForPw001[clarity];
                                sheet.Cells[filterRow + 60, 8].Value2 = Math.Round(clarityCountDividedForPw001, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCountForPw001 = clarityCountss02[clarity] / clarityPww02[clarity];
                                sheet.Cells[filterRow + 60, 9].Value2 = Math.Round(clarityPolishWtCountForPw001, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCrForPw001 = clarityPww02[clarity] / clarityWeightDictForPw002[clarity] * 100;
                                sheet.Cells[filterRow + 60, 10].Value2 = percentPwCrForPw001.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCrForPw001 = clarityPww02[clarity] / clarityWeightDictForPw001[clarity] * 100;
                                sheet.Cells[filterRow + 60, 11].Value2 = percentRoCrForPw001.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 60, 12].Value2 = clarityDolarDictForPw001[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolarForPw001 = clarityDolarDictForPw001[clarity] / clarityWeightDictForPw001[clarity];
                                sheet.Cells[filterRow + 60, 13].Value2 = Math.Round(clarityRoughCrtDolarForPw001, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolarForPw001 = clarityWeightDictForPw001[clarity] / clarityPww02[clarity];
                                sheet.Cells[filterRow + 60, 14].Value2 = Math.Round(clarityPolishCrtDolarForPw001, 2);

                                serialNumberForPw02++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRangeForPwp02 = sheet.Range[sheet.Cells[filterRow - serialNumberForPw02 + 60, 1], sheet.Cells[filterRow + 60, 14]];
                        tableRangeForPwp02.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRangeForPwp02.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRangeForPwp02.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 60, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 60, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeightForPw001 = clarityWeightDictForPw001.Values.Sum();
                        sheet.Cells[filterRow + 60, 4].Value2 = Math.Round(totalRoughtWeightForPw001, 2);

                        sheet.Cells[filterRow + 60, 5].Value2 = totalCountForPw001;

                        double partWTTotalForPw001 = clarityWeightDictForPw002.Values.Sum();
                        sheet.Cells[filterRow + 60, 6].Value2 = Math.Round(partWTTotalForPw001, 2);

                        sheet.Cells[filterRow + 60, 7].Value2 = Math.Round(totalPwForPw001, 3);

                        double totalSizeForPw001 = (totalCountForPw001 / totalRoughtWeightForPw001);
                        sheet.Cells[filterRow + 60, 8].Value2 = totalSizeForPw001.ToString("0.00");

                        double polishSizeForPw001 = (totalCountForPw001 / totalPwForPw001);
                        sheet.Cells[filterRow + 60, 9].Value2 = polishSizeForPw001.ToString("0.00");

                        double crPwPercentageForPw001 = (totalPwForPw001 / partWTTotalForPw001) * 100;
                        sheet.Cells[filterRow + 60, 10].Value2 = crPwPercentageForPw001.ToString("0.00") + "%";

                        double pwPercentageForPw001 = (totalPwForPw001 / totalRoughtWeightForPw001) * 100;
                        sheet.Cells[filterRow + 60, 11].Value2 = pwPercentageForPw001.ToString("0.00") + "%";

                        double dolarTotalForPw001 = clarityDolarDictForPw001.Values.Sum();
                        sheet.Cells[filterRow + 60, 12].Value2 = Math.Round(dolarTotalForPw001, 2);

                        double valueRoughForPw001 = (dolarTotalForPw001 / totalRoughtWeightForPw001);
                        sheet.Cells[filterRow + 60, 13].Value2 = valueRoughForPw001.ToString("0.00");

                        double valuePolishForPw001 = (dolarTotalForPw001 / totalPwForPw001);
                        sheet.Cells[filterRow + 60, 14].Value2 = valuePolishForPw001.ToString("0.00");

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeeForPwp02 = sheet.Range[sheet.Cells[filterRow + 60, 1], sheet.Cells[filterRow + 60, 14]];
                        //headerRangee1.Font.Bold = true;
                        headerRangeeForPwp02.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRangeeForPwp02.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion

                    #region 3rd Sieve PEAR Method

                    if (checkBoxRunCode003.Checked)
                    {
                        sheet.Cells[filterRow + 63, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 63, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 63, 5].Value2 = "Polish Weight";

                        string updatedValueForPw3 = textBox10.Text;
                        sheet.Cells[filterRow + 64, 3].Value2 = updatedValueForPw3;

                        string updatedValueForPw03 = txtPwRange03.Text;
                        sheet.Cells[filterRow + 64, 5].Value2 = updatedValueForPw03;

                        //sheet.Cells[filterRow + 10, 5].Value2 = "0.9 - 1.249";
                        txtPwRange03.Text = (string)sheet.Cells[filterRow + 64, 5].Value2;

                        sheet.Cells[filterRow + 64, 1].Value2 = "PEAR";

                        //sheet.Cells[filterRow + 10, 3].Value2 = "+000 -2";
                        textBox10.Text = (string)sheet.Cells[filterRow + 64, 3].Value2;

                        sheet.Cells[filterRow + 65, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 65, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 65, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 65, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 65, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 65, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 65, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 65, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 65, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 65, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 65, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 65, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 65, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 65, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 65, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeForPwp03 = sheet.Range[sheet.Cells[filterRow + 65, 1], sheet.Cells[filterRow + 65, 14]];
                        //headerRange1.Font.Bold = true;
                        headerRangeForPwp03.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the width range from the txtPwRange TextBox
                        string pwRangeText03 = txtPwRange03.Text;
                        string[] pwRangeParts03 = pwRangeText03.Split('-');
                        double minPw03 = 0.1;  // Default minimum pw
                        double maxPw03 = 0.2;  // Default maximum pw

                        // Parse the pw range values if the input is valid
                        if (pwRangeParts03.Length == 2 && double.TryParse(pwRangeParts03[0], out double parsedMinPw03) && double.TryParse(pwRangeParts03[1], out double parsedMaxPw03))
                        {
                            minPw03 = parsedMinPw03;
                            maxPw03 = parsedMaxPw03;
                        }

                        Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>> shapesDictionary03 =
                            new Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>>();

                        Dictionary<string, int> clarityCountss03 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPww03 = new Dictionary<string, double>();

                        string prevStoneShName03 = "";
                        // Loop through rows in excel sheet again to count for each stone name
                        for (int t = 2; t <= rowCount; t++)
                        {
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStoneShName03;
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            double polishWeight = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "pear" && double.TryParse(polishWeight.ToString(), out double pwValue))
                            {
                                if (pwValue >= minPw03 && pwValue <= maxPw03 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCountss03.ContainsKey(clarity))
                                        {
                                            clarityCountss03[clarity]++;
                                            clarityPww03[clarity] += polishWeight;
                                        }
                                        else
                                        {
                                            clarityCountss03.Add(clarity, 1);
                                            clarityPww03.Add(clarity, polishWeight);
                                        }

                                        if (shapesDictionary03.ContainsKey(stoneName))
                                        {
                                            shapesDictionary03[stoneName].Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)> shapes = new List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>();
                                            shapes.Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                            shapesDictionary03.Add(stoneName, shapes);
                                        }
                                    }
                                }
                            }
                            prevStoneShName03 = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING THIRD SIEVE FOR SHAPE 'PEAR' CHECKING IF THE PW FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        Dictionary<string, double> clarityWeightDictForPw0001 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> pearParts = partWts.Where(p => p.Item3.ToLower() == "pear").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForPw0001 = pearParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item4 >= minPw03 && p.Item4 <= maxPw03)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDictForPw0001.ContainsKey(clarity))
                                {
                                    clarityWeightDictForPw0001[clarity] += clarityWeightForPw0001;
                                }
                                else
                                {
                                    clarityWeightDictForPw0001[clarity] = clarityWeightForPw0001;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDictForPw0002 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesDictionary03.Values.SelectMany(shapes => shapes))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDictForPw0002.ContainsKey(clarity))
                            {
                                clarityWeightDictForPw0002[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDictForPw0002[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDictForPw0001 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesDictionary03.Values.SelectMany(shapes => shapes))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDictForPw0001.ContainsKey(clarity))
                                {
                                    clarityDolarDictForPw0001[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDictForPw0001[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCountForPw0001 = clarityCountss03.Values.Sum();
                        double totalPwForPw0001 = clarityPww03.Values.Sum();

                        int serialNumberForPw03 = 1;

                        int totalClarityFiltersForPw03 = clarityValues.Count();
                        int currentClarityFilterForPw03 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilterForPw03 / totalClarityFiltersForPw03 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING THIRD 'PEAR' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCountss03.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 66, 1].Value2 = serialNumberForPw03;
                                sheet.Cells[filterRow + 66, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 66, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPwForPw3 = clarityPww03[clarity] / totalPwForPw0001 * 100.0;
                                sheet.Cells[filterRow + 66, 3].Value2 = percentPwForPw3.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 66, 5].Value2 = clarityCountss03[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 66, 7].Value2 = clarityPww03[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeightForPw0001 = clarityWeightDictForPw0001.ContainsKey(clarity) ? clarityWeightDictForPw0001[clarity] : 0.000;
                                sheet.Cells[filterRow + 66, 4].Value2 = Math.Round(clarityWeightForPw0001, 3);

                                // Get the craft weight for this clarity
                                double clarityWeightForPw0002 = clarityWeightDictForPw0002.ContainsKey(clarity) ? clarityWeightDictForPw0002[clarity] : 0.000;
                                sheet.Cells[filterRow + 66, 6].Value2 = Math.Round(clarityWeightForPw0002, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDividedForPw0001 = clarityCountss03[clarity] / clarityWeightDictForPw0001[clarity];
                                sheet.Cells[filterRow + 66, 8].Value2 = Math.Round(clarityCountDividedForPw0001, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCountForPw0001 = clarityCountss03[clarity] / clarityPww03[clarity];
                                sheet.Cells[filterRow + 66, 9].Value2 = Math.Round(clarityPolishWtCountForPw0001, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCrForPw0001 = clarityPww03[clarity] / clarityWeightDictForPw0002[clarity] * 100;
                                sheet.Cells[filterRow + 66, 10].Value2 = percentPwCrForPw0001.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCrForPw0001 = clarityPww03[clarity] / clarityWeightDictForPw0001[clarity] * 100;
                                sheet.Cells[filterRow + 66, 11].Value2 = percentRoCrForPw0001.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 66, 12].Value2 = clarityDolarDictForPw0001[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolarForPw0001 = clarityDolarDictForPw0001[clarity] / clarityWeightDictForPw0001[clarity];
                                sheet.Cells[filterRow + 66, 13].Value2 = Math.Round(clarityRoughCrtDolarForPw0001, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolarForPw0001 = clarityWeightDictForPw0001[clarity] / clarityPww03[clarity];
                                sheet.Cells[filterRow + 66, 14].Value2 = Math.Round(clarityPolishCrtDolarForPw0001, 2);

                                serialNumberForPw03++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRangeForPwp03 = sheet.Range[sheet.Cells[filterRow - serialNumberForPw03 + 66, 1], sheet.Cells[filterRow + 66, 14]];
                        tableRangeForPwp03.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRangeForPwp03.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRangeForPwp03.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 66, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 66, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeightForPw0001 = clarityWeightDictForPw0001.Values.Sum();
                        sheet.Cells[filterRow + 66, 4].Value2 = Math.Round(totalRoughtWeightForPw0001, 2);

                        sheet.Cells[filterRow + 66, 5].Value2 = totalCountForPw0001;

                        double partWTTotalForPw0001 = clarityWeightDictForPw0002.Values.Sum();
                        sheet.Cells[filterRow + 66, 6].Value2 = Math.Round(partWTTotalForPw0001, 2);

                        sheet.Cells[filterRow + 66, 7].Value2 = Math.Round(totalPwForPw0001, 3);

                        double totalSizeForPw0001 = (totalCountForPw0001 / totalRoughtWeightForPw0001);
                        sheet.Cells[filterRow + 66, 8].Value2 = totalSizeForPw0001.ToString("0.00");

                        double polishSizeForPw0001 = (totalCountForPw0001 / totalPwForPw0001);
                        sheet.Cells[filterRow + 66, 9].Value2 = polishSizeForPw0001.ToString("0.00");

                        double crPwPercentageForPw0001 = (totalPwForPw0001 / partWTTotalForPw0001) * 100;
                        sheet.Cells[filterRow + 66, 10].Value2 = crPwPercentageForPw0001.ToString("0.00") + "%";

                        double pwPercentageForPw0001 = (totalPwForPw0001 / totalRoughtWeightForPw0001) * 100;
                        sheet.Cells[filterRow + 66, 11].Value2 = pwPercentageForPw0001.ToString("0.00") + "%";

                        double dolarTotalForPw0001 = clarityDolarDictForPw0001.Values.Sum();
                        sheet.Cells[filterRow + 66, 12].Value2 = Math.Round(dolarTotalForPw0001, 2);

                        double valueRoughForPw0001 = (dolarTotalForPw0001 / totalRoughtWeightForPw0001);
                        sheet.Cells[filterRow + 66, 13].Value2 = valueRoughForPw0001.ToString("0.00");

                        double valuePolishForPw0001 = (dolarTotalForPw0001 / totalPwForPw0001);
                        sheet.Cells[filterRow + 66, 14].Value2 = valuePolishForPw0001.ToString("0.00");


                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeeForPwp03 = sheet.Range[sheet.Cells[filterRow + 66, 1], sheet.Cells[filterRow + 66, 14]];
                        //headerRangee1.Font.Bold = true;
                        headerRangeeForPwp03.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRangeeForPwp03.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion

                    #region 4th Sieve PEAR Method

                    if (checkBoxRunCode004.Checked)
                    {
                        sheet.Cells[filterRow + 69, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 69, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 69, 5].Value2 = "Polish Weight";

                        string updatedValueForPw4 = textBox11.Text;
                        sheet.Cells[filterRow + 70, 3].Value2 = updatedValueForPw4;

                        string updatedValueForPw04 = txtPwRange04.Text;
                        sheet.Cells[filterRow + 70, 5].Value2 = updatedValueForPw04;

                        //sheet.Cells[filterRow + 10, 5].Value2 = "0.9 - 1.249";
                        txtPwRange04.Text = (string)sheet.Cells[filterRow + 70, 5].Value2;

                        sheet.Cells[filterRow + 70, 1].Value2 = "PEAR";

                        //sheet.Cells[filterRow + 10, 3].Value2 = "+000 -2";
                        textBox11.Text = (string)sheet.Cells[filterRow + 70, 3].Value2;

                        sheet.Cells[filterRow + 71, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 71, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 71, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 71, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 71, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 71, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 71, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 71, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 71, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 71, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 71, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 71, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 71, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 71, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 71, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeForPwp04 = sheet.Range[sheet.Cells[filterRow + 71, 1], sheet.Cells[filterRow + 71, 14]];
                        //headerRange1.Font.Bold = true;
                        headerRangeForPwp04.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the Pw range from the txtPwRange04 TextBox
                        string pwRangeText04 = txtPwRange04.Text;
                        string[] pwRangeParts04 = pwRangeText04.Split('-');
                        double minPw04 = 0.3;  // Default minimum pw
                        double maxPw04 = 0.4;  // Default maximum pw

                        // Parse the pw range values if the input is valid
                        if (pwRangeParts04.Length == 2 && double.TryParse(pwRangeParts04[0], out double parsedMinPw04) && double.TryParse(pwRangeParts04[1], out double parsedMaxPw04))
                        {
                            minPw04 = parsedMinPw04;
                            maxPw04 = parsedMaxPw04;
                        }

                        Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>> shapesDictionary04 =
                            new Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>>();

                        Dictionary<string, int> clarityCountss04 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPww04 = new Dictionary<string, double>();

                        string prevStoneShName04 = "";
                        // Loop through rows in excel sheet again to count for each stone name
                        for (int t = 2; t <= rowCount; t++)
                        {
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStoneShName04;
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            double polishWeight = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "pear" && double.TryParse(polishWeight.ToString(), out double pwValue))
                            {
                                if (pwValue >= minPw04 && pwValue <= maxPw04 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCountss04.ContainsKey(clarity))
                                        {
                                            clarityCountss04[clarity]++;
                                            clarityPww04[clarity] += polishWeight;
                                        }
                                        else
                                        {
                                            clarityCountss04.Add(clarity, 1);
                                            clarityPww04.Add(clarity, polishWeight);
                                        }

                                        if (shapesDictionary04.ContainsKey(stoneName))
                                        {
                                            shapesDictionary04[stoneName].Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)> shapes = new List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>();
                                            shapes.Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                            shapesDictionary04.Add(stoneName, shapes);
                                        }
                                    }
                                }
                            }
                            prevStoneShName04 = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING FOURTH SIEVE FOR SHAPE 'PEAR' CHECKING IF THE PW FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }



                        Dictionary<string, double> clarityWeightDictForPw00001 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> pearParts = partWts.Where(p => p.Item3.ToLower() == "pear").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForPw00001 = pearParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item4 >= minPw04 && p.Item4 <= maxPw04)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDictForPw00001.ContainsKey(clarity))
                                {
                                    clarityWeightDictForPw00001[clarity] += clarityWeightForPw00001;
                                }
                                else
                                {
                                    clarityWeightDictForPw00001[clarity] = clarityWeightForPw00001;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDictForPw00002 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesDictionary04.Values.SelectMany(shapes => shapes))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDictForPw00002.ContainsKey(clarity))
                            {
                                clarityWeightDictForPw00002[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDictForPw00002[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDictForPw00001 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesDictionary04.Values.SelectMany(shapes => shapes))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDictForPw00001.ContainsKey(clarity))
                                {
                                    clarityDolarDictForPw00001[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDictForPw00001[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCountForPw00001 = clarityCountss04.Values.Sum();
                        double totalPwForPw00001 = clarityPww04.Values.Sum();

                        int serialNumberForPw04 = 1;

                        int totalClarityFiltersForPw04 = clarityValues.Count();
                        int currentClarityFilterForPw04 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilterForPw04 / totalClarityFiltersForPw04 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING FOURTH 'PEAR' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCountss04.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 72, 1].Value2 = serialNumberForPw04;
                                sheet.Cells[filterRow + 72, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 72, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPwForPw4 = clarityPww04[clarity] / totalPwForPw00001 * 100.0;
                                sheet.Cells[filterRow + 72, 3].Value2 = percentPwForPw4.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 72, 5].Value2 = clarityCountss04[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 72, 7].Value2 = clarityPww04[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeightForPw00001 = clarityWeightDictForPw00001.ContainsKey(clarity) ? clarityWeightDictForPw00001[clarity] : 0.000;
                                sheet.Cells[filterRow + 72, 4].Value2 = Math.Round(clarityWeightForPw00001, 3);

                                // Get the craft weight for this clarity
                                double clarityWeightForPw00002 = clarityWeightDictForPw00002.ContainsKey(clarity) ? clarityWeightDictForPw00002[clarity] : 0.000;
                                sheet.Cells[filterRow + 72, 6].Value2 = Math.Round(clarityWeightForPw00002, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDividedForPw00001 = clarityCountss04[clarity] / clarityWeightDictForPw00001[clarity];
                                sheet.Cells[filterRow + 72, 8].Value2 = Math.Round(clarityCountDividedForPw00001, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCountForPw00001 = clarityCountss04[clarity] / clarityPww04[clarity];
                                sheet.Cells[filterRow + 72, 9].Value2 = Math.Round(clarityPolishWtCountForPw00001, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCrForPw00001 = clarityPww04[clarity] / clarityWeightDictForPw00002[clarity] * 100;
                                sheet.Cells[filterRow + 72, 10].Value2 = percentPwCrForPw00001.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCrForPw00001 = clarityPww04[clarity] / clarityWeightDictForPw00001[clarity] * 100;
                                sheet.Cells[filterRow + 72, 11].Value2 = percentRoCrForPw00001.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 72, 12].Value2 = clarityDolarDictForPw00001[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolarForPw00001 = clarityDolarDictForPw00001[clarity] / clarityWeightDictForPw00001[clarity];
                                sheet.Cells[filterRow + 72, 13].Value2 = Math.Round(clarityRoughCrtDolarForPw00001, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolarForPw00001 = clarityWeightDictForPw00001[clarity] / clarityPww04[clarity];
                                sheet.Cells[filterRow + 72, 14].Value2 = Math.Round(clarityPolishCrtDolarForPw00001, 2);

                                serialNumberForPw04++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRangeForPwp04 = sheet.Range[sheet.Cells[filterRow - serialNumberForPw04 + 72, 1], sheet.Cells[filterRow + 72, 14]];
                        tableRangeForPwp04.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRangeForPwp04.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRangeForPwp04.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 72, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 72, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeightForPw00001 = clarityWeightDictForPw00001.Values.Sum();
                        sheet.Cells[filterRow + 72, 4].Value2 = Math.Round(totalRoughtWeightForPw00001, 2);

                        sheet.Cells[filterRow + 72, 5].Value2 = totalCountForPw00001;

                        double partWTTotalForPw00001 = clarityWeightDictForPw00002.Values.Sum();
                        sheet.Cells[filterRow + 72, 6].Value2 = Math.Round(partWTTotalForPw00001, 2);

                        sheet.Cells[filterRow + 72, 7].Value2 = Math.Round(totalPwForPw00001, 3);

                        double totalSizeForPw00001 = (totalCountForPw00001 / totalRoughtWeightForPw00001);
                        sheet.Cells[filterRow + 72, 8].Value2 = totalSizeForPw00001.ToString("0.00");

                        double polishSizeForPw00001 = (totalCountForPw00001 / totalPwForPw00001);
                        sheet.Cells[filterRow + 72, 9].Value2 = polishSizeForPw00001.ToString("0.00");

                        double crPwPercentageForPw00001 = (totalPwForPw00001 / partWTTotalForPw00001) * 100;
                        sheet.Cells[filterRow + 72, 10].Value2 = crPwPercentageForPw00001.ToString("0.00") + "%";

                        double pwPercentageForPw00001 = (totalPwForPw00001 / totalRoughtWeightForPw00001) * 100;
                        sheet.Cells[filterRow + 72, 11].Value2 = pwPercentageForPw00001.ToString("0.00") + "%";

                        double dolarTotalForPw00001 = clarityDolarDictForPw00001.Values.Sum();
                        sheet.Cells[filterRow + 72, 12].Value2 = Math.Round(dolarTotalForPw00001, 2);

                        double valueRoughForPw00001 = (dolarTotalForPw00001 / totalRoughtWeightForPw00001);
                        sheet.Cells[filterRow + 72, 13].Value2 = valueRoughForPw00001.ToString("0.00");

                        double valuePolishForPw00001 = (dolarTotalForPw00001 / totalPwForPw00001);
                        sheet.Cells[filterRow + 72, 14].Value2 = valuePolishForPw00001.ToString("0.00");


                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeeForPwp04 = sheet.Range[sheet.Cells[filterRow + 72, 1], sheet.Cells[filterRow + 72, 14]];
                        //headerRangee1.Font.Bold = true;
                        headerRangeeForPwp04.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRangeeForPwp04.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion

                    #region 5th Sieve PEAR Method

                    if (checkBoxRunCode005.Checked)
                    {
                        sheet.Cells[filterRow + 75, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 75, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 75, 5].Value2 = "Polish Weight";

                        string updatedValueForPw5 = textBox12.Text;
                        sheet.Cells[filterRow + 76, 3].Value2 = updatedValueForPw5;

                        string updatedValueForPw05 = txtPwRange05.Text;
                        sheet.Cells[filterRow + 76, 5].Value2 = updatedValueForPw05;

                        //sheet.Cells[filterRow + 10, 5].Value2 = "0.9 - 1.249";
                        txtPwRange05.Text = (string)sheet.Cells[filterRow + 76, 5].Value2;

                        sheet.Cells[filterRow + 76, 1].Value2 = "PEAR";

                        //sheet.Cells[filterRow + 10, 3].Value2 = "+000 -2";
                        textBox12.Text = (string)sheet.Cells[filterRow + 76, 3].Value2;

                        sheet.Cells[filterRow + 77, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 77, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 77, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 77, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 77, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 77, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 77, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 77, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 77, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 77, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 77, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 77, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 77, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 77, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 77, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeForPwp05 = sheet.Range[sheet.Cells[filterRow + 77, 1], sheet.Cells[filterRow + 77, 14]];
                        //headerRange1.Font.Bold = true;
                        headerRangeForPwp05.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the Pw range from the txtPwRange04 TextBox
                        string pwRangeText05 = txtPwRange05.Text;
                        string[] pwRangeParts05 = pwRangeText05.Split('-');
                        double minPw05 = 0.5;  // Default minimum pw
                        double maxPw05 = 0.6;  // Default maximum pw

                        // Parse the pw range values if the input is valid
                        if (pwRangeParts05.Length == 2 && double.TryParse(pwRangeParts05[0], out double parsedMinPw05) && double.TryParse(pwRangeParts05[1], out double parsedMaxPw05))
                        {
                            minPw05 = parsedMinPw05;
                            maxPw05 = parsedMaxPw05;
                        }

                        Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>> shapesDictionary05 =
                            new Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>>();

                        Dictionary<string, int> clarityCountss05 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPww05 = new Dictionary<string, double>();

                        string prevStoneShName05 = "";
                        // Loop through rows in excel sheet again to count for each stone name
                        for (int t = 2; t <= rowCount; t++)
                        {
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStoneShName05;
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            double polishWeight = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "pear" && double.TryParse(polishWeight.ToString(), out double pwValue))
                            {
                                if (pwValue >= minPw05 && pwValue <= maxPw05 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCountss05.ContainsKey(clarity))
                                        {
                                            clarityCountss05[clarity]++;
                                            clarityPww05[clarity] += polishWeight;
                                        }
                                        else
                                        {
                                            clarityCountss05.Add(clarity, 1);
                                            clarityPww05.Add(clarity, polishWeight);
                                        }

                                        if (shapesDictionary05.ContainsKey(stoneName))
                                        {
                                            shapesDictionary05[stoneName].Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)> shapes = new List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>();
                                            shapes.Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                            shapesDictionary05.Add(stoneName, shapes);
                                        }
                                    }
                                }
                            }
                            prevStoneShName05 = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING FIFTH SIEVE FOR SHAPE 'PEAR' CHECKING IF THE PW FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                        }

                        Dictionary<string, double> clarityWeightDictForPw000001 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> pearParts = partWts.Where(p => p.Item3.ToLower() == "pear").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForPw000001 = pearParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item4 >= minPw05 && p.Item4 <= maxPw05)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDictForPw000001.ContainsKey(clarity))
                                {
                                    clarityWeightDictForPw000001[clarity] += clarityWeightForPw000001;
                                }
                                else
                                {
                                    clarityWeightDictForPw000001[clarity] = clarityWeightForPw000001;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDictForPw000002 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesDictionary05.Values.SelectMany(shapes => shapes))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDictForPw000002.ContainsKey(clarity))
                            {
                                clarityWeightDictForPw000002[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDictForPw000002[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDictForPw000001 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesDictionary05.Values.SelectMany(shapes => shapes))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDictForPw000001.ContainsKey(clarity))
                                {
                                    clarityDolarDictForPw000001[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDictForPw000001[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCountForPw000001 = clarityCountss05.Values.Sum();
                        double totalPwForPw000001 = clarityPww05.Values.Sum();

                        int serialNumberForPw05 = 1;

                        int totalClarityFiltersForPw05 = clarityValues.Count();
                        int currentClarityFilterForPw05 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilterForPw05 / totalClarityFiltersForPw05 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING FIFTH 'PEAR' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCountss05.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 78, 1].Value2 = serialNumberForPw05;
                                sheet.Cells[filterRow + 78, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 78, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPwForPw5 = clarityPww05[clarity] / totalPwForPw000001 * 100.0;
                                sheet.Cells[filterRow + 78, 3].Value2 = percentPwForPw5.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 78, 5].Value2 = clarityCountss05[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 78, 7].Value2 = clarityPww05[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeightForPw000001 = clarityWeightDictForPw000001.ContainsKey(clarity) ? clarityWeightDictForPw000001[clarity] : 0.000;
                                sheet.Cells[filterRow + 78, 4].Value2 = Math.Round(clarityWeightForPw000001, 3);

                                // Get the craft weight for this clarity
                                double clarityWeightForPw000002 = clarityWeightDictForPw000002.ContainsKey(clarity) ? clarityWeightDictForPw000002[clarity] : 0.000;
                                sheet.Cells[filterRow + 78, 6].Value2 = Math.Round(clarityWeightForPw000002, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDividedForPw000001 = clarityCountss05[clarity] / clarityWeightDictForPw000001[clarity];
                                sheet.Cells[filterRow + 78, 8].Value2 = Math.Round(clarityCountDividedForPw000001, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCountForPw000001 = clarityCountss05[clarity] / clarityPww05[clarity];
                                sheet.Cells[filterRow + 78, 9].Value2 = Math.Round(clarityPolishWtCountForPw000001, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCrForPw000001 = clarityPww05[clarity] / clarityWeightDictForPw000002[clarity] * 100;
                                sheet.Cells[filterRow + 78, 10].Value2 = percentPwCrForPw000001.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCrForPw000001 = clarityPww05[clarity] / clarityWeightDictForPw000001[clarity] * 100;
                                sheet.Cells[filterRow + 78, 11].Value2 = percentRoCrForPw000001.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 78, 12].Value2 = clarityDolarDictForPw000001[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolarForPw000001 = clarityDolarDictForPw000001[clarity] / clarityWeightDictForPw000001[clarity];
                                sheet.Cells[filterRow + 78, 13].Value2 = Math.Round(clarityRoughCrtDolarForPw000001, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolarForPw000001 = clarityWeightDictForPw000001[clarity] / clarityPww05[clarity];
                                sheet.Cells[filterRow + 78, 14].Value2 = Math.Round(clarityPolishCrtDolarForPw000001, 2);

                                serialNumberForPw05++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRangeForPwp05 = sheet.Range[sheet.Cells[filterRow - serialNumberForPw05 + 78, 1], sheet.Cells[filterRow + 78, 14]];
                        tableRangeForPwp05.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRangeForPwp05.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRangeForPwp05.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 78, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 78, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeightForPw000001 = clarityWeightDictForPw000001.Values.Sum();
                        sheet.Cells[filterRow + 78, 4].Value2 = Math.Round(totalRoughtWeightForPw000001, 2);

                        sheet.Cells[filterRow + 78, 5].Value2 = totalCountForPw000001;

                        double partWTTotalForPw000001 = clarityWeightDictForPw000002.Values.Sum();
                        sheet.Cells[filterRow + 78, 6].Value2 = Math.Round(partWTTotalForPw000001, 2);

                        sheet.Cells[filterRow + 78, 7].Value2 = Math.Round(totalPwForPw000001, 3);

                        double totalSizeForPw000001 = (totalCountForPw000001 / totalRoughtWeightForPw000001);
                        sheet.Cells[filterRow + 78, 8].Value2 = totalSizeForPw000001.ToString("0.00");

                        double polishSizeForPw000001 = (totalCountForPw000001 / totalPwForPw000001);
                        sheet.Cells[filterRow + 78, 9].Value2 = polishSizeForPw000001.ToString("0.00");

                        double crPwPercentageForPw000001 = (totalPwForPw000001 / partWTTotalForPw000001) * 100;
                        sheet.Cells[filterRow + 78, 10].Value2 = crPwPercentageForPw000001.ToString("0.00") + "%";

                        double pwPercentageForPw000001 = (totalPwForPw000001 / totalRoughtWeightForPw000001) * 100;
                        sheet.Cells[filterRow + 78, 11].Value2 = pwPercentageForPw000001.ToString("0.00") + "%";

                        double dolarTotalForPw000001 = clarityDolarDictForPw000001.Values.Sum();
                        sheet.Cells[filterRow + 78, 12].Value2 = Math.Round(dolarTotalForPw000001, 2);

                        double valueRoughForPw000001 = (dolarTotalForPw000001 / totalRoughtWeightForPw000001);
                        sheet.Cells[filterRow + 78, 13].Value2 = valueRoughForPw000001.ToString("0.00");

                        double valuePolishForPw000001 = (dolarTotalForPw000001 / totalPwForPw000001);
                        sheet.Cells[filterRow + 78, 14].Value2 = valuePolishForPw000001.ToString("0.00");


                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeeForPwp05 = sheet.Range[sheet.Cells[filterRow + 78, 1], sheet.Cells[filterRow + 78, 14]];
                        //headerRangee1.Font.Bold = true;
                        headerRangeeForPwp05.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRangeeForPwp05.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion


                    #region 1st Sieve MARQUISE Method

                    if (checkBoxRunCode006.Checked)
                    {
                        sheet.Cells[filterRow + 81, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 81, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 81, 5].Value2 = "Polish Weight";

                        string updatedValueForPwm1 = textBox13.Text;
                        sheet.Cells[filterRow + 82, 3].Value2 = updatedValueForPwm1;

                        string updatedValueForPwm01 = txtPwRange06.Text;
                        sheet.Cells[filterRow + 82, 5].Value2 = updatedValueForPwm01;

                        //sheet.Cells[filterRow + 10, 5].Value2 = "0.9 - 1.249";
                        txtPwRange06.Text = (string)sheet.Cells[filterRow + 82, 5].Value2;

                        sheet.Cells[filterRow + 82, 1].Value2 = "MARQUISE";

                        //sheet.Cells[filterRow + 10, 3].Value2 = "+000 -2";
                        textBox13.Text = (string)sheet.Cells[filterRow + 82, 3].Value2;

                        sheet.Cells[filterRow + 83, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 83, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 83, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 83, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 83, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 83, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 83, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 83, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 83, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 83, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 83, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 83, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 83, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 83, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 83, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeForPwm01 = sheet.Range[sheet.Cells[filterRow + 83, 1], sheet.Cells[filterRow + 83, 14]];
                        //headerRange1.Font.Bold = true;
                        headerRangeForPwm01.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the pw range from the txtPwRange01 TextBox
                        string pwRangeText06 = txtPwRange06.Text;
                        string[] pwRangeParts06 = pwRangeText06.Split('-');
                        double minPwm01 = 0.00;  // Default minimum pw
                        double maxPwm01 = 0.05;  // Default maximum pw

                        // Parse the width range values if the input is valid
                        if (pwRangeParts06.Length == 2 && double.TryParse(pwRangeParts06[0], out double parsedMinPw06) && double.TryParse(pwRangeParts06[1], out double parsedMaxPw06))
                        {
                            minPwm01 = parsedMinPw06;
                            maxPwm01 = parsedMaxPw06;
                        }

                        Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>> shapesMDictionary =
                            new Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>>();

                        Dictionary<string, int> clarityCountssm01 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPwwm01 = new Dictionary<string, double>();

                        string prevStoneShName = "";
                        // Loop through rows in excel sheet again to count Dolar for each stone name
                        for (int t = 2; t <= rowCount; t++)
                        {
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStoneShName;
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            double polishWeight = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "marquise" && double.TryParse(polishWeight.ToString(), out double pwValue))
                            {
                                if (pwValue >= minPwm01 && pwValue <= maxPwm01 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCountssm01.ContainsKey(clarity))
                                        {
                                            clarityCountssm01[clarity]++;
                                            clarityPwwm01[clarity] += polishWeight;
                                        }
                                        else
                                        {
                                            clarityCountssm01.Add(clarity, 1);
                                            clarityPwwm01.Add(clarity, polishWeight);
                                        }

                                        if (shapesMDictionary.ContainsKey(stoneName))
                                        {
                                            shapesMDictionary[stoneName].Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)> shapes = new List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>();
                                            shapes.Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                            shapesMDictionary.Add(stoneName, shapes);
                                        }
                                    }
                                }
                            }
                            prevStoneShName = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING FIRST SIEVE FOR SHAPE 'MARQUISE' CHECKING IF THE PW FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                        }

                        Dictionary<string, double> clarityWeightDictForPwm01 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> pearParts = partWts.Where(p => p.Item3.ToLower() == "marquise").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForPwm01 = pearParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item4 >= minPwm01 && p.Item4 <= maxPwm01)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDictForPwm01.ContainsKey(clarity))
                                {
                                    clarityWeightDictForPwm01[clarity] += clarityWeightForPwm01;
                                }
                                else
                                {
                                    clarityWeightDictForPwm01[clarity] = clarityWeightForPwm01;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDictForPwm02 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesMDictionary.Values.SelectMany(shapes => shapes))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDictForPwm02.ContainsKey(clarity))
                            {
                                clarityWeightDictForPwm02[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDictForPwm02[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDictForPwm01 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesMDictionary.Values.SelectMany(shapes => shapes))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDictForPwm01.ContainsKey(clarity))
                                {
                                    clarityDolarDictForPwm01[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDictForPwm01[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCountForPwm01 = clarityCountssm01.Values.Sum();
                        double totalPwForPwm01 = clarityPwwm01.Values.Sum();

                        int serialNumberForPwm01 = 1;

                        int totalClarityFiltersForPwm01 = clarityValues.Count();
                        int currentClarityFilterForPwm01 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilterForPwm01 / totalClarityFiltersForPwm01 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING FIRST 'MARQUISE' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCountssm01.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 84, 1].Value2 = serialNumberForPwm01;
                                sheet.Cells[filterRow + 84, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 84, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPwForPwm1 = clarityPwwm01[clarity] / totalPwForPwm01 * 100.0;
                                sheet.Cells[filterRow + 84, 3].Value2 = percentPwForPwm1.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 84, 5].Value2 = clarityCountssm01[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 84, 7].Value2 = clarityPwwm01[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeightForPwm01 = clarityWeightDictForPwm01.ContainsKey(clarity) ? clarityWeightDictForPwm01[clarity] : 0.000;
                                sheet.Cells[filterRow + 84, 4].Value2 = Math.Round(clarityWeightForPwm01, 3);

                                // Get the craft weight for this clarity
                                double clarityWeightForPwm02 = clarityWeightDictForPwm02.ContainsKey(clarity) ? clarityWeightDictForPwm02[clarity] : 0.000;
                                sheet.Cells[filterRow + 84, 6].Value2 = Math.Round(clarityWeightForPwm02, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDividedForPwm01 = clarityCountssm01[clarity] / clarityWeightDictForPwm01[clarity];
                                sheet.Cells[filterRow + 84, 8].Value2 = Math.Round(clarityCountDividedForPwm01, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCountForPwm01 = clarityCountssm01[clarity] / clarityPwwm01[clarity];
                                sheet.Cells[filterRow + 84, 9].Value2 = Math.Round(clarityPolishWtCountForPwm01, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCrForPwm01 = clarityPwwm01[clarity] / clarityWeightDictForPwm02[clarity] * 100;
                                sheet.Cells[filterRow + 84, 10].Value2 = percentPwCrForPwm01.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCrForPwm01 = clarityPwwm01[clarity] / clarityWeightDictForPwm01[clarity] * 100;
                                sheet.Cells[filterRow + 84, 11].Value2 = percentRoCrForPwm01.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 84, 12].Value2 = clarityDolarDictForPwm01[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolarForPwm01 = clarityDolarDictForPwm01[clarity] / clarityWeightDictForPwm01[clarity];
                                sheet.Cells[filterRow + 84, 13].Value2 = Math.Round(clarityRoughCrtDolarForPwm01, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolarForPwm01 = clarityWeightDictForPwm01[clarity] / clarityPwwm01[clarity];
                                sheet.Cells[filterRow + 84, 14].Value2 = Math.Round(clarityPolishCrtDolarForPwm01, 2);

                                serialNumberForPwm01++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRangeForPwm01 = sheet.Range[sheet.Cells[filterRow - serialNumberForPwm01 + 84, 1], sheet.Cells[filterRow + 84, 14]];
                        tableRangeForPwm01.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRangeForPwm01.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRangeForPwm01.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 84, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 84, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeightForPwm01 = clarityWeightDictForPwm01.Values.Sum();
                        sheet.Cells[filterRow + 84, 4].Value2 = Math.Round(totalRoughtWeightForPwm01, 2);

                        sheet.Cells[filterRow + 84, 5].Value2 = totalCountForPwm01;

                        double partWTTotalForPwm01 = clarityWeightDictForPwm02.Values.Sum();
                        sheet.Cells[filterRow + 84, 6].Value2 = Math.Round(partWTTotalForPwm01, 2);

                        sheet.Cells[filterRow + 84, 7].Value2 = Math.Round(totalPwForPwm01, 3);

                        double totalSizeForPwm01 = (totalCountForPwm01 / totalRoughtWeightForPwm01);
                        sheet.Cells[filterRow + 84, 8].Value2 = totalSizeForPwm01.ToString("0.00");

                        double polishSizeForPwm01 = (totalCountForPwm01 / totalPwForPwm01);
                        sheet.Cells[filterRow + 84, 9].Value2 = polishSizeForPwm01.ToString("0.00");

                        double crPwPercentageForPwm01 = (totalPwForPwm01 / partWTTotalForPwm01) * 100;
                        sheet.Cells[filterRow + 84, 10].Value2 = crPwPercentageForPwm01.ToString("0.00") + "%";

                        double pwPercentageForPwm01 = (totalPwForPwm01 / totalRoughtWeightForPwm01) * 100;
                        sheet.Cells[filterRow + 84, 11].Value2 = pwPercentageForPwm01.ToString("0.00") + "%";

                        double dolarTotalForPwm01 = clarityDolarDictForPwm01.Values.Sum();
                        sheet.Cells[filterRow + 84, 12].Value2 = Math.Round(dolarTotalForPwm01, 2);

                        double valueRoughForPwm01 = (dolarTotalForPwm01 / totalRoughtWeightForPwm01);
                        sheet.Cells[filterRow + 84, 13].Value2 = valueRoughForPwm01.ToString("0.00");

                        double valuePolishForPwm01 = (dolarTotalForPwm01 / totalPwForPwm01);
                        sheet.Cells[filterRow + 84, 14].Value2 = valuePolishForPwm01.ToString("0.00");

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeeForPwm01 = sheet.Range[sheet.Cells[filterRow + 84, 1], sheet.Cells[filterRow + 84, 14]];
                        //headerRangee1.Font.Bold = true;
                        headerRangeeForPwm01.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRangeeForPwm01.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion

                    #region 2nd Sieve MARQUISE Method

                    if (checkBoxRunCode007.Checked)
                    {
                        sheet.Cells[filterRow + 87, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 87, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 87, 5].Value2 = "Polish Weight";

                        string updatedValueForPwm2 = textBox14.Text;
                        sheet.Cells[filterRow + 88, 3].Value2 = updatedValueForPwm2;

                        string updatedValueForPwm02 = txtPwRange07.Text;
                        sheet.Cells[filterRow + 88, 5].Value2 = updatedValueForPwm02;

                        //sheet.Cells[filterRow + 10, 5].Value2 = "0.9 - 1.249";
                        txtPwRange07.Text = (string)sheet.Cells[filterRow + 88, 5].Value2;

                        sheet.Cells[filterRow + 88, 1].Value2 = "MARQUISE";

                        //sheet.Cells[filterRow + 10, 3].Value2 = "+000 -2";
                        textBox14.Text = (string)sheet.Cells[filterRow + 88, 3].Value2;

                        sheet.Cells[filterRow + 89, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 89, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 89, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 89, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 89, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 89, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 89, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 89, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 89, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 89, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 89, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 89, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 89, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 89, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 89, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeForPwm02 = sheet.Range[sheet.Cells[filterRow + 89, 1], sheet.Cells[filterRow + 89, 14]];
                        //headerRange1.Font.Bold = true;
                        headerRangeForPwm02.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the width range from the txtPwRange TextBox
                        string pwRangeText07 = txtPwRange07.Text;
                        string[] pwRangeParts07 = pwRangeText07.Split('-');
                        double minPwm02 = 0.05;  // Default minimum pw
                        double maxPwm02 = 0.1;  // Default maximum pw

                        // Parse the pw range values if the input is valid
                        if (pwRangeParts07.Length == 2 && double.TryParse(pwRangeParts07[0], out double parsedMinPw07) && double.TryParse(pwRangeParts07[1], out double parsedMaxPw07))
                        {
                            minPwm02 = parsedMinPw07;
                            maxPwm02 = parsedMaxPw07;
                        }

                        Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>> shapesMDictionary02 =
                            new Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>>();

                        Dictionary<string, int> clarityCountssm02 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPwwm02 = new Dictionary<string, double>();

                        string prevStoneShName02 = "";
                        // Loop through rows in excel sheet again to count for each stone name
                        for (int t = 2; t <= rowCount; t++)
                        {
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStoneShName02;
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            double polishWeight = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "marquise" && double.TryParse(polishWeight.ToString(), out double pwValue))
                            {
                                if (pwValue >= minPwm02 && pwValue <= maxPwm02 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCountssm02.ContainsKey(clarity))
                                        {
                                            clarityCountssm02[clarity]++;
                                            clarityPwwm02[clarity] += polishWeight;
                                        }
                                        else
                                        {
                                            clarityCountssm02.Add(clarity, 1);
                                            clarityPwwm02.Add(clarity, polishWeight);
                                        }

                                        if (shapesMDictionary02.ContainsKey(stoneName))
                                        {
                                            shapesMDictionary02[stoneName].Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)> shapes = new List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>();
                                            shapes.Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                            shapesMDictionary02.Add(stoneName, shapes);
                                        }
                                    }
                                }
                                //Console.WriteLine($"Stone Name: {stoneName} - Clarity: {clarity} - PartWT{partWTString} - PW: {polishWeight} - shape: {shape}");
                            }
                            prevStoneShName02 = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING SECOND SIEVE FOR SHAPE 'MARQUISE' CHECKING IF THE PW FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        Dictionary<string, double> clarityWeightDictForPwm001 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> pearParts = partWts.Where(p => p.Item3.ToLower() == "marquise").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForPwm001 = pearParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item4 >= minPwm02 && p.Item4 <= maxPwm02)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDictForPwm001.ContainsKey(clarity))
                                {
                                    clarityWeightDictForPwm001[clarity] += clarityWeightForPwm001;
                                }
                                else
                                {
                                    clarityWeightDictForPwm001[clarity] = clarityWeightForPwm001;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDictForPwm002 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesMDictionary02.Values.SelectMany(shapes => shapes))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDictForPwm002.ContainsKey(clarity))
                            {
                                clarityWeightDictForPwm002[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDictForPwm002[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDictForPwm001 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesMDictionary02.Values.SelectMany(shapes => shapes))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDictForPwm001.ContainsKey(clarity))
                                {
                                    clarityDolarDictForPwm001[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDictForPwm001[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCountForPwm001 = clarityCountssm02.Values.Sum();
                        double totalPwForPwm001 = clarityPwwm02.Values.Sum();

                        int serialNumberForPwm02 = 1;

                        int totalClarityFiltersForPwm02 = clarityValues.Count();
                        int currentClarityFilterForPwm02 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilterForPwm02 / totalClarityFiltersForPwm02 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING SECOND 'MARQUISE' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCountssm02.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 90, 1].Value2 = serialNumberForPwm02;
                                sheet.Cells[filterRow + 90, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 90, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPwForPwm2 = clarityPwwm02[clarity] / totalPwForPwm001 * 100.0;
                                sheet.Cells[filterRow + 90, 3].Value2 = percentPwForPwm2.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 90, 5].Value2 = clarityCountssm02[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 90, 7].Value2 = clarityPwwm02[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeightForPwm001 = clarityWeightDictForPwm001.ContainsKey(clarity) ? clarityWeightDictForPwm001[clarity] : 0.000;
                                sheet.Cells[filterRow + 90, 4].Value2 = Math.Round(clarityWeightForPwm001, 3);

                                // Get the craft weight for this clarity
                                double clarityWeightForPwm002 = clarityWeightDictForPwm002.ContainsKey(clarity) ? clarityWeightDictForPwm002[clarity] : 0.000;
                                sheet.Cells[filterRow + 90, 6].Value2 = Math.Round(clarityWeightForPwm002, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDividedForPwm001 = clarityCountssm02[clarity] / clarityWeightDictForPwm001[clarity];
                                sheet.Cells[filterRow + 90, 8].Value2 = Math.Round(clarityCountDividedForPwm001, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCountForPwm001 = clarityCountssm02[clarity] / clarityPwwm02[clarity];
                                sheet.Cells[filterRow + 90, 9].Value2 = Math.Round(clarityPolishWtCountForPwm001, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCrForPwm001 = clarityPwwm02[clarity] / clarityWeightDictForPwm002[clarity] * 100;
                                sheet.Cells[filterRow + 90, 10].Value2 = percentPwCrForPwm001.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCrForPwm001 = clarityPwwm02[clarity] / clarityWeightDictForPwm001[clarity] * 100;
                                sheet.Cells[filterRow + 90, 11].Value2 = percentRoCrForPwm001.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 90, 12].Value2 = clarityDolarDictForPwm001[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolarForPwm001 = clarityDolarDictForPwm001[clarity] / clarityWeightDictForPwm001[clarity];
                                sheet.Cells[filterRow + 90, 13].Value2 = Math.Round(clarityRoughCrtDolarForPwm001, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolarForPwm001 = clarityWeightDictForPwm001[clarity] / clarityPwwm02[clarity];
                                sheet.Cells[filterRow + 90, 14].Value2 = Math.Round(clarityPolishCrtDolarForPwm001, 2);

                                serialNumberForPwm02++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRangeForPwm02 = sheet.Range[sheet.Cells[filterRow - serialNumberForPwm02 + 90, 1], sheet.Cells[filterRow + 90, 14]];
                        tableRangeForPwm02.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRangeForPwm02.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRangeForPwm02.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 90, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 90, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeightForPwm001 = clarityWeightDictForPwm001.Values.Sum();
                        sheet.Cells[filterRow + 90, 4].Value2 = Math.Round(totalRoughtWeightForPwm001, 2);

                        sheet.Cells[filterRow + 90, 5].Value2 = totalCountForPwm001;

                        double partWTTotalForPwm001 = clarityWeightDictForPwm002.Values.Sum();
                        sheet.Cells[filterRow + 90, 6].Value2 = Math.Round(partWTTotalForPwm001, 2);

                        sheet.Cells[filterRow + 90, 7].Value2 = Math.Round(totalPwForPwm001, 3);

                        double totalSizeForPwm001 = (totalCountForPwm001 / totalRoughtWeightForPwm001);
                        sheet.Cells[filterRow + 90, 8].Value2 = totalSizeForPwm001.ToString("0.00");

                        double polishSizeForPwm001 = (totalCountForPwm001 / totalPwForPwm001);
                        sheet.Cells[filterRow + 90, 9].Value2 = polishSizeForPwm001.ToString("0.00");

                        double crPwPercentageForPwm001 = (totalPwForPwm001 / partWTTotalForPwm001) * 100;
                        sheet.Cells[filterRow + 90, 10].Value2 = crPwPercentageForPwm001.ToString("0.00") + "%";

                        double pwPercentageForPwm001 = (totalPwForPwm001 / totalRoughtWeightForPwm001) * 100;
                        sheet.Cells[filterRow + 90, 11].Value2 = pwPercentageForPwm001.ToString("0.00") + "%";

                        double dolarTotalForPwm001 = clarityDolarDictForPwm001.Values.Sum();
                        sheet.Cells[filterRow + 90, 12].Value2 = Math.Round(dolarTotalForPwm001, 2);

                        double valueRoughForPwm001 = (dolarTotalForPwm001 / totalRoughtWeightForPwm001);
                        sheet.Cells[filterRow + 90, 13].Value2 = valueRoughForPwm001.ToString("0.00");

                        double valuePolishForPwm001 = (dolarTotalForPwm001 / totalPwForPwm001);
                        sheet.Cells[filterRow + 90, 14].Value2 = valuePolishForPwm001.ToString("0.00");


                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeeForPwm02 = sheet.Range[sheet.Cells[filterRow + 90, 1], sheet.Cells[filterRow + 90, 14]];
                        //headerRangee1.Font.Bold = true;
                        headerRangeeForPwm02.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRangeeForPwm02.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion

                    #region 3rd Sieve MARQUISE Method

                    if (checkBoxRunCode008.Checked)
                    {
                        sheet.Cells[filterRow + 93, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 93, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 93, 5].Value2 = "Polish Weight";

                        string updatedValueForPw3 = textBox15.Text;
                        sheet.Cells[filterRow + 94, 3].Value2 = updatedValueForPw3;

                        string updatedValueForPw03 = txtPwRange08.Text;
                        sheet.Cells[filterRow + 94, 5].Value2 = updatedValueForPw03;

                        //sheet.Cells[filterRow + 10, 5].Value2 = "0.9 - 1.249";
                        txtPwRange08.Text = (string)sheet.Cells[filterRow + 94, 5].Value2;

                        sheet.Cells[filterRow + 94, 1].Value2 = "MARQUISE";

                        //sheet.Cells[filterRow + 10, 3].Value2 = "+000 -2";
                        textBox15.Text = (string)sheet.Cells[filterRow + 94, 3].Value2;

                        sheet.Cells[filterRow + 95, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 95, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 95, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 95, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 95, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 95, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 95, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 95, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 95, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 95, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 95, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 95, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 95, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 95, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 95, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeForPwm03 = sheet.Range[sheet.Cells[filterRow + 95, 1], sheet.Cells[filterRow + 95, 14]];
                        //headerRange1.Font.Bold = true;
                        headerRangeForPwm03.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the width range from the txtPwRange TextBox
                        string pwRangeText08 = txtPwRange08.Text;
                        string[] pwRangeParts08 = pwRangeText08.Split('-');
                        double minPwm03 = 0.1;  // Default minimum pw
                        double maxPwm03 = 0.2;  // Default maximum pw

                        // Parse the pw range values if the input is valid
                        if (pwRangeParts08.Length == 2 && double.TryParse(pwRangeParts08[0], out double parsedMinPw08) && double.TryParse(pwRangeParts08[1], out double parsedMaxPw08))
                        {
                            minPwm03 = parsedMinPw08;
                            maxPwm03 = parsedMaxPw08;
                        }

                        Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>> shapesMDictionary03 =
                            new Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>>();

                        Dictionary<string, int> clarityCountssm03 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPwwm03 = new Dictionary<string, double>();

                        string prevStoneShName03 = "";
                        // Loop through rows in excel sheet again to count for each stone name
                        for (int t = 2; t <= rowCount; t++)
                        {
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStoneShName03;
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            double polishWeight = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "marquise" && double.TryParse(polishWeight.ToString(), out double pwValue))
                            {
                                if (pwValue >= minPwm03 && pwValue <= maxPwm03 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCountssm03.ContainsKey(clarity))
                                        {
                                            clarityCountssm03[clarity]++;
                                            clarityPwwm03[clarity] += polishWeight;
                                        }
                                        else
                                        {
                                            clarityCountssm03.Add(clarity, 1);
                                            clarityPwwm03.Add(clarity, polishWeight);
                                        }

                                        if (shapesMDictionary03.ContainsKey(stoneName))
                                        {
                                            shapesMDictionary03[stoneName].Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)> shapes = new List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>();
                                            shapes.Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                            shapesMDictionary03.Add(stoneName, shapes);
                                        }
                                    }
                                }
                            }
                            prevStoneShName03 = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING THIRD SIEVE FOR SHAPE 'MARQUISE' CHECKING IF THE PW FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        Dictionary<string, double> clarityWeightDictForPwm0001 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> pearParts = partWts.Where(p => p.Item3.ToLower() == "marquise").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForPwm0001 = pearParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item4 >= minPwm03 && p.Item4 <= maxPwm03)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDictForPwm0001.ContainsKey(clarity))
                                {
                                    clarityWeightDictForPwm0001[clarity] += clarityWeightForPwm0001;
                                }
                                else
                                {
                                    clarityWeightDictForPwm0001[clarity] = clarityWeightForPwm0001;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDictForPwm0002 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesMDictionary03.Values.SelectMany(shapes => shapes))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDictForPwm0002.ContainsKey(clarity))
                            {
                                clarityWeightDictForPwm0002[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDictForPwm0002[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDictForPwm0001 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesMDictionary03.Values.SelectMany(shapes => shapes))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDictForPwm0001.ContainsKey(clarity))
                                {
                                    clarityDolarDictForPwm0001[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDictForPwm0001[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCountForPwm0001 = clarityCountssm03.Values.Sum();
                        double totalPwForPwm0001 = clarityPwwm03.Values.Sum();

                        int serialNumberForPwm03 = 1;

                        int totalClarityFiltersForPwm03 = clarityValues.Count();
                        int currentClarityFilterForPwm03 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilterForPwm03 / totalClarityFiltersForPwm03 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING THIRD 'MARQUISE' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCountssm03.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 96, 1].Value2 = serialNumberForPwm03;
                                sheet.Cells[filterRow + 96, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 96, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPwForPwm3 = clarityPwwm03[clarity] / totalPwForPwm0001 * 100.0;
                                sheet.Cells[filterRow + 96, 3].Value2 = percentPwForPwm3.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 96, 5].Value2 = clarityCountssm03[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 96, 7].Value2 = clarityPwwm03[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeightForPwm0001 = clarityWeightDictForPwm0001.ContainsKey(clarity) ? clarityWeightDictForPwm0001[clarity] : 0.000;
                                sheet.Cells[filterRow + 96, 4].Value2 = Math.Round(clarityWeightForPwm0001, 3);

                                // Get the craft weight for this clarity
                                double clarityWeightForPwm0002 = clarityWeightDictForPwm0002.ContainsKey(clarity) ? clarityWeightDictForPwm0002[clarity] : 0.000;
                                sheet.Cells[filterRow + 96, 6].Value2 = Math.Round(clarityWeightForPwm0002, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDividedForPwm0001 = clarityCountssm03[clarity] / clarityWeightDictForPwm0001[clarity];
                                sheet.Cells[filterRow + 96, 8].Value2 = Math.Round(clarityCountDividedForPwm0001, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCountForPwm0001 = clarityCountssm03[clarity] / clarityPwwm03[clarity];
                                sheet.Cells[filterRow + 96, 9].Value2 = Math.Round(clarityPolishWtCountForPwm0001, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCrForPwm0001 = clarityPwwm03[clarity] / clarityWeightDictForPwm0002[clarity] * 100;
                                sheet.Cells[filterRow + 96, 10].Value2 = percentPwCrForPwm0001.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCrForPwm0001 = clarityPwwm03[clarity] / clarityWeightDictForPwm0001[clarity] * 100;
                                sheet.Cells[filterRow + 96, 11].Value2 = percentRoCrForPwm0001.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 96, 12].Value2 = clarityDolarDictForPwm0001[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolarForPwm0001 = clarityDolarDictForPwm0001[clarity] / clarityWeightDictForPwm0001[clarity];
                                sheet.Cells[filterRow + 96, 13].Value2 = Math.Round(clarityRoughCrtDolarForPwm0001, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolarForPwm0001 = clarityWeightDictForPwm0001[clarity] / clarityPwwm03[clarity];
                                sheet.Cells[filterRow + 96, 14].Value2 = Math.Round(clarityPolishCrtDolarForPwm0001, 2);

                                serialNumberForPwm03++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRangeForPwm03 = sheet.Range[sheet.Cells[filterRow - serialNumberForPwm03 + 96, 1], sheet.Cells[filterRow + 96, 14]];
                        tableRangeForPwm03.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRangeForPwm03.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRangeForPwm03.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 96, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 96, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeightForPwm0001 = clarityWeightDictForPwm0001.Values.Sum();
                        sheet.Cells[filterRow + 96, 4].Value2 = Math.Round(totalRoughtWeightForPwm0001, 2);

                        sheet.Cells[filterRow + 96, 5].Value2 = totalCountForPwm0001;

                        double partWTTotalForPwm0001 = clarityWeightDictForPwm0002.Values.Sum();
                        sheet.Cells[filterRow + 96, 6].Value2 = Math.Round(partWTTotalForPwm0001, 2);

                        sheet.Cells[filterRow + 96, 7].Value2 = Math.Round(totalPwForPwm0001, 3);

                        double totalSizeForPwm0001 = (totalCountForPwm0001 / totalRoughtWeightForPwm0001);
                        sheet.Cells[filterRow + 96, 8].Value2 = totalSizeForPwm0001.ToString("0.00");

                        double polishSizeForPwm0001 = (totalCountForPwm0001 / totalPwForPwm0001);
                        sheet.Cells[filterRow + 96, 9].Value2 = polishSizeForPwm0001.ToString("0.00");

                        double crPwPercentageForPwm0001 = (totalPwForPwm0001 / partWTTotalForPwm0001) * 100;
                        sheet.Cells[filterRow + 96, 10].Value2 = crPwPercentageForPwm0001.ToString("0.00") + "%";

                        double pwPercentageForPwm0001 = (totalPwForPwm0001 / totalRoughtWeightForPwm0001) * 100;
                        sheet.Cells[filterRow + 96, 11].Value2 = pwPercentageForPwm0001.ToString("0.00") + "%";

                        double dolarTotalForPwm0001 = clarityDolarDictForPwm0001.Values.Sum();
                        sheet.Cells[filterRow + 96, 12].Value2 = Math.Round(dolarTotalForPwm0001, 2);

                        double valueRoughForPwm0001 = (dolarTotalForPwm0001 / totalRoughtWeightForPwm0001);
                        sheet.Cells[filterRow + 96, 13].Value2 = valueRoughForPwm0001.ToString("0.00");

                        double valuePolishForPwm0001 = (dolarTotalForPwm0001 / totalPwForPwm0001);
                        sheet.Cells[filterRow + 96, 14].Value2 = valuePolishForPwm0001.ToString("0.00");


                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeeForPwm03 = sheet.Range[sheet.Cells[filterRow + 96, 1], sheet.Cells[filterRow + 96, 14]];
                        //headerRangee1.Font.Bold = true;
                        headerRangeeForPwm03.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRangeeForPwm03.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion

                    #region 4th Sieve MARQUISE Method

                    if (checkBoxRunCode009.Checked)
                    {
                        sheet.Cells[filterRow + 99, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 99, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 99, 5].Value2 = "Polish Weight";

                        string updatedValueForPw4 = textBox16.Text;
                        sheet.Cells[filterRow + 100, 3].Value2 = updatedValueForPw4;

                        string updatedValueForPw04 = txtPwRange09.Text;
                        sheet.Cells[filterRow + 100, 5].Value2 = updatedValueForPw04;

                        //sheet.Cells[filterRow + 10, 5].Value2 = "0.9 - 1.249";
                        txtPwRange09.Text = (string)sheet.Cells[filterRow + 100, 5].Value2;

                        sheet.Cells[filterRow + 100, 1].Value2 = "MARQUISE";

                        //sheet.Cells[filterRow + 10, 3].Value2 = "+000 -2";
                        textBox16.Text = (string)sheet.Cells[filterRow + 100, 3].Value2;

                        sheet.Cells[filterRow + 101, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 101, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 101, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 101, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 101, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 101, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 101, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 101, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 101, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 101, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 101, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 101, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 101, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 101, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 101, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeForPwm04 = sheet.Range[sheet.Cells[filterRow + 101, 1], sheet.Cells[filterRow + 101, 14]];
                        //headerRange1.Font.Bold = true;
                        headerRangeForPwm04.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the Pw range from the txtPwRange04 TextBox
                        string pwRangeText09 = txtPwRange09.Text;
                        string[] pwRangeParts09 = pwRangeText09.Split('-');
                        double minPwm04 = 0.3;  // Default minimum pw
                        double maxPwm04 = 0.4;  // Default maximum pw

                        // Parse the pw range values if the input is valid
                        if (pwRangeParts09.Length == 2 && double.TryParse(pwRangeParts09[0], out double parsedMinPw09) && double.TryParse(pwRangeParts09[1], out double parsedMaxPw09))
                        {
                            minPwm04 = parsedMinPw09;
                            maxPwm04 = parsedMaxPw09;
                        }

                        Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>> shapesMDictionary04 =
                            new Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>>();

                        Dictionary<string, int> clarityCountssm04 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPwwm04 = new Dictionary<string, double>();

                        string prevStoneShName04 = "";
                        // Loop through rows in excel sheet again to count for each stone name
                        for (int t = 2; t <= rowCount; t++)
                        {
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStoneShName04;
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            double polishWeight = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "marquise" && double.TryParse(polishWeight.ToString(), out double pwValue))
                            {
                                if (pwValue >= minPwm04 && pwValue <= maxPwm04 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCountssm04.ContainsKey(clarity))
                                        {
                                            clarityCountssm04[clarity]++;
                                            clarityPwwm04[clarity] += polishWeight;
                                        }
                                        else
                                        {
                                            clarityCountssm04.Add(clarity, 1);
                                            clarityPwwm04.Add(clarity, polishWeight);
                                        }

                                        if (shapesMDictionary04.ContainsKey(stoneName))
                                        {
                                            shapesMDictionary04[stoneName].Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)> shapes = new List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>();
                                            shapes.Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                            shapesMDictionary04.Add(stoneName, shapes);
                                        }
                                    }
                                }
                            }
                            prevStoneShName04 = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING FOURTH SIEVE FOR SHAPE 'MARQUISE' CHECKING IF THE PW FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        Dictionary<string, double> clarityWeightDictForPwm00001 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> pearParts = partWts.Where(p => p.Item3.ToLower() == "marquise").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForPwm00001 = pearParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item4 >= minPwm04 && p.Item4 <= maxPwm04)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDictForPwm00001.ContainsKey(clarity))
                                {
                                    clarityWeightDictForPwm00001[clarity] += clarityWeightForPwm00001;
                                }
                                else
                                {
                                    clarityWeightDictForPwm00001[clarity] = clarityWeightForPwm00001;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDictForPwm00002 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesMDictionary04.Values.SelectMany(shapes => shapes))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDictForPwm00002.ContainsKey(clarity))
                            {
                                clarityWeightDictForPwm00002[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDictForPwm00002[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDictForPwm00001 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesMDictionary04.Values.SelectMany(shapes => shapes))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDictForPwm00001.ContainsKey(clarity))
                                {
                                    clarityDolarDictForPwm00001[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDictForPwm00001[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCountForPwm00001 = clarityCountssm04.Values.Sum();
                        double totalPwForPwm00001 = clarityPwwm04.Values.Sum();

                        int serialNumberForPwm04 = 1;

                        int totalClarityFiltersForPwm04 = clarityValues.Count();
                        int currentClarityFilterForPwm04 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilterForPwm04 / totalClarityFiltersForPwm04 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING FOURTH 'MARQUISE' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCountssm04.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 102, 1].Value2 = serialNumberForPwm04;
                                sheet.Cells[filterRow + 102, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 102, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPwForPwm4 = clarityPwwm04[clarity] / totalPwForPwm00001 * 100.0;
                                sheet.Cells[filterRow + 102, 3].Value2 = percentPwForPwm4.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 102, 5].Value2 = clarityCountssm04[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 102, 7].Value2 = clarityPwwm04[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeightForPwm00001 = clarityWeightDictForPwm00001.ContainsKey(clarity) ? clarityWeightDictForPwm00001[clarity] : 0.000;
                                sheet.Cells[filterRow + 102, 4].Value2 = Math.Round(clarityWeightForPwm00001, 3);

                                // Get the craft weight for this clarity
                                double clarityWeightForPwm00002 = clarityWeightDictForPwm00002.ContainsKey(clarity) ? clarityWeightDictForPwm00002[clarity] : 0.000;
                                sheet.Cells[filterRow + 102, 6].Value2 = Math.Round(clarityWeightForPwm00002, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDividedForPwm00001 = clarityCountssm04[clarity] / clarityWeightDictForPwm00001[clarity];
                                sheet.Cells[filterRow + 102, 8].Value2 = Math.Round(clarityCountDividedForPwm00001, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCountForPwm00001 = clarityCountssm04[clarity] / clarityPwwm04[clarity];
                                sheet.Cells[filterRow + 102, 9].Value2 = Math.Round(clarityPolishWtCountForPwm00001, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCrForPwm00001 = clarityPwwm04[clarity] / clarityWeightDictForPwm00002[clarity] * 100;
                                sheet.Cells[filterRow + 102, 10].Value2 = percentPwCrForPwm00001.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCrForPwm00001 = clarityPwwm04[clarity] / clarityWeightDictForPwm00001[clarity] * 100;
                                sheet.Cells[filterRow + 102, 11].Value2 = percentRoCrForPwm00001.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 102, 12].Value2 = clarityDolarDictForPwm00001[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolarForPwm00001 = clarityDolarDictForPwm00001[clarity] / clarityWeightDictForPwm00001[clarity];
                                sheet.Cells[filterRow + 102, 13].Value2 = Math.Round(clarityRoughCrtDolarForPwm00001, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolarForPwm00001 = clarityWeightDictForPwm00001[clarity] / clarityPwwm04[clarity];
                                sheet.Cells[filterRow + 102, 14].Value2 = Math.Round(clarityPolishCrtDolarForPwm00001, 2);

                                serialNumberForPwm04++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRangeForPwm04 = sheet.Range[sheet.Cells[filterRow - serialNumberForPwm04 + 102, 1], sheet.Cells[filterRow + 102, 14]];
                        tableRangeForPwm04.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRangeForPwm04.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRangeForPwm04.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 102, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 102, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeightForPwm00001 = clarityWeightDictForPwm00001.Values.Sum();
                        sheet.Cells[filterRow + 102, 4].Value2 = Math.Round(totalRoughtWeightForPwm00001, 2);

                        sheet.Cells[filterRow + 102, 5].Value2 = totalCountForPwm00001;

                        double partWTTotalForPwm00001 = clarityWeightDictForPwm00002.Values.Sum();
                        sheet.Cells[filterRow + 102, 6].Value2 = Math.Round(partWTTotalForPwm00001, 2);

                        sheet.Cells[filterRow + 102, 7].Value2 = Math.Round(totalPwForPwm00001, 3);

                        double totalSizeForPwm00001 = (totalCountForPwm00001 / totalRoughtWeightForPwm00001);
                        sheet.Cells[filterRow + 102, 8].Value2 = totalSizeForPwm00001.ToString("0.00");

                        double polishSizeForPwm00001 = (totalCountForPwm00001 / totalPwForPwm00001);
                        sheet.Cells[filterRow + 102, 9].Value2 = polishSizeForPwm00001.ToString("0.00");

                        double crPwPercentageForPwm00001 = (totalPwForPwm00001 / partWTTotalForPwm00001) * 100;
                        sheet.Cells[filterRow + 102, 10].Value2 = crPwPercentageForPwm00001.ToString("0.00") + "%";

                        double pwPercentageForPwm00001 = (totalPwForPwm00001 / totalRoughtWeightForPwm00001) * 100;
                        sheet.Cells[filterRow + 102, 11].Value2 = pwPercentageForPwm00001.ToString("0.00") + "%";

                        double dolarTotalForPwm00001 = clarityDolarDictForPwm00001.Values.Sum();
                        sheet.Cells[filterRow + 102, 12].Value2 = Math.Round(dolarTotalForPwm00001, 2);

                        double valueRoughForPwm00001 = (dolarTotalForPwm00001 / totalRoughtWeightForPwm00001);
                        sheet.Cells[filterRow + 102, 13].Value2 = valueRoughForPwm00001.ToString("0.00");

                        double valuePolishForPwm00001 = (dolarTotalForPwm00001 / totalPwForPwm00001);
                        sheet.Cells[filterRow + 102, 14].Value2 = valuePolishForPwm00001.ToString("0.00");


                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeeForPwm04 = sheet.Range[sheet.Cells[filterRow + 102, 1], sheet.Cells[filterRow + 102, 14]];
                        //headerRangee1.Font.Bold = true;
                        headerRangeeForPwm04.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRangeeForPwm04.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion

                    #region 5th Sieve MARQUISE Method

                    if (checkBoxRunCode010.Checked)
                    {
                        sheet.Cells[filterRow + 105, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 105, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 105, 5].Value2 = "Polish Weight";

                        string updatedValueForPw5 = textBox17.Text;
                        sheet.Cells[filterRow + 106, 3].Value2 = updatedValueForPw5;

                        string updatedValueForPw05 = txtPwRange10.Text;
                        sheet.Cells[filterRow + 106, 5].Value2 = updatedValueForPw05;

                        //sheet.Cells[filterRow + 10, 5].Value2 = "0.9 - 1.249";
                        txtPwRange10.Text = (string)sheet.Cells[filterRow + 106, 5].Value2;

                        sheet.Cells[filterRow + 106, 1].Value2 = "MARQUISE";

                        //sheet.Cells[filterRow + 10, 3].Value2 = "+000 -2";
                        textBox17.Text = (string)sheet.Cells[filterRow + 106, 3].Value2;

                        sheet.Cells[filterRow + 107, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 107, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 107, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 107, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 107, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 107, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 107, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 107, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 107, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 107, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 107, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 107, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 107, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 107, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 107, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeForPwm05 = sheet.Range[sheet.Cells[filterRow + 107, 1], sheet.Cells[filterRow + 107, 14]];
                        //headerRange1.Font.Bold = true;
                        headerRangeForPwm05.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the Pw range from the txtPwRange04 TextBox
                        string pwRangeText10 = txtPwRange10.Text;
                        string[] pwRangeParts10 = pwRangeText10.Split('-');
                        double minPwm05 = 0.5;  // Default minimum pw
                        double maxPwm05 = 0.6;  // Default maximum pw

                        // Parse the pw range values if the input is valid
                        if (pwRangeParts10.Length == 2 && double.TryParse(pwRangeParts10[0], out double parsedMinPw10) && double.TryParse(pwRangeParts10[1], out double parsedMaxPw10))
                        {
                            minPwm05 = parsedMinPw10;
                            maxPwm05 = parsedMaxPw10;
                        }

                        Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>> shapesMDictionary05 =
                            new Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>>();

                        Dictionary<string, int> clarityCountssm05 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPwwm05 = new Dictionary<string, double>();

                        string prevStoneShName05 = "";
                        // Loop through rows in excel sheet again to count for each stone name
                        for (int t = 2; t <= rowCount; t++)
                        {
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStoneShName05;
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            double polishWeight = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "marquise" && double.TryParse(polishWeight.ToString(), out double pwValue))
                            {
                                if (pwValue >= minPwm05 && pwValue <= maxPwm05 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCountssm05.ContainsKey(clarity))
                                        {
                                            clarityCountssm05[clarity]++;
                                            clarityPwwm05[clarity] += polishWeight;
                                        }
                                        else
                                        {
                                            clarityCountssm05.Add(clarity, 1);
                                            clarityPwwm05.Add(clarity, polishWeight);
                                        }

                                        if (shapesMDictionary05.ContainsKey(stoneName))
                                        {
                                            shapesMDictionary05[stoneName].Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)> shapes = new List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>();
                                            shapes.Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                            shapesMDictionary05.Add(stoneName, shapes);
                                        }
                                    }
                                }
                            }
                            prevStoneShName05 = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING FIFTH SIEVE FOR SHAPE 'MARQUISE' CHECKING IF THE PW FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        Dictionary<string, double> clarityWeightDictForPwm000001 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> pearParts = partWts.Where(p => p.Item3.ToLower() == "marquise").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForPwm000001 = pearParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item4 >= minPwm05 && p.Item4 <= maxPwm05)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDictForPwm000001.ContainsKey(clarity))
                                {
                                    clarityWeightDictForPwm000001[clarity] += clarityWeightForPwm000001;
                                }
                                else
                                {
                                    clarityWeightDictForPwm000001[clarity] = clarityWeightForPwm000001;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDictForPwm000002 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesMDictionary05.Values.SelectMany(shapes => shapes))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDictForPwm000002.ContainsKey(clarity))
                            {
                                clarityWeightDictForPwm000002[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDictForPwm000002[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDictForPwm000001 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesMDictionary05.Values.SelectMany(shapes => shapes))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDictForPwm000001.ContainsKey(clarity))
                                {
                                    clarityDolarDictForPwm000001[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDictForPwm000001[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCountForPwm000001 = clarityCountssm05.Values.Sum();
                        double totalPwForPwm000001 = clarityPwwm05.Values.Sum();

                        int serialNumberForPwm05 = 1;

                        int totalClarityFiltersForPwm05 = clarityValues.Count();
                        int currentClarityFilterForPwm05 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilterForPwm05 / totalClarityFiltersForPwm05 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING FIFTH 'MARQUISE' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCountssm05.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 108, 1].Value2 = serialNumberForPwm05;
                                sheet.Cells[filterRow + 108, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 108, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPwForPwm5 = clarityPwwm05[clarity] / totalPwForPwm000001 * 100.0;
                                sheet.Cells[filterRow + 108, 3].Value2 = percentPwForPwm5.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 108, 5].Value2 = clarityCountssm05[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 108, 7].Value2 = clarityPwwm05[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeightForPwm000001 = clarityWeightDictForPwm000001.ContainsKey(clarity) ? clarityWeightDictForPwm000001[clarity] : 0.000;
                                sheet.Cells[filterRow + 108, 4].Value2 = Math.Round(clarityWeightForPwm000001, 3);

                                // Get the craft weight for this clarity
                                double clarityWeightForPwm000002 = clarityWeightDictForPwm000002.ContainsKey(clarity) ? clarityWeightDictForPwm000002[clarity] : 0.000;
                                sheet.Cells[filterRow + 108, 6].Value2 = Math.Round(clarityWeightForPwm000002, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDividedForPwm000001 = clarityCountssm05[clarity] / clarityWeightDictForPwm000001[clarity];
                                sheet.Cells[filterRow + 108, 8].Value2 = Math.Round(clarityCountDividedForPwm000001, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCountForPwm000001 = clarityCountssm05[clarity] / clarityPwwm05[clarity];
                                sheet.Cells[filterRow + 108, 9].Value2 = Math.Round(clarityPolishWtCountForPwm000001, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCrForPwm000001 = clarityPwwm05[clarity] / clarityWeightDictForPwm000002[clarity] * 100;
                                sheet.Cells[filterRow + 108, 10].Value2 = percentPwCrForPwm000001.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCrForPwm000001 = clarityPwwm05[clarity] / clarityWeightDictForPwm000001[clarity] * 100;
                                sheet.Cells[filterRow + 108, 11].Value2 = percentRoCrForPwm000001.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 108, 12].Value2 = clarityDolarDictForPwm000001[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolarForPwm000001 = clarityDolarDictForPwm000001[clarity] / clarityWeightDictForPwm000001[clarity];
                                sheet.Cells[filterRow + 108, 13].Value2 = Math.Round(clarityRoughCrtDolarForPwm000001, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolarForPwm000001 = clarityWeightDictForPwm000001[clarity] / clarityPwwm05[clarity];
                                sheet.Cells[filterRow + 108, 14].Value2 = Math.Round(clarityPolishCrtDolarForPwm000001, 2);

                                serialNumberForPwm05++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRangeForPwm05 = sheet.Range[sheet.Cells[filterRow - serialNumberForPwm05 + 108, 1], sheet.Cells[filterRow + 108, 14]];
                        tableRangeForPwm05.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRangeForPwm05.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRangeForPwm05.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 108, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 108, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeightForPwm000001 = clarityWeightDictForPwm000001.Values.Sum();
                        sheet.Cells[filterRow + 108, 4].Value2 = Math.Round(totalRoughtWeightForPwm000001, 2);

                        sheet.Cells[filterRow + 108, 5].Value2 = totalCountForPwm000001;

                        double partWTTotalForPwm000001 = clarityWeightDictForPwm000002.Values.Sum();
                        sheet.Cells[filterRow + 108, 6].Value2 = Math.Round(partWTTotalForPwm000001, 2);

                        sheet.Cells[filterRow + 108, 7].Value2 = Math.Round(totalPwForPwm000001, 3);

                        double totalSizeForPwm000001 = (totalCountForPwm000001 / totalRoughtWeightForPwm000001);
                        sheet.Cells[filterRow + 108, 8].Value2 = totalSizeForPwm000001.ToString("0.00");

                        double polishSizeForPwm000001 = (totalCountForPwm000001 / totalPwForPwm000001);
                        sheet.Cells[filterRow + 108, 9].Value2 = polishSizeForPwm000001.ToString("0.00");

                        double crPwPercentageForPwm000001 = (totalPwForPwm000001 / partWTTotalForPwm000001) * 100;
                        sheet.Cells[filterRow + 108, 10].Value2 = crPwPercentageForPwm000001.ToString("0.00") + "%";

                        double pwPercentageForPwm000001 = (totalPwForPwm000001 / totalRoughtWeightForPwm000001) * 100;
                        sheet.Cells[filterRow + 108, 11].Value2 = pwPercentageForPwm000001.ToString("0.00") + "%";

                        double dolarTotalForPwm000001 = clarityDolarDictForPwm000001.Values.Sum();
                        sheet.Cells[filterRow + 108, 12].Value2 = Math.Round(dolarTotalForPwm000001, 2);

                        double valueRoughForPwm000001 = (dolarTotalForPwm000001 / totalRoughtWeightForPwm000001);
                        sheet.Cells[filterRow + 108, 13].Value2 = valueRoughForPwm000001.ToString("0.00");

                        double valuePolishForPwm000001 = (dolarTotalForPwm000001 / totalPwForPwm000001);
                        sheet.Cells[filterRow + 108, 14].Value2 = valuePolishForPwm000001.ToString("0.00");


                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeeForPwm05 = sheet.Range[sheet.Cells[filterRow + 108, 1], sheet.Cells[filterRow + 108, 14]];
                        //headerRangee1.Font.Bold = true;
                        headerRangeeForPwm05.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRangeeForPwm05.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion


                    #region 1st Sieve EMERALD 4STEP Method

                    if (checkBoxRunCode011.Checked)
                    {
                        sheet.Cells[filterRow + 111, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 111, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 111, 5].Value2 = "Polish Weight";

                        string updatedValueForPwm1 = textBox18.Text;
                        sheet.Cells[filterRow + 112, 3].Value2 = updatedValueForPwm1;

                        string updatedValueForPwm01 = txtPwRange11.Text;
                        sheet.Cells[filterRow + 112, 5].Value2 = updatedValueForPwm01;

                        //sheet.Cells[filterRow + 10, 5].Value2 = "0.9 - 1.249";
                        txtPwRange11.Text = (string)sheet.Cells[filterRow + 112, 5].Value2;

                        sheet.Cells[filterRow + 112, 1].Value2 = "EMERALD 4STEP";

                        //sheet.Cells[filterRow + 10, 3].Value2 = "+000 -2";
                        textBox18.Text = (string)sheet.Cells[filterRow + 112, 3].Value2;

                        sheet.Cells[filterRow + 113, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 113, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 113, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 113, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 113, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 113, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 113, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 113, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 113, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 113, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 113, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 113, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 113, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 113, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 113, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeForPwe01 = sheet.Range[sheet.Cells[filterRow + 113, 1], sheet.Cells[filterRow + 113, 14]];
                        //headerRange1.Font.Bold = true;
                        headerRangeForPwe01.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the pw range from the txtPwRange01 TextBox
                        string pwRangeText11 = txtPwRange11.Text;
                        string[] pwRangeParts11 = pwRangeText11.Split('-');
                        double minPwe01 = 0.00;  // Default minimum pw
                        double maxPwe01 = 0.05;  // Default maximum pw

                        // Parse the width range values if the input is valid
                        if (pwRangeParts11.Length == 2 && double.TryParse(pwRangeParts11[0], out double parsedMinPw11) && double.TryParse(pwRangeParts11[1], out double parsedMaxPw11))
                        {
                            minPwe01 = parsedMinPw11;
                            maxPwe01 = parsedMaxPw11;
                        }

                        Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>> shapesEDictionary =
                            new Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>>();

                        Dictionary<string, int> clarityCountsse01 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPwwe01 = new Dictionary<string, double>();

                        string prevStoneShName = "";
                        // Loop through rows in excel sheet again to count Dolar for each stone name
                        for (int t = 2; t <= rowCount; t++)
                        {
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStoneShName;
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            double polishWeight = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "emerald 4step" && double.TryParse(polishWeight.ToString(), out double pwValue))
                            {
                                if (pwValue >= minPwe01 && pwValue <= maxPwe01 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCountsse01.ContainsKey(clarity))
                                        {
                                            clarityCountsse01[clarity]++;
                                            clarityPwwe01[clarity] += polishWeight;
                                        }
                                        else
                                        {
                                            clarityCountsse01.Add(clarity, 1);
                                            clarityPwwe01.Add(clarity, polishWeight);
                                        }

                                        if (shapesEDictionary.ContainsKey(stoneName))
                                        {
                                            shapesEDictionary[stoneName].Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)> shapes = new List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>();
                                            shapes.Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                            shapesEDictionary.Add(stoneName, shapes);
                                        }
                                    }
                                }
                                //Console.WriteLine($"Stone Name: {stoneName} - Clarity: {clarity} - PartWT{partWTString} - PW: {polishWeight} - shape: {shape}");
                            }
                            prevStoneShName = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING FIRST SIEVE FOR SHAPE 'EMERALD 4STEP' CHECKING IF THE PW FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        Dictionary<string, double> clarityWeightDictForPwe01 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> pearParts = partWts.Where(p => p.Item3.ToLower() == "emerald 4step").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForPwe01 = pearParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item4 >= minPwe01 && p.Item4 <= maxPwe01)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDictForPwe01.ContainsKey(clarity))
                                {
                                    clarityWeightDictForPwe01[clarity] += clarityWeightForPwe01;
                                }
                                else
                                {
                                    clarityWeightDictForPwe01[clarity] = clarityWeightForPwe01;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDictForPwe02 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesEDictionary.Values.SelectMany(shapes => shapes))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDictForPwe02.ContainsKey(clarity))
                            {
                                clarityWeightDictForPwe02[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDictForPwe02[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDictForPwe01 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesEDictionary.Values.SelectMany(shapes => shapes))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDictForPwe01.ContainsKey(clarity))
                                {
                                    clarityDolarDictForPwe01[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDictForPwe01[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCountForPwe01 = clarityCountsse01.Values.Sum();
                        double totalPwForPwe01 = clarityPwwe01.Values.Sum();

                        int serialNumberForPwe01 = 1;

                        int totalClarityFiltersForPwe01 = clarityValues.Count();
                        int currentClarityFilterForPwe01 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilterForPwe01 / totalClarityFiltersForPwe01 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING FIRST 'EMERALD 4STEP' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCountsse01.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 114, 1].Value2 = serialNumberForPwe01;
                                sheet.Cells[filterRow + 114, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 114, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPwForPwe1 = clarityPwwe01[clarity] / totalPwForPwe01 * 100.0;
                                sheet.Cells[filterRow + 114, 3].Value2 = percentPwForPwe1.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 114, 5].Value2 = clarityCountsse01[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 114, 7].Value2 = clarityPwwe01[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeightForPwe01 = clarityWeightDictForPwe01.ContainsKey(clarity) ? clarityWeightDictForPwe01[clarity] : 0.000;
                                sheet.Cells[filterRow + 114, 4].Value2 = Math.Round(clarityWeightForPwe01, 3);

                                // Get the craft weight for this clarity
                                double clarityWeightForPwe02 = clarityWeightDictForPwe02.ContainsKey(clarity) ? clarityWeightDictForPwe02[clarity] : 0.000;
                                sheet.Cells[filterRow + 114, 6].Value2 = Math.Round(clarityWeightForPwe02, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDividedForPwe01 = clarityCountsse01[clarity] / clarityWeightDictForPwe01[clarity];
                                sheet.Cells[filterRow + 114, 8].Value2 = Math.Round(clarityCountDividedForPwe01, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCountForPwe01 = clarityCountsse01[clarity] / clarityPwwe01[clarity];
                                sheet.Cells[filterRow + 114, 9].Value2 = Math.Round(clarityPolishWtCountForPwe01, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCrForPwe01 = clarityPwwe01[clarity] / clarityWeightDictForPwe02[clarity] * 100;
                                sheet.Cells[filterRow + 114, 10].Value2 = percentPwCrForPwe01.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCrForPwe01 = clarityPwwe01[clarity] / clarityWeightDictForPwe01[clarity] * 100;
                                sheet.Cells[filterRow + 114, 11].Value2 = percentRoCrForPwe01.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 114, 12].Value2 = clarityDolarDictForPwe01[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolarForPwe01 = clarityDolarDictForPwe01[clarity] / clarityWeightDictForPwe01[clarity];
                                sheet.Cells[filterRow + 114, 13].Value2 = Math.Round(clarityRoughCrtDolarForPwe01, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolarForPwe01 = clarityWeightDictForPwe01[clarity] / clarityPwwe01[clarity];
                                sheet.Cells[filterRow + 114, 14].Value2 = Math.Round(clarityPolishCrtDolarForPwe01, 2);

                                serialNumberForPwe01++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRangeForPwe01 = sheet.Range[sheet.Cells[filterRow - serialNumberForPwe01 + 114, 1], sheet.Cells[filterRow + 114, 14]];
                        tableRangeForPwe01.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRangeForPwe01.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRangeForPwe01.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 114, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 114, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeightForPwe01 = clarityWeightDictForPwe01.Values.Sum();
                        sheet.Cells[filterRow + 114, 4].Value2 = Math.Round(totalRoughtWeightForPwe01, 2);

                        sheet.Cells[filterRow + 114, 5].Value2 = totalCountForPwe01;

                        double partWTTotalForPwe01 = clarityWeightDictForPwe02.Values.Sum();
                        sheet.Cells[filterRow + 114, 6].Value2 = Math.Round(partWTTotalForPwe01, 2);

                        sheet.Cells[filterRow + 114, 7].Value2 = Math.Round(totalPwForPwe01, 3);

                        double totalSizeForPwe01 = (totalCountForPwe01 / totalRoughtWeightForPwe01);
                        sheet.Cells[filterRow + 114, 8].Value2 = totalSizeForPwe01.ToString("0.00");

                        double polishSizeForPwe01 = (totalCountForPwe01 / totalPwForPwe01);
                        sheet.Cells[filterRow + 114, 9].Value2 = polishSizeForPwe01.ToString("0.00");

                        double crPwPercentageForPwe01 = (totalPwForPwe01 / partWTTotalForPwe01) * 100;
                        sheet.Cells[filterRow + 114, 10].Value2 = crPwPercentageForPwe01.ToString("0.00") + "%";

                        double pwPercentageForPwe01 = (totalPwForPwe01 / totalRoughtWeightForPwe01) * 100;
                        sheet.Cells[filterRow + 114, 11].Value2 = pwPercentageForPwe01.ToString("0.00") + "%";

                        double dolarTotalForPwe01 = clarityDolarDictForPwe01.Values.Sum();
                        sheet.Cells[filterRow + 114, 12].Value2 = Math.Round(dolarTotalForPwe01, 2);

                        double valueRoughForPwe01 = (dolarTotalForPwe01 / totalRoughtWeightForPwe01);
                        sheet.Cells[filterRow + 114, 13].Value2 = valueRoughForPwe01.ToString("0.00");

                        double valuePolishForPwe01 = (dolarTotalForPwe01 / totalPwForPwe01);
                        sheet.Cells[filterRow + 114, 14].Value2 = valuePolishForPwe01.ToString("0.00");

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeeForPwe01 = sheet.Range[sheet.Cells[filterRow + 114, 1], sheet.Cells[filterRow + 114, 14]];
                        //headerRangee1.Font.Bold = true;
                        headerRangeeForPwe01.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRangeeForPwe01.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion

                    #region 2nd Sieve EMERALD 4STEP Method

                    if (checkBoxRunCode012.Checked)
                    {
                        sheet.Cells[filterRow + 117, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 117, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 117, 5].Value2 = "Polish Weight";

                        string updatedValueForPwe2 = textBox19.Text;
                        sheet.Cells[filterRow + 118, 3].Value2 = updatedValueForPwe2;

                        string updatedValueForPwe02 = txtPwRange12.Text;
                        sheet.Cells[filterRow + 118, 5].Value2 = updatedValueForPwe02;

                        //sheet.Cells[filterRow + 10, 5].Value2 = "0.9 - 1.249";
                        txtPwRange12.Text = (string)sheet.Cells[filterRow + 118, 5].Value2;

                        sheet.Cells[filterRow + 118, 1].Value2 = "EMERALD 4STEP";

                        //sheet.Cells[filterRow + 10, 3].Value2 = "+000 -2";
                        textBox19.Text = (string)sheet.Cells[filterRow + 118, 3].Value2;

                        sheet.Cells[filterRow + 119, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 119, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 119, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 119, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 119, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 119, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 119, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 119, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 119, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 119, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 119, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 119, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 119, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 119, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 119, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeForPwe02 = sheet.Range[sheet.Cells[filterRow + 119, 1], sheet.Cells[filterRow + 119, 14]];
                        //headerRange1.Font.Bold = true;
                        headerRangeForPwe02.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the width range from the txtPwRange TextBox
                        string pwRangeText12 = txtPwRange12.Text;
                        string[] pwRangeParts12 = pwRangeText12.Split('-');
                        double minPwe02 = 0.05;  // Default minimum pw
                        double maxPwe02 = 0.1;  // Default maximum pw

                        // Parse the pw range values if the input is valid
                        if (pwRangeParts12.Length == 2 && double.TryParse(pwRangeParts12[0], out double parsedMinPw12) && double.TryParse(pwRangeParts12[1], out double parsedMaxPw12))
                        {
                            minPwe02 = parsedMinPw12;
                            maxPwe02 = parsedMaxPw12;
                        }

                        Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>> shapesEDictionary02 =
                            new Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>>();

                        Dictionary<string, int> clarityCountsse02 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPwwe02 = new Dictionary<string, double>();

                        string prevStoneShName02 = "";
                        // Loop through rows in excel sheet again to count for each stone name
                        for (int t = 2; t <= rowCount; t++)
                        {
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStoneShName02;
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            double polishWeight = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "emerald 4step" && double.TryParse(polishWeight.ToString(), out double pwValue))
                            {
                                if (pwValue >= minPwe02 && pwValue <= maxPwe02 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCountsse02.ContainsKey(clarity))
                                        {
                                            clarityCountsse02[clarity]++;
                                            clarityPwwe02[clarity] += polishWeight;
                                        }
                                        else
                                        {
                                            clarityCountsse02.Add(clarity, 1);
                                            clarityPwwe02.Add(clarity, polishWeight);
                                        }

                                        if (shapesEDictionary02.ContainsKey(stoneName))
                                        {
                                            shapesEDictionary02[stoneName].Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)> shapes = new List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>();
                                            shapes.Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                            shapesEDictionary02.Add(stoneName, shapes);
                                        }
                                    }
                                }
                            }
                            prevStoneShName02 = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING SECOND SIEVE FOR SHAPE 'EMERALD 4STEP' CHECKING IF THE PW FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        Dictionary<string, double> clarityWeightDictForPwe001 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> pearParts = partWts.Where(p => p.Item3.ToLower() == "emerald 4step").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForPwe001 = pearParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item4 >= minPwe02 && p.Item4 <= maxPwe02)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDictForPwe001.ContainsKey(clarity))
                                {
                                    clarityWeightDictForPwe001[clarity] += clarityWeightForPwe001;
                                }
                                else
                                {
                                    clarityWeightDictForPwe001[clarity] = clarityWeightForPwe001;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDictForPwe002 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesEDictionary02.Values.SelectMany(shapes => shapes))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDictForPwe002.ContainsKey(clarity))
                            {
                                clarityWeightDictForPwe002[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDictForPwe002[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDictForPwe001 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesEDictionary02.Values.SelectMany(shapes => shapes))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDictForPwe001.ContainsKey(clarity))
                                {
                                    clarityDolarDictForPwe001[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDictForPwe001[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCountForPwe001 = clarityCountsse02.Values.Sum();
                        double totalPwForPwe001 = clarityPwwe02.Values.Sum();

                        int serialNumberForPwe02 = 1;

                        int totalClarityFiltersForPwe02 = clarityValues.Count();
                        int currentClarityFilterForPwe02 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilterForPwe02 / totalClarityFiltersForPwe02 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING SECOND 'EMERALD 4STEP' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCountsse02.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 120, 1].Value2 = serialNumberForPwe02;
                                sheet.Cells[filterRow + 120, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 120, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPwForPwe2 = clarityPwwe02[clarity] / totalPwForPwe001 * 100.0;
                                sheet.Cells[filterRow + 120, 3].Value2 = percentPwForPwe2.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 120, 5].Value2 = clarityCountsse02[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 120, 7].Value2 = clarityPwwe02[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeightForPwe001 = clarityWeightDictForPwe001.ContainsKey(clarity) ? clarityWeightDictForPwe001[clarity] : 0.000;
                                sheet.Cells[filterRow + 120, 4].Value2 = Math.Round(clarityWeightForPwe001, 3);

                                // Get the craft weight for this clarity
                                double clarityWeightForPwe002 = clarityWeightDictForPwe002.ContainsKey(clarity) ? clarityWeightDictForPwe002[clarity] : 0.000;
                                sheet.Cells[filterRow + 120, 6].Value2 = Math.Round(clarityWeightForPwe002, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDividedForPwe001 = clarityCountsse02[clarity] / clarityWeightDictForPwe001[clarity];
                                sheet.Cells[filterRow + 120, 8].Value2 = Math.Round(clarityCountDividedForPwe001, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCountForPwe001 = clarityCountsse02[clarity] / clarityPwwe02[clarity];
                                sheet.Cells[filterRow + 120, 9].Value2 = Math.Round(clarityPolishWtCountForPwe001, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCrForPwe001 = clarityPwwe02[clarity] / clarityWeightDictForPwe002[clarity] * 100;
                                sheet.Cells[filterRow + 120, 10].Value2 = percentPwCrForPwe001.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCrForPwe001 = clarityPwwe02[clarity] / clarityWeightDictForPwe001[clarity] * 100;
                                sheet.Cells[filterRow + 120, 11].Value2 = percentRoCrForPwe001.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 120, 12].Value2 = clarityDolarDictForPwe001[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolarForPwe001 = clarityDolarDictForPwe001[clarity] / clarityWeightDictForPwe001[clarity];
                                sheet.Cells[filterRow + 120, 13].Value2 = Math.Round(clarityRoughCrtDolarForPwe001, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolarForPwe001 = clarityWeightDictForPwe001[clarity] / clarityPwwe02[clarity];
                                sheet.Cells[filterRow + 120, 14].Value2 = Math.Round(clarityPolishCrtDolarForPwe001, 2);

                                serialNumberForPwe02++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRangeForPwe02 = sheet.Range[sheet.Cells[filterRow - serialNumberForPwe02 + 120, 1], sheet.Cells[filterRow + 120, 14]];
                        tableRangeForPwe02.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRangeForPwe02.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRangeForPwe02.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 120, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 120, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeightForPwe001 = clarityWeightDictForPwe001.Values.Sum();
                        sheet.Cells[filterRow + 120, 4].Value2 = Math.Round(totalRoughtWeightForPwe001, 2);

                        sheet.Cells[filterRow + 120, 5].Value2 = totalCountForPwe001;

                        double partWTTotalForPwe001 = clarityWeightDictForPwe002.Values.Sum();
                        sheet.Cells[filterRow + 120, 6].Value2 = Math.Round(partWTTotalForPwe001, 2);

                        sheet.Cells[filterRow + 120, 7].Value2 = Math.Round(totalPwForPwe001, 3);

                        double totalSizeForPwe001 = (totalCountForPwe001 / totalRoughtWeightForPwe001);
                        sheet.Cells[filterRow + 120, 8].Value2 = totalSizeForPwe001.ToString("0.00");

                        double polishSizeForPwe001 = (totalCountForPwe001 / totalPwForPwe001);
                        sheet.Cells[filterRow + 120, 9].Value2 = polishSizeForPwe001.ToString("0.00");

                        double crPwPercentageForPwe001 = (totalPwForPwe001 / partWTTotalForPwe001) * 100;
                        sheet.Cells[filterRow + 120, 10].Value2 = crPwPercentageForPwe001.ToString("0.00") + "%";

                        double pwPercentageForPwm001 = (totalPwForPwe001 / totalRoughtWeightForPwe001) * 100;
                        sheet.Cells[filterRow + 120, 11].Value2 = pwPercentageForPwm001.ToString("0.00") + "%";

                        double dolarTotalForPwe001 = clarityDolarDictForPwe001.Values.Sum();
                        sheet.Cells[filterRow + 120, 12].Value2 = Math.Round(dolarTotalForPwe001, 2);

                        double valueRoughForPwe001 = (dolarTotalForPwe001 / totalRoughtWeightForPwe001);
                        sheet.Cells[filterRow + 120, 13].Value2 = valueRoughForPwe001.ToString("0.00");

                        double valuePolishForPwe001 = (dolarTotalForPwe001 / totalPwForPwe001);
                        sheet.Cells[filterRow + 120, 14].Value2 = valuePolishForPwe001.ToString("0.00");


                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeeForPwe02 = sheet.Range[sheet.Cells[filterRow + 120, 1], sheet.Cells[filterRow + 120, 14]];
                        //headerRangee1.Font.Bold = true;
                        headerRangeeForPwe02.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRangeeForPwe02.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion

                    #region 3rd Sieve EMERALD 4STEP Method

                    if (checkBoxRunCode013.Checked)
                    {
                        sheet.Cells[filterRow + 123, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 123, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 123, 5].Value2 = "Polish Weight";

                        string updatedValueForPw3 = textBox20.Text;
                        sheet.Cells[filterRow + 124, 3].Value2 = updatedValueForPw3;

                        string updatedValueForPw03 = txtPwRange13.Text;
                        sheet.Cells[filterRow + 124, 5].Value2 = updatedValueForPw03;

                        //sheet.Cells[filterRow + 10, 5].Value2 = "0.9 - 1.249";
                        txtPwRange13.Text = (string)sheet.Cells[filterRow + 124, 5].Value2;

                        sheet.Cells[filterRow + 124, 1].Value2 = "EMERALD 4STEP";

                        //sheet.Cells[filterRow + 10, 3].Value2 = "+000 -2";
                        textBox20.Text = (string)sheet.Cells[filterRow + 124, 3].Value2;

                        sheet.Cells[filterRow + 125, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 125, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 125, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 125, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 125, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 125, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 125, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 125, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 125, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 125, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 125, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 125, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 125, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 125, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 125, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeForPwe03 = sheet.Range[sheet.Cells[filterRow + 125, 1], sheet.Cells[filterRow + 125, 14]];
                        //headerRange1.Font.Bold = true;
                        headerRangeForPwe03.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the width range from the txtPwRange TextBox
                        string pwRangeText13 = txtPwRange13.Text;
                        string[] pwRangeParts13 = pwRangeText13.Split('-');
                        double minPwe03 = 0.1;  // Default minimum pw
                        double maxPwe03 = 0.2;  // Default maximum pw

                        // Parse the pw range values if the input is valid
                        if (pwRangeParts13.Length == 2 && double.TryParse(pwRangeParts13[0], out double parsedMinPw13) && double.TryParse(pwRangeParts13[1], out double parsedMaxPw13))
                        {
                            minPwe03 = parsedMinPw13;
                            maxPwe03 = parsedMaxPw13;
                        }

                        Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>> shapesEDictionary03 =
                            new Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>>();

                        Dictionary<string, int> clarityCountsse03 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPwwe03 = new Dictionary<string, double>();

                        string prevStoneShName03 = "";
                        // Loop through rows in excel sheet again to count for each stone name
                        for (int t = 2; t <= rowCount; t++)
                        {
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStoneShName03;
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            double polishWeight = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "emerald 4step" && double.TryParse(polishWeight.ToString(), out double pwValue))
                            {
                                if (pwValue >= minPwe03 && pwValue <= maxPwe03 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCountsse03.ContainsKey(clarity))
                                        {
                                            clarityCountsse03[clarity]++;
                                            clarityPwwe03[clarity] += polishWeight;
                                        }
                                        else
                                        {
                                            clarityCountsse03.Add(clarity, 1);
                                            clarityPwwe03.Add(clarity, polishWeight);
                                        }

                                        if (shapesEDictionary03.ContainsKey(stoneName))
                                        {
                                            shapesEDictionary03[stoneName].Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)> shapes = new List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>();
                                            shapes.Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                            shapesEDictionary03.Add(stoneName, shapes);
                                        }
                                    }
                                }
                            }
                            prevStoneShName03 = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING THIRD SIEVE FOR SHAPE 'EMERALD 4STEP' CHECKING IF THE PW FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        Dictionary<string, double> clarityWeightDictForPwe0001 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> pearParts = partWts.Where(p => p.Item3.ToLower() == "emerald 4step").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForPwe0001 = pearParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item4 >= minPwe03 && p.Item4 <= maxPwe03)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDictForPwe0001.ContainsKey(clarity))
                                {
                                    clarityWeightDictForPwe0001[clarity] += clarityWeightForPwe0001;
                                }
                                else
                                {
                                    clarityWeightDictForPwe0001[clarity] = clarityWeightForPwe0001;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDictForPwe0002 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesEDictionary03.Values.SelectMany(shapes => shapes))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDictForPwe0002.ContainsKey(clarity))
                            {
                                clarityWeightDictForPwe0002[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDictForPwe0002[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDictForPwe0001 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesEDictionary03.Values.SelectMany(shapes => shapes))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDictForPwe0001.ContainsKey(clarity))
                                {
                                    clarityDolarDictForPwe0001[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDictForPwe0001[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCountForPwe0001 = clarityCountsse03.Values.Sum();
                        double totalPwForPwe0001 = clarityPwwe03.Values.Sum();

                        int serialNumberForPwe03 = 1;

                        int totalClarityFiltersForPwe03 = clarityValues.Count();
                        int currentClarityFilterForPwe03 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilterForPwe03 / totalClarityFiltersForPwe03 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING THIRD 'EMERALD 4STEP' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCountsse03.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 126, 1].Value2 = serialNumberForPwe03;
                                sheet.Cells[filterRow + 126, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 126, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPwForPwe3 = clarityPwwe03[clarity] / totalPwForPwe0001 * 100.0;
                                sheet.Cells[filterRow + 126, 3].Value2 = percentPwForPwe3.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 126, 5].Value2 = clarityCountsse03[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 126, 7].Value2 = clarityPwwe03[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeightForPwe0001 = clarityWeightDictForPwe0001.ContainsKey(clarity) ? clarityWeightDictForPwe0001[clarity] : 0.000;
                                sheet.Cells[filterRow + 126, 4].Value2 = Math.Round(clarityWeightForPwe0001, 3);

                                // Get the craft weight for this clarity
                                double clarityWeightForPwe0002 = clarityWeightDictForPwe0002.ContainsKey(clarity) ? clarityWeightDictForPwe0002[clarity] : 0.000;
                                sheet.Cells[filterRow + 126, 6].Value2 = Math.Round(clarityWeightForPwe0002, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDividedForPwe0001 = clarityCountsse03[clarity] / clarityWeightDictForPwe0001[clarity];
                                sheet.Cells[filterRow + 126, 8].Value2 = Math.Round(clarityCountDividedForPwe0001, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCountForPwe0001 = clarityCountsse03[clarity] / clarityPwwe03[clarity];
                                sheet.Cells[filterRow + 126, 9].Value2 = Math.Round(clarityPolishWtCountForPwe0001, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCrForPwe0001 = clarityPwwe03[clarity] / clarityWeightDictForPwe0002[clarity] * 100;
                                sheet.Cells[filterRow + 126, 10].Value2 = percentPwCrForPwe0001.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCrForPwe0001 = clarityPwwe03[clarity] / clarityWeightDictForPwe0001[clarity] * 100;
                                sheet.Cells[filterRow + 126, 11].Value2 = percentRoCrForPwe0001.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 126, 12].Value2 = clarityDolarDictForPwe0001[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolarForPwe0001 = clarityDolarDictForPwe0001[clarity] / clarityWeightDictForPwe0001[clarity];
                                sheet.Cells[filterRow + 126, 13].Value2 = Math.Round(clarityRoughCrtDolarForPwe0001, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolarForPwe0001 = clarityWeightDictForPwe0001[clarity] / clarityPwwe03[clarity];
                                sheet.Cells[filterRow + 126, 14].Value2 = Math.Round(clarityPolishCrtDolarForPwe0001, 2);

                                serialNumberForPwe03++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRangeForPwe03 = sheet.Range[sheet.Cells[filterRow - serialNumberForPwe03 + 126, 1], sheet.Cells[filterRow + 126, 14]];
                        tableRangeForPwe03.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRangeForPwe03.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRangeForPwe03.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 126, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 126, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeightForPwe0001 = clarityWeightDictForPwe0001.Values.Sum();
                        sheet.Cells[filterRow + 126, 4].Value2 = Math.Round(totalRoughtWeightForPwe0001, 2);

                        sheet.Cells[filterRow + 126, 5].Value2 = totalCountForPwe0001;

                        double partWTTotalForPwe0001 = clarityWeightDictForPwe0002.Values.Sum();
                        sheet.Cells[filterRow + 126, 6].Value2 = Math.Round(partWTTotalForPwe0001, 2);

                        sheet.Cells[filterRow + 126, 7].Value2 = Math.Round(totalPwForPwe0001, 3);

                        double totalSizeForPwe0001 = (totalCountForPwe0001 / totalRoughtWeightForPwe0001);
                        sheet.Cells[filterRow + 126, 8].Value2 = totalSizeForPwe0001.ToString("0.00");

                        double polishSizeForPwe0001 = (totalCountForPwe0001 / totalPwForPwe0001);
                        sheet.Cells[filterRow + 126, 9].Value2 = polishSizeForPwe0001.ToString("0.00");

                        double crPwPercentageForPwe0001 = (totalPwForPwe0001 / partWTTotalForPwe0001) * 100;
                        sheet.Cells[filterRow + 126, 10].Value2 = crPwPercentageForPwe0001.ToString("0.00") + "%";

                        double pwPercentageForPwe0001 = (totalPwForPwe0001 / totalRoughtWeightForPwe0001) * 100;
                        sheet.Cells[filterRow + 126, 11].Value2 = pwPercentageForPwe0001.ToString("0.00") + "%";

                        double dolarTotalForPwe0001 = clarityDolarDictForPwe0001.Values.Sum();
                        sheet.Cells[filterRow + 126, 12].Value2 = Math.Round(dolarTotalForPwe0001, 2);

                        double valueRoughForPwe0001 = (dolarTotalForPwe0001 / totalRoughtWeightForPwe0001);
                        sheet.Cells[filterRow + 126, 13].Value2 = valueRoughForPwe0001.ToString("0.00");

                        double valuePolishForPwe0001 = (dolarTotalForPwe0001 / totalPwForPwe0001);
                        sheet.Cells[filterRow + 126, 14].Value2 = valuePolishForPwe0001.ToString("0.00");


                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeeForPwe03 = sheet.Range[sheet.Cells[filterRow + 126, 1], sheet.Cells[filterRow + 126, 14]];
                        //headerRangee1.Font.Bold = true;
                        headerRangeeForPwe03.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRangeeForPwe03.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion

                    #region 4th Sieve EMERALD 4STEP Method

                    if (checkBoxRunCode014.Checked)
                    {
                        sheet.Cells[filterRow + 129, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 129, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 129, 5].Value2 = "Polish Weight";

                        string updatedValueForPw4 = textBox21.Text;
                        sheet.Cells[filterRow + 130, 3].Value2 = updatedValueForPw4;

                        string updatedValueForPw04 = txtPwRange14.Text;
                        sheet.Cells[filterRow + 130, 5].Value2 = updatedValueForPw04;

                        //sheet.Cells[filterRow + 10, 5].Value2 = "0.9 - 1.249";
                        txtPwRange14.Text = (string)sheet.Cells[filterRow + 130, 5].Value2;

                        sheet.Cells[filterRow + 130, 1].Value2 = "EMERALD 4STEP";

                        //sheet.Cells[filterRow + 10, 3].Value2 = "+000 -2";
                        textBox21.Text = (string)sheet.Cells[filterRow + 130, 3].Value2;

                        sheet.Cells[filterRow + 131, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 131, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 131, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 131, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 131, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 131, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 131, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 131, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 131, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 131, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 131, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 131, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 131, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 131, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 131, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeForPwe04 = sheet.Range[sheet.Cells[filterRow + 131, 1], sheet.Cells[filterRow + 131, 14]];
                        //headerRange1.Font.Bold = true;
                        headerRangeForPwe04.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the Pw range from the txtPwRange04 TextBox
                        string pwRangeText14 = txtPwRange14.Text;
                        string[] pwRangeParts14 = pwRangeText14.Split('-');
                        double minPwe04 = 0.3;  // Default minimum pw
                        double maxPwe04 = 0.4;  // Default maximum pw

                        // Parse the pw range values if the input is valid
                        if (pwRangeParts14.Length == 2 && double.TryParse(pwRangeParts14[0], out double parsedMinPw14) && double.TryParse(pwRangeParts14[1], out double parsedMaxPw14))
                        {
                            minPwe04 = parsedMinPw14;
                            maxPwe04 = parsedMaxPw14;
                        }

                        Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>> shapesEDictionary04 =
                            new Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>>();

                        Dictionary<string, int> clarityCountsse04 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPwwe04 = new Dictionary<string, double>();

                        string prevStoneShName04 = "";
                        // Loop through rows in excel sheet again to count for each stone name
                        for (int t = 2; t <= rowCount; t++)
                        {
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStoneShName04;
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            double polishWeight = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "emerald 4step" && double.TryParse(polishWeight.ToString(), out double pwValue))
                            {
                                if (pwValue >= minPwe04 && pwValue <= maxPwe04 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCountsse04.ContainsKey(clarity))
                                        {
                                            clarityCountsse04[clarity]++;
                                            clarityPwwe04[clarity] += polishWeight;
                                        }
                                        else
                                        {
                                            clarityCountsse04.Add(clarity, 1);
                                            clarityPwwe04.Add(clarity, polishWeight);
                                        }

                                        if (shapesEDictionary04.ContainsKey(stoneName))
                                        {
                                            shapesEDictionary04[stoneName].Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)> shapes = new List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>();
                                            shapes.Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                            shapesEDictionary04.Add(stoneName, shapes);
                                        }
                                    }
                                }
                            }
                            prevStoneShName04 = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING FOURTH SIEVE FOR SHAPE 'EMERALD 4STEP' CHECKING IF THE PW FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        Dictionary<string, double> clarityWeightDictForPwe00001 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> pearParts = partWts.Where(p => p.Item3.ToLower() == "emerald 4step").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForPwe00001 = pearParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item4 >= minPwe04 && p.Item4 <= maxPwe04)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDictForPwe00001.ContainsKey(clarity))
                                {
                                    clarityWeightDictForPwe00001[clarity] += clarityWeightForPwe00001;
                                }
                                else
                                {
                                    clarityWeightDictForPwe00001[clarity] = clarityWeightForPwe00001;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDictForPwe00002 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesEDictionary04.Values.SelectMany(shapes => shapes))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDictForPwe00002.ContainsKey(clarity))
                            {
                                clarityWeightDictForPwe00002[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDictForPwe00002[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDictForPwe00001 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesEDictionary04.Values.SelectMany(shapes => shapes))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDictForPwe00001.ContainsKey(clarity))
                                {
                                    clarityDolarDictForPwe00001[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDictForPwe00001[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCountForPwe00001 = clarityCountsse04.Values.Sum();
                        double totalPwForPwe00001 = clarityPwwe04.Values.Sum();

                        int serialNumberForPwe04 = 1;

                        int totalClarityFiltersForPwe04 = clarityValues.Count();
                        int currentClarityFilterForPwe04 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilterForPwe04 / totalClarityFiltersForPwe04 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING FOURTH 'EMERALD 4 STEP' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCountsse04.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 132, 1].Value2 = serialNumberForPwe04;
                                sheet.Cells[filterRow + 132, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 132, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPwForPwe4 = clarityPwwe04[clarity] / totalPwForPwe00001 * 100.0;
                                sheet.Cells[filterRow + 132, 3].Value2 = percentPwForPwe4.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 132, 5].Value2 = clarityCountsse04[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 132, 7].Value2 = clarityPwwe04[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeightForPwe00001 = clarityWeightDictForPwe00001.ContainsKey(clarity) ? clarityWeightDictForPwe00001[clarity] : 0.000;
                                sheet.Cells[filterRow + 132, 4].Value2 = Math.Round(clarityWeightForPwe00001, 3);

                                // Get the craft weight for this clarity
                                double clarityWeightForPwe00002 = clarityWeightDictForPwe00002.ContainsKey(clarity) ? clarityWeightDictForPwe00002[clarity] : 0.000;
                                sheet.Cells[filterRow + 132, 6].Value2 = Math.Round(clarityWeightForPwe00002, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDividedForPwe00001 = clarityCountsse04[clarity] / clarityWeightDictForPwe00001[clarity];
                                sheet.Cells[filterRow + 132, 8].Value2 = Math.Round(clarityCountDividedForPwe00001, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCountForPwe00001 = clarityCountsse04[clarity] / clarityPwwe04[clarity];
                                sheet.Cells[filterRow + 132, 9].Value2 = Math.Round(clarityPolishWtCountForPwe00001, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCrForPwe00001 = clarityPwwe04[clarity] / clarityWeightDictForPwe00002[clarity] * 100;
                                sheet.Cells[filterRow + 132, 10].Value2 = percentPwCrForPwe00001.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCrForPwe00001 = clarityPwwe04[clarity] / clarityWeightDictForPwe00001[clarity] * 100;
                                sheet.Cells[filterRow + 132, 11].Value2 = percentRoCrForPwe00001.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 132, 12].Value2 = clarityDolarDictForPwe00001[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolarForPwe00001 = clarityDolarDictForPwe00001[clarity] / clarityWeightDictForPwe00001[clarity];
                                sheet.Cells[filterRow + 132, 13].Value2 = Math.Round(clarityRoughCrtDolarForPwe00001, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolarForPwe00001 = clarityWeightDictForPwe00001[clarity] / clarityPwwe04[clarity];
                                sheet.Cells[filterRow + 132, 14].Value2 = Math.Round(clarityPolishCrtDolarForPwe00001, 2);

                                serialNumberForPwe04++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRangeForPwe04 = sheet.Range[sheet.Cells[filterRow - serialNumberForPwe04 + 132, 1], sheet.Cells[filterRow + 132, 14]];
                        tableRangeForPwe04.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRangeForPwe04.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRangeForPwe04.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 132, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 132, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeightForPwe00001 = clarityWeightDictForPwe00001.Values.Sum();
                        sheet.Cells[filterRow + 132, 4].Value2 = Math.Round(totalRoughtWeightForPwe00001, 2);

                        sheet.Cells[filterRow + 132, 5].Value2 = totalCountForPwe00001;

                        double partWTTotalForPwe00001 = clarityWeightDictForPwe00002.Values.Sum();
                        sheet.Cells[filterRow + 132, 6].Value2 = Math.Round(partWTTotalForPwe00001, 2);

                        sheet.Cells[filterRow + 132, 7].Value2 = Math.Round(totalPwForPwe00001, 3);

                        double totalSizeForPwe00001 = (totalCountForPwe00001 / totalRoughtWeightForPwe00001);
                        sheet.Cells[filterRow + 132, 8].Value2 = totalSizeForPwe00001.ToString("0.00");

                        double polishSizeForPwe00001 = (totalCountForPwe00001 / totalPwForPwe00001);
                        sheet.Cells[filterRow + 132, 9].Value2 = polishSizeForPwe00001.ToString("0.00");

                        double crPwPercentageForPwe00001 = (totalPwForPwe00001 / partWTTotalForPwe00001) * 100;
                        sheet.Cells[filterRow + 132, 10].Value2 = crPwPercentageForPwe00001.ToString("0.00") + "%";

                        double pwPercentageForPwe00001 = (totalPwForPwe00001 / totalRoughtWeightForPwe00001) * 100;
                        sheet.Cells[filterRow + 132, 11].Value2 = pwPercentageForPwe00001.ToString("0.00") + "%";

                        double dolarTotalForPwe00001 = clarityDolarDictForPwe00001.Values.Sum();
                        sheet.Cells[filterRow + 132, 12].Value2 = Math.Round(dolarTotalForPwe00001, 2);

                        double valueRoughForPwe00001 = (dolarTotalForPwe00001 / totalRoughtWeightForPwe00001);
                        sheet.Cells[filterRow + 132, 13].Value2 = valueRoughForPwe00001.ToString("0.00");

                        double valuePolishForPwe00001 = (dolarTotalForPwe00001 / totalPwForPwe00001);
                        sheet.Cells[filterRow + 132, 14].Value2 = valuePolishForPwe00001.ToString("0.00");


                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeeForPwe04 = sheet.Range[sheet.Cells[filterRow + 132, 1], sheet.Cells[filterRow + 132, 14]];
                        //headerRangee1.Font.Bold = true;
                        headerRangeeForPwe04.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRangeeForPwe04.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion

                    #region 5th Sieve EMERALD 4STEP Method

                    if (checkBoxRunCode015.Checked)
                    {
                        sheet.Cells[filterRow + 135, 1].Value2 = "SHAPE";
                        sheet.Cells[filterRow + 135, 3].Value2 = "Sieve";
                        sheet.Cells[filterRow + 135, 5].Value2 = "Polish Weight";

                        string updatedValueForPw5 = textBox22.Text;
                        sheet.Cells[filterRow + 136, 3].Value2 = updatedValueForPw5;

                        string updatedValueForPw05 = txtPwRange15.Text;
                        sheet.Cells[filterRow + 136, 5].Value2 = updatedValueForPw05;

                        //sheet.Cells[filterRow + 10, 5].Value2 = "0.9 - 1.249";
                        txtPwRange15.Text = (string)sheet.Cells[filterRow + 136, 5].Value2;

                        sheet.Cells[filterRow + 136, 1].Value2 = "EMERALD 4STEP";

                        //sheet.Cells[filterRow + 10, 3].Value2 = "+000 -2";
                        textBox22.Text = (string)sheet.Cells[filterRow + 136, 3].Value2;

                        sheet.Cells[filterRow + 137, 1].Value2 = "Sr No.";
                        sheet.Cells[filterRow + 137, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center
                        sheet.Cells[filterRow + 137, 2].Value2 = "Purity";
                        sheet.Cells[filterRow + 137, 3].Value2 = "%of PolishWeight";
                        sheet.Cells[filterRow + 137, 4].Value2 = "Rough CRT";
                        sheet.Cells[filterRow + 137, 5].Value2 = "PolishPCs";
                        sheet.Cells[filterRow + 137, 6].Value2 = "Craft Weight";
                        sheet.Cells[filterRow + 137, 7].Value2 = "PolishWeight";
                        sheet.Cells[filterRow + 137, 8].Value2 = "Rough Size";
                        sheet.Cells[filterRow + 137, 9].Value2 = "Polish Size";
                        sheet.Cells[filterRow + 137, 10].Value2 = "Craft To Polish %";
                        sheet.Cells[filterRow + 137, 11].Value2 = "Rough To Polish %";
                        sheet.Cells[filterRow + 137, 12].Value2 = "Polish Dollar";
                        sheet.Cells[filterRow + 137, 13].Value2 = "Value/Rough Cts";
                        sheet.Cells[filterRow + 137, 14].Value2 = "Value/Polish Cts";

                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeForPwe05 = sheet.Range[sheet.Cells[filterRow + 137, 1], sheet.Cells[filterRow + 137, 14]];
                        //headerRange1.Font.Bold = true;
                        headerRangeForPwe05.Interior.Color = System.Drawing.Color.LightGray;

                        // Retrieve the Pw range from the txtPwRange04 TextBox
                        string pwRangeText15 = txtPwRange15.Text;
                        string[] pwRangeParts15 = pwRangeText15.Split('-');
                        double minPwe05 = 0.5;  // Default minimum pw
                        double maxPwe05 = 0.6;  // Default maximum pw

                        // Parse the pw range values if the input is valid
                        if (pwRangeParts15.Length == 2 && double.TryParse(pwRangeParts15[0], out double parsedMinPw15) && double.TryParse(pwRangeParts15[1], out double parsedMaxPw15))
                        {
                            minPwe05 = parsedMinPw15;
                            maxPwe05 = parsedMaxPw15;
                        }

                        Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>> shapesEDictionary05 =
                            new Dictionary<string, List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>>();

                        Dictionary<string, int> clarityCountsse05 = new Dictionary<string, int>();

                        Dictionary<string, double> clarityPwwe05 = new Dictionary<string, double>();

                        string prevStoneShName05 = "";
                        // Loop through rows in excel sheet again to count for each stone name
                        for (int t = 2; t <= rowCount; t++)
                        {
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStoneShName05;
                            string shape = (range.Cells[t, 9].Value2 != null) ? range.Cells[t, 9].Value2.ToString() : "";
                            string partWTString = (range.Cells[t, 10].Value2 != null) ? range.Cells[t, 10].Value2.ToString() : "0.000";
                            double polishWeight = (range.Cells[t, 11].Value2 != null) ? range.Cells[t, 11].Value2 : 0.000;
                            string clarity = (range.Cells[t, 16].Value2 != null) ? range.Cells[t, 16].Value2.ToString() : "";
                            string PoDolar = (range.Cells[t, 18].Value2 != null) ? range.Cells[t, 18].Value2.ToString() : "0.000";

                            if (!string.IsNullOrEmpty(stoneName) && !string.IsNullOrEmpty(shape) && shape.ToLower() == "emerald 4step" && double.TryParse(polishWeight.ToString(), out double pwValue))
                            {
                                if (pwValue >= minPwe05 && pwValue <= maxPwe05 && clarityValues.Contains(clarity))
                                {
                                    // Check if the stone name has the specified clarity
                                    if (clarityDict.ContainsKey(stoneName))
                                    {
                                        // Add the clarity and polish weight to the respective dictionaries
                                        if (clarityCountsse05.ContainsKey(clarity))
                                        {
                                            clarityCountsse05[clarity]++;
                                            clarityPwwe05[clarity] += polishWeight;
                                        }
                                        else
                                        {
                                            clarityCountsse05.Add(clarity, 1);
                                            clarityPwwe05.Add(clarity, polishWeight);
                                        }

                                        if (shapesEDictionary05.ContainsKey(stoneName))
                                        {
                                            shapesEDictionary05[stoneName].Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                        }
                                        else
                                        {
                                            List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)> shapes = new List<(string shape, double polishWeight, double partWT, string clarity, string PoDolar)>();
                                            shapes.Add((shape, pwValue, double.Parse(partWTString), clarity, PoDolar));
                                            shapesEDictionary05.Add(stoneName, shapes);
                                        }
                                    }
                                }
                            }
                            prevStoneShName05 = stoneName;

                            // Update the progress bar and label text
                            int progressPercentage = (int)Math.Round((double)t / totalRows * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "CALCULATING FIFTH SIEVE FOR SHAPE 'EMERALD 4STEP' CHECKING IF THE PW FALLS WITHIN THE SPECIFIED RANGE";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));
                        }

                        Dictionary<string, double> clarityWeightDictForPwe000001 = new Dictionary<string, double>();

                        foreach (string stoneName in partWtForWdDict.Keys)
                        {
                            // Get the list of part weights for the current stone name
                            List<Tuple<double, double, string, double>> partWts = partWtForWdDict[stoneName];

                            // Filter out only the parts with a shape of "Pear"
                            List<Tuple<double, double, string, double>> pearParts = partWts.Where(p => p.Item3.ToLower() == "emerald 4step").ToList();

                            // Get the list of clarities for the current stone name
                            List<string> clarities = clarityDict.ContainsKey(stoneName) ? clarityDict[stoneName] : new List<string>();

                            // Loop through each unique clarity for the current stone name
                            foreach (string clarity in clarities.Distinct())
                            {
                                // Sum the weights of the parts with the current clarity that fall within the specified polish weight (pw) range
                                double clarityWeightForPwe000001 = pearParts
                                    .Where(p => clarities[partWts.IndexOf(p)] == clarity && p.Item4 >= minPwe05 && p.Item4 <= maxPwe05)
                                    .Sum(p => p.Item1);

                                // Add the clarity weight to the clarityWeightDictForPw01
                                if (clarityWeightDictForPwe000001.ContainsKey(clarity))
                                {
                                    clarityWeightDictForPwe000001[clarity] += clarityWeightForPwe000001;
                                }
                                else
                                {
                                    clarityWeightDictForPwe000001[clarity] = clarityWeightForPwe000001;
                                }
                            }
                        }

                        Dictionary<string, double> clarityWeightDictForPwe000002 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesEDictionary05.Values.SelectMany(shapes => shapes))
                        {
                            double partWT = stoneData.partWT;
                            string clarity = stoneData.clarity;

                            // Sum the partWT values clarity-wise
                            if (clarityWeightDictForPwe000002.ContainsKey(clarity))
                            {
                                clarityWeightDictForPwe000002[clarity] += partWT;
                            }
                            else
                            {
                                clarityWeightDictForPwe000002[clarity] = partWT;
                            }
                        }

                        Dictionary<string, double> clarityDolarDictForPwe000001 = new Dictionary<string, double>();

                        // Iterate over the stone names in shapesDictionary
                        foreach (var stoneData in shapesEDictionary05.Values.SelectMany(shapes => shapes))
                        {
                            string PoDolarString = stoneData.PoDolar;
                            string clarity = stoneData.clarity;

                            // Convert PoDolarString to double
                            if (double.TryParse(PoDolarString, out double PoDolar))
                            {
                                // Sum the Dolar values clarity-wise
                                if (clarityDolarDictForPwe000001.ContainsKey(clarity))
                                {
                                    clarityDolarDictForPwe000001[clarity] += PoDolar;
                                }
                                else
                                {
                                    clarityDolarDictForPwe000001[clarity] = PoDolar;
                                }
                            }
                        }

                        // Calculate total count and total Pw
                        int totalCountForPwe000001 = clarityCountsse05.Values.Sum();
                        double totalPwForPwe000001 = clarityPwwe05.Values.Sum();

                        int serialNumberForPwe05 = 1;

                        int totalClarityFiltersForPwe05 = clarityValues.Count();
                        int currentClarityFilterForPwe05 = 0;

                        // Loop through each clarity filter and write the results to the worksheet
                        foreach (string clarity in clarityValues)
                        {
                            // Update progress
                            currentClarityFilter++;
                            int progressPercentage = (int)Math.Round((double)currentClarityFilterForPwe05 / totalClarityFiltersForPwe05 * 100);
                            progressPercentage = Math.Max(0, Math.Min(progressPercentage, 100)); // Clamp the value between 0 and 100
                            Invoke(new System.Action(() =>
                            {
                                progressBar1.Value = progressPercentage;
                                string currentMethod = "FINALIZING FIFTH 'EMERALD 4STEP' SIEVE TOTAL VALUES";
                                lblProgress.Text = string.Format("{0}: {1}%", currentMethod, progressPercentage);
                            }));

                            if (clarityCountsse05.ContainsKey(clarity))
                            {
                                // Serial number wise
                                sheet.Cells[filterRow + 138, 1].Value2 = serialNumberForPwe05;
                                sheet.Cells[filterRow + 138, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                                // Get the clarity Filter
                                sheet.Cells[filterRow + 138, 2].Value2 = clarity;

                                // Get % of polish weight
                                double percentPwForPwe5 = clarityPwwe05[clarity] / totalPwForPwe000001 * 100.0;
                                sheet.Cells[filterRow + 138, 3].Value2 = percentPwForPwe5.ToString("0.00") + "%";

                                // Write the total count of roughPcs for the clarity
                                sheet.Cells[filterRow + 138, 5].Value2 = clarityCountsse05[clarity];

                                // Write the Part Weight for the clarity
                                sheet.Cells[filterRow + 138, 7].Value2 = clarityPwwe05[clarity];

                                // Get the roughCRT weight for clarity
                                double clarityWeightForPwe000001 = clarityWeightDictForPwe000001.ContainsKey(clarity) ? clarityWeightDictForPwe000001[clarity] : 0.000;
                                sheet.Cells[filterRow + 138, 4].Value2 = Math.Round(clarityWeightForPwe000001, 3);

                                // Get the craft weight for this clarity
                                double clarityWeightForPwe000002 = clarityWeightDictForPwe000002.ContainsKey(clarity) ? clarityWeightDictForPwe000002[clarity] : 0.000;
                                sheet.Cells[filterRow + 138, 6].Value2 = Math.Round(clarityWeightForPwe000002, 3);

                                // Calculate the division of clarityCounts by clarityWeightDict
                                double clarityCountDividedForPwe000001 = clarityCountsse05[clarity] / clarityWeightDictForPwe000001[clarity];
                                sheet.Cells[filterRow + 138, 8].Value2 = Math.Round(clarityCountDividedForPwe000001, 2);

                                // Calculate the division of clarityCounts by clarityPolishWeight
                                double clarityPolishWtCountForPwe000001 = clarityCountsse05[clarity] / clarityPwwe05[clarity];
                                sheet.Cells[filterRow + 138, 9].Value2 = Math.Round(clarityPolishWtCountForPwe000001, 2);

                                // Calculate the % of polish weight with craft weight
                                double percentPwCrForPwe000001 = clarityPwwe05[clarity] / clarityWeightDictForPwe000002[clarity] * 100;
                                sheet.Cells[filterRow + 138, 10].Value2 = percentPwCrForPwe000001.ToString("0.00") + "%";

                                // Calculate the % rough weight with craft weight
                                double percentRoCrForPwe000001 = clarityPwwe05[clarity] / clarityWeightDictForPwe000001[clarity] * 100;
                                sheet.Cells[filterRow + 138, 11].Value2 = percentRoCrForPwe000001.ToString("0.00") + "%";

                                // Calculate dolar sum clarity wise
                                sheet.Cells[filterRow + 138, 12].Value2 = clarityDolarDictForPwe000001[clarity];

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityRoughCrtDolarForPwe000001 = clarityDolarDictForPwe000001[clarity] / clarityWeightDictForPwe000001[clarity];
                                sheet.Cells[filterRow + 138, 13].Value2 = Math.Round(clarityRoughCrtDolarForPwe000001, 2);

                                // Calculate division of dolar by rough crt clarity wise
                                double clarityPolishCrtDolarForPwe000001 = clarityWeightDictForPwe000001[clarity] / clarityPwwe05[clarity];
                                sheet.Cells[filterRow + 138, 14].Value2 = Math.Round(clarityPolishCrtDolarForPwe000001, 2);

                                serialNumberForPwe05++;
                                filterRow++;
                            }
                        }

                        // Apply table formatting
                        Microsoft.Office.Interop.Excel.Range tableRangeForPwe05 = sheet.Range[sheet.Cells[filterRow - serialNumberForPwe05 + 138, 1], sheet.Cells[filterRow + 138, 14]];
                        tableRangeForPwe05.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tableRangeForPwe05.Font.Size = 11;
                        //tableRange.Font.Name = "Arial";
                        tableRangeForPwe05.Columns.AutoFit();

                        // Write the total count and total Pw to the worksheet
                        sheet.Cells[filterRow + 138, 1].Value2 = "Total";
                        sheet.Cells[filterRow + 138, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // Align center

                        double totalRoughtWeightForPwe000001 = clarityWeightDictForPwe000001.Values.Sum();
                        sheet.Cells[filterRow + 138, 4].Value2 = Math.Round(totalRoughtWeightForPwe000001, 2);

                        sheet.Cells[filterRow + 138, 5].Value2 = totalCountForPwe000001;

                        double partWTTotalForPwe000001 = clarityWeightDictForPwe000002.Values.Sum();
                        sheet.Cells[filterRow + 138, 6].Value2 = Math.Round(partWTTotalForPwe000001, 2);

                        sheet.Cells[filterRow + 138, 7].Value2 = Math.Round(totalPwForPwe000001, 3);

                        double totalSizeForPwe000001 = (totalCountForPwe000001 / totalRoughtWeightForPwe000001);
                        sheet.Cells[filterRow + 138, 8].Value2 = totalSizeForPwe000001.ToString("0.00");

                        double polishSizeForPwe000001 = (totalCountForPwe000001 / totalPwForPwe000001);
                        sheet.Cells[filterRow + 138, 9].Value2 = polishSizeForPwe000001.ToString("0.00");

                        double crPwPercentageForPwe000001 = (totalPwForPwe000001 / partWTTotalForPwe000001) * 100;
                        sheet.Cells[filterRow + 138, 10].Value2 = crPwPercentageForPwe000001.ToString("0.00") + "%";

                        double pwPercentageForPwe000001 = (totalPwForPwe000001 / totalRoughtWeightForPwe000001) * 100;
                        sheet.Cells[filterRow + 138, 11].Value2 = pwPercentageForPwe000001.ToString("0.00") + "%";

                        double dolarTotalForPwe000001 = clarityDolarDictForPwe000001.Values.Sum();
                        sheet.Cells[filterRow + 138, 12].Value2 = Math.Round(dolarTotalForPwe000001, 2);

                        double valueRoughForPwe000001 = (dolarTotalForPwe000001 / totalRoughtWeightForPwe000001);
                        sheet.Cells[filterRow + 138, 13].Value2 = valueRoughForPwe000001.ToString("0.00");

                        double valuePolishForPwe000001 = (dolarTotalForPwe000001 / totalPwForPwe000001);
                        sheet.Cells[filterRow + 138, 14].Value2 = valuePolishForPwe000001.ToString("0.00");


                        // Apply formatting to header cells
                        Microsoft.Office.Interop.Excel.Range headerRangeeForPwe05 = sheet.Range[sheet.Cells[filterRow + 138, 1], sheet.Cells[filterRow + 138, 14]];
                        //headerRangee1.Font.Bold = true;
                        headerRangeeForPwe05.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        headerRangeeForPwe05.Interior.Color = System.Drawing.Color.LightGray;
                    }

                    #endregion

                    // Assuming you have a checkbox control named checkBoxValuePerRoughCT
                    if (checkBoxValuePerRoughCT.Checked)
                    {
                        // Calculate and add the ValuePerRoughCT column (between columns F and G)
                        sheet.Cells[filterRow + 137, 7].EntireColumn.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftToRight, System.Type.Missing);
                        range = sheet.UsedRange;
                        range.Cells[1, 7].Value2 = "ValuePerRoughCT";

                        // Loop through the rows and set the calculated values for the new column
                        string prevStoneNamePerCT = "";
                        for (int t = 2; t <= rowCount; t++)
                        {
                            string stoneName = (range.Cells[t, 1].Value2 != null) ? range.Cells[t, 1].Value2.ToString() : prevStoneNamePerCT;
                            double roughWeight = (range.Cells[t, 2].Value2 != null) ? range.Cells[t, 2].Value2 : 0.0;
                            double totalValue = (range.Cells[t, 6].Value2 != null) ? range.Cells[t, 6].Value2 : 0.0;

                            if (!string.IsNullOrEmpty(stoneName) && roughWeight > 0)
                            {
                                double valuePerRoughCT = totalValue / roughWeight;

                                // Set the ValuePerRoughCT in the new column (column G)
                                sheet.Cells[t, 7].Value2 = valuePerRoughCT;
                            }
                        }
                    }

                    wb.Save();

                    MessageBox.Show("RoughPCs: " + stoneNameCount +
                        "\n \nTotal RoughWeight: " + totalRoughtWeight.ToString("0.000") +
                        "\n \nTotal PartWT: " + partWTTotal.ToString("0.00") +
                        "\n \nTotal Pw: " + pwTotal.ToString("0.00") +
                        "\n \nR2P(%): " + pwPercentage.ToString("0.00") + "%" +
                        "\n \nTotal Dolar: " + dolarTotal.ToString("0.00") +
                        "\n \nValue/Rough: " + valueRough.ToString("0.00") +
                        "\n \nValue/Polish: " + valuePolish.ToString("0.00") +
                        "\n \nPolishPCs: " + partCount.ToString("0") +
                        "\n \nPolish Size: " + polishSize.ToString("0.00") +
                        "\n \nSize: " + totalSize.ToString("0.00"),
                        "SUMMARY", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                finally
                {
                    // Release COM objects in reverse order of creation
                    if (range != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                        range = null;
                    }
                    if (sheet != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                        sheet = null;
                    }
                    if (wb != null)
                    {
                        wb.Close(false); // Close without saving changes
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                        wb = null;
                    }
                    if (excel != null)
                    {
                        excel.Quit();
                        excel.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                        excel = null;
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

                DataTable dataTable = LoadExcelIntoDataTable(filePath);

                this.Invoke((MethodInvoker)delegate
                {
                    // Bind DataTable to kryptonDataGridView1
                    kryptonDataGridView1.DataSource = dataTable;
                });

            });
        }

        #endregion

        private DataTable LoadExcelIntoDataTable(string filePath)
        {
            DataTable dataTable = new DataTable();

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    if (result.Tables.Count > 0)
                    {
                        dataTable = result.Tables[0];
                    }
                }
            }

            return dataTable;
        }


        private async void btnConvertAndSum_Click(object sender, EventArgs e)
        {
            // Disable all buttons and textboxes except progressBar1
            btnConvert.Enabled = false;
            btnSumConvert.Enabled = false;
            btnConvertAndSum.Enabled = false;
            txtFolder.Enabled = false;
            txtExcelFilePath.Enabled = false;
            btnBrowse.Enabled = false;
            btnExcelBrowse.Enabled = false;
            btnSaveSettings.Enabled = false;
            kryptonHeaderGroup1.Enabled = false;
            kryptonHeaderGroup2.Enabled = false;
            kryptonHeaderGroup3.Enabled = false;
            kryptonHeaderGroup4.Enabled = false;

            try
            {
                // Call btnConvert_Click and wait for it to complete
                await btnConvert_ClickAsync(sender, e);

                // Marshal the UI update for progressBar1 back to the UI thread
                Invoke(new System.Action(() => progressBar1.Value = 0));

                // Pass the excelWorkbook and excelApp objects to btnSumConvert_ClickAsync
                await btnSumConvert_ClickAsync(sender, e);
            }
            catch (Exception ex)
            {
                // Handle any exceptions that occur during the process
                Invoke(new System.Action(() =>
                {
                    MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }));
            }
            finally
            {
                // Enable the buttons and textboxes after the task completes or an error occurs
                btnConvert.Enabled = true;
                btnSumConvert.Enabled = true;
                btnConvertAndSum.Enabled = true;
                txtFolder.Enabled = true;
                txtExcelFilePath.Enabled = true;
                btnBrowse.Enabled = true;
                btnExcelBrowse.Enabled = true;
                btnSaveSettings.Enabled = true;
                kryptonHeaderGroup1.Enabled = true;
                kryptonHeaderGroup2.Enabled = true;
                kryptonHeaderGroup3.Enabled = true;
                kryptonHeaderGroup4.Enabled = true;
            }
        }

        #region Settings for Sieves

        private void LoadSettings()
        {
            // Read the settings from the text file
            string settingsFilePath = "settings.txt";
            if (File.Exists(settingsFilePath))
            {
                string[] lines = File.ReadAllLines(settingsFilePath);
                if (lines.Length >= 51)  // Updated condition to account for the additional empty line
                {
                    textBox1.Text = lines[1];
                    textBox2.Text = lines[2];
                    textBox3.Text = lines[3];
                    textBox4.Text = lines[4];
                    textBox5.Text = lines[5];
                    textBox6.Text = lines[6];
                    textBox7.Text = lines[7];

                    // Skip the empty line
                    txtWidthRange.Text = lines[10];
                    txtWidthRange01.Text = lines[11];
                    txtWidthRange02.Text = lines[12];
                    txtWidthRange03.Text = lines[13];
                    txtWidthRange04.Text = lines[14];
                    txtWidthRange05.Text = lines[15];
                    txtWidthRange06.Text = lines[16];

                    // Skip the empty line
                    textBox8.Text = lines[19];
                    textBox9.Text = lines[20];
                    textBox10.Text = lines[21];
                    textBox11.Text = lines[22];
                    textBox12.Text = lines[23];

                    // Skip the empty line
                    txtPwRange01.Text = lines[26];
                    txtPwRange02.Text = lines[27];
                    txtPwRange03.Text = lines[28];
                    txtPwRange04.Text = lines[29];
                    txtPwRange05.Text = lines[30];

                    // Skip the empty line
                    textBox13.Text = lines[33];
                    textBox14.Text = lines[34];
                    textBox15.Text = lines[35];
                    textBox16.Text = lines[36];
                    textBox17.Text = lines[37];

                    // Skip the empty line
                    txtPwRange06.Text = lines[40];
                    txtPwRange07.Text = lines[41];
                    txtPwRange08.Text = lines[42];
                    txtPwRange09.Text = lines[43];
                    txtPwRange10.Text = lines[44];

                    // Skip the empty line
                    textBox18.Text = lines[47];
                    textBox19.Text = lines[48];
                    textBox20.Text = lines[49];
                    textBox21.Text = lines[50];
                    textBox22.Text = lines[51];

                    // Skip the empty line
                    txtPwRange11.Text = lines[54];
                    txtPwRange12.Text = lines[55];
                    txtPwRange13.Text = lines[56];
                    txtPwRange14.Text = lines[57];
                    txtPwRange15.Text = lines[58];
                }
            }
        }

        private void SaveSettings()
        {
            // Save the settings to the text file
            string settingsFilePath = "settings.txt";
            string[] lines = new string[]
            {
                "Shape: Round",
                textBox1.Text,
                textBox2.Text,
                textBox3.Text,
                textBox4.Text,
                textBox5.Text,
                textBox6.Text,
                textBox7.Text,

                "", // Empty line

                "Width Ranges(Round):",
                txtWidthRange.Text,
                txtWidthRange01.Text,
                txtWidthRange02.Text,
                txtWidthRange03.Text,
                txtWidthRange04.Text,
                txtWidthRange05.Text,
                txtWidthRange06.Text,

                "", // Empty line

                "Shape: Pear",
                textBox8.Text,
                textBox9.Text,
                textBox10.Text,
                textBox11.Text,
                textBox12.Text,

                "", // Empty line

                "PW Ranges(Pear):",
                txtPwRange01.Text,
                txtPwRange02.Text,
                txtPwRange03.Text,
                txtPwRange04.Text,
                txtPwRange05.Text,

                "", // Empty line

                "Shape: Marquise",
                textBox13.Text,
                textBox14.Text,
                textBox15.Text,
                textBox16.Text,
                textBox17.Text,

                "", // Empty line

                "PW Ranges(Marquise):",
                txtPwRange06.Text,
                txtPwRange07.Text,
                txtPwRange08.Text,
                txtPwRange09.Text,
                txtPwRange10.Text,

                "", // Empty line

                "Shape: Emerald 4Step",
                textBox18.Text,
                textBox19.Text,
                textBox20.Text,
                textBox21.Text,
                textBox22.Text,

                "", // Empty line

                "PW Ranges(Emerald 4Step):",
                txtPwRange11.Text,
                txtPwRange12.Text,
                txtPwRange13.Text,
                txtPwRange14.Text,
                txtPwRange15.Text,
            };
            File.WriteAllLines(settingsFilePath, lines);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Load the settings from the text file
            LoadSettings();

            // Check the chkSelectAll checkboxes
            chkSelectAll01.Checked = true;
            chkSelectAll02.Checked = true;
            chkSelectAll03.Checked = true;
            chkSelectAll04.Checked = true;
        }

        private void btnSaveSettings_Click(object sender, EventArgs e)
        {
            // Save the settings to the text file
            SaveSettings();
        }

        #endregion

        #region Checkboxes Selection

        private void chkSelectAll01_CheckedChanged(object sender, EventArgs e)
        {
            bool isChecked = chkSelectAll01.Checked;

            // Update the state of other checkboxes based on the checked state of chkSelectAll01
            checkBoxRunCode01.Checked = isChecked;
            checkBoxRunCode02.Checked = isChecked;
            checkBoxRunCode03.Checked = isChecked;
            checkBoxRunCode04.Checked = isChecked;
            checkBoxRunCode05.Checked = isChecked;
            checkBoxRunCode06.Checked = isChecked;
            checkBoxRunCode07.Checked = isChecked;
        }

        private void chkSelectAll02_CheckedChanged(object sender, EventArgs e)
        {
            bool isChecked = chkSelectAll02.Checked;

            // Update the state of other checkboxes based on the checked state of chkSelectAll02
            checkBoxRunCode001.Checked = isChecked;
            checkBoxRunCode002.Checked = isChecked;
            checkBoxRunCode003.Checked = isChecked;
            checkBoxRunCode004.Checked = isChecked;
            checkBoxRunCode005.Checked = isChecked;
        }

        private void chkSelectAll03_CheckedChanged(object sender, EventArgs e)
        {
            bool isChecked = chkSelectAll03.Checked;

            // Update the state of other checkboxes based on the checked state of chkSelectAll02
            checkBoxRunCode006.Checked = isChecked;
            checkBoxRunCode007.Checked = isChecked;
            checkBoxRunCode008.Checked = isChecked;
            checkBoxRunCode009.Checked = isChecked;
            checkBoxRunCode010.Checked = isChecked;
        }

        private void chkSelectAll04_CheckedChanged(object sender, EventArgs e)
        {
            bool isChecked = chkSelectAll04.Checked;

            // Update the state of other checkboxes based on the checked state of chkSelectAll02
            checkBoxRunCode011.Checked = isChecked;
            checkBoxRunCode012.Checked = isChecked;
            checkBoxRunCode013.Checked = isChecked;
            checkBoxRunCode014.Checked = isChecked;
            checkBoxRunCode015.Checked = isChecked;
        }

        #endregion

        private void btnReloadDGVSheet_Click(object sender, EventArgs e)
        {
            // Assuming filePath is the path to your Excel file
            string filePath = txtExcelFilePath.Text;

            // Check if filePath is blank
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("Please Select Excel File to Continue.", "Excel File Not Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Load Excel data into DataTable
            DataTable dataTable = LoadExcelIntoDataTable(filePath);

            // Update DataGridView on the UI thread
            this.Invoke((MethodInvoker)delegate
            {
                kryptonDataGridView1.DataSource = dataTable;
            });
        }
    }
}
