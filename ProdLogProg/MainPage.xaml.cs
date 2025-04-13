using System;
using System.Collections.Generic;
using System.IO;
using Xamarin.Forms;
using Xamarin.Essentials;

namespace ProdLogProg
{
    public partial class MainPage : ContentPage
    {
        private List<ProductionData> dataList = new List<ProductionData>();

        public MainPage()
        {
            InitializeComponent();

        }

        public class ProductionData
        {
            public string Product { get; set; }
            public string TimePerUnit { get; set; }
            public string Quantity { get; set; }
            public string MONumber { get; set; }
            public string Comments { get; set; }
            public string Result { get; set; }
        }

        private void BtnAddData_Clicked(object sender, EventArgs e)
        {
            try
            {
                var data = new ProductionData
                {
                    Product = txtProduct.Text,
                    TimePerUnit = txtTime.Text,
                    Quantity = txtQuantity.Text,
                    MONumber = txtMONumber.Text,
                    Comments = txtComments.Text,
                    Result = lblResult.Text
                };

                // Add the collected data to the list
                dataList.Add(data);

                // Provide feedback to the user
                DisplayAlert("Success", "Data added successfully.", "OK");

                txtProduct.Text = string.Empty;
                txtTime.Text = string.Empty;
                txtQuantity.Text = string.Empty;
                txtMONumber.Text = string.Empty;
                txtComments.Text = string.Empty;
                lblResult.Text = string.Empty;
            }
            catch (Exception ex)
            {
                DisplayAlert("Error", ex.Message, "OK");
            }
        }

        private void BtnCalculate_Clicked(object sender, EventArgs e)
        {
            try
            {
                double timePerUnit = Convert.ToDouble(txtTime.Text);
                int buildQuantity = Convert.ToInt32(txtQuantity.Text);
                string selectedProduct = txtProduct.Text;
                double panelsPerHour = 3600 / timePerUnit;
                double productionRunTime = buildQuantity / panelsPerHour;
                int hours = (int)productionRunTime;
                int minutes = (int)((productionRunTime - hours) * 60);

                lblResult.Text = $"Run Time: {hours} hours {minutes} mins";
            }
            catch (Exception)
            {
                lblResult.Text = "Invalid input";
            }
        }

        private void BtnClear_Clicked(object sender, EventArgs e)
        {
            txtProduct.Text = string.Empty;
            txtTime.Text = string.Empty;
            txtQuantity.Text = string.Empty;
            lblResult.Text = string.Empty;
            txtMONumber.Text = string.Empty;
            txtComments.Text = string.Empty;
            dataList.Clear();
        }

        private async void BtnSaveToExcel_Clicked(object sender, EventArgs e)
        {
            try
            {
                // Prompt the user for a file name
                string fileName = await DisplayPromptAsync("Save File", "Enter a name for the Excel file:");

                if (string.IsNullOrEmpty(fileName))
                {
                    await DisplayAlert("Cancelled", "No file name provided.", "OK");
                    return;
                }

                // Ensure the file name ends with .xlsx
                if (!fileName.EndsWith(".xlsx"))
                {
                    fileName += ".xlsx";
                }

                // Create the Excel file in memory
                using (var stream = new MemoryStream())
                using (var workbook = new ClosedXML.Excel.XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Production Data");

                    // Write headers
                    worksheet.Cell(1, 1).Value = "Product";
                    worksheet.Cell(1, 2).Value = "Time per Unit";
                    worksheet.Cell(1, 3).Value = "Quantity";
                    worksheet.Cell(1, 4).Value = "MO Number";
                    worksheet.Cell(1, 5).Value = "Comments";
                    worksheet.Cell(1, 6).Value = "Result";

                    // Append data
                    for (int i = 0; i < dataList.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = dataList[i].Product;
                        worksheet.Cell(i + 2, 2).Value = dataList[i].TimePerUnit;
                        worksheet.Cell(i + 2, 3).Value = dataList[i].Quantity;
                        worksheet.Cell(i + 2, 4).Value = dataList[i].MONumber;
                        worksheet.Cell(i + 2, 5).Value = dataList[i].Comments;
                        worksheet.Cell(i + 2, 6).Value = dataList[i].Result;
                    }

                    // Save the workbook to the memory stream
                    workbook.SaveAs(stream);

                    // Call the Android-specific code to save the file
                    DependencyService.Get<IFileService>().SaveFile(fileName, stream.ToArray());
                }

                // Inform the user
                await DisplayAlert("Success", $"File saved successfully as {fileName}", "OK");
            }
            catch (Exception ex)
            {
                await DisplayAlert("Error", $"Failed to save Excel file: {ex.Message}", "OK");
            }
        }

        private async void BtnImportData_Clicked(object sender, EventArgs e)
        {
            try
            {
                // Define custom file types for Excel files
                var excelFileType = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>>
                {
                    { DevicePlatform.iOS, new[] { "com.microsoft.excel.xlsx" } },
                    { DevicePlatform.Android, new[] { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } },
                    { DevicePlatform.UWP, new[] { ".xlsx" } }
                });

                // Use FilePicker to select the file
                var result = await FilePicker.PickAsync(new PickOptions
                {
                    PickerTitle = "Select an Excel File",
                    FileTypes = excelFileType
                });

                if (result == null)
                {
                    // User canceled the file selection
                    await DisplayAlert("Cancelled", "No file selected.", "OK");
                    return;
                }

                // Get the selected file's path
                string filePath = result.FullPath;

                // Open the workbook
                using (var workbook = new ClosedXML.Excel.XLWorkbook(filePath))
                {
                    // Access the worksheet
                    var worksheet = workbook.Worksheet("Production Data");

                    // Find the last row with data
                    var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;

                    // Loop through rows and import data
                    for (int row = 2; row <= lastRow; row++)
                    {
                        var data = new ProductionData
                        {
                            Product = worksheet.Cell(row, 1).GetValue<string>(),
                            TimePerUnit = worksheet.Cell(row, 2).GetValue<string>(),
                            Quantity = worksheet.Cell(row, 3).GetValue<string>(),
                            MONumber = worksheet.Cell(row, 4).GetValue<string>(),
                            Comments = worksheet.Cell(row, 5).GetValue<string>(),
                            Result = worksheet.Cell(row, 6).GetValue<string>()
                        };

                        // Add the imported data to the list
                        dataList.Add(data);
                    }
                }

                await DisplayAlert("Success", "Data imported successfully.", "OK");
            }
            catch (Exception ex)
            {
                await DisplayAlert("Error", $"Failed to import data: {ex.Message}", "OK");
            }
        }

        private void BtnExit_Clicked(object sender, EventArgs e)
        {
            System.Diagnostics.Process.GetCurrentProcess().CloseMainWindow();
        }
    }
}
