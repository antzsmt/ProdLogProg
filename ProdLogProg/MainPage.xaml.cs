using System;
using System.Collections.Generic;
using System.IO;
using Xamarin.Forms;
using Xamarin.Essentials;
using System.Linq;

namespace ProdLogProg
{
    public partial class MainPage : ContentPage
    {
        private List<ProductionData> dataList = new List<ProductionData>();

        private int selectedRow = -1;

        public MainPage()
        {
            InitializeComponent();

        }

        protected override void OnAppearing()
        {
            base.OnAppearing();

            // Set the focus to the "Name" Entry field when the page appears
            txtProduct.Focus();

            // Get half of the screen width
            double halfScreenWidth = Application.Current.MainPage.Width / 2;

            // Set the width for buttons
            btnCalculate.WidthRequest = halfScreenWidth;
            btnAddData.WidthRequest = halfScreenWidth;
            btnClear.WidthRequest = halfScreenWidth;
            btnSaveToExcel.WidthRequest = halfScreenWidth;
            btnImportData.WidthRequest = halfScreenWidth;
            btnExit.WidthRequest = halfScreenWidth;

            // Check if the last saved Excel file path exists
            if (Application.Current.Properties.ContainsKey("LastExcelFilePath"))
            {
                string lastFilePath = Application.Current.Properties["LastExcelFilePath"] as string;

                if (!string.IsNullOrEmpty(lastFilePath) && File.Exists(lastFilePath))
                {
                    // File exists, proceed to load
                    LoadExcelFile(lastFilePath);
                }
                else
                {
                    // File not found or path is invalid
                    DisplayAlert("Info", "No previous file found. Starting with an empty grid.", "OK");
                }
            }
            else
            {
                // No file path was stored
                DisplayAlert("Info", "No previous file found. Starting with an empty grid.", "OK");
            }
        }

        private void OnRowTapped(int rowIndex)
        {
            if (selectedRow == rowIndex)
            {
                // If the tapped row is already selected, unselect it
                foreach (var child in DataGrid.Children)
                {
                    if (Grid.GetRow(child) == rowIndex)
                    {
                        (child as Label).BackgroundColor = Color.Transparent; // Remove highlight
                    }
                }

                selectedRow = -1; // Clear the selection
            }
            else
            {
                // Select the new row
                selectedRow = rowIndex;

                // Highlight the selected row visually
                foreach (var child in DataGrid.Children)
                {
                    if (Grid.GetRow(child) == rowIndex)
                    {
                        (child as Label).BackgroundColor = Color.LightGray; // Highlight the row
                    }
                    else
                    {
                        (child as Label).BackgroundColor = Color.Transparent; // Remove highlight from other rows
                    }
                }
            }
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

                // Add data to the list
                dataList.Add(data);

                // Add a new row to the grid dynamically
                int newRow = DataGrid.RowDefinitions.Count;
                DataGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                // Create a TapGestureRecognizer for row selection
                var tapGestureRecognizer = new TapGestureRecognizer();
                tapGestureRecognizer.Tapped += (s, args) => OnRowTapped(newRow); // Pass the correct row index

                // Create labels for each column and attach gesture recognizers
                var productLabel = new Label
                {
                    Text = data.Product,
                    HorizontalTextAlignment = TextAlignment.Center,
                    VerticalTextAlignment = TextAlignment.Center,
                    TextColor = Color.Black,
                    FontSize = 14
                };
                productLabel.GestureRecognizers.Add(tapGestureRecognizer);

                var timeLabel = new Label
                {
                    Text = data.TimePerUnit,
                    HorizontalTextAlignment = TextAlignment.Center,
                    VerticalTextAlignment = TextAlignment.Center,
                    TextColor = Color.Black,
                    FontSize = 14
                };
                timeLabel.GestureRecognizers.Add(tapGestureRecognizer);

                var quantityLabel = new Label
                {
                    Text = data.Quantity,
                    HorizontalTextAlignment = TextAlignment.Center,
                    VerticalTextAlignment = TextAlignment.Center,
                    TextColor = Color.Black,
                    FontSize = 14
                };
                quantityLabel.GestureRecognizers.Add(tapGestureRecognizer);

                var monumberLabel = new Label
                {
                    Text = data.MONumber,
                    HorizontalTextAlignment = TextAlignment.Center,
                    VerticalTextAlignment = TextAlignment.Center,
                    TextColor = Color.Black,
                    FontSize = 14
                };
                monumberLabel.GestureRecognizers.Add(tapGestureRecognizer);

                var commentsLabel = new Label
                {
                    Text = data.Comments,
                    HorizontalTextAlignment = TextAlignment.Center,
                    VerticalTextAlignment = TextAlignment.Center,
                    TextColor = Color.Black,
                    FontSize = 14
                };
                commentsLabel.GestureRecognizers.Add(tapGestureRecognizer);

                var resultLabel = new Label
                {
                    Text = data.Result,
                    HorizontalTextAlignment = TextAlignment.Center,
                    VerticalTextAlignment = TextAlignment.Center,
                    TextColor = Color.Black,
                    FontSize = 14
                };
                resultLabel.GestureRecognizers.Add(tapGestureRecognizer);

                // Add labels to the grid
                DataGrid.Children.Add(productLabel, 0, newRow);
                DataGrid.Children.Add(timeLabel, 1, newRow);
                DataGrid.Children.Add(quantityLabel, 2, newRow);
                DataGrid.Children.Add(monumberLabel, 3, newRow);
                DataGrid.Children.Add(commentsLabel, 4, newRow);
                DataGrid.Children.Add(resultLabel, 5, newRow);

                // Clear input fields
                txtProduct.Text = string.Empty;
                txtTime.Text = string.Empty;
                txtQuantity.Text = string.Empty;
                txtMONumber.Text = string.Empty;
                txtComments.Text = string.Empty;
                lblResult.Text = string.Empty;

                DisplayAlert("Success", "Data added successfully.", "OK");
            }
            catch (Exception ex)
            {
                DisplayAlert("Error", ex.Message, "OK");
            }
        }


        private void BtnRemoveData_Clicked(object sender, EventArgs e)
        {
            if (selectedRow == -1)
            {
                DisplayAlert("No Selection", "Please select a row to remove.", "OK");
                return;
            }

            // Remove the selected row from the data list
            dataList.RemoveAt(selectedRow - 1); // Adjust index if the grid starts from row 1

            // Remove the row from the grid
            var childrenToRemove = DataGrid.Children.Where(c => Grid.GetRow(c) == selectedRow).ToList();
            foreach (var child in childrenToRemove)
            {
                DataGrid.Children.Remove(child);
            }

            // Update row indices
            foreach (var child in DataGrid.Children)
            {
                int row = Grid.GetRow(child);
                if (row > selectedRow)
                {
                    Grid.SetRow(child, row - 1); // Shift rows up
                }
            }

            // Clear selection
            selectedRow = -1;
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

                lblResult.Text = $"R/T: {hours} hours {minutes} mins";
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

                // Check if the file already exists
                string filePath = DependencyService.Get<IFileService>().GetSavedFilePath(fileName);
                if (File.Exists(filePath))
                {
                    bool overwrite = await DisplayAlert(
                        "File Exists",
                        $"The file '{fileName}' already exists. Do you want to overwrite it?",
                        "Overwrite",
                        "Cancel"
                    );

                    if (!overwrite)
                    {
                        await DisplayAlert("Cancelled", "File not overwritten.", "OK");
                        return;
                    }
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

                    // Save the file
                    DependencyService.Get<IFileService>().SaveFile(fileName, stream.ToArray());

                    // Update the saved file path in persistent storage
                    Application.Current.Properties["LastExcelFilePath"] = filePath;
                    await Application.Current.SavePropertiesAsync();
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
                        var product = worksheet.Cell(row, 1).GetValue<string>();
                        var timePerUnit = worksheet.Cell(row, 2).GetValue<string>();
                        var quantity = worksheet.Cell(row, 3).GetValue<string>();
                        var moNumber = worksheet.Cell(row, 4).GetValue<string>();
                        var comments = worksheet.Cell(row, 5).GetValue<string>();
                        var resultCell = worksheet.Cell(row, 6).GetValue<string>();

                        // Skip empty rows
                        if (string.IsNullOrWhiteSpace(product) &&
                            string.IsNullOrWhiteSpace(timePerUnit) &&
                            string.IsNullOrWhiteSpace(quantity) &&
                            string.IsNullOrWhiteSpace(moNumber) &&
                            string.IsNullOrWhiteSpace(comments) &&
                            string.IsNullOrWhiteSpace(resultCell))
                        {
                            continue;
                        }

                        var data = new ProductionData
                        {
                            Product = product,
                            TimePerUnit = timePerUnit,
                            Quantity = quantity,
                            MONumber = moNumber,
                            Comments = comments,
                            Result = resultCell
                        };

                        // Add the imported data to the list
                        dataList.Add(data);

                        // Dynamically add data to the grid
                        int newRow = DataGrid.RowDefinitions.Count;
                        DataGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                        // Add labels with gesture recognizers
                        var tapGestureRecognizer = new TapGestureRecognizer();
                        tapGestureRecognizer.Tapped += (s, args) => OnRowTapped(newRow);

                        var productLabel = new Label
                        {
                            Text = data.Product,
                            HorizontalTextAlignment = TextAlignment.Center,
                            VerticalTextAlignment = TextAlignment.Center,
                            TextColor = Color.Black,
                            FontSize = 14
                        };
                        productLabel.GestureRecognizers.Add(tapGestureRecognizer);
                        DataGrid.Children.Add(productLabel, 0, newRow);

                        var timeLabel = new Label
                        {
                            Text = data.TimePerUnit,
                            HorizontalTextAlignment = TextAlignment.Center,
                            VerticalTextAlignment = TextAlignment.Center,
                            TextColor = Color.Black,
                            FontSize = 14
                        };
                        timeLabel.GestureRecognizers.Add(tapGestureRecognizer);
                        DataGrid.Children.Add(timeLabel, 1, newRow);

                        var quantityLabel = new Label
                        {
                            Text = data.Quantity,
                            HorizontalTextAlignment = TextAlignment.Center,
                            VerticalTextAlignment = TextAlignment.Center,
                            TextColor = Color.Black,
                            FontSize = 14
                        };
                        quantityLabel.GestureRecognizers.Add(tapGestureRecognizer);
                        DataGrid.Children.Add(quantityLabel, 2, newRow);

                        var monumberLabel = new Label
                        {
                            Text = data.MONumber,
                            HorizontalTextAlignment = TextAlignment.Center,
                            VerticalTextAlignment = TextAlignment.Center,
                            TextColor = Color.Black,
                            FontSize = 14
                        };
                        monumberLabel.GestureRecognizers.Add(tapGestureRecognizer);
                        DataGrid.Children.Add(monumberLabel, 3, newRow);

                        var commentsLabel = new Label
                        {
                            Text = data.Comments,
                            HorizontalTextAlignment = TextAlignment.Center,
                            VerticalTextAlignment = TextAlignment.Center,
                            TextColor = Color.Black,
                            FontSize = 14
                        };
                        commentsLabel.GestureRecognizers.Add(tapGestureRecognizer);
                        DataGrid.Children.Add(commentsLabel, 4, newRow);

                        var resultLabel = new Label
                        {
                            Text = data.Result,
                            HorizontalTextAlignment = TextAlignment.Center,
                            VerticalTextAlignment = TextAlignment.Center,
                            TextColor = Color.Black,
                            FontSize = 14
                        };
                        resultLabel.GestureRecognizers.Add(tapGestureRecognizer);
                        DataGrid.Children.Add(resultLabel, 5, newRow);
                    }
                }

                await DisplayAlert("Success", "Data imported successfully and displayed.", "OK");
            }
            catch (Exception ex)
            {
                await DisplayAlert("Error", $"Failed to import data: {ex.Message}", "OK");
            }
        }

        // Method to load Excel file
        private void LoadExcelFile(string filePath)
        {
            try
            {
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

                        // Add the imported data to the dataList
                        dataList.Add(data);

                        // Dynamically add data to the grid
                        int newRow = DataGrid.RowDefinitions.Count;
                        DataGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                        var tapGestureRecognizer = new TapGestureRecognizer();
                        tapGestureRecognizer.Tapped += (s, args) => OnRowTapped(newRow);

                        var productLabel = new Label
                        {
                            Text = data.Product,
                            HorizontalTextAlignment = TextAlignment.Center,
                            VerticalTextAlignment = TextAlignment.Center,
                            TextColor = Color.Black,
                            FontSize = 14
                        };
                        productLabel.GestureRecognizers.Add(tapGestureRecognizer);
                        DataGrid.Children.Add(productLabel, 0, newRow);

                        var timeLabel = new Label
                        {
                            Text = data.TimePerUnit,
                            HorizontalTextAlignment = TextAlignment.Center,
                            VerticalTextAlignment = TextAlignment.Center,
                            TextColor = Color.Black,
                            FontSize = 14
                        };
                        timeLabel.GestureRecognizers.Add(tapGestureRecognizer);
                        DataGrid.Children.Add(timeLabel, 1, newRow);

                        var quantityLabel = new Label
                        {
                            Text = data.Quantity,
                            HorizontalTextAlignment = TextAlignment.Center,
                            VerticalTextAlignment = TextAlignment.Center,
                            TextColor = Color.Black,
                            FontSize = 14
                        };
                        quantityLabel.GestureRecognizers.Add(tapGestureRecognizer);
                        DataGrid.Children.Add(quantityLabel, 2, newRow);

                        var monumberLabel = new Label
                        {
                            Text = data.MONumber,
                            HorizontalTextAlignment = TextAlignment.Center,
                            VerticalTextAlignment = TextAlignment.Center,
                            TextColor = Color.Black,
                            FontSize = 14
                        };
                        monumberLabel.GestureRecognizers.Add(tapGestureRecognizer);
                        DataGrid.Children.Add(monumberLabel, 3, newRow);

                        var commentsLabel = new Label
                        {
                            Text = data.Comments,
                            HorizontalTextAlignment = TextAlignment.Center,
                            VerticalTextAlignment = TextAlignment.Center,
                            TextColor = Color.Black,
                            FontSize = 14
                        };
                        commentsLabel.GestureRecognizers.Add(tapGestureRecognizer);
                        DataGrid.Children.Add(commentsLabel, 4, newRow);

                        var resultLabel = new Label
                        {
                            Text = data.Result,
                            HorizontalTextAlignment = TextAlignment.Center,
                            VerticalTextAlignment = TextAlignment.Center,
                            TextColor = Color.Black,
                            FontSize = 14
                        };
                        resultLabel.GestureRecognizers.Add(tapGestureRecognizer);
                        DataGrid.Children.Add(resultLabel, 5, newRow);
                    }

                    DisplayAlert("Success", "Data imported successfully.", "OK");
                }
            }
            catch (Exception ex)
            {
                DisplayAlert("Error", $"Failed to load Excel file: {ex.Message}", "OK");
            }
        }

        private void BtnExit_Clicked(object sender, EventArgs e)
        {
            System.Diagnostics.Process.GetCurrentProcess().CloseMainWindow();
        }
    }
}