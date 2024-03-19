using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OfficeOpenXml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Excel_To_TXT
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Filter to allow only excel files and selecting multiple of them 
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            // Checks if the file is selected in Dialog
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                int folderIndex = 1; 

                foreach (string selectedFilePath in openFileDialog.FileNames)
                {
                    // Create a folder on the desktop to store the folders for each column
                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string mainFolderPath = Path.Combine(desktopPath, $"ExcelToTextFiles_{folderIndex}");
                    Directory.CreateDirectory(mainFolderPath);

                    using (var excelPackage = new ExcelPackage(new FileInfo(selectedFilePath)))
                    {
                        var worksheet = excelPackage.Workbook.Worksheets[0];
                        int rowCount = worksheet.Dimension.Rows;
                        int colCount = worksheet.Dimension.Columns;

                        // Iterate through each column
                        for (int col = 1; col <= colCount; col++)
                        {
                            // Get the name text from the first row of the current column
                            string headerText = FixFileName(worksheet.Cells[1, col].Value?.ToString());

                            // Create a folder for the current column
                            string columnFolderPath = Path.Combine(mainFolderPath, headerText);
                            Directory.CreateDirectory(columnFolderPath);

                            // Iterate through each row starting from the second row
                            for (int row = 2; row <= rowCount; row++)
                            {
                                // Get the text from the current cell
                                string text = worksheet.Cells[row, col].Value?.ToString();
                                string fileNameText = FixFileName(text);

                                // If the text is null or empty, skip to the next row
                                if (string.IsNullOrEmpty(text))
                                    continue;

                                // Create a text file with the row's value as its name and write the text into it
                                string fileName = $"{fileNameText}.txt";
                                string filePath = Path.Combine(columnFolderPath, fileName);
                                File.WriteAllText(filePath, text.ToLower());
                            }
                        }
                    }

                    folderIndex++; 
                }

                MessageBox.Show("Text files have been created successfully.");
            }
        }

        // Method to make sure the files are not illegaly named 
        private string FixFileName(string fileName)
            {
                // Remove any characters that are not letters, numbers, or underscores

                return Regex.Replace(fileName, "[^a-zA-Z0-9_]", " ");
            }

            private void textBox1_TextChanged(object sender, EventArgs e)
            {

            }
    }
} 
