using System;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Drawing;
using System.IO;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace ExcelAddInDemo
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var filePath = new FileInfo("C:\\Users\\Kalina Aleksandrova\\Documents\\EmployeeData.xlsx");

            if (File.Exists(filePath.FullName))
                File.Delete(filePath.FullName);

            using (var package = new ExcelPackage(filePath))
            {
                // Create a new worksheet
                var worksheet = package.Workbook.Worksheets.Add("Employee Data");

                // Add some header cells
                worksheet.Cells["A1"].Value = "ID";
                worksheet.Cells["B1"].Value = "Name";
                worksheet.Cells["C1"].Value = "Position";
                worksheet.Cells["D1"].Value = "Salary";
                worksheet.Cells["E1"].Value = "Tax";
                worksheet.Cells["F1"].Value = "Net Salary";

                // Apply some styling to the headers
                using (var range = worksheet.Cells["A1:F1"])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    range.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                }

                // Add employee data
                var employees = new[]
                {
                    new { Id = 1, Name = "Peter Petrov", Position = "Manager", Salary = 5500 },
                    new { Id = 2, Name = "Ivan Ivanov", Position = "Engineer", Salary = 4500 },
                    new { Id = 3, Name = "Maria Marinova", Position = "Recruiter", Salary = 3000 }
                };

                int row = 2;
                foreach (var employee in employees)
                {
                    worksheet.Cells[row, 1].Value = employee.Id;
                    worksheet.Cells[row, 2].Value = employee.Name;
                    worksheet.Cells[row, 3].Value = employee.Position;
                    worksheet.Cells[row, 4].Value = employee.Salary;

                    // Add a formula for tax (10% of salary) and net salary
                    worksheet.Cells[row, 5].Formula = $"D{row}*0.1";
                    worksheet.Cells[row, 6].Formula = $"D{row}-E{row}";
                    row++;
                }

                // Apply some currency format to salary, tax, and net salary columns
                using (var range = worksheet.Cells["D2:F4"])
                {
                    range.Style.Numberformat.Format = "$#,##0.00";
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }

                worksheet.Cells.AutoFitColumns();

                // Add a chart to represent salary data
                var chart = worksheet.Drawings.AddChart("SalaryChart", eChartType.ColumnClustered);
                chart.Title.Text = "Employee Salary Data";
                chart.SetPosition(6, 0, 1, 0);
                chart.SetSize(600, 300);

                // Add series to the chart
                var series = chart.Series.Add(worksheet.Cells["D2:D4"], worksheet.Cells["B2:B4"]);
                series.Header = "Salary";

                // Save the workbook
                package.Save();
            }

            // Read the data back from the file
            using (var package = new ExcelPackage(filePath))
            {
                var worksheet = package.Workbook.Worksheets["Employee Data"];

                for (int row = 2; row <= 4; row++)
                {
                    var id = worksheet.Cells[row, 1].Text;
                    var name = worksheet.Cells[row, 2].Text;
                    var position = worksheet.Cells[row, 3].Text;
                    var salary = worksheet.Cells[row, 4].Text;
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
