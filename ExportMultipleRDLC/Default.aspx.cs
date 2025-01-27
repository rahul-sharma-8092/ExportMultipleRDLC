using Microsoft.Reporting.WebForms;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Web.UI;

namespace ExportMultipleRDLC
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            Page.Title = "Multi RDLC to Excel";
            if (!Page.IsPostBack)
            {
                BindRDLC();
            }
        }

        protected void BtnExportAllReport_Click(object sender, EventArgs e)
        {
            string fileName = "All_Combined_Reports_Rahul_Sharma_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

            byte[] reportByte1 = GenerateExcelByte(ReportViewer1.LocalReport);
            byte[] reportByte2 = GenerateExcelByte(ReportViewer2.LocalReport);
            byte[] reportByte3 = GenerateExcelByte(ReportViewer3.LocalReport);
            byte[] reportByte4 = GenerateExcelByte(ReportViewer4.LocalReport);

            // Using EpPlus [V4.5.3.3] free version
            byte[] excelFile = CreateMultiSheetExcel(new[]
            {
                ("CarReport", reportByte1),
                ("DeveloperReport", reportByte2),
                ("ProjectReport", reportByte3),
                ("GamesReport", reportByte4)
            });

            string base64Excel = Convert.ToBase64String(excelFile, 0, excelFile.Length);

            string script = $"downloadExcelFromBase64('{base64Excel}', '{fileName}');";

            ScriptManager.RegisterStartupScript(this, this.GetType(), "DownloadExcel", script, true);
        }

        #region Combine multiple RDLC into Excel ==> EpPlus [4.5.3.3] free version
        public byte[] CreateMultiSheetExcel((string sheetName, byte[] data)[] reports)
        {
            string outputDirectory = Server.MapPath("~/Reports/ExcelSheet");
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            // Save all reports as individual Excel files
            foreach (var (sheetName, data) in reports)
            {
                string fileName = sheetName + "-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
                string filePath = Path.Combine(outputDirectory,  fileName);
                File.WriteAllBytes(filePath, data);
            }

            try
            {
                using (var package = new ExcelPackage())
                {
                    foreach (var (sheetName, data) in reports)
                    {
                        // Create a new worksheet for each report
                        var worksheet = package.Workbook.Worksheets.Add(sheetName);

                        // Load the byte array (Excel data) into a MemoryStream
                        using (var memoryStream = new MemoryStream(data))
                        using (var tempPackage = new ExcelPackage(memoryStream))
                        {
                            var tempSheet = tempPackage.Workbook.Worksheets[1];

                            // Copy the entire worksheet data and design
                            for (int row = tempSheet.Dimension.Start.Row; row <= tempSheet.Dimension.End.Row; row++)
                            {
                                for (int col = tempSheet.Dimension.Start.Column; col <= tempSheet.Dimension.End.Column; col++)
                                {
                                    worksheet.Cells[row, col].Value = tempSheet.Cells[row, col].Value;

                                    // Copy font style
                                    var font = tempSheet.Cells[row, col].Style.Font;
                                    worksheet.Cells[row, col].Style.Font.Size = font.Size;
                                    worksheet.Cells[row, col].Style.Font.Bold = font.Bold;
                                    worksheet.Cells[row, col].Style.Font.Italic = font.Italic;
                                    worksheet.Cells[row, col].Style.Font.UnderLine = font.UnderLine;
                                    worksheet.Cells[row, col].Style.Font.Strike = font.Strike;
                                    worksheet.Cells[row, col].Style.Font.Family = font.Family;
                                    worksheet.Cells[row, col].Style.Font.Name = font.Name;
                                    worksheet.Cells[row, col].Style.Font.UnderLineType = font.UnderLineType;
                                    worksheet.Cells[row, col].Style.Font.VerticalAlign = font.VerticalAlign;

                                    string rgb = "";
                                    if (!string.IsNullOrEmpty(font.Color.Rgb))
                                    {
                                        rgb = "#" + font.Color.Rgb.Substring(2);
                                        System.Drawing.Color fontColor = ColorTranslator.FromHtml(rgb);
                                        worksheet.Cells[row, col].Style.Font.Color.SetColor(fontColor);
                                    }

                                    // Copy fill style
                                    worksheet.Cells[row, col].Style.Fill.PatternType = tempSheet.Cells[row, col].Style.Fill.PatternType;

                                    if (!string.IsNullOrEmpty(tempSheet.Cells[row, col].Style.Fill.BackgroundColor.Rgb))
                                    {
                                        rgb = "#" + tempSheet.Cells[row, col].Style.Fill.BackgroundColor.Rgb.Substring(2);
                                        System.Drawing.Color bgColor = ColorTranslator.FromHtml(rgb);
                                        worksheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(bgColor);
                                    }

                                    // Copy border style
                                    worksheet.Cells[row, col].Style.Border.Top.Style = tempSheet.Cells[row, col].Style.Border.Top.Style;
                                    worksheet.Cells[row, col].Style.Border.Left.Style = tempSheet.Cells[row, col].Style.Border.Left.Style;
                                    worksheet.Cells[row, col].Style.Border.Right.Style = tempSheet.Cells[row, col].Style.Border.Right.Style;
                                    worksheet.Cells[row, col].Style.Border.Bottom.Style = tempSheet.Cells[row, col].Style.Border.Bottom.Style;

                                    // Copy number format style
                                    worksheet.Cells[row, col].Style.Numberformat.Format = tempSheet.Cells[row, col].Style.Numberformat.Format;

                                    // Copy alignment style
                                    worksheet.Cells[row, col].Style.HorizontalAlignment = tempSheet.Cells[row, col].Style.HorizontalAlignment;
                                    worksheet.Cells[row, col].Style.VerticalAlignment = tempSheet.Cells[row, col].Style.VerticalAlignment;
                                    worksheet.Cells[row, col].Style.WrapText = tempSheet.Cells[row, col].Style.WrapText;
                                    worksheet.Cells[row, col].Style.Indent = tempSheet.Cells[row, col].Style.Indent;
                                    worksheet.Cells[row, col].Style.ShrinkToFit = tempSheet.Cells[row, col].Style.ShrinkToFit;
                                    worksheet.Cells[row, col].Style.TextRotation = tempSheet.Cells[row, col].Style.TextRotation;
                                    worksheet.Cells[row, col].Style.ReadingOrder = tempSheet.Cells[row, col].Style.ReadingOrder;
                                    worksheet.Cells[row, col].Style.QuotePrefix = tempSheet.Cells[row, col].Style.QuotePrefix;
                                    worksheet.Cells[row, col].Style.Hidden = tempSheet.Cells[row, col].Style.Hidden;
                                    worksheet.Cells[row, col].Style.Locked = tempSheet.Cells[row, col].Style.Locked;
                                }
                            }

                            // Copy merged cells
                            foreach (var mergeCell in tempSheet.MergedCells)
                            {
                                worksheet.Cells[mergeCell].Merge = true;
                            }

                            // Copy column widths
                            for (int col = tempSheet.Dimension.Start.Column; col <= tempSheet.Dimension.End.Column; col++)
                            {
                                worksheet.Column(col).Width = tempSheet.Column(col).Width;
                            }

                            // Copy row heights
                            for (int row = tempSheet.Dimension.Start.Row; row <= tempSheet.Dimension.End.Row; row++)
                            {
                                worksheet.Row(row).Height = tempSheet.Row(row).Height;
                            }

                            // Copy worksheet image
                            foreach (var drawing in tempSheet.Drawings)
                            {
                                if (drawing is ExcelPicture sourcePicture)
                                {
                                    var picture = worksheet.Drawings.AddPicture(sourcePicture.Name, sourcePicture.Image);

                                    // Set the position (row, column, and offsets)
                                    picture.SetPosition(
                                        sourcePicture.From.Row, sourcePicture.From.RowOff,
                                        sourcePicture.From.Column, sourcePicture.From.ColumnOff
                                    );

                                    var widthField = typeof(ExcelPicture).GetField("_width", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                                    var heightField = typeof(ExcelPicture).GetField("_height", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

                                    int width = (int)widthField.GetValue(sourcePicture);
                                    int height = (int)heightField.GetValue(sourcePicture);

                                    picture.SetSize(width, height);
                                }
                            }

                            // Copy worksheet properties (optional)
                            worksheet.View.PageLayoutView = tempSheet.View.PageLayoutView;
                            worksheet.View.ShowGridLines = tempSheet.View.ShowGridLines;
                        }
                    }

                    return package.GetAsByteArray();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region Generate Report Excel byte array from RDLC
        public byte[] GenerateExcelByte(Microsoft.Reporting.WebForms.LocalReport ReportViewer)
        {
            string format = "ExcelOpenXml";
            string mimeType = string.Empty;
            string encoding = string.Empty;
            string extension = string.Empty;
            string[] streamids;
            Warning[] warnings;

            byte[] bytes = ReportViewer.Render(format, null, out mimeType, out encoding, out extension, out streamids, out warnings);

            return bytes;
        }
        #endregion

        #region Bind RDLC
        public void BindRDLC()
        {
            //ReportViewer1
            DataTable carData = GetCarData();
            ReportViewer1.Reset();
            ReportViewer1.LocalReport.ReportPath = Server.MapPath("~/Reports/RDLC/CarReport.rdlc");

            ReportDataSource reportDataSource1 = new ReportDataSource("Car", carData);

            ReportViewer1.LocalReport.DataSources.Clear();
            ReportViewer1.LocalReport.DataSources.Add(reportDataSource1);
            ReportViewer1.LocalReport.Refresh();


            //ReportViewer2
            DataTable developerData = GetDeveloperData();
            ReportViewer2.Reset();
            ReportViewer2.LocalReport.ReportPath = Server.MapPath("~/Reports/RDLC/DeveloperReport.rdlc");

            ReportDataSource reportDataSource2 = new ReportDataSource("Developer", developerData);

            ReportViewer2.LocalReport.DataSources.Clear();
            ReportViewer2.LocalReport.DataSources.Add(reportDataSource2);
            ReportViewer2.LocalReport.Refresh();


            //ReportViewer3
            DataTable projectData = GetProjectData();
            ReportViewer3.Reset();
            ReportViewer3.LocalReport.ReportPath = Server.MapPath("~/Reports/RDLC/ProjectReport.rdlc");
            ReportDataSource reportDataSource3 = new ReportDataSource("Project", projectData);
            ReportViewer3.LocalReport.DataSources.Clear();
            ReportViewer3.LocalReport.DataSources.Add(reportDataSource3);
            ReportViewer3.LocalReport.Refresh();

            //ReportViewer4
            DataTable gameData = GetGameData();
            ReportViewer4.Reset();
            ReportViewer4.LocalReport.ReportPath = Server.MapPath("~/Reports/RDLC/GamesReport.rdlc");
            ReportDataSource reportDataSource4 = new ReportDataSource("Games", gameData);
            ReportViewer4.LocalReport.DataSources.Clear();
            ReportViewer4.LocalReport.DataSources.Add(reportDataSource4);
            ReportViewer4.LocalReport.Refresh();
        }
        #endregion

        #region DataTables
        public DataTable GetCarData()
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Company", typeof(string));
            dataTable.Columns.Add("Model", typeof(string));
            dataTable.Columns.Add("EngineType", typeof(string));
            dataTable.Columns.Add("Price", typeof(int));
            dataTable.Columns.Add("UserExperience", typeof(string));

            dataTable.Rows.Add("Toyota", "Corolla", "Petrol", 25000, "Good");
            dataTable.Rows.Add("Honda", "Civic", "Diesel", 27000, "Excellent");
            dataTable.Rows.Add("Ford", "Focus", "Electric", 30000, "Average");
            dataTable.Rows.Add("BMW", "X5", "Hybrid", 60000, "Excellent");
            dataTable.Rows.Add("Audi", "A4", "Petrol", 35000, "Good");
            dataTable.Rows.Add("Mercedes", "C-Class", "Diesel", 45000, "Good");
            dataTable.Rows.Add("Tesla", "Model S", "Electric", 80000, "Excellent");
            dataTable.Rows.Add("Nissan", "Altima", "Petrol", 28000, "Average");
            dataTable.Rows.Add("Chevrolet", "Malibu", "Hybrid", 35000, "Good");
            dataTable.Rows.Add("Hyundai", "Sonata", "Diesel", 32000, "Average");
            dataTable.Rows.Add("Toyota", "Corolla", "Petrol", 25000, "Good");
            dataTable.Rows.Add("Honda", "Civic", "Diesel", 27000, "Excellent");
            dataTable.Rows.Add("Ford", "Focus", "Electric", 30000, "Average");
            dataTable.Rows.Add("BMW", "X5", "Hybrid", 60000, "Excellent");
            dataTable.Rows.Add("Audi", "A4", "Petrol", 35000, "Good");
            dataTable.Rows.Add("Mercedes", "C-Class", "Diesel", 45000, "Good");
            dataTable.Rows.Add("Tesla", "Model S", "Electric", 80000, "Excellent");
            dataTable.Rows.Add("Nissan", "Altima", "Petrol", 28000, "Average");
            dataTable.Rows.Add("Chevrolet", "Malibu", "Hybrid", 35000, "Good");
            dataTable.Rows.Add("Hyundai", "Sonata", "Diesel", 32000, "Average");
            dataTable.Rows.Add("Ford", "Focus", "Electric", 30000, "Average");
            dataTable.Rows.Add("BMW", "X5", "Hybrid", 60000, "Excellent");
            dataTable.Rows.Add("Audi", "A4", "Petrol", 35000, "Good");
            dataTable.Rows.Add("Mercedes", "C-Class", "Diesel", 45000, "Good");
            dataTable.Rows.Add("Tesla", "Model S", "Electric", 80000, "Excellent");
            dataTable.Rows.Add("Nissan", "Altima", "Petrol", 28000, "Average");
            dataTable.Rows.Add("Chevrolet", "Malibu", "Hybrid", 35000, "Good");
            dataTable.Rows.Add("Hyundai", "Sonata", "Diesel", 32000, "Average");
            dataTable.Rows.Add("Toyota", "Corolla", "Petrol", 25000, "Good");
            dataTable.Rows.Add("Honda", "Civic", "Diesel", 27000, "Excellent");
            dataTable.Rows.Add("Ford", "Focus", "Electric", 30000, "Average");
            dataTable.Rows.Add("BMW", "X5", "Hybrid", 60000, "Excellent");
            dataTable.Rows.Add("Audi", "A4", "Petrol", 35000, "Good");
            dataTable.Rows.Add("Mercedes", "C-Class", "Diesel", 45000, "Good");
            dataTable.Rows.Add("Tesla", "Model S", "Electric", 80000, "Excellent");
            dataTable.Rows.Add("Nissan", "Altima", "Petrol", 28000, "Average");
            dataTable.Rows.Add("Chevrolet", "Malibu", "Hybrid", 35000, "Good");
            dataTable.Rows.Add("Ford", "Focus", "Electric", 30000, "Average");
            dataTable.Rows.Add("BMW", "X5", "Hybrid", 60000, "Excellent");
            dataTable.Rows.Add("Audi", "A4", "Petrol", 35000, "Good");
            dataTable.Rows.Add("Mercedes", "C-Class", "Diesel", 45000, "Good");
            dataTable.Rows.Add("Tesla", "Model S", "Electric", 80000, "Excellent");
            dataTable.Rows.Add("Nissan", "Altima", "Petrol", 28000, "Average");
            dataTable.Rows.Add("Chevrolet", "Malibu", "Hybrid", 35000, "Good");
            dataTable.Rows.Add("Hyundai", "Sonata", "Diesel", 32000, "Average");
            dataTable.Rows.Add("Toyota", "Corolla", "Petrol", 25000, "Good");
            dataTable.Rows.Add("Honda", "Civic", "Diesel", 27000, "Excellent");
            dataTable.Rows.Add("Ford", "Focus", "Electric", 30000, "Average");
            dataTable.Rows.Add("BMW", "X5", "Hybrid", 60000, "Excellent");
            dataTable.Rows.Add("Audi", "A4", "Petrol", 35000, "Good");
            dataTable.Rows.Add("Mercedes", "C-Class", "Diesel", 45000, "Good");
            dataTable.Rows.Add("Tesla", "Model S", "Electric", 80000, "Excellent");
            dataTable.Rows.Add("Nissan", "Altima", "Petrol", 28000, "Average");
            dataTable.Rows.Add("Chevrolet", "Malibu", "Hybrid", 35000, "Good");

            return dataTable;
        }

        public DataTable GetDeveloperData()
        {
            // Create a DataTable
            DataTable developerTable = new DataTable();

            // Define the columns
            developerTable.Columns.Add("Name", typeof(string));
            developerTable.Columns.Add("Department", typeof(string));
            developerTable.Columns.Add("Salary", typeof(int));
            developerTable.Columns.Add("Experience", typeof(int));

            // Add rows of data
            developerTable.Rows.Add("Alice Johnson", "Software Development", 85000, 5);
            developerTable.Rows.Add("Bob Smith", "Web Development", 78000, 4);
            developerTable.Rows.Add("Charlie Brown", "Mobile Development", 95000, 6);
            developerTable.Rows.Add("Diana Prince", "Data Science", 105000, 7);
            developerTable.Rows.Add("Ethan Hunt", "Software Development", 89000, 5);
            developerTable.Rows.Add("Fiona Adams", "UI/UX Design", 72000, 3);
            developerTable.Rows.Add("George Michaels", "Database Management", 83000, 4);
            developerTable.Rows.Add("Hannah Lee", "Cloud Computing", 98000, 6);
            developerTable.Rows.Add("Ian Carter", "Cybersecurity", 102000, 8);
            developerTable.Rows.Add("Jessica Day", "Software Development", 87000, 5);
            developerTable.Rows.Add("Kevin Parker", "Web Development", 81000, 4);
            developerTable.Rows.Add("Laura Kim", "Mobile Development", 95000, 6);
            developerTable.Rows.Add("Michael Scott", "Data Science", 100000, 7);
            developerTable.Rows.Add("Nancy Drew", "UI/UX Design", 75000, 3);
            developerTable.Rows.Add("Oliver Stone", "Database Management", 85000, 5);
            developerTable.Rows.Add("Paula White", "Cloud Computing", 97000, 6);
            developerTable.Rows.Add("Quentin Blake", "Cybersecurity", 103000, 8);
            developerTable.Rows.Add("Rachel Green", "Software Development", 88000, 5);
            developerTable.Rows.Add("Steve Rogers", "Web Development", 79000, 4);
            developerTable.Rows.Add("Tina Turner", "Mobile Development", 94000, 6);

            return developerTable;
        }

        public DataTable GetProjectData()
        {
            // Create a DataTable
            DataTable projectTable = new DataTable();

            // Define the columns
            projectTable.Columns.Add("ProjectName", typeof(string));
            projectTable.Columns.Add("Framework", typeof(string));
            projectTable.Columns.Add("ProjectType", typeof(string));
            projectTable.Columns.Add("Client", typeof(string));

            // Add rows of data
            projectTable.Rows.Add("Inventory Management System", ".NET Core", "Web Application", "Acme Corp");
            projectTable.Rows.Add("E-Commerce Platform", "React", "Web Application", "Tech Innovators");
            projectTable.Rows.Add("Mobile Banking App", "Flutter", "Mobile Application", "Global Bank");
            projectTable.Rows.Add("Healthcare Portal", "Angular", "Web Application", "MediCare Inc.");
            projectTable.Rows.Add("CRM Software", "Java Spring", "Desktop Application", "Sales Experts");
            projectTable.Rows.Add("Social Media App", "Node.js", "Web Application", "ConnectWorld");
            projectTable.Rows.Add("Weather Monitoring System", "Python Flask", "IoT Application", "Climate Solutions");
            projectTable.Rows.Add("Education LMS", "Vue.js", "Web Application", "EdTech Leaders");
            projectTable.Rows.Add("Travel Booking System", ".NET Framework", "Desktop Application", "Go Travel");
            projectTable.Rows.Add("Gaming Platform", "Unity", "Game Development", "FunSoft Games");
            projectTable.Rows.Add("Fitness Tracker App", "Swift", "Mobile Application", "FitLife Inc.");
            projectTable.Rows.Add("Online Food Delivery System", "PHP Laravel", "Web Application", "Yummy Bites");
            projectTable.Rows.Add("Blockchain Wallet", "React", "Web Application", "Crypto Wallets Ltd.");
            projectTable.Rows.Add("E-Library System", "Django", "Web Application", "EduBooks Co.");
            projectTable.Rows.Add("Virtual Reality App", "Unity", "Game Development", "VisionVR");
            projectTable.Rows.Add("Warehouse Management System", "Java", "Desktop Application", "LogiTrack Systems");
            projectTable.Rows.Add("Video Streaming Platform", "React", "Web Application", "StreamNow");
            projectTable.Rows.Add("AI Chatbot", "Python TensorFlow", "AI Application", "SmartAssist Inc.");
            projectTable.Rows.Add("Online Voting System", ".NET Core", "Web Application", "GovTech Solutions");
            projectTable.Rows.Add("Event Management Platform", "Vue.js", "Web Application", "PlanIt Events");

            return projectTable;
        }

        public DataTable GetGameData()
        {
            // Create a DataTable
            DataTable gameTable = new DataTable();

            // Define the columns
            gameTable.Columns.Add("GameName", typeof(string));
            gameTable.Columns.Add("PCorMobile", typeof(string));
            gameTable.Columns.Add("Rating", typeof(int));
            gameTable.Columns.Add("Multiplayer", typeof(string));
            gameTable.Columns.Add("MinRequirements", typeof(string));

            // Add rows of data
            gameTable.Rows.Add("Call of Duty: Modern Warfare", "PC", 5, "Yes", "16GB RAM, GTX 1080");
            gameTable.Rows.Add("Fortnite", "PC", 4, "Yes", "8GB RAM, GTX 660");
            gameTable.Rows.Add("PUBG Mobile", "Mobile", 4, "Yes", "2GB RAM, Snapdragon 625");
            gameTable.Rows.Add("Minecraft", "PC", 5, "Yes", "4GB RAM, Integrated Graphics");
            gameTable.Rows.Add("Among Us", "Mobile", 4, "Yes", "1GB RAM, Any Processor");
            gameTable.Rows.Add("The Witcher 3", "PC", 5, "No", "8GB RAM, GTX 970");
            gameTable.Rows.Add("Clash of Clans", "Mobile", 4, "Yes", "1GB RAM, Any Processor");
            gameTable.Rows.Add("Assassin's Creed Valhalla", "PC", 5, "No", "16GB RAM, RTX 2060");
            gameTable.Rows.Add("Candy Crush Saga", "Mobile", 3, "No", "512MB RAM, Any Processor");
            gameTable.Rows.Add("League of Legends", "PC", 5, "Yes", "4GB RAM, GTX 560");
            gameTable.Rows.Add("Valorant", "PC", 5, "Yes", "4GB RAM, Intel HD Graphics 4000");
            gameTable.Rows.Add("Free Fire", "Mobile", 4, "Yes", "2GB RAM, Snapdragon 450");
            gameTable.Rows.Add("Genshin Impact", "PC", 4, "Yes", "8GB RAM, GTX 1060");
            gameTable.Rows.Add("Subway Surfers", "Mobile", 3, "No", "1GB RAM, Any Processor");
            gameTable.Rows.Add("Apex Legends", "PC", 4, "Yes", "8GB RAM, GTX 970");
            gameTable.Rows.Add("Battlefield V", "PC", 5, "Yes", "12GB RAM, RTX 2060");
            gameTable.Rows.Add("Roblox", "PC", 4, "Yes", "4GB RAM, Integrated Graphics");
            gameTable.Rows.Add("Temple Run", "Mobile", 3, "No", "1GB RAM, Any Processor");
            gameTable.Rows.Add("Overwatch 2", "PC", 5, "Yes", "6GB RAM, GTX 960");
            gameTable.Rows.Add("Hearthstone", "Mobile", 4, "Yes", "2GB RAM, Snapdragon 625");

            return gameTable;
        }
        #endregion
    }
}