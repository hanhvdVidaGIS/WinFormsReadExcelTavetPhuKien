using System.Globalization;
using System.IO;
using System.Reflection.PortableExecutable;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
namespace WinFormsReadExcelTavetPhuKien
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ComboboxItem item = new ComboboxItem();
            item.Text = "Tà vẹt";
            item.Value = "tavet";
            comboBox1.Items.Add(item);
            item = new ComboboxItem();
            item.Text = "Phụ kiện";
            item.Value = "phukien";
            comboBox1.Items.Add(item);

            comboBox1.SelectedIndex = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create an instance of OpenFileDialog
            OpenFileDialog openFileDialog = new OpenFileDialog();

            //Set the filter to show only Excel files
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            //Check if the user selected a file
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                //Get the file path
                string filePath = openFileDialog.FileName;

                //Display the file path in the text box
                richTextBox1.Text = "Reading File...";

                //reading excel


                //Clear the rich text box
                richTextBox1.Clear();

                this.ReadExcel(filePath);
            }
        }
        private void ReadExcel(string filePath)
        {
            var typeTableVal = ((ComboboxItem)comboBox1.SelectedItem).Value as string;
            var typeTableText = ((ComboboxItem)comboBox1.SelectedItem).Text as string;

            FileInfo fileInfo = new FileInfo(filePath);
            //package read
            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault()!;

            //package response
            ExcelPackage packageResponse = new ExcelPackage();
            ExcelWorksheet SheetResponse = packageResponse.Workbook.Worksheets.Add(typeTableText ?? "Du lieu chuan hoa");

            //Chèn alias vào row 1 đàu tiên
            for (int i = 1; i < 10; i++)
            {
                string textAlias = "";
                switch (i)
                {
                    case 1:
                        textAlias = "STT";
                        break;
                    case 2:
                        textAlias = "Tuyến đường sắt";
                        break;
                    case 3:
                        textAlias = "Khu gian";
                        break;
                    case 4:
                        textAlias = "Tên cầu ray";
                        break;
                    case 5:
                        textAlias = "Lý trình đầu (Km)";
                        break;
                    case 6:
                        textAlias = "Lý trình cuối (Km)";
                        break;
                    case 7:
                        textAlias = typeTableVal == "tavet" ? "Loại tà vẹt đường sắt" : "Loại phụ kiện";
                        break;
                    case 8:
                        textAlias = "Số lượng (tốt)";
                        break;
                    case 9:
                        textAlias = "Số lượng (xấu)";
                        break;
                }
                var value_alias = SheetResponse.Cells[2, i].RichText.Add(textAlias);
                value_alias.Bold = true;
            }

            // get number of rows and columns in the sheet
            int rows = worksheet.Dimension.Rows;
            int columns = worksheet.Dimension.Columns;

            // loop through the worksheet rows and columns
            int currentRow = 3;
            for (int i = 5; i <= rows; i++)
            {
                bool isnullRow = true;
                List<string> values = new List<string>();
                for (int j = 1; j <= columns; j++)
                {
                    string content = worksheet.Cells[i, j].Value?.ToString() ?? "";
                    values.Add(content);
                    if (!string.IsNullOrEmpty(content.Trim()))
                    {
                        isnullRow = false;
                    }
                }
                if (isnullRow)
                {
                    richTextBox1.Text = ($"Row {i} of total {rows} isnullRow \n");
                    break;
                }
                
                if (typeTableVal == "tavet")
                {
                    //value tavet
                    this.HandleQuantityTavet(values, ref currentRow, ref SheetResponse);
                    richTextBox1.Text = ($"Reading Row {i} of total {rows} \n");
                }
                else if (typeTableVal == "phukien")
                {
                    //value phukien
                    this.HandleQuantityPhukien(values, ref currentRow, ref SheetResponse);
                    richTextBox1.Text = ($"Reading Row {i} of total {rows} \n");
                }

                
            }
            SheetResponse.Cells[3, 5, currentRow - 1, 6].Style.Numberformat.Format = $"K\\m 0. + 000#######";
            SheetResponse.Calculate();
            SheetResponse.Cells[SheetResponse.Dimension.Address].AutoFitColumns();
            using (ExcelRange Rng = SheetResponse.Cells[2, 1, 2, 9])
            {
                Rng.IsRichText = true;
                Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                Rng.AutoFitColumns();
            }
            using (ExcelRange Rng = SheetResponse.Cells[3, 1, currentRow - 1, 9])
            {
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                Rng.AutoFitColumns();
            }
            using (ExcelRange Rng = SheetResponse.Cells[1, 1, currentRow - 1, 9])
            {
                Rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }
            if (package != null) package.Dispose();
            richTextBox1.AppendText($"Successfully read file");
            string excelName = $"{typeTableText}.{DateTime.Now.ToString("yyyy.MM.dd.HH.mm.ss")}"; //$"{typeTableText}.{DateTime.Now.Year}.{DateTime.Now.Month}.{DateTime.Now.Day}.{DateTime.Now.Hour}.{DateTime.Now.Minute}";
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Title = "Export Excel";
            saveFile.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm";
            saveFile.FileName = excelName;
            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    packageResponse.SaveAs(saveFile.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Xuat file loi! \n" + ex.Message);
                }
            }
        }
        private void HandleQuantityTavet(List<string> values, ref int currentRow, ref ExcelWorksheet SheetResponse)
        {
            int.TryParse(values[9 - 1], out int countTavet_BTDUL_Tot);
            int.TryParse(values[10 - 1], out int countTavet_BTDUL_Xau);
            int.TryParse(values[15 - 1], out int countTavet_BT_Tot);
            int.TryParse(values[16 - 1], out int countTavet_BT_Xau);
            int.TryParse(values[21 - 1], out int countTavet_SAT_Tot);
            int.TryParse(values[22 - 1], out int countTavet_SAT_Xau);
            int.TryParse(values[27 - 1], out int countTavet_GO_Tot);
            int.TryParse(values[28 - 1], out int countTavet_GO_Xau);
            int.TryParse(values[33 - 1], out int countTavet_COMPOSITE_Tot);
            int.TryParse(values[34 - 1], out int countTavet_COMPOSITE_Xau);

            string typeTavet = "";
            if (countTavet_BTDUL_Tot > 0 || countTavet_BTDUL_Xau > 0)
            {
                typeTavet = "Tà vẹt bê tông dự ứng lực";

                SheetResponse.Cells[currentRow, 1].Value = currentRow - 2;
                SheetResponse.Cells[currentRow, 2].Value = values[1];
                SheetResponse.Cells[currentRow, 3].Value = values[2];
                
                SheetResponse.Cells[currentRow, 4].Value = values[5];
                if (!string.IsNullOrEmpty(values[6]))
                {
                    double.TryParse(values[6], out double lytrinh_from);
                    SheetResponse.Cells[currentRow, 5].Value = lytrinh_from;
                }
                if (!string.IsNullOrEmpty(values[7]))
                {
                    double.TryParse(values[7], out double lytrinh_to);
                    SheetResponse.Cells[currentRow, 6].Value = lytrinh_to;
                }
                SheetResponse.Cells[currentRow, 7].Value = typeTavet;
                SheetResponse.Cells[currentRow, 8].Value = countTavet_BTDUL_Tot;
                SheetResponse.Cells[currentRow, 9].Value = countTavet_BTDUL_Xau;

                currentRow++;
            }
            if (countTavet_BT_Tot > 0 || countTavet_BT_Xau > 0)
            {
                typeTavet = "Tà vẹt bê tông";

                SheetResponse.Cells[currentRow, 1].Value = currentRow - 2;
                SheetResponse.Cells[currentRow, 2].Value = values[1];
                SheetResponse.Cells[currentRow, 3].Value = values[2];
                SheetResponse.Cells[currentRow, 4].Value = values[5];
                if (!string.IsNullOrEmpty(values[6]))
                {
                    double.TryParse(values[6], out double lytrinh_from);
                    SheetResponse.Cells[currentRow, 5].Value = lytrinh_from;
                }
                if (!string.IsNullOrEmpty(values[7]))
                {
                    double.TryParse(values[7], out double lytrinh_to);
                    SheetResponse.Cells[currentRow, 6].Value = lytrinh_to;
                }

                SheetResponse.Cells[currentRow, 7].Value = typeTavet;
                SheetResponse.Cells[currentRow, 8].Value = countTavet_BT_Tot;
                SheetResponse.Cells[currentRow, 9].Value = countTavet_BT_Xau;

                currentRow++;
            }
            if (countTavet_SAT_Tot > 0 || countTavet_SAT_Tot > 0)
            {
                typeTavet = "Tà vẹt sắt";

                SheetResponse.Cells[currentRow, 1].Value = currentRow - 2;
                SheetResponse.Cells[currentRow, 2].Value = values[1];
                SheetResponse.Cells[currentRow, 3].Value = values[2];
                SheetResponse.Cells[currentRow, 4].Value = values[5];
                if (!string.IsNullOrEmpty(values[6]))
                {
                    double.TryParse(values[6], out double lytrinh_from);
                    SheetResponse.Cells[currentRow, 5].Value = lytrinh_from;
                }
                if (!string.IsNullOrEmpty(values[7]))
                {
                    double.TryParse(values[7], out double lytrinh_to);
                    SheetResponse.Cells[currentRow, 6].Value = lytrinh_to;
                }

                SheetResponse.Cells[currentRow, 7].Value = typeTavet;
                SheetResponse.Cells[currentRow, 8].Value = countTavet_SAT_Tot;
                SheetResponse.Cells[currentRow, 9].Value = countTavet_SAT_Tot;

                currentRow++;
            }
            if (countTavet_GO_Tot > 0 || countTavet_GO_Xau > 0)
            {
                typeTavet = "Tà vẹt gỗ";

                SheetResponse.Cells[currentRow, 1].Value = currentRow - 2;
                SheetResponse.Cells[currentRow, 2].Value = values[1];
                SheetResponse.Cells[currentRow, 3].Value = values[2];
                SheetResponse.Cells[currentRow, 4].Value = values[5];
                if (!string.IsNullOrEmpty(values[6]))
                {
                    double.TryParse(values[6], out double lytrinh_from);
                    SheetResponse.Cells[currentRow, 5].Value = lytrinh_from;
                }
                if (!string.IsNullOrEmpty(values[7]))
                {
                    double.TryParse(values[7], out double lytrinh_to);
                    SheetResponse.Cells[currentRow, 6].Value = lytrinh_to;
                }

                SheetResponse.Cells[currentRow, 7].Value = typeTavet;
                SheetResponse.Cells[currentRow, 8].Value = countTavet_GO_Tot;
                SheetResponse.Cells[currentRow, 9].Value = countTavet_GO_Xau;

                currentRow++;
            }
            if (countTavet_COMPOSITE_Tot > 0 || countTavet_COMPOSITE_Xau > 0)
            {
                typeTavet = "Tà vẹt composite";

                SheetResponse.Cells[currentRow, 1].Value = currentRow - 2;
                SheetResponse.Cells[currentRow, 2].Value = values[1];
                SheetResponse.Cells[currentRow, 3].Value = values[2];
                SheetResponse.Cells[currentRow, 4].Value = values[5];
                if (!string.IsNullOrEmpty(values[6]))
                {
                    double.TryParse(values[6], out double lytrinh_from);
                    SheetResponse.Cells[currentRow, 5].Value = lytrinh_from;
                }
                if (!string.IsNullOrEmpty(values[7]))
                {
                    double.TryParse(values[7], out double lytrinh_to);
                    SheetResponse.Cells[currentRow, 6].Value = lytrinh_to;
                }

                SheetResponse.Cells[currentRow, 7].Value = typeTavet;
                SheetResponse.Cells[currentRow, 8].Value = countTavet_COMPOSITE_Tot;
                SheetResponse.Cells[currentRow, 9].Value = countTavet_COMPOSITE_Xau;

                currentRow++;
            }
        }
        private void HandleQuantityPhukien(List<string> values, ref int currentRow, ref ExcelWorksheet SheetResponse)
        {
            int.TryParse(values[13 - 1], out int countPhukien_CC_BTDUL_Tot);
            int.TryParse(values[14 - 1], out int countPhukien_CC_BTDUL_Xau);
            int.TryParse(values[11 - 1], out int countPhukien_CDH_BTDUL_Tot);
            int.TryParse(values[12 - 1], out int countPhukien_CDH_BTDUL_Xau);

            int.TryParse(values[17 - 1], out int countPhukien_CC_BT_Tot);
            int.TryParse(values[18 - 1], out int countPhukien_CC_BT_Xau);
            int.TryParse(values[19 - 1], out int countPhukien_CDH_BT_Tot);
            int.TryParse(values[20 - 1], out int countPhukien_CDH_BT_Xau);

            int.TryParse(values[23 - 1], out int countPhukien_CC_SAT_Tot);
            int.TryParse(values[24 - 1], out int countPhukien_CC_SAT_Xau);
            int.TryParse(values[25 - 1], out int countPhukien_CDH_SAT_Tot);
            int.TryParse(values[26 - 1], out int countPhukien_CDH_SAT_Xau);

            int.TryParse(values[29 - 1], out int countPhukien_CC_GO_Tot);
            int.TryParse(values[30 - 1], out int countPhukien_CC_GO_Xau);
            int.TryParse(values[31 - 1], out int countPhukien_CDH_GO_Tot);
            int.TryParse(values[32 - 1], out int countPhukien_CDH_GO_Xau);

            int.TryParse(values[35 - 1], out int countPhukien_CDH_COMPOSITE_Tot);
            int.TryParse(values[36 - 1], out int countPhukien_CDH_COMPOSITE_Xau);

            int countPhukien_CC_TOT = countPhukien_CC_BTDUL_Tot + countPhukien_CC_BT_Tot + countPhukien_CC_SAT_Tot + countPhukien_CC_GO_Tot;
            int countPhukien_CC_XAU = countPhukien_CC_BTDUL_Xau + countPhukien_CC_BT_Xau + countPhukien_CC_SAT_Xau + countPhukien_CC_GO_Xau;
            int countPhukien_CDH_TOT = countPhukien_CDH_BTDUL_Tot + countPhukien_CDH_BT_Tot + countPhukien_CDH_SAT_Tot + countPhukien_CDH_GO_Tot + countPhukien_CDH_COMPOSITE_Tot;
            int countPhukien_CDH_XAU = countPhukien_CDH_BTDUL_Xau + countPhukien_CDH_BT_Xau + countPhukien_CDH_SAT_Xau + countPhukien_CDH_GO_Xau + countPhukien_CDH_COMPOSITE_Xau;

            string typePhukien = "";
            if (countPhukien_CC_TOT > 0 || countPhukien_CC_XAU > 0)
            {
                typePhukien = "Cóc cứng";

                SheetResponse.Cells[currentRow, 1].Value = currentRow - 2;
                SheetResponse.Cells[currentRow, 2].Value = values[1];
                SheetResponse.Cells[currentRow, 3].Value = values[2];
                SheetResponse.Cells[currentRow, 4].Value = values[5];
                if (!string.IsNullOrEmpty(values[6]))
                {
                    double.TryParse(values[6], out double lytrinh_from);
                    SheetResponse.Cells[currentRow, 5].Value = lytrinh_from;
                }
                if (!string.IsNullOrEmpty(values[7]))
                {
                    double.TryParse(values[7], out double lytrinh_to);
                    SheetResponse.Cells[currentRow, 6].Value = lytrinh_to;
                }

                SheetResponse.Cells[currentRow, 7].Value = typePhukien;
                SheetResponse.Cells[currentRow, 8].Value = countPhukien_CC_TOT;
                SheetResponse.Cells[currentRow, 9].Value = countPhukien_CC_XAU;

                currentRow++;
            }
            if (countPhukien_CDH_TOT > 0 || countPhukien_CDH_XAU > 0)
            {
                typePhukien = "Cóc đàn hồi";

                SheetResponse.Cells[currentRow, 1].Value = currentRow - 2;
                SheetResponse.Cells[currentRow, 2].Value = values[1];
                SheetResponse.Cells[currentRow, 3].Value = values[2];
                SheetResponse.Cells[currentRow, 4].Value = values[5];
                if (!string.IsNullOrEmpty(values[6]))
                {
                    double.TryParse(values[6], out double lytrinh_from);
                    SheetResponse.Cells[currentRow, 5].Value = lytrinh_from;
                }
                if (!string.IsNullOrEmpty(values[7]))
                {
                    double.TryParse(values[7], out double lytrinh_to);
                    SheetResponse.Cells[currentRow, 6].Value = lytrinh_to;
                }

                SheetResponse.Cells[currentRow, 7].Value = typePhukien;
                SheetResponse.Cells[currentRow, 8].Value = countPhukien_CDH_TOT;
                SheetResponse.Cells[currentRow, 9].Value = countPhukien_CDH_XAU;

                currentRow++;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Clear the rich text box
            richTextBox1.Clear();
        }
    }
}
