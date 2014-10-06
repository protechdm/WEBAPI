using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using CompareCloudware.Domain.Models;

using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;

using C = DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;

using System.IO;

namespace CompareCloudwareWebAPI.Helpers
{
    #region PACKAGE
    public class Package
    {
        public string Company { get; set; }
        public double Weight { get; set; }
        public long TrackingNumber { get; set; }
        public DateTime DateOrder { get; set; }
        public bool HasCompleted { get; set; }
    }
    #endregion

    #region CUSTOMSTYLESHEET
    public class CustomStylesheet : Stylesheet
    {
        public CustomStylesheet()
        {
            //var fonts = new Fonts();
            var fonts = new DocumentFormat.OpenXml.Spreadsheet.Fonts();
            var font = new DocumentFormat.OpenXml.Spreadsheet.Font();
            var fontName = new FontName {Val = StringValue.FromString("Arial")};
            var fontSize = new FontSize {Val = DoubleValue.FromDouble(11)};
            font.FontName = fontName;
            font.FontSize = fontSize;
            fonts.Append(font);
            //Font Index 1
            font = new DocumentFormat.OpenXml.Spreadsheet.Font();
            fontName = new FontName {Val = StringValue.FromString("Arial")};
            fontSize = new FontSize {Val = DoubleValue.FromDouble(12)};
            font.FontName = fontName;
            font.FontSize = fontSize;
            font.Bold = new Bold();
            fonts.Append(font);
            fonts.Count = UInt32Value.FromUInt32((uint)fonts.ChildElements.Count);
            var fills = new Fills();
            //var fill = new Fill();
            var fill = new DocumentFormat.OpenXml.Spreadsheet.Fill();
            //var patternFill = new PatternFill {PatternType = PatternValues.None};
            var patternFill = new DocumentFormat.OpenXml.Spreadsheet.PatternFill { PatternType = PatternValues.None };
            fill.PatternFill = patternFill;
            fills.Append(fill);
            //fill = new Fill();
            fill = new DocumentFormat.OpenXml.Spreadsheet.Fill();
            //patternFill = new PatternFill { PatternType = PatternValues.Gray125 };
            patternFill = new DocumentFormat.OpenXml.Spreadsheet.PatternFill { PatternType = PatternValues.Gray125 };
            fill.PatternFill = patternFill;
            fills.Append(fill);
            //Fill index  2
            //fill = new Fill();
            fill = new DocumentFormat.OpenXml.Spreadsheet.Fill();
            //patternFill = new PatternFill {PatternType = PatternValues.Solid, 
            //                               ForegroundColor = new ForegroundColor()};
            patternFill = new DocumentFormat.OpenXml.Spreadsheet.PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor()
            };
            patternFill.ForegroundColor = 
               TranslateForeground(System.Drawing.Color.LightBlue);
            //patternFill.BackgroundColor = 
            //    new BackgroundColor {Rgb = patternFill.ForegroundColor.Rgb};
            patternFill.BackgroundColor =
                new DocumentFormat.OpenXml.Spreadsheet.BackgroundColor { Rgb = patternFill.ForegroundColor.Rgb };
            fill.PatternFill = patternFill;
            fills.Append(fill);
            //Fill index  3
            //fill = new Fill();
            fill = new DocumentFormat.OpenXml.Spreadsheet.Fill();
            //patternFill = new PatternFill
            //{
            //    PatternType = PatternValues.Solid, 
            //                  ForegroundColor = new ForegroundColor()};
            patternFill = new DocumentFormat.OpenXml.Spreadsheet.PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor()
            };
            patternFill.ForegroundColor = 
               TranslateForeground(System.Drawing.Color.DodgerBlue);
            //patternFill.BackgroundColor = 
            //   new BackgroundColor {Rgb = patternFill.ForegroundColor.Rgb};
            patternFill.BackgroundColor =
               new DocumentFormat.OpenXml.Spreadsheet.BackgroundColor { Rgb = patternFill.ForegroundColor.Rgb };
            fill.PatternFill = patternFill;
            fills.Append(fill);
            fills.Count = UInt32Value.FromUInt32((uint)fills.ChildElements.Count);
            var borders = new Borders();
            //var border = new Border
            //            {
            //                LeftBorder = new LeftBorder(),
            //                RightBorder = new RightBorder(),
            //                TopBorder = new TopBorder(),
            //                BottomBorder = new BottomBorder(),
            //                DiagonalBorder = new DiagonalBorder()
            //            };
            var border = new Border
            {
                LeftBorder = new DocumentFormat.OpenXml.Spreadsheet.LeftBorder(),
                RightBorder = new DocumentFormat.OpenXml.Spreadsheet.RightBorder(),
                TopBorder = new DocumentFormat.OpenXml.Spreadsheet.TopBorder(),
                BottomBorder = new DocumentFormat.OpenXml.Spreadsheet.BottomBorder(),
                DiagonalBorder = new DocumentFormat.OpenXml.Spreadsheet.DiagonalBorder()
            };
            borders.Append(border);
            //All Boarder Index 1
            //border = new Border
            //             {
            //                 LeftBorder = new LeftBorder {Style = BorderStyleValues.Thin},
            //                 RightBorder = new RightBorder {Style = BorderStyleValues.Thin},
            //                 TopBorder = new TopBorder {Style = BorderStyleValues.Thin},
            //                 BottomBorder = new BottomBorder {Style = BorderStyleValues.Thin},
            //                 DiagonalBorder = new DiagonalBorder()
            //             };
            border = new Border
            {
                LeftBorder = new DocumentFormat.OpenXml.Spreadsheet.LeftBorder { Style = BorderStyleValues.Thin },
                RightBorder = new DocumentFormat.OpenXml.Spreadsheet.RightBorder { Style = BorderStyleValues.Thin },
                TopBorder = new DocumentFormat.OpenXml.Spreadsheet.TopBorder { Style = BorderStyleValues.Thin },
                BottomBorder = new DocumentFormat.OpenXml.Spreadsheet.BottomBorder { Style = BorderStyleValues.Thin },
                DiagonalBorder = new DiagonalBorder()
            };
            borders.Append(border);
            //Top and Bottom Boarder Index 2
            //border = new Border
            //{
            //    LeftBorder = new LeftBorder(),
            //    RightBorder = new RightBorder (),
            //    TopBorder = new TopBorder { Style = BorderStyleValues.Thin },
            //    BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin },
            //    DiagonalBorder = new DiagonalBorder()
            //};
            border = new Border
            {
                LeftBorder = new DocumentFormat.OpenXml.Spreadsheet.LeftBorder(),
                RightBorder = new DocumentFormat.OpenXml.Spreadsheet.RightBorder(),
                TopBorder = new DocumentFormat.OpenXml.Spreadsheet.TopBorder { Style = BorderStyleValues.Thin },
                BottomBorder = new DocumentFormat.OpenXml.Spreadsheet.BottomBorder { Style = BorderStyleValues.Thin },
                DiagonalBorder = new DiagonalBorder()
            };
            borders.Append(border);
            borders.Count = UInt32Value.FromUInt32((uint)borders.ChildElements.Count);
            var cellStyleFormats = new CellStyleFormats();
            var cellFormat = new CellFormat {NumberFormatId = 0, 
                                 FontId = 0, FillId = 0, BorderId = 0};
            cellStyleFormats.Append(cellFormat);
            cellStyleFormats.Count = 
               UInt32Value.FromUInt32((uint)cellStyleFormats.ChildElements.Count);
            uint iExcelIndex = 164;
            var numberingFormats = new NumberingFormats();
            var cellFormats = new CellFormats();
            cellFormat = new CellFormat {NumberFormatId = 0, FontId = 0, 
                             FillId = 0, BorderId = 0, FormatId = 0};
            cellFormats.Append(cellFormat);
            //var nformatDateTime = new NumberingFormat
            //         {
            //             NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
            //             FormatCode = StringValue.FromString("dd/mm/yyyy hh:mm:ss")
            //         };
            var nformatDateTime = new DocumentFormat.OpenXml.Spreadsheet.NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                FormatCode = StringValue.FromString("dd/mm/yyyy hh:mm:ss")
            };
            numberingFormats.Append(nformatDateTime);
            //var nformat4Decimal = new NumberingFormat
            //         {
            //             NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
            //             FormatCode = StringValue.FromString("#,##0.0000")
            //         };
            var nformat4Decimal = new DocumentFormat.OpenXml.Spreadsheet.NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                FormatCode = StringValue.FromString("#,##0.0000")
            };
            numberingFormats.Append(nformat4Decimal);
            //var nformat2Decimal = new NumberingFormat
            //          {
            //              NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
            //              FormatCode = StringValue.FromString("#,##0.00")
            //          };
            var nformat2Decimal = new DocumentFormat.OpenXml.Spreadsheet.NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                FormatCode = StringValue.FromString("#,##0.00")
            };
            numberingFormats.Append(nformat2Decimal);
            //var nformatForcedText = new NumberingFormat
            //           {
            //               NumberFormatId = UInt32Value.FromUInt32(iExcelIndex),
            //               FormatCode = StringValue.FromString("@")
            //           };
            var nformatForcedText = new DocumentFormat.OpenXml.Spreadsheet.NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex),
                FormatCode = StringValue.FromString("@")
            };
            numberingFormats.Append(nformatForcedText);
            // index 1
            // Cell Standard Date format 
            cellFormat = new CellFormat
                 {
                     NumberFormatId = 14,
                     FontId = 0,
                     FillId = 0,
                     BorderId = 0,
                     FormatId = 0,
                     ApplyNumberFormat = BooleanValue.FromBoolean(true)
                 };
            cellFormats.Append(cellFormat);
            // Index 2
            // Cell Standard Number format with 2 decimal placing
            cellFormat = new CellFormat
                 {
                     NumberFormatId = 4,
                     FontId = 0,
                     FillId = 0,
                     BorderId = 0,
                     FormatId = 0,
                     ApplyNumberFormat = BooleanValue.FromBoolean(true)
                 };
            cellFormats.Append(cellFormat);
            // Index 3
            // Cell Date time custom format
            cellFormat = new CellFormat
                 {
                     NumberFormatId = nformatDateTime.NumberFormatId,
                     FontId = 0,
                     FillId = 0,
                     BorderId = 0,
                     FormatId = 0,
                     ApplyNumberFormat = BooleanValue.FromBoolean(true)
                 };
            cellFormats.Append(cellFormat);
            // Index 4
            // Cell 4 decimal custom format
            cellFormat = new CellFormat
                 {
                     NumberFormatId = nformat4Decimal.NumberFormatId,
                     FontId = 0,
                     FillId = 0,
                     BorderId = 0,
                     FormatId = 0,
                     ApplyNumberFormat = BooleanValue.FromBoolean(true)
                 };
            cellFormats.Append(cellFormat);
            // Index 5
            // Cell 2 decimal custom format
            cellFormat = new CellFormat
                 {
                     NumberFormatId = nformat2Decimal.NumberFormatId,
                     FontId = 0,
                     FillId = 0,
                     BorderId = 0,
                     FormatId = 0,
                     ApplyNumberFormat = BooleanValue.FromBoolean(true)
                 };
            cellFormats.Append(cellFormat);
            // Index 6
            // Cell forced number text custom format
            cellFormat = new CellFormat
                 {
                     NumberFormatId = nformatForcedText.NumberFormatId,
                     FontId = 0,
                     FillId = 0,
                     BorderId = 0,
                     FormatId = 0,
                     ApplyNumberFormat = BooleanValue.FromBoolean(true)
                 };
            cellFormats.Append(cellFormat);
            // Index 7
            // Cell text with font 12 
            cellFormat = new CellFormat
                 {
                     NumberFormatId = nformatForcedText.NumberFormatId,
                     FontId = 1,
                     FillId = 0,
                     BorderId = 0,
                     FormatId = 0,
                     ApplyNumberFormat = BooleanValue.FromBoolean(true)
                 };
            cellFormats.Append(cellFormat);
            // Index 8
            // Cell text
            cellFormat = new CellFormat
                 {
                     NumberFormatId = nformatForcedText.NumberFormatId,
                     FontId = 0,
                     FillId = 0,
                     BorderId = 1,
                     FormatId = 0,
                     ApplyNumberFormat = BooleanValue.FromBoolean(true)
                 };
            cellFormats.Append(cellFormat);
            // Index 9
            // Coloured 2 decimal cell text
            cellFormat = new CellFormat
                     {
                         NumberFormatId = nformat2Decimal.NumberFormatId,
                         FontId = 0,
                         FillId = 2,
                         BorderId = 2,
                         FormatId = 0,
                         ApplyNumberFormat = BooleanValue.FromBoolean(true)
                     };
            cellFormats.Append(cellFormat);
            // Index 10
            // Coloured cell text
            cellFormat = new CellFormat
                     {
                         NumberFormatId = nformatForcedText.NumberFormatId,
                         FontId = 0,
                         FillId = 2,
                         BorderId = 2,
                         FormatId = 0,
                         ApplyNumberFormat = BooleanValue.FromBoolean(true)
                     };
            cellFormats.Append(cellFormat);
            // Index 11
            // Coloured cell text
            cellFormat = new CellFormat
                 {
                     NumberFormatId = nformatForcedText.NumberFormatId,
                     FontId = 1,
                     FillId = 3,
                     BorderId = 2,
                     FormatId = 0,
                     ApplyNumberFormat = BooleanValue.FromBoolean(true)
                 };
            cellFormats.Append(cellFormat);
            numberingFormats.Count = 
              UInt32Value.FromUInt32((uint)numberingFormats.ChildElements.Count);
            cellFormats.Count = UInt32Value.FromUInt32((uint)cellFormats.ChildElements.Count);
            this.Append(numberingFormats);
            this.Append(fonts);
            this.Append(fills);
            this.Append(borders);
            this.Append(cellStyleFormats);
            this.Append(cellFormats);
            var css = new CellStyles();
            var cs = new CellStyle {Name = StringValue.FromString("Normal"), 
                                    FormatId = 0, BuiltinId = 0};
            css.Append(cs);
            css.Count = UInt32Value.FromUInt32((uint)css.ChildElements.Count);
            this.Append(css);
            var dfs = new DifferentialFormats {Count = 0};
            this.Append(dfs);
            var tss = new TableStyles
                  {
                      Count = 0,
                      DefaultTableStyle = StringValue.FromString("TableStyleMedium9"),
                      DefaultPivotStyle = StringValue.FromString("PivotStyleLight16")
                  };
            this.Append(tss);
        }
        private static DocumentFormat.OpenXml.Spreadsheet.ForegroundColor TranslateForeground(System.Drawing.Color fillColor)
        {
            return new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor()
           {
               Rgb = new HexBinaryValue()
                     {
                         Value =
                             System.Drawing.ColorTranslator.ToHtml(
                             System.Drawing.Color.FromArgb(
                                 fillColor.A,
                                 fillColor.R,
                                 fillColor.G,
                                 fillColor.B)).Replace("#", "")
                     }
           };
        }
    }
    #endregion

    #region CUSTOMCOLUMN
    public class CustomColumn : Column
    {
        public CustomColumn(UInt32 startColumnIndex, 
               UInt32 endColumnIndex, double columnWidth)
        {
            this.Min = startColumnIndex;
            this.Max = endColumnIndex;
            this.Width = columnWidth;
            this.CustomWidth = true;
        }
    }
    #endregion

    #region TEXTCELL
    public class TextCell : Cell
    {
        public TextCell(string header, string text, int index)
        {
            this.DataType = CellValues.InlineString;
            this.CellReference = header + index;
            //Add text to the text cell.
            //this.InlineString = new InlineString { Text = new Text { Text = text } };
            this.InlineString = new InlineString { Text = new DocumentFormat.OpenXml.Spreadsheet.Text { Text = text } };
        }
    }
    public class NumberCell : Cell
    {
        public NumberCell(string header, string text, int index)
        {
            this.DataType = CellValues.Number;
            this.CellReference = header + index;
            this.CellValue = new CellValue(text);
        }
    }
    public class FormatedNumberCell : NumberCell
    {
        public FormatedNumberCell(string header, string text, int index)
            : base(header, text, index)
        {
            this.StyleIndex = 2;
        }
    }
    public class DateCell : Cell
    {
        public DateCell(string header, DateTime dateTime, int index)
        {
            this.DataType = CellValues.Date;
            this.CellReference = header + index;
            this.StyleIndex = 1;
            this.CellValue = new CellValue { Text = dateTime.ToOADate().ToString() }; ;
        }
    }
    public class FomulaCell : Cell
    {
        public FomulaCell(string header, string text, int index)
        {
            this.CellFormula = new CellFormula { CalculateCell = true, Text = text };
            this.DataType = CellValues.Number;
            this.CellReference = header + index;
            this.StyleIndex = 2;
        }
    }
    public class HeaderCell : TextCell
    {
        public HeaderCell(string header, string text, int index) : 
               base(header, text, index)
        {
            this.StyleIndex = 11;
        }
    }
    #endregion

    #region EXCELHELPER
    public class ExcelHelper
    {
        #region CREATE
        /// <summary>
        /// Write excel file of a list of object as T
        /// Assume that maximum of 24 columns 
        /// </summary>
        /// <typeparam name="T">Object type to pass in</typeparam>
        /// <param name="fileName">Full path of the file name of excel spreadsheet</param>
        /// <param name="objects">list of the object type</param>
        /// <param name="sheetName">Sheet names of Excel File</param>
        /// <param name="headerNames">Header names of the object</param>
        public void Create<T>(
            string fileName,
            List<T> objects,
            string sheetName,
            List<string> headerNames)
        {
            //Open the copied template workbook. 
            using (SpreadsheetDocument myWorkbook = 
                   SpreadsheetDocument.Create(fileName, 
                   SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = myWorkbook.AddWorkbookPart();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                // Create Styles and Insert into Workbook
                var stylesPart = 
                    myWorkbook.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                Stylesheet styles = new CustomStylesheet();
                styles.Save(stylesPart);
                string relId = workbookPart.GetIdOfPart(worksheetPart);
                var workbook = new Workbook();
                var fileVersion = 
                    new FileVersion { ApplicationName = 
                    "Microsoft Office Excel" };
                var worksheet = new Worksheet();
                int numCols = headerNames.Count;
                var columns = new Columns();
                for (int col = 0; col < numCols; col++)
                {
                    int width = headerNames[col].Length + 5;
                    Column c = new CustomColumn((UInt32)col + 1, 
                                  (UInt32)numCols + 1, width);
                    columns.Append(c);
                }
                worksheet.Append(columns);
                var sheets = new Sheets();
                var sheet = new Sheet { Name = sheetName, SheetId = 1, Id = relId };
                sheets.Append(sheet);
                workbook.Append(fileVersion);
                workbook.Append(sheets);
                SheetData sheetData = CreateSheetData(objects, headerNames);
                worksheet.Append(sheetData);
                worksheetPart.Worksheet = worksheet;
                worksheetPart.Worksheet.Save();
                myWorkbook.WorkbookPart.Workbook = workbook;
                myWorkbook.WorkbookPart.Workbook.Save();
                myWorkbook.Close();
            }
        }
        #endregion

        #region CREATESHEETDATA
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T">Object type to pass in</typeparam>
        /// <param name="objects">list of the object type</param>
        /// <param name="headerNames">Header names of the object</param>
        /// <returns></returns>
        private static SheetData CreateSheetData<T>(List<T> objects, 
                       List<string> headerNames)
        {
            var sheetData = new SheetData();
            if (objects != null)
            {
                //Get fields names of object
                List<string> fields = GetPropertyInfo<T>();
                //Get a list of A to Z
                var az = new List<Char>(Enumerable.Range('A', 'Z' - 
                                      'A' + 1).Select(i => (Char)i).ToArray());
                //A to E number of columns 
                List<Char> headers = az.GetRange(0, fields.Count);
                int numRows = objects.Count;
                int numCols = fields.Count;
                var header = new Row();
                int index = 1;
                header.RowIndex = (uint)index;
                for (int col = 0; col < numCols; col++)
                {
                    var c = new HeaderCell(headers[col].ToString(), 
                                           headerNames[col], index);
                    header.Append(c);
                }
                sheetData.Append(header);
                for (int i = 0; i < numRows; i++)
                {
                    index++;
                    var obj1 = objects[i];
                    var r = new Row { RowIndex = (uint)index };
                    for (int col = 0; col < numCols; col++)
                    {
                        string fieldName = fields[col];
                        PropertyInfo myf = obj1.GetType().GetProperty(fieldName);
                        if (myf != null)
                        {
                            object obj = myf.GetValue(obj1, null);
                            if (obj != null)
                            {
                                if (obj.GetType() == typeof(string))
                                {
                                    var c = new TextCell(headers[col].ToString(), 
                                                obj.ToString(), index);
                                    r.Append(c);
                                }
                                else if (obj.GetType() == typeof(bool))
                                {
                                    string value = 
                                      (bool)obj ? "Yes" : "No";
                                    var c = new TextCell(headers[col].ToString(), 
                                                         value, index);
                                    r.Append(c);
                                }
                                else if (obj.GetType() == typeof(DateTime))
                                {
                                    var c = new DateCell(headers[col].ToString(), 
                                               (DateTime)obj, index);
                                    r.Append(c);
                                }
                                else if (obj.GetType() == typeof(decimal) || 
                                         obj.GetType() == typeof(double))
                                {
                                    var c = new FormatedNumberCell(
                                                 headers[col].ToString(), 
                                                 obj.ToString(), index);
                                    r.Append(c);
                                }
                                else
                                {
                                    long value;
                                    if (long.TryParse(obj.ToString(), out value))
                                    {
                                        var c = new NumberCell(headers[col].ToString(), 
                                                    obj.ToString(), index);
                                        r.Append(c);
                                    }
                                    else
                                    {
                                        var c = new TextCell(headers[col].ToString(), 
                                                    obj.ToString(), index);
                                        r.Append(c);
                                    }
                                }
                            }
                        }
                    }
                    sheetData.Append(r);
                }
                index++;
                Row total = new Row();
                total.RowIndex = (uint)index;
                for (int col = 0; col < numCols; col++)
                {
                    var obj1 = objects[0];
                    string fieldName = fields[col];
                    PropertyInfo myf = obj1.GetType().GetProperty(fieldName);
                    if (myf != null)
                    {
                        object obj = myf.GetValue(obj1, null);
                        if (obj != null)
                        {
                            if (col == 0)
                            {
                                var c = new TextCell(headers[col].ToString(), 
                                                     "Total", index);
                                c.StyleIndex = 10;
                                total.Append(c);
                            }
                            else if (obj.GetType() == typeof(decimal) || 
                                     obj.GetType() == typeof(double))
                            {
                                string headerCol = headers[col].ToString();
                                string firstRow = headerCol + "2";
                                string lastRow = headerCol + (numRows + 1);
                                string formula = "=SUM(" + firstRow + " : " + lastRow + ")";
                                //Console.WriteLine(formula);
                                var c = new FomulaCell(headers[col].ToString(), 
                                                       formula, index);
                                c.StyleIndex = 9;
                                total.Append(c);
                            }
                            else
                            {
                                var c = new TextCell(headers[col].ToString(), 
                                                     string.Empty, index);
                                c.StyleIndex = 10;
                                total.Append(c);
                            }
                        }
                    }
                }
                sheetData.Append(total);
            }
            return sheetData;
        }
        #endregion

        #region GETPROPERTYINFO
        private static List<string> GetPropertyInfo<T>()
        {
            PropertyInfo[] propertyInfos = typeof(T).GetProperties();
            // write property names
            return propertyInfos.Select(propertyInfo => propertyInfo.Name).ToList();
        }
        #endregion

        #region CREATE ANALYTICS SUMMARY
        /// <summary>
        /// Write excel file of a list of object as T
        /// Assume that maximum of 24 columns 
        /// </summary>
        /// <typeparam name="T">Object type to pass in</typeparam>
        /// <param name="fileName">Full path of the file name of excel spreadsheet</param>
        /// <param name="objects">list of the object type</param>
        /// <param name="sheetName">Sheet names of Excel File</param>
        /// <param name="headerNames">Header names of the object</param>
        public void CreateAnalyticsSummary(
            string fileName,
            List<SiteAnalyticsVendorSummary> objects,
            string vendorName,
            DateTime startDate,
            DateTime endDate,
            string sheetName)
        {
            //Open the copied template workbook. 
            using (SpreadsheetDocument myWorkbook =
                   SpreadsheetDocument.Create(fileName,
                   SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = myWorkbook.AddWorkbookPart();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                // Create Styles and Insert into Workbook
                var stylesPart =
                    myWorkbook.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                Stylesheet styles = new CustomStylesheet();
                styles.Save(stylesPart);
                string relId = workbookPart.GetIdOfPart(worksheetPart);
                var workbook = new Workbook();
                var fileVersion =
                    new FileVersion
                    {
                        ApplicationName =
                            "Microsoft Office Excel"
                    };
                var worksheet = new Worksheet();
                //int numCols = headerNames.Count;
                //var columns = new Columns();
                //for (int col = 0; col < numCols; col++)
                //{
                //    int width = headerNames[col].Length + 5;
                //    Column c = new CustomColumn((UInt32)col + 1,
                //                  (UInt32)numCols + 1, width);
                //    columns.Append(c);
                //}
                //worksheet.Append(columns);
                var sheets = new Sheets();
                var sheet = new Sheet { Name = sheetName, SheetId = 1, Id = relId };
                sheets.Append(sheet);
                workbook.Append(fileVersion);
                workbook.Append(sheets);
                SheetData sheetData = CreateVendorAnalyticsSheetData(objects,vendorName,startDate,endDate);
                worksheet.Append(sheetData);

                Dictionary<string,int> chartValues = new Dictionary<string,int>();
                chartValues.Add("VALUE 1",100);
                chartValues.Add("VALUE 2",200);
                chartValues.Add("VALUE 3",300);


                worksheetPart.Worksheet = worksheet;
                worksheetPart.Worksheet.Save();
                myWorkbook.WorkbookPart.Workbook = workbook;
                myWorkbook.WorkbookPart.Workbook.Save();

                //InsertChartInSpreadsheet("", sheetName, "TITLE", chartValues, myWorkbook);
                InsertChartInSpreadsheetUsingSDK("", sheetName, "TITLE", chartValues, myWorkbook);
                

                myWorkbook.Close();
            }
        }
        #endregion

        #region CREATEVENDORANALYTICSSHEETDATA
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T">Object type to pass in</typeparam>
        /// <param name="objects">list of the object type</param>
        /// <param name="headerNames">Header names of the object</param>
        /// <returns></returns>
        private static SheetData CreateVendorAnalyticsSheetData(List<SiteAnalyticsVendorSummary> analytics,string vendorName,DateTime startDate,DateTime endDate)
        {
            var sheetData = new SheetData();
            NumberCell nc;
            if (analytics != null)
            {
                ////Get fields names of object
                //List<string> fields = GetPropertyInfo<T>();
                ////Get a list of A to Z
                //var az = new List<Char>(Enumerable.Range('A', 'Z' -
                //                      'A' + 1).Select(i => (Char)i).ToArray());
                ////A to E number of columns 
                //List<Char> headers = az.GetRange(0, fields.Count);
                //int numRows = objects.Count;
                //int numCols = fields.Count;
                //var header = new Row();
                int index = 1;
                //header.RowIndex = (uint)index;
                //for (int col = 0; col < numCols; col++)
                //{
                //    var c = new HeaderCell(headers[col].ToString(),
                //                           headerNames[col], index);
                //    header.Append(c);
                //}
                //sheetData.Append(header);
                var r = new Row { RowIndex = (uint)index };
                var c = new TextCell("A",
                            "Vendor name analytics summary", index);
                r.Append(c);
                sheetData.Append(r);

                #region IMPRESSIONS
                index +=2;
                r = new Row { RowIndex = (uint)index };
                c = new TextCell("A",
                            "Impressions", index);
                r.Append(c);
                sheetData.Append(r);

                index++;
                r = new Row { RowIndex = (uint)index };
                c = new TextCell("A",
                            "Total portfolio", index);
                r.Append(c);

                string totalPortfolioImpressions = analytics.Sum(x => x.Impressions).ToString();
                c = new TextCell("B",totalPortfolioImpressions, index);
                r.Append(c);
                c = new TextCell("C", "between", index);
                r.Append(c);
                c = new TextCell("D", startDate.ToShortDateString(), index);
                r.Append(c);
                c = new TextCell("E", "and", index);
                r.Append(c);
                c = new TextCell("F", endDate.ToShortDateString(), index);
                r.Append(c);

                sheetData.Append(r);


                index+=2;

                foreach (SiteAnalyticsVendorSummary savs in analytics)
                {
                    r = new Row { RowIndex = (uint)index };
                    c = new TextCell("A",
                                savs.ServiceName, index);
                    r.Append(c);

                    nc = new NumberCell("B", savs.Impressions.ToString(), index);
                    r.Append(nc);
                    c = new TextCell("C", "between", index);
                    r.Append(c);
                    c = new TextCell("D", startDate.ToShortDateString(), index);
                    r.Append(c);
                    c = new TextCell("E", "and", index);
                    r.Append(c);
                    c = new TextCell("F", endDate.ToShortDateString(), index);
                    r.Append(c);

                    sheetData.Append(r);
                    index++;
                }
                #endregion

                #region COMPARISON RESULTS IMPRESSIONS
                index += 2;
                r = new Row { RowIndex = (uint)index };
                c = new TextCell("A",
                            "Comparison results impressions", index);
                r.Append(c);
                sheetData.Append(r);

                index++;
                r = new Row { RowIndex = (uint)index };
                c = new TextCell("A",
                            "Total portfolio", index);
                r.Append(c);

                totalPortfolioImpressions = analytics.Sum(x => x.ComparisonResultImpressions).ToString();
                c = new TextCell("B", totalPortfolioImpressions, index);
                r.Append(c);
                c = new TextCell("C", "between", index);
                r.Append(c);
                c = new TextCell("D", startDate.ToShortDateString(), index);
                r.Append(c);
                c = new TextCell("E", "and", index);
                r.Append(c);
                c = new TextCell("F", endDate.ToShortDateString(), index);
                r.Append(c);

                sheetData.Append(r);


                index += 2;

                foreach (SiteAnalyticsVendorSummary savs in analytics)
                {
                    r = new Row { RowIndex = (uint)index };
                    c = new TextCell("A",
                                savs.ServiceName, index);
                    r.Append(c);

                    c = new TextCell("B", savs.ComparisonResultImpressions.ToString(), index);
                    r.Append(c);
                    c = new TextCell("C", "between", index);
                    r.Append(c);
                    c = new TextCell("D", startDate.ToShortDateString(), index);
                    r.Append(c);
                    c = new TextCell("E", "and", index);
                    r.Append(c);
                    c = new TextCell("F", endDate.ToShortDateString(), index);
                    r.Append(c);

                    sheetData.Append(r);
                    index++;
                }
                #endregion

                #region SHOP VISITS
                index += 2;
                r = new Row { RowIndex = (uint)index };
                c = new TextCell("A",
                            "Shop visits", index);
                r.Append(c);
                sheetData.Append(r);

                index++;
                r = new Row { RowIndex = (uint)index };
                c = new TextCell("A",
                            "Total portfolio", index);
                r.Append(c);

                totalPortfolioImpressions = analytics.Sum(x => x.ShopVisits).ToString();
                c = new TextCell("B", totalPortfolioImpressions, index);
                r.Append(c);
                c = new TextCell("C", "between", index);
                r.Append(c);
                c = new TextCell("D", startDate.ToShortDateString(), index);
                r.Append(c);
                c = new TextCell("E", "and", index);
                r.Append(c);
                c = new TextCell("F", endDate.ToShortDateString(), index);
                r.Append(c);

                sheetData.Append(r);


                index += 2;

                foreach (SiteAnalyticsVendorSummary savs in analytics)
                {
                    r = new Row { RowIndex = (uint)index };
                    c = new TextCell("A",
                                savs.ServiceName, index);
                    r.Append(c);

                    c = new TextCell("B", savs.ShopVisits.ToString(), index);
                    r.Append(c);
                    c = new TextCell("C", "between", index);
                    r.Append(c);
                    c = new TextCell("D", startDate.ToShortDateString(), index);
                    r.Append(c);
                    c = new TextCell("E", "and", index);
                    r.Append(c);
                    c = new TextCell("F", endDate.ToShortDateString(), index);
                    r.Append(c);

                    sheetData.Append(r);
                    index++;
                }
                #endregion

                #region SHOP CONTENT CONSUMPTION
                index += 2;
                r = new Row { RowIndex = (uint)index };
                c = new TextCell("A",
                            "Shop content consumption", index);
                r.Append(c);
                sheetData.Append(r);

                index++;
                r = new Row { RowIndex = (uint)index };
                c = new TextCell("A",
                            "Total portfolio", index);
                r.Append(c);

                totalPortfolioImpressions = analytics.Sum(x => x.ShopContentConsumption).ToString();
                c = new TextCell("B", totalPortfolioImpressions, index);
                r.Append(c);
                c = new TextCell("C", "between", index);
                r.Append(c);
                c = new TextCell("D", startDate.ToShortDateString(), index);
                r.Append(c);
                c = new TextCell("E", "and", index);
                r.Append(c);
                c = new TextCell("F", endDate.ToShortDateString(), index);
                r.Append(c);

                sheetData.Append(r);


                index += 2;

                foreach (SiteAnalyticsVendorSummary savs in analytics)
                {
                    r = new Row { RowIndex = (uint)index };
                    c = new TextCell("A",
                                savs.ServiceName, index);
                    r.Append(c);

                    c = new TextCell("B", savs.ShopContentConsumption.ToString(), index);
                    r.Append(c);
                    c = new TextCell("C", "between", index);
                    r.Append(c);
                    c = new TextCell("D", startDate.ToShortDateString(), index);
                    r.Append(c);
                    c = new TextCell("E", "and", index);
                    r.Append(c);
                    c = new TextCell("F", endDate.ToShortDateString(), index);
                    r.Append(c);

                    sheetData.Append(r);
                    index++;
                }
                #endregion

                #region SHOP LEADS (TRY/BUY REQUESTS - NOT EMAILS)
                index += 2;
                r = new Row { RowIndex = (uint)index };
                c = new TextCell("A",
                            "Shop leads (Try/Buy requests - not emails)", index);
                r.Append(c);
                sheetData.Append(r);

                index++;
                r = new Row { RowIndex = (uint)index };
                c = new TextCell("A",
                            "Total portfolio", index);
                r.Append(c);

                totalPortfolioImpressions = analytics.Sum(x => x.ShopLeads).ToString();
                c = new TextCell("B", totalPortfolioImpressions, index);
                r.Append(c);
                c = new TextCell("C", "between", index);
                r.Append(c);
                c = new TextCell("D", startDate.ToShortDateString(), index);
                r.Append(c);
                c = new TextCell("E", "and", index);
                r.Append(c);
                c = new TextCell("F", endDate.ToShortDateString(), index);
                r.Append(c);

                sheetData.Append(r);


                index += 2;

                foreach (SiteAnalyticsVendorSummary savs in analytics)
                {
                    r = new Row { RowIndex = (uint)index };
                    c = new TextCell("A",
                                savs.ServiceName, index);
                    r.Append(c);

                    c = new TextCell("B", savs.ShopLeads.ToString(), index);
                    r.Append(c);
                    c = new TextCell("C", "between", index);
                    r.Append(c);
                    c = new TextCell("D", startDate.ToShortDateString(), index);
                    r.Append(c);
                    c = new TextCell("E", "and", index);
                    r.Append(c);
                    c = new TextCell("F", endDate.ToShortDateString(), index);
                    r.Append(c);

                    sheetData.Append(r);
                    index++;
                }
                #endregion


                #region CRAP
                //for (int i = 0; i < numRows; i++)
                //{
                //    index++;
                //    var obj1 = objects[i];
                //    var r = new Row { RowIndex = (uint)index };
                //    for (int col = 0; col < numCols; col++)
                //    {
                //        string fieldName = fields[col];
                //        PropertyInfo myf = obj1.GetType().GetProperty(fieldName);
                //        if (myf != null)
                //        {
                //            object obj = myf.GetValue(obj1, null);
                //            if (obj != null)
                //            {
                //                #region STRING
                //                if (obj.GetType() == typeof(string))
                //                {
                //                    var c = new TextCell(headers[col].ToString(),
                //                                obj.ToString(), index);
                //                    r.Append(c);
                //                }
                //                #endregion
                //                #region BOOL
                //                else if (obj.GetType() == typeof(bool))
                //                {
                //                    string value =
                //                      (bool)obj ? "Yes" : "No";
                //                    var c = new TextCell(headers[col].ToString(),
                //                                         value, index);
                //                    r.Append(c);
                //                }
                //                #endregion
                //                #region DATETIME
                //                else if (obj.GetType() == typeof(DateTime))
                //                {
                //                    var c = new DateCell(headers[col].ToString(),
                //                               (DateTime)obj, index);
                //                    r.Append(c);
                //                }
                //                #endregion
                //                #region DECIMAL/DOUBLE
                //                else if (obj.GetType() == typeof(decimal) ||
                //                         obj.GetType() == typeof(double))
                //                {
                //                    var c = new FormatedNumberCell(
                //                                 headers[col].ToString(),
                //                                 obj.ToString(), index);
                //                    r.Append(c);
                //                }
                //                #endregion
                //                #region LONG
                //                else
                //                {
                //                    long value;
                //                    if (long.TryParse(obj.ToString(), out value))
                //                    {
                //                        var c = new NumberCell(headers[col].ToString(),
                //                                    obj.ToString(), index);
                //                        r.Append(c);
                //                    }
                //                    else
                //                    {
                //                        var c = new TextCell(headers[col].ToString(),
                //                                    obj.ToString(), index);
                //                        r.Append(c);
                //                    }
                //                }
                //                #endregion
                //            }
                //        }
                //    }
                //    sheetData.Append(r);
                //}
                #endregion
                index++;
                Row total = new Row();
                total.RowIndex = (uint)index;
                #region TOTALS
                //for (int col = 0; col < numCols; col++)
                //{
                //    var obj1 = objects[0];
                //    string fieldName = fields[col];
                //    PropertyInfo myf = obj1.GetType().GetProperty(fieldName);
                //    if (myf != null)
                //    {
                //        object obj = myf.GetValue(obj1, null);
                //        if (obj != null)
                //        {
                //            if (col == 0)
                //            {
                //                var c = new TextCell(headers[col].ToString(),
                //                                     "Total", index);
                //                c.StyleIndex = 10;
                //                total.Append(c);
                //            }
                //            else if (obj.GetType() == typeof(decimal) ||
                //                     obj.GetType() == typeof(double))
                //            {
                //                string headerCol = headers[col].ToString();
                //                string firstRow = headerCol + "2";
                //                string lastRow = headerCol + (numRows + 1);
                //                string formula = "=SUM(" + firstRow + " : " + lastRow + ")";
                //                //Console.WriteLine(formula);
                //                var c = new FomulaCell(headers[col].ToString(),
                //                                       formula, index);
                //                c.StyleIndex = 9;
                //                total.Append(c);
                //            }
                //            else
                //            {
                //                var c = new TextCell(headers[col].ToString(),
                //                                     string.Empty, index);
                //                c.StyleIndex = 10;
                //                total.Append(c);
                //            }
                //        }
                //    }
                //}
                #endregion
                sheetData.Append(total);
            }
            return sheetData;
        }
        #endregion

        #region INSERTCHARTINTOSPREADSHEET
        // Given a document name, a worksheet name, a chart title, and a Dictionary collection of text keys
        // and corresponding integer data, creates a column chart with the text as the series and the integers as the values.
        private static void InsertChartInSpreadsheet(string docName, string worksheetName, string title,
        Dictionary<string, int> data, SpreadsheetDocument document)
        {
            // Open the document for editing.
            //using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
            {
                IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().
        Where(s => s.Name == worksheetName);
                if (sheets.Count() == 0)
                {
                    // The specified worksheet does not exist.
                    return;
                }
                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);

                // Add a new drawing to the worksheet.
                DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                worksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing() { Id = worksheetPart.GetIdOfPart(drawingsPart) });
                worksheetPart.Worksheet.Save();

                // Add a new chart and set the chart language to English-US.
                ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();

                GenerateChartPart1Content(chartPart);
                return;

                chartPart.ChartSpace = new ChartSpace();
                chartPart.ChartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });
                DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartPart.ChartSpace.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>(
                    new DocumentFormat.OpenXml.Drawing.Charts.Chart());

                // Create a new clustered column chart.
                PlotArea plotArea = chart.AppendChild<PlotArea>(new PlotArea());
                Layout layout = plotArea.AppendChild<Layout>(new Layout());
                //BarChart barChart = plotArea.AppendChild<BarChart>(new BarChart(new BarDirection() { Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column) },
                //    new BarGrouping() { Val = new EnumValue<BarGroupingValues>(BarGroupingValues.Clustered) }));
                //PieChart barChart = plotArea.AppendChild<PieChart>(new PieChart(new BarDirection() { Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column) },
                //    new BarGrouping() { Val = new EnumValue<BarGroupingValues>(BarGroupingValues.Clustered) }));
                PieChart barChart = plotArea.AppendChild<PieChart>(new PieChart());

                uint i = 0;

                // Iterate through each key in the Dictionary collection and add the key to the chart Series
                // and add the corresponding value to the chart Values.
                foreach (string key in data.Keys)
                {
                    //BarChartSeries barChartSeries = barChart.AppendChild<BarChartSeries>(new BarChartSeries(new Index()
                    //{
                    //    Val =
                    //        new UInt32Value(i)
                    //},
                    //    new Order() { Val = new UInt32Value(i) },
                    //    new SeriesText(new NumericValue() { Text = key })));

                    PieChartSeries barChartSeries = barChart.AppendChild<PieChartSeries>(new PieChartSeries(new Index()
                    {
                        Val =
                            new UInt32Value(i)
                    },
                        new Order() { Val = new UInt32Value(i) },
                        new SeriesText(new NumericValue() { Text = key })));

                    StringLiteral strLit = barChartSeries.AppendChild<CategoryAxisData>(new CategoryAxisData()).AppendChild<StringLiteral>(new StringLiteral());
                    strLit.Append(new PointCount() { Val = new UInt32Value(1U) });
                    strLit.AppendChild<StringPoint>(new StringPoint() { Index = new UInt32Value(0U) }).Append(new NumericValue(title));

                    NumberLiteral numLit = barChartSeries.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Values>(
                        new DocumentFormat.OpenXml.Drawing.Charts.Values()).AppendChild<NumberLiteral>(new NumberLiteral());
                    numLit.Append(new FormatCode("General"));
                    numLit.Append(new PointCount() { Val = new UInt32Value(1U) });
                    numLit.AppendChild<NumericPoint>(new NumericPoint() { Index = new UInt32Value(0u) }).Append
        (new NumericValue(data[key].ToString()));

                    i++;
                }

                barChart.Append(new AxisId() { Val = new UInt32Value(48650112u) });
                barChart.Append(new AxisId() { Val = new UInt32Value(48672768u) });

                #region CATEGORYAXIS
                // Add the Category Axis.
                CategoryAxis catAx = plotArea.AppendChild<CategoryAxis>(new CategoryAxis(new AxisId() { Val = new UInt32Value(48650112u) }, new Scaling(new Orientation()
                {
                    Val = new EnumValue<DocumentFormat.
                        OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
                }),
                    new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
                    new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                    new CrossingAxis() { Val = new UInt32Value(48672768U) },
                    new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                    new AutoLabeled() { Val = new BooleanValue(true) },
                    new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) },
                    new LabelOffset() { Val = new UInt16Value((ushort)100) }));
                #endregion

                #region VALUEAXIS
                // Add the Value Axis.
                ValueAxis valAx = plotArea.AppendChild<ValueAxis>(new ValueAxis(new AxisId() { Val = new UInt32Value(48672768u) },
                    new Scaling(new Orientation()
                    {
                        Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(
                            DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
                    }),
                    new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
                    new MajorGridlines(),
                    new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat()
                    {
                        FormatCode = new StringValue("General"),
                        SourceLinked = new BooleanValue(true)
                    }, new TickLabelPosition()
                    {
                        Val = new EnumValue<TickLabelPositionValues>
                            (TickLabelPositionValues.NextTo)
                    }, new CrossingAxis() { Val = new UInt32Value(48650112U) },
                    new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                    new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }));
                #endregion

                #region LEGEND
                // Add the chart Legend.
                Legend legend = chart.AppendChild<Legend>(new Legend(new LegendPosition() { Val = new EnumValue<LegendPositionValues>(LegendPositionValues.Right) },
                    new Layout()));

                chart.Append(new PlotVisibleOnly() { Val = new BooleanValue(true) });
                #endregion

                #region NEW PIE
        //        // Create a new clustered column chart.
        //        PlotArea plotArea2 = chart.AppendChild<PlotArea>(new PlotArea());
        //        Layout layout2 = plotArea2.AppendChild<Layout>(new Layout());
        //        PieChart pieChart = plotArea.AppendChild<PieChart>(new PieChart(new BarDirection() { Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column) },
        //            new BarGrouping() { Val = new EnumValue<BarGroupingValues>(BarGroupingValues.Clustered) }));

        //        // Iterate through each key in the Dictionary collection and add the key to the chart Series
        //        // and add the corresponding value to the chart Values.
        //        foreach (string key in data.Keys)
        //        {
        //            PieChartSeries pieChartSeries = pieChart.AppendChild<PieChartSeries>(new PieChartSeries(new Index()
        //            {
        //                Val =
        //                    new UInt32Value(i)
        //            },
        //                new Order() { Val = new UInt32Value(i) },
        //                new SeriesText(new NumericValue() { Text = key })));

        //            StringLiteral strLit = pieChartSeries.AppendChild<CategoryAxisData>(new CategoryAxisData()).AppendChild<StringLiteral>(new StringLiteral());
        //            strLit.Append(new PointCount() { Val = new UInt32Value(1U) });
        //            strLit.AppendChild<StringPoint>(new StringPoint() { Index = new UInt32Value(0U) }).Append(new NumericValue(title));

        //            NumberLiteral numLit = pieChartSeries.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Values>(
        //                new DocumentFormat.OpenXml.Drawing.Charts.Values()).AppendChild<NumberLiteral>(new NumberLiteral());
        //            numLit.Append(new FormatCode("General"));
        //            numLit.Append(new PointCount() { Val = new UInt32Value(1U) });
        //            numLit.AppendChild<NumericPoint>(new NumericPoint() { Index = new UInt32Value(0u) }).Append
        //(new NumericValue(data[key].ToString()));

        //            i++;
        //        }
                #endregion

                // Save the chart part.
                chartPart.ChartSpace.Save();

                // Position the chart on the worksheet using a TwoCellAnchor object.
                drawingsPart.WorksheetDrawing = new WorksheetDrawing();
                TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild<TwoCellAnchor>(new TwoCellAnchor());
                twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(new ColumnId("9"),
                    new ColumnOffset("581025"),
                    new RowId("17"),
                    new RowOffset("114300")));
                twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(new ColumnId("17"),
                    new ColumnOffset("276225"),
                    new RowId("32"),
                    new RowOffset("0")));

                #region GRAPHICFRAME
                // Append a GraphicFrame to the TwoCellAnchor object.
                DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame =
                    twoCellAnchor.AppendChild<DocumentFormat.OpenXml.
        Drawing.Spreadsheet.GraphicFrame>(new DocumentFormat.OpenXml.Drawing.
        Spreadsheet.GraphicFrame());
                graphicFrame.Macro = "";

                graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties() { Id = new UInt32Value(2u), Name = "Chart 1" },
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()));

                graphicFrame.Append(new Transform(new Offset() { X = 0L, Y = 0L },
                                                                        new Extents() { Cx = 0L, Cy = 0L }));

                graphicFrame.Append(new Graphic(new GraphicData(new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) }) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }));

                twoCellAnchor.Append(new ClientData());
                #endregion

                // Save the WorksheetDrawing object.
                drawingsPart.WorksheetDrawing.Save();



            }
        }
        #endregion

        #region INSERTCHARTINTOSPREADSHEETUSINGSDK
        // Given a document name, a worksheet name, a chart title, and a Dictionary collection of text keys
        // and corresponding integer data, creates a column chart with the text as the series and the integers as the values.
        private static void InsertChartInSpreadsheetUsingSDK(string docName, string worksheetName, string title,
        Dictionary<string, int> data, SpreadsheetDocument document)
        {
            // Open the document for editing.
            //using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
            {
                IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().
        Where(s => s.Name == worksheetName);
                if (sheets.Count() == 0)
                {
                    // The specified worksheet does not exist.
                    return;
                }
                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);

                // Add a new drawing to the worksheet.
                DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                worksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing() { Id = worksheetPart.GetIdOfPart(drawingsPart) });
                worksheetPart.Worksheet.Save();

                // Add a new chart and set the chart language to English-US.
                ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();

                GenerateChartPart1Content(chartPart);


                // Save the chart part.
                chartPart.ChartSpace.Save();


                // Position the chart on the worksheet using a TwoCellAnchor object.
                drawingsPart.WorksheetDrawing = new WorksheetDrawing();
                TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild<TwoCellAnchor>(new TwoCellAnchor());
                twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(new ColumnId("9"),
                    new ColumnOffset("581025"),
                    new RowId("17"),
                    new RowOffset("114300")));
                twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(new ColumnId("17"),
                    new ColumnOffset("276225"),
                    new RowId("32"),
                    new RowOffset("0")));

                // Save the WorksheetDrawing object.
                drawingsPart.WorksheetDrawing.Save();



            }
        }
        #endregion

        #region GenerateChartPart1Content
        // Generates content of chartPart1.
        private static void GenerateChartPart1Content(ChartPart chartPart1)
        {
            C.ChartSpace chartSpace1 = new C.ChartSpace();
            chartSpace1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartSpace1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            chartSpace1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            C.EditingLanguage editingLanguage1 = new C.EditingLanguage() { Val = "en-US" };

            C.Chart chart1 = new C.Chart();

            C.PlotArea plotArea1 = new C.PlotArea();
            C.Layout layout1 = new C.Layout();

            C.PieChart pieChart1 = new C.PieChart();
            C.VaryColors varyColors1 = new C.VaryColors() { Val = true };

            C.PieChartSeries pieChartSeries1 = new C.PieChartSeries();
            C.Index index1 = new C.Index() { Val = (UInt32Value)0U };
            C.Order order1 = new C.Order() { Val = (UInt32Value)0U };

            C.Values values1 = new C.Values();

            C.NumberReference numberReference1 = new C.NumberReference();
            C.Formula formula1 = new C.Formula();
            formula1.Text = "\'Vendor Analytics\'!$B$6:$B$13";

            C.NumberingCache numberingCache1 = new C.NumberingCache();
            C.FormatCode formatCode1 = new C.FormatCode();
            formatCode1.Text = "General";
            C.PointCount pointCount1 = new C.PointCount() { Val = (UInt32Value)8U };

            C.NumericPoint numericPoint1 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue1 = new C.NumericValue();
            numericValue1.Text = "853";

            numericPoint1.Append(numericValue1);

            C.NumericPoint numericPoint2 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue2 = new C.NumericValue();
            numericValue2.Text = "853";

            numericPoint2.Append(numericValue2);

            C.NumericPoint numericPoint3 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue3 = new C.NumericValue();
            numericValue3.Text = "829";

            numericPoint3.Append(numericValue3);

            C.NumericPoint numericPoint4 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue4 = new C.NumericValue();
            numericValue4.Text = "844";

            numericPoint4.Append(numericValue4);

            C.NumericPoint numericPoint5 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue5 = new C.NumericValue();
            numericValue5.Text = "829";

            numericPoint5.Append(numericValue5);

            C.NumericPoint numericPoint6 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue6 = new C.NumericValue();
            numericValue6.Text = "858";

            numericPoint6.Append(numericValue6);

            C.NumericPoint numericPoint7 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue7 = new C.NumericValue();
            numericValue7.Text = "845";

            numericPoint7.Append(numericValue7);

            C.NumericPoint numericPoint8 = new C.NumericPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue8 = new C.NumericValue();
            numericValue8.Text = "858";

            numericPoint8.Append(numericValue8);

            numberingCache1.Append(formatCode1);
            numberingCache1.Append(pointCount1);
            numberingCache1.Append(numericPoint1);
            numberingCache1.Append(numericPoint2);
            numberingCache1.Append(numericPoint3);
            numberingCache1.Append(numericPoint4);
            numberingCache1.Append(numericPoint5);
            numberingCache1.Append(numericPoint6);
            numberingCache1.Append(numericPoint7);
            numberingCache1.Append(numericPoint8);

            numberReference1.Append(formula1);
            numberReference1.Append(numberingCache1);

            values1.Append(numberReference1);

            pieChartSeries1.Append(index1);
            pieChartSeries1.Append(order1);
            pieChartSeries1.Append(values1);
            C.FirstSliceAngle firstSliceAngle1 = new C.FirstSliceAngle() { Val = (UInt16Value)0U };

            pieChart1.Append(varyColors1);
            pieChart1.Append(pieChartSeries1);
            pieChart1.Append(firstSliceAngle1);

            plotArea1.Append(layout1);
            plotArea1.Append(pieChart1);

            C.Legend legend1 = new C.Legend();
            C.LegendPosition legendPosition1 = new C.LegendPosition() { Val = C.LegendPositionValues.Right };
            C.Layout layout2 = new C.Layout();

            C.TextProperties textProperties1 = new C.TextProperties();
            A.BodyProperties bodyProperties1 = new A.BodyProperties();
            A.ListStyle listStyle1 = new A.ListStyle();

            A.Paragraph paragraph1 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties() { RightToLeft = false };
            A.DefaultRunProperties defaultRunProperties1 = new A.DefaultRunProperties();

            paragraphProperties1.Append(defaultRunProperties1);
            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(endParagraphRunProperties1);

            textProperties1.Append(bodyProperties1);
            textProperties1.Append(listStyle1);
            textProperties1.Append(paragraph1);

            legend1.Append(legendPosition1);
            legend1.Append(layout2);
            legend1.Append(textProperties1);
            C.PlotVisibleOnly plotVisibleOnly1 = new C.PlotVisibleOnly() { Val = true };

            chart1.Append(plotArea1);
            chart1.Append(legend1);
            chart1.Append(plotVisibleOnly1);

            C.PrintSettings printSettings1 = new C.PrintSettings();
            C.HeaderFooter headerFooter1 = new C.HeaderFooter();
            C.PageMargins pageMargins2 = new C.PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            C.PageSetup pageSetup1 = new C.PageSetup();

            printSettings1.Append(headerFooter1);
            printSettings1.Append(pageMargins2);
            printSettings1.Append(pageSetup1);

            chartSpace1.Append(editingLanguage1);
            chartSpace1.Append(chart1);
            chartSpace1.Append(printSettings1);

            chartPart1.ChartSpace = chartSpace1;
        }
        #endregion

        
    }
    #endregion

    #region CRAP
    public class ExcelCreate
    {
        #region Main
        public void Main()
        {

            try
            {
                #region DUMMY DATA

                List<Package> packages =
                    new List<Package>
                            { new Package { Company = "Coho Vineyard", Weight = 25.2, 
                                  TrackingNumber = 89453312L, 
                                  DateOrder = DateTime.Today, HasCompleted = false },
                              new Package { Company = "Lucerne Publishing", Weight = 18.7, 
                                  TrackingNumber = 89112755L, 
                                  DateOrder = DateTime.Today, HasCompleted = false },
                              new Package { Company = "Wingtip Toys", Weight = 6.0, 
                                  TrackingNumber = 299456122L, 
                                  DateOrder = DateTime.Today, HasCompleted = false },
                              new Package { Company = "Adventure Works", Weight = 33.8, 
                                  TrackingNumber = 4665518773L, 
                                  DateOrder =  DateTime.Today.AddDays(-4), 
                                  HasCompleted = true },
                              new Package { Company = "Test Works", Weight = 35.8, 
                                  TrackingNumber = 4665518774L, 
                                  DateOrder =  DateTime.Today.AddDays(-2), 
                                  HasCompleted = true },
                              new Package { Company = "Good Works", Weight = 48.8, 
                                  TrackingNumber = 4665518775L, 
                                  DateOrder =  DateTime.Today.AddDays(-1), HasCompleted = true },

                            };
                #endregion
                List<string> headerNames =
                   new List<string> { "Company", 
                       "Weight", "Tracking Number", 
                       "Date Order", "Completed" };

                ExcelHelper excelFacade = new ExcelHelper();
                excelFacade.Create<Package>(@"H:\temp\output1.xlsx",
                            packages, "Packages", headerNames);

                Console.WriteLine("Completed");
            }
            catch (Exception e)
            {
                System.Diagnostics.Debugger.Break();
            }

        }
        #endregion

        #region CreateVendorAnalyticsSummary
        public void CreateVendorAnalyticsSummary(List<SiteAnalyticsVendorSummary> siteAnalytics, string vendorName, DateTime startDate, DateTime endDate)
        {

            try
            {
                GeneratedExcel ge = new GeneratedExcel();
                string fileName = vendorName + "_" + 
                    startDate.Day.ToString() + "-" + startDate.Month.ToString() + "-" + startDate.Year.ToString() +
                    "_to_" + 
                    endDate.Day.ToString() + "-" + endDate.Month.ToString() + "-" + endDate.Year.ToString() 
                    ;
                //ge.CreatePackage(@"h:\temp\generated.xlsx", siteAnalytics, vendorName, startDate, endDate);
                ge.CreatePackage(fileName, siteAnalytics, vendorName, startDate, endDate);
                return;
                #region DUMMY DATA

                //List<Package> packages =
                //    new List<Package>
                //            { new Package { Company = "Coho Vineyard", Weight = 25.2, 
                //                  TrackingNumber = 89453312L, 
                //                  DateOrder = DateTime.Today, HasCompleted = false },
                //              new Package { Company = "Lucerne Publishing", Weight = 18.7, 
                //                  TrackingNumber = 89112755L, 
                //                  DateOrder = DateTime.Today, HasCompleted = false },
                //              new Package { Company = "Wingtip Toys", Weight = 6.0, 
                //                  TrackingNumber = 299456122L, 
                //                  DateOrder = DateTime.Today, HasCompleted = false },
                //              new Package { Company = "Adventure Works", Weight = 33.8, 
                //                  TrackingNumber = 4665518773L, 
                //                  DateOrder =  DateTime.Today.AddDays(-4), 
                //                  HasCompleted = true },
                //              new Package { Company = "Test Works", Weight = 35.8, 
                //                  TrackingNumber = 4665518774L, 
                //                  DateOrder =  DateTime.Today.AddDays(-2), 
                //                  HasCompleted = true },
                //              new Package { Company = "Good Works", Weight = 48.8, 
                //                  TrackingNumber = 4665518775L, 
                //                  DateOrder =  DateTime.Today.AddDays(-1), HasCompleted = true },

                //            };
                #endregion
                //List<string> headerNames =
                //   new List<string> { "Company", 
                //       "Weight", "Tracking Number", 
                //       "Date Order", "Completed" };

                ExcelHelper excelFacade = new ExcelHelper();
                excelFacade.CreateAnalyticsSummary(@"H:\temp\output1.xlsx",
                            siteAnalytics, vendorName, startDate, endDate, "Vendor Analytics");

                Console.WriteLine("Completed");
            }
            catch (Exception e)
            {
                System.Diagnostics.Debugger.Break();
            }

        }
        #endregion

        #region CreateVendorAnalyticsSummaryAsStream
        public MemoryStream CreateVendorAnalyticsSummaryAsStream(List<SiteAnalyticsVendorSummary> siteAnalytics, string vendorName, DateTime startDate, DateTime endDate)
        {

            try
            {
                GeneratedExcel ge = new GeneratedExcel();
                string fileName = vendorName + "_" +
                    startDate.Day.ToString() + "-" + startDate.Month.ToString() + "-" + startDate.Year.ToString() +
                    "_to_" +
                    endDate.Day.ToString() + "-" + endDate.Month.ToString() + "-" + endDate.Year.ToString()
                    ;
                //ge.CreatePackage(@"h:\temp\generated.xlsx", siteAnalytics, vendorName, startDate, endDate);
                MemoryStream ms = ge.CreatePackageAsStream(fileName, siteAnalytics, vendorName, startDate, endDate);
                return ms;
            }
            catch (Exception e)
            {
                System.Diagnostics.Debugger.Break();
                return null;
            }

        }
        #endregion

    }
#endregion


}