using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using CompareCloudware.Domain.Models;
using System.IO;

namespace CompareCloudwareWebAPI.Helpers
{
    public class GeneratedExcel
    {
        string impressionsLabelRange = null;
        string impressionsDataRange = null;
        string comparisonResultsImpressionsLabelRange = null;
        string comparisonResultsImpressionsDataRange = null;
        string shopVisitsLabelRange = null;
        string shopVisitsDataRange = null;
        string shopContentConsumptionLabelRange = null;
        string shopContentConsumptionDataRange = null;
        string shopLeadsLabelRange = null;
        string shopLeadsDataRange = null;
        int sectionHeight = 10;
        int sectionSpacing = 3;
        int graphHeight = 10;
        // Creates a SpreadsheetDocument.
        public void CreatePackage(string filePath, List<SiteAnalyticsVendorSummary> analytics, string vendorName, DateTime startDate, DateTime endDate)
        {
            using(SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                CreateParts(package,analytics,vendorName,startDate,endDate);

                
                package.Close();
            }
        }

        public MemoryStream CreatePackageAsStream(string filePath, List<SiteAnalyticsVendorSummary> analytics, string vendorName, DateTime startDate, DateTime endDate)
        {
            MemoryStream ms = new MemoryStream();
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
            {
                CreateParts(package, analytics, vendorName, startDate, endDate);


                package.Close();
            }
            return ms;
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(SpreadsheetDocument document, List<SiteAnalyticsVendorSummary> analytics,string vendorName,DateTime startDate,DateTime endDate)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPart1Content(workbookPart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateWorkbookStylesPart1Content(workbookStylesPart1);

            //ThemePart themePart1 = workbookPart1.AddNewPart<ThemePart>("rId2");
            //GenerateThemePart1Content(themePart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetPart1Content(worksheetPart1,analytics,vendorName,startDate,endDate);

            DrawingsPart drawingsPart1 = worksheetPart1.AddNewPart<DrawingsPart>("rId2");
            GenerateDrawingsPart1Content(drawingsPart1);

            ChartPart chartPart2 = drawingsPart1.AddNewPart<ChartPart>("rId1");
            GenerateImpressionsPieChartContent(chartPart2);

            ChartPart chartPart1 = drawingsPart1.AddNewPart<ChartPart>("rId2");
            GenerateComparisonResultsImpressionsPieChartContent(chartPart1);

            ChartPart chartPart3 = drawingsPart1.AddNewPart<ChartPart>("rId3");
            GenerateShopVisitsPieChartContent(chartPart3);
            ChartPart chartPart4 = drawingsPart1.AddNewPart<ChartPart>("rId4");
            GenerateShopContentConsumptionPieChartContent(chartPart4);
            ChartPart chartPart5 = drawingsPart1.AddNewPart<ChartPart>("rId5");
            GenerateShopLeadsPieChartContent(chartPart5);


            //SharedStringTablePart sharedStringTablePart1 = workbookPart1.AddNewPart<SharedStringTablePart>("rId4");
            //GenerateSharedStringTablePart1Content(sharedStringTablePart1,analytics,vendorName,startDate,endDate);

            //SetPackageProperties(document);
        }

        #region GenerateExtendedFilePropertiesPart1Content
        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Excel";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector(){ BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Worksheets";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "1";

            variant2.Append(vTInt321);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector(){ BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)1U };
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Vendor Analytics";

            vTVector2.Append(vTLPSTR2);

            titlesOfParts1.Append(vTVector2);
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "12.0000";

            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }
        #endregion

        #region GenerateWorkbookPart1Content
        // Generates content of workbookPart1.
        private void GenerateWorkbookPart1Content(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook();
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            FileVersion fileVersion1 = new FileVersion(){ ApplicationName = "xl", LastEdited = "4", LowestEdited = "4", BuildVersion = "4507" };
            WorkbookProperties workbookProperties1 = new WorkbookProperties(){ DefaultThemeVersion = (UInt32Value)124226U };

            BookViews bookViews1 = new BookViews();
            WorkbookView workbookView1 = new WorkbookView(){ XWindow = 210, YWindow = 495, WindowWidth = (UInt32Value)18855U, WindowHeight = (UInt32Value)11445U };

            bookViews1.Append(workbookView1);

            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet(){ Name = "Vendor Analytics", SheetId = (UInt32Value)1U, Id = "rId1" };

            sheets1.Append(sheet1);
            CalculationProperties calculationProperties1 = new CalculationProperties(){ CalculationId = (UInt32Value)125725U };
            FileRecoveryProperties fileRecoveryProperties1 = new FileRecoveryProperties(){ RepairLoad = true };

            workbook1.Append(fileVersion1);
            workbook1.Append(workbookProperties1);
            workbook1.Append(bookViews1);
            workbook1.Append(sheets1);
            workbook1.Append(calculationProperties1);
            workbook1.Append(fileRecoveryProperties1);

            workbookPart1.Workbook = workbook1;
        }
        #endregion

        #region GenerateWorkbookStylesPart1Content
        // Generates content of workbookStylesPart1.
        private void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet();

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)3U };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            FontName fontName1 = new FontName() { Val = "Arial" };

            font1.Append(fontSize1);
            font1.Append(fontName1);

            Font font2 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            FontName fontName2 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };

            font2.Append(bold1);
            font2.Append(fontSize2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering1);

            Font font3 = new Font();
            Bold bold2 = new Bold();
            FontSize fontSize3 = new FontSize() { Val = 14D };
            FontName fontName3 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };

            font3.Append(bold2);
            font3.Append(fontSize3);
            font3.Append(fontName3);
            font3.Append(fontFamilyNumbering2);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);

            Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);

            Borders borders1 = new Borders() { Count = (UInt32Value)1U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            borders1.Append(border1);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)3U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true };

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium9", DefaultPivotStyle = "PivotStyleLight16" };

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);

            workbookStylesPart1.Stylesheet = stylesheet1;
        }
        #endregion

        #region GenerateThemePart1Content
        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme(){ Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme(){ Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor(){ Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor(){ Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex(){ Val = "1F497D" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex(){ Val = "EEECE1" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex(){ Val = "4F81BD" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex(){ Val = "C0504D" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex(){ Val = "9BBB59" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex(){ Val = "8064A2" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex(){ Val = "4BACC6" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex(){ Val = "F79646" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex(){ Val = "0000FF" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex(){ Val = "800080" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme(){ Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont(){ Typeface = "Cambria" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont(){ Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont(){ Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont(){ Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont(){ Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont(){ Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont(){ Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont(){ Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont(){ Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont(){ Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont(){ Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont(){ Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont(){ Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont(){ Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont(){ Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont(){ Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont(){ Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont(){ Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont(){ Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont(){ Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont(){ Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont(){ Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont(){ Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont(){ Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont(){ Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont(){ Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont(){ Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont(){ Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont(){ Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont(){ Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont(){ Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont(){ Script = "Uigh", Typeface = "Microsoft Uighur" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont(){ Typeface = "Calibri" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont(){ Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont(){ Typeface = "" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont(){ Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont(){ Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont(){ Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont(){ Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont(){ Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont(){ Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont(){ Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont(){ Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont(){ Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont(){ Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont(){ Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont(){ Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont(){ Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont(){ Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont(){ Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont(){ Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont(){ Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont(){ Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont(){ Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont(){ Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont(){ Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont(){ Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont(){ Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont(){ Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont(){ Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont(){ Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont(){ Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont(){ Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont(){ Script = "Uigh", Typeface = "Microsoft Uighur" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont30);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme(){ Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint1 = new A.Tint(){ Val = 50000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation(){ Val = 300000 };

            schemeColor2.Append(tint1);
            schemeColor2.Append(saturationModulation1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop(){ Position = 35000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint2 = new A.Tint(){ Val = 37000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation(){ Val = 300000 };

            schemeColor3.Append(tint2);
            schemeColor3.Append(saturationModulation2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint3 = new A.Tint(){ Val = 15000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation(){ Val = 350000 };

            schemeColor4.Append(tint3);
            schemeColor4.Append(saturationModulation3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill(){ Angle = 16200000, Scaled = true };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Shade shade1 = new A.Shade(){ Val = 51000 };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation(){ Val = 130000 };

            schemeColor5.Append(shade1);
            schemeColor5.Append(saturationModulation4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop(){ Position = 80000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Shade shade2 = new A.Shade(){ Val = 93000 };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation(){ Val = 130000 };

            schemeColor6.Append(shade2);
            schemeColor6.Append(saturationModulation5);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Shade shade3 = new A.Shade(){ Val = 94000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation(){ Val = 135000 };

            schemeColor7.Append(shade3);
            schemeColor7.Append(saturationModulation6);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill(){ Angle = 16200000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline(){ Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Shade shade4 = new A.Shade(){ Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation(){ Val = 105000 };

            schemeColor8.Append(shade4);
            schemeColor8.Append(saturationModulation7);

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);

            A.Outline outline2 = new A.Outline(){ Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);

            A.Outline outline3 = new A.Outline(){ Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();

            A.EffectList effectList1 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow(){ BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex(){ Val = "000000" };
            A.Alpha alpha1 = new A.Alpha(){ Val = 38000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList1.Append(outerShadow1);

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();

            A.EffectList effectList2 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow(){ BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex(){ Val = "000000" };
            A.Alpha alpha2 = new A.Alpha(){ Val = 35000 };

            rgbColorModelHex12.Append(alpha2);

            outerShadow2.Append(rgbColorModelHex12);

            effectList2.Append(outerShadow2);

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow3 = new A.OuterShadow(){ BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex(){ Val = "000000" };
            A.Alpha alpha3 = new A.Alpha(){ Val = 35000 };

            rgbColorModelHex13.Append(alpha3);

            outerShadow3.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow3);

            A.Scene3DType scene3DType1 = new A.Scene3DType();

            A.Camera camera1 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.Rotation rotation1 = new A.Rotation(){ Latitude = 0, Longitude = 0, Revolution = 0 };

            camera1.Append(rotation1);

            A.LightRig lightRig1 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };
            A.Rotation rotation2 = new A.Rotation(){ Latitude = 0, Longitude = 0, Revolution = 1200000 };

            lightRig1.Append(rotation2);

            scene3DType1.Append(camera1);
            scene3DType1.Append(lightRig1);

            A.Shape3DType shape3DType1 = new A.Shape3DType();
            A.BevelTop bevelTop1 = new A.BevelTop(){ Width = 63500L, Height = 25400L };

            shape3DType1.Append(bevelTop1);

            effectStyle3.Append(effectList3);
            effectStyle3.Append(scene3DType1);
            effectStyle3.Append(shape3DType1);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.GradientFill gradientFill3 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor12 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint4 = new A.Tint(){ Val = 40000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation(){ Val = 350000 };

            schemeColor12.Append(tint4);
            schemeColor12.Append(saturationModulation8);

            gradientStop7.Append(schemeColor12);

            A.GradientStop gradientStop8 = new A.GradientStop(){ Position = 40000 };

            A.SchemeColor schemeColor13 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint(){ Val = 45000 };
            A.Shade shade5 = new A.Shade(){ Val = 99000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation(){ Val = 350000 };

            schemeColor13.Append(tint5);
            schemeColor13.Append(shade5);
            schemeColor13.Append(saturationModulation9);

            gradientStop8.Append(schemeColor13);

            A.GradientStop gradientStop9 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Shade shade6 = new A.Shade(){ Val = 20000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation(){ Val = 255000 };

            schemeColor14.Append(shade6);
            schemeColor14.Append(saturationModulation10);

            gradientStop9.Append(schemeColor14);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);

            A.PathGradientFill pathGradientFill1 = new A.PathGradientFill(){ Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle1 = new A.FillToRectangle(){ Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill1.Append(fillToRectangle1);

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(pathGradientFill1);

            A.GradientFill gradientFill4 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop10 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor15 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint(){ Val = 80000 };
            A.SaturationModulation saturationModulation11 = new A.SaturationModulation(){ Val = 300000 };

            schemeColor15.Append(tint6);
            schemeColor15.Append(saturationModulation11);

            gradientStop10.Append(schemeColor15);

            A.GradientStop gradientStop11 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor16 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Shade shade7 = new A.Shade(){ Val = 30000 };
            A.SaturationModulation saturationModulation12 = new A.SaturationModulation(){ Val = 200000 };

            schemeColor16.Append(shade7);
            schemeColor16.Append(saturationModulation12);

            gradientStop11.Append(schemeColor16);

            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);

            A.PathGradientFill pathGradientFill2 = new A.PathGradientFill(){ Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle2 = new A.FillToRectangle(){ Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };

            pathGradientFill2.Append(fillToRectangle2);

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(pathGradientFill2);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(gradientFill3);
            backgroundFillStyleList1.Append(gradientFill4);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            themePart1.Theme = theme1;
        }
        #endregion

        #region GenerateWorksheetPart1Content - MODIFY COMMENTED
        // Generates content of worksheetPart1.
        private void GenerateWorksheetPart1Content(WorksheetPart worksheetPart1, List<SiteAnalyticsVendorSummary> analytics, string vendorName, DateTime startDate, DateTime endDate)
        {
            Worksheet worksheet1 = new Worksheet();
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1:L13" };

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection1 = new Selection() { SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1:L1" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 14.25D };

            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 30.625D, CustomWidth = true };
            columns1.Append(column1);

            SheetData sheetData1 = new SheetData();


            Cell c;
            CellValue cv;
            Row r;

            UInt32Value rowIndex;
            string cellReference;
            int startRange;
            int endRange;
            rowIndex = 1U;
            #region TITLE
            cellReference = "A" + rowIndex.ToString();
            r = new Row() { RowIndex = rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:12" } };
            c = new Cell() { CellReference = cellReference, StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
            cv = new CellValue();
            //cellValue1.Text = "0";
            cv.Text = vendorName + " analytics summary";
            c.Append(cv);
            r.Append(c);
            sheetData1.Append(r);
            #endregion

            rowIndex +=2;
            #region IMPRESSIONS TITLE
            cellReference = "A" + rowIndex.ToString();
            r = new Row() { RowIndex = rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };
            c = new Cell() { CellReference = cellReference, StyleIndex = (UInt32Value)2U, DataType = CellValues.String };
            cv = new CellValue();
            //cellValue1.Text = "0";
            cv.Text = "Impressions";
            c.Append(cv);
            r.Append(c);
            sheetData1.Append(r);


            //Row row2 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };
            //Cell cell2 = new Cell() { CellReference = "A3", DataType = CellValues.SharedString };
            //CellValue cellValue2 = new CellValue();
            //cellValue2.Text = "1";
            //cell2.Append(cellValue2);
            //row2.Append(cell2);
            #endregion

            rowIndex++;
            #region IMPRESSIONS TOTAL PORTFOLIO
            cellReference = "A" + rowIndex.ToString();
            r = new Row() { RowIndex = rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };
            c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
            cv = new CellValue();
            cv.Text = "Total portfolio";
            c.Append(cv);
            r.Append(c);

            int totalPortfolioImpressions = analytics.Sum(x => x.Impressions);
            cellReference = "B" + rowIndex.ToString();
            c = new Cell() { CellReference = cellReference, DataType = CellValues.Number };
            cv = new CellValue();
            cv.Text = totalPortfolioImpressions.ToString();
            c.Append(cv);
            r.Append(c);

            cellReference = "C" + rowIndex.ToString();
            c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
            cv = new CellValue();
            cv.Text = "between " + startDate.ToShortDateString() + " and " + endDate.ToShortDateString();
            c.Append(cv);
            r.Append(c);

            //cellReference = "D" + rowIndex.ToString();
            //c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
            //cv = new CellValue();
            //cv.Text = startDate.ToShortDateString();
            //c.Append(cv);
            //r.Append(c);

            //cellReference = "E" + rowIndex.ToString();
            //c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
            //cv = new CellValue();
            //cv.Text = "and";
            //c.Append(cv);
            //r.Append(c);

            //cellReference = "F" + rowIndex.ToString();
            //c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
            //cv = new CellValue();
            //cv.Text = endDate.ToShortDateString();
            //c.Append(cv);
            //r.Append(c);

            sheetData1.Append(r);



            //Row row3 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };

            //Cell cell3 = new Cell() { CellReference = "A4", DataType = CellValues.SharedString };
            //CellValue cellValue3 = new CellValue();
            //cellValue3.Text = "2";

            //cell3.Append(cellValue3);

            //Cell cell4 = new Cell() { CellReference = "B4", DataType = CellValues.Number };
            //CellValue cellValue4 = new CellValue();
            //cellValue4.Text = "3";

            //cell4.Append(cellValue4);

            //Cell cell5 = new Cell() { CellReference = "C4", DataType = CellValues.SharedString };
            //CellValue cellValue5 = new CellValue();
            //cellValue5.Text = "4";

            //cell5.Append(cellValue5);

            //Cell cell6 = new Cell() { CellReference = "D4", DataType = CellValues.SharedString };
            //CellValue cellValue6 = new CellValue();
            //cellValue6.Text = "5";

            //cell6.Append(cellValue6);

            //Cell cell7 = new Cell() { CellReference = "E4", DataType = CellValues.SharedString };
            //CellValue cellValue7 = new CellValue();
            //cellValue7.Text = "6";

            //cell7.Append(cellValue7);

            //Cell cell8 = new Cell() { CellReference = "F4", DataType = CellValues.SharedString };
            //CellValue cellValue8 = new CellValue();
            //cellValue8.Text = "7";

            //cell8.Append(cellValue8);

            //row3.Append(cell3);
            //row3.Append(cell4);
            //row3.Append(cell5);
            //row3.Append(cell6);
            //row3.Append(cell7);
            //row3.Append(cell8);
            #endregion

            rowIndex +=2;
            #region IMPRESSIONS SERVICES
            startRange = (int)rowIndex.Value;
            foreach (SiteAnalyticsVendorSummary savs in analytics)
            {
                cellReference = "A" + rowIndex.ToString();
                r = new Row() { RowIndex = rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };
                c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
                cv = new CellValue();
                cv.Text = savs.ServiceName;
                c.Append(cv);
                r.Append(c);

                cellReference = "B" + rowIndex.ToString();
                c = new Cell() { CellReference = cellReference, DataType = CellValues.Number };
                cv = new CellValue();
                cv.Text = savs.Impressions.ToString();
                c.Append(cv);
                r.Append(c);

                cellReference = "C" + rowIndex.ToString();
                c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
                cv = new CellValue();
                cv.Text = "between " + startDate.ToShortDateString() + " and " + endDate.ToShortDateString();
                c.Append(cv);
                r.Append(c);

                //cellReference = "D" + rowIndex.ToString();
                //c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
                //cv = new CellValue();
                //cv.Text = startDate.ToShortDateString();
                //c.Append(cv);
                //r.Append(c);

                //cellReference = "E" + rowIndex.ToString();
                //c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
                //cv = new CellValue();
                //cv.Text = "and";
                //c.Append(cv);
                //r.Append(c);

                //cellReference = "F" + rowIndex.ToString();
                //c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
                //cv = new CellValue();
                //cv.Text = endDate.ToShortDateString();
                //c.Append(cv);
                //r.Append(c);

                rowIndex++;

                sheetData1.Append(r);
            }
            endRange = (int)rowIndex.Value-1;

            impressionsLabelRange = "\'Vendor Analytics\'!$A$" + startRange.ToString() + ":$A$" + endRange.ToString();
            impressionsDataRange = "\'Vendor Analytics\'!$B$" + startRange.ToString() + ":$B$" + endRange.ToString();
            //formula3.Text = "\'Vendor Analytics\'!$A$6:$A$13";

                //Row row4 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };

            //Cell cell9 = new Cell() { CellReference = "A6", DataType = CellValues.SharedString };
            //CellValue cellValue9 = new CellValue();
            //cellValue9.Text = "8";

            //cell9.Append(cellValue9);

            //Cell cell10 = new Cell() { CellReference = "B6" };
            //CellValue cellValue10 = new CellValue();
            //cellValue10.Text = "851";

            //cell10.Append(cellValue10);

            //Cell cell11 = new Cell() { CellReference = "C6", DataType = CellValues.SharedString };
            //CellValue cellValue11 = new CellValue();
            //cellValue11.Text = "4";

            //cell11.Append(cellValue11);

            //Cell cell12 = new Cell() { CellReference = "D6", DataType = CellValues.SharedString };
            //CellValue cellValue12 = new CellValue();
            //cellValue12.Text = "5";

            //cell12.Append(cellValue12);

            //Cell cell13 = new Cell() { CellReference = "E6", DataType = CellValues.SharedString };
            //CellValue cellValue13 = new CellValue();
            //cellValue13.Text = "6";

            //cell13.Append(cellValue13);

            //Cell cell14 = new Cell() { CellReference = "F6", DataType = CellValues.SharedString };
            //CellValue cellValue14 = new CellValue();
            //cellValue14.Text = "7";

            //cell14.Append(cellValue14);

            //row4.Append(cell9);
            //row4.Append(cell10);
            //row4.Append(cell11);
            //row4.Append(cell12);
            //row4.Append(cell13);
            //row4.Append(cell14);


            #endregion

            //rowIndex += 2;
            rowIndex = 16;
            #region COMPARISON RESULTS IMPRESSIONS TITLE
            cellReference = "A" + rowIndex.ToString();
            r = new Row() { RowIndex = rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };
            c = new Cell() { CellReference = cellReference, StyleIndex = (UInt32Value)2U, DataType = CellValues.String };
            cv = new CellValue();
            //cellValue1.Text = "0";
            cv.Text = "Comparison results impressions";
            c.Append(cv);
            r.Append(c);
            sheetData1.Append(r);
            #endregion

            rowIndex++;
            #region COMPARISON RESULTS IMPRESSIONS TOTAL PORTFOLIO
            cellReference = "A" + rowIndex.ToString();
            r = new Row() { RowIndex = rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };
            c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
            cv = new CellValue();
            cv.Text = "Total portfolio";
            c.Append(cv);
            r.Append(c);

            int totalPortfolioComparisonResultsImpressions = analytics.Sum(x => x.ComparisonResultImpressions);
            cellReference = "B" + rowIndex.ToString();
            c = new Cell() { CellReference = cellReference, DataType = CellValues.Number };
            cv = new CellValue();
            cv.Text = totalPortfolioComparisonResultsImpressions.ToString();
            c.Append(cv);
            r.Append(c);

            cellReference = "C" + rowIndex.ToString();
            c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
            cv = new CellValue();
            cv.Text = "between " + startDate.ToShortDateString() + " and " + endDate.ToShortDateString();
            c.Append(cv);
            r.Append(c);

            sheetData1.Append(r);
            #endregion

            rowIndex += 2;
            #region COMPARISON RESULTS IMPRESSIONS SERVICES
            startRange = (int)rowIndex.Value;
            foreach (SiteAnalyticsVendorSummary savs in analytics)
            {
                cellReference = "A" + rowIndex.ToString();
                r = new Row() { RowIndex = rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };
                c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
                cv = new CellValue();
                cv.Text = savs.ServiceName;
                c.Append(cv);
                r.Append(c);

                cellReference = "B" + rowIndex.ToString();
                c = new Cell() { CellReference = cellReference, DataType = CellValues.Number };
                cv = new CellValue();
                cv.Text = savs.ComparisonResultImpressions.ToString();
                c.Append(cv);
                r.Append(c);

                cellReference = "C" + rowIndex.ToString();
                c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
                cv = new CellValue();
                cv.Text = "between " + startDate.ToShortDateString() + " and " + endDate.ToShortDateString();
                c.Append(cv);
                r.Append(c);

                rowIndex++;

                sheetData1.Append(r);
            }
            endRange = (int)rowIndex.Value - 1;

            comparisonResultsImpressionsLabelRange = "\'Vendor Analytics\'!$A$" + startRange.ToString() + ":$A$" + endRange.ToString();
            comparisonResultsImpressionsDataRange = "\'Vendor Analytics\'!$B$" + startRange.ToString() + ":$B$" + endRange.ToString();
            //formula3.Text = "\'Vendor Analytics\'!$A$6:$A$13";
            #endregion

            //rowIndex += 2;
            rowIndex = 29;
            #region SHOP VISITS TITLE
            cellReference = "A" + rowIndex.ToString();
            r = new Row() { RowIndex = rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };
            c = new Cell() { CellReference = cellReference, StyleIndex = (UInt32Value)2U, DataType = CellValues.String };
            cv = new CellValue();
            //cellValue1.Text = "0";
            cv.Text = "Shop visits";
            c.Append(cv);
            r.Append(c);
            sheetData1.Append(r);
            #endregion

            rowIndex++;
            #region SHOP VISITS TOTAL PORTFOLIO
            cellReference = "A" + rowIndex.ToString();
            r = new Row() { RowIndex = rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };
            c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
            cv = new CellValue();
            cv.Text = "Total portfolio";
            c.Append(cv);
            r.Append(c);

            int totalPortfolioShopVisits = analytics.Sum(x => x.ShopVisits);
            cellReference = "B" + rowIndex.ToString();
            c = new Cell() { CellReference = cellReference, DataType = CellValues.Number };
            cv = new CellValue();
            cv.Text = totalPortfolioShopVisits.ToString();
            c.Append(cv);
            r.Append(c);

            cellReference = "C" + rowIndex.ToString();
            c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
            cv = new CellValue();
            cv.Text = "between " + startDate.ToShortDateString() + " and " + endDate.ToShortDateString();
            c.Append(cv);
            r.Append(c);

            sheetData1.Append(r);
            #endregion

            rowIndex += 2;
            #region SHOP VISITS SERVICES
            startRange = (int)rowIndex.Value;
            foreach (SiteAnalyticsVendorSummary savs in analytics)
            {
                cellReference = "A" + rowIndex.ToString();
                r = new Row() { RowIndex = rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };
                c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
                cv = new CellValue();
                cv.Text = savs.ServiceName;
                c.Append(cv);
                r.Append(c);

                cellReference = "B" + rowIndex.ToString();
                c = new Cell() { CellReference = cellReference, DataType = CellValues.Number };
                cv = new CellValue();
                cv.Text = savs.ShopVisits.ToString();
                c.Append(cv);
                r.Append(c);

                cellReference = "C" + rowIndex.ToString();
                c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
                cv = new CellValue();
                cv.Text = "between " + startDate.ToShortDateString() + " and " + endDate.ToShortDateString();
                c.Append(cv);
                r.Append(c);

                rowIndex++;

                sheetData1.Append(r);
            }
            endRange = (int)rowIndex.Value - 1;

            shopVisitsLabelRange = "\'Vendor Analytics\'!$A$" + startRange.ToString() + ":$A$" + endRange.ToString();
            shopVisitsDataRange = "\'Vendor Analytics\'!$B$" + startRange.ToString() + ":$B$" + endRange.ToString();
            //formula3.Text = "\'Vendor Analytics\'!$A$6:$A$13";
            #endregion

            //rowIndex += 2;
            rowIndex = 42;
            #region SHOP CONTENT CONSUMPTION TITLE
            cellReference = "A" + rowIndex.ToString();
            r = new Row() { RowIndex = rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };
            c = new Cell() { CellReference = cellReference, StyleIndex = (UInt32Value)2U, DataType = CellValues.String };
            cv = new CellValue();
            //cellValue1.Text = "0";
            cv.Text = "Shop content consumption";
            c.Append(cv);
            r.Append(c);
            sheetData1.Append(r);
            #endregion

            rowIndex++;
            #region SHOP CONTENT CONSUMPTION TOTAL PORTFOLIO
            cellReference = "A" + rowIndex.ToString();
            r = new Row() { RowIndex = rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };
            c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
            cv = new CellValue();
            cv.Text = "Total portfolio";
            c.Append(cv);
            r.Append(c);

            int totalPortfolioShopContentConsumption = analytics.Sum(x => x.ShopContentConsumption);
            cellReference = "B" + rowIndex.ToString();
            c = new Cell() { CellReference = cellReference, DataType = CellValues.Number };
            cv = new CellValue();
            cv.Text = totalPortfolioShopContentConsumption.ToString();
            c.Append(cv);
            r.Append(c);

            cellReference = "C" + rowIndex.ToString();
            c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
            cv = new CellValue();
            cv.Text = "between " + startDate.ToShortDateString() + " and " + endDate.ToShortDateString();
            c.Append(cv);
            r.Append(c);

            sheetData1.Append(r);
            #endregion

            rowIndex += 2;
            #region SHOP CONTENT CONSUMPTION SERVICES
            startRange = (int)rowIndex.Value;
            foreach (SiteAnalyticsVendorSummary savs in analytics)
            {
                cellReference = "A" + rowIndex.ToString();
                r = new Row() { RowIndex = rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };
                c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
                cv = new CellValue();
                cv.Text = savs.ServiceName;
                c.Append(cv);
                r.Append(c);

                cellReference = "B" + rowIndex.ToString();
                c = new Cell() { CellReference = cellReference, DataType = CellValues.Number };
                cv = new CellValue();
                cv.Text = savs.ShopContentConsumption.ToString();
                c.Append(cv);
                r.Append(c);

                cellReference = "C" + rowIndex.ToString();
                c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
                cv = new CellValue();
                cv.Text = "between " + startDate.ToShortDateString() + " and " + endDate.ToShortDateString();
                c.Append(cv);
                r.Append(c);

                rowIndex++;

                sheetData1.Append(r);
            }
            endRange = (int)rowIndex.Value - 1;

            shopContentConsumptionLabelRange = "\'Vendor Analytics\'!$A$" + startRange.ToString() + ":$A$" + endRange.ToString();
            shopContentConsumptionDataRange = "\'Vendor Analytics\'!$B$" + startRange.ToString() + ":$B$" + endRange.ToString();
            //formula3.Text = "\'Vendor Analytics\'!$A$6:$A$13";
            #endregion

            //rowIndex += 2;
            rowIndex = 55;
            #region SHOP LEADS TITLE
            cellReference = "A" + rowIndex.ToString();
            r = new Row() { RowIndex = rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };
            c = new Cell() { CellReference = cellReference, StyleIndex = (UInt32Value)2U, DataType = CellValues.String };
            cv = new CellValue();
            //cellValue1.Text = "0";
            cv.Text = "Shop leads (Try/Buy requests - not emails)";
            c.Append(cv);
            r.Append(c);
            sheetData1.Append(r);
            #endregion

            rowIndex++;
            #region SHOP LEADS TOTAL PORTFOLIO
            cellReference = "A" + rowIndex.ToString();
            r = new Row() { RowIndex = rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };
            c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
            cv = new CellValue();
            cv.Text = "Total portfolio";
            c.Append(cv);
            r.Append(c);

            int totalPortfolioShopLeads = analytics.Sum(x => x.ShopLeads);
            cellReference = "B" + rowIndex.ToString();
            c = new Cell() { CellReference = cellReference, DataType = CellValues.Number };
            cv = new CellValue();
            cv.Text = totalPortfolioShopLeads.ToString();
            c.Append(cv);
            r.Append(c);

            cellReference = "C" + rowIndex.ToString();
            c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
            cv = new CellValue();
            cv.Text = "between " + startDate.ToShortDateString() + " and " + endDate.ToShortDateString();
            c.Append(cv);
            r.Append(c);

            sheetData1.Append(r);
            #endregion

            rowIndex += 2;
            #region SHOP LEADS SERVICES
            startRange = (int)rowIndex.Value;
            foreach (SiteAnalyticsVendorSummary savs in analytics)
            {
                cellReference = "A" + rowIndex.ToString();
                r = new Row() { RowIndex = rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };
                c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
                cv = new CellValue();
                cv.Text = savs.ServiceName;
                c.Append(cv);
                r.Append(c);

                cellReference = "B" + rowIndex.ToString();
                c = new Cell() { CellReference = cellReference, DataType = CellValues.Number };
                cv = new CellValue();
                cv.Text = savs.ShopLeads.ToString();
                c.Append(cv);
                r.Append(c);

                cellReference = "C" + rowIndex.ToString();
                c = new Cell() { CellReference = cellReference, DataType = CellValues.String };
                cv = new CellValue();
                cv.Text = "between " + startDate.ToShortDateString() + " and " + endDate.ToShortDateString();
                c.Append(cv);
                r.Append(c);

                rowIndex++;

                sheetData1.Append(r);
            }
            endRange = (int)rowIndex.Value - 1;

            shopLeadsLabelRange = "\'Vendor Analytics\'!$A$" + startRange.ToString() + ":$A$" + endRange.ToString();
            shopLeadsDataRange = "\'Vendor Analytics\'!$B$" + startRange.ToString() + ":$B$" + endRange.ToString();
            //formula3.Text = "\'Vendor Analytics\'!$A$6:$A$13";
            #endregion

            #region CRAP
            //Row row5 = new Row() { RowIndex = (UInt32Value)7U, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };

            //Cell cell15 = new Cell() { CellReference = "A7", DataType = CellValues.SharedString };
            //CellValue cellValue15 = new CellValue();
            //cellValue15.Text = "9";

            //cell15.Append(cellValue15);

            //Cell cell16 = new Cell() { CellReference = "B7" };
            //CellValue cellValue16 = new CellValue();
            //cellValue16.Text = "851";

            //cell16.Append(cellValue16);

            //Cell cell17 = new Cell() { CellReference = "C7", DataType = CellValues.SharedString };
            //CellValue cellValue17 = new CellValue();
            //cellValue17.Text = "4";

            //cell17.Append(cellValue17);

            //Cell cell18 = new Cell() { CellReference = "D7", DataType = CellValues.SharedString };
            //CellValue cellValue18 = new CellValue();
            //cellValue18.Text = "5";

            //cell18.Append(cellValue18);

            //Cell cell19 = new Cell() { CellReference = "E7", DataType = CellValues.SharedString };
            //CellValue cellValue19 = new CellValue();
            //cellValue19.Text = "6";

            //cell19.Append(cellValue19);

            //Cell cell20 = new Cell() { CellReference = "F7", DataType = CellValues.SharedString };
            //CellValue cellValue20 = new CellValue();
            //cellValue20.Text = "7";

            //cell20.Append(cellValue20);

            //row5.Append(cell15);
            //row5.Append(cell16);
            //row5.Append(cell17);
            //row5.Append(cell18);
            //row5.Append(cell19);
            //row5.Append(cell20);

            //Row row6 = new Row() { RowIndex = (UInt32Value)8U, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };

            //Cell cell21 = new Cell() { CellReference = "A8", DataType = CellValues.SharedString };
            //CellValue cellValue21 = new CellValue();
            //cellValue21.Text = "10";

            //cell21.Append(cellValue21);

            //Cell cell22 = new Cell() { CellReference = "B8" };
            //CellValue cellValue22 = new CellValue();
            //cellValue22.Text = "827";

            //cell22.Append(cellValue22);

            //Cell cell23 = new Cell() { CellReference = "C8", DataType = CellValues.SharedString };
            //CellValue cellValue23 = new CellValue();
            //cellValue23.Text = "4";

            //cell23.Append(cellValue23);

            //Cell cell24 = new Cell() { CellReference = "D8", DataType = CellValues.SharedString };
            //CellValue cellValue24 = new CellValue();
            //cellValue24.Text = "5";

            //cell24.Append(cellValue24);

            //Cell cell25 = new Cell() { CellReference = "E8", DataType = CellValues.SharedString };
            //CellValue cellValue25 = new CellValue();
            //cellValue25.Text = "6";

            //cell25.Append(cellValue25);

            //Cell cell26 = new Cell() { CellReference = "F8", DataType = CellValues.SharedString };
            //CellValue cellValue26 = new CellValue();
            //cellValue26.Text = "7";

            //cell26.Append(cellValue26);

            //row6.Append(cell21);
            //row6.Append(cell22);
            //row6.Append(cell23);
            //row6.Append(cell24);
            //row6.Append(cell25);
            //row6.Append(cell26);

            //Row row7 = new Row() { RowIndex = (UInt32Value)9U, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };

            //Cell cell27 = new Cell() { CellReference = "A9", DataType = CellValues.SharedString };
            //CellValue cellValue27 = new CellValue();
            //cellValue27.Text = "11";

            //cell27.Append(cellValue27);

            //Cell cell28 = new Cell() { CellReference = "B9" };
            //CellValue cellValue28 = new CellValue();
            //cellValue28.Text = "842";

            //cell28.Append(cellValue28);

            //Cell cell29 = new Cell() { CellReference = "C9", DataType = CellValues.SharedString };
            //CellValue cellValue29 = new CellValue();
            //cellValue29.Text = "4";

            //cell29.Append(cellValue29);

            //Cell cell30 = new Cell() { CellReference = "D9", DataType = CellValues.SharedString };
            //CellValue cellValue30 = new CellValue();
            //cellValue30.Text = "5";

            //cell30.Append(cellValue30);

            //Cell cell31 = new Cell() { CellReference = "E9", DataType = CellValues.SharedString };
            //CellValue cellValue31 = new CellValue();
            //cellValue31.Text = "6";

            //cell31.Append(cellValue31);

            //Cell cell32 = new Cell() { CellReference = "F9", DataType = CellValues.SharedString };
            //CellValue cellValue32 = new CellValue();
            //cellValue32.Text = "7";

            //cell32.Append(cellValue32);

            //row7.Append(cell27);
            //row7.Append(cell28);
            //row7.Append(cell29);
            //row7.Append(cell30);
            //row7.Append(cell31);
            //row7.Append(cell32);

            //Row row8 = new Row() { RowIndex = (UInt32Value)10U, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };

            //Cell cell33 = new Cell() { CellReference = "A10", DataType = CellValues.SharedString };
            //CellValue cellValue33 = new CellValue();
            //cellValue33.Text = "12";

            //cell33.Append(cellValue33);

            //Cell cell34 = new Cell() { CellReference = "B10" };
            //CellValue cellValue34 = new CellValue();
            //cellValue34.Text = "827";

            //cell34.Append(cellValue34);

            //Cell cell35 = new Cell() { CellReference = "C10", DataType = CellValues.SharedString };
            //CellValue cellValue35 = new CellValue();
            //cellValue35.Text = "4";

            //cell35.Append(cellValue35);

            //Cell cell36 = new Cell() { CellReference = "D10", DataType = CellValues.SharedString };
            //CellValue cellValue36 = new CellValue();
            //cellValue36.Text = "5";

            //cell36.Append(cellValue36);

            //Cell cell37 = new Cell() { CellReference = "E10", DataType = CellValues.SharedString };
            //CellValue cellValue37 = new CellValue();
            //cellValue37.Text = "6";

            //cell37.Append(cellValue37);

            //Cell cell38 = new Cell() { CellReference = "F10", DataType = CellValues.SharedString };
            //CellValue cellValue38 = new CellValue();
            //cellValue38.Text = "7";

            //cell38.Append(cellValue38);

            //row8.Append(cell33);
            //row8.Append(cell34);
            //row8.Append(cell35);
            //row8.Append(cell36);
            //row8.Append(cell37);
            //row8.Append(cell38);

            //Row row9 = new Row() { RowIndex = (UInt32Value)11U, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };

            //Cell cell39 = new Cell() { CellReference = "A11", DataType = CellValues.SharedString };
            //CellValue cellValue39 = new CellValue();
            //cellValue39.Text = "13";

            //cell39.Append(cellValue39);

            //Cell cell40 = new Cell() { CellReference = "B11" };
            //CellValue cellValue40 = new CellValue();
            //cellValue40.Text = "856";

            //cell40.Append(cellValue40);

            //Cell cell41 = new Cell() { CellReference = "C11", DataType = CellValues.SharedString };
            //CellValue cellValue41 = new CellValue();
            //cellValue41.Text = "4";

            //cell41.Append(cellValue41);

            //Cell cell42 = new Cell() { CellReference = "D11", DataType = CellValues.SharedString };
            //CellValue cellValue42 = new CellValue();
            //cellValue42.Text = "5";

            //cell42.Append(cellValue42);

            //Cell cell43 = new Cell() { CellReference = "E11", DataType = CellValues.SharedString };
            //CellValue cellValue43 = new CellValue();
            //cellValue43.Text = "6";

            //cell43.Append(cellValue43);

            //Cell cell44 = new Cell() { CellReference = "F11", DataType = CellValues.SharedString };
            //CellValue cellValue44 = new CellValue();
            //cellValue44.Text = "7";

            //cell44.Append(cellValue44);

            //row9.Append(cell39);
            //row9.Append(cell40);
            //row9.Append(cell41);
            //row9.Append(cell42);
            //row9.Append(cell43);
            //row9.Append(cell44);

            //Row row10 = new Row() { RowIndex = (UInt32Value)12U, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };

            //Cell cell45 = new Cell() { CellReference = "A12", DataType = CellValues.SharedString };
            //CellValue cellValue45 = new CellValue();
            //cellValue45.Text = "14";

            //cell45.Append(cellValue45);

            //Cell cell46 = new Cell() { CellReference = "B12" };
            //CellValue cellValue46 = new CellValue();
            //cellValue46.Text = "843";

            //cell46.Append(cellValue46);

            //Cell cell47 = new Cell() { CellReference = "C12", DataType = CellValues.SharedString };
            //CellValue cellValue47 = new CellValue();
            //cellValue47.Text = "4";

            //cell47.Append(cellValue47);

            //Cell cell48 = new Cell() { CellReference = "D12", DataType = CellValues.SharedString };
            //CellValue cellValue48 = new CellValue();
            //cellValue48.Text = "5";

            //cell48.Append(cellValue48);

            //Cell cell49 = new Cell() { CellReference = "E12", DataType = CellValues.SharedString };
            //CellValue cellValue49 = new CellValue();
            //cellValue49.Text = "6";

            //cell49.Append(cellValue49);

            //Cell cell50 = new Cell() { CellReference = "F12", DataType = CellValues.SharedString };
            //CellValue cellValue50 = new CellValue();
            //cellValue50.Text = "7";

            //cell50.Append(cellValue50);

            //row10.Append(cell45);
            //row10.Append(cell46);
            //row10.Append(cell47);
            //row10.Append(cell48);
            //row10.Append(cell49);
            //row10.Append(cell50);

            //Row row11 = new Row() { RowIndex = (UInt32Value)13U, Spans = new ListValue<StringValue>() { InnerText = "1:6" } };

            //Cell cell51 = new Cell() { CellReference = "A13", DataType = CellValues.SharedString };
            //CellValue cellValue51 = new CellValue();
            //cellValue51.Text = "15";

            //cell51.Append(cellValue51);

            //Cell cell52 = new Cell() { CellReference = "B13" };
            //CellValue cellValue52 = new CellValue();
            //cellValue52.Text = "856";

            //cell52.Append(cellValue52);

            //Cell cell53 = new Cell() { CellReference = "C13", DataType = CellValues.SharedString };
            //CellValue cellValue53 = new CellValue();
            //cellValue53.Text = "4";

            //cell53.Append(cellValue53);

            //Cell cell54 = new Cell() { CellReference = "D13", DataType = CellValues.SharedString };
            //CellValue cellValue54 = new CellValue();
            //cellValue54.Text = "5";

            //cell54.Append(cellValue54);

            //Cell cell55 = new Cell() { CellReference = "E13", DataType = CellValues.SharedString };
            //CellValue cellValue55 = new CellValue();
            //cellValue55.Text = "6";

            //cell55.Append(cellValue55);

            //Cell cell56 = new Cell() { CellReference = "F13", DataType = CellValues.SharedString };
            //CellValue cellValue56 = new CellValue();
            //cellValue56.Text = "7";

            //cell56.Append(cellValue56);

            //row11.Append(cell51);
            //row11.Append(cell52);
            //row11.Append(cell53);
            //row11.Append(cell54);
            //row11.Append(cell55);
            //row11.Append(cell56);
            #endregion

            //sheetData1.Append(row1);
            //sheetData1.Append(row2);
            //sheetData1.Append(row3);
            //sheetData1.Append(row4);
            //sheetData1.Append(row5);
            //sheetData1.Append(row6);
            //sheetData1.Append(row7);
            //sheetData1.Append(row8);
            //sheetData1.Append(row9);
            //sheetData1.Append(row10);
            //sheetData1.Append(row11);

            MergeCells mergeCells1 = new MergeCells() { Count = (UInt32Value)1U };
            MergeCell mergeCell1 = new MergeCell() { Reference = "A1:L1" };
            mergeCells1.Append(mergeCell1);

            PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup1 = new PageSetup() { PaperSize = (UInt32Value)9U, Orientation = OrientationValues.Portrait, HorizontalDpi = (UInt32Value)300U, VerticalDpi = (UInt32Value)300U, Id = "rId1" };
            Drawing drawing1 = new Drawing() { Id = "rId2" };

            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(columns1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(mergeCells1);
            //worksheet1.Append(pageMargins1);
            //worksheet1.Append(pageSetup1);
            worksheet1.Append(drawing1);

            worksheetPart1.Worksheet = worksheet1;
        }
        #endregion

        #region GenerateDrawingsPart1Content
        // Generates content of drawingsPart1.
        private void GenerateDrawingsPart1Content(DrawingsPart drawingsPart1)
        {
            Xdr.WorksheetDrawing worksheetDrawing1 = new Xdr.WorksheetDrawing();
            worksheetDrawing1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheetDrawing1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Xdr.TwoCellAnchor twoCellAnchor1 = new Xdr.TwoCellAnchor();

            Xdr.FromMarker fromMarker1 = new Xdr.FromMarker();
            Xdr.ColumnId columnId1 = new Xdr.ColumnId();
            //columnId1.Text = "7";
            columnId1.Text = "6";
            Xdr.ColumnOffset columnOffset1 = new Xdr.ColumnOffset();
            //columnOffset1.Text = "95250";
            columnOffset1.Text = "0";
            Xdr.RowId rowId1 = new Xdr.RowId();
            //rowId1.Text = "3";
            rowId1.Text = (3 + (sectionHeight * 0)).ToString();
            Xdr.RowOffset rowOffset1 = new Xdr.RowOffset();
            //rowOffset1.Text = "171450";
            rowOffset1.Text = "0";

            fromMarker1.Append(columnId1);
            fromMarker1.Append(columnOffset1);
            fromMarker1.Append(rowId1);
            fromMarker1.Append(rowOffset1);

            Xdr.ToMarker toMarker1 = new Xdr.ToMarker();
            Xdr.ColumnId columnId2 = new Xdr.ColumnId();
            //columnId2.Text = "13";
            columnId2.Text = "12";
            Xdr.ColumnOffset columnOffset2 = new Xdr.ColumnOffset();
            //columnOffset2.Text = "466725";
            columnOffset2.Text = "0";
            Xdr.RowId rowId2 = new Xdr.RowId();
            //rowId2.Text = "17";
            //rowId2.Text = "13";
            rowId2.Text = (3 + (sectionHeight * 0) + graphHeight).ToString();
            Xdr.RowOffset rowOffset2 = new Xdr.RowOffset();
            //rowOffset2.Text = "171450";
            rowOffset2.Text = "0";

            toMarker1.Append(columnId2);
            toMarker1.Append(columnOffset2);
            toMarker1.Append(rowId2);
            toMarker1.Append(rowOffset2);

            Xdr.GraphicFrame graphicFrame1 = new Xdr.GraphicFrame(){ Macro = "" };

            Xdr.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = new Xdr.NonVisualGraphicFrameProperties();
            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Chart 1" };
            Xdr.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Xdr.NonVisualGraphicFrameDrawingProperties();

            nonVisualGraphicFrameProperties1.Append(nonVisualDrawingProperties1);
            nonVisualGraphicFrameProperties1.Append(nonVisualGraphicFrameDrawingProperties1);

            Xdr.Transform transform1 = new Xdr.Transform();
            A.Offset offset1 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents(){ Cx = 0L, Cy = 0L };

            transform1.Append(offset1);
            transform1.Append(extents1);

            A.Graphic graphic1 = new A.Graphic();

            A.GraphicData graphicData1 = new A.GraphicData(){ Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

            C.ChartReference chartReference1 = new C.ChartReference(){ Id = "rId1" };
            chartReference1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartReference1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            graphicData1.Append(chartReference1);

            graphic1.Append(graphicData1);

            graphicFrame1.Append(nonVisualGraphicFrameProperties1);
            graphicFrame1.Append(transform1);
            graphicFrame1.Append(graphic1);
            Xdr.ClientData clientData1 = new Xdr.ClientData();

            twoCellAnchor1.Append(fromMarker1);
            twoCellAnchor1.Append(toMarker1);
            twoCellAnchor1.Append(graphicFrame1);
            twoCellAnchor1.Append(clientData1);


            //SECOND GRAPH
            Xdr.TwoCellAnchor twoCellAnchor2 = new Xdr.TwoCellAnchor();

            Xdr.FromMarker fromMarker2 = new Xdr.FromMarker();
            Xdr.ColumnId columnId3 = new Xdr.ColumnId();
            columnId3.Text = "6";
            Xdr.ColumnOffset columnOffset3 = new Xdr.ColumnOffset();
            columnOffset3.Text = "0";
            Xdr.RowId rowId3 = new Xdr.RowId();
            //rowId3.Text = "16";
            rowId3.Text = (3 + ((sectionHeight + sectionSpacing) * 1)).ToString();
            Xdr.RowOffset rowOffset3 = new Xdr.RowOffset();
            rowOffset3.Text = "0";

            fromMarker2.Append(columnId3);
            fromMarker2.Append(columnOffset3);
            fromMarker2.Append(rowId3);
            fromMarker2.Append(rowOffset3);

            Xdr.ToMarker toMarker2 = new Xdr.ToMarker();
            Xdr.ColumnId columnId4 = new Xdr.ColumnId();
            columnId4.Text = "12";
            Xdr.ColumnOffset columnOffset4 = new Xdr.ColumnOffset();
            columnOffset4.Text = "0";
            Xdr.RowId rowId4 = new Xdr.RowId();
            //rowId4.Text = "26";
            rowId4.Text = (3 + ((sectionHeight + sectionSpacing) * 1) + graphHeight).ToString();
            Xdr.RowOffset rowOffset4 = new Xdr.RowOffset();
            rowOffset4.Text = "0";

            toMarker2.Append(columnId4);
            toMarker2.Append(columnOffset4);
            toMarker2.Append(rowId4);
            toMarker2.Append(rowOffset4);

            Xdr.GraphicFrame graphicFrame2 = new Xdr.GraphicFrame() { Macro = "" };

            Xdr.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties2 = new Xdr.NonVisualGraphicFrameProperties();
            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties2 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Chart 2" };
            Xdr.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties2 = new Xdr.NonVisualGraphicFrameDrawingProperties();

            nonVisualGraphicFrameProperties2.Append(nonVisualDrawingProperties2);
            nonVisualGraphicFrameProperties2.Append(nonVisualGraphicFrameDrawingProperties2);

            Xdr.Transform transform2 = new Xdr.Transform();
            A.Offset offset2 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents2 = new A.Extents() { Cx = 0L, Cy = 0L };

            transform2.Append(offset2);
            transform2.Append(extents2);

            A.Graphic graphic2 = new A.Graphic();

            A.GraphicData graphicData2 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

            C.ChartReference chartReference2 = new C.ChartReference() { Id = "rId2" };
            chartReference2.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartReference2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            graphicData2.Append(chartReference2);

            graphic2.Append(graphicData2);

            graphicFrame2.Append(nonVisualGraphicFrameProperties2);
            graphicFrame2.Append(transform2);
            graphicFrame2.Append(graphic2);
            Xdr.ClientData clientData2 = new Xdr.ClientData();

            twoCellAnchor2.Append(fromMarker2);
            twoCellAnchor2.Append(toMarker2);
            twoCellAnchor2.Append(graphicFrame2);
            twoCellAnchor2.Append(clientData2);



            //THIRD GRAPH
            Xdr.TwoCellAnchor twoCellAnchor3 = new Xdr.TwoCellAnchor();

            Xdr.FromMarker fromMarker3 = new Xdr.FromMarker();
            Xdr.ColumnId columnId5 = new Xdr.ColumnId();
            columnId5.Text = "6";
            Xdr.ColumnOffset columnOffset5 = new Xdr.ColumnOffset();
            columnOffset5.Text = "0";
            Xdr.RowId rowId5 = new Xdr.RowId();
            //rowId5.Text = "29";
            rowId5.Text = (3 + ((sectionHeight + sectionSpacing) * 2)).ToString();
            Xdr.RowOffset rowOffset5 = new Xdr.RowOffset();
            rowOffset5.Text = "0";

            fromMarker3.Append(columnId5);
            fromMarker3.Append(columnOffset5);
            fromMarker3.Append(rowId5);
            fromMarker3.Append(rowOffset5);

            Xdr.ToMarker toMarker3 = new Xdr.ToMarker();
            Xdr.ColumnId columnId6 = new Xdr.ColumnId();
            columnId6.Text = "12";
            Xdr.ColumnOffset columnOffset6 = new Xdr.ColumnOffset();
            columnOffset6.Text = "0";
            Xdr.RowId rowId6 = new Xdr.RowId();
            //rowId6.Text = "39";
            rowId6.Text = (3 + ((sectionHeight + sectionSpacing) * 2) + graphHeight).ToString();
            Xdr.RowOffset rowOffset6 = new Xdr.RowOffset();
            rowOffset6.Text = "0";

            toMarker3.Append(columnId6);
            toMarker3.Append(columnOffset6);
            toMarker3.Append(rowId6);
            toMarker3.Append(rowOffset6);

            Xdr.GraphicFrame graphicFrame3 = new Xdr.GraphicFrame() { Macro = "" };

            Xdr.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties3 = new Xdr.NonVisualGraphicFrameProperties();
            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties3 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Chart 3" };
            Xdr.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties3 = new Xdr.NonVisualGraphicFrameDrawingProperties();

            nonVisualGraphicFrameProperties3.Append(nonVisualDrawingProperties3);
            nonVisualGraphicFrameProperties3.Append(nonVisualGraphicFrameDrawingProperties3);

            Xdr.Transform transform3 = new Xdr.Transform();
            A.Offset offset3 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents3 = new A.Extents() { Cx = 0L, Cy = 0L };

            transform3.Append(offset3);
            transform3.Append(extents3);

            A.Graphic graphic3 = new A.Graphic();

            A.GraphicData graphicData3 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

            C.ChartReference chartReference3 = new C.ChartReference() { Id = "rId3" };
            chartReference3.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartReference3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            graphicData3.Append(chartReference3);

            graphic3.Append(graphicData3);

            graphicFrame3.Append(nonVisualGraphicFrameProperties3);
            graphicFrame3.Append(transform3);
            graphicFrame3.Append(graphic3);
            Xdr.ClientData clientData3 = new Xdr.ClientData();

            twoCellAnchor3.Append(fromMarker3);
            twoCellAnchor3.Append(toMarker3);
            twoCellAnchor3.Append(graphicFrame3);
            twoCellAnchor3.Append(clientData3);


            
            
            
            //FOURTH GRAPH
            Xdr.TwoCellAnchor twoCellAnchor4 = new Xdr.TwoCellAnchor();

            Xdr.FromMarker fromMarker4 = new Xdr.FromMarker();
            Xdr.ColumnId columnId7 = new Xdr.ColumnId();
            columnId7.Text = "6";
            Xdr.ColumnOffset columnOffset7 = new Xdr.ColumnOffset();
            //columnOffset7.Text = "76200";
            columnOffset7.Text = "0";
            Xdr.RowId rowId7 = new Xdr.RowId();
            //rowId7.Text = "42";
            rowId7.Text = (3 + ((sectionHeight + sectionSpacing) * 3)).ToString();
            Xdr.RowOffset rowOffset7 = new Xdr.RowOffset();
            rowOffset7.Text = "0";

            fromMarker4.Append(columnId7);
            fromMarker4.Append(columnOffset7);
            fromMarker4.Append(rowId7);
            fromMarker4.Append(rowOffset7);

            Xdr.ToMarker toMarker4 = new Xdr.ToMarker();
            Xdr.ColumnId columnId8 = new Xdr.ColumnId();
            columnId8.Text = "12";
            Xdr.ColumnOffset columnOffset8 = new Xdr.ColumnOffset();
            columnOffset8.Text = "0";
            Xdr.RowId rowId8 = new Xdr.RowId();
            //rowId8.Text = "52";
            rowId8.Text = (3 + ((sectionHeight + sectionSpacing) * 3) + graphHeight).ToString();
            Xdr.RowOffset rowOffset8 = new Xdr.RowOffset();
            rowOffset8.Text = "0";

            toMarker4.Append(columnId8);
            toMarker4.Append(columnOffset8);
            toMarker4.Append(rowId8);
            toMarker4.Append(rowOffset8);

            Xdr.GraphicFrame graphicFrame4 = new Xdr.GraphicFrame() { Macro = "" };

            Xdr.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties4 = new Xdr.NonVisualGraphicFrameProperties();
            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties4 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Chart 4" };
            Xdr.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties4 = new Xdr.NonVisualGraphicFrameDrawingProperties();

            nonVisualGraphicFrameProperties4.Append(nonVisualDrawingProperties4);
            nonVisualGraphicFrameProperties4.Append(nonVisualGraphicFrameDrawingProperties4);

            Xdr.Transform transform4 = new Xdr.Transform();
            A.Offset offset4 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents4 = new A.Extents() { Cx = 0L, Cy = 0L };

            transform4.Append(offset4);
            transform4.Append(extents4);

            A.Graphic graphic4 = new A.Graphic();

            A.GraphicData graphicData4 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

            C.ChartReference chartReference4 = new C.ChartReference() { Id = "rId4" };
            chartReference4.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartReference4.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            graphicData4.Append(chartReference4);

            graphic4.Append(graphicData4);

            graphicFrame4.Append(nonVisualGraphicFrameProperties4);
            graphicFrame4.Append(transform4);
            graphicFrame4.Append(graphic4);
            Xdr.ClientData clientData4 = new Xdr.ClientData();

            twoCellAnchor4.Append(fromMarker4);
            twoCellAnchor4.Append(toMarker4);
            twoCellAnchor4.Append(graphicFrame4);
            twoCellAnchor4.Append(clientData4);

            //FIFTH GRAPH
            Xdr.TwoCellAnchor twoCellAnchor5 = new Xdr.TwoCellAnchor();

            Xdr.FromMarker fromMarker5 = new Xdr.FromMarker();
            Xdr.ColumnId columnId9 = new Xdr.ColumnId();
            //columnId9.Text = "6";
            //columnId9.Text = "8";
            columnId9.Text = "6";
            Xdr.ColumnOffset columnOffset9 = new Xdr.ColumnOffset();
            //columnOffset9.Text = "66675";
            //columnOffset9.Text = "266700";
            columnOffset9.Text = "0";
            Xdr.RowId rowId9 = new Xdr.RowId();
            //rowId9.Text = "54";
            //rowId9.Text = "55";
            rowId9.Text = (3 + ((sectionHeight + sectionSpacing) * 4)).ToString();
            Xdr.RowOffset rowOffset9 = new Xdr.RowOffset();
            //rowOffset9.Text = "180975";
            rowOffset9.Text = "0";

            fromMarker5.Append(columnId9);
            fromMarker5.Append(columnOffset9);
            fromMarker5.Append(rowId9);
            fromMarker5.Append(rowOffset9);

            Xdr.ToMarker toMarker5 = new Xdr.ToMarker();
            Xdr.ColumnId columnId10 = new Xdr.ColumnId();
            //columnId10.Text = "11";
            columnId10.Text = "12";
            Xdr.ColumnOffset columnOffset10 = new Xdr.ColumnOffset();
            //columnOffset10.Text = "476249";
            columnOffset10.Text = "0";
            Xdr.RowId rowId10 = new Xdr.RowId();
            //rowId10.Text = "65";
            //rowId10.Text = "65";
            rowId10.Text = (3 + ((sectionHeight + sectionSpacing) * 4) + graphHeight).ToString();
            Xdr.RowOffset rowOffset10 = new Xdr.RowOffset();
            //rowOffset10.Text = "9524";
            rowOffset10.Text = "0";

            toMarker5.Append(columnId10);
            toMarker5.Append(columnOffset10);
            toMarker5.Append(rowId10);
            toMarker5.Append(rowOffset10);

            Xdr.GraphicFrame graphicFrame5 = new Xdr.GraphicFrame() { Macro = "" };

            Xdr.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties5 = new Xdr.NonVisualGraphicFrameProperties();
            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties5 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Chart 5" };
            Xdr.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties5 = new Xdr.NonVisualGraphicFrameDrawingProperties();

            nonVisualGraphicFrameProperties5.Append(nonVisualDrawingProperties5);
            nonVisualGraphicFrameProperties5.Append(nonVisualGraphicFrameDrawingProperties5);

            Xdr.Transform transform5 = new Xdr.Transform();
            A.Offset offset5 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents5 = new A.Extents() { Cx = 0L, Cy = 0L };

            transform5.Append(offset5);
            transform5.Append(extents5);

            A.Graphic graphic5 = new A.Graphic();

            A.GraphicData graphicData5 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

            C.ChartReference chartReference5 = new C.ChartReference() { Id = "rId5" };
            chartReference5.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartReference5.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            graphicData5.Append(chartReference5);

            graphic5.Append(graphicData5);

            graphicFrame5.Append(nonVisualGraphicFrameProperties5);
            graphicFrame5.Append(transform5);
            graphicFrame5.Append(graphic5);
            Xdr.ClientData clientData5 = new Xdr.ClientData();

            twoCellAnchor5.Append(fromMarker5);
            twoCellAnchor5.Append(toMarker5);
            twoCellAnchor5.Append(graphicFrame5);
            twoCellAnchor5.Append(clientData5);







            worksheetDrawing1.Append(twoCellAnchor1);
            worksheetDrawing1.Append(twoCellAnchor2);
            worksheetDrawing1.Append(twoCellAnchor3);
            worksheetDrawing1.Append(twoCellAnchor4);
            worksheetDrawing1.Append(twoCellAnchor5);

            drawingsPart1.WorksheetDrawing = worksheetDrawing1;
        }
        #endregion

        #region GenerateComparisonResultsImpressionsPieChartContent - MODIFY
        // Generates content of chartPart1.
        private void GenerateComparisonResultsImpressionsPieChartContent(ChartPart chartPart1)
        {
            C.ChartSpace chartSpace1 = new C.ChartSpace();
            chartSpace1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartSpace1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            chartSpace1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            C.EditingLanguage editingLanguage1 = new C.EditingLanguage() { Val = "en-US" };

            C.Chart chart1 = new C.Chart();

            C.PlotArea plotArea1 = new C.PlotArea();

            C.Layout layout1 = new C.Layout();

            C.ManualLayout manualLayout1 = new C.ManualLayout();
            C.LayoutTarget layoutTarget1 = new C.LayoutTarget() { Val = C.LayoutTargetValues.Inner };
            C.LeftMode leftMode1 = new C.LeftMode() { Val = C.LayoutModeValues.Edge };
            C.TopMode topMode1 = new C.TopMode() { Val = C.LayoutModeValues.Edge };
            C.Left left1 = new C.Left() { Val = 4.2408147613687039E-2D };
            C.Top top1 = new C.Top() { Val = 0.1256544502617801D };
            C.Width width1 = new C.Width() { Val = 0.40775681341719078D };
            C.Height height1 = new C.Height() { Val = 0.67888307155322858D };

            manualLayout1.Append(layoutTarget1);
            manualLayout1.Append(leftMode1);
            manualLayout1.Append(topMode1);
            manualLayout1.Append(left1);
            manualLayout1.Append(top1);
            manualLayout1.Append(width1);
            manualLayout1.Append(height1);

            layout1.Append(manualLayout1);

            C.PieChart pieChart1 = new C.PieChart();
            C.VaryColors varyColors1 = new C.VaryColors() { Val = true };

            C.PieChartSeries pieChartSeries1 = new C.PieChartSeries();
            C.Index index1 = new C.Index() { Val = (UInt32Value)0U };
            C.Order order1 = new C.Order() { Val = (UInt32Value)0U };

            C.DataLabels dataLabels1 = new C.DataLabels();
            C.ShowPercent showPercent1 = new C.ShowPercent() { Val = true };
            C.ShowLeaderLines showLeaderLines1 = new C.ShowLeaderLines() { Val = true };

            dataLabels1.Append(showPercent1);
            dataLabels1.Append(showLeaderLines1);

            C.CategoryAxisData categoryAxisData1 = new C.CategoryAxisData();

            C.StringReference stringReference1 = new C.StringReference();
            C.Formula formula1 = new C.Formula();
            //formula1.Text = "\'Vendor Analytics\'!$A$19:$A$26";
            formula1.Text = comparisonResultsImpressionsLabelRange;

            C.StringCache stringCache1 = new C.StringCache();
            C.PointCount pointCount1 = new C.PointCount() { Val = (UInt32Value)8U };

            C.StringPoint stringPoint1 = new C.StringPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue1 = new C.NumericValue();
            numericValue1.Text = "Sales Cloud Contact Manager";

            stringPoint1.Append(numericValue1);

            C.StringPoint stringPoint2 = new C.StringPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue2 = new C.NumericValue();
            numericValue2.Text = "Sales Cloud Group";

            stringPoint2.Append(numericValue2);

            C.StringPoint stringPoint3 = new C.StringPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue3 = new C.NumericValue();
            numericValue3.Text = "Sales Cloud Professional";

            stringPoint3.Append(numericValue3);

            C.StringPoint stringPoint4 = new C.StringPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue4 = new C.NumericValue();
            numericValue4.Text = "Sales Cloud Enterprise";

            stringPoint4.Append(numericValue4);

            C.StringPoint stringPoint5 = new C.StringPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue5 = new C.NumericValue();
            numericValue5.Text = "Sales Cloud Unlimited";

            stringPoint5.Append(numericValue5);

            C.StringPoint stringPoint6 = new C.StringPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue6 = new C.NumericValue();
            numericValue6.Text = "Service Cloud Professional";

            stringPoint6.Append(numericValue6);

            C.StringPoint stringPoint7 = new C.StringPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue7 = new C.NumericValue();
            numericValue7.Text = "Service Cloud Enterprise";

            stringPoint7.Append(numericValue7);

            C.StringPoint stringPoint8 = new C.StringPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue8 = new C.NumericValue();
            numericValue8.Text = "Service Cloud Unlimited";

            stringPoint8.Append(numericValue8);

            stringCache1.Append(pointCount1);
            stringCache1.Append(stringPoint1);
            stringCache1.Append(stringPoint2);
            stringCache1.Append(stringPoint3);
            stringCache1.Append(stringPoint4);
            stringCache1.Append(stringPoint5);
            stringCache1.Append(stringPoint6);
            stringCache1.Append(stringPoint7);
            stringCache1.Append(stringPoint8);

            stringReference1.Append(formula1);
            stringReference1.Append(stringCache1);

            categoryAxisData1.Append(stringReference1);

            C.Values values1 = new C.Values();

            C.NumberReference numberReference1 = new C.NumberReference();
            C.Formula formula2 = new C.Formula();
            //formula2.Text = "\'Vendor Analytics\'!$B$19:$B$26";
            formula2.Text = comparisonResultsImpressionsDataRange;

            C.NumberingCache numberingCache1 = new C.NumberingCache();
            C.FormatCode formatCode1 = new C.FormatCode();
            formatCode1.Text = "General";
            C.PointCount pointCount2 = new C.PointCount() { Val = (UInt32Value)8U };

            C.NumericPoint numericPoint1 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue9 = new C.NumericValue();
            numericValue9.Text = "5";

            numericPoint1.Append(numericValue9);

            C.NumericPoint numericPoint2 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue10 = new C.NumericValue();
            numericValue10.Text = "6";

            numericPoint2.Append(numericValue10);

            C.NumericPoint numericPoint3 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue11 = new C.NumericValue();
            numericValue11.Text = "6";

            numericPoint3.Append(numericValue11);

            C.NumericPoint numericPoint4 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue12 = new C.NumericValue();
            numericValue12.Text = "4";

            numericPoint4.Append(numericValue12);

            C.NumericPoint numericPoint5 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue13 = new C.NumericValue();
            numericValue13.Text = "8";

            numericPoint5.Append(numericValue13);

            C.NumericPoint numericPoint6 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue14 = new C.NumericValue();
            numericValue14.Text = "10";

            numericPoint6.Append(numericValue14);

            C.NumericPoint numericPoint7 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue15 = new C.NumericValue();
            numericValue15.Text = "12";

            numericPoint7.Append(numericValue15);

            C.NumericPoint numericPoint8 = new C.NumericPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue16 = new C.NumericValue();
            numericValue16.Text = "13";

            numericPoint8.Append(numericValue16);

            numberingCache1.Append(formatCode1);
            numberingCache1.Append(pointCount2);
            numberingCache1.Append(numericPoint1);
            numberingCache1.Append(numericPoint2);
            numberingCache1.Append(numericPoint3);
            numberingCache1.Append(numericPoint4);
            numberingCache1.Append(numericPoint5);
            numberingCache1.Append(numericPoint6);
            numberingCache1.Append(numericPoint7);
            numberingCache1.Append(numericPoint8);

            numberReference1.Append(formula2);
            numberReference1.Append(numberingCache1);

            values1.Append(numberReference1);

            pieChartSeries1.Append(index1);
            pieChartSeries1.Append(order1);
            pieChartSeries1.Append(dataLabels1);
            pieChartSeries1.Append(categoryAxisData1);
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

            C.ManualLayout manualLayout2 = new C.ManualLayout();
            C.LeftMode leftMode2 = new C.LeftMode() { Val = C.LayoutModeValues.Edge };
            C.TopMode topMode2 = new C.TopMode() { Val = C.LayoutModeValues.Edge };
            C.Left left2 = new C.Left() { Val = 0.52736050586889915D };
            C.Top top2 = new C.Top() { Val = 6.4585653398003817E-2D };
            C.Width width2 = new C.Width() { Val = 0.46006074712359069D };
            C.Height height2 = new C.Height() { Val = 0.88479036979016368D };

            manualLayout2.Append(leftMode2);
            manualLayout2.Append(topMode2);
            manualLayout2.Append(left2);
            manualLayout2.Append(top2);
            manualLayout2.Append(width2);
            manualLayout2.Append(height2);

            layout2.Append(manualLayout2);
            C.Overlay overlay1 = new C.Overlay() { Val = true };

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
            legend1.Append(overlay1);
            legend1.Append(textProperties1);
            C.PlotVisibleOnly plotVisibleOnly1 = new C.PlotVisibleOnly() { Val = true };

            chart1.Append(plotArea1);
            chart1.Append(legend1);
            chart1.Append(plotVisibleOnly1);

            //C.PrintSettings printSettings1 = new C.PrintSettings();
            C.HeaderFooter headerFooter1 = new C.HeaderFooter();
            C.PageMargins pageMargins2 = new C.PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            C.PageSetup pageSetup1 = new C.PageSetup();

            //printSettings1.Append(headerFooter1);
            //printSettings1.Append(pageMargins2);
            //printSettings1.Append(pageSetup1);

            chartSpace1.Append(editingLanguage1);
            chartSpace1.Append(chart1);
            //chartSpace1.Append(printSettings1);

            chartPart1.ChartSpace = chartSpace1;
        }
        #endregion

        #region GenerateImpressionsPieChartContent - MODIFY
        // Generates content of chartPart2.
        private void GenerateImpressionsPieChartContent(ChartPart chartPart2)
        {
            C.ChartSpace chartSpace2 = new C.ChartSpace();
            chartSpace2.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartSpace2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            chartSpace2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            C.EditingLanguage editingLanguage2 = new C.EditingLanguage() { Val = "en-US" };

            C.Chart chart2 = new C.Chart();

            C.PlotArea plotArea2 = new C.PlotArea();

            C.Layout layout3 = new C.Layout();

            C.ManualLayout manualLayout3 = new C.ManualLayout();
            C.LayoutTarget layoutTarget2 = new C.LayoutTarget() { Val = C.LayoutTargetValues.Inner };
            C.LeftMode leftMode3 = new C.LeftMode() { Val = C.LayoutModeValues.Edge };
            C.TopMode topMode3 = new C.TopMode() { Val = C.LayoutModeValues.Edge };
            C.Left left3 = new C.Left() { Val = 4.2408147613687053E-2D };
            C.Top top3 = new C.Top() { Val = 0.1256544502617801D };
            C.Width width3 = new C.Width() { Val = 0.40775681341719078D };
            C.Height height3 = new C.Height() { Val = 0.6788830715532288D };

            manualLayout3.Append(layoutTarget2);
            manualLayout3.Append(leftMode3);
            manualLayout3.Append(topMode3);
            manualLayout3.Append(left3);
            manualLayout3.Append(top3);
            manualLayout3.Append(width3);
            manualLayout3.Append(height3);

            layout3.Append(manualLayout3);

            C.PieChart pieChart2 = new C.PieChart();
            C.VaryColors varyColors2 = new C.VaryColors() { Val = true };

            C.PieChartSeries pieChartSeries2 = new C.PieChartSeries();
            C.Index index2 = new C.Index() { Val = (UInt32Value)0U };
            C.Order order2 = new C.Order() { Val = (UInt32Value)0U };

            C.DataLabels dataLabels2 = new C.DataLabels();
            C.ShowPercent showPercent2 = new C.ShowPercent() { Val = true };
            C.ShowLeaderLines showLeaderLines2 = new C.ShowLeaderLines() { Val = true };

            dataLabels2.Append(showPercent2);
            dataLabels2.Append(showLeaderLines2);

            C.CategoryAxisData categoryAxisData2 = new C.CategoryAxisData();

            C.StringReference stringReference2 = new C.StringReference();
            C.Formula formula3 = new C.Formula();
            //formula3.Text = "\'Vendor Analytics\'!$A$6:$A$13";
            formula3.Text = impressionsLabelRange;

            C.StringCache stringCache2 = new C.StringCache();
            C.PointCount pointCount3 = new C.PointCount() { Val = (UInt32Value)8U };

            //C.StringPoint stringPoint9 = new C.StringPoint() { Index = (UInt32Value)0U };
            //C.NumericValue numericValue17 = new C.NumericValue();
            //numericValue17.Text = "Sales Cloud Contact Manager";
            //stringPoint9.Append(numericValue17);

            //C.StringPoint stringPoint10 = new C.StringPoint() { Index = (UInt32Value)1U };
            //C.NumericValue numericValue18 = new C.NumericValue();
            //numericValue18.Text = "Sales Cloud Group";
            //stringPoint10.Append(numericValue18);

            //C.StringPoint stringPoint11 = new C.StringPoint() { Index = (UInt32Value)2U };
            //C.NumericValue numericValue19 = new C.NumericValue();
            //numericValue19.Text = "Sales Cloud Professional";
            //stringPoint11.Append(numericValue19);

            //C.StringPoint stringPoint12 = new C.StringPoint() { Index = (UInt32Value)3U };
            //C.NumericValue numericValue20 = new C.NumericValue();
            //numericValue20.Text = "Sales Cloud Enterprise";
            //stringPoint12.Append(numericValue20);

            //C.StringPoint stringPoint13 = new C.StringPoint() { Index = (UInt32Value)4U };
            //C.NumericValue numericValue21 = new C.NumericValue();
            //numericValue21.Text = "Sales Cloud Unlimited";
            //stringPoint13.Append(numericValue21);

            //C.StringPoint stringPoint14 = new C.StringPoint() { Index = (UInt32Value)5U };
            //C.NumericValue numericValue22 = new C.NumericValue();
            //numericValue22.Text = "Service Cloud Professional";
            //stringPoint14.Append(numericValue22);

            //C.StringPoint stringPoint15 = new C.StringPoint() { Index = (UInt32Value)6U };
            //C.NumericValue numericValue23 = new C.NumericValue();
            //numericValue23.Text = "Service Cloud Enterprise";
            //stringPoint15.Append(numericValue23);

            //C.StringPoint stringPoint16 = new C.StringPoint() { Index = (UInt32Value)7U };
            //C.NumericValue numericValue24 = new C.NumericValue();
            //numericValue24.Text = "Service Cloud Unlimited";
            //stringPoint16.Append(numericValue24);

            //stringCache2.Append(pointCount3);
            //stringCache2.Append(stringPoint9);
            //stringCache2.Append(stringPoint10);
            //stringCache2.Append(stringPoint11);
            //stringCache2.Append(stringPoint12);
            //stringCache2.Append(stringPoint13);
            //stringCache2.Append(stringPoint14);
            //stringCache2.Append(stringPoint15);
            //stringCache2.Append(stringPoint16);

            stringReference2.Append(formula3);
            //stringReference2.Append(stringCache2);

            categoryAxisData2.Append(stringReference2);

            C.Values values2 = new C.Values();

            C.NumberReference numberReference2 = new C.NumberReference();
            C.Formula formula4 = new C.Formula();
            //formula4.Text = "\'Vendor Analytics\'!$B$6:$B$13";
            formula4.Text = impressionsDataRange;

            C.NumberingCache numberingCache2 = new C.NumberingCache();
            C.FormatCode formatCode2 = new C.FormatCode();
            formatCode2.Text = "General";
            C.PointCount pointCount4 = new C.PointCount() { Val = (UInt32Value)8U };

            //C.NumericPoint numericPoint9 = new C.NumericPoint() { Index = (UInt32Value)0U };
            //C.NumericValue numericValue25 = new C.NumericValue();
            //numericValue25.Text = "829";
            //numericPoint9.Append(numericValue25);

            //C.NumericPoint numericPoint10 = new C.NumericPoint() { Index = (UInt32Value)1U };
            //C.NumericValue numericValue26 = new C.NumericValue();
            //numericValue26.Text = "829";
            //numericPoint10.Append(numericValue26);

            //C.NumericPoint numericPoint11 = new C.NumericPoint() { Index = (UInt32Value)2U };
            //C.NumericValue numericValue27 = new C.NumericValue();
            //numericValue27.Text = "806";
            //numericPoint11.Append(numericValue27);

            //C.NumericPoint numericPoint12 = new C.NumericPoint() { Index = (UInt32Value)3U };
            //C.NumericValue numericValue28 = new C.NumericValue();
            //numericValue28.Text = "820";
            //numericPoint12.Append(numericValue28);

            //C.NumericPoint numericPoint13 = new C.NumericPoint() { Index = (UInt32Value)4U };
            //C.NumericValue numericValue29 = new C.NumericValue();
            //numericValue29.Text = "806";
            //numericPoint13.Append(numericValue29);

            //C.NumericPoint numericPoint14 = new C.NumericPoint() { Index = (UInt32Value)5U };
            //C.NumericValue numericValue30 = new C.NumericValue();
            //numericValue30.Text = "834";
            //numericPoint14.Append(numericValue30);

            //C.NumericPoint numericPoint15 = new C.NumericPoint() { Index = (UInt32Value)6U };
            //C.NumericValue numericValue31 = new C.NumericValue();
            //numericValue31.Text = "822";
            //numericPoint15.Append(numericValue31);

            //C.NumericPoint numericPoint16 = new C.NumericPoint() { Index = (UInt32Value)7U };
            //C.NumericValue numericValue32 = new C.NumericValue();
            //numericValue32.Text = "834";
            //numericPoint16.Append(numericValue32);

            numberingCache2.Append(formatCode2);
            numberingCache2.Append(pointCount4);
            //numberingCache2.Append(numericPoint9);
            //numberingCache2.Append(numericPoint10);
            //numberingCache2.Append(numericPoint11);
            //numberingCache2.Append(numericPoint12);
            //numberingCache2.Append(numericPoint13);
            //numberingCache2.Append(numericPoint14);
            //numberingCache2.Append(numericPoint15);
            //numberingCache2.Append(numericPoint16);

            numberReference2.Append(formula4);
            numberReference2.Append(numberingCache2);

            values2.Append(numberReference2);

            pieChartSeries2.Append(index2);
            pieChartSeries2.Append(order2);
            pieChartSeries2.Append(dataLabels2);
            pieChartSeries2.Append(categoryAxisData2);
            pieChartSeries2.Append(values2);
            C.FirstSliceAngle firstSliceAngle2 = new C.FirstSliceAngle() { Val = (UInt16Value)0U };

            pieChart2.Append(varyColors2);
            pieChart2.Append(pieChartSeries2);
            pieChart2.Append(firstSliceAngle2);

            plotArea2.Append(layout3);
            plotArea2.Append(pieChart2);

            C.Legend legend2 = new C.Legend();
            C.LegendPosition legendPosition2 = new C.LegendPosition() { Val = C.LegendPositionValues.Right };

            C.Layout layout4 = new C.Layout();

            C.ManualLayout manualLayout4 = new C.ManualLayout();
            C.LeftMode leftMode4 = new C.LeftMode() { Val = C.LayoutModeValues.Edge };
            C.TopMode topMode4 = new C.TopMode() { Val = C.LayoutModeValues.Edge };
            C.Left left4 = new C.Left() { Val = 0.52736050586889904D };
            C.Top top4 = new C.Top() { Val = 6.4585653398003831E-2D };
            C.Width width4 = new C.Width() { Val = 0.46006074712359074D };
            C.Height height4 = new C.Height() { Val = 0.88479036979016357D };

            manualLayout4.Append(leftMode4);
            manualLayout4.Append(topMode4);
            manualLayout4.Append(left4);
            manualLayout4.Append(top4);
            manualLayout4.Append(width4);
            manualLayout4.Append(height4);

            layout4.Append(manualLayout4);
            C.Overlay overlay2 = new C.Overlay() { Val = true };

            C.TextProperties textProperties2 = new C.TextProperties();
            A.BodyProperties bodyProperties2 = new A.BodyProperties();
            A.ListStyle listStyle2 = new A.ListStyle();

            A.Paragraph paragraph2 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties() { RightToLeft = false };
            A.DefaultRunProperties defaultRunProperties2 = new A.DefaultRunProperties();

            paragraphProperties2.Append(defaultRunProperties2);
            A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(endParagraphRunProperties2);

            textProperties2.Append(bodyProperties2);
            textProperties2.Append(listStyle2);
            textProperties2.Append(paragraph2);

            legend2.Append(legendPosition2);
            legend2.Append(layout4);
            legend2.Append(overlay2);
            legend2.Append(textProperties2);
            C.PlotVisibleOnly plotVisibleOnly2 = new C.PlotVisibleOnly() { Val = true };

            chart2.Append(plotArea2);
            chart2.Append(legend2);
            chart2.Append(plotVisibleOnly2);

            C.PrintSettings printSettings2 = new C.PrintSettings();
            C.HeaderFooter headerFooter2 = new C.HeaderFooter();
            C.PageMargins pageMargins3 = new C.PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            C.PageSetup pageSetup2 = new C.PageSetup();

            printSettings2.Append(headerFooter2);
            printSettings2.Append(pageMargins3);
            printSettings2.Append(pageSetup2);

            chartSpace2.Append(editingLanguage2);
            chartSpace2.Append(chart2);
            chartSpace2.Append(printSettings2);

            chartPart2.ChartSpace = chartSpace2;
        }
        #endregion

        #region GenerateShopContentConsumptionPieChartContent
        // Generates content of chartPart5.
        private void GenerateShopContentConsumptionPieChartContent(ChartPart chartPart5)
        {
            C.ChartSpace chartSpace5 = new C.ChartSpace();
            chartSpace5.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartSpace5.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            chartSpace5.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            C.EditingLanguage editingLanguage5 = new C.EditingLanguage() { Val = "en-US" };

            C.Chart chart5 = new C.Chart();

            C.PlotArea plotArea5 = new C.PlotArea();

            C.Layout layout9 = new C.Layout();

            C.ManualLayout manualLayout9 = new C.ManualLayout();
            C.LayoutTarget layoutTarget5 = new C.LayoutTarget() { Val = C.LayoutTargetValues.Inner };
            C.LeftMode leftMode9 = new C.LeftMode() { Val = C.LayoutModeValues.Edge };
            C.TopMode topMode9 = new C.TopMode() { Val = C.LayoutModeValues.Edge };
            C.Left left9 = new C.Left() { Val = 4.2408147613687046E-2D };
            C.Top top9 = new C.Top() { Val = 0.1256544502617801D };
            C.Width width9 = new C.Width() { Val = 0.40775681341719078D };
            C.Height height9 = new C.Height() { Val = 0.67888307155322902D };

            manualLayout9.Append(layoutTarget5);
            manualLayout9.Append(leftMode9);
            manualLayout9.Append(topMode9);
            manualLayout9.Append(left9);
            manualLayout9.Append(top9);
            manualLayout9.Append(width9);
            manualLayout9.Append(height9);

            layout9.Append(manualLayout9);

            C.PieChart pieChart5 = new C.PieChart();
            C.VaryColors varyColors5 = new C.VaryColors() { Val = true };

            C.PieChartSeries pieChartSeries5 = new C.PieChartSeries();
            C.Index index5 = new C.Index() { Val = (UInt32Value)0U };
            C.Order order5 = new C.Order() { Val = (UInt32Value)0U };

            C.DataLabels dataLabels5 = new C.DataLabels();
            C.ShowPercent showPercent5 = new C.ShowPercent() { Val = true };
            C.ShowLeaderLines showLeaderLines5 = new C.ShowLeaderLines() { Val = true };

            dataLabels5.Append(showPercent5);
            dataLabels5.Append(showLeaderLines5);

            C.CategoryAxisData categoryAxisData5 = new C.CategoryAxisData();

            C.StringReference stringReference5 = new C.StringReference();
            C.Formula formula9 = new C.Formula();
            //formula9.Text = "\'Vendor Analytics\'!$A$45:$A$52";
            formula9.Text = shopContentConsumptionLabelRange;

            C.StringCache stringCache5 = new C.StringCache();
            C.PointCount pointCount9 = new C.PointCount() { Val = (UInt32Value)8U };

            C.StringPoint stringPoint33 = new C.StringPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue65 = new C.NumericValue();
            numericValue65.Text = "Sales Cloud Contact Manager";

            stringPoint33.Append(numericValue65);

            C.StringPoint stringPoint34 = new C.StringPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue66 = new C.NumericValue();
            numericValue66.Text = "Sales Cloud Group";

            stringPoint34.Append(numericValue66);

            C.StringPoint stringPoint35 = new C.StringPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue67 = new C.NumericValue();
            numericValue67.Text = "Sales Cloud Professional";

            stringPoint35.Append(numericValue67);

            C.StringPoint stringPoint36 = new C.StringPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue68 = new C.NumericValue();
            numericValue68.Text = "Sales Cloud Enterprise";

            stringPoint36.Append(numericValue68);

            C.StringPoint stringPoint37 = new C.StringPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue69 = new C.NumericValue();
            numericValue69.Text = "Sales Cloud Unlimited";

            stringPoint37.Append(numericValue69);

            C.StringPoint stringPoint38 = new C.StringPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue70 = new C.NumericValue();
            numericValue70.Text = "Service Cloud Professional";

            stringPoint38.Append(numericValue70);

            C.StringPoint stringPoint39 = new C.StringPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue71 = new C.NumericValue();
            numericValue71.Text = "Service Cloud Enterprise";

            stringPoint39.Append(numericValue71);

            C.StringPoint stringPoint40 = new C.StringPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue72 = new C.NumericValue();
            numericValue72.Text = "Service Cloud Unlimited";

            stringPoint40.Append(numericValue72);

            stringCache5.Append(pointCount9);
            stringCache5.Append(stringPoint33);
            stringCache5.Append(stringPoint34);
            stringCache5.Append(stringPoint35);
            stringCache5.Append(stringPoint36);
            stringCache5.Append(stringPoint37);
            stringCache5.Append(stringPoint38);
            stringCache5.Append(stringPoint39);
            stringCache5.Append(stringPoint40);

            stringReference5.Append(formula9);
            stringReference5.Append(stringCache5);

            categoryAxisData5.Append(stringReference5);

            C.Values values5 = new C.Values();

            C.NumberReference numberReference5 = new C.NumberReference();
            C.Formula formula10 = new C.Formula();
            //formula10.Text = "\'Vendor Analytics\'!$B$45:$B$52";
            formula10.Text = shopContentConsumptionDataRange;

            C.NumberingCache numberingCache5 = new C.NumberingCache();
            C.FormatCode formatCode5 = new C.FormatCode();
            formatCode5.Text = "General";
            C.PointCount pointCount10 = new C.PointCount() { Val = (UInt32Value)8U };

            C.NumericPoint numericPoint33 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue73 = new C.NumericValue();
            numericValue73.Text = "0";

            numericPoint33.Append(numericValue73);

            C.NumericPoint numericPoint34 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue74 = new C.NumericValue();
            numericValue74.Text = "0";

            numericPoint34.Append(numericValue74);

            C.NumericPoint numericPoint35 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue75 = new C.NumericValue();
            numericValue75.Text = "0";

            numericPoint35.Append(numericValue75);

            C.NumericPoint numericPoint36 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue76 = new C.NumericValue();
            numericValue76.Text = "0";

            numericPoint36.Append(numericValue76);

            C.NumericPoint numericPoint37 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue77 = new C.NumericValue();
            numericValue77.Text = "0";

            numericPoint37.Append(numericValue77);

            C.NumericPoint numericPoint38 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue78 = new C.NumericValue();
            numericValue78.Text = "24";

            numericPoint38.Append(numericValue78);

            C.NumericPoint numericPoint39 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue79 = new C.NumericValue();
            numericValue79.Text = "23";

            numericPoint39.Append(numericValue79);

            C.NumericPoint numericPoint40 = new C.NumericPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue80 = new C.NumericValue();
            numericValue80.Text = "18";

            numericPoint40.Append(numericValue80);

            numberingCache5.Append(formatCode5);
            numberingCache5.Append(pointCount10);
            numberingCache5.Append(numericPoint33);
            numberingCache5.Append(numericPoint34);
            numberingCache5.Append(numericPoint35);
            numberingCache5.Append(numericPoint36);
            numberingCache5.Append(numericPoint37);
            numberingCache5.Append(numericPoint38);
            numberingCache5.Append(numericPoint39);
            numberingCache5.Append(numericPoint40);

            numberReference5.Append(formula10);
            numberReference5.Append(numberingCache5);

            values5.Append(numberReference5);

            pieChartSeries5.Append(index5);
            pieChartSeries5.Append(order5);
            pieChartSeries5.Append(dataLabels5);
            pieChartSeries5.Append(categoryAxisData5);
            pieChartSeries5.Append(values5);
            C.FirstSliceAngle firstSliceAngle5 = new C.FirstSliceAngle() { Val = (UInt16Value)0U };

            pieChart5.Append(varyColors5);
            pieChart5.Append(pieChartSeries5);
            pieChart5.Append(firstSliceAngle5);

            plotArea5.Append(layout9);
            plotArea5.Append(pieChart5);

            C.Legend legend5 = new C.Legend();
            C.LegendPosition legendPosition5 = new C.LegendPosition() { Val = C.LegendPositionValues.Right };

            C.Layout layout10 = new C.Layout();

            C.ManualLayout manualLayout10 = new C.ManualLayout();
            C.LeftMode leftMode10 = new C.LeftMode() { Val = C.LayoutModeValues.Edge };
            C.TopMode topMode10 = new C.TopMode() { Val = C.LayoutModeValues.Edge };
            C.Left left10 = new C.Left() { Val = 0.52736050586889882D };
            C.Top top10 = new C.Top() { Val = 6.4585653398003831E-2D };
            C.Width width10 = new C.Width() { Val = 0.4600607471235908D };
            C.Height height10 = new C.Height() { Val = 0.88479036979016357D };

            manualLayout10.Append(leftMode10);
            manualLayout10.Append(topMode10);
            manualLayout10.Append(left10);
            manualLayout10.Append(top10);
            manualLayout10.Append(width10);
            manualLayout10.Append(height10);

            layout10.Append(manualLayout10);
            C.Overlay overlay5 = new C.Overlay() { Val = true };

            C.TextProperties textProperties5 = new C.TextProperties();
            A.BodyProperties bodyProperties5 = new A.BodyProperties();
            A.ListStyle listStyle5 = new A.ListStyle();

            A.Paragraph paragraph5 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties() { RightToLeft = false };
            A.DefaultRunProperties defaultRunProperties5 = new A.DefaultRunProperties();

            paragraphProperties5.Append(defaultRunProperties5);
            A.EndParagraphRunProperties endParagraphRunProperties5 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(endParagraphRunProperties5);

            textProperties5.Append(bodyProperties5);
            textProperties5.Append(listStyle5);
            textProperties5.Append(paragraph5);

            legend5.Append(legendPosition5);
            legend5.Append(layout10);
            legend5.Append(overlay5);
            legend5.Append(textProperties5);
            C.PlotVisibleOnly plotVisibleOnly5 = new C.PlotVisibleOnly() { Val = true };

            chart5.Append(plotArea5);
            chart5.Append(legend5);
            chart5.Append(plotVisibleOnly5);

            C.PrintSettings printSettings5 = new C.PrintSettings();
            C.HeaderFooter headerFooter5 = new C.HeaderFooter();
            C.PageMargins pageMargins6 = new C.PageMargins() { Left = 0.70000000000000018D, Right = 0.70000000000000018D, Top = 0.75000000000000022D, Bottom = 0.75000000000000022D, Header = 0.3000000000000001D, Footer = 0.3000000000000001D };
            C.PageSetup pageSetup5 = new C.PageSetup();

            printSettings5.Append(headerFooter5);
            printSettings5.Append(pageMargins6);
            printSettings5.Append(pageSetup5);

            chartSpace5.Append(editingLanguage5);
            chartSpace5.Append(chart5);
            chartSpace5.Append(printSettings5);

            chartPart5.ChartSpace = chartSpace5;
        }
        #endregion

        #region GenerateShopLeadsPieChartContent
        private void GenerateShopLeadsPieChartContent(ChartPart chartPart4)
        {
            C.ChartSpace chartSpace4 = new C.ChartSpace();
            chartSpace4.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartSpace4.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            chartSpace4.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            C.EditingLanguage editingLanguage4 = new C.EditingLanguage() { Val = "en-US" };

            C.Chart chart4 = new C.Chart();

            C.PlotArea plotArea4 = new C.PlotArea();

            C.Layout layout7 = new C.Layout();

            C.ManualLayout manualLayout7 = new C.ManualLayout();
            C.LayoutTarget layoutTarget4 = new C.LayoutTarget() { Val = C.LayoutTargetValues.Inner };
            C.LeftMode leftMode7 = new C.LeftMode() { Val = C.LayoutModeValues.Edge };
            C.TopMode topMode7 = new C.TopMode() { Val = C.LayoutModeValues.Edge };
            C.Left left7 = new C.Left() { Val = 4.2408147613687046E-2D };
            //C.Left left7 = new C.Left() { Val = 4.3408147613687053E-2D };
            C.Top top7 = new C.Top() { Val = 0.1256544502617801D };
            C.Width width7 = new C.Width() { Val = 0.40775681341719078D };
            C.Height height7 = new C.Height() { Val = 0.67888307155322902D };

            manualLayout7.Append(layoutTarget4);
            manualLayout7.Append(leftMode7);
            manualLayout7.Append(topMode7);
            manualLayout7.Append(left7);
            manualLayout7.Append(top7);
            manualLayout7.Append(width7);
            manualLayout7.Append(height7);

            layout7.Append(manualLayout7);

            C.PieChart pieChart4 = new C.PieChart();
            C.VaryColors varyColors4 = new C.VaryColors() { Val = true };

            C.PieChartSeries pieChartSeries4 = new C.PieChartSeries();
            C.Index index4 = new C.Index() { Val = (UInt32Value)0U };
            C.Order order4 = new C.Order() { Val = (UInt32Value)0U };

            C.DataLabels dataLabels4 = new C.DataLabels();
            C.ShowPercent showPercent4 = new C.ShowPercent() { Val = true };
            C.ShowLeaderLines showLeaderLines4 = new C.ShowLeaderLines() { Val = true };

            dataLabels4.Append(showPercent4);
            dataLabels4.Append(showLeaderLines4);

            C.CategoryAxisData categoryAxisData4 = new C.CategoryAxisData();

            C.StringReference stringReference4 = new C.StringReference();
            C.Formula formula7 = new C.Formula();
            //formula7.Text = "\'Vendor Analytics\'!$A$58:$A$65";
            formula7.Text = shopLeadsLabelRange;

            C.StringCache stringCache4 = new C.StringCache();
            C.PointCount pointCount7 = new C.PointCount() { Val = (UInt32Value)8U };

            C.StringPoint stringPoint25 = new C.StringPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue49 = new C.NumericValue();
            numericValue49.Text = "Sales Cloud Contact Manager";

            stringPoint25.Append(numericValue49);

            C.StringPoint stringPoint26 = new C.StringPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue50 = new C.NumericValue();
            numericValue50.Text = "Sales Cloud Group";

            stringPoint26.Append(numericValue50);

            C.StringPoint stringPoint27 = new C.StringPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue51 = new C.NumericValue();
            numericValue51.Text = "Sales Cloud Professional";

            stringPoint27.Append(numericValue51);

            C.StringPoint stringPoint28 = new C.StringPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue52 = new C.NumericValue();
            numericValue52.Text = "Sales Cloud Enterprise";

            stringPoint28.Append(numericValue52);

            C.StringPoint stringPoint29 = new C.StringPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue53 = new C.NumericValue();
            numericValue53.Text = "Sales Cloud Unlimited";

            stringPoint29.Append(numericValue53);

            C.StringPoint stringPoint30 = new C.StringPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue54 = new C.NumericValue();
            numericValue54.Text = "Service Cloud Professional";

            stringPoint30.Append(numericValue54);

            C.StringPoint stringPoint31 = new C.StringPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue55 = new C.NumericValue();
            numericValue55.Text = "Service Cloud Enterprise";

            stringPoint31.Append(numericValue55);

            C.StringPoint stringPoint32 = new C.StringPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue56 = new C.NumericValue();
            numericValue56.Text = "Service Cloud Unlimited";

            stringPoint32.Append(numericValue56);

            stringCache4.Append(pointCount7);
            stringCache4.Append(stringPoint25);
            stringCache4.Append(stringPoint26);
            stringCache4.Append(stringPoint27);
            stringCache4.Append(stringPoint28);
            stringCache4.Append(stringPoint29);
            stringCache4.Append(stringPoint30);
            stringCache4.Append(stringPoint31);
            stringCache4.Append(stringPoint32);

            stringReference4.Append(formula7);
            stringReference4.Append(stringCache4);

            categoryAxisData4.Append(stringReference4);

            C.Values values4 = new C.Values();

            C.NumberReference numberReference4 = new C.NumberReference();
            C.Formula formula8 = new C.Formula();
            //formula8.Text = "\'Vendor Analytics\'!$B$58:$B$65";
            formula8.Text = shopLeadsDataRange;

            C.NumberingCache numberingCache4 = new C.NumberingCache();
            C.FormatCode formatCode4 = new C.FormatCode();
            formatCode4.Text = "General";
            C.PointCount pointCount8 = new C.PointCount() { Val = (UInt32Value)8U };

            C.NumericPoint numericPoint25 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue57 = new C.NumericValue();
            numericValue57.Text = "0";

            numericPoint25.Append(numericValue57);

            C.NumericPoint numericPoint26 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue58 = new C.NumericValue();
            numericValue58.Text = "0";

            numericPoint26.Append(numericValue58);

            C.NumericPoint numericPoint27 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue59 = new C.NumericValue();
            numericValue59.Text = "0";

            numericPoint27.Append(numericValue59);

            C.NumericPoint numericPoint28 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue60 = new C.NumericValue();
            numericValue60.Text = "0";

            numericPoint28.Append(numericValue60);

            C.NumericPoint numericPoint29 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue61 = new C.NumericValue();
            numericValue61.Text = "0";

            numericPoint29.Append(numericValue61);

            C.NumericPoint numericPoint30 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue62 = new C.NumericValue();
            numericValue62.Text = "0";

            numericPoint30.Append(numericValue62);

            C.NumericPoint numericPoint31 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue63 = new C.NumericValue();
            numericValue63.Text = "0";

            numericPoint31.Append(numericValue63);

            C.NumericPoint numericPoint32 = new C.NumericPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue64 = new C.NumericValue();
            numericValue64.Text = "0";

            numericPoint32.Append(numericValue64);

            numberingCache4.Append(formatCode4);
            numberingCache4.Append(pointCount8);
            numberingCache4.Append(numericPoint25);
            numberingCache4.Append(numericPoint26);
            numberingCache4.Append(numericPoint27);
            numberingCache4.Append(numericPoint28);
            numberingCache4.Append(numericPoint29);
            numberingCache4.Append(numericPoint30);
            numberingCache4.Append(numericPoint31);
            numberingCache4.Append(numericPoint32);

            numberReference4.Append(formula8);
            numberReference4.Append(numberingCache4);

            values4.Append(numberReference4);

            pieChartSeries4.Append(index4);
            pieChartSeries4.Append(order4);
            pieChartSeries4.Append(dataLabels4);
            pieChartSeries4.Append(categoryAxisData4);
            pieChartSeries4.Append(values4);
            C.FirstSliceAngle firstSliceAngle4 = new C.FirstSliceAngle() { Val = (UInt16Value)0U };

            pieChart4.Append(varyColors4);
            pieChart4.Append(pieChartSeries4);
            pieChart4.Append(firstSliceAngle4);

            plotArea4.Append(layout7);
            plotArea4.Append(pieChart4);

            C.Legend legend4 = new C.Legend();
            C.LegendPosition legendPosition4 = new C.LegendPosition() { Val = C.LegendPositionValues.Right };

            C.Layout layout8 = new C.Layout();

            C.ManualLayout manualLayout8 = new C.ManualLayout();
            C.LeftMode leftMode8 = new C.LeftMode() { Val = C.LayoutModeValues.Edge };
            C.TopMode topMode8 = new C.TopMode() { Val = C.LayoutModeValues.Edge };
            C.Left left8 = new C.Left() { Val = 0.52736050586889882D };
            //C.Left left8 = new C.Left() { Val = 0.62736050586889904D };
            C.Top top8 = new C.Top() { Val = 6.4585653398003831E-2D };
            C.Width width8 = new C.Width() { Val = 0.4600607471235908D };
            C.Height height8 = new C.Height() { Val = 0.88479036979016357D };

            manualLayout8.Append(leftMode8);
            manualLayout8.Append(topMode8);
            manualLayout8.Append(left8);
            manualLayout8.Append(top8);
            manualLayout8.Append(width8);
            manualLayout8.Append(height8);

            layout8.Append(manualLayout8);
            C.Overlay overlay4 = new C.Overlay() { Val = true };

            C.TextProperties textProperties4 = new C.TextProperties();
            A.BodyProperties bodyProperties4 = new A.BodyProperties();
            A.ListStyle listStyle4 = new A.ListStyle();

            A.Paragraph paragraph4 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties() { RightToLeft = false };
            A.DefaultRunProperties defaultRunProperties4 = new A.DefaultRunProperties();

            paragraphProperties4.Append(defaultRunProperties4);
            A.EndParagraphRunProperties endParagraphRunProperties4 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(endParagraphRunProperties4);

            textProperties4.Append(bodyProperties4);
            textProperties4.Append(listStyle4);
            textProperties4.Append(paragraph4);

            legend4.Append(legendPosition4);
            legend4.Append(layout8);
            legend4.Append(overlay4);
            legend4.Append(textProperties4);
            C.PlotVisibleOnly plotVisibleOnly4 = new C.PlotVisibleOnly() { Val = true };

            chart4.Append(plotArea4);
            chart4.Append(legend4);
            chart4.Append(plotVisibleOnly4);

            C.PrintSettings printSettings4 = new C.PrintSettings();
            C.HeaderFooter headerFooter4 = new C.HeaderFooter();
            C.PageMargins pageMargins5 = new C.PageMargins() { Left = 0.70000000000000018D, Right = 0.70000000000000018D, Top = 0.75000000000000022D, Bottom = 0.75000000000000022D, Header = 0.3000000000000001D, Footer = 0.3000000000000001D };
            C.PageSetup pageSetup4 = new C.PageSetup();

            printSettings4.Append(headerFooter4);
            printSettings4.Append(pageMargins5);
            printSettings4.Append(pageSetup4);

            chartSpace4.Append(editingLanguage4);
            chartSpace4.Append(chart4);
            chartSpace4.Append(printSettings4);

            chartPart4.ChartSpace = chartSpace4;
        }
        #endregion

        #region GenerateShopVisitsPieChartContent
        private void GenerateShopVisitsPieChartContent(ChartPart chartPart1)
        {
            C.ChartSpace chartSpace1 = new C.ChartSpace();
            chartSpace1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartSpace1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            chartSpace1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            C.EditingLanguage editingLanguage1 = new C.EditingLanguage() { Val = "en-US" };

            C.Chart chart1 = new C.Chart();

            C.PlotArea plotArea1 = new C.PlotArea();

            C.Layout layout1 = new C.Layout();

            C.ManualLayout manualLayout1 = new C.ManualLayout();
            C.LayoutTarget layoutTarget1 = new C.LayoutTarget() { Val = C.LayoutTargetValues.Inner };
            C.LeftMode leftMode1 = new C.LeftMode() { Val = C.LayoutModeValues.Edge };
            C.TopMode topMode1 = new C.TopMode() { Val = C.LayoutModeValues.Edge };
            C.Left left1 = new C.Left() { Val = 4.2408147613687046E-2D };
            C.Top top1 = new C.Top() { Val = 0.1256544502617801D };
            C.Width width1 = new C.Width() { Val = 0.40775681341719078D };
            C.Height height1 = new C.Height() { Val = 0.67888307155322902D };

            manualLayout1.Append(layoutTarget1);
            manualLayout1.Append(leftMode1);
            manualLayout1.Append(topMode1);
            manualLayout1.Append(left1);
            manualLayout1.Append(top1);
            manualLayout1.Append(width1);
            manualLayout1.Append(height1);

            layout1.Append(manualLayout1);

            C.PieChart pieChart1 = new C.PieChart();
            C.VaryColors varyColors1 = new C.VaryColors() { Val = true };

            C.PieChartSeries pieChartSeries1 = new C.PieChartSeries();
            C.Index index1 = new C.Index() { Val = (UInt32Value)0U };
            C.Order order1 = new C.Order() { Val = (UInt32Value)0U };

            C.DataLabels dataLabels1 = new C.DataLabels();
            C.ShowPercent showPercent1 = new C.ShowPercent() { Val = true };
            C.ShowLeaderLines showLeaderLines1 = new C.ShowLeaderLines() { Val = true };

            dataLabels1.Append(showPercent1);
            dataLabels1.Append(showLeaderLines1);

            C.CategoryAxisData categoryAxisData1 = new C.CategoryAxisData();

            C.StringReference stringReference1 = new C.StringReference();
            C.Formula formula1 = new C.Formula();
            //formula1.Text = "\'Vendor Analytics\'!$A$32:$A$39";
            formula1.Text = shopVisitsLabelRange;

            C.StringCache stringCache1 = new C.StringCache();
            C.PointCount pointCount1 = new C.PointCount() { Val = (UInt32Value)8U };

            C.StringPoint stringPoint1 = new C.StringPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue1 = new C.NumericValue();
            numericValue1.Text = "Sales Cloud Contact Manager";

            stringPoint1.Append(numericValue1);

            C.StringPoint stringPoint2 = new C.StringPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue2 = new C.NumericValue();
            numericValue2.Text = "Sales Cloud Group";

            stringPoint2.Append(numericValue2);

            C.StringPoint stringPoint3 = new C.StringPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue3 = new C.NumericValue();
            numericValue3.Text = "Sales Cloud Professional";

            stringPoint3.Append(numericValue3);

            C.StringPoint stringPoint4 = new C.StringPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue4 = new C.NumericValue();
            numericValue4.Text = "Sales Cloud Enterprise";

            stringPoint4.Append(numericValue4);

            C.StringPoint stringPoint5 = new C.StringPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue5 = new C.NumericValue();
            numericValue5.Text = "Sales Cloud Unlimited";

            stringPoint5.Append(numericValue5);

            C.StringPoint stringPoint6 = new C.StringPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue6 = new C.NumericValue();
            numericValue6.Text = "Service Cloud Professional";

            stringPoint6.Append(numericValue6);

            C.StringPoint stringPoint7 = new C.StringPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue7 = new C.NumericValue();
            numericValue7.Text = "Service Cloud Enterprise";

            stringPoint7.Append(numericValue7);

            C.StringPoint stringPoint8 = new C.StringPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue8 = new C.NumericValue();
            numericValue8.Text = "Service Cloud Unlimited";

            stringPoint8.Append(numericValue8);

            stringCache1.Append(pointCount1);
            stringCache1.Append(stringPoint1);
            stringCache1.Append(stringPoint2);
            stringCache1.Append(stringPoint3);
            stringCache1.Append(stringPoint4);
            stringCache1.Append(stringPoint5);
            stringCache1.Append(stringPoint6);
            stringCache1.Append(stringPoint7);
            stringCache1.Append(stringPoint8);

            stringReference1.Append(formula1);
            stringReference1.Append(stringCache1);

            categoryAxisData1.Append(stringReference1);

            C.Values values1 = new C.Values();

            C.NumberReference numberReference1 = new C.NumberReference();
            C.Formula formula2 = new C.Formula();
            //formula2.Text = "\'Vendor Analytics\'!$B$32:$B$39";
            formula2.Text = shopVisitsDataRange;

            C.NumberingCache numberingCache1 = new C.NumberingCache();
            C.FormatCode formatCode1 = new C.FormatCode();
            formatCode1.Text = "General";
            C.PointCount pointCount2 = new C.PointCount() { Val = (UInt32Value)8U };

            C.NumericPoint numericPoint1 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue9 = new C.NumericValue();
            numericValue9.Text = "65";

            numericPoint1.Append(numericValue9);

            C.NumericPoint numericPoint2 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue10 = new C.NumericValue();
            numericValue10.Text = "23";

            numericPoint2.Append(numericValue10);

            C.NumericPoint numericPoint3 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue11 = new C.NumericValue();
            numericValue11.Text = "15";

            numericPoint3.Append(numericValue11);

            C.NumericPoint numericPoint4 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue12 = new C.NumericValue();
            numericValue12.Text = "17";

            numericPoint4.Append(numericValue12);

            C.NumericPoint numericPoint5 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue13 = new C.NumericValue();
            numericValue13.Text = "19";

            numericPoint5.Append(numericValue13);

            C.NumericPoint numericPoint6 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue14 = new C.NumericValue();
            numericValue14.Text = "18";

            numericPoint6.Append(numericValue14);

            C.NumericPoint numericPoint7 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue15 = new C.NumericValue();
            numericValue15.Text = "17";

            numericPoint7.Append(numericValue15);

            C.NumericPoint numericPoint8 = new C.NumericPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue16 = new C.NumericValue();
            numericValue16.Text = "17";

            numericPoint8.Append(numericValue16);

            numberingCache1.Append(formatCode1);
            numberingCache1.Append(pointCount2);
            numberingCache1.Append(numericPoint1);
            numberingCache1.Append(numericPoint2);
            numberingCache1.Append(numericPoint3);
            numberingCache1.Append(numericPoint4);
            numberingCache1.Append(numericPoint5);
            numberingCache1.Append(numericPoint6);
            numberingCache1.Append(numericPoint7);
            numberingCache1.Append(numericPoint8);

            numberReference1.Append(formula2);
            numberReference1.Append(numberingCache1);

            values1.Append(numberReference1);

            pieChartSeries1.Append(index1);
            pieChartSeries1.Append(order1);
            pieChartSeries1.Append(dataLabels1);
            pieChartSeries1.Append(categoryAxisData1);
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

            C.ManualLayout manualLayout2 = new C.ManualLayout();
            C.LeftMode leftMode2 = new C.LeftMode() { Val = C.LayoutModeValues.Edge };
            C.TopMode topMode2 = new C.TopMode() { Val = C.LayoutModeValues.Edge };
            C.Left left2 = new C.Left() { Val = 0.52736050586889882D };
            C.Top top2 = new C.Top() { Val = 6.4585653398003831E-2D };
            C.Width width2 = new C.Width() { Val = 0.4600607471235908D };
            C.Height height2 = new C.Height() { Val = 0.88479036979016357D };

            manualLayout2.Append(leftMode2);
            manualLayout2.Append(topMode2);
            manualLayout2.Append(left2);
            manualLayout2.Append(top2);
            manualLayout2.Append(width2);
            manualLayout2.Append(height2);

            layout2.Append(manualLayout2);
            C.Overlay overlay1 = new C.Overlay() { Val = true };

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
            legend1.Append(overlay1);
            legend1.Append(textProperties1);
            C.PlotVisibleOnly plotVisibleOnly1 = new C.PlotVisibleOnly() { Val = true };

            chart1.Append(plotArea1);
            chart1.Append(legend1);
            chart1.Append(plotVisibleOnly1);

            C.PrintSettings printSettings1 = new C.PrintSettings();
            C.HeaderFooter headerFooter1 = new C.HeaderFooter();
            C.PageMargins pageMargins2 = new C.PageMargins() { Left = 0.70000000000000018D, Right = 0.70000000000000018D, Top = 0.75000000000000022D, Bottom = 0.75000000000000022D, Header = 0.3000000000000001D, Footer = 0.3000000000000001D };
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

        #region GenerateSharedStringTablePart1Content - MODIFY
        // Generates content of sharedStringTablePart1.
        private void GenerateSharedStringTablePart1Content(SharedStringTablePart sharedStringTablePart1, List<SiteAnalyticsVendorSummary> analytics, string vendorName, DateTime startDate, DateTime endDate)
        {
            SharedStringTable sharedStringTable1 = new SharedStringTable() { Count = (UInt32Value)48U, UniqueCount = (UInt32Value)16U };

            SharedStringItem ssi;
            Text t;

            ssi = new SharedStringItem();
            t = new Text();
            t.Text = "Vendor name analytics summary generated";
            ssi.Append(t);
            sharedStringTable1.Append(ssi);

            ssi = new SharedStringItem();
            t = new Text();
            t.Text = "Impressions";
            ssi.Append(t);
            sharedStringTable1.Append(ssi);

            ssi = new SharedStringItem();
            t = new Text();
            t.Text = "Total portfolio";
            ssi.Append(t);
            sharedStringTable1.Append(ssi);

            int totalPortfolioImpressions = analytics.Sum(x => x.Impressions);
            ssi = new SharedStringItem();
            t = new Text();
            t.Text = totalPortfolioImpressions.ToString();
            ssi.Append(t);
            sharedStringTable1.Append(ssi);

            ssi = new SharedStringItem();
            t = new Text();
            t.Text = "between";
            ssi.Append(t);
            sharedStringTable1.Append(ssi);

            ssi = new SharedStringItem();
            t = new Text();
            t.Text = startDate.ToShortDateString();
            ssi.Append(t);
            sharedStringTable1.Append(ssi);

            ssi = new SharedStringItem();
            t = new Text();
            t.Text = "and";
            ssi.Append(t);
            sharedStringTable1.Append(ssi);

            ssi = new SharedStringItem();
            t = new Text();
            t.Text = endDate.ToShortDateString();
            ssi.Append(t);
            sharedStringTable1.Append(ssi);

            foreach (SiteAnalyticsVendorSummary savs in analytics)
            {
                ssi = new SharedStringItem();
                t = new Text();
                t.Text = savs.ServiceName + "GENERATED";
                ssi.Append(t);
                sharedStringTable1.Append(ssi);

                //ssi = new SharedStringItem();
                //t = new Text();
                //t.Text = savs.Impressions.ToString();
                //ssi.Append(t);
                //sharedStringTable1.Append(ssi);

                //ssi = new SharedStringItem();
                //t = new Text();
                //t.Text = "between";
                //ssi.Append(t);
                //sharedStringTable1.Append(ssi);

                //ssi = new SharedStringItem();
                //t = new Text();
                //t.Text = startDate.ToShortDateString();
                //ssi.Append(t);
                //sharedStringTable1.Append(ssi);

                //ssi = new SharedStringItem();
                //t = new Text();
                //t.Text = "and";
                //ssi.Append(t);
                //sharedStringTable1.Append(ssi);

                //ssi = new SharedStringItem();
                //t = new Text();
                //t.Text = endDate.ToShortDateString();
                //ssi.Append(t);
                //sharedStringTable1.Append(ssi);
               
            }

            //SharedStringItem sharedStringItem1 = new SharedStringItem();
            //Text text1 = new Text();
            //text1.Text = "Vendor name analytics summary";

            //sharedStringItem1.Append(text1);

            //SharedStringItem sharedStringItem2 = new SharedStringItem();
            //Text text2 = new Text();
            //text2.Text = "Impressions";

            //sharedStringItem2.Append(text2);

            //SharedStringItem sharedStringItem3 = new SharedStringItem();
            //Text text3 = new Text();
            //text3.Text = "Total portfolio";

            //sharedStringItem3.Append(text3);

            //SharedStringItem sharedStringItem4 = new SharedStringItem();
            //Text text4 = new Text();
            //text4.Text = "6753";

            //sharedStringItem4.Append(text4);

            //SharedStringItem sharedStringItem5 = new SharedStringItem();
            //Text text5 = new Text();
            //text5.Text = "between";

            //sharedStringItem5.Append(text5);

            //SharedStringItem sharedStringItem6 = new SharedStringItem();
            //Text text6 = new Text();
            //text6.Text = "01/08/2013";

            //sharedStringItem6.Append(text6);

            //SharedStringItem sharedStringItem7 = new SharedStringItem();
            //Text text7 = new Text();
            //text7.Text = "and";

            //sharedStringItem7.Append(text7);

            //SharedStringItem sharedStringItem8 = new SharedStringItem();
            //Text text8 = new Text();
            //text8.Text = "31/08/2013";

            //sharedStringItem8.Append(text8);

            //SharedStringItem sharedStringItem9 = new SharedStringItem();
            //Text text9 = new Text();
            //text9.Text = "Sales Cloud Contact Manager";

            //sharedStringItem9.Append(text9);

            //SharedStringItem sharedStringItem10 = new SharedStringItem();
            //Text text10 = new Text();
            //text10.Text = "Sales Cloud Group";

            //sharedStringItem10.Append(text10);

            //SharedStringItem sharedStringItem11 = new SharedStringItem();
            //Text text11 = new Text();
            //text11.Text = "Sales Cloud Professional";

            //sharedStringItem11.Append(text11);

            //SharedStringItem sharedStringItem12 = new SharedStringItem();
            //Text text12 = new Text();
            //text12.Text = "Sales Cloud Enterprise";

            //sharedStringItem12.Append(text12);

            //SharedStringItem sharedStringItem13 = new SharedStringItem();
            //Text text13 = new Text();
            //text13.Text = "Sales Cloud Unlimited";

            //sharedStringItem13.Append(text13);

            //SharedStringItem sharedStringItem14 = new SharedStringItem();
            //Text text14 = new Text();
            //text14.Text = "Service Cloud Professional";

            //sharedStringItem14.Append(text14);

            //SharedStringItem sharedStringItem15 = new SharedStringItem();
            //Text text15 = new Text();
            //text15.Text = "Service Cloud Enterprise";

            //sharedStringItem15.Append(text15);

            //SharedStringItem sharedStringItem16 = new SharedStringItem();
            //Text text16 = new Text();
            //text16.Text = "Service Cloud Unlimited";

            //sharedStringItem16.Append(text16);

            //sharedStringTable1.Append(sharedStringItem1);
            //sharedStringTable1.Append(sharedStringItem2);
            //sharedStringTable1.Append(sharedStringItem3);
            //sharedStringTable1.Append(sharedStringItem4);
            //sharedStringTable1.Append(sharedStringItem5);
            //sharedStringTable1.Append(sharedStringItem6);
            //sharedStringTable1.Append(sharedStringItem7);
            //sharedStringTable1.Append(sharedStringItem8);
            //sharedStringTable1.Append(sharedStringItem9);
            //sharedStringTable1.Append(sharedStringItem10);
            //sharedStringTable1.Append(sharedStringItem11);
            //sharedStringTable1.Append(sharedStringItem12);
            //sharedStringTable1.Append(sharedStringItem13);
            //sharedStringTable1.Append(sharedStringItem14);
            //sharedStringTable1.Append(sharedStringItem15);
            //sharedStringTable1.Append(sharedStringItem16);

            sharedStringTablePart1.SharedStringTable = sharedStringTable1;
        }
        #endregion

        #region SetPackageProperties
        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2013-08-31T13:02:32Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Glyn";
        }
        #endregion

    }
}
