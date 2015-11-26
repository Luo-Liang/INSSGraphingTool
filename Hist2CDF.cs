using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Collections.Concurrent;

namespace intelliSys.GraphUtils
{
    public class Hist2CDF
    {
        /// <summary>
        /// Takes a CSV file and sucks it into the specified worksheet of this workbook at the specified range
        /// </summary>
        /// <param name="importFileName">Specifies the full path to the .CSV file to import</param>
        /// <param name="destinationSheet">Excel.Worksheet object corresponding to the destination worksheet.</param>
        /// <param name="destinationRange">Excel.Range object specifying the destination cell(s)</param>
        /// <param name="columnDataTypes">Column data type specifier array. For the QueryTable.TextFileColumnDataTypes property.</param>
        /// <param name="autoFitColumns">Specifies whether to do an AutoFit on all imported columns.</param>
        public static void ImportCSV(string importFileName, Worksheet destinationSheet,
            Range destinationRange, int[] columnDataTypes, bool autoFitColumns)
        {
            destinationSheet.QueryTables.Add(
                "TEXT;" + Path.GetFullPath(importFileName),
            destinationRange, Type.Missing);
            destinationSheet.QueryTables[1].Name = Path.GetFileNameWithoutExtension(importFileName);
            destinationSheet.QueryTables[1].FieldNames = true;
            destinationSheet.QueryTables[1].RowNumbers = false;
            destinationSheet.QueryTables[1].FillAdjacentFormulas = false;
            destinationSheet.QueryTables[1].PreserveFormatting = true;
            destinationSheet.QueryTables[1].RefreshOnFileOpen = false;
            destinationSheet.QueryTables[1].RefreshStyle = XlCellInsertionMode.xlInsertDeleteCells;
            destinationSheet.QueryTables[1].SavePassword = false;
            destinationSheet.QueryTables[1].SaveData = true;
            destinationSheet.QueryTables[1].AdjustColumnWidth = true;
            destinationSheet.QueryTables[1].RefreshPeriod = 0;
            destinationSheet.QueryTables[1].TextFilePromptOnRefresh = false;
            destinationSheet.QueryTables[1].TextFilePlatform = 437;
            destinationSheet.QueryTables[1].TextFileStartRow = 1;
            destinationSheet.QueryTables[1].TextFileParseType = XlTextParsingType.xlDelimited;
            destinationSheet.QueryTables[1].TextFileTextQualifier = XlTextQualifier.xlTextQualifierDoubleQuote;
            destinationSheet.QueryTables[1].TextFileConsecutiveDelimiter = false;
            destinationSheet.QueryTables[1].TextFileTabDelimiter = false;
            destinationSheet.QueryTables[1].TextFileSemicolonDelimiter = false;
            destinationSheet.QueryTables[1].TextFileCommaDelimiter = true;
            destinationSheet.QueryTables[1].TextFileSpaceDelimiter = false;
            destinationSheet.QueryTables[1].TextFileColumnDataTypes = columnDataTypes;

            destinationSheet.QueryTables[1].Refresh(false);

            if (autoFitColumns == true)
                destinationSheet.QueryTables[1].Destination.EntireColumn.AutoFit();

            // cleanup
        }

        static void DrawLineChart(Worksheet sheet, string SeriesName, double maximum, double minimum)
        {
            DrawCombinedLineChart(new Worksheet[] { sheet }, sheet, SeriesName, maximum, minimum);
        }

        static void DrawCombinedLineChart(IEnumerable<Worksheet> source, Worksheet dest, string SeriesName, double maximum, double minimum)
        {
            Excel.ChartObjects xlCharts = (Excel.ChartObjects)dest.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 800, 400);
            Excel.Chart chartPage = myChart.Chart;
            myChart.Select();

            chartPage.ChartType = Excel.XlChartType.xlXYScatterLines;
            Microsoft.Office.Interop.Excel.Application xla = new Microsoft.Office.Interop.Excel.Application();
            Excel.SeriesCollection seriesCollection = chartPage.SeriesCollection();
            Excel.Axis xAxis = (Excel.Axis)chartPage.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
            xAxis.MaximumScale = maximum;
            //xAxis.LogBase = 10;
            Excel.Axis yAxis = (Excel.Axis)chartPage.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
            yAxis.MaximumScale = 1;

            foreach (var sheet in source)
            {
                Excel.Series series1 = seriesCollection.NewSeries();
                series1.Name = sheet.Name;
                series1.XValues = sheet.UsedRange.get_Range("A:A");
                series1.Values = sheet.UsedRange.get_Range("B:B");
                series1.Smooth = true;
                series1.MarkerStyle = XlMarkerStyle.xlMarkerStyleNone;
            }
        }
        const string Guid = "53321e6f-826b-49d7-a4e8-0aa5ec87f49e";
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.White;
            string TargetDirectory = args[0];
            double Stepping = double.Parse(args[1]);
            double Minimum = double.Parse(args[2]);
            double Maximum = 0;// double.Parse(args[3]);
            string GrouppingRule = "";
            if (args.Length > 3) GrouppingRule = args[3];
            Console.WriteLine("Getting rid of Excels...");
            foreach (var proc in Process.GetProcessesByName("excel")) proc.Kill();
            Console.WriteLine("Getting rid of stale files...");
            foreach (var file in Directory.GetFiles(TargetDirectory).Where(file => (file).Contains(Guid))) { File.Delete(file); }
            Application app = new Application();
            app.Visible = false;
            app.Workbooks.Add();
            var sheet1 = app.ActiveWorkbook.Worksheets[1];
            var allfiles = Directory.GetFiles(TargetDirectory, "*.txt");
            var groups = allfiles.GroupBy<string, string>(o => Path.GetFileNameWithoutExtension(o.Substring(1 + o.IndexOf(GrouppingRule))));
            var regex = new Regex("\\d+", RegexOptions.Compiled);
            foreach (var files in groups)
            {
                ConcurrentDictionary<string, double> EmpericalMaximum = new ConcurrentDictionary<string, double>();
                Worksheet groupSheet = app.ActiveWorkbook.Worksheets.Add();
                groupSheet.Name = $"{files.Key}";
                Parallel.ForEach(files, (file =>
                {
                    Console.WriteLine($"Parsing...{Path.GetFileNameWithoutExtension(file)}");
                    var streamReader = new StreamReader(file);
                    var streamWriter = new StreamWriter($"{file}-{Guid}.txt");
                    Dictionary<double, double> CDF = new Dictionary<double, double>();
                    int lineCount = 0;
                    EmpericalMaximum[file] = double.MinValue;
                    while (!streamReader.EndOfStream)
                    {
                        lineCount++;
                        var value = streamReader.ReadLine();
                        var dValue = double.Parse(regex.Match(value).Groups[0].Value);
                        var key = ((int)((dValue - Minimum) / Stepping)) * Stepping + Minimum;
                        if (CDF.ContainsKey(key) == false) CDF[key] = 0;
                        CDF[key]++;
                        EmpericalMaximum[file] = dValue > EmpericalMaximum[file] ? dValue : EmpericalMaximum[file];
                    }
                    var total = CDF.Values.Sum();
                    Console.WriteLine($"Parsing...{Path.GetFileNameWithoutExtension(file)}...done  {total}  items discovered.");
                    var lastKnownGood = 0.0;
                    for (double i = Minimum; i < EmpericalMaximum[file]; i += Stepping)
                    {
                        if (CDF.ContainsKey(i))
                        {
                            lastKnownGood += CDF[i] / total;
                        }
                        streamWriter.WriteLine($"{i},{lastKnownGood}");
                    }
                    streamReader.Dispose();
                    streamWriter.Flush();
                    streamWriter.Dispose();
                    streamWriter.Close();
                    Monitor.Enter(app);
                    Console.WriteLine($"Processing...{Path.GetFileNameWithoutExtension(file)}...Connecting Excel Imports");
                    Worksheet sheet = app.ActiveWorkbook.Worksheets.Add();
                    sheet.Name = Path.GetFileNameWithoutExtension(file);
                    ImportCSV($"{file}-{Guid}.txt", sheet, (Range)(sheet).get_Range("$A$1"), new int[] { 1, 1 }, true);
                    Console.WriteLine($"Processing...{Path.GetFileNameWithoutExtension(file)}...Excel Finished Importing");
                    Monitor.Exit(app);
                    try
                    {
                        File.Delete($"{file}-{Guid}.txt");
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(ex.Message);
                        Console.WriteLine("Will attempt removal next time.");
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                }));
                Console.WriteLine($"Finalizing for group {files.Key}");
                Maximum = Math.Ceiling(EmpericalMaximum.Values.Max() / Stepping) * Stepping;
                DrawCombinedLineChart(app.ActiveWorkbook.Worksheets.Cast<Worksheet>().Where(o => o.Name.Contains(files.Key) && o.Name != files.Key), groupSheet, $"CDF-Group-{files.Key}", Maximum, Minimum);
            }
            Console.WriteLine("Taking Care of a Few Things...");
            sheet1.Delete();
            app.Visible = true;
        }
    }
}
