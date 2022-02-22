using ChartJSCore.Helpers;
using ChartJSCore.Models;
using OfficeOpenXml;
using System.Globalization;

namespace BinanceCardCharts
{
    public static class CardUsageBarChart
    {
        public static double TotalAmountSpent { get; set; }
        public static double EstimatedCashback { get; set; }

        private const string BINANCE_FILE = "C:\\temp\\binance\\binance_card_chart_data.xlsx";

        public static Chart GenerateDailyChart()
        {
            var chart = new Chart
            {
                Type = Enums.ChartType.Bar,
                Data = new Data
                {
                    Labels = new List<string>(),
                    YLabels = new List<string>
                    {
                        "EUR"
                    }
                },
                Options = new Options
                {
                    Scales = new Dictionary<string, Scale>
                    {
                        {
                            "x", new BarScale
                            {
                                GridLines = new GridLine()
                                {
                                    OffsetGridLines = true
                                }
                            }
                        },
                        {
                            "y", new CartesianScale
                            {
                                Ticks = new CartesianLinearTick
                                {
                                    BeginAtZero = true
                                }
                            }
                        }
                    },
                    Layout = new Layout
                    {
                        Padding = new Padding
                        {
                            PaddingObject = new PaddingObject
                            {
                                Left = 10,
                                Right = 12
                            }
                        }
                    }
                }
            };

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Read excel from Binance
            using var binanceExcel = new ExcelPackage(new FileInfo(fileName: BINANCE_FILE));
            var binanceSheet = binanceExcel.Workbook.Worksheets["sheet1"];
            int binanceStart = binanceSheet.Dimension.Start.Row + 1;
            int binanceEnd = binanceSheet.Dimension.End.Row;

            var transactionsPerDay = new Dictionary<int, double?>();

            for (int row = binanceStart; row <= binanceEnd; row++)
            {
                // Set date
                string[] dateParts = binanceSheet.Cells[row, 1].Text.Split(' ');

                var transactionDate = new DateTime(
                    year: Convert.ToInt32(dateParts[5]),
                    month: GetMonth(dateParts[1]),
                    day: Convert.ToInt32(dateParts[2])
                );

                // Amount spent in EUR
                if (transactionsPerDay.ContainsKey(transactionDate.DayOfYear))
                {
                    transactionsPerDay[transactionDate.DayOfYear] += Convert.ToDouble($"{binanceSheet.Cells[row, 3].Text}", CultureInfo.InvariantCulture);
                }
                else
                {
                    transactionsPerDay.Add(transactionDate.DayOfYear, Convert.ToDouble($"{binanceSheet.Cells[row, 3].Text}", CultureInfo.InvariantCulture));
                }

                var transDateString = transactionDate.ToString("yyyy-MM-dd");

                if (!chart.Data.Labels.Contains(transDateString))
                {
                    chart.Data.Labels.Add(transDateString);
                }
            }

            chart.Data.Labels = chart.Data.Labels.Reverse().ToList();

            chart.Data.Datasets = new List<Dataset>
            {
                new BarDataset
                {
                    Label = "Amount spent per day (EUR)",
                    Data = transactionsPerDay.Values.Reverse().ToList(),
                    BackgroundColor = new List<ChartColor>
                    {
                        ChartColor.FromRgba(54, 162, 235, 0.2),
                    },
                    BorderColor = new List<ChartColor>
                    {
                        ChartColor.FromRgb(54, 162, 235),
                    },
                    BorderWidth = new List<int> { 1 },
                    BarPercentage = 0.5,
                    BarThickness = 10,
                    MaxBarThickness = 15,
                    MinBarLength = 1
                }
            };

            TotalAmountSpent = transactionsPerDay.Values.Sum() ?? 0;
            EstimatedCashback = Math.Round(TotalAmountSpent * 0.02, 1);

            return chart;
        }

        private static int GetMonth(string month) => month switch
        {
            "Jan" => 1,
            "Feb" => 2,
            "Mar" => 3,
            "Apr" => 4,
            "May" => 5,
            "Jun" => 6,
            "Jul" => 7,
            "Aug" => 8,
            "Sep" => 9,
            "Oct" => 10,
            "Nov" => 11,
            "Dec" => 12,
            _ => 0,
        };
    }
}