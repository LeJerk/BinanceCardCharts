using ChartJSCore.Helpers;
using ChartJSCore.Models;
using OfficeOpenXml;
using System.Globalization;

namespace BinanceCardCharts
{
    public static class CardUsageBarChart
    {
        private const string BINANCE_FILE = "C:\\temp\\binance\\binance_card_chart_data.xlsx";

        public static Chart GenerateDailyChart()
        {
            var chart = new Chart { Type = Enums.ChartType.Bar };
            var data = new Data
            {
                Labels = new List<string>()
            };

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Read excel from Binance
            using var binanceExcel = new ExcelPackage(new FileInfo(fileName: BINANCE_FILE));
            var binanceSheet = binanceExcel.Workbook.Worksheets["sheet1"];
            int binanceStart = binanceSheet.Dimension.Start.Row + 1;
            int binanceEnd = binanceSheet.Dimension.End.Row;

            var transactionsPerDay = new Dictionary<int, double>();


            for (int row = binanceStart; row <= binanceEnd; row++)
            {
                // Set date
                string[] dateParts = binanceSheet.Cells[row, 1].Text.Split(' ');

                var transactionDate = new DateTime(
                    year: Convert.ToInt32(dateParts[5]),
                    month: GetMonth(dateParts[1]),
                    day: Convert.ToInt32(dateParts[2])
                );

                // Set amount and currency
                string[] assetUsed = binanceSheet.Cells[row, 6].Text.Split(' ');
                double amountSpent = Convert.ToDouble($"{assetUsed[1].Replace(";", "")}", CultureInfo.InvariantCulture);

                if (transactionsPerDay.ContainsKey(transactionDate.DayOfYear))
                {
                    transactionsPerDay[transactionDate.DayOfYear] += amountSpent;
                }
                else
                {
                    transactionsPerDay.Add(transactionDate.DayOfYear, amountSpent);
                }

                data.Labels.Add(transactionDate.ToString("yyyy-MM-dd"));
            }

            var dataset = new BarDataset
            {
                Label = "# of Votes",
                Data = new List<double?> { 12, 19, 3, null, 2, 3 },
                BackgroundColor = new List<ChartColor>
                {
                    ChartColor.FromRgba(255, 99, 132, 0.2),
                    ChartColor.FromRgba(54, 162, 235, 0.2),
                    ChartColor.FromRgba(255, 206, 86, 0.2),
                    ChartColor.FromRgba(75, 192, 192, 0.2),
                    ChartColor.FromRgba(153, 102, 255, 0.2),
                    ChartColor.FromRgba(255, 159, 64, 0.2)
                },
                BorderColor = new List<ChartColor>
                {
                    ChartColor.FromRgb(255, 99, 132),
                    ChartColor.FromRgb(54, 162, 235),
                    ChartColor.FromRgb(255, 206, 86),
                    ChartColor.FromRgb(75, 192, 192),
                    ChartColor.FromRgb(153, 102, 255),
                    ChartColor.FromRgb(255, 159, 64)
                },
                BorderWidth = new List<int> { 1 },
                BarPercentage = 0.5,
                BarThickness = 6,
                MaxBarThickness = 8,
                MinBarLength = 2
            };

            data.Datasets = new List<Dataset> { dataset };
            chart.Data = data;
            chart.Options = new Options
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
            };

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