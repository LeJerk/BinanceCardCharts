using ChartJSCore.Helpers;
using ChartJSCore.Models;
using OfficeOpenXml;
using System.Globalization;
using System.Text.RegularExpressions;

namespace BinanceCardCharts
{
    public static class CardUsageBarChart
    {
        private const string BINANCE_FILE = "C:\\temp\\binance\\binance_card_chart_data.xlsx";
        private const string EUR = "EUR";
        private static readonly Regex _regex = new("[^0-9.]");

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

                // Calculate amount spent in EUR
                string[] assetUsed = binanceSheet.Cells[row, 6].Text.Split(' ');
                double amountSpent = ConvertAmountToEur(assetUsed, binanceSheet.Cells[row, 7].Text.Split('='));

                if (transactionsPerDay.ContainsKey(transactionDate.DayOfYear))
                {
                    transactionsPerDay[transactionDate.DayOfYear] += amountSpent;
                }
                else
                {
                    transactionsPerDay.Add(transactionDate.DayOfYear, amountSpent);
                }

                var transDateString = transactionDate.ToString("yyyy-MM-dd");

                if (!data.Labels.Contains(transDateString))
                {
                    data.Labels.Add(transDateString);
                }
            }

            var dataset = new BarDataset
            {
                Label = "Amount spent per day",
                Data = transactionsPerDay.Values.ToList(),
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
                BarThickness = 10,
                MaxBarThickness = 15,
                MinBarLength = 1
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

        private static double ConvertAmountToEur(string[] assetUsed, string[] exchangeRates)
        {
            // Using only EUR to finance transaction
            if (exchangeRates.Length == 1)
            {
                return Convert.ToDouble($"{assetUsed[1].Replace(";", "")}", CultureInfo.InvariantCulture);
            }

            // Using only one asset to finance transaction
            if (assetUsed.Length == 2)
            {
                double beforeConvertion = Convert.ToDouble($"{assetUsed[1].Replace(";", "")}", CultureInfo.InvariantCulture);
                double convertionRate = Convert.ToDouble(_regex.Replace(exchangeRates[1], ""), CultureInfo.InvariantCulture);
                return Math.Round(beforeConvertion / convertionRate, 2);
            }

            // Using assets two assets finance transaction
            if (assetUsed.Length == 4)
            {
                double totalEurAmount = 0;

                if (assetUsed[0].Equals(EUR))
                {
                    totalEurAmount = Convert.ToDouble($"{assetUsed[1].Replace(";", "")}", CultureInfo.InvariantCulture);
                }
/*                else
                {
                    double beforeConvertion = Convert.ToDouble($"{assetUsed[1].Replace(";", "")}", CultureInfo.InvariantCulture);
                    double convertionRate = Convert.ToDouble(_regex.Replace(exchangeRates[2], ""), CultureInfo.InvariantCulture);
                    totalEurAmount += Math.Round(totalEurAmount + (beforeConvertion / convertionRate), 2);
                }*/

                double beforeConvertion = Convert.ToDouble($"{assetUsed[3].Replace(";", "")}", CultureInfo.InvariantCulture);
                double convertionRate = Convert.ToDouble(_regex.Replace(exchangeRates[1], ""), CultureInfo.InvariantCulture);
                return Math.Round(totalEurAmount + (beforeConvertion / convertionRate), 2);
            }

            throw new NotImplementedException();
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