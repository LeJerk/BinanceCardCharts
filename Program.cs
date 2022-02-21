using ChartJSCore.Helpers;
using ChartJSCore.Models;
using IronPdf;
using Razor.Templating.Core;

namespace BinanceCardToKoinly
{
    internal static class Program
    {
        private const string BINANCE_FILE = "C:\\temp\\binance\\binance_card_chart.pdf";

        internal static async Task Main(string[] args)
        {
            var html = await RazorTemplateEngine.RenderAsync("C:\\Dev\\BinanceCardCharts\\html\\BinanceChart.cshtml");

            new ChromePdfRenderer()
                .RenderHTMLFileAsPdf(html)
                .SaveAs(BINANCE_FILE);
        }

        internal static Chart GenerateBarChart()
        {
            var chart = new Chart { Type = Enums.ChartType.Bar };
            var data = new Data
            {
                Labels = new List<string>
                {
                    "Red",
                    "Blue",
                    "Yellow",
                    "Green",
                    "Purple",
                    "Orange"
                }
            };

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
    }
}