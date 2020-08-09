using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Style;
using sc2DataReader.GameData;

namespace sc2DataReader
{
    class Column
    {
        protected readonly string title;

        public Column(string title)
        {
            this.title = title;
        }
    }

    class Column<T1> : Column
    {
        private readonly Action<T1, ExcelRange> action;

        public Column(string title, Action<T1, ExcelRange> action):base(title)
        {
            this.action = action;
        }

        public void WriteTitle(ExcelRange cell)
        {
            cell.Value = this.title;
        }

        public void WriteCell(ExcelRange cell, T1 data)
        {
            this.action(data, cell);
        }
    }

    class ExcelWriter
    {
        private static Column<GameStats>[] sheet1Columns;
        private static Column<WinLose>[] sheet2Columns;
        private static Column<WinLose>[] sheet3Columns;
        private static Column<WinLose>[] sheet4Columns;
        private static Column<WinLose>[] raceSummaryColumns;
        private static Column<WinLose>[] buildSummaryColumns;

        private static readonly Color winColor = Color.YellowGreen;
        private static readonly Color minorWinColor = Color.LightGreen;
        private static readonly Color drawColor = Color.Yellow;

        private static readonly Color lossColor = Color.Salmon;
        private static readonly Color crashColor = Color.CornflowerBlue;


        static ExcelWriter()
        {
            sheet1Columns = new[]
            {
                new Column<GameStats>("Game started", (stats, cell) => {
                    cell.Value = stats.StartedOn;
                    cell.Style.Numberformat.Format = "yyyy-mm-dd HH:mm:ss";
                }),
                new Column<GameStats>("Opponent", (stats, cell) => cell.Value = stats.Opponent),
                new Column<GameStats>("E Race", (stats, cell) => cell.Value = stats.OpponentRace),
                new Column<GameStats>("Map", (stats, cell) => cell.Value = stats.Map),
                new Column<GameStats>("Build", (stats, cell) => cell.Value = stats.Build),
                new Column<GameStats>("E Build", (stats, cell) => cell.Value = stats.DummyBuild),
                new Column<GameStats>("Result", (stats, cell) => {
                    var result = stats.ResultToString();
                    cell.Value = result;
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;

                    switch (result) {
                        case "Victory":
                            cell.Style.Fill.BackgroundColor.SetColor(winColor);
                            break;
                        case "Draw":
                            cell.Style.Fill.BackgroundColor.SetColor(drawColor);
                            break;
                        case "Defeat":
                            cell.Style.Fill.BackgroundColor.SetColor(lossColor);
                            break;
                        case "Crash":
                            cell.Style.Fill.BackgroundColor.SetColor(crashColor);
                            break;
                    }   
                }),
                new Column<GameStats>("Duration", (stats, cell) =>
                {
                    cell.Value = stats.Duration;
                    cell.Style.Numberformat.Format = "mm:ss";
                }),
                new Column<GameStats>("Avg (ms)", (stats, cell) =>
                {
                    cell.Value = stats.StepTimeAvg;
                }),
                new Column<GameStats>("Score", (stats, cell) =>
                {
                    cell.Value = stats.Score;
                    cell.Style.Numberformat.Format = "# ### ### ##0.##";
                }),
                new Column<GameStats>("RushDistance", (stats, cell) => cell.Value = stats.RushDistance),
                new Column<GameStats>("MineralsLost", (stats, cell) => cell.Value = stats.MineralsLost),
                new Column<GameStats>("GasLost", (stats, cell) => cell.Value = stats.GasLost),

                new Column<GameStats>("E MineralsLost", (stats, cell) => cell.Value = stats.EnemyMineralsLost),
                new Column<GameStats>("E GasLost", (stats, cell) => cell.Value = stats.EnemyGasLost),

                new Column<GameStats>("MaxMinerals", (stats, cell) => cell.Value = stats.MaxMinerals),
                new Column<GameStats>("MaxGas", (stats, cell) => cell.Value = stats.MaxGas),

                new Column<GameStats>("Average Minerals", (stats, cell) => {
                    cell.Value = stats.AverageMinerals;

                    if (stats.AverageMinerals > 1000)
                    {
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(drawColor);
                    }
                }),
                new Column<GameStats>("Average Gas", (stats, cell) => {
                    cell.Value = stats.AverageGas;

                    if (stats.AverageGas > 1000)
                    {
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(drawColor);
                    }
                }),
                new Column<GameStats>("Workers", (stats, cell) => cell.Value = stats.Workers),
                new Column<GameStats>("E Workers", (stats, cell) => cell.Value = stats.EnemyWorkers),
                new Column<GameStats>("Loss", (stats, cell) =>
                {
                    cell.Value = stats.Loss;
                    cell.Style.Numberformat.Format = "# ### ### ##0.##";
                }),
                new Column<GameStats>("Game name", (stats, cell) => cell.Value = stats.GameName),
                new Column<GameStats>("Bot Version", (stats, cell) => cell.Value = stats.BotVersion),
            };

            sheet2Columns = new[]
            {
                new Column<WinLose>("Opponent", (stats, cell) => cell.Value = stats.Stats.FirstOrDefault()?.Opponent),
                new Column<WinLose>("Win percentage", (stats, cell) =>
                {
                    cell.Value = stats.WinPercentage;
                    cell.Style.Numberformat.Format = "#0.00 %";
                    ColorizeCell(cell, stats.WinPercentage);
                }),
                new Column<WinLose>("Games", (stats, cell) => cell.Value = stats.TotalGames),
                new Column<WinLose>("Crashes", (stats, cell) => cell.Value = stats.Crashes),
                new Column<WinLose>("Losses", (stats, cell) => cell.Value = stats.Losses),
                new Column<WinLose>("Draws", (stats, cell) => cell.Value = stats.Draws),
                new Column<WinLose>("Wins", (stats, cell) => cell.Value = stats.Wins),
                new Column<WinLose>("Unknown", (stats, cell) => cell.Value = stats.Unknown),
            };

            sheet3Columns = new[]
            {
                new Column<WinLose>("Map", (stats, cell) =>
                {
                    cell.Value = stats.Stats.FirstOrDefault()?.Map;
                }),
                new Column<WinLose>("Win percentage", (stats, cell) =>
                {
                    cell.Value = stats.WinPercentage;
                    cell.Style.Numberformat.Format = "#0.00 %";

                    ColorizeCell(cell, stats.WinPercentage);
                }),
                new Column<WinLose>("Games", (stats, cell) => cell.Value = stats.TotalGames),
                new Column<WinLose>("Crashes", (stats, cell) => cell.Value = stats.Crashes),
                new Column<WinLose>("Losses", (stats, cell) => cell.Value = stats.Losses),
                new Column<WinLose>("Draws", (stats, cell) => cell.Value = stats.Draws),
                new Column<WinLose>("Wins", (stats, cell) => cell.Value = stats.Wins),
                new Column<WinLose>("Unknown", (stats, cell) => cell.Value = stats.Unknown),
            };

            sheet4Columns = new[]
            {
                new Column<WinLose>("Opponent", (stats, cell) => cell.Value = stats.Stats.FirstOrDefault()?.Opponent),
                new Column<WinLose>("Map", (stats, cell) => cell.Value = stats.Stats.FirstOrDefault()?.Map),
                new Column<WinLose>("Win percentage", (stats, cell) =>
                {
                    cell.Value = stats.WinPercentage;
                    cell.Style.Numberformat.Format = "#0.00 %";
                    ColorizeCell(cell, stats.WinPercentage);
                }),
                new Column<WinLose>("Games", (stats, cell) => cell.Value = stats.TotalGames),
                new Column<WinLose>("Crashes", (stats, cell) => cell.Value = stats.Crashes),
                new Column<WinLose>("Losses", (stats, cell) => cell.Value = stats.Losses),
                new Column<WinLose>("Draws", (stats, cell) => cell.Value = stats.Draws),
                new Column<WinLose>("Wins", (stats, cell) => cell.Value = stats.Wins),
                new Column<WinLose>("Unknown", (stats, cell) => cell.Value = stats.Unknown),
            };

            raceSummaryColumns = new[]
            {
                new Column<WinLose>("Matchup", (stats, cell) =>
                {
                    cell.Value = stats.Stats.FirstOrDefault()?.RaceMatchup;
                }),
                new Column<WinLose>("Win percentage", (stats, cell) =>
                {
                    cell.Value = stats.WinPercentage;
                    cell.Style.Numberformat.Format = "#0.00 %";

                    ColorizeCell(cell, stats.WinPercentage);
                }),
                new Column<WinLose>("Games", (stats, cell) => cell.Value = stats.TotalGames),
                new Column<WinLose>("Crashes", (stats, cell) => cell.Value = stats.Crashes),
                new Column<WinLose>("Losses", (stats, cell) => cell.Value = stats.Losses),
                new Column<WinLose>("Draws", (stats, cell) => cell.Value = stats.Draws),
                new Column<WinLose>("Wins", (stats, cell) => cell.Value = stats.Wins),
                new Column<WinLose>("Unknown", (stats, cell) => cell.Value = stats.Unknown),
            };

            buildSummaryColumns = new[]
            {
                new Column<WinLose>("Build", (stats, cell) =>
                {
                    cell.Value = stats.Stats.FirstOrDefault()?.Build;
                }),
                new Column<WinLose>("Win percentage", (stats, cell) =>
                {
                    cell.Value = stats.WinPercentage;
                    cell.Style.Numberformat.Format = "#0.00 %";

                    ColorizeCell(cell, stats.WinPercentage);
                }),
                new Column<WinLose>("Games", (stats, cell) => cell.Value = stats.TotalGames),
                new Column<WinLose>("Crashes", (stats, cell) => cell.Value = stats.Crashes),
                new Column<WinLose>("Losses", (stats, cell) => cell.Value = stats.Losses),
                new Column<WinLose>("Draws", (stats, cell) => cell.Value = stats.Draws),
                new Column<WinLose>("Wins", (stats, cell) => cell.Value = stats.Wins),
                new Column<WinLose>("Unknown", (stats, cell) => cell.Value = stats.Unknown),
            };
        }

        public static void Write(Stats stats, string fullPath)
        {
            ExcelPackage package = new ExcelPackage();
            package.Compatibility.IsWorksheets1Based = true;
            // Add a new worksheet to the empty workbook
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Games");

            WriteGameSheet(stats, worksheet);

            ExcelWorksheet summaryWorksheet = package.Workbook.Worksheets.Add("Opponent summary");
            WriteOpponentSheet(stats, summaryWorksheet);

            ExcelWorksheet mapSummaryWorksheet = package.Workbook.Worksheets.Add("Map summary");
            WriteMapSheet(stats, mapSummaryWorksheet);

            ExcelWorksheet opponentMapSummaryWorksheet = package.Workbook.Worksheets.Add("Opponent/map summary");
            WriteOpponentMapSheet(stats, opponentMapSummaryWorksheet);

            ExcelWorksheet raceSummaryWorksheet = package.Workbook.Worksheets.Add("Race summary");
            WriteRaceSheet(stats, raceSummaryWorksheet);

            ExcelWorksheet buildSummaryWorksheet = package.Workbook.Worksheets.Add("Build summary");
            WriteBuildSheet(stats, buildSummaryWorksheet);

            ExcelWorksheet buildOpponentWorksheet = package.Workbook.Worksheets.Add("Build + Opponent");
            WriteOpponentBuildSheet(stats, buildOpponentWorksheet);

            Save(fullPath, package);
        }

        private static void Save(string fullPath, ExcelPackage package)
        {
            try
            {
                package.SaveAs(new FileInfo(fullPath));
            }
            catch (Exception e)
            {
                Console.WriteLine(e);

                Console.WriteLine("\nPress 'y' to try saving again");
                var key = Console.ReadKey();
                if (key.KeyChar == 'y')
                {
                    Save(fullPath, package);
                }
            }
            
        }

        private static void WriteOpponentSheet(Stats stats, ExcelWorksheet worksheet)
        {
            var values = stats.VsDict.Values.OrderByDescending(opp => opp.WinPercentage);
            var rowIndex = WriteSheet(values, worksheet, sheet2Columns);
            worksheet.Cells[rowIndex, 1].Value = "Total";
            worksheet.Cells[rowIndex, 2].Value = (float)(stats.Wins.Count()) / stats.AllGames.Count();
            worksheet.Cells[rowIndex, 2].Style.Numberformat.Format = "#0.00 %";
            worksheet.Cells[rowIndex, 3].Formula = $"=SUM(C2:C{rowIndex - 1})";
            worksheet.Cells[rowIndex, 4].Formula = $"=SUM(D2:D{rowIndex - 1})";
            worksheet.Cells[rowIndex, 5].Formula = $"=SUM(E2:E{rowIndex - 1})";
            worksheet.Cells[rowIndex, 6].Formula = $"=SUM(F2:F{rowIndex - 1})";
            worksheet.Cells[rowIndex, 7].Formula = $"=SUM(G2:G{rowIndex - 1})";
            worksheet.Cells[rowIndex, 8].Formula = $"=SUM(H2:H{rowIndex - 1})";
        }

        private static void WriteOpponentMapSheet(Stats stats, ExcelWorksheet worksheet)
        {
            WriteSheet(stats.MapVsDict.Values, worksheet, sheet4Columns);
        }

        private static void WriteMapSheet(Stats stats, ExcelWorksheet worksheet)
        {
            var values = stats.MapDict.Values.OrderByDescending(map => map.WinPercentage);
            var rowIndex = WriteSheet(values, worksheet, sheet3Columns);
        }

        private static void WriteRaceSheet(Stats stats, ExcelWorksheet worksheet)
        {
            var values = stats.ByMatchupDict.Values.OrderByDescending(race => race.WinPercentage);
            WriteSheet(values, worksheet, raceSummaryColumns);
        }

        private static void WriteBuildSheet(Stats stats, ExcelWorksheet worksheet)
        {
            var values = stats.ByBuildDict.Values.OrderByDescending(build => build.WinPercentage);
            WriteSheet(values, worksheet, buildSummaryColumns);
        }

        private static void WriteGameSheet(Stats stats, ExcelWorksheet worksheet)
        {
            var values = stats.AllGames.OrderBy(game => game.StartedOn);
            WriteSheet(values, worksheet, sheet1Columns);
        }

        private static void WriteOpponentBuildSheet(Stats stats, ExcelWorksheet worksheet)
        {
            var buildOpponentColumns = new List<Column<WinLose>>();
            buildOpponentColumns.Add(
                new Column<WinLose>("Opponent", (_stats, cell) => cell.Value = _stats.Stats.FirstOrDefault()?.Opponent)
            );

            foreach (var key in stats.ByBuildDict.Keys)
            {
                buildOpponentColumns.Add(new Column<WinLose>($"{key} %", (winLose, cell) =>
                {
                    var games = winLose.Stats.Where(x => x.Build == key).ToArray();
                    var wins = games.Count(x => x.Result == Result.Victory);
                    if (games.Length == 0)
                    {

                    }
                    else
                    {
                        var winRate = (float)wins / games.Length;
                        cell.Value = winRate;
                        cell.Style.Numberformat.Format = "#0.00 %";

                        ColorizeCell(cell, winRate);
                    }
                }));

                buildOpponentColumns.Add(new Column<WinLose>($"Games", (winLose, cell) =>
                {
                    var games = winLose.Stats.Where(x => x.Build == key).ToArray();
                    var wins = games.Count(x => x.Result == Result.Victory);
                    if (games.Length == 0)
                    {

                    }
                    else
                    {
                        var winRate = (float)wins / games.Length;
                        cell.Value = $"{wins} - {games.Length - wins}";
                        //cell.Style.Numberformat.Format = "#0.00 %";

                        ColorizeCell(cell, winRate);
                    }
                }));
            }

            var values = stats.VsDict.Values.OrderByDescending(build => build.WinPercentage);
            WriteSheet(values, worksheet, buildOpponentColumns.ToArray());
        }

        private static void ColorizeCell(ExcelRange cell, float winPercentage)
        {
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            if (winPercentage >= 0.9)
            {
                cell.Style.Fill.BackgroundColor.SetColor(winColor);
            }
            else if (winPercentage >= 0.75)
            {
                cell.Style.Fill.BackgroundColor.SetColor(minorWinColor);
            }
            else if (winPercentage >= 0.5)
            {
                cell.Style.Fill.BackgroundColor.SetColor(drawColor);
            }
            else if (winPercentage >= 0)
            {
                cell.Style.Fill.BackgroundColor.SetColor(lossColor);
            }
        }

        private static int WriteSheet<T1>(IEnumerable<T1> data, ExcelWorksheet worksheet, Column<T1>[] columns)
        {
            var rowIndex = 1;
            for (int i = 0; i < columns.Length; i++)
            {
                var cell = worksheet.Cells[rowIndex, i + 1];
                columns[i].WriteTitle(cell);
            }

            rowIndex++;

            foreach (var row in data)
            {
                for (int i = 0; i < columns.Length; i++)
                {
                    var cell = worksheet.Cells[rowIndex, i + 1];
                    columns[i].WriteCell(cell, row);
                }

                rowIndex++;
            }

            for (int i = 0; i < columns.Length; i++)
            {
                worksheet.Column(i + 1).AutoFit(10, 100);
            }

            var range = worksheet.Cells[1, 1, rowIndex, columns.Length];
            range.AutoFilter = true;
            return rowIndex;
        }
    }
}
