using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using RestSharp;
using RestSharp.Authenticators;
using sc2DataReader.GameData;

namespace sc2DataReader
{
    class AiArenaBotList
    {
        public int count { get; set; }
        public List<AiArenaBot> results { get; set; }
    }
    class AiArenaBot
    {
        public string name { get; set; }
        public DateTime created { get; set; }
        public string game_display_id { get; set; }
    }

    class Program
    {
        private static Dictionary<string, string> BotNameMap = new Dictionary<string, string>();

        private const bool dropUnknownGamesExcel = true;

        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                args = new[] {"games"};
            }

            bool automaticallOopenFile = false;
            if (args.Contains("-y")) {
                automaticallOopenFile = true;
            }

            if (args.Contains("-api"))
            {
                string token = null;

                for (int i = 0; i < args.Length - 1; i++)
                {
                    if (args[i] == "-api")
                    {
                        token = args[i + 1];
                    }
                }
                // https://aiarena.net/api/bots/?active=true
                var client = new RestClient("https://aiarena.net/api/");
                //client.Authenticator = new HttpBasicAuthenticator("DrInfy", "YHj9Te0OfSK5RKpUvtbQ");
                client.AddDefaultHeader("Authorization", $"Token {token}");

                var request = new RestRequest("bots/?active=true", DataFormat.Json);
                var restResponse = client.Get<AiArenaBotList>(request);

                if (restResponse.IsSuccessful)
                {
                    foreach (AiArenaBot bot in restResponse.Data.results)
                    {
                        BotNameMap[bot.game_display_id] = bot.name;
                    }
                }
            }

            var path = args[0];
            if (!Directory.Exists(path))
            {
                Console.WriteLine($"Directory doesn't exist: '{path}'");
                return;
            }

            var stats = new Stats();

            foreach (var file in Directory.GetFiles(path, "*.log"))
            {
                var gameStats = ParseLogFile(file);
                stats.Add(gameStats);
            }

            Console.WriteLine();
            stats.Write();

            var now = DateTime.Now;
            var filename = $"stats-{now.Year}-{now.Month}-{now.Day}.xlsx";
            var fullPath = Path.Combine(path, filename);

            Console.WriteLine($"Writing stats to file {fullPath}");

            if (dropUnknownGamesExcel)
            {
                stats.DropUnknown();
            }

            ExcelWriter.Write(stats, fullPath);

            if (automaticallOopenFile)
            {
                OpenFile(fullPath);
            } else
            {
                Console.WriteLine("\nPress 'y' to open the file.\n\nYou can also use '-y' as command line parameter to automatically open it.");
                var key = Console.ReadKey();
                if (key.KeyChar == 'y')
                {
                    OpenFile(fullPath);
                }
            }
        }

        private static GameStats ParseLogFile(string file)
        {
            var gameStats = new GameStats();

            var fileName = Path.GetFileNameWithoutExtension(file);
            gameStats.GameName = fileName;

            var splits = fileName.Replace("random_learner", "randomlearner").Replace("_v2", "v2").Replace(" ", "_").Split('_');

            if (splits.Length < 3)
            {
                Console.WriteLine("Invalid file: " + file);
                return null;
            }

            gameStats.Map = splits[1];
            gameStats.Opponent = splits[0];

            if (BotNameMap.TryGetValue(gameStats.Opponent, out var botName))
            {
                gameStats.Opponent = botName;
            }

            // todo: change run_custom.py so that the timestamp is easier to parse
            string timestampStr;
            if (!splits[1].StartsWith("2"))
            {
                gameStats.Map = splits[1];
                timestampStr = $"{splits[2]} {splits[3]}:{splits[4]}:{splits[5]}";
            }
            else
            {
                gameStats.Map = "Unknown";
                timestampStr = splits[1] + " " + string.Join(":", splits.Skip(2));
            }

            var timestamp = DateTime.ParseExact(timestampStr, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

            try
            {
                gameStats.StartedOn = timestamp;

                string contents = File.ReadAllText(file);

                var lines = contents.Split('\r', '\n');
                var ownUnits = false;
                var enemyUnits = false;

                foreach (var fullLine in lines)
                {
                    var line = "";
                    try
                    {
                        line = fullLine.Split("|", 2, StringSplitOptions.None)[1].Trim();
                    }
                    catch (Exception e)
                    {
                        continue;
                    }
                    if (line.StartsWith("[EDGE]"))
                    {
                        var words = line.Split(' ', StringSplitOptions.RemoveEmptyEntries);

                        if(TryParseDuration(words[1], out TimeSpan? duration))
                        {
                            gameStats.Duration = duration.Value;
                        }

                        if (line.Contains("[Start] My race"))
                        {
                            gameStats.MyRace = words.Last();
                        }

                        if (line.Contains("[Start] Opponent race"))
                        {
                            gameStats.OpponentRace = words.Last();
                        }

                        if (line.Contains("[Start] OpponentId"))
                        {
                            gameStats.OpponentId = words.Last();
                        }

                        if (line.Contains("[Chat] Sharpened Edge"))
                        {
                            gameStats.BotVersion = words.Last();
                        }

                        if (line.Contains("[MapInfo] Map"))
                        {
                            var index = line.IndexOf("[MapInfo] Map") + "[MapInfo] Map".Length + 1;
                            gameStats.Map = line.Substring(index);
                        }

                        if (line.Contains("[Build]"))
                        {
                            gameStats.Build = words.Last().Trim();
                        }

                        if (line.Contains("Duration:"))
                        {
                            if (TryParseDuration(words.Last(), out TimeSpan? finalDuration))
                            {
                                gameStats.Duration = finalDuration.Value;
                            }
                        }

                        if (line.Contains("Minerals max"))
                        {
                            gameStats.MaxMinerals = int.Parse(words[4]);
                            gameStats.AverageMinerals = int.Parse(words[6]);
                        }

                        if (line.Contains("Vespene max"))
                        {
                            gameStats.MaxGas = int.Parse(words[4]);
                            gameStats.AverageGas = int.Parse(words[6]);
                        }

                        if (line.Contains("My lost units minerals and gas:"))
                        {
                            words = line.Split('(', ',', ')');
                            gameStats.MineralsLost = int.Parse(words[words.Length - 3], NumberFormatInfo.InvariantInfo);
                            gameStats.GasLost = int.Parse(words[words.Length - 2], NumberFormatInfo.InvariantInfo);
                        }

                        if (line.Contains("Enemy lost units minerals and gas:"))
                        {
                            words = line.Split('(', ',', ')');
                            gameStats.EnemyMineralsLost = int.Parse(words[words.Length - 3], NumberFormatInfo.InvariantInfo);
                            gameStats.EnemyGasLost = int.Parse(words[words.Length - 2], NumberFormatInfo.InvariantInfo);
                        }

                        if (line.Contains("Step time min:"))
                        {
                            gameStats.StepTimeMin = int.Parse(words.Last());
                        }

                        if (line.Contains("Step time avg:"))
                        {
                            gameStats.StepTimeAvg = int.Parse(words.Last());
                        }

                        if (line.Contains("Step time max:"))
                        {
                            gameStats.StepTimeMax = int.Parse(words.Last());
                        }

                        if (line.Contains("Own units:"))
                        {
                            ownUnits = true;
                            enemyUnits = false;
                        }

                        if (line.Contains("FinalScore:"))
                        {
                            gameStats.FinalScore = float.Parse(words.Last(), NumberStyles.Float, CultureInfo.InvariantCulture);
                        }
                        

                        if (line.Contains("Enemy units:"))
                        {
                            ownUnits = false;
                            enemyUnits = true;
                        }

                        if (line.Contains("Result"))
                        {
                            ownUnits = false;
                            enemyUnits = false;
                        }

                        if (ownUnits || enemyUnits)
                        {
                            if (line.Contains("DRONE") || line.Contains("PROBE") || line.Contains("SCV"))
                            {
                                var lineSplits = line.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                                var value = 0;

                                for (int i = 0; i < lineSplits.Length-1; i++)
                                {
                                    if (lineSplits[i] == "total:")
                                    {
                                        value = int.Parse(lineSplits[i + 1]);
                                    }
                                }
                                if (ownUnits)
                                {
                                    gameStats.Workers = value;
                                }
                                else
                                {
                                    gameStats.EnemyWorkers = value;
                                }
                            }

                        }

                        if (line.Contains("Result: "))
                        {
                            if (line.Contains("Victory"))
                            {
                                gameStats.Result = Result.Victory;
                            }
                            else if (line.Contains("Defeat"))
                            {
                                gameStats.Result = Result.Defeat;
                            }
                            else if (line.Contains("Tie"))
                            {
                                gameStats.Result = Result.Draw;
                            }
                        }
                        else if (line.Contains("Episode:"))
                        {
                            var strings = line.Split('|');

                            foreach (var split in strings)
                            {
                                var texts = split.Split(":");

                                if (texts.Length == 2 && texts[0].Trim() == "Loss")
                                {
                                    if (float.TryParse(texts[1], NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
                                    {
                                        gameStats.Loss = result;
                                    }
                                }

                                if (texts.Length == 2 && texts[0].Trim() == "Steps")
                                {
                                    if (float.TryParse(texts[1], NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
                                    {
                                        gameStats.Steps = result;
                                    }
                                }
                            }
                        }
                    }
                    else if (line.Contains("[Build]"))
                    {
                        var words = line.Split("[Build]");
                        gameStats.DummyBuild = words.Last().Trim();
                    }
                    

                    if (line.Contains("rush distance"))
                    {
                        var words = line.Split(':');
                        gameStats.RushDistance = words.Last().Trim();
                        gameStats.RushDistance = float.Parse(gameStats.RushDistance, CultureInfo.InvariantCulture).ToString("N2");
                    }
                    
                    if (line.Contains("[Surrender]"))
                    {
                        gameStats.Result = Result.Defeat;
                    }
                }

                if (contents.Contains("AI step threw an error"))
                {
                    gameStats.Result = Result.Crash;
                }
            }
            catch (IOException e)
            {
                // The file is probably being used by another process. Just print exception message
                // instead of whole stacktrace.
                Console.WriteLine($"Error on file: {file}");
                Console.WriteLine(e.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error on file: {file}");
                Console.WriteLine(e);
            }

            return gameStats;
        }

        private static void OpenFile(string fullPath)
        {
            Console.WriteLine("\nOpening...");

            FileInfo fi = new FileInfo(fullPath);
            if (fi.Exists)
            {
                Process.Start(new ProcessStartInfo(fullPath) { UseShellExecute = true });
            }
        }

        private static bool TryParseDuration(string durationStr, out TimeSpan? duration)
        {
            if (DateTime.TryParseExact(durationStr, "mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
            {
                duration = date.TimeOfDay;
                return true;
            }

            duration = null;
            return false;
        }
    }
}