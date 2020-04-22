using System;
using System.Collections.Generic;
using System.Linq;

namespace sc2DataReader.GameData
{
    class Stats
    {
        private WinLose all = new WinLose();
        public Dictionary<string, WinLose> VsDict = new Dictionary<string, WinLose>();
        public Dictionary<string, WinLose> MapDict = new Dictionary<string, WinLose>();
        public Dictionary<string, WinLose> MapVsDict = new Dictionary<string, WinLose>();
        public Dictionary<string, WinLose> ByMatchupDict = new Dictionary<string, WinLose>();
        public Dictionary<string, WinLose> ByBuildDict = new Dictionary<string, WinLose>();

        public IEnumerable<GameStats> Wins => this.all.Stats.Where(x => x.Result == Result.Victory);
        public IEnumerable<GameStats> Draws => this.all.Stats.Where(x => x.Result == Result.Draw);
        public IEnumerable<GameStats> Losses => this.all.Stats.Where(x => x.Result == Result.Defeat);
        public IEnumerable<GameStats> Unknown => this.all.Stats.Where(x => x.Result == Result.Unknown);
        public IEnumerable<GameStats> Crashes => this.all.Stats.Where(x => x.Result == Result.Crash);

        public IEnumerable<GameStats> AllGames => this.all.Stats;

        public void Add(GameStats result)
        {
            var opponent = result.Opponent;
            var map = result.Map;

            this.all.Add(result);

            if (!this.VsDict.TryGetValue(opponent, out var vsStat))
            {
                vsStat = new WinLose();
                this.VsDict[opponent] = vsStat;
            }

            vsStat.Add(result);

            if (!this.MapDict.TryGetValue(map, out var mapStat))
            {
                mapStat = new WinLose();
                this.MapDict[map] = mapStat;
            }

            mapStat.Add(result);

            if (!this.MapVsDict.TryGetValue(opponent + "_" + map, out var mapVsStat))
            {
                mapVsStat = new WinLose();
                this.MapVsDict[opponent + "_" + map] = mapVsStat;
            }

            mapVsStat.Add(result);
            
            if(!this.ByMatchupDict.TryGetValue(result.RaceMatchup, out var ByMatchupStat))
            {
                ByMatchupStat = new WinLose();
                this.ByMatchupDict[result.RaceMatchup] = ByMatchupStat;
            }

            ByMatchupStat.Add(result);

            if (!this.ByBuildDict.TryGetValue(result.Build, out var ByBuildStat))
            {
                ByBuildStat = new WinLose();
                this.ByBuildDict[result.Build] = ByBuildStat;
            }

            ByBuildStat.Add(result);
        }

        public void Write()
        {
            Console.WriteLine($"{this.all.Wins} - {this.all.Losses} (Draws: {this.all.Draws} Unknown: {this.all.Unknown}, Crashed: {this.all.Crashes})");
            Console.WriteLine($"Win percentage: {this.all.Wins * 100f / this.all.Stats.Count} %");
            Console.WriteLine("");

            if (this.Crashes.Any())
            {
                Console.WriteLine("Crashes:");

                foreach (var crash in this.Crashes)
                {
                    Console.WriteLine(crash);
                }
                Console.WriteLine("");
            }

            if (this.Unknown.Any())
            {
                Console.WriteLine("Results unknown:");

                foreach (var unknowns in this.Unknown)
                {
                    Console.WriteLine(unknowns);
                }
                Console.WriteLine("");
            }

            if (this.Draws.Any())
            {
                Console.WriteLine("");
                Console.WriteLine("Draws:");

                foreach (var draw in this.Draws)
                {
                    Console.WriteLine(draw);
                }
                Console.WriteLine("");
            }

            //var lastLine = DateTime.Now.Date.ToString("d");
            //const string separator = ";";
            //lastLine += separator + this.all.Stats.Average(x => x.Score);
            //lastLine += separator + this.all.Wins * 100f / this.all.Stats.Count;

            //foreach (KeyValuePair<string, WinLose> valuePair in this.VsDict.OrderBy(x => x.Key))
            //{
            //    var winlose = valuePair.Value;
            //    lastLine += separator + winlose.WinPercentage * 100f;
            //}

            //lastLine += separator + this.all.Stats.Count;

            //Console.WriteLine(lastLine);
        }

        public void DropUnknown()
        {
            void clearDict(Dictionary<string, WinLose> dict)
            {
                var remove = new HashSet<string>();
                foreach (KeyValuePair<string, WinLose> keyValuePair in dict)
                {
                    keyValuePair.Value.Stats.RemoveAll(x => x.Result == Result.Unknown);
                    if (keyValuePair.Value.Stats.Count == 0)
                    {
                        remove.Add(keyValuePair.Key);
                    }
                }

                foreach (var key in remove)
                {
                    dict.Remove(key);
                }
            }

            this.all.Stats.RemoveAll(x => x.Result == Result.Unknown);
            clearDict(this.VsDict);
            clearDict(this.MapDict);
            clearDict(this.MapVsDict);
            clearDict(this.ByMatchupDict);
            clearDict(this.ByBuildDict);

        }
    }
}