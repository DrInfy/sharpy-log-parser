using System;

namespace sc2DataReader.GameData
{
    /// <summary>
    /// Selected statistics about a single game.
    /// </summary>
    public class GameStats
    {
        /// <summary>
        /// When the game was started.
        /// </summary>
        public DateTime StartedOn { get; set; }

        /// <summary>
        /// Bot version the game was played on.
        /// </summary>
        public string BotVersion { get; set; }

        public Result Result { get; set; }
        public string GameName { get; set; }

        public string MyRace { get; set; }

        public string Opponent { get; set; }
        public string OpponentId { get; set; }
        public string OpponentRace { get; set; }

        /// <summary>
        /// Returns matchup of the game, eg. "PvZ", short for "Protoss vs. Zerg".
        /// </summary>
        public string RaceMatchup {
            get
            {
                try
                {
                    return $"{MyRace[0]}v{OpponentRace[0]}";
                } catch (Exception)
                {
                    return "Unknown";
                }
            }
        }

        public string Map { get; set; }
        public TimeSpan Duration { get; set; }

        public int StepTimeMin { get; set; }
        public int StepTimeAvg { get; set; }
        public int StepTimeMax { get; set; }

        public int MyLost => this.MineralsLost + this.GasLost;
        public int MineralsLost { get; set; }
        public int GasLost { get; set; }
        public int EnemyLost => this.EnemyMineralsLost + this.EnemyGasLost;
        public int EnemyMineralsLost { get; set; }
        public int EnemyGasLost { get; set; }
        public int Workers { get; set; }
        public int EnemyWorkers { get; set; }

        private string _build;
        public string Build
        {
            get => _build ?? "Unknown";
            set => _build = value;
        }
        public string DummyBuild { get; set; }
        public string RushDistance { get; set; }

        public int MaxGas { get; set; }
        public int MaxMinerals { get; set; }

        public float AverageMinerals { get; set; }
        public float AverageGas { get; set; }
        public float Loss { get; set; }

        public double Score
        {
            get
            {
                var myLostResources = this.MineralsLost + this.GasLost;
                var enemyLostResources = this.EnemyMineralsLost + this.EnemyGasLost;
                var points = Math.Min((enemyLostResources + 50d) / (myLostResources + 50d) * 10d, 35d);

                switch (this.Result)
                {
                    case Result.Victory:
                        points += 50;
                        points += Math.Min(15, Math.Max(0, 22 - (7 + this.Duration.TotalMinutes) * 0.6));
                        break;
                    case Result.Draw:
                        points += 25;
                        break;
                    case Result.Defeat:
                        points += 15 - Math.Min(15, Math.Max(0, 22 - (7 + this.Duration.TotalMinutes) * 0.6));
                        break;
                    case Result.Crash:
                    case Result.Unknown:
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }

                return points;
            }
        }


        public override string ToString()
        {
            return $"{this.GameName}\r\n" +
                   $"\tDuration: {this.Duration:mm\\:ss} Build: {this.Build} | Dummy build: {this.DummyBuild ?? "-"}";
        }

        public string ResultToString()
        {
            return this.Result.ToString();
        }
    }
}