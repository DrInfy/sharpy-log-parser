using System.Collections.Generic;
using System.Linq;

namespace sc2DataReader.GameData
{
    class WinLose
    {
        public List<GameStats> Stats = new List<GameStats>();

        public int TotalGames => Stats.Count - Unknown;
        public int Wins => Stats.Count(x => x.Result == Result.Victory);
        public int Losses => Stats.Count(x => x.Result == Result.Defeat);
        public int Draws => Stats.Count(x => x.Result == Result.Draw);
        public int Unknown => Stats.Count(x => x.Result == Result.Unknown);
        public int Crashes => Stats.Count(x => x.Result == Result.Crash);

        public float WinPercentage => (float) Wins / TotalGames;

        public void Add(GameStats result)
        {
            this.Stats.Add(result);
        }
    }
}