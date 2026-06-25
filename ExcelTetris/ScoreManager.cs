namespace ExcelTetris
{
    public class ScoreManager
    {
        public int CurrentScore { get; private set; }
        public int HighScore { get; private set; }

        public ScoreManager()
        {
            CurrentScore = 0;
            HighScore = HighScoreManager.LoadHighScore();
        }

        public void AddLines(int linesCount)
        {
            switch (linesCount)
            {
                case 1: CurrentScore += 100; break;
                case 2: CurrentScore += 300; break;
                case 3: CurrentScore += 500; break;
                case 4: CurrentScore += 800; break;
            }

            if (CurrentScore > HighScore)
            {
                HighScore = CurrentScore;
                HighScoreManager.SaveHighScore(HighScore);
            }
        }

        public void Reset()
        {
            CurrentScore = 0;
            HighScore = HighScoreManager.LoadHighScore();
        }

        public void ForceResetHighScore()
        {
            HighScoreManager.ResetHighScore();
            HighScore = 0;
        }
    }
}