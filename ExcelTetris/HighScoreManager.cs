using System;
using System.IO;

namespace ExcelTetris
{
    public static class HighScoreManager
    {
        private static readonly string FolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelTetris");
        private static readonly string FilePath = Path.Combine(FolderPath, "highscore.txt");

        public static int LoadHighScore()
        {
            try
            {
                if (File.Exists(FilePath))
                {
                    string content = File.ReadAllText(FilePath);
                    if (int.TryParse(content, out int score))
                    {
                        return score;
                    }
                }
            }
            catch
            {
                // 실패 시 기본값 0 반환
            }
            return 0;
        }

        public static void SaveHighScore(int score)
        {
            try
            {
                if (!Directory.Exists(FolderPath))
                {
                    Directory.CreateDirectory(FolderPath);
                }
                File.WriteAllText(FilePath, score.ToString());
            }
            catch
            {
                // 쓰기 실패 처리 무시
            }
        }

        public static void ResetHighScore()
        {
            SaveHighScore(0);
        }
    }
}