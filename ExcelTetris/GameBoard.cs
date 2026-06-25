using System.Drawing;

namespace ExcelTetris
{
    public class GameBoard
    {
        public const int Width = 10;
        public const int Height = 20;

        public Color[,] Grid { get; private set; }

        public GameBoard()
        {
            Grid = new Color[Height, Width];
            Clear();
        }

        public void Clear()
        {
            for (int r = 0; r < Height; r++)
            {
                for (int c = 0; c < Width; c++)
                {
                    Grid[r, c] = Color.Empty;
                }
            }
        }

        public bool CheckCollision(Tetromino block, int offsetX, int offsetY)
        {
            int size = block.Matrix.GetLength(0);
            for (int r = 0; r < size; r++)
            {
                for (int c = 0; c < size; c++)
                {
                    if (block.Matrix[r, c] != 0)
                    {
                        int targetX = block.X + c + offsetX;
                        int targetY = block.Y + r + offsetY;

                        if (targetX < 0 || targetX >= Width || targetY >= Height)
                        {
                            return true;
                        }

                        if (targetY >= 0)
                        {
                            if (Grid[targetY, targetX] != Color.Empty)
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            return false;
        }

        public void PlaceBlock(Tetromino block)
        {
            int size = block.Matrix.GetLength(0);
            for (int r = 0; r < size; r++)
            {
                for (int c = 0; c < size; c++)
                {
                    if (block.Matrix[r, c] != 0)
                    {
                        int targetX = block.X + c;
                        int targetY = block.Y + r;

                        if (targetY >= 0 && targetY < Height && targetX >= 0 && targetX < Width)
                        {
                            Grid[targetY, targetX] = block.BlockColor;
                        }
                    }
                }
            }
        }

        public int ClearLines()
        {
            int linesCleared = 0;
            for (int r = Height - 1; r >= 0; r--)
            {
                if (IsLineFull(r))
                {
                    DeleteLine(r);
                    linesCleared++;
                    r++; // 삭제 및 정밀 밀림 처리를 위한 위치 조정
                }
            }
            return linesCleared;
        }

        private bool IsLineFull(int row)
        {
            for (int c = 0; c < Width; c++)
            {
                if (Grid[row, c] == Color.Empty)
                {
                    return false;
                }
            }
            return true;
        }

        private void DeleteLine(int row)
        {
            for (int r = row; r > 0; r--)
            {
                for (int c = 0; c < Width; c++)
                {
                    Grid[r, c] = Grid[r - 1, c];
                }
            }
            for (int c = 0; c < Width; c++)
            {
                Grid[0, c] = Color.Empty;
            }
        }
    }
}