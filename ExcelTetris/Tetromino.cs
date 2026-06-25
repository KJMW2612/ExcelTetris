using System;
using System.Drawing;

namespace ExcelTetris
{
    public enum TetrominoType { Empty, I, O, T, S, Z, J, L }

    public class Tetromino
    {
        public TetrominoType Type { get; set; }
        public int[,] Matrix { get; set; }
        public Color BlockColor { get; set; }
        public int X { get; set; }
        public int Y { get; set; }

        public Tetromino(TetrominoType type)
        {
            Type = type;
            InitializeMatrixAndColor();
        }

        private void InitializeMatrixAndColor()
        {
            switch (Type)
            {
                case TetrominoType.I:
                    Matrix = new int[,] {
                        {0,0,0,0},
                        {1,1,1,1},
                        {0,0,0,0},
                        {0,0,0,0}
                    };
                    BlockColor = Color.FromArgb(0, 240, 240); // Cyan
                    break;
                case TetrominoType.O:
                    Matrix = new int[,] {
                        {1,1},
                        {1,1}
                    };
                    BlockColor = Color.FromArgb(240, 240, 0); // Yellow
                    break;
                case TetrominoType.T:
                    Matrix = new int[,] {
                        {0,1,0},
                        {1,1,1},
                        {0,0,0}
                    };
                    BlockColor = Color.FromArgb(160, 0, 240); // Purple
                    break;
                case TetrominoType.S:
                    Matrix = new int[,] {
                        {0,1,1},
                        {1,1,0},
                        {0,0,0}
                    };
                    BlockColor = Color.FromArgb(0, 240, 0); // Green
                    break;
                case TetrominoType.Z:
                    Matrix = new int[,] {
                        {1,1,0},
                        {0,1,1},
                        {0,0,0}
                    };
                    BlockColor = Color.FromArgb(240, 0, 0); // Red
                    break;
                case TetrominoType.J:
                    Matrix = new int[,] {
                        {1,0,0},
                        {1,1,1},
                        {0,0,0}
                    };
                    BlockColor = Color.FromArgb(0, 0, 240); // Blue
                    break;
                case TetrominoType.L:
                    Matrix = new int[,] {
                        {0,0,1},
                        {1,1,1},
                        {0,0,0}
                    };
                    BlockColor = Color.FromArgb(240, 160, 0); // Orange
                    break;
                default:
                    Matrix = new int[0, 0];
                    BlockColor = Color.White;
                    break;
            }
        }

        public Tetromino Clone()
        {
            return new Tetromino(Type)
            {
                X = this.X,
                Y = this.Y,
                BlockColor = this.BlockColor,
                Matrix = (int[,])this.Matrix.Clone()
            };
        }

        public void RotateClockwise()
        {
            int size = Matrix.GetLength(0);
            int[,] rotated = new int[size, size];
            for (int r = 0; r < size; r++)
            {
                for (int c = 0; c < size; c++)
                {
                    rotated[c, size - 1 - r] = Matrix[r, c];
                }
            }
            Matrix = rotated;
        }

        public void RotateCounterClockwise()
        {
            int size = Matrix.GetLength(0);
            int[,] rotated = new int[size, size];
            for (int r = 0; r < size; r++)
            {
                for (int c = 0; c < size; c++)
                {
                    rotated[size - 1 - c, r] = Matrix[r, c];
                }
            }
            Matrix = rotated;
        }
    }
}