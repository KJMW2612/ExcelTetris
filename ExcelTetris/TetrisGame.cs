using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTetris
{
    public class TetrisGame
    {
        private Timer _timer;
        private GameBoard _board;
        private ScoreManager _scoreManager;
        private InputManager _inputManager;

        private Tetromino _currentBlock;
        private Tetromino _nextBlock;
        private Tetromino _holdBlock;

        private List<TetrominoType> _bag = new List<TetrominoType>();
        private Random _random = new Random();

        private bool _isGameOver;
        private bool _hasHeldThisTurn;
        private bool _isRunning;
        private bool _isBossModeActive;

        private bool _isGrayscaleMode;

        private int _level;
        private int _linesClearedTotal;

        private Excel.Worksheet _tetrisSheet;
        private Excel.Worksheet _previousSheet;

        private Color[,] _screenCache = new Color[GameBoard.Height, GameBoard.Width];

        public TetrisGame()
        {
            _board = new GameBoard();
            _scoreManager = new ScoreManager();

            _inputManager = new InputManager(
                HandleKeyPress,
                () => _isBossModeActive,
                () => _isRunning && !_isGameOver && !_isBossModeActive
            );

            _timer = new Timer();
            _timer.Interval = 500;
            _timer.Tick += (s, e) => UpdateGame();
        }

        public void Start()
        {
            ForceExitEditMode();

            if (_isRunning)
            {
                End();
            }

            InitializeGameSheet();
            _board.Clear();
            _scoreManager.Reset();
            _bag.Clear();

            _linesClearedTotal = 0;
            _level = 1;

            _currentBlock = SpawnBlock();
            _nextBlock = SpawnBlock();
            _holdBlock = null;

            _isGameOver = false;
            _hasHeldThisTurn = false;
            _isBossModeActive = false;
            _isRunning = true;

            ClearCache();

            _inputManager.Start();
            _timer.Interval = GetSpeedForLevel(_level);
            _timer.Start();

            Render();
        }

        private void ForceExitEditMode()
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            if (app == null) return;

            bool isEditing = false;
            try
            {
                app.Interactive = false;
                app.Interactive = true;
            }
            catch
            {
                isEditing = true;
            }

            if (isEditing)
            {
                try
                {
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    System.Threading.Thread.Sleep(50);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"편집 모드 자동 무력화 실패: {ex.Message}");
                }
            }
        }

        // [신규 보안 해결책] 저장 단추 입력 시 파일 내의 테트리스 게임 흔적을 즉석에서 완전하게 소멸시키는 가드 함수
        public void CleanUpBeforeSave(Excel.Workbook wb)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            bool originalDisplayAlerts = app.DisplayAlerts;

            try
            {
                // 시트 영구삭제 경고창("이 시트를 영구적으로 삭제합니까?")이 팝업되는 것을 비동기로 방지합니다.
                app.DisplayAlerts = false;

                Excel.Worksheet targetSheet = null;
                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    if (ws.Name == "Tetris")
                    {
                        targetSheet = ws;
                        break;
                    }
                }

                if (targetSheet != null)
                {
                    // 1. 동작 중인 실시간 낙하 타이머 중단
                    End();

                    // 2. 현재 열려있는 테트리스 시트 이외의 업무용 대체 시트 색인
                    Excel.Worksheet alternativeSheet = null;
                    foreach (Excel.Worksheet ws in wb.Worksheets)
                    {
                        if (ws.Name != "Tetris")
                        {
                            alternativeSheet = ws;
                            break;
                        }
                    }

                    // 3. 만약 파일 내에 업무용 시트가 아예 존재하지 않는 최악의 경우, 일반 시트 강제 가설
                    if (alternativeSheet == null)
                    {
                        alternativeSheet = wb.Worksheets.Add() as Excel.Worksheet;
                    }

                    // 4. 활성화 화면을 업무용 시트로 피신 변경 (활성 시트 자체는 삭제가 거부되는 엑셀 룰 방어)
                    alternativeSheet.Activate();

                    // 5. 파일 내부에서 테트리스 흔적 시트 영구 파괴
                    targetSheet.Delete();
                    _tetrisSheet = null;
                }
            }
            catch { }
            finally
            {
                // 원래 상태로 경고 알림 복구
                app.DisplayAlerts = originalDisplayAlerts;
            }
        }

        public void End()
        {
            _timer.Stop();
            _isRunning = false;

            if (_tetrisSheet != null)
            {
                try
                {
                    _tetrisSheet.Unprotect();

                    Excel.Range goRange = _tetrisSheet.Range["C11:L12"];
                    goRange.Merge();
                    goRange.Value2 = "GAME OVER";
                    goRange.Font.Color = ColorTranslator.ToOle(Color.Red);
                    goRange.Font.Size = 20;
                    goRange.Font.Bold = true;
                    goRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    goRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                    _tetrisSheet.Protect(
                        Password: Type.Missing,
                        DrawingObjects: false,
                        Contents: false,
                        Scenarios: false,
                        UserInterfaceOnly: true
                    );
                    _tetrisSheet.EnableSelection = Excel.XlEnableSelection.xlNoSelection;
                }
                catch { }
            }
        }

        public void Shutdown()
        {
            _timer.Stop();
            _inputManager.Stop();
        }

        public void ResetHighScore()
        {
            _scoreManager.ForceResetHighScore();
            if (_isRunning && !_isBossModeActive)
            {
                Render();
            }
        }

        public void ToggleGrayscale()
        {
            _isGrayscaleMode = !_isGrayscaleMode;
            ClearCache();
            Render();
        }

        private Color GetDisplayColor(Color originalColor)
        {
            if (originalColor == Color.Empty || originalColor == Color.FromArgb(32, 32, 32) || originalColor == Color.FromArgb(20, 20, 20))
            {
                return originalColor;
            }

            if (_isGrayscaleMode)
            {
                int gray = (int)(originalColor.R * 0.299 + originalColor.G * 0.587 + originalColor.B * 0.114);
                if (gray < 70) gray = 70;
                return Color.FromArgb(gray, gray, gray);
            }

            return originalColor;
        }

        private void InitializeGameSheet()
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            _previousSheet = wb.ActiveSheet as Excel.Worksheet;

            bool isNewSheet = true;
            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                if (ws.Name == "Tetris")
                {
                    _tetrisSheet = ws;
                    isNewSheet = false;
                    break;
                }
            }

            if (isNewSheet)
            {
                _tetrisSheet = wb.Worksheets.Add() as Excel.Worksheet;
                _tetrisSheet.Name = "Tetris";
                _tetrisSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                _tetrisSheet.Activate();

                FormatGameLayout(_tetrisSheet);
            }
            else
            {
                _tetrisSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                _tetrisSheet.Activate();

                ResetGameLayout();
            }
        }

        private void ResetGameLayout()
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            app.ScreenUpdating = false;

            try
            {
                _tetrisSheet.Unprotect();
            }
            catch { }

            Excel.Range boardRange = _tetrisSheet.Range["C3:L22"];
            boardRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(32, 32, 32));

            Excel.Range goRange = _tetrisSheet.Range["C11:L12"];
            try
            {
                goRange.UnMerge();
            }
            catch { }
            goRange.Value2 = null;
            goRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(32, 32, 32));
            ApplyBorders(boardRange);

            RenderMiniPreview("O3:R6", null);
            RenderMiniPreview("O9:R12", null);

            _tetrisSheet.Range["O15"].Value2 = "0";
            _tetrisSheet.Range["O21"].Value2 = "1";

            _tetrisSheet.Protect(
                Password: Type.Missing,
                DrawingObjects: false,
                Contents: false,
                Scenarios: false,
                UserInterfaceOnly: true
            );
            _tetrisSheet.EnableSelection = Excel.XlEnableSelection.xlNoSelection;

            app.ScreenUpdating = true;
        }

        private void FormatGameLayout(Excel.Worksheet ws)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            app.ScreenUpdating = false;

            try
            {
                ws.Unprotect();
            }
            catch { }

            ws.Cells.Clear();

            for (int c = 1; c <= 26; c++)
            {
                Excel.Range col = ws.Columns[c] as Excel.Range;
                if (c >= 3 && c <= 12) col.ColumnWidth = 3.5;
                else if (c >= 15 && c <= 18) col.ColumnWidth = 3.5;
                else if (c >= 21 && c <= 24) col.ColumnWidth = 3.5;
                else col.ColumnWidth = 1.5;
            }

            for (int r = 1; r <= 23; r++)
            {
                Excel.Range row = ws.Rows[r] as Excel.Range;
                row.RowHeight = 22;
            }

            Excel.Range boardRange = ws.Range[ws.Cells[3, 3], ws.Cells[22, 12]];
            boardRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(32, 32, 32));
            ApplyBorders(boardRange);

            Excel.Range holdLabel = ws.Range["O2:R2"];
            holdLabel.Merge();
            holdLabel.Value2 = "HOLD";
            StyleLabel(holdLabel);

            Excel.Range holdBox = ws.Range["O3:R6"];
            holdBox.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(32, 32, 32));
            ApplyBorders(holdBox);

            Excel.Range nextLabel = ws.Range["O8:R8"];
            nextLabel.Merge();
            nextLabel.Value2 = "NEXT";
            StyleLabel(nextLabel);

            Excel.Range nextBox = ws.Range["O9:R12"];
            nextBox.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(32, 32, 32));
            ApplyBorders(nextBox);

            Excel.Range scoreLabel = ws.Range["O14:R14"];
            scoreLabel.Merge();
            scoreLabel.Value2 = "SCORE";
            StyleLabel(scoreLabel);

            Excel.Range scoreVal = ws.Range["O15:R15"];
            scoreVal.Merge();
            scoreVal.Value2 = "0";
            StyleValue(scoreVal);

            Excel.Range hsLabel = ws.Range["O17:R17"];
            hsLabel.Merge();
            hsLabel.Value2 = "HIGH SCORE";
            StyleLabel(hsLabel);

            Excel.Range hsVal = ws.Range["O18:R18"];
            hsVal.Merge();
            hsVal.Value2 = "0";
            StyleValue(hsVal);

            Excel.Range levelLabel = ws.Range["O20:R20"];
            levelLabel.Merge();
            levelLabel.Value2 = "LEVEL";
            StyleLabel(levelLabel);

            Excel.Range levelVal = ws.Range["O21:R21"];
            levelVal.Merge();
            levelVal.Value2 = "1";
            StyleValue(levelVal);

            // ----------------- [조작 가이드라인 우측 패널 레이아웃 적용] -----------------
            Excel.Range ctrlLabel = ws.Range["U2:X2"];
            ctrlLabel.Merge();
            ctrlLabel.Value2 = "CONTROLS";
            StyleLabel(ctrlLabel);
            ctrlLabel.Font.Color = ColorTranslator.ToOle(Color.Black);
            ctrlLabel.Font.Size = 11;

            string[] controlLines = new string[] {
                "A : 왼쪽 이동",
                "D : 오른쪽 이동",
                "S : 부드러운 하강",
                "Space : 즉시 낙하",
                "W : 시계 회전",
                "E : 반시계 회전",
                "Q : 홀드(보관/교체)",
                "F : 흑백 모드 토글",
                "ESC : 긴급 숨김"
            };

            for (int i = 0; i < controlLines.Length; i++)
            {
                int targetRow = 4 + (i * 2);
                Excel.Range lineRange = ws.Range[ws.Cells[targetRow, 21], ws.Cells[targetRow, 24]];
                lineRange.Merge();
                lineRange.Value2 = controlLines[i];
                lineRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                lineRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                lineRange.Font.Name = "Consolas";
                lineRange.Font.Size = 10;

                lineRange.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;

                if (i == 8)
                {
                    lineRange.Font.Color = ColorTranslator.ToOle(Color.FromArgb(235, 30, 30));
                    lineRange.Font.Bold = true;
                }
                else
                {
                    lineRange.Font.Color = ColorTranslator.ToOle(Color.Black);
                    lineRange.Font.Bold = false;
                }
            }
            // -------------------------------------------------------------------------

            try
            {
                ws.Range["A1:Z23"].Select();
                app.ActiveWindow.Zoom = true;
                ws.Range["A1"].Select();
            }
            catch { }

            ws.Protect(
                Password: Type.Missing,
                DrawingObjects: false,
                Contents: false,
                Scenarios: false,
                UserInterfaceOnly: true
            );
            ws.EnableSelection = Excel.XlEnableSelection.xlNoSelection;

            app.ScreenUpdating = true;
        }

        private void ApplyBorders(Excel.Range range)
        {
            Excel.Borders borders = range.Borders;
            borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            borders.Weight = Excel.XlBorderWeight.xlThin;
            borders.Color = ColorTranslator.ToOle(Color.FromArgb(64, 64, 64));
        }

        private void StyleLabel(Excel.Range range)
        {
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.Font.Bold = true;
            range.Font.Name = "Consolas";
            range.Font.Size = 10;
            range.Font.Color = ColorTranslator.ToOle(Color.FromArgb(120, 120, 120));
        }

        private void StyleValue(Excel.Range range)
        {
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.Font.Bold = true;
            range.Font.Name = "Consolas";
            range.Font.Size = 11;
            range.Font.Color = ColorTranslator.ToOle(Color.White);
            range.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(20, 20, 20));
            ApplyBorders(range);
        }

        private void ClearCache()
        {
            for (int r = 0; r < GameBoard.Height; r++)
            {
                for (int c = 0; c < GameBoard.Width; c++)
                {
                    _screenCache[r, c] = Color.Empty;
                }
            }
        }

        private Tetromino SpawnBlock()
        {
            TetrominoType type = GetNextFromBag();
            var block = new Tetromino(type);
            block.X = (GameBoard.Width - block.Matrix.GetLength(1)) / 2;
            block.Y = 0;
            return block;
        }

        private TetrominoType GetNextFromBag()
        {
            if (_bag.Count == 0)
            {
                var candidates = new List<TetrominoType> {
                    TetrominoType.I, TetrominoType.O, TetrominoType.T,
                    TetrominoType.S, TetrominoType.Z, TetrominoType.J, TetrominoType.L
                };

                for (int i = candidates.Count - 1; i > 0; i--)
                {
                    int idx = _random.Next(i + 1);
                    var temp = candidates[i];
                    candidates[i] = candidates[idx];
                    candidates[idx] = temp;
                }
                _bag.AddRange(candidates);
            }

            TetrominoType next = _bag[0];
            _bag.RemoveAt(0);
            return next;
        }

        private void HandleKeyPress(Keys key)
        {
            if (key == Keys.Escape)
            {
                ToggleBossMode();
                return;
            }

            if (!_isRunning || _isBossModeActive || _isGameOver) return;

            switch (key)
            {
                case Keys.A:
                    MoveBlock(-1, 0);
                    break;
                case Keys.D:
                    MoveBlock(1, 0);
                    break;
                case Keys.S:
                    MoveBlock(0, 1);
                    break;
                case Keys.W:
                    RotateBlock(true);
                    break;
                case Keys.E:
                    RotateBlock(false);
                    break;
                case Keys.Space:
                    HardDrop();
                    break;
                case Keys.Q:
                    ExecuteHold();
                    break;
                case Keys.F:
                    ToggleGrayscale();
                    break;
            }
            Render();
        }

        private void MoveBlock(int dx, int dy)
        {
            if (!_board.CheckCollision(_currentBlock, dx, dy))
            {
                _currentBlock.X += dx;
                _currentBlock.Y += dy;
            }
            else if (dy > 0)
            {
                LockAndSpawnNext();
            }
        }

        private void RotateBlock(bool clockwise)
        {
            var temp = _currentBlock.Clone();
            if (clockwise) temp.RotateClockwise();
            else temp.RotateCounterClockwise();

            if (!_board.CheckCollision(temp, 0, 0))
            {
                if (clockwise) _currentBlock.RotateClockwise();
                else _currentBlock.RotateCounterClockwise();
            }
        }

        private void HardDrop()
        {
            while (!_board.CheckCollision(_currentBlock, 0, 1))
            {
                _currentBlock.Y++;
            }
            LockAndSpawnNext();
        }

        private void ExecuteHold()
        {
            if (_hasHeldThisTurn) return;

            if (_holdBlock == null)
            {
                _holdBlock = new Tetromino(_currentBlock.Type);
                _currentBlock = _nextBlock;
                _nextBlock = SpawnBlock();
            }
            else
            {
                var temp = _holdBlock;
                _holdBlock = new Tetromino(_currentBlock.Type);
                _currentBlock = temp;
                _currentBlock.X = (GameBoard.Width - _currentBlock.Matrix.GetLength(1)) / 2;
                _currentBlock.Y = 0;
            }

            _hasHeldThisTurn = true;
        }

        private void LockAndSpawnNext()
        {
            _board.PlaceBlock(_currentBlock);
            int cleared = _board.ClearLines();
            if (cleared > 0)
            {
                _scoreManager.AddLines(cleared);

                _linesClearedTotal += cleared;
                int newLevel = (_linesClearedTotal / 10) + 1;

                if (newLevel != _level)
                {
                    _level = newLevel;
                    _timer.Interval = GetSpeedForLevel(_level);
                }
            }

            _currentBlock = _nextBlock;
            _nextBlock = SpawnBlock();
            _hasHeldThisTurn = false;

            if (_board.CheckCollision(_currentBlock, 0, 0))
            {
                _isGameOver = true;
                End();
            }
        }

        private int GetSpeedForLevel(int level)
        {
            double seconds = Math.Pow(0.8 - ((level - 1) * 0.007), level - 1);
            int ms = (int)(seconds * 1000);
            return Math.Max(100, ms);
        }

        private void UpdateGame()
        {
            if (_isGameOver || _isBossModeActive || !_isRunning) return;

            MoveBlock(0, 1);
            Render();
        }

        private void ToggleBossMode()
        {
            Excel.Application app = Globals.ThisAddIn.Application;

            if (!_isBossModeActive)
            {
                _isBossModeActive = true;
                _timer.Stop();

                app.ScreenUpdating = false;

                if (_previousSheet != null && _previousSheet.Name != "Tetris")
                {
                    _previousSheet.Activate();
                }
                else
                {
                    Excel.Worksheet defaultWS = app.ActiveWorkbook.Worksheets.Add() as Excel.Worksheet;
                    _previousSheet = defaultWS;
                    _previousSheet.Activate();
                }

                _tetrisSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                app.ScreenUpdating = true;
            }
            else
            {
                _isBossModeActive = false;

                _tetrisSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                _tetrisSheet.Activate();

                ClearCache();
                Render();

                if (_isRunning && !_isGameOver)
                {
                    _timer.Start();
                }
            }
        }

        private void Render()
        {
            if (_tetrisSheet == null || _isBossModeActive) return;

            try
            {
                Color[,] visualGrid = new Color[GameBoard.Height, GameBoard.Width];
                for (int r = 0; r < GameBoard.Height; r++)
                {
                    for (int c = 0; c < GameBoard.Width; c++)
                    {
                        visualGrid[r, c] = _board.Grid[r, c];
                    }
                }

                if (_currentBlock != null && !_isGameOver)
                {
                    int size = _currentBlock.Matrix.GetLength(0);
                    for (int r = 0; r < size; r++)
                    {
                        for (int c = 0; c < size; c++)
                        {
                            if (_currentBlock.Matrix[r, c] != 0)
                            {
                                int ty = _currentBlock.Y + r;
                                int tx = _currentBlock.X + c;
                                if (ty >= 0 && ty < GameBoard.Height && tx >= 0 && tx < GameBoard.Width)
                                {
                                    visualGrid[ty, tx] = _currentBlock.BlockColor;
                                }
                            }
                        }
                    }
                }

                Excel.Application app = Globals.ThisAddIn.Application;
                app.ScreenUpdating = false;

                for (int r = 0; r < GameBoard.Height; r++)
                {
                    for (int c = 0; c < GameBoard.Width; c++)
                    {
                        Color rawColor = visualGrid[r, c] == Color.Empty ? Color.FromArgb(32, 32, 32) : visualGrid[r, c];
                        Color targetColor = GetDisplayColor(rawColor);

                        if (_screenCache[r, c] != targetColor)
                        {
                            Excel.Range cell = _tetrisSheet.Cells[r + 3, c + 3] as Excel.Range;
                            cell.Interior.Color = ColorTranslator.ToOle(targetColor);
                            _screenCache[r, c] = targetColor;
                        }
                    }
                }

                RenderMiniPreview("O3:R6", _holdBlock);
                RenderMiniPreview("O9:R12", _nextBlock);

                _tetrisSheet.Range["O15"].Value2 = _scoreManager.CurrentScore.ToString();
                _tetrisSheet.Range["O18"].Value2 = _scoreManager.HighScore.ToString();
                _tetrisSheet.Range["O21"].Value2 = _level.ToString();

                app.ScreenUpdating = true;
            }
            catch (System.Runtime.InteropServices.COMException ex) when ((uint)ex.ErrorCode == 0x800A03EC)
            {
                // COM 바인딩 버그 방지
            }
            catch
            {
                // 예외 제어
            }
        }

        private void RenderMiniPreview(string rangeAddress, Tetromino block)
        {
            Excel.Range container = _tetrisSheet.Range[rangeAddress];
            container.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(32, 32, 32));

            if (block == null) return;

            int size = block.Matrix.GetLength(0);
            int startRow = container.Row;
            int startCol = container.Column;

            int offsetX = (4 - size) / 2;
            int offsetY = (4 - size) / 2;

            for (int r = 0; r < size; r++)
            {
                for (int c = 0; c < size; c++)
                    if (block.Matrix[r, c] != 0)
                    {
                        Excel.Range cell = _tetrisSheet.Cells[startRow + r + offsetY, startCol + c + offsetX] as Excel.Range;
                        Color targetColor = GetDisplayColor(block.BlockColor);
                        cell.Interior.Color = ColorTranslator.ToOle(targetColor);
                    }
            }
        }
    }
}