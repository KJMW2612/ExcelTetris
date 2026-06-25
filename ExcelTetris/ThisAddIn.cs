using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTetris
{
    public partial class ThisAddIn
    {
        private static ThisAddIn _instance;
        public static ThisAddIn Instance => _instance;

        public TetrisGame Game { get; private set; }

        private System.Windows.Forms.Control _syncControl;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _instance = this;

            _syncControl = new System.Windows.Forms.Control();
            _syncControl.CreateControl();

            Game = new TetrisGame();

            // [추가] 엑셀이 저장을 개시하는 순간 작동할 사전 감지 이벤트를 연결합니다.
            this.Application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            if (Game != null)
            {
                Game.Shutdown();
            }

            if (_syncControl != null)
            {
                _syncControl.Dispose();
            }

            // 이벤트 언바인딩 진행
            this.Application.WorkbookBeforeSave -= Application_WorkbookBeforeSave;
        }

        // [추가] 저장 단추가 입력되었을 때 자동으로 가동되는 샌드박스 보안 청소 핸들러
        private void Application_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            if (Game != null)
            {
                // 디스크에 기록을 쓰기 전, 파일 내의 "Tetris" 시트를 강제로 언로드 및 영구 소멸시킵니다.
                Game.CleanUpBeforeSave(Wb);
            }
        }

        public void SafeInvoke(Action action)
        {
            if (_syncControl != null && _syncControl.IsHandleCreated)
            {
                _syncControl.BeginInvoke(action);
            }
            else
            {
                action();
            }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}