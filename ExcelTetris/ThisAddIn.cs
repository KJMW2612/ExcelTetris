using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTetris
{
    public partial class ThisAddIn
    {
        private static ThisAddIn _instance;
        public static ThisAddIn Instance => _instance;

        public TetrisGame Game { get; private set; }

        // 메인 UI 스레드에서 동작을 대기하는 윈도우 메시지 마샬러
        private System.Windows.Forms.Control _syncControl;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _instance = this;

            // 1. 메인 UI 스레드에서 더미 컨트롤을 인스턴스화하고 강제로 HWND 생성
            _syncControl = new System.Windows.Forms.Control();
            _syncControl.CreateControl();

            Game = new TetrisGame();
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
        }

        // 백그라운드 호출이나 훅 스레드 이벤트를 메인 Excel UI 스레드로 100% 보장하며 넘겨주는 함수
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