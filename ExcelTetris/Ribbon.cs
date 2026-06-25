using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace ExcelTetris
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ExcelTetris.Ribbon.xml");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnStartGame(Office.IRibbonControl control)
        {
            try
            {
                // "시작" 버튼 하나로 작동 중인 게임이 있을 경우 자동으로 폐기 종료 후 다시 시작합니다.
                ThisAddIn.Instance.Game.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"게임을 시작할 수 없습니다: {ex.Message}", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnEndGame(Office.IRibbonControl control)
        {
            try
            {
                ThisAddIn.Instance.Game.End();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"게임을 종료하는 중 오류가 발생했습니다: {ex.Message}", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnToggleGrayscale(Office.IRibbonControl control)
        {
            try
            {
                ThisAddIn.Instance.Game.ToggleGrayscale();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"흑백 모드 전환 중 오류가 발생했습니다: {ex.Message}", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnResetHighScore(Office.IRibbonControl control)
        {
            try
            {
                ThisAddIn.Instance.Game.ResetHighScore();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"최고 점수를 초기화하는 중 오류가 발생했습니다: {ex.Message}", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] names = asm.GetManifestResourceNames();
            for (int i = 0; i < names.Length; ++i)
            {
                if (string.Compare(resourceName, names[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(names[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
    }
}