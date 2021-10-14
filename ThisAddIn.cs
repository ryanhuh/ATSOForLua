using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using ExcelAddIn1.Utils;

namespace ExcelAddIn1
{    
    public partial class ThisAddIn
    {
        private ContextMenuHandler _contextMenuHandler;        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _contextMenuHandler = new ContextMenuHandler();
            Application.SheetBeforeRightClick += Application_SheetBeforeRightClick;            

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }        
        private void Application_SheetBeforeRightClick(object Sh, Excel.Range Target, ref bool Cancel)
        {
            _contextMenuHandler.InitializelMenu();
        }

        #region VSTO에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
