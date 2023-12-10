using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddIn_VSTO_Sample
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //将以下代码添加到 ThisAddIn_Startup 事件处理程序中。 此代码通过将 CustomTaskPane 对象添加到 MyUserControl
            //集合来创建新 CustomTaskPanes 。 代码还将显示任务窗格。
            myUserControl1 = new MyUserControl();
            myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl1, "My Task Pane");
            myCustomTaskPane.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private MyUserControl myUserControl1;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }

}
