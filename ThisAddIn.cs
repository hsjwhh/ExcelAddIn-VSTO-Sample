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
        private MyUserControl myUserControl1;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {   
            // 添加 taskpane
            CustomTaskPanesAdd();
            // 添加选中单元格事件
            this.Application.SheetSelectionChange += Application_SheetSelectionChange;
        }

        #region 选中单元格有条件格式的话就 f9 刷新计算
        /// <summary>
        /// 单元格格式 cell("ROW")=ROW(),实现单击单元格给高亮该行
        /// </summary>
        /// <param name="Sh"></param>
        /// <param name="Target"></param>
        private void Application_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            // 获取点击的单元格
            Excel.Range clickedCell = Target.Cells[1, 1];

            // 检查单元格是否有条件格式
            if (HasConditionalFormat(clickedCell))
            {
                // 模拟按下 F9 键
                clickedCell.Application.SendKeys("{F9}");
            }
        }

        private bool HasConditionalFormat(Excel.Range cell)
        {
            // 获取单元格的条件格式
            Excel.FormatConditions conditions = cell.FormatConditions;

            // 检查是否有条件格式
            return (conditions != null && conditions.Count > 0);
        }
        #endregion

        private void CustomTaskPanesAdd()
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
