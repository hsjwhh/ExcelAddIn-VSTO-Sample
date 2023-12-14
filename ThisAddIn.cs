using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Core;

namespace ExcelAddIn_VSTO_Sample
{
    public partial class ThisAddIn
    {
        private MyUserControl myUserControl1;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {   
            // 添加 taskpane
            // CustomTaskPanesAdd();
            // 添加选中单元格事件
            this.Application.SheetSelectionChange += Application_SheetSelectionChange;
            // 添加自定义右键菜单项目
            this.Application.SheetBeforeRightClick += Application_SheetBeforeRightClick;
        }

        #region 添加右键菜单项目
        private void Application_SheetBeforeRightClick(object Sh, Excel.Range Target, ref bool Cancel)
        {
            // 获取右键点击的单元格
            // Excel.Range clickedCell = Target.Cells[1, 1];
            // 获取 CommandBars 对象
            // 调整行高，所以右键所在行的菜单添加自定义项目
            Office.CommandBar originalContextMen = Target.Application.CommandBars["Row"];
            originalContextMen.Reset();
            // 添加自定义菜单项到菜单第一位置
            CommandBarControl customMenuItem = originalContextMen.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, 1, true);
            // 转换为CommandBarButton
            CommandBarButton customButton = (CommandBarButton)customMenuItem;
            // 设置菜单项属性
            customButton.Caption = "设置行高为 25";
            customButton.Tag = "SetRowHeight";
            // customButton.FaceId = 22;
            customButton.Click += SetRowHeightMenuItemClick;
        }

        private void SetRowHeightMenuItemClick(Office.CommandBarButton ctrl, ref bool cancelDefault)
        {
            // 获取当前选定的单元格
            Excel.Range selectedRange = this.Application.Selection;

            // 设置所选行的高度为25
            selectedRange.Rows.RowHeight = 25;
        }
        #endregion

        #region 选中单元格有条件格式的话就 刷新计算
        /// <summary>
        /// 单元格格式 cell("ROW")=ROW(),实现单击单元格给高亮该行
        /// 在这个示例中，`HasConditionalFormat` 方法检查给定单元格是否应用了条件格式。如果应用了条件格式，就触发 F9 操作。
        /// 请记住，这种方法是基于检查条件格式的数量，可能并不是非常严格的验证，因为条件格式的具体设置可能会更加复杂。
        /// 你可能需要根据你的具体需求进行更详细的条件格式检查。
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
                // 手动刷新计算
                clickedCell.Application.Calculate();
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
