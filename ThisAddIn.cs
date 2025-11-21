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

        public bool spotlightEnable = false;
        private MyUserControl myUserControl1;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 添加 taskpane
            // CustomTaskPanesAdd();
            // 读取上次保存的状态
            spotlightEnable = Properties.Settings.Default.SpotlightEnabled;
            if (spotlightEnable)
            {
                // 添加选中单元格事件
                this.Application.SheetSelectionChange += Application_SheetSelectionChange;
            }
            
            // 添加自定义右键菜单项目
            //this.Application.SheetBeforeRightClick += Application_SheetBeforeRightClick;

        }


        #region 添加右键菜单项目


        private void ToggleSpotlightMenuItemClick(Office.CommandBarButton ctrl, ref bool cancelDefault)
        {
            try
            {
                // 切换状态（你已有的方法）
                ToggleSpotlight();

                // 更新菜单显示文本（当前上下文的菜单按钮实例）
                try
                {
                    ctrl.Caption = spotlightEnable ? "聚光灯：已开启" : "聚光灯：已关闭";
                }
                catch { }
            }
            catch { }
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
        /// 单元格条件格式=OR(AND(ROW()>=sRow,ROW()<=eRow),AND(COLUMN()>=sColumn,COLUMN()<=eColumn)),实现单击单元格给高亮该行
        /// 在这个示例中，`HasConditionalFormat` 方法检查给定单元格是否应用了条件格式。如果应用了条件格式，就更新指定名称的范围。
        /// 请记住，这种方法是基于检查条件格式的数量，可能并不是非常严格的验证，因为条件格式的具体设置可能会更加复杂。
        /// 你可能需要根据你的具体需求进行更详细的条件格式检查。
        /// </summary>
        /// <param name="Sh"></param>
        /// <param name="Target"></param>
        private void Application_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            try
            {
                // 获取点击的单元格
                Excel.Range clickedCell = Target.Cells[1, 1];

                // 检查单元格是否有条件格式
                if (HasConditionalFormat(clickedCell))
                {
                    // 手动刷新计算
                    // clickedCell.Application.Calculate();

                    // 获取选中区域的起始行、结束行、起始列和结束列
                    int startRow = Target.Row;
                    int endRow = Target.Cells[Target.Cells.Count].Row;
                    int startColumn = Target.Column;
                    int endColumn = Target.Cells[Target.Cells.Count].Column;

                    // 在 ActiveWorkbook 中添加名称
                    AddName("sRow", startRow);
                    AddName("eRow", endRow);
                    AddName("sColumn", startColumn);
                    AddName("eColumn", endColumn);
                }
            }
            catch (Exception ex)
            {
                // 处理异常，可以根据实际情况进行处理
                // 例如，记录日志或显示错误消息
            }
        }

        private bool HasConditionalFormat(Excel.Range cell)
        {
            // 获取单元格的条件格式
            Excel.FormatConditions conditions = cell.FormatConditions;
            
            if (conditions != null & conditions.Count > 0)
            {
                foreach (Excel.FormatCondition condition in conditions)
                {
                    // 检查是否有指定条件格式
                    if (condition.Type == 2 & condition.Formula1 == "=OR(CELL(\"row\")=ROW(),CELL(\"col\")=COLUMN())")
                    {
                        return true;
                    }
                }
            }           
            return false;
        }

        // 添加名称
        private void AddName(string name, int refersTo)
        {
            Excel.Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (activeWorkbook != null)
            {
                activeWorkbook.Names.Add(name, refersTo);
            }
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

        public void ToggleSpotlight()
        {
            spotlightEnable = !spotlightEnable;
            if (spotlightEnable)
            {
                // 启用时订阅（如果你的原来实现是这样）
                try { this.Application.SheetSelectionChange += Application_SheetSelectionChange; } catch { }
            }
            else
            {
                try { this.Application.SheetSelectionChange -= Application_SheetSelectionChange; } catch { }
            }
        }

        // 2) 提供给 Ribbon 显示标签的字符串
        public string GetSpotlightCaption()
        {
            return spotlightEnable ? "聚光灯：已开启" : "聚光灯：已关闭";
        }

        // 3) 提供给 Ribbon 获取当前状态（toggle 的 getPressed）
        public bool IsSpotlightEnabled()
        {
            return spotlightEnable;
        }

        // 4) 设置选中行高（被 ContextMenu 调用）
        public void SetSelectionRowHeight(double height)
        {
            try
            {
                var sel = this.Application.Selection as Microsoft.Office.Interop.Excel.Range;
                if (sel != null)
                    sel.Rows.RowHeight = height;
            }
            catch { /* 忽略 */ }
        }

        // 5) 你的自定义按钮回调（可选）
        public void OnCustomButtonClicked()
        {
            System.Windows.Forms.MessageBox.Show("自定义按钮被点击（来自 AggregatorRibbon）");
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new SingleResourceRibbon("ExcelAddIn_VSTO_Sample.RibbonMerged.xml");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // 退出时保存 聚光灯 开关状态（保留原有行为）
            Properties.Settings.Default.SpotlightEnabled = spotlightEnable;
            Properties.Settings.Default.Save();
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
