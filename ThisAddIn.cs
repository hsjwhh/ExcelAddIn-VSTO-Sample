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
        private bool spotlightEnable;
        private MyUserControl myUserControl1;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        // 新增字段 spotlightButton、spotlightListButton、rowButton 用于存储右键菜单按钮的引用及绑定状态
        private Office.CommandBarButton spotlightButton;
        private Office.CommandBarButton spotlightListButton;
        private bool spotlightHandlerAttached = false;
        private bool spotlightListHandlerAttached = false;
        private Office.CommandBarButton rowButton;
        private bool rowHandlerAttached = false;
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
            this.Application.SheetBeforeRightClick += Application_SheetBeforeRightClick;

        }

        // 添加自定义菜单项目
        private Office.CommandBarButton AddContextMenuItem(
            string menuName,
            string tag,
            string caption,
            bool handlerattached,
            Office._CommandBarButtonEvents_ClickEventHandler handler)
        {
            Office.CommandBar menu = null;
            try
            {
                menu = Globals.ThisAddIn.Application.CommandBars[menuName];
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Menu '{menuName}' not found: {ex.Message}");
                return null;
            }

            if (menu == null) return null;

            Office.CommandBarButton button = null;

            // 查找是否已存在
            foreach (Office.CommandBarControl ctrl in menu.Controls)
            {
                try
                {
                    if (ctrl != null && ctrl.Tag == tag)
                    {
                        button = ctrl as Office.CommandBarButton;
                        break;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error checking control tag: {ex.Message}");
                }
            }

            // 如果不存在则创建
            if (button == null)
            {
                var created = menu.Controls.Add(
                    MsoControlType.msoControlButton,
                    Type.Missing,
                    Type.Missing,
                    menu.Controls.Count + 1,
                    true);

                button = (Office.CommandBarButton)created;
                button.Tag = tag;
                button.Caption = caption;
            }
            else
            {
                // 已存在时可选择更新 Caption
                if (!string.IsNullOrEmpty(caption))
                    button.Caption = caption;
            }

            // 绑定事件（避免重复绑定）
            try
            {
                if (!handlerattached)
                {
                    button.Click -= handler;
                    button.Click += handler;
                    if (tag == "ToggleSpotlight") spotlightHandlerAttached = true;
                    if (tag == "ToggleListSpotlight") spotlightListHandlerAttached = true;
                    if (tag == "SetRowHeight") rowHandlerAttached = true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error binding event for {tag}: {ex.Message}");
            }

            // 保存引用（可选）
            if (tag == "ToggleSpotlight") spotlightButton = button;
            if (tag == "ToggleListSpotlight") spotlightListButton = button;
            if (tag == "SetRowHeight") rowButton = button;

            return button;
        }

        // 移除自定义菜单项
        private void RemoveContextMenuItem(string menuName, string tag)
        {
            Office.CommandBar menu = null;
            try
            {
                menu = Globals.ThisAddIn.Application.CommandBars[menuName];
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Menu '{menuName}' not found: {ex.Message}");
                return;
            }

            if (menu == null) return;

            try
            {
                for (int i = 1; i <= menu.Controls.Count; i++)
                {
                    var ctrl = menu.Controls[i];
                    if (ctrl != null && ctrl.Tag == tag)
                    {
                        ctrl.Delete();
                        System.Diagnostics.Debug.WriteLine($"Removed context menu item '{tag}' from '{menuName}'");
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error removing item '{tag}' from '{menuName}': {ex.Message}");
            }
        }


        #region 添加右键菜单项目
        private void Application_SheetBeforeRightClick(object Sh, Excel.Range Target, ref bool Cancel)
        {
            try
            {
                // --- 单元格菜单：聚光灯开关（查找已存在控件或创建，并确保事件只绑定一次） ---
                string caption = spotlightEnable ? "聚光灯：已开启" : "聚光灯：已关闭";
                if (Target.Columns.Count == Target.Worksheet.Columns.Count)
                {
                    // 调整行高，所以右键所在行的菜单添加自定义项目
                    // 说明选中的是整行（点击了行号区域）
                    AddContextMenuItem("Row", "SetRowHeight", "设置行高为 25", rowHandlerAttached, SetRowHeightMenuItemClick);
                }
                else if (Target.ListObject != null)
                {
                    // 表格中的单元格
                    AddContextMenuItem("List Range Popup", "ToggleListSpotlight", caption, spotlightListHandlerAttached, ToggleSpotlightMenuItemClick);
                }
                else
                {
                    // 普通的单元格
                    AddContextMenuItem("Cell", "ToggleSpotlight", caption, spotlightHandlerAttached, ToggleSpotlightMenuItemClick);
                }
            }
            catch
            {
                // 忽略右键菜单修改失败，避免影响 Excel
            }
         
        }

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
                this.Application.SheetSelectionChange += Application_SheetSelectionChange;
            }
            else
            { 
                this.Application.SheetSelectionChange -= Application_SheetSelectionChange;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // 在关闭时做清理
            RemoveContextMenuItem("Row", "SetRowHeight");
            RemoveContextMenuItem("List Range Popup", "ToggleListSpotlight");
            RemoveContextMenuItem("Cell", "ToggleSpotlight");

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
