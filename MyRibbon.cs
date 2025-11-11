using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelAddIn_VSTO_Sample
{
    public partial class MyRibbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            spotlightCheckBox.Checked = Properties.Settings.Default.SpotlightEnabled;
        }

        private void spotlightCheckBox_Click(object sender, RibbonControlEventArgs e)
        {
            // 读取勾选状态并保存
            Properties.Settings.Default.SpotlightEnabled = spotlightCheckBox.Checked;
            Properties.Settings.Default.Save();

            // 你的业务逻辑
            Globals.ThisAddIn.ToggleSpotlight();
        }
    }
}
