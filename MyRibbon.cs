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
            spotlightToggleButton.Checked = Properties.Settings.Default.SpotlightEnabled;
        }

        private void spotlightToggleButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ToggleSpotlight();
            // 同步按钮状态
            spotlightToggleButton.Checked = Properties.Settings.Default.SpotlightEnabled;
        }
    }
}
