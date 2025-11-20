using ExcelAddIn_VSTO_Sample;
using Microsoft.Office.Core;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

[ComVisible(true)]
public class AggregatorRibbon : IRibbonExtensibility
{
    private readonly string[] resourceNames;
    private IRibbonUI ribbon;

    /// <summary>
    /// resourceNames: 嵌入资源名数组（按顺序合并），例如 {"Your.Namespace.DesignerRibbon.xml", "Your.Namespace.ContextMenus.xml"}
    /// </summary>
    public AggregatorRibbon(string[] resourceNames)
    {
        this.resourceNames = resourceNames ?? new string[0];
    }

    public string GetCustomUI(string ribbonID)
    {
        var sbChildren = new StringBuilder();
        var asm = Assembly.GetExecutingAssembly();

        foreach (var res in resourceNames)
        {
            try
            {
                using (var stream = asm.GetManifestResourceStream(res))
                {
                    if (stream == null)
                    {
                        sbChildren.AppendLine($"<!-- resource '{res}' not found -->");
                        continue;
                    }
                    using (var reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        var txt = reader.ReadToEnd();
                        // Extract children: if txt contains <customUI>..</customUI>, extract inner; else assume fragment
                        sbChildren.AppendLine(ExtractCustomUiChildren(txt));
                    }
                }
            }
            catch (Exception ex)
            {
                sbChildren.AppendLine($"<!-- failed load resource '{res}': {ex.Message} -->");
            }
        }

        var final = $@"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
                    {sbChildren}
                    </customUI>";
        return final;
    }

    private string ExtractCustomUiChildren(string xml)
    {
        if (string.IsNullOrWhiteSpace(xml)) return string.Empty;
        var trimmed = xml.Trim();
        var lower = trimmed.ToLowerInvariant();
        var startTag = "<customui";
        var endTag = "</customui>";
        var idxStart = lower.IndexOf(startTag, StringComparison.OrdinalIgnoreCase);
        var idxEnd = lower.LastIndexOf(endTag, StringComparison.OrdinalIgnoreCase);
        if (idxStart >= 0 && idxEnd > idxStart)
        {
            var gt = trimmed.IndexOf('>', idxStart);
            if (gt >= 0 && gt + 1 < trimmed.Length)
            {
                return trimmed.Substring(gt + 1, idxEnd - (gt + 1));
            }
        }
        // 如果不是完整 customUI，则直接返回（片段）
        return trimmed;
    }

    public void Ribbon_Load(IRibbonUI ribbonUI)
    {
        this.ribbon = ribbonUI;
    }

    // ---------- 回调：UI -> 委托到 ThisAddIn 的公用方法 ----------
    public void OnCustomButtonClick(IRibbonControl control)
    {
        try
        {
            // 你可以在 ThisAddIn 中实现更复杂逻辑
            Globals.ThisAddIn.OnCustomButtonClicked();
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show("OnCustomButtonClick 错误: " + ex.Message);
        }
    }

    // toggleButton 的 onAction 回调签名： (IRibbonControl control, bool pressed)
    public void OnToggleSpotlight(IRibbonControl control, bool pressed)
    {
        try
        {
            Globals.ThisAddIn.ToggleSpotlight();
            // 刷新 context menu labels
            this.ribbon?.InvalidateControl("btnToggleSpotlightCell");
            this.ribbon?.InvalidateControl("btnToggleSpotlightList");
            this.ribbon?.InvalidateControl("spotlightToggle");
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show("OnToggleSpotlight 错误: " + ex.Message);
        }
    }

    // 当从 ContextMenu 点击切换聚光灯时（signature without pressed）
    public void OnToggleSpotlightFromContext(IRibbonControl control)
    {
        try
        {
            Globals.ThisAddIn.ToggleSpotlight();
            this.ribbon?.InvalidateControl("btnToggleSpotlightCell");
            this.ribbon?.InvalidateControl("btnToggleSpotlightList");
            this.ribbon?.InvalidateControl("spotlightToggle");
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show("OnToggleSpotlightFromContext 错误: " + ex.Message);
        }
    }

    // getLabel 回调
    public string GetSpotlightLabel(IRibbonControl control)
    {
        try
        {
            return Globals.ThisAddIn.GetSpotlightCaption();
        }
        catch
        {
            return "聚光灯";
        }
    }

    // toggle 的 getPressed 回调
    public bool GetSpotlightPressed(IRibbonControl control)
    {
        try
        {
            return Globals.ThisAddIn.IsSpotlightEnabled();
        }
        catch
        {
            return false;
        }
    }

    // Context menu 设置行高
    public void OnSetRowHeight(IRibbonControl control)
    {
        try
        {
            Globals.ThisAddIn.SetSelectionRowHeight(25);
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show("OnSetRowHeight 错误: " + ex.Message);
        }
    }
}
