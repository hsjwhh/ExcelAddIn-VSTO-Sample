using ExcelAddIn_VSTO_Sample;
using Microsoft.Office.Core;                // IRibbonExtensibility, IRibbonUI, IRibbonControl 等类型
using System;                               // 基本框架类型（Exception 等）
using System.IO;                            // 处理 Stream/StreamReader
using System.Reflection;                    // 读取嵌入资源需要使用 Assembly
using System.Runtime.InteropServices;       // 用于 [ComVisible]

// [ComVisible(true)]：让此类对 COM 可见，Office 通过 COM 调用你的回调时需要这个。
[ComVisible(true)]
public class SingleResourceRibbon : IRibbonExtensibility
{
    // 保存要读取的嵌入资源名称（例如 "ExcelAddIn_VSTO_Sample.RibbonMerged.xml"）
    private readonly string _resourceName;

    // 在 Ribbon_Load 回调里会得到 IRibbonUI 的实例，保存起来用于后续 InvalidateControl。
    private IRibbonUI _ribbon;

    // 构造器：传入嵌入资源完整名（含命名空间前缀）
    public SingleResourceRibbon(string resourceName)
    {
        _resourceName = resourceName;
    }

    // IRibbonExtensibility 必须实现的方法：返回给 Office 的 Ribbon XML 字符串
    public string GetCustomUI(string ribbonID)
    {
        try
        {
            // 获取当前程序集（DLL），我们把 Ribbon XML 做成了嵌入资源打进这个程序集
            var asm = Assembly.GetExecutingAssembly();

            // 通过资源名打开流（如果找不到 resource 会返回 null）
            using (Stream s = asm.GetManifestResourceStream(_resourceName))
            {
                if (s == null)
                {
                    // 若资源找不到，弹窗提示（便于调试），并返回一个空的 customUI，避免 Excel 抛更严重错误
                    System.Windows.Forms.MessageBox.Show(
                        "无法读取资源：" + _resourceName,
                        "Ribbon XML 读取失败");

                    return "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'></customUI>";
                }

                // 读取流内容并作为字符串返回（即完整的 customUI XML）
                using (StreamReader r = new StreamReader(s))
                    return r.ReadToEnd();
            }
        }
        catch (Exception ex)
        {
            // 任何异常都显示给开发者并返回空 customUI，避免 Excel 加载失败
            System.Windows.Forms.MessageBox.Show("GetCustomUI: " + ex.Message);
            return "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'></customUI>";
        }
    }

    // Office 会在加载 Ribbon 时调用 Ribbon_Load，把 IRibbonUI 实例传进来
    // 保存下来后就可以在回调里调用 _ribbon.InvalidateControl(...) 刷新控件状态/label
    public void Ribbon_Load(IRibbonUI ribbonUI)
    {
        _ribbon = ribbonUI;
    }

    // ---------- 以下为 XML 中声明的回调方法（public 且签名要和 XML 指定的回调类型一致） ----------

    // Ribbon Tab 中按钮的 onAction 回调（button 的 onAction 只有一个 IRibbonControl 参数）
    public void OnCustomButtonClick(IRibbonControl control)
        => Globals.ThisAddIn.OnCustomButtonClicked();

    // getVisible 回调（用于控制 btClose 是否显示）
    public bool GetBtCloseVisible(IRibbonControl control) => false;

    // getVisible 回调（控制 spotlight toggle 是否可见）
    public bool GetSpotlightVisible(IRibbonControl control) => true;

    // toggleButton 的 onAction（注意签名：(IRibbonControl, bool pressed)）
    // 当用户通过 Ribbon 上的 toggle 切换时调用
    public void OnToggleSpotlight(IRibbonControl control, bool pressed)
    {
        Globals.ThisAddIn.ToggleSpotlight();
        InvalidateSpotControls();
    }

    // toggleButton 的 getPressed 回调，用来让 Ribbon 显示 toggle 当前的按下状态
    public bool GetSpotlightPressed(IRibbonControl control)
        => Globals.ThisAddIn.IsSpotlightEnabled();

    // Context menu 的 button 的 onAction（单参数签名）
    public void OnToggleSpotlightFromContext(IRibbonControl control)
    {
        Globals.ThisAddIn.ToggleSpotlight();
        InvalidateSpotControls();
    }

    // getLabel 回调：ContextMenu 上显示动态文字（例如 “聚光灯：已开启/已关闭”）
    public string GetSpotlightLabel(IRibbonControl control)
        => Globals.ThisAddIn.GetSpotlightCaption();

    // Context menu 设置行高的回调
    public void OnSetRowHeight(IRibbonControl control)
        => Globals.ThisAddIn.SetSelectionRowHeight(25);

    // 工具方法：使与聚光灯相关的几个控件刷新显示（label/pressed 状态）
    private void InvalidateSpotControls()
    {
        _ribbon?.InvalidateControl("spotlightToggle");
        _ribbon?.InvalidateControl("btnToggleSpotlightCell");
        _ribbon?.InvalidateControl("btnToggleSpotlightList");
    }
}