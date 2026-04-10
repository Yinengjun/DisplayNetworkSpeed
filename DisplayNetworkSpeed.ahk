#Requires AutoHotkey v2.0
#SingleInstance Force
ListLines(false)
DetectHiddenWindows(true)
SetControlDelay(-1)

A_TrayMenu.Delete()
A_TrayMenu.Add("设置", ShowSettings)
A_TrayMenu.Add("关于", ShowAbout)
A_TrayMenu.Add()
A_TrayMenu.Add("重启程序", RestartApp)
A_TrayMenu.Add("退出程序", ExitAppHandler)

ConfigFile := A_ScriptDir "\config.ini"

DefaultInterval := 1000
DefaultGuiWidth := 120
DefaultGuiHeight := 44
DefaultFontName := "Segoe UI Variable"
DefaultFontSize := 11
DefaultFontWeight := "Bold"
DefaultBgColorMode := "深色预设"
DefaultBgColor := "0x0B1113"
DefaultOnlyText := false
DefaultBgTransparency := 255
DefaultNumRightMargin := 6
DefaultArrowWidth := 24
DefaultPositionCorner := "右下角"
DefaultOffsetX := 15
DefaultOffsetY := 10
DefaultLimitOffset := true
DefaultThresh1 := 50 * 1024
DefaultThresh2 := 500 * 1024
DefaultThresh3 := 2 * 1024 * 1024
DefaultColorVeryLow := "CFCFCF"
DefaultColorLow := "A8D5A2"
DefaultColorMed := "7FD3D6"
DefaultColorHigh := "F2C08C"
DefaultEnableSmoothing := true
DefaultEMAFactor := "0.35"
DefaultConfirmNeeded := 2
DefaultStatsScope := "全部网卡"
DefaultStatsAdapters := ""
DefaultDataSource := "WMI"
DefaultAutoRestart := false
DefaultMouseThrough := true
DefaultDisplayTarget := "主屏幕"
DefaultDragPositioning := false
DefaultAutoStart := false
DefaultAutoStartScope := "当前用户"
DefaultShowLineMarkers := true
DefaultLineMarkerStyle := "↑ ↓"
DefaultColorChangeWithValue := false

LoadConfig()

NumWidth := GuiWidth - ArrowWidth - NumRightMargin
UpY := 4
DownY := 22
ArrowX := NumWidth

global UpNum, UpArrow, DownNum, DownArrow
global emaUp := 0, emaDown := 0
global pendingUp := "", pendingDown := ""
global pendingCountUp := 0, pendingCountDown := 0
global lastColorUp := "", lastColorDown := ""
global lastDisplayColorUp := "", lastDisplayColorDown := ""
global lastTextUp := "", lastTextDown := ""
global recv := 0, sent := 0
global q, item, sSent, sRecv, candidateUp, candidateDown
global Display
global hGui
global LimitOffset
global StatsScope, SelectedAdapters
global AdapterList := ""

global DragPositioning
global DragStartX, DragStartY
global DragStartGuiX, DragStartGuiY
global IsDragging := false

global PickerActive := false
global CurrentPickTarget := ""
global CurrentPickFmt := ""
global hPicker := 0
global PickerLastRGB := "FFFFFF"
global PickerLastBGR := "FFFFFF"
global PickPrev

global SettingsGui := 0
global AdapterSelectGui := 0
global PickerGui := 0
global AboutGui := 0
global MainGui := 0

global wmi, WmiWarned
global DataSource
global IpPrevTick := 0, IpPrevMap := Map()
global IpHelperAvailable := true, IpHelperWarned := false
WmiWarned := false
wmi := ""
InitDataSource()

CreateGuiAndShow(ColorVeryLow)
SetTimer(UpdateNet, Interval)
UpdateNet()

LoadConfig() {
    global

    if (!FileExist(ConfigFile))
        CreateDefaultConfig()

    Interval := IniRead(ConfigFile, "General", "Interval", DefaultInterval)
    AutoRestart := IniRead(ConfigFile, "General", "AutoRestart", DefaultAutoRestart)
    MouseThrough := IniRead(ConfigFile, "Settings", "MouseThrough", DefaultMouseThrough)
    GuiWidth := IniRead(ConfigFile, "GUI", "Width", DefaultGuiWidth)
    GuiHeight := IniRead(ConfigFile, "GUI", "Height", DefaultGuiHeight)
    FontName := IniRead(ConfigFile, "GUI", "FontName", DefaultFontName)
    FontSize := IniRead(ConfigFile, "GUI", "FontSize", DefaultFontSize)
    FontWeight := IniRead(ConfigFile, "GUI", "FontWeight", DefaultFontWeight)
    BgColorMode := IniRead(ConfigFile, "GUI", "BgColorMode", DefaultBgColorMode)
    BgColor := IniRead(ConfigFile, "GUI", "BgColor", DefaultBgColor)
    OnlyText := IniRead(ConfigFile, "GUI", "OnlyText", DefaultOnlyText)
    BgTransparency := IniRead(ConfigFile, "GUI", "BgTransparency", DefaultBgTransparency)
    NumRightMargin := IniRead(ConfigFile, "GUI", "NumRightMargin", DefaultNumRightMargin)
    ArrowWidth := IniRead(ConfigFile, "GUI", "ArrowWidth", DefaultArrowWidth)
    PositionCorner := IniRead(ConfigFile, "Position", "Corner", DefaultPositionCorner)
    OffsetX := IniRead(ConfigFile, "Position", "OffsetX", DefaultOffsetX)
    OffsetY := IniRead(ConfigFile, "Position", "OffsetY", DefaultOffsetY)
    Display := IniRead(ConfigFile, "GUI", "Display", DefaultDisplayTarget)
    LimitOffset := IniRead(ConfigFile, "Position", "LimitOffset", DefaultLimitOffset)

    Thresh1 := IniRead(ConfigFile, "Thresholds", "Thresh1", DefaultThresh1)
    Thresh2 := IniRead(ConfigFile, "Thresholds", "Thresh2", DefaultThresh2)
    Thresh3 := IniRead(ConfigFile, "Thresholds", "Thresh3", DefaultThresh3)

    ColorVeryLow := IniRead(ConfigFile, "Colors", "VeryLow", DefaultColorVeryLow)
    ColorLow := IniRead(ConfigFile, "Colors", "Low", DefaultColorLow)
    ColorMed := IniRead(ConfigFile, "Colors", "Medium", DefaultColorMed)
    ColorHigh := IniRead(ConfigFile, "Colors", "High", DefaultColorHigh)

    EnableSmoothing := IniRead(ConfigFile, "Advanced", "EnableSmoothing", DefaultEnableSmoothing)
    EMAFactor := IniRead(ConfigFile, "Advanced", "EMAFactor", DefaultEMAFactor)
    ConfirmNeeded := IniRead(ConfigFile, "Advanced", "ConfirmNeeded", DefaultConfirmNeeded)
    StatsScope := IniRead(ConfigFile, "Advanced", "StatsScope", DefaultStatsScope)
    SelectedAdapters := IniRead(ConfigFile, "Advanced", "StatsAdapters", DefaultStatsAdapters)
    DataSource := IniRead(ConfigFile, "Advanced", "DataSource", DefaultDataSource)

    DragPositioning := IniRead(ConfigFile, "Position", "DragPositioning", DefaultDragPositioning)

    AutoStart := IniRead(ConfigFile, "General", "AutoStart", DefaultAutoStart)
    AutoStartScope := IniRead(ConfigFile, "General", "AutoStartScope", DefaultAutoStartScope)

    ShowLineMarkers := IniRead(ConfigFile, "GUI", "ShowLineMarkers", DefaultShowLineMarkers)
    LineMarkerStyle := IniRead(ConfigFile, "GUI", "LineMarkerStyle", DefaultLineMarkerStyle)
    ColorChangeWithValue := IniRead(ConfigFile, "GUI", "ColorChangeWithValue", DefaultColorChangeWithValue)
    CombinedMode := ColorChangeWithValue && ShowLineMarkers

    Interval := RegExReplace(Interval, "[,\s]", "")
    GuiWidth := RegExReplace(GuiWidth, "[,\s]", "")
    GuiHeight := RegExReplace(GuiHeight, "[,\s]", "")
    FontSize := RegExReplace(FontSize, "[,\s]", "")
    BgTransparency := RegExReplace(BgTransparency, "[,\s]", "")
    NumRightMargin := RegExReplace(NumRightMargin, "[,\s]", "")
    ArrowWidth := RegExReplace(ArrowWidth, "[,\s]", "")
    OffsetX := RegExReplace(OffsetX, "[,\s]", "")
    OffsetY := RegExReplace(OffsetY, "[,\s]", "")
    Thresh1 := RegExReplace(Thresh1, "[,\s]", "")
    Thresh2 := RegExReplace(Thresh2, "[,\s]", "")
    Thresh3 := RegExReplace(Thresh3, "[,\s]", "")
    ConfirmNeeded := RegExReplace(ConfirmNeeded, "[,\s]", "")

    if (Interval < 100 || Interval > 5000)
        Interval := DefaultInterval
    if (GuiWidth < 80 || GuiWidth > 300)
        GuiWidth := DefaultGuiWidth
    if (GuiHeight < 30 || GuiHeight > 100)
        GuiHeight := DefaultGuiHeight
    if (FontSize < 8 || FontSize > 24)
        FontSize := DefaultFontSize
    if (BgTransparency < 0 || BgTransparency > 255)
        BgTransparency := DefaultBgTransparency
    if (BgColorMode = "浅色预设")
        BgColor := "0xF5F5F5"
    else if (BgColorMode = "深色预设")
        BgColor := "0x0B1113"

    if (StatsScope != "自动" && StatsScope != "全部网卡" && StatsScope != "自定义")
        StatsScope := DefaultStatsScope

    DataSource := NormalizeDataSource(DataSource)

    if (!EnableSmoothing)
        EMAFactor := 1.0
}

CreateDefaultConfig() {
    global

    IniWrite(DefaultInterval, ConfigFile, "General", "Interval")
    IniWrite(DefaultAutoRestart, ConfigFile, "General", "AutoRestart")
    IniWrite(DefaultMouseThrough, ConfigFile, "Settings", "MouseThrough")
    IniWrite(DefaultGuiWidth, ConfigFile, "GUI", "Width")
    IniWrite(DefaultGuiHeight, ConfigFile, "GUI", "Height")
    IniWrite(DefaultFontName, ConfigFile, "GUI", "FontName")
    IniWrite(DefaultFontSize, ConfigFile, "GUI", "FontSize")
    IniWrite(DefaultFontWeight, ConfigFile, "GUI", "FontWeight")
    IniWrite(DefaultBgColorMode, ConfigFile, "GUI", "BgColorMode")
    IniWrite(DefaultBgColor, ConfigFile, "GUI", "BgColor")
    IniWrite(DefaultOnlyText, ConfigFile, "GUI", "OnlyText")
    IniWrite(DefaultBgTransparency, ConfigFile, "GUI", "BgTransparency")
    IniWrite(DefaultShowLineMarkers, ConfigFile, "GUI", "ShowLineMarkers")
    IniWrite(DefaultLineMarkerStyle, ConfigFile, "GUI", "LineMarkerStyle")
    IniWrite(DefaultColorChangeWithValue, ConfigFile, "GUI", "ColorChangeWithValue")
    IniWrite(DefaultNumRightMargin, ConfigFile, "GUI", "NumRightMargin")
    IniWrite(DefaultArrowWidth, ConfigFile, "GUI", "ArrowWidth")
    IniWrite(DefaultPositionCorner, ConfigFile, "Position", "Corner")
    IniWrite(DefaultOffsetX, ConfigFile, "Position", "OffsetX")
    IniWrite(DefaultOffsetY, ConfigFile, "Position", "OffsetY")
    IniWrite(DefaultDisplayTarget, ConfigFile, "GUI", "Display")
    IniWrite(DefaultLimitOffset, ConfigFile, "Position", "LimitOffset")

    IniWrite(DefaultThresh1, ConfigFile, "Thresholds", "Thresh1")
    IniWrite(DefaultThresh2, ConfigFile, "Thresholds", "Thresh2")
    IniWrite(DefaultThresh3, ConfigFile, "Thresholds", "Thresh3")

    IniWrite(DefaultColorVeryLow, ConfigFile, "Colors", "VeryLow")
    IniWrite(DefaultColorLow, ConfigFile, "Colors", "Low")
    IniWrite(DefaultColorMed, ConfigFile, "Colors", "Medium")
    IniWrite(DefaultColorHigh, ConfigFile, "Colors", "High")

    IniWrite(DefaultEnableSmoothing, ConfigFile, "Advanced", "EnableSmoothing")
    IniWrite(DefaultEMAFactor, ConfigFile, "Advanced", "EMAFactor")
    IniWrite(DefaultConfirmNeeded, ConfigFile, "Advanced", "ConfirmNeeded")
    IniWrite(DefaultStatsScope, ConfigFile, "Advanced", "StatsScope")
    IniWrite(DefaultStatsAdapters, ConfigFile, "Advanced", "StatsAdapters")
    IniWrite(DefaultDataSource, ConfigFile, "Advanced", "DataSource")

    IniWrite(DefaultDragPositioning, ConfigFile, "Position", "DragPositioning")

    IniWrite(DefaultAutoStart, ConfigFile, "General", "AutoStart")
    IniWrite(DefaultAutoStartScope, ConfigFile, "General", "AutoStartScope")
}

ShowSettings(*) {
    global SettingsGui, FontName, FontSize, FontWeight, BgColorMode, BgColor
    global GuiWidth, GuiHeight, Interval, AutoRestart, MouseThrough
    global DragPositioning
    global PositionCorner, LimitOffset, OffsetX, OffsetY
    global Display, Thresh1, Thresh2, Thresh3
    global ColorVeryLow, ColorLow, ColorMed, ColorHigh
    global EnableSmoothing, EMAFactor, ConfirmNeeded
    global StatsScope, AutoStart, AutoStartScope
    global ShowLineMarkers, LineMarkerStyle, ColorChangeWithValue
    global DataSource

    if (SettingsGui)
        SettingsGui.Destroy()

    SettingsGui := Gui("+Owner")
    SettingsGui.OnEvent("Close", SettingsGuiClose)

    tab := SettingsGui.Add("Tab3", "", ["常规", "界面", "行标", "位置", "网速阈值", "颜色", "高级"])

    tab.UseTab("常规")
    SettingsGui.Add("Text", "x20 y50", "刷新间隔 (毫秒):")
    IntervalClean := RegExReplace(Interval, "[^\d-]", "")
    if (IntervalClean = "")
        IntervalClean := 1000
    SettingsGui.Add("Edit", "x140 y46 w80 vInterval Number", IntervalClean)
    SettingsGui.Add("UpDown", "vIntervalUD Range100-5000 0x80", IntervalClean)

    SettingsGui.Add("Checkbox", "x20 y80 vAutoRestart", "保存后重启不二次确认").Value := AutoRestart

    mouseCtrl := SettingsGui.Add("Checkbox", "x20 y110 vMouseThrough", "鼠标穿透")
    mouseCtrl.Value := MouseThrough
    mouseCtrl.OnEvent("Click", MouseThroughChanged)

    tab.UseTab("界面")
    SettingsGui.Add("Text", "x20 y50", "窗口宽度:")
    SettingsGui.Add("Edit", "x100 y46 w50 vGuiWidth Number", GuiWidth)
    SettingsGui.Add("UpDown", "vGuiWidthUD Range80-300", GuiWidth)
    SettingsGui.Add("Text", "x180 y50", "窗口高度:")
    SettingsGui.Add("Edit", "x250 y46 w50 vGuiHeight Number", GuiHeight)
    SettingsGui.Add("UpDown", "vGuiHeightUD Range30-100", GuiHeight)

    SettingsGui.Add("Text", "x20 y80", "字体名称:")
    fontPresetList := ["Segoe UI Variable", "Segoe UI", "Microsoft YaHei", "Consolas", "Cascadia Mono", "Cascadia Code", "Sarasa Mono SC", "SimHei", "SimSun", "Arial", "Times New Roman", "自定义"]
    fontPresetCtrl := SettingsGui.Add("DropDownList", "x100 y76 w160 vFontNamePreset", fontPresetList)
    fontPresetCtrl.OnEvent("Change", FontNamePresetChange)
    SettingsGui.Add("Text", "x270 y80 vFontNameCustomLabel", "自定义:")
    SettingsGui.Add("Edit", "x320 y76 w100 vFontNameCustom", FontName)

    if (HasValue(fontPresetList, FontName)) {
        fontPresetCtrl.Choose(FontName)
        SettingsGui["FontNameCustom"].Visible := false
        SettingsGui["FontNameCustomLabel"].Visible := false
    } else {
        fontPresetCtrl.Choose("自定义")
        SettingsGui["FontNameCustom"].Visible := true
        SettingsGui["FontNameCustomLabel"].Visible := true
    }

    SettingsGui.Add("Text", "x20 y110", "字号:")
    SettingsGui.Add("Edit", "x60 y106 w40 vFontSize", FontSize)
    SettingsGui.Add("UpDown", "vFontSizeUD Range8-24", FontSize)

    SettingsGui.Add("Text", "x120 y110", "字体粗细:")
    fontWeightCtrl := SettingsGui.Add("DropDownList", "x180 y106 w80 vFontWeight", ["Normal", "Bold"])
    fontWeightCtrl.Choose(FontWeight = "Bold" ? 2 : 1)

    SettingsGui.Add("Text", "x20 y140", "背景色模式:")
    bgModeCtrl := SettingsGui.Add("DropDownList", "x90 y136 w80 vBgColorMode", ["浅色预设", "深色预设", "自定义"])
    bgModeCtrl.Choose(BgColorMode = "浅色预设" ? 1 : (BgColorMode = "深色预设" ? 2 : 3))
    bgModeCtrl.OnEvent("Change", BgColorModeChange)

    SettingsGui.Add("Text", "x180 y140", "自定义背景色:")
    bgColorCtrl := SettingsGui.Add("Edit", "x260 y136 w80 vBgColor", BgColor)
    bgColorCtrl.OnEvent("Change", BgColorChanged)
    SettingsGui.Add("Progress", "x350 y136 w30 h20 vPrevBgColor +Border", 100)
    SettingsGui.Add("Button", "x390 y136 w45 h20 vPickBgBtn", "取色").OnEvent("Click", PickBgColor)

    if (BgColorMode != "自定义") {
        SettingsGui["BgColor"].Enabled := false
        SettingsGui["PrevBgColor"].Enabled := false
        SettingsGui["PickBgBtn"].Enabled := false
    }

    onlyTextCtrl := SettingsGui.Add("Checkbox", "x20 y170 vOnlyText", "只显示文字")
    onlyTextCtrl.Value := OnlyText
    onlyTextCtrl.OnEvent("Click", OnlyTextChanged)

    SettingsGui.Add("Text", "x20 y200", "背景透明度 (0-255):")
    SettingsGui.Add("Edit", "x150 y196 w50 vBgTransparency Number", BgTransparency)
    SettingsGui.Add("UpDown", "vBgTransparencyUD Range0-255", BgTransparency)
    SettingsGui.Add("Text", "x210 y200", "(0=完全透明，255=完全不透明)")

    mCount := MonitorGetCount()
    dispOpt := ["主屏幕", "全部"]
    Loop mCount
        dispOpt.Push("显示器" A_Index)
    SettingsGui.Add("Text", "x20 y230", "显示器:")
    dispCtrl := SettingsGui.Add("DropDownList", "x80 y226 w140 vDisplay", dispOpt)
    if (Display = "")
        Display := "主屏幕"
    dispCtrl.Choose(Display)

    tab.UseTab("行标")
    showLineCtrl := SettingsGui.Add("Checkbox", "x20 y50 vShowLineMarkers", "显示上下行标志")
    showLineCtrl.Value := ShowLineMarkers
    showLineCtrl.OnEvent("Click", ShowLineMarkersChanged)

    colorChangeCtrl := SettingsGui.Add("Checkbox", "x20 y80 vColorChangeWithValue", "标志随数值变色")
    colorChangeCtrl.Value := ColorChangeWithValue
    colorChangeCtrl.OnEvent("Click", ColorChangeWithValueChanged)

    SettingsGui.Add("Text", "x20 y110", "上下行标志样式:")
    lineMarkerStyles := ["↑ ↓", "TX RX", "⬆️⬇️", "🔼🔽", "⏫⏬", "上传 下载", "上行 下行", "发送 接收"]
    lineMarkerCtrl := SettingsGui.Add("DropDownList", "x140 y106 w120 vLineMarkerStyle", lineMarkerStyles)
    lineMarkerCtrl.Choose(LineMarkerStyle)
    lineMarkerCtrl.OnEvent("Change", LineMarkerStyleChange)

    if (!ShowLineMarkers) {
        SettingsGui["LineMarkerStyle"].Enabled := false
    }

    tab.UseTab("位置")
    SettingsGui.Add("Text", "x20 y50", "位置角落:")
    posCornerCtrl := SettingsGui.Add("DropDownList", "x100 y46 w100 vPositionCorner", ["右下角", "右上角", "左下角", "左上角"])
    posCornerCtrl.Choose(PositionCorner = "右下角" ? 1 : (PositionCorner = "右上角" ? 2 : (PositionCorner = "左下角" ? 3 : 4)))
    SettingsGui.Add("Checkbox", "x210 y48 vLimitOffset", "限制偏离量（防止超出屏幕）").Value := LimitOffset

    SettingsGui.Add("Text", "x20 y80", "横向偏移:")
    SettingsGui.Add("Edit", "x100 y76 w60 vOffsetX", OffsetX).OnEvent("Change", SignedIntEditFilter)
    SettingsGui.Add("UpDown", "vOffsetXUD Range-1000-1000", OffsetX)
    SettingsGui.Add("Text", "x170 y80", "(正数向右，负数向左)")

    SettingsGui.Add("Text", "x20 y110", "纵向偏移:")
    SettingsGui.Add("Edit", "x100 y106 w60 vOffsetY", OffsetY).OnEvent("Change", SignedIntEditFilter)
    SettingsGui.Add("UpDown", "vOffsetYUD Range-200-200", OffsetY)
    SettingsGui.Add("Text", "x170 y110", "(正数向上，负数向下)")

    dragCtrl := SettingsGui.Add("Checkbox", "x20 y140 vDragPositioning", "启用拖动定位")
    dragCtrl.Value := DragPositioning
    dragCtrl.OnEvent("Click", DragPositioningChanged)
    SettingsGui.Add("Text", "x20 y170 w350", "(启用后可拖动窗口调整位置，偏移量将自动计算保存)")

    if (DragPositioning) {
        SettingsGui["PositionCorner"].Enabled := false
        SettingsGui["OffsetX"].Enabled := false
        SettingsGui["OffsetXUD"].Enabled := false
        SettingsGui["OffsetY"].Enabled := false
        SettingsGui["OffsetYUD"].Enabled := false
        SettingsGui["LimitOffset"].Enabled := false
    }

    tab.UseTab("网速阈值")
    SettingsGui.Add("Text", "x20 y50", "很低速阈值 (KB/s):")
    SettingsGui.Add("Edit", "x150 y46 w60 vThresh1KB Number", Round(Thresh1 / 1024))
    SettingsGui.Add("UpDown", "vThresh1KBUD Range1-1000", Round(Thresh1 / 1024))

    SettingsGui.Add("Text", "x20 y80", "低速阈值 (KB/s):")
    SettingsGui.Add("Edit", "x150 y76 w60 vThresh2KB Number", Round(Thresh2 / 1024))
    SettingsGui.Add("UpDown", "vThresh2KBUD Range1-5000", Round(Thresh2 / 1024))

    SettingsGui.Add("Text", "x20 y110", "中速阈值 (MB/s):")
    SettingsGui.Add("Edit", "x150 y106 w60 vThresh3MB Number", Round(Thresh3 / 1024 / 1024))
    SettingsGui.Add("UpDown", "vThresh3MBUD Range1-100", Round(Thresh3 / 1024 / 1024))

    tab.UseTab("颜色")
    SettingsGui.Add("Text", "x20 y50", "很低速颜色:")
    SettingsGui.Add("Edit", "x120 y46 w80 vColorVeryLow", ColorVeryLow).OnEvent("Change", ColorEditChanged)
    SettingsGui.Add("Progress", "x205 y46 w30 h20 vPrevVeryLow +Border", 100)
    SettingsGui.Add("Button", "x240 y44 w45 h22", "取色").OnEvent("Click", PickVeryLow)

    SettingsGui.Add("Text", "x20 y80", "低速颜色:")
    SettingsGui.Add("Edit", "x120 y76 w80 vColorLow", ColorLow).OnEvent("Change", ColorEditChanged)
    SettingsGui.Add("Progress", "x205 y76 w30 h20 vPrevLow +Border", 100)
    SettingsGui.Add("Button", "x240 y74 w45 h22", "取色").OnEvent("Click", PickLow)

    SettingsGui.Add("Text", "x20 y110", "中速颜色:")
    SettingsGui.Add("Edit", "x120 y106 w80 vColorMed", ColorMed).OnEvent("Change", ColorEditChanged)
    SettingsGui.Add("Progress", "x205 y106 w30 h20 vPrevMed +Border", 100)
    SettingsGui.Add("Button", "x240 y104 w45 h22", "取色").OnEvent("Click", PickMed)

    SettingsGui.Add("Text", "x20 y140", "高速颜色:")
    SettingsGui.Add("Edit", "x120 y136 w80 vColorHigh", ColorHigh).OnEvent("Change", ColorEditChanged)
    SettingsGui.Add("Progress", "x205 y136 w30 h20 vPrevHigh +Border", 100)
    SettingsGui.Add("Button", "x240 y134 w45 h22", "取色").OnEvent("Click", PickHigh)

    tab.UseTab("高级")
    SettingsGui.Add("Text", "x20 y50", "数据来源:")
    dataSourceCtrl := SettingsGui.Add("DropDownList", "x90 y46 w160 vDataSource", ["WMI (默认)", "IP Helper API"])
    dataSourceCtrl.OnEvent("Change", DataSourceChanged)
    dataSourceCtrl.Choose(DataSourceToDisplay(DataSource))
    SettingsGui.Add("Text", "x260 y44 w170 h34 vDataSourceWarn", "")

    SettingsGui.Add("Text", "x20 y80", "统计范围:")
    statsScopeCtrl := SettingsGui.Add("DropDownList", "x90 y76 w120 vStatsScope", ["自动", "全部网卡", "自定义"])
    statsScopeCtrl.OnEvent("Change", StatsScopeChanged)
    SettingsGui.Add("Button", "x220 y74 w60 h22 vSelectAdaptersBtn", "选择...").OnEvent("Click", SelectAdapters)
    SettingsGui.Add("Text", "x290 y80 w140 vStatsScopeHint", "")

    statsScopeCtrl.Choose(StatsScope)
    StatsScopeChanged()
    UpdateDataSourceStatusText()

    smoothCtrl := SettingsGui.Add("Checkbox", "x20 y110 vEnableSmoothing", "平滑处理")
    smoothCtrl.Value := EnableSmoothing

    SettingsGui.Add("Text", "x20 y140", "EMA 平滑因子 (0-1):")
    SettingsGui.Add("Edit", "x150 y136 w50 vEMAFactor", EMAFactor)
    SettingsGui.Add("Text", "x210 y140", "（若启用平滑处理推荐0.35）")

    SettingsGui.Add("Text", "x20 y170", "防抖确认次数:")
    SettingsGui.Add("Edit", "x150 y166 w50 vConfirmNeeded Number", ConfirmNeeded)
    SettingsGui.Add("UpDown", "vConfirmNeededUD Range1-10", ConfirmNeeded)
    SettingsGui.Add("Text", "x210 y170", "（颜色更新需连续出现几次）")

    autoStartCtrl := SettingsGui.Add("Checkbox", "x20 y200 vAutoStart", "开机自启动")
    autoStartCtrl.Value := AutoStart
    autoStartCtrl.OnEvent("Click", AutoStartChanged)

    SettingsGui.Add("Text", "x160 y200", "范围:")
    autoStartScopeCtrl := SettingsGui.Add("DropDownList", "x200 y196 w90 vAutoStartScope", ["当前用户", "所有用户"])
    autoStartScopeCtrl.Choose(AutoStartScope = "当前用户" ? 1 : 2)

    SettingsGui.Add("Button", "x300 y196 w45 h22", "当前").OnEvent("Click", OpenCurrentStartup)
    SettingsGui.Add("Button", "x350 y196 w45 h22", "全局").OnEvent("Click", OpenGlobalStartup)

    if (!AutoStart) {
        SettingsGui["AutoStartScope"].Enabled := false
    }

    tab.UseTab()
    SettingsGui.Add("Button", "x200 y270 w60 h30", "保存").OnEvent("Click", SaveSettings)
    SettingsGui.Add("Button", "x270 y270 w60 h30", "取消").OnEvent("Click", CloseSettings)
    SettingsGui.Add("Button", "x340 y270 w80 h30", "恢复默认").OnEvent("Click", ResetSettings)

    InitColorPreview()
    OnlyTextChanged()

    SettingsGui.Title := "网速监控设置"
    SettingsGui.Show("w450 h320")
}

MouseThroughChanged(*) {
    global MouseThrough, DragPositioning, SettingsGui
    MouseThrough := SettingsGui["MouseThrough"].Value
    if (MouseThrough && DragPositioning) {
        MsgBox("启用拖动定位时，鼠标穿透功能将被禁用以确保拖动操作正常工作。`n`n保存设置后将自动应用此调整。", "提示", 0x30)
    }
}

AutoStartChanged(*) {
    global AutoStart, SettingsGui
    AutoStart := SettingsGui["AutoStart"].Value
    SettingsGui["AutoStartScope"].Enabled := AutoStart
}

StatsScopeChanged(*) {
    global StatsScope, SettingsGui
    StatsScope := SettingsGui["StatsScope"].Text
    SettingsGui["SelectAdaptersBtn"].Enabled := (StatsScope = "自定义")
    UpdateStatsScopeHint()
}

DataSourceChanged(*) {
    UpdateDataSourceStatusText()
}

SelectAdapters(*) {
    ShowAdapterSelectDialog()
}

DragPositioningChanged(*) {
    global DragPositioning, MouseThrough, SettingsGui
    DragPositioning := SettingsGui["DragPositioning"].Value
    if (DragPositioning) {
        SettingsGui["PositionCorner"].Enabled := false
        SettingsGui["OffsetX"].Enabled := false
        SettingsGui["OffsetXUD"].Enabled := false
        SettingsGui["OffsetY"].Enabled := false
        SettingsGui["OffsetYUD"].Enabled := false
        SettingsGui["LimitOffset"].Enabled := false
        if (MouseThrough) {
            MsgBox("启用拖动定位时，鼠标穿透功能将被禁用以确保拖动操作正常工作。`n`n保存设置后将自动应用此调整。", "提示", 0x30)
        }
    } else {
        SettingsGui["PositionCorner"].Enabled := true
        SettingsGui["OffsetX"].Enabled := true
        SettingsGui["OffsetXUD"].Enabled := true
        SettingsGui["OffsetY"].Enabled := true
        SettingsGui["OffsetYUD"].Enabled := true
        SettingsGui["LimitOffset"].Enabled := true
    }
}

ShowLineMarkersChanged(*) {
    global ShowLineMarkers, ColorChangeWithValue, CombinedMode, SettingsGui
    ShowLineMarkers := SettingsGui["ShowLineMarkers"].Value
    SettingsGui["LineMarkerStyle"].Enabled := ShowLineMarkers
    CombinedMode := ColorChangeWithValue && ShowLineMarkers
}

ColorChangeWithValueChanged(*) {
    global ColorChangeWithValue, ShowLineMarkers, CombinedMode, SettingsGui
    ColorChangeWithValue := SettingsGui["ColorChangeWithValue"].Value
    CombinedMode := ColorChangeWithValue && ShowLineMarkers
}

LineMarkerStyleChange(*) {
    global LineMarkerStyle, SettingsGui
    LineMarkerStyle := SettingsGui["LineMarkerStyle"].Text
}

InitColorPreview(*) {
    UpdateColorPreviews()
}

ColorEditChanged(*) {
    UpdateColorPreviews()
}

BgColorChanged(*) {
    UpdateColorPreviews()
}

UpdateColorPreviews() {
    global SettingsGui, ColorVeryLow, ColorLow, ColorMed, ColorHigh, BgColorMode, BgColor
    if (!SettingsGui)
        return
    SettingsGui["PrevVeryLow"].Opt("c" ColorVeryLow)
    SettingsGui["PrevLow"].Opt("c" ColorLow)
    SettingsGui["PrevMed"].Opt("c" ColorMed)
    SettingsGui["PrevHigh"].Opt("c" ColorHigh)

    if (BgColorMode = "自定义") {
        bgr := RgbOrBgrToBgrNo0x(BgColor)
        SettingsGui["PrevBgColor"].Opt("c" bgr)
        SettingsGui["PrevBgColor"].Opt("+Background" bgr)
        SettingsGui["PrevBgColor"].Enabled := true
    } else {
        SettingsGui["PrevBgColor"].Enabled := false
    }
}

UpdateStatsScopeHint() {
    global SettingsGui, StatsScope, SelectedAdapters
    if (!SettingsGui)
        return
    if (StatsScope = "自定义") {
        if (SelectedAdapters = "")
            hint := "已选: 0"
        else
            hint := "已选: " StrSplit(SelectedAdapters, "|").Length
    } else if (StatsScope = "自动") {
        hint := "当前: 自动"
    } else {
        hint := "当前: 全部"
    }
    SettingsGui["StatsScopeHint"].Text := hint
}

UpdateDataSourceStatusText() {
    global SettingsGui
    if (!SettingsGui)
        return

    selected := NormalizeDataSource(SettingsGui["DataSource"].Text)
    warn := ""

    if (selected = "IPHelper") {
        if (!CheckIpHelperAvailable())
            warn := "IP Helper 不可用，保存后回退 WMI"
    }

    if (warn = "") {
        SettingsGui["DataSourceWarn"].Opt("c008000")
        SettingsGui["DataSourceWarn"].Text := "可用"
    } else {
        SettingsGui["DataSourceWarn"].Opt("cC00000")
        SettingsGui["DataSourceWarn"].Text := warn
    }
}

RgbOrBgrToBgrNo0x(c) {
    s := Trim(c)
    if (SubStr(s, 1, 2) = "0x" || SubStr(s, 1, 2) = "0X") {
        rgb := SubStr(s, 3)
        if (StrLen(rgb) = 6) {
            r := SubStr(rgb, 1, 2), g := SubStr(rgb, 3, 2), b := SubStr(rgb, 5, 2)
            return b g r
        }
        return "000000"
    }
    if RegExMatch(s, "^[0-9A-Fa-f]{6}$")
        return s
    return "000000"
}

BgColorModeChange(*) {
    global BgColorMode, SettingsGui
    BgColorMode := SettingsGui["BgColorMode"].Text
    if (BgColorMode = "自定义") {
        SettingsGui["BgColor"].Enabled := true
        SettingsGui["PrevBgColor"].Enabled := true
        SettingsGui["PickBgBtn"].Enabled := true
    } else {
        SettingsGui["BgColor"].Enabled := false
        SettingsGui["PrevBgColor"].Enabled := false
        SettingsGui["PickBgBtn"].Enabled := false
    }
    UpdateColorPreviews()
}

FontNamePresetChange(*) {
    global SettingsGui
    if (SettingsGui["FontNamePreset"].Text = "自定义") {
        SettingsGui["FontNameCustom"].Visible := true
        SettingsGui["FontNameCustomLabel"].Visible := true
    } else {
        SettingsGui["FontNameCustom"].Visible := false
        SettingsGui["FontNameCustomLabel"].Visible := false
    }
}

SignedIntEditFilter(ctrl, *) {
    val := ctrl.Value
    cleaned := RegExReplace(val, "[^\d-]", "")
    if (StrLen(cleaned) > 1)
        cleaned := SubStr(cleaned, 1, 1) . RegExReplace(SubStr(cleaned, 2), "-", "")
    if (cleaned != val)
        ctrl.Value := cleaned
}

OnlyTextChanged(*) {
    global OnlyText, SettingsGui
    OnlyText := SettingsGui["OnlyText"].Value
    SettingsGui["BgTransparency"].Enabled := !OnlyText
    SettingsGui["BgTransparencyUD"].Enabled := !OnlyText
}

SaveSettings(*) {
    global SettingsGui, ConfigFile
    global Interval, AutoRestart, MouseThrough
    global GuiWidth, GuiHeight, FontName, FontSize, FontWeight
    global BgColorMode, BgColor, OnlyText, BgTransparency
    global NumRightMargin, ArrowWidth, PositionCorner, OffsetX, OffsetY
    global Display, LimitOffset
    global Thresh1, Thresh2, Thresh3
    global ColorVeryLow, ColorLow, ColorMed, ColorHigh
    global EnableSmoothing, EMAFactor, ConfirmNeeded
    global StatsScope, SelectedAdapters
    global DragPositioning
    global AutoStart, AutoStartScope
    global ShowLineMarkers, LineMarkerStyle, ColorChangeWithValue
    global DataSource
    global CombinedMode

    values := SettingsGui.Submit(false)

    if (values.FontNamePreset = "自定义")
        FontName := Trim(values.FontNameCustom)
    else
        FontName := Trim(values.FontNamePreset)

    Interval := RegExReplace(values.Interval, "[,\s]", "")
    GuiWidth := RegExReplace(values.GuiWidth, "[,\s]", "")
    GuiHeight := RegExReplace(values.GuiHeight, "[,\s]", "")
    FontSize := RegExReplace(values.FontSize, "[,\s]", "")
    BgTransparency := RegExReplace(values.BgTransparency, "[,\s]", "")
    OffsetX := RegExReplace(values.OffsetX, "[,\s]", "")
    OffsetY := RegExReplace(values.OffsetY, "[,\s]", "")
    Thresh1KB := RegExReplace(values.Thresh1KB, "[,\s]", "")
    Thresh2KB := RegExReplace(values.Thresh2KB, "[,\s]", "")
    Thresh3MB := RegExReplace(values.Thresh3MB, "[,\s]", "")
    ConfirmNeeded := RegExReplace(values.ConfirmNeeded, "[,\s]", "")
    EMAFactor := Trim(values.EMAFactor)
    StatsScope := Trim(values.StatsScope)
    DataSource := NormalizeDataSource(values.DataSource)
    LimitOffset := values.LimitOffset ? 1 : 0
    DragPositioning := values.DragPositioning ? 1 : 0

    MouseThrough := values.MouseThrough
    BgColorMode := values.BgColorMode
    BgColor := values.BgColor
    OnlyText := values.OnlyText
    ColorVeryLow := values.ColorVeryLow
    ColorLow := values.ColorLow
    ColorMed := values.ColorMed
    ColorHigh := values.ColorHigh
    EnableSmoothing := values.EnableSmoothing
    AutoRestart := values.AutoRestart
    AutoStart := values.AutoStart
    AutoStartScope := values.AutoStartScope
    ShowLineMarkers := values.ShowLineMarkers
    LineMarkerStyle := values.LineMarkerStyle
    ColorChangeWithValue := values.ColorChangeWithValue
    Display := values.Display
    FontWeight := values.FontWeight

    if (DragPositioning)
        MouseThrough := 0

    if (Interval < 100 || Interval > 5000)
        Interval := 1000
    if (GuiWidth < 80 || GuiWidth > 300)
        GuiWidth := 120
    if (GuiHeight < 30 || GuiHeight > 100)
        GuiHeight := 44
    if (FontSize < 8 || FontSize > 24)
        FontSize := 11
    if (BgTransparency < 0 || BgTransparency > 255)
        BgTransparency := 255
    if (Thresh1KB < 1 || Thresh1KB > 1000)
        Thresh1KB := 50
    if (Thresh2KB < 1 || Thresh2KB > 5000)
        Thresh2KB := 500
    if (Thresh3MB < 1 || Thresh3MB > 100)
        Thresh3MB := 2
    if (ConfirmNeeded < 1 || ConfirmNeeded > 10)
        ConfirmNeeded := 2
    if (StatsScope != "自动" && StatsScope != "全部网卡" && StatsScope != "自定义")
        StatsScope := "全部网卡"
    if (DataSource = "")
        DataSource := "WMI"
    if (EMAFactor = "")
        EMAFactor := DefaultEMAFactor
    else if (EMAFactor < 0)
        EMAFactor := "0"
    else if (EMAFactor > 1)
        EMAFactor := "1"

    if (!IsHex6(ColorVeryLow) || !IsHex6(ColorLow) || !IsHex6(ColorMed) || !IsHex6(ColorHigh)) {
        MsgBox("颜色必须为6位16进制（如 CFCFCF）。请检查颜色输入后重试。", "保存失败", 0x10)
        return
    }

    if (BgColorMode = "自定义") {
        nBg := NormalizeBgColor(BgColor)
        if (nBg = "") {
            MsgBox("自定义背景色必须为 0xRRGGBB 或 RRGGBB。请检查后重试。", "保存失败", 0x10)
            return
        }
        BgColor := nBg
    }

    t1 := Thresh1KB * 1024
    t2 := Thresh2KB * 1024
    t3 := Thresh3MB * 1024 * 1024
    adj := false
    if (t1 >= t2) {
        t2 := t1 + 1024
        Thresh2KB := Ceil(t2 / 1024.0)
        adj := true
    }
    if (t2 >= t3) {
        t3 := t2 + 1024
        Thresh3MB := Ceil(t3 / 1024.0 / 1024.0)
        adj := true
    }
    if (adj)
        MsgBox("网速阈值已自动调整为递增顺序。保存后生效。", "已自动调整", 0x30)

    IniWrite(Interval, ConfigFile, "General", "Interval")
    IniWrite(AutoRestart, ConfigFile, "General", "AutoRestart")
    IniWrite(MouseThrough, ConfigFile, "Settings", "MouseThrough")
    IniWrite(GuiWidth, ConfigFile, "GUI", "Width")
    IniWrite(GuiHeight, ConfigFile, "GUI", "Height")
    IniWrite(FontName, ConfigFile, "GUI", "FontName")
    IniWrite(FontSize, ConfigFile, "GUI", "FontSize")
    IniWrite(FontWeight, ConfigFile, "GUI", "FontWeight")
    IniWrite(BgColorMode, ConfigFile, "GUI", "BgColorMode")
    IniWrite(BgColor, ConfigFile, "GUI", "BgColor")
    IniWrite(OnlyText, ConfigFile, "GUI", "OnlyText")
    IniWrite(BgTransparency, ConfigFile, "GUI", "BgTransparency")
    IniWrite(ShowLineMarkers, ConfigFile, "GUI", "ShowLineMarkers")
    IniWrite(LineMarkerStyle, ConfigFile, "GUI", "LineMarkerStyle")
    IniWrite(ColorChangeWithValue, ConfigFile, "GUI", "ColorChangeWithValue")
    CombinedMode := ColorChangeWithValue && ShowLineMarkers
    IniWrite(PositionCorner, ConfigFile, "Position", "Corner")
    IniWrite(OffsetX, ConfigFile, "Position", "OffsetX")
    IniWrite(OffsetY, ConfigFile, "Position", "OffsetY")
    IniWrite(Display, ConfigFile, "GUI", "Display")
    IniWrite(LimitOffset, ConfigFile, "Position", "LimitOffset")

    IniWrite(t1, ConfigFile, "Thresholds", "Thresh1")
    IniWrite(t2, ConfigFile, "Thresholds", "Thresh2")
    IniWrite(t3, ConfigFile, "Thresholds", "Thresh3")

    IniWrite(ColorVeryLow, ConfigFile, "Colors", "VeryLow")
    IniWrite(ColorLow, ConfigFile, "Colors", "Low")
    IniWrite(ColorMed, ConfigFile, "Colors", "Medium")
    IniWrite(ColorHigh, ConfigFile, "Colors", "High")

    IniWrite(EnableSmoothing, ConfigFile, "Advanced", "EnableSmoothing")
    IniWrite(EMAFactor, ConfigFile, "Advanced", "EMAFactor")
    IniWrite(ConfirmNeeded, ConfigFile, "Advanced", "ConfirmNeeded")
    IniWrite(StatsScope, ConfigFile, "Advanced", "StatsScope")
    IniWrite(SelectedAdapters, ConfigFile, "Advanced", "StatsAdapters")
    IniWrite(DataSource, ConfigFile, "Advanced", "DataSource")

    IniWrite(DragPositioning, ConfigFile, "Position", "DragPositioning")

    oldAutoStart := IniRead(ConfigFile, "General", "AutoStart", 0)
    oldScope := IniRead(ConfigFile, "General", "AutoStartScope", "当前用户")

    IniWrite(AutoStart, ConfigFile, "General", "AutoStart")
    IniWrite(AutoStartScope, ConfigFile, "General", "AutoStartScope")

    if (AutoStart && !oldAutoStart) {
        if (!CreateAutoStartShortcut(AutoStartScope)) {
            MsgBox("创建开机自启动快捷方式失败，为所有用户创建需要以管理员身份运行程序。", "提示", 0x30)
        }
    } else if (!AutoStart && oldAutoStart) {
        DeleteAutoStartShortcut(oldScope)
    } else if (AutoStart && oldAutoStart && AutoStartScope != oldScope) {
        DeleteAutoStartShortcut(oldScope)
        if (!CreateAutoStartShortcut(AutoStartScope)) {
            MsgBox("更新开机自启动快捷方式失败，为所有用户更新需要以管理员身份运行程序。", "提示", 0x30)
        }
    }

    if (AutoRestart) {
        Reload()
    } else {
        if (MsgBox("设置已保存！需要重启程序以应用新设置。是否现在重启？", "设置已保存", 0x4) = "Yes")
            Reload()
    }
}

CloseSettings(*) {
    global SettingsGui
    if (SettingsGui)
        SettingsGui.Destroy()
}

ResetSettings(*) {
    if (MsgBox("确定要重置所有设置为默认值吗？", "确认重置", 0x4) = "Yes") {
        FileDelete(ConfigFile)
        CreateDefaultConfig()
        if (SettingsGui)
            SettingsGui.Destroy()
        MsgBox("设置已重置为默认值！请重启程序以应用新设置。", "提示")
    }
}

SettingsGuiClose(*) {
    CloseSettings()
}

UpdateNet(*) {
    global CombinedMode, recv, sent, sSent, sRecv, candidateUp, candidateDown, MainGui
    global EMAFactor, ConfirmNeeded
    global pendingUp, pendingDown, pendingCountUp, pendingCountDown
    global lastColorUp, lastColorDown, lastDisplayColorUp, lastDisplayColorDown
    global lastTextUp, lastTextDown, emaUp, emaDown
    global StatsScope, SelectedAdapters

    recv := 0
    sent := 0
    sSent := ""
    sRecv := ""
    candidateUp := ""
    candidateDown := ""

    GetSpeedData(&recv, &sent)

    if (EMAFactor = 1)
        emaUp := sent, emaDown := recv
    else {
        emaUp := (emaUp ? emaUp * (1 - EMAFactor) + sent * EMAFactor : sent)
        emaDown := (emaDown ? emaDown * (1 - EMAFactor) + recv * EMAFactor : recv)
    }

    candidateUp := GetColorBySpeed(emaUp)
    candidateDown := GetColorBySpeed(emaDown)

    if (candidateUp = pendingUp)
        pendingCountUp += 1
    else {
        pendingUp := candidateUp
        pendingCountUp := 1
    }
    if (pendingCountUp >= ConfirmNeeded && pendingUp != lastColorUp) {
        lastColorUp := pendingUp
        pendingCountUp := 0
    }

    if (candidateDown = pendingDown)
        pendingCountDown += 1
    else {
        pendingDown := candidateDown
        pendingCountDown := 1
    }
    if (pendingCountDown >= ConfirmNeeded && pendingDown != lastColorDown) {
        lastColorDown := pendingDown
        pendingCountDown := 0
    }

    sSent := FormatSpeed(emaUp)
    sRecv := FormatSpeed(emaDown)

    if (CombinedMode) {
        GetLineMarkers(&upSymbol, &downSymbol)
        sSentWithArrow := sSent " " upSymbol
        sRecvWithArrow := sRecv " " downSymbol

        if (sSentWithArrow != lastTextUp) {
            MainGui["UpNum"].Text := sSentWithArrow
            lastTextUp := sSentWithArrow
        }
        if (sRecvWithArrow != lastTextDown) {
            MainGui["DownNum"].Text := sRecvWithArrow
            lastTextDown := sRecvWithArrow
        }

        if (lastColorUp != lastDisplayColorUp) {
            MainGui["UpNum"].SetFont("c" lastColorUp)
            lastDisplayColorUp := lastColorUp
        }
        if (lastColorDown != lastDisplayColorDown) {
            MainGui["DownNum"].SetFont("c" lastColorDown)
            lastDisplayColorDown := lastColorDown
        }
    } else {
        if (sSent != lastTextUp) {
            MainGui["UpNum"].Text := sSent
            lastTextUp := sSent
        }
        if (sRecv != lastTextDown) {
            MainGui["DownNum"].Text := sRecv
            lastTextDown := sRecv
        }

        if (lastColorUp != lastDisplayColorUp) {
            MainGui["UpNum"].SetFont("c" lastColorUp)
            MainGui["UpArrow"].SetFont("c" lastColorUp)
            lastDisplayColorUp := lastColorUp
        }
        if (lastColorDown != lastDisplayColorDown) {
            MainGui["DownNum"].SetFont("c" lastColorDown)
            MainGui["DownArrow"].SetFont("c" lastColorDown)
            lastDisplayColorDown := lastColorDown
        }
    }

    EnsureTopmostOnTaskbarActive()
}

EnsureTopmostOnTaskbarActive(*) {
    global hGui
    if (!hGui)
        return
    WinSetAlwaysOnTop(true, "ahk_id " hGui)
}

NormalizeDataSource(value) {
    v := Trim(value)
    if (v = "WMI" || v = "IPHelper")
        return v
    if (v = "WMI (默认)")
        return "WMI"
    if (v = "IP Helper API")
        return "IPHelper"
    return "WMI"
}

DataSourceToDisplay(value) {
    if (value = "IPHelper")
        return 2
    return 1
}

CheckIpHelperAvailable() {
    size := 0
    status := 0
    try {
        status := DllCall("iphlpapi.dll\GetIfTable", "ptr", 0, "uint*", &size, "int", false, "uint")
    } catch as e {
        return false
    }
    return (status = 0 || status = 122)
}

InitDataSource() {
    global DataSource, wmi, WmiWarned
    global IpPrevTick, IpPrevMap
    global IpHelperAvailable

    wmi := ""

    if (DataSource = "WMI") {
        InitWmi()
    } else if (DataSource = "IPHelper") {
        IpHelperAvailable := true
        IpPrevTick := 0
        IpPrevMap := Map()
    }
}

SwitchToWmiDataSource() {
    global DataSource, wmi
    DataSource := "WMI"
    if (!wmi)
        InitWmi()
}

HandleIpHelperFailure(&recv, &sent, tipText) {
    global IpHelperAvailable, IpHelperWarned, IpPrevTick, IpPrevMap
    IpHelperAvailable := false
    IpPrevTick := 0
    IpPrevMap := Map()
    if (!IpHelperWarned) {
        IpHelperWarned := true
        TrayTip(tipText, "网速监控")
    }
    SwitchToWmiDataSource()
    GetSpeedDataWmi(&recv, &sent)
}

InitWmi() {
    global wmi, WmiWarned
    wmi := ""
    try {
        wmi := ComObjGet("winmgmts:{impersonationLevel=impersonate}!//./root/cimv2")
        test := wmi.ExecQuery("SELECT BytesReceivedPersec, BytesSentPersec FROM Win32_PerfFormattedData_Tcpip_NetworkInterface")
    } catch as e {
        try {
            test := wmi.ExecQuery("SELECT BytesReceivedPersec, BytesSentPersec FROM Win32_PerfFormattedData_Tcpip_TCPv4")
        } catch as e2 {
            wmi := ""
        }
    }
    if (!wmi && !WmiWarned) {
        WmiWarned := true
        TrayTip("WMI 接口不可用，速度显示可能为 0。可尝试以管理员运行或重启系统。", "网速监控")
    }
}

GetSpeedData(&recv, &sent) {
    global DataSource, IpHelperAvailable
    recv := 0
    sent := 0
    if (DataSource = "IPHelper") {
        if (!IpHelperAvailable) {
            SwitchToWmiDataSource()
            GetSpeedDataWmi(&recv, &sent)
            return
        }
        GetSpeedDataIpHelper(&recv, &sent)
        return
    }
    GetSpeedDataWmi(&recv, &sent)
}

GetSpeedDataWmi(&recv, &sent) {
    global wmi, StatsScope, SelectedAdapters
    if (!wmi)
        return
    try {
        q := wmi.ExecQuery("SELECT BytesReceivedPersec, BytesSentPersec, Name FROM Win32_PerfFormattedData_Tcpip_NetworkInterface")
        if (!q.Count)
            q := wmi.ExecQuery("SELECT BytesReceivedPersec, BytesSentPersec, Name FROM Win32_PerfFormattedData_Tcpip_TCPv4")
        for item in q {
            name := item.Name
            r := item.BytesReceivedPersec ? item.BytesReceivedPersec : 0
            s := item.BytesSentPersec ? item.BytesSentPersec : 0

            if (StatsScope = "自定义" && SelectedAdapters != "" && !IsAdapterSelected(name))
                continue
            if (StatsScope = "自动" && (r + s = 0))
                continue

            recv += r
            sent += s
        }
    } catch as e {
        recv := 0
        sent := 0
    }
}

GetSpeedDataIpHelper(&recv, &sent) {
    global StatsScope, SelectedAdapters
    global IpPrevTick, IpPrevMap
    global IpHelperAvailable

    rows := GetIpHelperRows()
    if (Type(rows) != "Array") {
        HandleIpHelperFailure(&recv, &sent, "IP Helper API 不可用，已回退到 WMI 数据源。")
        return
    }

    now := A_TickCount
    currentMap := Map()
    for row in rows {
        currentMap[row.idx] := {in: row.inOctets, out: row.outOctets}
    }

    if (IpPrevTick = 0) {
        IpPrevTick := now
        IpPrevMap := currentMap
        recv := 0
        sent := 0
        return
    }

    deltaMs := now - IpPrevTick
    if (deltaMs <= 0) {
        IpPrevTick := now
        IpPrevMap := currentMap
        recv := 0
        sent := 0
        return
    }

    totalDeltaIn := 0
    totalDeltaOut := 0
    wrap32 := 4294967296
    wrapHigh := 0xF0000000
    wrapLow := 0x0FFFFFFF

    for row in rows {
        if (StatsScope = "自定义" && SelectedAdapters != "") {
            if (!IsAdapterSelected(row.name) && !IsAdapterSelected(row.desc) && !IsAdapterSelected(row.alias))
                continue
        }

        if (!IpPrevMap.Has(row.idx))
            continue

        prev := IpPrevMap[row.idx]
        deltaIn := row.inOctets - prev.in
        deltaOut := row.outOctets - prev.out

        if (deltaIn < 0) {
            if (prev.in >= wrapHigh && row.inOctets <= wrapLow)
                deltaIn += wrap32
            else
                deltaIn := 0
        }
        if (deltaOut < 0) {
            if (prev.out >= wrapHigh && row.outOctets <= wrapLow)
                deltaOut += wrap32
            else
                deltaOut := 0
        }

        if (StatsScope = "自动" && (deltaIn + deltaOut = 0))
            continue

        totalDeltaIn += deltaIn
        totalDeltaOut += deltaOut
    }

    recv := (totalDeltaIn * 1000) / deltaMs
    sent := (totalDeltaOut * 1000) / deltaMs

    IpPrevTick := now
    IpPrevMap := currentMap
}

GetIpHelperRows() {
    size := 0
    status := 0

    try {
        status := DllCall("iphlpapi.dll\GetIfTable", "ptr", 0, "uint*", &size, "int", false, "uint")
    } catch as e {
        return ""
    }

    if (status != 0 && status != 122)
        return ""
    if (size <= 0)
        return []

    table := Buffer(size, 0)
    try {
        status := DllCall("iphlpapi.dll\GetIfTable", "ptr", table.Ptr, "uint*", &size, "int", false, "uint")
    } catch as e {
        return ""
    }

    if (status != 0)
        return ""

    try {
        rows := []
        numEntries := NumGet(table, 0, "uint")
        rowSize := 860
        rowBase := table.Ptr + 4

        Loop numEntries {
            rowPtr := rowBase + (A_Index - 1) * rowSize
            idx := NumGet(rowPtr + 512, "uint")
            alias := Trim(StrGet(rowPtr, 256, "UTF-16"))

            descLen := NumGet(rowPtr + 600, "uint")
            if (descLen > 256)
                descLen := 256
            desc := Trim(StrGet(rowPtr + 604, descLen, "CP0"))

            name := (desc != "" ? desc : alias)
            inOctets := NumGet(rowPtr + 552, "uint")
            outOctets := NumGet(rowPtr + 576, "uint")

            rows.Push({idx: idx, name: name, alias: alias, desc: desc, inOctets: inOctets, outOctets: outOctets})
        }
        return rows
    } catch as e {
        return ""
    }
}

GetColorBySpeed(val) {
    global Thresh1, Thresh2, Thresh3
    global ColorVeryLow, ColorLow, ColorMed, ColorHigh

    if (val < Thresh1)
        return ColorVeryLow
    else if (val < Thresh2)
        return ColorLow
    else if (val < Thresh3)
        return ColorMed
    else
        return ColorHigh
}

CreateGuiAndShow(hexColor) {
    global GuiWidth, GuiHeight, FontName, FontSize, FontWeight
    global NumWidth, ArrowWidth, UpY, DownY, ArrowX, BgColor, BgTransparency, OnlyText
    global UpNum, UpArrow, DownNum, DownArrow
    global hGui, MouseThrough, DragPositioning, CombinedMode
    global MainGui

    MainGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    MainGui.MarginX := 0
    MainGui.MarginY := 0
    fontOpt := "s" FontSize
    if (FontWeight = "Bold")
        fontOpt .= " Bold"
    MainGui.SetFont(fontOpt, FontName)
    hGui := MainGui.Hwnd

    if (CombinedMode) {
        GetLineMarkers(&upSymbol, &downSymbol)
        MainGui.Add("Text", "x0 y" UpY " w" GuiWidth " vUpNum Right c" hexColor " BackgroundTrans", "初始化...")
        MainGui.Add("Text", "x0 y" DownY " w" GuiWidth " vDownNum Right c" hexColor " BackgroundTrans", "初始化...")
        MainGui.Add("Text", "x-100 y-100 w1 vUpArrow Center c" hexColor " BackgroundTrans", "")
        MainGui.Add("Text", "x-100 y-100 w1 vDownArrow Center c" hexColor " BackgroundTrans", "")
    } else {
        GetLineMarkers(&upSymbol, &downSymbol)
        MainGui.Add("Text", "x0 y" UpY " w" NumWidth " vUpNum Right c" hexColor " BackgroundTrans", "初始化...")
        MainGui.Add("Text", "x" ArrowX " y" UpY " w" ArrowWidth " vUpArrow Center c" hexColor " BackgroundTrans", upSymbol)
        MainGui.Add("Text", "x0 y" DownY " w" NumWidth " vDownNum Right c" hexColor " BackgroundTrans", "初始化...")
        MainGui.Add("Text", "x" ArrowX " y" DownY " w" ArrowWidth " vDownArrow Center c" hexColor " BackgroundTrans", downSymbol)
    }

    MainGui.BackColor := BgColor

    PositionGui()
    ApplyGuiTransparency()

    if (MouseThrough && !DragPositioning)
        WinSetExStyle("+0x20", "ahk_id " hGui)
    else
        WinSetExStyle("-0x20", "ahk_id " hGui)

    if (DragPositioning) {
        OnMessage(0x0201, OnLButtonDown)
        OnMessage(0x0202, OnLButtonUp)
        OnMessage(0x0200, OnMouseMove)
    }
}

ApplyGuiTransparency() {
    global hGui, BgColor, BgTransparency, OnlyText
    if (OnlyText) {
        WinSetTransparent("Off", "ahk_id " hGui)
        WinSetTransColor(BgColor " 255", "ahk_id " hGui)
    } else {
        WinSetTransColor("Off", "ahk_id " hGui)
        WinSetTransparent(BgTransparency, "ahk_id " hGui)
    }
}

GetTargetWorkArea(&sx, &sy, &sw, &sh) {
    global Display
    if (Display = "全部") {
        sx := SysGet(76)
        sy := SysGet(77)
        sw := SysGet(78)
        sh := SysGet(79)
        return
    } else if (Display = "主屏幕" || Display = "") {
        idx := MonitorGetPrimary()
    } else {
        idx := RegExReplace(Display, "\D", "")
        if (idx = "")
            idx := 1
    }
    MonitorGetWorkArea(idx, &waLeft, &waTop, &waRight, &waBottom)
    sx := waLeft, sy := waTop, sw := waRight - waLeft, sh := waBottom - waTop
}

PositionGui() {
    global GuiWidth, GuiHeight, PositionCorner, OffsetX, OffsetY, hGui, LimitOffset, MainGui
    GetTargetWorkArea(&screenX, &screenY, &screenW, &screenH)

    if (PositionCorner = "右下角") {
        baseX := screenX + screenW - GuiWidth
        baseY := screenY + screenH - GuiHeight
        x := baseX - OffsetX
        y := baseY - OffsetY
    } else if (PositionCorner = "右上角") {
        baseX := screenX + screenW - GuiWidth
        baseY := screenY
        x := baseX - OffsetX
        y := baseY + OffsetY
    } else if (PositionCorner = "左下角") {
        baseX := screenX
        baseY := screenY + screenH - GuiHeight
        x := baseX + OffsetX
        y := baseY - OffsetY
    } else {
        baseX := screenX
        baseY := screenY
        x := baseX + OffsetX
        y := baseY + OffsetY
    }

    if (LimitOffset) {
        maxX := screenX + screenW - GuiWidth
        maxY := screenY + screenH - GuiHeight
        if (x < screenX)
            x := screenX
        if (y < screenY)
            y := screenY
        if (x > maxX)
            x := maxX
        if (y > maxY)
            y := maxY
    }

    MainGui.Show("x" x " y" y " w" GuiWidth " h" GuiHeight " NoActivate")
}

GetLineMarkers(&upSymbol, &downSymbol) {
    global ShowLineMarkers, LineMarkerStyle

    if (!ShowLineMarkers) {
        upSymbol := ""
        downSymbol := ""
        return
    }

    if (LineMarkerStyle = "↑ ↓") {
        upSymbol := "↑"
        downSymbol := "↓"
    } else if (LineMarkerStyle = "TX RX") {
        upSymbol := "TX"
        downSymbol := "RX"
    } else if (LineMarkerStyle = "⬆️⬇️") {
        upSymbol := "⬆️"
        downSymbol := "⬇️"
    } else if (LineMarkerStyle = "🔼🔽") {
        upSymbol := "🔼"
        downSymbol := "🔽"
    } else if (LineMarkerStyle = "⏫⏬") {
        upSymbol := "⏫"
        downSymbol := "⏬"
    } else if (LineMarkerStyle = "上传 下载") {
        upSymbol := "上传"
        downSymbol := "下载"
    } else if (LineMarkerStyle = "上行 下行") {
        upSymbol := "上行"
        downSymbol := "下行"
    } else if (LineMarkerStyle = "发送 接收") {
        upSymbol := "发送"
        downSymbol := "接收"
    } else {
        upSymbol := "↑"
        downSymbol := "↓"
    }
}

FormatSpeed(val) {
    if (val >= 1048576)
        return Round(val / 1048576, 2) " MB/s"
    else if (val >= 1024)
        return Round(val / 1024, 1) " KB/s"
    else
        return Round(val, 0) "  B/s"
}

HasValue(arr, value) {
    for item in arr {
        if (item = value)
            return true
    }
    return false
}

IsHex6(s) {
    return RegExMatch(s, "^[0-9A-Fa-f]{6}$")
}

NormalizeBgColor(s) {
    s := Trim(s)
    if (SubStr(s, 1, 2) = "0x" || SubStr(s, 1, 2) = "0X") {
        if (StrLen(s) = 8)
            return s
        else
            return ""
    } else if RegExMatch(s, "^[0-9A-Fa-f]{6}$") {
        return "0x" s
    }
    return ""
}

GetAdapterList() {
    global DataSource
    if (DataSource = "IPHelper")
        return GetAdapterListIpHelper()
    return GetAdapterListWmi()
}

GetAdapterListWmi() {
    global wmi
    list := ""
    if (wmi) {
        try {
            q := wmi.ExecQuery("SELECT Name FROM Win32_PerfFormattedData_Tcpip_NetworkInterface")
            for item in q {
                name := item.Name
                if (name != "" && !InStr("|" list "|", "|" name "|"))
                    list .= (list = "" ? "" : "|") name
            }
        } catch as e {
            list := ""
        }
    }
    return list
}

GetAdapterListIpHelper() {
    rows := GetIpHelperRows()
    if (Type(rows) != "Array")
        return GetAdapterListWmi()

    list := ""
    for row in rows {
        name := row.name
        if (name != "" && !InStr("|" list "|", "|" name "|"))
            list .= (list = "" ? "" : "|") name
    }

    if (list = "")
        return GetAdapterListWmi()
    return list
}

IsAdapterSelected(name) {
    global SelectedAdapters
    if (SelectedAdapters = "")
        return false
    return InStr("|" SelectedAdapters "|", "|" name "|")
}

ShowAdapterSelectDialog() {
    global AdapterList, SelectedAdapters, AdapterSelectGui

    AdapterList := GetAdapterList()
    if (AdapterList = "") {
        MsgBox("未获取到网卡列表，请检查当前数据来源是否可用。", "提示", 0x30)
        return
    }

    if (AdapterSelectGui)
        AdapterSelectGui.Destroy()

    AdapterSelectGui := Gui("+AlwaysOnTop +ToolWindow")
    AdapterSelectGui.MarginX := 10
    AdapterSelectGui.MarginY := 10
    AdapterSelectGui.Add("Text", "x10 y10", "请选择需要统计的网卡（可多选）:")
    lv := AdapterSelectGui.Add("ListView", "x10 y30 w420 h220 Checked vAdapterListView", ["网卡名称"])
    AdapterSelectGui.Add("Button", "x10 y260 w60 h24", "确定").OnEvent("Click", AdapterSelectOK)
    AdapterSelectGui.Add("Button", "x80 y260 w60 h24", "取消").OnEvent("Click", AdapterSelectCancel)
    AdapterSelectGui.Add("Button", "x350 y260 w80 h24", "刷新列表").OnEvent("Click", AdapterSelectRefresh)
    AdapterSelectGui.OnEvent("Close", AdapterSelectCancel)

    for name in StrSplit(AdapterList, "|") {
        row := lv.Add("", name)
        if (IsAdapterSelected(name))
            lv.Modify(row, "Check")
    }

    AdapterSelectGui.Title := "网卡选择"
    AdapterSelectGui.Show("w440 h300")
}

AdapterSelectOK(*) {
    global SelectedAdapters, AdapterSelectGui
    lv := AdapterSelectGui["AdapterListView"]
    selected := ""
    row := 0
    while (row := lv.GetNext(row, "Checked")) {
        name := lv.GetText(row, 1)
        selected .= (selected = "" ? "" : "|") name
    }
    SelectedAdapters := selected
    UpdateStatsScopeHint()
    AdapterSelectGui.Destroy()
}

AdapterSelectCancel(*) {
    global AdapterSelectGui
    if (AdapterSelectGui)
        AdapterSelectGui.Destroy()
}

AdapterSelectRefresh(*) {
    global AdapterList, AdapterSelectGui
    AdapterList := GetAdapterList()
    if (AdapterList = "") {
        MsgBox("未获取到网卡列表，请检查当前数据来源是否可用。", "提示", 0x30)
        return
    }
    lv := AdapterSelectGui["AdapterListView"]
    lv.Delete()
    for name in StrSplit(AdapterList, "|") {
        row := lv.Add("", name)
        if (IsAdapterSelected(name))
            lv.Modify(row, "Check")
    }
}

StartColorPicker(targetControl) {
    global PickerActive, CurrentPickTarget, hPicker, PickerLastRGB
    global PickerGui

    if (PickerActive)
        return

    PickerActive := true
    CurrentPickTarget := targetControl
    PickerLastRGB := "FFFFFF"

    KeyWait("LButton")

    if (PickerGui)
        PickerGui.Destroy()

    PickerGui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20")
    PickerGui.MarginX := 8
    PickerGui.MarginY := 8
    PickerGui.BackColor := "0xF5F5F5"
    PickerGui.SetFont("s10", "Segoe UI")

    PickerGui.Add("Text", "x8 y8 c000000", "左键确定 右键取消")
    PickerGui.Add("Progress", "x8 y28 w110 h50 vPickPrev +Border", 100)

    CoordMode("Mouse", "Screen")
    MouseGetPos(&mx, &my)
    px := mx + 20
    py := my + 20
    PickerGui.Show("x" px " y" py " AutoSize NoActivate")

    Hotkey("~LButton Up", PickerConfirm, "On")
    Hotkey("~RButton Up", PickerCancel, "On")
    Hotkey("Esc", PickerCancel, "On")
    Hotkey("Enter", PickerConfirm, "On")
    Hotkey("Space", PickerConfirm, "On")

    SetTimer(PickerTick, 30)
}

PickerTick(*) {
    global PickerActive, PickerLastRGB, PickerGui
    if (!PickerActive)
        return

    CoordMode("Pixel", "Screen")
    CoordMode("Mouse", "Screen")

    MouseGetPos(&mx, &my)
    px := mx + 20
    py := my + 20
    PickerGui.Show("x" px " y" py " NoActivate")

    colRGB := PixelGetColor(mx, my, "RGB") & 0xFFFFFF
    hexRGB := Format("{:06X}", colRGB)

    PickerLastRGB := hexRGB
    PickerGui["PickPrev"].Opt("c" hexRGB)
    PickerGui["PickPrev"].Opt("+Background" hexRGB)
}

PickerConfirm(*) {
    global PickerActive, CurrentPickTarget, PickerLastRGB, SettingsGui
    if (!PickerActive)
        return

    SettingsGui[CurrentPickTarget].Value := PickerLastRGB
    InitColorPreview()
    PickerStop()
}

PickerCancel(*) {
    PickerStop()
}

PickerStop() {
    global PickerActive, PickerGui
    SetTimer(PickerTick, 0)
    Hotkey("~LButton Up", "Off")
    Hotkey("~RButton Up", "Off")
    Hotkey("Esc", "Off")
    Hotkey("Enter", "Off")
    Hotkey("Space", "Off")
    if (PickerGui)
        PickerGui.Destroy()
    PickerActive := false
}

PickVeryLow(*) {
    StartColorPicker("ColorVeryLow")
}

PickLow(*) {
    StartColorPicker("ColorLow")
}

PickMed(*) {
    StartColorPicker("ColorMed")
}

PickHigh(*) {
    StartColorPicker("ColorHigh")
}

PickBgColor(*) {
    StartColorPicker("BgColor")
}

ShowAbout(*) {
    global AboutGui
    if (AboutGui)
        AboutGui.Destroy()

    AboutGui := Gui("+AlwaysOnTop +ToolWindow")
    AboutGui.MarginX := 20
    AboutGui.MarginY := 20
    AboutGui.SetFont("s11", "Segoe UI Variable")

    AboutGui.SetFont("s16 Bold")
    AboutGui.Add("Text", "x20 y20 w400 Center", "Display Network Speed")

    AboutGui.SetFont("s10 Normal")
    AboutGui.Add("Text", "x20 y50 w400 Center c666666", "显示网速")

    AboutGui.Add("Progress", "x20 y80 w400 h2 BackgroundDDDDDD", 100)

    AboutGui.SetFont("s10")
    AboutGui.Add("Text", "x20 y100", "版本: v1.0.0")
    AboutGui.Add("Text", "x20 y125", "基于: AutoHotkey v1")
    AboutGui.Add("Text", "x20 y150", "作者：YI")
    AboutGui.Add("Text", "x20 y175", "项目地址：")

    urlCtrl := AboutGui.Add("Text", "x20 y195 w400", "https://github.com/Yinengjun/DisplayNetworkSpeed")
    urlCtrl.SetFont("c0000FF")
    urlCtrl.OnEvent("Click", OpenProjectURL)

    AboutGui.SetFont("s11 Bold")
    AboutGui.Add("Text", "x20 y230", "主要功能")

    AboutGui.SetFont("s10 Normal")
    AboutGui.Add("Text", "x30 y255", "• 实时网速显示 (WMI接口)")
    AboutGui.Add("Text", "x30 y275", "• 智能颜色编码 (4档阈值)")
    AboutGui.Add("Text", "x30 y295", "• EMA平滑处理")
    AboutGui.Add("Text", "x30 y315", "• 多显示器支持")
    AboutGui.Add("Text", "x30 y335", "• 高度自定义配置")

    AboutGui.Add("Button", "x20 y380 w120 h30", "打开配置文件夹").OnEvent("Click", OpenConfigFolder)

    AboutGui.Title := "关于 Display Network Speed"
    AboutGui.Show("w440 h430")
}

OpenProjectURL(*) {
    Run("https://github.com/Yinengjun/DisplayNetworkSpeed")
}

OpenConfigFolder(*) {
    Run("explorer " Chr(34) A_ScriptDir Chr(34))
}

RestartApp(*) {
    Reload()
}

ExitAppHandler(*) {
    ExitApp()
}

OnLButtonDown(wParam, lParam, msg, hwnd) {
    global hGui, DragPositioning, IsDragging
    global DragStartX, DragStartY, DragStartGuiX, DragStartGuiY
    if (!DragPositioning || hwnd != hGui)
        return

    CoordMode("Mouse", "Screen")
    MouseGetPos(&DragStartX, &DragStartY)
    WinGetPos(&DragStartGuiX, &DragStartGuiY, , , "ahk_id " hGui)

    IsDragging := true
    DllCall("SetCapture", "Ptr", hGui)
}

OnMouseMove(wParam, lParam, msg, hwnd) {
    global hGui, DragPositioning, IsDragging
    global DragStartX, DragStartY, DragStartGuiX, DragStartGuiY
    if (!DragPositioning || !IsDragging || hwnd != hGui)
        return

    CoordMode("Mouse", "Screen")
    MouseGetPos(&currentX, &currentY)

    newX := DragStartGuiX + (currentX - DragStartX)
    newY := DragStartGuiY + (currentY - DragStartY)
    WinMove(newX, newY, , , "ahk_id " hGui)
}

OnLButtonUp(wParam, lParam, msg, hwnd) {
    global hGui, DragPositioning, IsDragging
    global DragStartX, DragStartY, DragStartGuiX, DragStartGuiY
    global PositionCorner, OffsetX, OffsetY, LimitOffset
    global GuiWidth, GuiHeight

    if (!DragPositioning || !IsDragging || hwnd != hGui)
        return

    IsDragging := false
    DllCall("ReleaseCapture")
    WinGetPos(&currentGuiX, &currentGuiY, , , "ahk_id " hGui)
    GetTargetWorkArea(&screenX, &screenY, &screenW, &screenH)

    centerX := currentGuiX + GuiWidth / 2
    centerY := currentGuiY + GuiHeight / 2
    screenCenterX := screenX + screenW / 2
    screenCenterY := screenY + screenH / 2

    if (centerX >= screenCenterX && centerY >= screenCenterY)
        PositionCorner := "右下角"
    else if (centerX >= screenCenterX && centerY < screenCenterY)
        PositionCorner := "右上角"
    else if (centerX < screenCenterX && centerY >= screenCenterY)
        PositionCorner := "左下角"
    else
        PositionCorner := "左上角"

    if (PositionCorner = "右下角") {
        baseX := screenX + screenW - GuiWidth
        baseY := screenY + screenH - GuiHeight
    } else if (PositionCorner = "右上角") {
        baseX := screenX + screenW - GuiWidth
        baseY := screenY
    } else if (PositionCorner = "左下角") {
        baseX := screenX
        baseY := screenY + screenH - GuiHeight
    } else {
        baseX := screenX
        baseY := screenY
    }

    if (PositionCorner = "右下角") {
        OffsetX := baseX - currentGuiX
        OffsetY := baseY - currentGuiY
    } else if (PositionCorner = "右上角") {
        OffsetX := baseX - currentGuiX
        OffsetY := currentGuiY - baseY
    } else if (PositionCorner = "左下角") {
        OffsetX := currentGuiX - baseX
        OffsetY := baseY - currentGuiY
    } else {
        OffsetX := currentGuiX - baseX
        OffsetY := currentGuiY - baseY
    }

    if (LimitOffset) {
        if (OffsetX < -500)
            OffsetX := -500
        if (OffsetX > 500)
            OffsetX := 500
        if (OffsetY < -200)
            OffsetY := -200
        if (OffsetY > 200)
            OffsetY := 200
    }

    SavePositionConfig()
}

SavePositionConfig() {
    global ConfigFile, PositionCorner, OffsetX, OffsetY, LimitOffset
    IniWrite(PositionCorner, ConfigFile, "Position", "Corner")
    IniWrite(OffsetX, ConfigFile, "Position", "OffsetX")
    IniWrite(OffsetY, ConfigFile, "Position", "OffsetY")
    IniWrite(LimitOffset, ConfigFile, "Position", "LimitOffset")
}

CreateAutoStartShortcut(scope) {
    if (scope = "当前用户")
        startupFolder := A_AppData "\Microsoft\Windows\Start Menu\Programs\Startup"
    else
        startupFolder := "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"

    shortcutPath := startupFolder "\DisplayNetworkSpeed.lnk"
    try {
        FileCreateShortcut(A_ScriptFullPath, shortcutPath)
        return true
    } catch as e {
        return false
    }
}

DeleteAutoStartShortcut(scope) {
    if (scope = "当前用户")
        startupFolder := A_AppData "\Microsoft\Windows\Start Menu\Programs\Startup"
    else
        startupFolder := "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"

    shortcutPath := startupFolder "\DisplayNetworkSpeed.lnk"
    try {
        FileDelete(shortcutPath)
        return true
    } catch as e {
        return false
    }
}

CheckAutoStartShortcut(scope) {
    if (scope = "当前用户")
        startupFolder := A_AppData "\Microsoft\Windows\Start Menu\Programs\Startup"
    else
        startupFolder := "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
    shortcutPath := startupFolder "\DisplayNetworkSpeed.lnk"
    return FileExist(shortcutPath)
}

OpenCurrentStartup(*) {
    Run("explorer " Chr(34) A_AppData "\Microsoft\Windows\Start Menu\Programs\Startup" Chr(34))
}

OpenGlobalStartup(*) {
    Run("explorer " Chr(34) "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup" Chr(34))
}
