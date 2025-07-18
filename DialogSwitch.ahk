#Requires AutoHotkey v2.0
#SingleInstance Force
FileEncoding "UTF-8"

; ================== 1. 初始化与全局变量 ==================
#Include <UIA>
SetWorkingDir A_InitialWorkingDir

; --- 全局变量定义 ---
global iniPath := A_InitialWorkingDir "\DialogSwitch.ini"
global actWinC := 0
global main := Menu()
global root := Map()
global dopusRT_Path := "F:\SoftWare\Explorer\Directory Opus\dopusrt.exe"
global g_qt := ""
global g_application := ""
global editGui := Gui()
global extendArray := Array()

; --- 开关与状态变量 ---
global isPathCollectorEnabled := true
global isSingleGroupExplore := true
global groupUngroupedItems := false
global pname := ""
global ptitle := ""

; --- 初始化操作 ---
;global qt := ComObject("QTTabBarLib.Scripting")
InitializeEditConfigGui()
InitializeTrayMenu()
; =============================================================

; ================== 2. 托盘菜单管理 ==================
InitializeTrayMenu() {
    global MyTrayMenu := A_TrayMenu
    global shortcutPath

    ; 首先，删除所有默认的标准菜单项
    MyTrayMenu.Delete()
    MyTrayMenu.Add("修改配置", ShowEditConfigGui)
    MyTrayMenu.Add("编辑此脚本", (*) => Edit())
    MyTrayMenu.Add()
    MyTrayMenu.Add("打开自启目录", OpenStartupFolder)
    MyTrayMenu.Add("开机自启", ToggleStartupShortcut)
    MyTrayMenu.Add("启用路径收集器", TogglePathCollector)
    if (isPathCollectorEnabled)
        MyTrayMenu.Check("启用路径收集器")
    MyTrayMenu.Add()
    MyTrayMenu.Add("WindowSpy", (*) => Run(A_AhkPath "\..\..\WindowSpy.ahk"))
    MyTrayMenu.Add("挂起/恢复热键", (*) => Suspend())
    MyTrayMenu.Add("重载", (*) => Reload())
    MyTrayMenu.Add("退出", (*) => ExitApp())

    local startupFolder := A_Startup
    shortcutPath := startupFolder "\DialogSwitch.lnk"
    UpdateTrayMenuState()
}

ShowEditConfigGui(*) {
    editGui.Show()
}

TogglePathCollector(ItemName, ItemPos, MyMenu) {
    global isPathCollectorEnabled
    isPathCollectorEnabled := !isPathCollectorEnabled
    MyMenu.ToggleCheck(ItemName)
}

ToggleStartupShortcut(ItemName, ItemPos, MyMenu) {
    if (FileExist(shortcutPath)) {
        MyMenu.Uncheck(ItemName)
        DeleteStartupShortcut()
    } else {
        MyMenu.Check(ItemName)
        CreateStartupShortcut()
    }
}

OpenStartupFolder(*) {
    Run("explorer.exe shell:startup")
}

CreateStartupShortcut() {
    try {
        FileCreateShortcut(A_ScriptFullPath, shortcutPath, A_InitialWorkingDir)
        MsgBox("开机自启动已设置！")
    } catch {
        MsgBox("创建快捷方式失败，请检查权限。")
    }
}

DeleteStartupShortcut() {
    try {
        FileDelete(shortcutPath)
    } catch {
        MsgBox("删除快捷方式失败，请检查权限。")
    }
}

UpdateTrayMenuState() {
    if (FileExist(shortcutPath))
        MyTrayMenu.Check("开机自启")
    else
        MyTrayMenu.Uncheck("开机自启")
}
; =============================================================

; ================== 3. 配置编辑GUI与文件创建 ==================
CreateDefaultConfigFile() {
    global iniPath

    local template := "
    ( LTrim
    ; =========================================================================
    ; DialogSwitch by Ahykal - 对话框路径切换器配置文件
    ; =========================================================================
    ; 欢迎！本文件是 DialogSwitch 的配置文件，首次运行脚本时会自动创建。
    ; 建议：如果使用记事本等外部编辑器修改，请确保保存时使用 UTF-8 编码。
    ; =========================================================================
    ; 如何添加路径 (两种方法)
    ; =========================================================================
    ; **1: 使用“快速添加”功能**
    ; -------------------------------------------------------------------------
    ; 1. 在任意“打开/保存”对话框中，进入想收藏的文件夹。
    ; 2. 按下热键弹出菜单。
    ; 3. 点击菜单底部的 [快速添加]。
    ;    - 如果选中了文件/文件夹，它们的路径会被添加。
    ;    - 如果没选中任何文件，当前所在的文件夹路径会被添加。
    ; 4. 脚本会自动将新路径写入本文件并重载。
    ; -------------------------------------------------------------------------
    ; **2: 手动编辑此文件**
    ; -------------------------------------------------------------------------
    ; 当需要进行批量修改、或设置复杂的分组和规则时，可以手动编辑。
    ; 编辑保存后，右键托盘图标 -> 操作 -> 修改配置 (或直接重载脚本) 即可生效。
    ; =========================================================================
    ; 语法参考
    ; =========================================================================
    ; ; (分号)  : 行首的分号代表注释，脚本会忽略。
    ; 空白行    : 会被忽略，您可以用它来分隔配置块，以提高可读性。
    ;
    ; - (一个减号)  : 定义【程序】。后跟进程名，如 '- Notepad.exe'。
    ;                 特殊值: '- -' 代表“所有程序”，用于设置全局通用规则。
    ;                 [!] 提示: 如果想在文件开头直接定义全局路径，可以省略 '- -' 等头部。
    ;
    ; -- (两个减号) : 定义【窗口标题】。支持正则表达式。
    ;                 例如 '-- 打开' 只匹配标题为“打开”的对话框。
    ;                 特殊值: '-- -' 代表“所有标题”。
    ;
    ; --- (三个减号): 定义【控件文本】。用于匹配对话框内某些特定文本。
    ;                 不常用，通常设为 '--- -' 即可。
    ;
    ; ---- (四个减号): 定义【一级分组】(子菜单)。
    ; ----- (五个减号): 定义【二级分组】，可无限增加减号创建更多层级。
    ;
    ; 无前缀的行   : 定义一个【菜单项】。
    ;                 格式:   显示名称|实际路径
    ;                 如果省略'|'和显示名称，则路径本身就是显示名称。
    ; =========================================================================
    ; 配置示例
    ; =========================================================================
    ; 示例 1: 全局通用规则 (对所有程序的所有对话框生效)
    ; 可以直接在这里写，脚本会自动识别为全局规则。
    ; -------------------------------------------------------------------------
    桌面|C:\Users\{1}\Desktop
    下载|C:\Users\{1}\Downloads
    文档|C:\Users\{1}\Documents
    ; 使用分组
    ---- 我的工作目录
    D:\Projects\Project_A
    D:\Projects\Project_B
    ---- 常用工具
    ----- 开发工具
    C:\Program Files\Microsoft VS Code
    ----- 设计软件
    C:\Program Files\Adobe
    ; -------------------------------------------------------------------------
    ; 示例 2: 特定程序规则 (仅对 Notepad.exe 生效)
    ; -------------------------------------------------------------------------
    - Notepad.exe
    ; 规则A: 当在记事本中“另存为”时
    -- 另存为
    C:\Temp\Notes
    ; 规则B: 当在记事本中“打开”时
    -- 打开
    C:\Users\{1}\Documents
    ; -------------------------------------------------------------------------
    ; 示例 3: 高级规则 (仅对 code.exe 生效)
    ; -------------------------------------------------------------------------
    - Code.exe
    ---- 前端项目
    D:\Project\WebApp\VueProject
    D:\Project\WebApp\ReactProject
    -- 保存工作区
    ; processName 为 code.exe 且 title 为 保存工作区
    D:\WorkSpace
    )"

    local defaultContent := Format(template, A_UserName)

    try {
        FileAppend(defaultContent, iniPath, "UTF-8")
    } catch {
        MsgBox("无法创建默认配置文件，请检查脚本所在目录的写入权限！", "错误", 48)
    }
    return defaultContent
}

InitializeEditConfigGui() {
    editGui.AddEdit("r40 vConfig w900 WantTab")
    editGui["Config"].Value := FileExist(iniPath) ? FileRead(iniPath, "UTF-8") : CreateDefaultConfigFile()
    editGui.AddEdit("r4 vTitle w300 ys WantTab")
    ParseConfigToBuildMapping(editGui["Config"].Value)
    editGui.AddEdit("r35 vText w300 WantTab")
    editGui.AddButton(, "确认").OnEvent("Click", SaveConfigAndReload)
}

SaveConfigAndReload(*) {
    FileDelete(iniPath)
    FileAppend(editGui["Config"].Value, iniPath, "UTF-8")
    Reload()
}
; =============================================================

; ================== 4. 配置文件解析 ==================
ParseConfigToBuildMapping(configTxt) {
    m1 := "", m2 := "", m3 := ""
    groupPath := Array()

    loop parse, configTxt, "`n", "`r" {
        loopField := Trim(A_LoopField)

        if (loopField = "")
            continue

        count := CountLeadingChars(loopField, "-")

        if (count = 1) {
            m1 := LTrim(loopField, "- ")
            m2 := "-"
            m3 := "-"
            groupPath := []
            EnsureProcessNameMapExists(m1)
        } else if (count = 2) {
            m2 := LTrim(loopField, "-- ")
            m3 := "-"
            groupPath := []
            EnsureWindowTitleMapExists(m1, m2)
        } else if (count = 3) {
            m3 := LTrim(loopField, "--- ")
            groupPath := []
            EnsureControlTextMapExists(m1, m2, m3)
        } else if (count >= 4) {
            groupLevel := count - 3
            groupName := LTrim(loopField, "- ")
            while (groupPath.Length >= groupLevel) {
                groupPath.Pop()
            }
            groupPath.Push(groupName)
        }
        else {
            if (SubStr(loopField, 1, 1) != ";") {
                ; 如果还没有设置任何上下文，则自动设为全局上下文
                if (m1 = "") {
                    m1 := m2 := m3 := "-"
                }

                EnsureControlTextMapExists(m1, m2, m3)
                currentMap := root[m1][m2][m3]
                for g in groupPath {
                    if (!currentMap.Has(g))
                        currentMap[g] := Map()
                    currentMap := currentMap[g]
                }
                if (!currentMap.Has("-"))
                    currentMap["-"] := Array()
                currentMap["-"].Push(ParseMenuItemString(loopField))
            }
        }
    }
}

EnsureProcessNameMapExists(m) {
    if (!root.Has(m))
        root[m] := Map()
}

EnsureWindowTitleMapExists(m, n) {
    EnsureProcessNameMapExists(m)
    if (!root[m].Has(n))
        root[m][n] := Map()
}

EnsureControlTextMapExists(m, n, t) {
    EnsureWindowTitleMapExists(m, n)
    if (!root[m][n].Has(t))
        root[m][n][t] := Map()
}

ParseMenuItemString(str) {
    parts := StrSplit(str, "|")
    if (parts.Length = 1) {
        local path := Trim(parts[1])
        return [path, path]
    } else {
        return [Trim(parts[1]), Trim(parts[2])]
    }
}

CountLeadingChars(str, char) {
    count := 0
    loop StrLen(str) {
        if (SubStr(str, A_Index, 1) = char)
            count++
        else
            break
    }
    return count
}
; =============================================================

; ================== 5. 上下文菜单构建 ==================
FindAndBuildMenuForProcess(pname, ptitle, ptext, p*) {
    if (root.Has("-")) {
        a := root["-"]
        FindAndBuildMenuForTitle(a, ptitle, ptext, p*)
    }
    if (root.Has(pname)) {
        b := root[pname]
        FindAndBuildMenuForTitle(b, ptitle, ptext, p*)
    }
}

FindAndBuildMenuForTitle(map, ptitle, ptext, p*) {
    if (map.Has("-")) {
        FindAndBuildMenuForText(map["-"], ptext, p*)
    }
    for k in map {
        if (RegExMatch(ptitle, k))
            FindAndBuildMenuForText(map[k], ptext, p*)
    }
}

FindAndBuildMenuForText(map, ptext, p*) {
    if (p.Length > 0) {
        if (map.Has("-")) {
            FindAndBuildMenuForExtend(map["-"], p*)
        }
        for k in map {
            if (RegExMatch(ptext, k))
                FindAndBuildMenuForExtend(map[k], p*)
        }
    } else {
        if (map.Has("-")) {
            BuildMainMenuFromMap(map["-"])
        }
        for k in map {
            if (RegExMatch(ptext, k))
                BuildMainMenuFromMap(map[k])
        }
    }
}


FindAndBuildMenuForExtend(map, p*) {
    ; 基本情况 1: 如果没有更多的扩展参数 (p* 为空),
    if (p.Length = 0) {
        BuildMainMenuFromMap(map)
        return
    }

    ; 获取当前需要处理的上下文参数 (例如 "Main")
    local currentContextKey := p[1]

    ; 检查当前 map 是否包含这个上下文作为子分组 (键)
    if (map.Has(currentContextKey)) {
        ; 找到匹配的子分组
        ; 准备一个新数组，包含剩下的所有参数
        local remainingParams := []
        local i := 2
        while (i <= p.Length) {
            remainingParams.Push(p[i])
            i++
        }

        ; 使用子分组的 map 和剩下的参数进行递归调用
        FindAndBuildMenuForExtend(map[currentContextKey], remainingParams*)

    } else {
        ; 在当前 map 中找不到匹配的子分组 (例如在 "Main" 分组下找不到 "First")
        ; 显示当前层级的菜单
        BuildMainMenuFromMap(map)
        return
    }
}

BuildMainMenuFromMap(map) {
    local localMap := map.Clone()

    if (localMap.Has("-")) {
        local defaultItems := localMap.Delete("-")
        if (groupUngroupedItems) {
            local subMenu := Menu()
            for e in defaultItems
                subMenu.Add(e[1], HandleMenuItemClick.Bind(e[2]))
            main.Add("-", subMenu)
        } else {
            for e in defaultItems
                main.Add(e[1], HandleMenuItemClick.Bind(e[2]))
        }
    }
    ; // 遍历所有分组
    for groupName, groupData in localMap {
        if (groupData is Object and not HasProp(groupData, "Length")) {
            ;// 如果是分组对象，则递归构建子菜单
            local subMenu := Menu()
            ;// 递归调用以构建子菜单
            BuildSubMenuRecursively(subMenu, groupData)
            main.Add(groupName, subMenu)
        }
    }
}

BuildSubMenuRecursively(parentMenu, menuDataMap) {
    if (menuDataMap.Has("-")) {
        local directItems := menuDataMap["-"]
        for item in directItems
            parentMenu.Add(item[1], HandleMenuItemClick.Bind(item[2]))
    }

    for groupName, groupData in menuDataMap {
        if (groupName != "-" and groupData is Object and not HasProp(groupData, "Length")) {
            local newSubMenu := Menu()
            BuildSubMenuRecursively(newSubMenu, groupData)
            parentMenu.Add(groupName, newSubMenu)
        }
    }
}
; =============================================================

; ================== 6. 路径处理与UIA交互 ==================
BuildExplorerPathsMenu() {
    if (isSingleGroupExplore) {
        parent := Menu()
        AddExplorerPathsToMenu(parent)
        main.Add("explorer", parent)
    } else {
        AddExplorerPathsToMenu(main)
    }
}

AddExplorerPathsFromQT(parent) {
    global g_qt
    if (g_qt == "") { ; 未尝试
        try {
            g_qt := ComObject("QTTabBarLib.Scripting") ; 
        } catch {
            g_qt := 0 ; 失败则标记为 "已尝试但失败" (Failed)
            global g_application
            g_application := ComObject("Shell.Application")
        }
    }
    if IsObject(g_qt) {
        parent.Add("——————QTTabBar——————", HandleMenuItemClick)
        parent.Disable("——————QTTabBar——————")
        for wnd in g_qt.Windows {
            for tab in wnd.Tabs {
                try folder := SubStr(tab.Path, 1, 2) = "::" ? "Shell:" . tab.Path : tab.path
                if (!folder)
                    continue
                parent.Add(folder, HandleMenuItemClick)
            }
        }
        return true
    }
    return false
}

AddExplorerPathsFromDefault(parent) {
    parent.Add "——————Explorer——————", HandleMenuItemClick
    parent.Disable "——————Explorer——————"
    for exp in g_application.Windows {
        path := exp.Document.Folder.Self.Path
        try folder := SubStr(path, 1, 2) = "::" ? "Shell:" . path : path
        if (!folder)
            continue
        parent.Add(folder, HandleMenuItemClick)
    }
}

/**
 * [辅助函数] 通过 dopusrt.exe /info 获取标签页信息，并解析XML文件以提取路径。
 * @returns {Array} 一个包含所有标签页路径的数组。如果失败则返回空数组。
 */
GetDopusTabPaths() {
    if (dopusRT_Path) {

        if !FileExist(dopusRT_Path)
            return []

        tempFile := A_Temp . "\dopus_tabs_" . A_TickCount . ".xml" ; 注意扩展名改为 .xml

        try {
            RunWait('"' dopusRT_Path '" /info "' . tempFile . '",paths', , "Hide")

            loop 50 {
                if FileExist(tempFile)
                    break
                Sleep 10
            }

            if !FileExist(tempFile)
                return []

            ; ======================= 解析XML =======================
            paths := []
            try {
                ; 1. 创建 MSXML DOM Document 对象
                xmlDoc := ComObject("Msxml2.DOMDocument.6.0")
                xmlDoc.async := false ; 确保同步加载

                ; 2. 加载临时 XML 文件
                if !xmlDoc.load(tempFile) {
                    ; 如果加载失败，可以添加错误日志，然后返回
                    return []
                }

                ; 检查解析错误
                if (xmlDoc.parseError.errorCode != 0) {
                    ; 可以添加错误日志，例如:
                    ; MsgBox "XML Parse Error:`n" xmlDoc.parseError.reason
                    return []
                }

                ; 3. 使用 XPath 查询选择所有 <path> 节点
                pathNodes := xmlDoc.selectNodes("//path")

                ; 4. 遍历节点并提取其文本内容 (即路径)
                for node in pathNodes {
                    paths.Push(node.text)
                }
            } catch as e {
                ; 捕获潜在的COM错误
                MsgBox "XML解析时发生COM错误: " . e.Message
                return []
            }
            return paths
            ; ================================================================
        }
        catch {
            return []
        }
        finally {
            if FileExist(tempFile)
                FileDelete(tempFile)
        }
    }
    return [] ; 如果 dopusRT_Path 未设置或无效，则返回空数组
}

AddExplorerPathsFromDopus(parent) {
    ; --- 1. 从 Directory Opus 获取路径 (使用 /info 命令) ---
    dopusPaths := GetDopusTabPaths()
    if (dopusPaths.Length > 0) {
        parent.Add("—————Directory Opus—————", HandleMenuItemClick)
        parent.Disable("—————Directory Opus—————")
        for _, path in dopusPaths {
            if !path
                continue
            try folder := SubStr(path, 1, 2) = "::" ? "Shell:" . path : path
            parent.Add(folder, HandleMenuItemClick)
        }
    }
}

AddExplorerPathsToMenu(parent) {
    if (!AddExplorerPathsFromQT(parent))
        AddExplorerPathsFromDefault(parent)
    AddExplorerPathsFromDopus(parent)

}

HandleMenuItemClick(Item, *) {
Reactive:
    {
        try {
            WinID := WinExist("A")
            WinActivate WinID
        } catch as e {
            Sleep 50
            goto Reactive
        }
    }
    if (RegExMatch(Item, "S)^.:") || RegExMatch(Item, "S)^\\\\file"))
        SetDialogPathViaEdit1(Item)
    else
        SetDialogPathViaAddressBar(Item)
}

SetDialogPathViaEdit1(FolderPath) {
    OldText := ControlGetText("Edit1")
    ControlFocus("Edit1")
    loop 5 {
        ControlSetText(FolderPath, "Edit1")
        Sleep 50
        if (ControlGetText("Edit1") = FolderPath)
            break
    }
    Sleep 50
    Send "{Enter}"
    Sleep 50
    if (!OldText)
        return
    loop 5 {
        ControlSetText(OldText, "Edit1")
        Sleep 50
        if (ControlGetText("Edit1") = OldText)
            break
    }
}

SetDialogPathViaAddressBar(FolderPath) {
    ControlFocus "Edit2"
    Send "{f4}"
    Sleep 50
    ControlSetText FolderPath, "Edit2"
    Sleep 50
    Send "{Enter}"
}

CollectAndAddPathToConfig(*) {
    dialogData := GetDialogInfoWithUIA()

    if (dialogData.error != "") {
        MsgBox("路径收集失败: `n" . dialogData.error, "UIA 错误")
        return
    }

    pathsToAdd := Array()
    if (dialogData.selectedFiles.Length > 0) {
        for path in dialogData.selectedFiles
            pathsToAdd.Push(path)
    } else if (dialogData.currentDir != "" && dialogData.currentDir != "（未能获取路径）") {
        pathsToAdd.Push(dialogData.currentDir)
    }

    if (pathsToAdd.Length = 0) {
        MsgBox("未能获取到任何可用的路径。", "路径收集器")
        return
    }

    WritePathsToConfigFile(pname, ptitle, pathsToAdd)
}

GetDialogInfoWithUIA() {
    output := { error: "", currentDir: "（未能获取路径）", selectedFiles: Array() }

    try {
        if !(dlg := WinExist("A")) || WinGetClass() != "#32770" {
            output.error := "活动窗口不是 #32770 对话框。"
            return output
        }

        dlgElement := UIA.ElementFromHandle(dlg)
        
        local currentDir := ""
        if (rebarPane := dlgElement.FindFirst({ ClassName: "ReBarWindow32" }))
            if (addressBandRoot := rebarPane.FindFirst({ ClassName: "Address Band Root" }))
                if (breadcrumbParent := addressBandRoot.FindFirst({ ClassName: "Breadcrumb Parent" }))
                    if (toolbar := breadcrumbParent.FindFirst({ AutomationId: "1001" })) {
                        rawName := toolbar.Name
                        processedPath := RegExReplace(rawName, "i)^\w+:\s+", "")
                        if (processedPath == rawName) {
                            parts := StrSplit(rawName, ":",, 2)
                            if (parts.Length > 1) 
                                processedPath := Trim(parts[2])
                        }
                        currentDir := processedPath
                        output.currentDir := currentDir
                    }

        if (duiView := dlgElement.FindFirst({ ClassName: "DUIViewWndClassName" }))
            if (shellView := duiView.FindFirst({ ControlType: 50033, ClassName: "DUIListView" })) ; Pane
                if (listView := shellView.FindFirst({ ControlType: 50008, ClassName: "UIItemsView" })) { ; List
                    selection := listView.FindAll({ SelectionItemIsSelected: true })
                    if selection.Length > 0 {
                        if (currentDir) {
                            basePath := RTrim(currentDir, "\") . "\"
                            for item in selection
                                output.selectedFiles.Push(basePath . item.Name)
                        } else {
                            output.error .= (output.error ? "`n" : "") . "错误：选中了文件，但无法确定其所在目录。"
                        }
                    }
                }

    } catch as err {
        output.error := err.Message . "`nStack: " . err.Stack
    }

    return output
}

; GetDialogInfoWithUIA() {
;     output := { error: "", currentDir: "（未能获取路径）", selectedFiles: Array() }

;     try {
;         if !(dlg := WinExist("A")) || WinGetClass() != "#32770" {
;             output.error := "活动窗口不是 #32770 对话框。"
;             return output
;         }

;         dlgElement := UIA.ElementFromHandle(dlg)

;         if (rebarPane := dlgElement.FindFirst({ ClassName: "ReBarWindow32" }))
;             if (addressBandRoot := rebarPane.FindFirst({ ClassName: "Address Band Root" }))
;                 if (breadcrumbParent := addressBandRoot.FindFirst({ ClassName: "Breadcrumb Parent" }))
;                     if (toolbar := breadcrumbParent.FindFirst({ LocalizedControlType: "工具栏", ClassName: "ToolbarWindow32" }))
;                         output.currentDir := RegExReplace(toolbar.CurrentName, "^地址: ", "")

;         if (duiView := dlgElement.FindFirst({ ClassName: "DUIViewWndClassName" }))
;             if (shellView := duiView.FindFirst({ LocalizedControlType: "窗格", ClassName: "DUIListView" }))
;                 if (listView := shellView.FindFirst({ LocalizedControlType: "列表", ClassName: "UIItemsView" })) {
;                     selection := listView.FindAll({ SelectionItemIsSelected: true })
;                     if selection.Length > 0 {
;                         basePath := RTrim(output.currentDir, "\") . "\"
;                         for item in selection
;                             output.selectedFiles.Push(basePath . item.CurrentName)
;                     }
;                 }
;     } catch as err {
;         output.error := err.Message . "`nStack: " . err.Stack
;     }

;     return output
; }

WritePathsToConfigFile(pname, ptitle, pathsArray) {
    configFileContent := FileExist(iniPath) ? FileRead(iniPath, "UTF-8") : ""
    pnameSection := "- " . pname
    ptitleSection := "-- " . ptitle
    lines := StrSplit(configFileContent, "`n", "`r")

    pnameLineNum := -1, ptitleLineNum := -1

    for i, line in lines {
        if (Trim(line) = pnameSection) {
            pnameLineNum := i
            break
        }
    }

    if (pnameLineNum = -1) {
        if (lines.Length > 0 && Trim(lines[lines.Length]) != "")
            lines.Push("")
        lines.Push(pnameSection)
        pnameLineNum := lines.Length
    }

    i := pnameLineNum + 1
    while (i <= lines.Length) {
        trimmedLine := Trim(lines[i])
        if (trimmedLine = ptitleSection) {
            ptitleLineNum := i
            break
        }
        if (SubStr(trimmedLine, 1, 2) = "- ")
            break
        i++
    }

    if (ptitleLineNum = -1) {
        insertionPoint := pnameLineNum + 1
        i := pnameLineNum + 1
        while (i <= lines.Length) {
            if (SubStr(Trim(lines[i]), 1, 2) = "- ") {
                insertionPoint := i
                break
            }
            insertionPoint := i + 1
            i++
        }
        lines.InsertAt(insertionPoint, ptitleSection)
        ptitleLineNum := insertionPoint
    }

    blockStart := ptitleLineNum + 1
    blockEnd := lines.Length

    i := blockStart
    while (i <= lines.Length) {
        trimmedLine := Trim(lines[i])
        if (CountLeadingChars(trimmedLine, "-") > 0) {
            blockEnd := i - 1
            break
        }
        i++
    }

    existingPaths := Map()
    i := blockStart
    while (i <= blockEnd) {
        if (i <= lines.Length && Trim(lines[i]) != "")
            existingPaths[Trim(lines[i])] := true
        i++
    }

    addedCount := 0
    pathsToInsert := Array()
    for path in pathsArray {
        if (!existingPaths.Has(path)) {
            pathsToInsert.Push(path)
            existingPaths[path] := true
            addedCount++
        }
    }

    if (pathsToInsert.Length > 0) {
        insertionPoint := blockEnd + 1
        for path in pathsToInsert {
            lines.InsertAt(insertionPoint, path)
            insertionPoint++
        }
    }

    newContent := ""
    for line in lines
        newContent .= line . "`n"

    FileDelete(iniPath)
    FileAppend(Trim(newContent, "`n`r "), iniPath, "UTF-8")

    if (addedCount > 0) {
        MsgBox(addedCount . " 条新路径已成功添加到配置中！`n即将重载脚本以应用更改。", "路径收集器")
        Reload()
    } else {
        MsgBox("所有路径均已存在，未作修改。", "路径收集器")
    }
}
; =============================================================

; ================== 7. 核心逻辑与热键 ==================
Retry(handle, time) {
ReActive:
    {
        try {
            handle()
        } catch as e {
            Sleep time
            goto ReActive
        }
    }
}

ShowContextMenu(p*) {
Re:
    {
        try {
            actWin := WinExist("A")
        } catch as e {
            Sleep 50
            goto Re
        }
    }

    if (actWinC != actWin) {
        main.Delete()

        global pname := WinGetProcessName()
        global ptitle := WinGetTitle()
        ptext := WinGetText()

        temptxt := "- " . pname . "`n`t- " . ptitle
        editGui["Title"].Value := temptxt
        editGui["Text"].Value := ptext

        BuildExplorerPathsMenu()
        FindAndBuildMenuForProcess(pname, ptitle, ptext, p*)
        if (isPathCollectorEnabled && WinGetClass("A") = "#32770") {
            main.Add()
            main.Add("快速添加", CollectAndAddPathToConfig)
        }
    }
    global actWinC := actWin
    main.Show()
    return
}


#HotIf WinActive("ahk_class #32770")
LWin:: {
    ShowContextMenu()
}

MButton::{
    ShowContextMenu()
}


