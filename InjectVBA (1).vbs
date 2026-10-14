'==============================================================================
' InjectVBA.vbs —— 把 src/*.bas 的 VBA 代码导入 Master，并生成可点击按钮
' 用法：双击同目录的 Build_Master.bat（或：cscript //nologo InjectVBA.vbs）
' 前置条件：Excel 需信任“对 VBA 工程对象模型的访问”
'   文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置 → 勾选“信任对 VBA 工程对象模型的访问”
' 若未信任，本脚本会弹出英文指引并退出；此时请改用 README 里的“手动导入”（Alt+F11 导入 .bas），
' 手动导入不需要该信任设置。
' 本脚本可重复运行：会先删除旧按钮与旧模块，再重新导入，不会重复生成按钮。
' 容错：若 dist 下的 .xlsm 打不开（例如旧版外壳导致损坏），自动回退到
' build/Master_Shell.xlsx 重建，并另存为合法的 .xlsm。
'==============================================================================
Option Explicit

Dim fso, buildDir, rootDir, srcDir, distDir
Dim masterPath, shellXlsxPath, openedFromXlsx
Dim xlApp, wb, vbProj
Dim basFiles, i, btnDefs, j
Dim ws, shp, rng, shapeName, caption, macroName, cellAddr

Set fso = CreateObject("Scripting.FileSystemObject")
buildDir = fso.GetParentFolderName(WScript.ScriptFullName)
rootDir = fso.GetParentFolderName(buildDir)
srcDir = fso.BuildPath(rootDir, "src")
distDir = fso.BuildPath(rootDir, "dist")
masterPath = fso.BuildPath(distDir, "IUT_Bencase_Automaiton_Master.xlsm")
shellXlsxPath = fso.BuildPath(buildDir, "Master_Shell.xlsx")
openedFromXlsx = False

If Not fso.FileExists(masterPath) And Not fso.FileExists(shellXlsxPath) Then
    MsgBox "Master file not found: " & vbCrLf & masterPath & vbCrLf & _
           "Shell backup also not found: " & vbCrLf & shellXlsxPath, vbExclamation, "Injection Failed"
    WScript.Quit 1
End If

basFiles = Array("modConfig.bas", "modFileOps.bas", "modDiscovery.bas", "modLog.bas", "modRunner.bas", "modUI.bas")
For i = 0 To UBound(basFiles)
    If Not fso.FileExists(fso.BuildPath(srcDir, basFiles(i))) Then
        MsgBox "VBA source file not found: " & vbCrLf & fso.BuildPath(srcDir, basFiles(i)), vbExclamation, "Injection Failed"
        WScript.Quit 1
    End If
Next

' --- 预检：Master 是否已在 Excel 中打开？（存在 ~$ 锁定文件即表示已打开）---
' 修复（2026-10-13，Build 报错“Cannot open the Master file”反馈）：
' 最常见原因是 Master_Shell.xlsx 或 Master .xlsm 还开在 Excel 里（文件被锁定），
' 新 Excel 实例打不开被锁定的文件。提前检查锁定文件，给出明确提示，
' 而不是让用户看到含糊的 COM 错误。
Dim lockXlsx, lockXlsm
lockXlsx = fso.BuildPath(buildDir, "~$Master_Shell.xlsx")
lockXlsm = fso.BuildPath(distDir, "~$IUT_Bencase_Automaiton_Master.xlsm")
If fso.FileExists(lockXlsx) Or fso.FileExists(lockXlsm) Then
    MsgBox "Master appears to be open in Excel (a lock file ~$... was found)." & vbCrLf & vbCrLf & _
           "Please close ALL Excel windows (check the taskbar; if unsure, open Task Manager and end any EXCEL.EXE), then rerun Build_Master.bat." & vbCrLf & vbCrLf & _
           "If no Excel window is open, the lock file may be stale: delete the ~$ file in build/ (or dist/) and rerun.", vbExclamation, "Injection Failed"
    WScript.Quit 1
End If

On Error Resume Next
Set xlApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Cannot start Excel. Please confirm Excel is installed." & vbCrLf & Err.Description, vbExclamation, "Injection Failed"
    WScript.Quit 1
End If
On Error GoTo 0

xlApp.Visible = False
xlApp.DisplayAlerts = False

Dim attempted, lastErr
attempted = ""
lastErr = ""
Set wb = OpenMaster()
If wb Is Nothing Then
    MsgBox "Cannot open the Master file." & vbCrLf & vbCrLf & _
           "Attempted to open:" & vbCrLf & attempted & vbCrLf & _
           "Excel error: " & lastErr & vbCrLf & vbCrLf & _
           "Please try in this order:" & vbCrLf & _
           "1) Task Manager -> end ALL EXCEL.EXE (earlier failed runs may have left hidden Excel processes), then rerun." & vbCrLf & _
           "2) Excel -> File -> Options -> Advanced -> General: UNCHECK 'Ignore other applications that use Dynamic Data Exchange (DDE)', then rerun." & vbCrLf & _
           "3) If it still fails, use Manual Import instead (no automation needed): open Master_Shell.xlsx, Save As .xlsm, Alt+F11, import the 6 .bas files. See README 'Manual Import'.", vbExclamation, "Injection Failed"
    xlApp.Quit
    WScript.Quit 1
End If
On Error GoTo 0

' --- 打开 Master：重试 + 修复加载 + 受保护视图回退 ---
' 修复（2026-10-13，Unblock + 受保护视图回退仍失败后的进一步加固）：
' 剩余可能原因：(a) Excel 启动未完成就调 Open（启动竞态）；(b) 文件有轻微损坏，
' 手动打开时 Excel 会修复、自动化 Open 默认却直接失败；(c) 受保护视图；
' (d) Excel 勾选了“忽略其他使用 DDE 的应用程序”导致自动化被拒绝（需用户手动取消勾选，见上文报错）。
' 本函数逐个候选文件尝试：普通打开 -> 等待重试 -> 修复加载（CorruptLoad=xlRepairFile），
' 最后再试受保护视图 Edit。
Function OpenMaster()
    Dim w, cand, attempt, candidates, ci
    Set w = Nothing
    candidates = Array(masterPath, shellXlsxPath)
    For ci = 0 To UBound(candidates)
        cand = candidates(ci)
        If fso.FileExists(cand) Then
            attempted = attempted & cand & vbCrLf
            For attempt = 1 To 3
                On Error Resume Next
                Err.Clear
                If attempt < 3 Then
                    Set w = xlApp.Workbooks.Open(cand)
                Else
                    ' 第 3 次：修复加载（CorruptLoad=1 xlRepairFile），对应手动打开时的“修复”行为。
                    ' Workbooks.Open 的第 15 个参数是 CorruptLoad，中间参数用 Empty 表示缺省。
                    Set w = xlApp.Workbooks.Open(cand, False, False, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, 1)
                End If
                If Err.Number <> 0 Then lastErr = Err.Description
                If Err.Number = 0 And Not w Is Nothing Then Exit For
                Set w = Nothing
                On Error GoTo 0
                WScript.Sleep 1500
            Next
            If Not w Is Nothing Then
                If InStr(1, w.FullName, "Master_Shell.xlsx", vbTextCompare) > 0 Then openedFromXlsx = True
                Exit For
            End If
        End If
    Next
    ' 受保护视图回退（下载文件的 Mark of the Web 导致 Open 进了受保护视图）
    If w Is Nothing Then
        On Error Resume Next
        Err.Clear
        If xlApp.ProtectedViewWindows.Count > 0 Then
            attempted = attempted & "(Protected View -> Edit)" & vbCrLf
            Set w = xlApp.ProtectedViewWindows(1).Edit
            If Err.Number <> 0 Then lastErr = Err.Description
            If Err.Number = 0 And Not w Is Nothing Then
                If InStr(1, w.FullName, "Master_Shell.xlsx", vbTextCompare) > 0 Then openedFromXlsx = True
            End If
        End If
        On Error GoTo 0
    End If
    Set OpenMaster = w
End Function

' --- 检查 VBA 工程访问权限（需要信任设置）---
On Error Resume Next
Set vbProj = wb.VBProject
If Err.Number <> 0 Or vbProj Is Nothing Then
    MsgBox "Excel does not trust 'Access to the VBA project object model', so automatic injection is not possible." & vbCrLf & vbCrLf & _
           "Choose one of the following:" & vbCrLf & _
           "A) Enable trust and retry: File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> " & vbCrLf & _
           "   check 'Trust access to the VBA project object model' -> OK -> rerun Build_Master.bat" & vbCrLf & _
           "B) Manual import (no trust setting needed): see the 'Manual Import' section in README.md; import src/*.bas with Alt+F11.", _
           vbExclamation, "Trust Access Required"
    wb.Close False
    xlApp.Quit
    WScript.Quit 1
End If
On Error GoTo 0

' --- 删除旧模块（按名删除，使脚本可重复运行）---
Dim comp, toDelete, k, modName
For i = 0 To UBound(basFiles)
    modName = Replace(basFiles(i), ".bas", "")
    On Error Resume Next
    Set comp = vbProj.VBComponents(modName)
    If Err.Number = 0 And Not comp Is Nothing Then
        vbProj.VBComponents.Remove comp
    End If
    Err.Clear
    On Error GoTo 0
Next

' --- 导入 .bas 模块 ---
For i = 0 To UBound(basFiles)
    On Error Resume Next
    vbProj.VBComponents.Import fso.BuildPath(srcDir, basFiles(i))
    If Err.Number <> 0 Then
        MsgBox "Failed to import " & basFiles(i) & ": " & vbCrLf & Err.Description, vbExclamation, "Injection Failed"
        wb.Close False
        xlApp.Quit
        WScript.Quit 1
    End If
    On Error GoTo 0
Next

' --- 生成可点击按钮（形状按钮，OnAction 绑定宏）---
' 每项：工作表 | 定位单元格 | 按钮名 | 标题 | 绑定宏
btnDefs = Array( _
    Array("Control Panel", "F4",  "btnBrowse",      "Browse...",       "BrowseFolder"), _
    Array("Control Panel", "F5",  "btnRefresh",     "Refresh Status",      "RefreshStatus"), _
    Array("Control Panel", "F6",  "btnHelp",        "Help",      "ShowHelp"), _
    Array("Control Panel", "E9",  "btnRun_CCB",     "Run",          "Run_CCB"), _
    Array("Control Panel", "E10", "btnRun_ACWB",    "Run",          "Run_ACWB"), _
    Array("Control Panel", "E11", "btnRun_PEICB",   "Run",          "Run_PEICB"), _
    Array("Control Panel", "E12", "btnRun_OSTC",    "Run",          "Run_OSTC"), _
    Array("Control Panel", "E13", "btnRun_CGEB",    "Run",          "Run_CGEB"), _
    Array("Control Panel", "E14", "btnRun_NLDB",    "Run",          "Run_NLDB"), _
    Array("Control Panel", "C17", "btnRefreshList", "Refresh Script List",  "RefreshMacroListForSelectedTest"), _
    Array("Control Panel", "D17", "btnScanOne",     "Scan Selected Test",  "ScanSelectedTest"), _
    Array("Control Panel", "E17", "btnScanAll",     "Scan All Tests",  "ScanAllTests"), _
    Array("Control Panel", "C18", "btnRunSelected", "Run Selected Script",  "RunSelected"), _
    Array("Control Panel", "A22", "btnRunAll",      "Run All Tests",  "RunAllTests"), _
    Array("Run Log",       "H2",  "btnClearLog",    "Clear Log",      "ClearLog") _
)

For j = 0 To UBound(btnDefs)
    Dim sheetName
    sheetName = btnDefs(j)(0)
    cellAddr = btnDefs(j)(1)
    shapeName = btnDefs(j)(2)
    caption = btnDefs(j)(3)
    macroName = btnDefs(j)(4)
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    If Err.Number = 0 Then
        ' 删除同名旧按钮（可重复运行；倒序遍历，避免边遍历边删除）
        Dim si
        For si = ws.Shapes.Count To 1 Step -1
            If ws.Shapes(si).Name = shapeName Then ws.Shapes(si).Delete
        Next
        Set rng = ws.Range(cellAddr)
        ' 合并单元格取整个合并区的位置与尺寸
        If rng.MergeCells Then Set rng = rng.MergeArea
        Set shp = ws.Shapes.AddShape(5, rng.Left + 1, rng.Top + 1, rng.Width - 2, rng.Height - 2) ' 5 = msoShapeRoundedRectangle
        shp.Name = shapeName
        shp.TextFrame.Characters.Text = caption
        shp.TextFrame.HorizontalAlignment = 2 ' xlHAlignCenter
        shp.TextFrame.VerticalAlignment = 2   ' xlVAlignCenter
        shp.TextFrame.Characters.Font.Size = 11
        shp.TextFrame.Characters.Font.Bold = True
        shp.TextFrame.Characters.Font.Color = RGB(31, 78, 120)
        shp.Fill.ForeColor.RGB = RGB(221, 235, 247)
        shp.Fill.Solid
        shp.Line.ForeColor.RGB = RGB(180, 198, 231)
        shp.Line.Weight = 1
        shp.OnAction = macroName
        ' 清掉外壳里的占位文字（按钮已盖住该单元格）
        rng.Value = ""
    End If
    Err.Clear
    On Error GoTo 0
Next

' --- 保存并退出 ---
' 从 .xlsx 外壳重建时，用 SaveAs 存为 .xlsm（52 = xlOpenXMLWorkbookMacroEnabled），
' 由 Excel 自身生成合法的 .xlsm；否则直接保存。
If openedFromXlsx Then
    On Error Resume Next
    ' 若旧的 .xlsm 还在，先删掉，避免 SaveAs 被旧文件干扰
    wb.SaveAs masterPath, 52
    If Err.Number <> 0 Then
        MsgBox "Save As .xlsm failed: " & vbCrLf & Err.Description, vbExclamation, "Injection Failed"
        wb.Close False
        xlApp.Quit
        WScript.Quit 1
    End If
    On Error GoTo 0
Else
    wb.Save
End If
wb.Close False
xlApp.Quit

MsgBox "Injection complete!" & vbCrLf & vbCrLf & _
       "The Master file now contains all VBA code and clickable buttons." & vbCrLf & _
       "Please open " & masterPath & ", click [Enable Macros], and it is ready to use." & vbCrLf & vbCrLf & _
       "(Optional) In Alt+F11, paste src/ThisWorkbook_Code.txt into ThisWorkbook. See README.", _
       vbInformation, "Injection Complete"
