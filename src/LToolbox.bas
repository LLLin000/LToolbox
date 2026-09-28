Option Explicit

' L的工具箱 —— PowerPoint 图片标定、比例尺与原像素批量裁取
' 核心逻辑在 engine.py / LToolbox.exe；本模块只负责功能区交互。
' 本文件是模板：__PYW__ / __ENG__ / __VERSION__ 由 tools/build_ppam.py 替换。

Private Function LaunchBase() As String
    Dim q As String
    q = Chr(34)
    LaunchBase = q & "__PYW__" & q & " " & q & "__ENG__" & q
End Function

Private Sub RunPy(ByVal action As String, Optional ByVal arg As String = "")
    On Error GoTo CatchErr
    Dim sh As Object, cmd As String, logFile As String, fso As Object, ts As Object
    logFile = Environ("TEMP") & "\LToolbox_run.log"
    
    cmd = LaunchBase & " " & action
    If Len(arg) > 0 Then cmd = cmd & " " & arg
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(logFile, True, True)
    ts.WriteLine "Time: " & Now
    ts.WriteLine "Command: " & cmd
    ts.Close
    
    Set sh = CreateObject("WScript.Shell")
    sh.Run cmd, 0, False
    Exit Sub
CatchErr:
    MsgBox "工具启动失败: " & Err.Description, vbCritical, "L的工具箱"
End Sub


Private Function HasFrame() As Boolean
    Dim i As Long, s As Shape
    On Error Resume Next
    For i = 1 To Application.ActiveWindow.Selection.ShapeRange.Count
        If Application.ActiveWindow.Selection.ShapeRange(i).Type = 1 Then
            HasFrame = True
            Exit Function
        End If
    Next i
    Err.Clear
    For Each s In Application.ActiveWindow.View.Slide.Shapes
        If Left$(s.Name, 6) = "LT_box" Then
            HasFrame = True
            Exit Function
        End If
    Next s
End Function

' 1. 标定所选图片 / 检查所选图片状态
Public Sub LT_Calibrate(Optional control As IRibbonControl = Nothing)
    On Error GoTo ShowStatus
    Dim i As Long, hasLine As Boolean, v As String
    For i = 1 To Application.ActiveWindow.Selection.ShapeRange.Count
        If Application.ActiveWindow.Selection.ShapeRange(i).Type = 9 Then
            hasLine = True
            Exit For
        End If
    Next i
    If Not hasLine Then GoTo ShowStatus
    v = InputBox("请输入这条线对应的真实长度（单位 µm，例如 500）：" & vbCrLf & _
                 "标定会绑定到直线所在的具体图片。", _
                 "① 标定图片", "500")
    If Len(Trim(v)) = 0 Then Exit Sub
    RunPy "calibrate", Trim(v)
    Exit Sub
ShowStatus:
    Err.Clear
    RunPy "status"
End Sub

' 2. 为所选已标定图片生成固定物理尺寸取景框
Public Sub LT_MakeBox(Optional control As IRibbonControl = Nothing)
    Dim v As String
    v = InputBox("请输入取景框的实际边长（单位 µm，例如 200）：", _
                 "② 生成取景框", "200")
    If Len(Trim(v)) = 0 Then Exit Sub
    RunPy "makebox", Trim(v)
End Sub

' 3. 在所选框右下角生成规范比例尺
Public Sub LT_Scalebar(Optional control As IRibbonControl = Nothing)
    If Not HasFrame() Then
        MsgBox "没有找到取景框。请先点击【② 生成取景框】，" & _
               "或选中你自己画的矩形。", vbExclamation, "缺少取景框"
        Exit Sub
    End If
    RunPy "scalebar"
End Sub

' 原图像素批量裁取：选中图片和一个或多个框
Public Sub LT_BatchCrop(Optional control As IRibbonControl = Nothing)
    RunPy "crop"
End Sub

Public Sub LT_Help(Optional control As IRibbonControl = Nothing)
    MsgBox "【L的工具箱 v__VERSION__】" & vbCrLf & vbCrLf & _
           "① 标定 / 状态：在图片自带比例尺上画等长直线并选中；标定直接绑定该图片。选中图片后点击①可显示该图是否已标定。" & vbCrLf & vbCrLf & _
           "② 生成取景框：先选中已标定图片，再输入目标视野边长。旧红框保留，新框自动避让。" & vbCrLf & vbCrLf & _
           "③ 处理：选中框后，在框内右下角添加规范比例尺；不裁图、不删除红框。" & vbCrLf & vbCrLf & _
           "【批量原图裁取】" & vbCrLf & _
           "同时选中一张或多张图片和图片内的所有矩形框，再点击【批量原图裁取】。工具按 PPT 内嵌原图像素裁切，把每张小图放在对应框的原位置；原图和红框不变。此功能不要求标定。", _
           vbInformation, "L的工具箱"
End Sub
