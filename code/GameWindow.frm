VERSION 5.00
Begin VB.Form GameWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "ClassYET"
   ClientHeight    =   6672
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   9000
      Top             =   240
   End
End
Attribute VB_Name = "GameWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================
'   页面管理器
    Dim EC As GMan
'==================================================
'   在此处放置你的页面类模块声明
    Dim MainPage As MainPage
'==================================================

Private Sub DrawTimer_Timer()
    '绘制
    EC.Display
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '发送字符输入
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub

Private Sub Form_Load()
    SetWindowPos Me.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE

    ReDim Stu(0): ReDim Wo(0)
    Dim temp As String, ary() As String
    Open App.path & "\data\students.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, temp
        ary = Split(temp, ";")
        If UBound(ary) = 2 Then
            ReDim Preserve Stu(UBound(Stu) + 1)
            With Stu(UBound(Stu))
                .Name = ary(0)
                .Number = ary(1)
                .Sex = ary(2)
                For i = 1 To Len(.Name)
                    ReDim Preserve Wo(UBound(Wo) + 1)
                    Wo(UBound(Wo)) = Mid(.Name, i, 1)
                Next
            End With
        ElseIf UBound(ary) > 0 Then
            MsgBox "Configure Error : '" & temp & "'", 48, "摇号机学生配置错误"
        End If
    Loop
    Close #1
    ReDim Ignored(0)
    Open App.path & "\data\config.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, temp
        ary = Split(temp, "=")
        Select Case ary(0)
            Case "title": Me.Caption = ary(1) & "专属摇号机"
            Case "ignore"
                ary = Split(ary(1), ";")
                For i = 0 To UBound(ary)
                    ReDim Preserve Ignored(UBound(Ignored) + 1)
                    Ignored(UBound(Ignored)) = ary(i)
                Next
        End Select
    Loop
    Close #1
    
    '初始化Emerald（在此处可以修改窗口大小哟~）
    StartEmerald Me.Hwnd, Screen.Width / Screen.TwipsPerPixelX + 2, Screen.Height / Screen.TwipsPerPixelY + 2
    '创建字体
    Set EF = New GFont
    EF.AddFont App.path & "\UI.TTF"
    EF.MakeFont "叶立群几何极简体"
    '创建页面管理器
    Set EC = New GMan
    EC.Layered False
    
    '创建存档（可选）
    'Set ESave = New GSaving
    'ESave.Create "emerald.test", "Emerald.test"
    
    '创建音乐列表
    Set MusicList = New GMusicList
    MusicList.Create App.path & "\music"
    
    '在此处初始化你的页面
    '=============================================
    '示例：TestPage.cls
    '     Set TestPage = New TestPage
    '公共部分：Dim TestPage As TestPage
        Set MainPage = New MainPage
    '=============================================

    '开始显示
    Me.Show
    DrawTimer.Enabled = True

    '设置活动页面
    EC.ActivePage = "MainPage"
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    UpdateMouse X, y, 1, button
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    If Mouse.State = 0 Then
        UpdateMouse X, y, 0, button
    Else
        Mouse.X = X: Mouse.y = y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    UpdateMouse X, y, 2, button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '终止绘制
    DrawTimer.Enabled = False
    '释放Emerald资源
    EndEmerald
End Sub
