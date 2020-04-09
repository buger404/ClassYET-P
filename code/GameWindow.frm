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
   StartUpPosition =   2  '��Ļ����
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
'   ҳ�������
    Dim EC As GMan
'==================================================
'   �ڴ˴��������ҳ����ģ������
    Dim MainPage As MainPage
'==================================================

Private Sub DrawTimer_Timer()
    '����
    EC.Display
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '�����ַ�����
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
            MsgBox "Configure Error : '" & temp & "'", 48, "ҡ�Ż�ѧ�����ô���"
        End If
    Loop
    Close #1
    ReDim Ignored(0)
    Open App.path & "\data\config.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, temp
        ary = Split(temp, "=")
        Select Case ary(0)
            Case "title": Me.Caption = ary(1) & "ר��ҡ�Ż�"
            Case "ignore"
                ary = Split(ary(1), ";")
                For i = 0 To UBound(ary)
                    ReDim Preserve Ignored(UBound(Ignored) + 1)
                    Ignored(UBound(Ignored)) = ary(i)
                Next
        End Select
    Loop
    Close #1
    
    '��ʼ��Emerald���ڴ˴������޸Ĵ��ڴ�СӴ~��
    StartEmerald Me.Hwnd, Screen.Width / Screen.TwipsPerPixelX + 2, Screen.Height / Screen.TwipsPerPixelY + 2
    '��������
    Set EF = New GFont
    EF.AddFont App.path & "\UI.TTF"
    EF.MakeFont "Ҷ��Ⱥ���μ�����"
    '����ҳ�������
    Set EC = New GMan
    EC.Layered False
    
    '�����浵����ѡ��
    'Set ESave = New GSaving
    'ESave.Create "emerald.test", "Emerald.test"
    
    '���������б�
    Set MusicList = New GMusicList
    MusicList.Create App.path & "\music"
    
    '�ڴ˴���ʼ�����ҳ��
    '=============================================
    'ʾ����TestPage.cls
    '     Set TestPage = New TestPage
    '�������֣�Dim TestPage As TestPage
        Set MainPage = New MainPage
    '=============================================

    '��ʼ��ʾ
    Me.Show
    DrawTimer.Enabled = True

    '���ûҳ��
    EC.ActivePage = "MainPage"
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, y As Single)
    '���������Ϣ
    UpdateMouse X, y, 1, button
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, y As Single)
    '���������Ϣ
    If Mouse.State = 0 Then
        UpdateMouse X, y, 0, button
    Else
        Mouse.X = X: Mouse.y = y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
    '���������Ϣ
    UpdateMouse X, y, 2, button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '��ֹ����
    DrawTimer.Enabled = False
    '�ͷ�Emerald��Դ
    EndEmerald
End Sub
