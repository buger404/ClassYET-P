Attribute VB_Name = "RandomCore"
Public Type Student
    Name As String
    Sex As String
    Number As String
End Type
Public Type Stick
    X As Long
    y As Long
    Person As Student
    time As Long
End Type
Public Stu() As Student
Public Ignored() As String
Public Sticks() As Stick
Public Wo() As String

Public Function CheckIgnore(Index As Integer) As Boolean
    Dim i As Integer
    For i = 1 To UBound(Ignored)
        If Ignored(i) = Stu(Index).Number Then CheckIgnore = True: Exit Function
    Next
    CheckIgnore = False
End Function
