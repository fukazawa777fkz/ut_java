Public Enum srcPos
    offsetPhysical = 10
    offsetType = 11
    offsetRequired = 12
    offsetMin = 13
    offsetMax = 14
    offsetEnum = 15
End Enum


Public Enum destPos
    offsetPhysical = 10
    offsetType = 11
    offsetRequired = 12
    offsetMin = 13
    offsetMax = 14
    offsetEnum = 15
End Enum


Public Enum eTestType
    NORMAL
    ABNORMAL
End Enum

Public Const FixedFiledNum = 2
Public Const TestTypePos = -2




'Public Sub openFolder()
'
'    Dim strExe As String
'    strExe = "C:\Windows\explorer.exe"
'
'
'    Shell strExe & " " & ActiveWorkbook.Path, vbNormalFocus
'
'End Sub
'
'
'
'Public Sub open_file()
'
'    Dim strExe As String
'    strExe = "C:\Program Files (x86)\sakura\sakura.exe"
'
'    For Each c In ActiveWindow.RangeSelection
'        Shell strExe & " " & c.value, vbNormalFocus
'    Next
'
'    'Shell strExe & " " & ActiveCell.Value, vbNormalFocus
'
'End Sub
'
'Public Sub open_file_eclipse()
'
'    Dim strExe As String
'    strExe = "C:\pleiades\eclipse\eclipse.exe"
'
'
'    Shell strExe & " " & ActiveCell.value, vbNormalFocus
'
'End Sub
'
'
'Public Sub a1macro()
'    Dim c As Worksheet
'    For Each c In ActiveWorkbook.Worksheets
'        c.Activate
'        c.Range("a1").Activate
'        c.Range("a1").Show
'        c.Activate
'        ActiveWindow.Zoom = 100
'    Next
'    ActiveWorkbook.Worksheets(1).Activate
'
'End Sub
'
'Public Sub makeSheet()
'
'    Dim wStart As Worksheet: Set wStart = ActiveSheet
'    Dim iMax As Integer: iMax = InputBox("最大値を入力", "", 10)
'
'    Dim st As Worksheet: Set st = ActiveSheet
'    For i = 1 To iMax
'        Set st = ActiveWorkbook.Worksheets.Add(After:=st)
'        On Error Resume Next
'        st.name = i
'    Next
'
'    wStart.Activate
'
'End Sub
'
'
''Public Sub openURLPathForGoogle2()
''    For Each c In ActiveWindow.RangeSelection
''
''        c.Activate
''        openURLPathForGoogle
''    Next
''End Sub
'
'
'
'Public Sub aaa()
'
'End Sub
'
'Public Sub numberling()
'    Dim c As Range
'    Dim iPearent As Integer
'    iPearent = 0
'    Dim iChild As Integer
'    iChild = 0
'
'    iPearent = InputBox("start", "", iPearent)
'    For Each c In ActiveWindow.RangeSelection
'        If c.Offset(0, 1).value <> "" & c.Offset(0, 1).value <> c.Offset(0, 1).value Then
'            iPearent = iPearent + 1
'            iChild = 0
'        End If
'        iChild = iChild + 1
'        c.value = "'" + CStr(iPearent) & "-" & CStr(iChild)
'    Next
'End Sub
'
'
'Public Sub set_active_line()
'    Range(ActiveCell.Address & ":" & ActiveCell.Offset(0, 100).Address).Interior.ColorIndex = Yellow_Green
'End Sub
'Public Sub clear_active_line()
'    Range(ActiveCell.Address & ":" & ActiveCell.Offset(0, 200).Address).Interior.ColorIndex = xlNone
'End Sub
'

