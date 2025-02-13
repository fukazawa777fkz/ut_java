Option Explicit
Private writeCurrent As Integer

' �إå�������
Private fileNumber As Integer
Private endpoint As String
Private functionName As String
Private methodType As String
Private formTagL As String
Private formTagU As String

' ����
Private srcRangeSelection As Range
Private srcRowSelection As Range
Private startRange As Range

' ���ϥե�����ե�ѥ�
Private filePath As String

' �ꥹ��
Private minList As Collection ' ItemInfo
Private maxList As Collection ' ItemInfo
Private requiredList As Collection ' ItemInfo
Private enumList As Collection ' ItemInfo
Private defaultList As Collection ' ItemInfo

' ���ס������Τ褦�ʥƥ��Ȥ��������
'    @Nested
'    @DisplayName("/regist-review")
'    class registReview {
'        @Test
'        public void �����_�Ǿ�() throws Exception {
' ��         mockMvc.perform(POST("/regist-review")
' ��             .param("restaurantId", "0")
' ��             .param("userId", "aaa")
' ��             .param("visitDate", "2025-02-12")
' ��             .param("rating", "0")
' ��             .param("comment", ""))
' ��             .andExpect(status().isOk())
' ��             .andExpect(view().name(""))
' ��             .andExpect(model().attributeHasFieldErrors("reviewRegistForm","restaurantId","userId","rating","comment"))
' ��             .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","restaurantId","Min"))
' ��             .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
' ��         ��  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
' ��             .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));
'        }
'    }
Public Sub CreateTestCode()
    
    Call Initialize
    Dim i As Integer
    Dim testCaseValues As Variant
    testCaseValues = srcRangeSelection.value  ' �ƥ��ȥ������ǡ���
    ' ���饹���
    CreateClassDefine
    
    Dim c As Range
    
    For Each c In srcRowSelection
    'For i = LBound(testCaseValues, 1) To UBound(testCaseValues, 1)
        ' �ƥ��ȴؿ���������ꥯ������
        ' �ƥ��ȴؿ�
        Call CreateFuncRequest(c)
        
        ' �ꥯ�����ȥѥ�᡼��
        Call CreateRequestParam1(c)
        
        ' �ǥե���ȥ�å�
        Call CreateMock(c)
        
        ' post get�᥽�å�
        Call CreateMetod(c)
        
        ' �ꥯ�����ȤΥѥ�᡼����������
        ' Call CreateRequestParam2(c)
        Call CreateRequestParam3(c)
        
        ' HTTP���ơ�����
        Call CreateHttpStatus
        
        ' HTML̾
        Debug.Print c.Offset(0, -2).value & ":" & c.Offset(0, -1).value
        Call CreateReturnHtmlName(c)
        
        ' ���顼����
        Call CreateErrorInfo(c)
        
        ' �ƥ��ȴؿ������λ
        Print #fileNumber, "        }"
        Print #fileNumber, ""
    Next
    
    ' ���饹�����λ
    Print #fileNumber, "    }"
    
    Call Terminate
End Sub

Private Sub Initialize()
    
    ' header����
    endpoint = Range("D2").value        '����ɥݥ����
    methodType = Range("D3").value      'POST�Ȥ�
    functionName = Range("D4").value    '���饹̾
    
    'Form̾
    formTagU = Range("D5").value
    formTagL = LCase(Left(formTagU, 1)) & Mid(formTagU, 2)
    

    ' input����
    Set srcRangeSelection = ActiveWindow.RangeSelection
    Set srcRowSelection = Range(srcRangeSelection.Columns(1).Address)
    Set startRange = ActiveCell
    
    '���ϥե�����
    Dim folderPath As String: folderPath = ActiveWorkbook.Path & "\test\"
    ' �ե������̵����к���
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
    filePath = folderPath & getClassName & "Test.java"
    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    

End Sub

Private Sub Terminate()
    Close #fileNumber
    
    If MsgBox("���������ե�����򳫤���", vbYesNo) = vbYes Then
        Dim strExe As String
        strExe = "C:\Program Files (x86)\sakura\sakura.exe"
        Shell strExe & " " & filePath, vbNormalFocus
    End If

End Sub

Private Sub CreateClassDefine()
    Print #fileNumber, Space(4) & "@Nested"
    Print #fileNumber, Space(4) & "@DisplayName(""" & endpoint & """)"
    Print #fileNumber, Space(4) & "class " & functionName & " {"
End Sub

' ���ס��ꥯ�����Ȥ�������ʬ���������
'
' �㡡�������έ����������
' ��   @Test
' ��   public void �����_�Ǿ�() throws Exception {
' ��       // �ꥯ������
' ��       // �ǥե���ȥ�å�
' ��       // defaultMock();
' ��       mockMvc.perform(POST("/regist-review")
'          .param("restaurantId", "0")
'          .param("userId", "aaa")
'          .param("visitDate", "2025-02-12")
'          .param("rating", "0")
'          .param("comment", ""))
'          .andExpect(status().isOk())
'          .andExpect(view().name(""))
'          .andExpect(model().attributeHasFieldErrors("reviewRegistForm","restaurantId","userId","rating","comment"))
'          .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","restaurantId","Min"))
'          .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
'      ��  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
'          .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));
Private Sub CreateFuncRequest(c As Range)
    Print #fileNumber, "        @Test"
    Print #fileNumber, "        public void " & GetTestFunctionName(c) & "() throws Exception {"
    Print #fileNumber, ""
End Sub

Private Sub CreateRequestParam1(rngCurrent As Range)
    Dim currentFieldValues As Variant
    Dim paramMaxCount As Integer
    Dim paramCount As Integer
    Dim j As Integer
    Dim paramLine As String
    Dim ret As Variant
    Dim fields As Range
    
    
    Print #fileNumber, "            // ================== �ꥯ������ =================="
    Print #fileNumber, "            " + formTagU & " " & formTagL & " = new " & formTagU & "();"
    
    Set fields = Range(srcRangeSelection.Rows(rngCurrent.row - startRange.row + 1).Address)
    
    ' �롼�׽���
    Dim c As Range
    Dim rowRange As Range
    For Each c In fields
        ' �ѥ�᡼������
        Dim fieldName As String: fieldName = c.Offset(startRange.row - rngCurrent.row - 1, 0).value
        Dim fieldValue As String: fieldValue = convertValue(c.value)
        
        If c.value = "null" Then
            GoTo ContinueLoop
        End If
        paramLine = "            " & formTagL & ".set" & UCase(Left(fieldName, 1)) & Mid(fieldName, 2) & "(" & fieldValue & ");"
        ' �ե�����񤭹���
        Print #fileNumber, "    " + paramLine

ContinueLoop:
    Next



End Sub


Private Sub CreateMock(c As Range)
    Print #fileNumber, "            // ================== ��å� =================="
    Print #fileNumber, "            defaultMock();"
    Print #fileNumber, ""
End Sub

Private Sub CreateMetod(c As Range)
    Print #fileNumber, "            // ================== �¹� =================="
    Print #fileNumber, "            mockMvc.perform(" & methodType & "(""" & endpoint & """)"
End Sub


' ���ס��ꥯ�����ȤΥѥ�᡼����ʬ���������
'
' �㡡�������έ����������
'      mockMvc.perform(POST("/regist-review")
' ��       .param("restaurantId", "0")
' ��       .param("userId", "aaa")
' ��       .param("visitDate", "2025-02-12")
' ��       .param("rating", "0")
' ��       .param("comment", ""))
'          .andExpect(status().isOk())
'          .andExpect(view().name(""))
'          .andExpect(model().attributeHasFieldErrors("reviewRegistForm","restaurantId","userId","rating","comment"))
'          .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","restaurantId","Min"))
'  ��      .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
'      ��  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
'  ��      .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));

Private Sub CreateRequestParam2(rngCurrent As Range)
    Dim currentFieldValues As Variant
    Dim paramMaxCount As Integer
    Dim paramCount As Integer
    Dim j As Integer
    Dim paramLine As String
    Dim ret As Variant
    Dim fields As Range
    
    
    Set fields = Range(srcRangeSelection.Rows(rngCurrent.row - startRange.row + 1).Address)
    
    ' ͭ�������ǿ��򥫥����
    paramMaxCount = CountNonNullElements(fields)
    paramCount = 0
    
    ' �롼�׽���
    Dim c As Range
    Dim rowRange As Range
    
    For Each c In fields
        
        ' �ѥ�᡼������
        Dim fieldName As String: fieldName = c.Offset(startRange.row - rngCurrent.row - 1, 0).value
        Dim fieldValue As String: fieldValue = convertValue(c.value)
        If c.value = "null" Then
            GoTo ContinueLoop
        End If
        paramLine = "                .param(" & Chr(34) & fieldName & Chr(34) & ", " & fieldValue & ")"
        
        ' �Ǹ�Υѥ�᡼������
        paramCount = paramCount + 1
        If paramCount = paramMaxCount Then
            paramLine = paramLine & ")"
        End If
        
        ' �ե�����񤭹���
        Print #fileNumber, "    " + paramLine

ContinueLoop:
    Next
End Sub

' //                     .flashAttr("taskRegistForm", form)) // �ե����४�֥������Ȥ�����
Private Sub CreateRequestParam3(rngCurrent As Range)
    
    ' �ե�����񤭹���
    Print #fileNumber, "                .flashAttr(" & Chr(34) & formTagL & Chr(34) & " , " & formTagL & ")) // �ե����४�֥�������"

End Sub

' ���ס��ꥯ�����Ȥ�HTTP���ơ��������������
'
' �㡡�������έ����������
'      mockMvc.perform(POST("/regist-review")
'          .param("restaurantId", "0")
'          .param("userId", "aaa")
'          .param("visitDate", "2025-02-12")
'          .param("rating", "0")
'          .param("comment", ""))
' ��       .andExpect(status().isOk())
'          .andExpect(view().name(""))
'          .andExpect(model().attributeHasFieldErrors("reviewRegistForm","restaurantId","userId","rating","comment"))
'          .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","restaurantId","Min"))
' ��       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
'      ��  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
' ��       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));
Private Sub CreateHttpStatus()
    Print #fileNumber, "                .andExpect(status().isOk())"
End Sub

' ���ס�html̾�βս���������
'
' �㡡�������έ����������
'      mockMvc.perform(POST("/regist-review")
'          .param("restaurantId", "0")
'          .param("userId", "aaa")
'          .param("visitDate", "2025-02-12")
'          .param("rating", "0")
'          .param("comment", ""))
'          .andExpect(status().isOk())
' ��       .andExpect(view().name("task-regist-confirm"))
'          .andExpect(model().attributeHasFieldErrors("reviewRegistForm","restaurantId","userId","rating","comment"))
'          .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","restaurantId","Min"))
' ��       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
'      ��  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
' ��       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));
Private Sub CreateReturnHtmlName(rngCurrent As Range)
    Print #fileNumber, "                .andExpect(view().name(""" & rngCurrent.Offset(0, srcRangeSelection.Columns.count + 1).value & """))"
End Sub


' ���ס����顼���ڤκ���
'
' �㡡�������έ��ȭ����������
'      mockMvc.perform(POST("/regist-review")
'          .param("restaurantId", "0")
'          .param("userId", "aaa")
'          .param("visitDate", "2025-02-12")
'          .param("rating", "0")
'          .param("comment", ""))
'          .andExpect(status().isOk())
'          .andExpect(view().name(""))
' ��       .andExpect(model().attributeHasFieldErrors("reviewRegistForm","restaurantId","userId","rating","comment"))
' ��       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","restaurantId","Min"))
' ��       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
' ��   ��  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
' ��       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));
Private Sub CreateErrorInfo(rngCurrent As Range)
    If isNormalTermination(rngCurrent) Then
        Print #fileNumber, "                .andExpect(model().hasNoErrors());"
    Else
        Dim fieldErrors As String
        fieldErrors = getFieldErrors(rngCurrent)
        ' �������
        Print #fileNumber, "                .andExpect(model().attributeHasFieldErrors(" & fieldErrors & "))"
        
        ' �������
        Dim rngFieldErrors As Range: Set rngFieldErrors = getFieldErrorRange(rngCurrent)
        Dim errorCount As Integer: errorCount = 0   ' �ǽ�Ƚ������ѡ����顼�Ǹ��";"������롣
        Dim errorNum As Integer:   errorNum = CountNonEmptyElements(rngFieldErrors)
        Dim c As Range
        For Each c In rngFieldErrors
            If c.value <> "" Then
                ' �ե������̾
                Dim errorField As String: errorField = c.Offset(startRange.row - rngCurrent.row - 1, 0).value
                ' ���顼������
                Dim errorCode As String: errorCode = c.value
                ' ���Σ�
                Dim fieldErrorCode As String: fieldErrorCode = """" & formTagL & """, """ & errorField & """, """ & errorCode & """"
                ' ���Σ�
                Dim andExpect As String: andExpect = "                .andExpect(model().attributeHasFieldErrorCode(" & fieldErrorCode & "))"
                        
                ' ��λȽ��ʺǸ��";"���դ����
                If errorCount + 1 = errorNum Then
                    andExpect = andExpect & ";"
                End If
                
                ' ����
                Print #fileNumber, andExpect
                errorCount = errorCount + 1
                
            End If
        Next
    End If
End Sub


Private Function FindNonEmptyCell(rngCurrent As Range) As String

    If rngCurrent.Offset(0, TestTypePos).value <> "" Then
        FindNonEmptyCell = rngCurrent.Offset(0, TestTypePos).value
        Exit Function
    End If

    Dim row As Long: row = rngCurrent.row
    Dim value As String
    Dim index As Integer
    
    ' ���Ǥʤ����뤬���Ĥ���ޤǥ롼��
    Do While row > 1 ' 1���ܤޤǤ�����
        value = rngCurrent.Offset(-index, TestTypePos).value
        
        If Not IsEmpty(value) And value <> "" Then
            FindNonEmptyCell = value
            Exit Function
        End If
        
        ' ��ľ�ιԤ˰�ư
        row = row - 1
        index = index + 1
    Loop
    
    FindNonEmptyCell = Null
End Function


Private Function GetTestFunctionName(c As Range) As String
    Dim TestType As String: TestType = FindNonEmptyCell(c)
    Dim testItem As String: testItem = c.Offset(0, -1).value
    GetTestFunctionName = TestType & "_" & Format(c.Offset(0, -3).value, "000") & "_" & testItem
End Function

Private Function getClassName() As String
    getClassName = Range("d4").value
    getClassName = UCase(Left(getClassName, 1)) & Mid(getClassName, 2)
End Function


' ͭ�������ǿ��򥫥����
Private Function CountNonNullElements(rowRange As Range) As Integer
    Dim count As Integer: count = 0
    Dim c As Range
    For Each c In rowRange
        If c.value <> "null" Then
            count = count + 1
        End If
    Next
    
    CountNonNullElements = count
End Function

'' ɽ���ͤ����
'Private Function GetDisplayValue(value As Variant) As Variant
'    Dim result(1) As Variant
'
'    Select Case value
'        Case "����": result(0) = "LocalDate.now().minusDays(1).toString()": result(1) = 0
'        Case "����": result(0) = "LocalDate.now().toString()": result(1) = 0
'        Case "����": result(0) = "LocalDate.now().plusDays(1).toString()": result(1) = 0
'        Case "������": result(0) = "LocalDate.now().plusDays(2).toString()": result(1) = 0
'        Case "��������": result(0) = "LocalDate.now().plusDays(3).toString()": result(1) = 0
'        Case Else
'            If LCase(value) = "null" Then
'                result(0) = "null": result(1) = 2
'            ElseIf IsDate(value) Then
'                result(0) = Format(value, "yyyy-mm-dd"): result(1) = 1
'            Else
'                result(0) = value: result(1) = 1
'            End If
'    End Select
'
'    GetDisplayValue = result
'End Function
Private Function convertValue(value As String) As Variant
    Dim result As String

    Select Case value
    '
'        Case "����": result = "LocalDate.now().minusDays(1).toString()"
'        Case "����": result = "LocalDate.now().toString()"
'        Case "����": result = "LocalDate.now().plusDays(1).toString()"
'        Case "������": result = "LocalDate.now().plusDays(2).toString()"
'        Case "��������": result = "LocalDate.now().plusDays(3).toString()"
        Case "����": result = "java.sql.Date.valueOf(LocalDate.now().minusDays(1).toString())"
        Case "����": result = "java.sql.Date.valueOf(LocalDate.now().toString())"
        Case "����": result = "java.sql.Date.valueOf(LocalDate.now().plusDays(1).toString())"
        Case "������": result = "java.sql.Date.valueOf(LocalDate.now().plusDays(2).toString())"
        Case "��������": result = "java.sql.Date.valueOf(LocalDate.now().plusDays(3).toString())"
        Case Else
            If LCase(value) = "null" Then
                result = "null"
            ElseIf IsDate(value) Then
                result = Format(value, "yyyy-mm-dd")
                result = "java.sql.Date.valueOf(" & Chr(34) & result & Chr(34) & ")"
            Else
                result = value
                result = Chr(34) & result & Chr(34)
            End If
    End Select

    convertValue = result
End Function

' �ե�������ͤ����
Private Function GetCurrentFieldValues(rowIndex As Integer) As Variant
    Dim r1 As Range, r2 As Range
    Set r1 = startRange.Offset(rowIndex, 0)
    Set r2 = startRange.Offset(rowIndex, fieldCount - 1)
    GetCurrentFieldValues = sheet.Range(r1.Address & ":" & r2.Address).value
End Function

' �ե������̾�����
Private Function GetFieldName(columnIndex As Integer) As String

    GetFieldName = startRange.Offset(-1, columnIndex - startRange.row).value
End Function


'Private Function getReturnHtml(rowIndex As Integer) As String
'    'getReturnHtml = startRange.Offset(rowIndex - startRange.row, srcRangeSelection.Columns.count + 1).value
'    getReturnHtml = startRange.Offset(0, srcRangeSelection.Columns.count + 1).value
'End Function

Private Function isNormalTermination(rngCurrent As Range) As Boolean

    'Dim testcase As String: testcase = FindNonEmptyCell(startRange.Offset(c.row, -2))
    Dim testcase As String: testcase = FindNonEmptyCell(rngCurrent)
    If testcase = "�����" Then
      isNormalTermination = True
    Else
      isNormalTermination = False
    End If

End Function

Private Function getFieldErrors(rngCurrent As Range) As String

    
    Dim fields As String: fields = gerCurrentErrors(rngCurrent)
    
    getFieldErrors = Chr(34) + formTagL + Chr(34) + "," + fields
    
End Function

Function gerCurrentErrors(rngCurrent As Range) As String
    
    Dim ret As String
    
    ' errorCode�ιԡ�Min, Size�ʤɽ񤫤�Ƥ��륻���ϰϡˤ����
    Dim currentErrorFieldRange As Range: Set currentErrorFieldRange = getFieldErrorRange(rngCurrent)
    
    ' ���顼�����ä���Min, Size�ʤɽ񤫤�Ƥ���˥ե������̾���֤�
    Dim c As Range
    For Each c In currentErrorFieldRange
        
        If c.value <> "" Then
            If ret <> "" Then
                ret = ret & ","
            End If
            ret = ret & """" & c.Offset(startRange.row - rngCurrent.row - 1, 0).value & """"
        End If
    Next
    gerCurrentErrors = ret
End Function

'Private Function getFieldErrorValues(rowIndex) As String
'    Var r1 = this.startRange.Offset(rowIndex, this.offsetErrorCode)
'    Var r2 = this.startRange.Offset(rowIndex, (this.offsetErrorCode + this.fieldCount) - 1)
'    Var currentErrorFeildRange = this.sheet.getRange(r1.getA1Notation() + ":" + r2.getA1Notation())
'    return currentErrorFeildRange.getValues()
'End Function

Private Function getFieldErrorRange(rngCurrent As Range) As Range
    
    ' ���顼��������ܤ��֤���
    ' errorCode�ιԡ�Min, Size�ʤɽ񤫤�Ƥ��륻���ϰϡˤ��֤���
    Dim fields As Range: Set fields = Range(srcRangeSelection.Rows(rngCurrent.row).Address)
    Dim r1 As Range: Set r1 = rngCurrent.Offset(0, fields.Columns.count + FixedFiledNum)
    Dim r2 As Range: Set r2 = rngCurrent.Offset(0, fields.Columns.count + FixedFiledNum + fields.Columns.count - 1)
    Set getFieldErrorRange = Range(r1.Address & ":" & r2.Address)
End Function


Private Function CountNonEmptyElements(rngFieldErrors As Range) As Integer
    Dim count As Integer
    Dim c As Range
    For Each c In rngFieldErrors
        If c.value <> "" Then
            count = count + 1
        End If
    Next
    CountNonEmptyElements = count
End Function
