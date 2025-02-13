Attribute VB_Name = "m3_createJunit"
Option Explicit
Private writeCurrent As Integer

' �w�b�_�쐬�p
Private fileNumber As Integer
Private endpoint As String
Private functionName As String
Private methodType As String
Private formTag As String

' �Z��
Private srcRangeSelection As Range
Private srcRowSelection As Range
Private startRange As Range


' ���X�g
Private minList As Collection ' ItemInfo
Private maxList As Collection ' ItemInfo
Private requiredList As Collection ' ItemInfo
Private enumList As Collection ' ItemInfo
Private defaultList As Collection ' ItemInfo

' �T�v�F���L�̂悤�ȃe�X�g���쐬����
'    @Nested
'    @DisplayName("/regist-review")
'    class registReview {
'        @Test
'        public void ����n_�ŏ�() throws Exception {
' �@         mockMvc.perform(POST("/regist-review")
' �A             .param("restaurantId", "0")
' ��             .param("userId", "aaa")
' ��             .param("visitDate", "2025-02-12")
' ��             .param("rating", "0")
' ��             .param("comment", ""))
' �B             .andExpect(status().isOk())
' �C             .andExpect(view().name(""))
' �D             .andExpect(model().attributeHasFieldErrors("reviewRegistForm","restaurantId","userId","rating","comment"))
' �E             .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","restaurantId","Min"))
' ��             .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
' ��         �@  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
' ��             .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));
'        }
'    }
Public Sub CreateTestCode()
    
    Call Initialize
    Dim i As Integer
    Dim testCaseValues As Variant
    testCaseValues = srcRangeSelection.value  ' �e�X�g�P�[�X�f�[�^
    ' �N���X��`
    CreateClassDefine
    
    Dim c As Range
    
    For Each c In srcRowSelection
    'For i = LBound(testCaseValues, 1) To UBound(testCaseValues, 1)
        ' �e�X�g�֐��쐬�`���N�G�X�g
        ' �e�X�g�֐�
        Call CreateFuncRequest(c)
        
        ' ���N�G�X�g�p�����[�^
        Call CreateRequestParam1(c)
        
        ' �f�t�H���g���b�N
        Call CreateMock(c)
        
        ' post get���\�b�h
        Call CreateMetod(c)
        
        ' ���N�G�X�g�̃p�����[�^�ݒ���쐬
        Call CreateRequestParam2(c)
        
        ' HTTP�X�e�[�^�X
        Call CreateHttpStatus
        
        ' HTML��
        Call CreateReturnHtmlName(c.row)
        
        ' �G���[���
        Call CreateErrorInfo(c)
        
        ' �e�X�g�֐���`�I��
        Print #fileNumber, "        }"
        Print #fileNumber, ""
    Next
    
    ' �N���X��`�I��
    Print #fileNumber, "    }"
    
    Call Terminate
End Sub

Private Sub Initialize()
    
    ' header�Z��
    endpoint = Range("D2").value        '�G���h�|�C���g
    methodType = Range("D3").value      'POST�Ƃ�
    functionName = Range("D4").value    '�N���X��
    
    'Form��
    formTag = Range("D5").value
    formTag = LCase(Left(formTag, 1)) & Mid(formTag, 2)
    

    ' input�Z��
    Set srcRangeSelection = ActiveWindow.RangeSelection
    Set srcRowSelection = Range(srcRangeSelection.Columns(1).Address)
    Set startRange = ActiveCell
    
    '�o�̓t�@�C��
    Dim filePath As String: filePath = ActiveWorkbook.Path & "\" & getClassName & "Test.java"
    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    

End Sub

Private Sub Terminate()
    Close #fileNumber
End Sub

Private Sub CreateClassDefine()
    Print #fileNumber, Space(4) & "@Nested"
    Print #fileNumber, Space(4) & "@DisplayName(""" & endpoint & """)"
    Print #fileNumber, Space(4) & "class " & functionName & " {"
End Sub

' �T�v�F���N�G�X�g�̑��M�������쐬����
'
' ��@�F���L�̇@���쐬����
' �@   @Test
' �@   public void ����n_�ŏ�() throws Exception {
' �@       // ���N�G�X�g
' �@       // �f�t�H���g���b�N
' �@       // defaultMock();
' �@       mockMvc.perform(POST("/regist-review")
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
'      �@  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
'          .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));
Private Sub CreateFuncRequest(c As Range)
    Print #fileNumber, "        @Test"
    Print #fileNumber, "        public void " & GetTestFunctionName(c) & "() throws Exception {"
    Print #fileNumber, ""
End Sub

Private Sub CreateRequestParam1(c As Range)
    Print #fileNumber, "            // ================== ���N�G�X�g =================="
    Print #fileNumber, ""
End Sub


Private Sub CreateMock(c As Range)
    Print #fileNumber, "            // ================== ���b�N =================="
    Print #fileNumber, "            defaultMock();"
    Print #fileNumber, ""
End Sub

Private Sub CreateMetod(c As Range)
    Print #fileNumber, "            // ================== ���s =================="
    Print #fileNumber, "            mockMvc.perform(" & methodType & "(""" & endpoint & """)"
End Sub


' �T�v�F���N�G�X�g�̃p�����[�^�������쐬����
'
' ��@�F���L�̇A���쐬����
'      mockMvc.perform(POST("/regist-review")
' �A       .param("restaurantId", "0")
' �A       .param("userId", "aaa")
' �A       .param("visitDate", "2025-02-12")
' �A       .param("rating", "0")
' �A       .param("comment", ""))
'          .andExpect(status().isOk())
'          .andExpect(view().name(""))
'          .andExpect(model().attributeHasFieldErrors("reviewRegistForm","restaurantId","userId","rating","comment"))
'          .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","restaurantId","Min"))
'  �@      .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
'      �@  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
'  �@      .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));

Private Sub CreateRequestParam2(rngCurrent As Range)
    Dim currentFieldValues As Variant
    Dim paramMaxCount As Integer
    Dim paramCount As Integer
    Dim j As Integer
    Dim paramLine As String
    Dim ret As Variant
    Dim fields As Range
    
    
    Set fields = Range(srcRangeSelection.Rows(rngCurrent.row - startRange.row + 1).Address)
    
    ' �L���ȗv�f�����J�E���g
    paramMaxCount = CountNonNullElements(fields)
    paramCount = 0
    
    ' ���[�v����
    'For j = LBound(currentFieldValues, 2) To UBound(currentFieldValues, 2)
    Dim c As Range
    Dim rowRange As Range
    
    For Each c In fields
        
        ' �p�����[�^����
        Dim fieldName As String: fieldName = c.Offset(startRange.row - rngCurrent.row - 1, 0).value
        Dim fieldValue As String: fieldValue = convertValue(c.value)
        If c.value = "null" Then
            GoTo ContinueLoop
        End If
        paramLine = "                .param(" & Chr(34) & fieldName & Chr(34) & ", " & fieldValue & ")"
        
        ' �Ō�̃p�����[�^����
        paramCount = paramCount + 1
        If paramCount = paramMaxCount Then
            paramLine = paramLine & ")"
        End If
        
        ' �t�@�C����������
        Print #fileNumber, "    " + paramLine

ContinueLoop:
    Next
End Sub

' �T�v�F���N�G�X�g��HTTP�X�e�[�^�X���쐬����
'
' ��@�F���L�̇B���쐬����
'      mockMvc.perform(POST("/regist-review")
'          .param("restaurantId", "0")
'          .param("userId", "aaa")
'          .param("visitDate", "2025-02-12")
'          .param("rating", "0")
'          .param("comment", ""))
' �B       .andExpect(status().isOk())
'          .andExpect(view().name(""))
'          .andExpect(model().attributeHasFieldErrors("reviewRegistForm","restaurantId","userId","rating","comment"))
'          .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","restaurantId","Min"))
' �@       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
'      �@  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
' �@       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));
Private Sub CreateHttpStatus()
    Print #fileNumber, "                .andExpect(status().isOk())"
End Sub

' �T�v�Fhtml���̉ӏ����쐬����
'
' ��@�F���L�̇C���쐬����
'      mockMvc.perform(POST("/regist-review")
'          .param("restaurantId", "0")
'          .param("userId", "aaa")
'          .param("visitDate", "2025-02-12")
'          .param("rating", "0")
'          .param("comment", ""))
'          .andExpect(status().isOk())
' �C       .andExpect(view().name(""))
'          .andExpect(model().attributeHasFieldErrors("reviewRegistForm","restaurantId","userId","rating","comment"))
'          .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","restaurantId","Min"))
' �@       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
'      �@  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
' �@       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));
Private Sub CreateReturnHtmlName(i As Integer)
    Print #fileNumber, "                .andExpect(view().name(""" & getReturnHtml(i) & """))"
End Sub


' �T�v�F�G���[���؂̍쐬
'
' ��@�F���L�̇D�ƇE���쐬����
'      mockMvc.perform(POST("/regist-review")
'          .param("restaurantId", "0")
'          .param("userId", "aaa")
'          .param("visitDate", "2025-02-12")
'          .param("rating", "0")
'          .param("comment", ""))
'          .andExpect(status().isOk())
'          .andExpect(view().name(""))
' �D       .andExpect(model().attributeHasFieldErrors("reviewRegistForm","restaurantId","userId","rating","comment"))
' �E       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","restaurantId","Min"))
' ��       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
' ��   �@  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
' ��       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));
Private Sub CreateErrorInfo(rngCurrent As Range)
    If isNormalTermination(rngCurrent) Then
        Print #fileNumber, "                .andExpect(model().hasNoErrors());"
    Else
        Dim fieldErrors As String
        fieldErrors = getFieldErrors(rngCurrent)
        ' �D���쐬
        Print #fileNumber, "                .andExpect(model().attributeHasFieldErrors(" & fieldErrors & "))"
        
        ' �E���쐬
        Dim rngFieldErrors As Range: Set rngFieldErrors = getFieldErrorRange(rngCurrent)
        Dim errorCount As Integer: errorCount = 0   ' �ŏI����ɗ��p�B�G���[�Ō��";"�Œ��߂�B
        Dim errorNum As Integer:   errorNum = CountNonEmptyElements(rngFieldErrors)
        Dim c As Range
        For Each c In rngFieldErrors
            If c.value <> "" Then
                ' �t�B�[���h��
                Dim errorField As String: errorField = c.Offset(startRange.row - rngCurrent.row - 1, 0).value
                ' �G���[�R�[�h
                Dim errorCode As String: errorCode = c.value
                ' ���̂P
                Dim fieldErrorCode As String: fieldErrorCode = """" & formTag & """, """ & errorField & """, """ & errorCode & """"
                ' ���̂Q
                Dim andExpect As String: andExpect = "                .andExpect(model().attributeHasFieldErrorCode(" & fieldErrorCode & "))"
                        
                ' �I������i�Ō��";"��t����j
                If errorCount + 1 = errorNum Then
                    andExpect = andExpect & ";"
                End If
                
                ' �o��
                Print #fileNumber, andExpect
                errorCount = errorCount + 1
                
            End If
        Next

'        Dim errorValues As Variant, errorCount As Integer, errorNum As Integer
'        errorValues = getFieldErrorValues(rngCurrent, i)
'        Dim errorNum As Integer:   errorNum = CountNonEmptyElements(errorValues)
'        errorCount = 0
'
'        Dim errorIndex As Integer
'        For errorIndex = LBound(errorValues, 2) To UBound(errorValues, 2)
'            If errorValues(1, errorIndex) <> "" Then
'                Dim errorField As String, fieldErrorCode As String
'                'errorField = GetErrorFieldName(i, errorIndex)
'                fieldErrorCode = """" & formTag & """, """ & errorField & """, """ & errorValues(1, errorIndex) & """"
'
'                Dim andExpect As String
'                andExpect = "                .andExpect(model().attributeHasFieldErrorCode(" & fieldErrorCode & "))"
'                If errorCount + 1 = errorNum Then andExpect = andExpect & ";"
'
'                Print #fileNumber, andExpect
'                errorCount = errorCount + 1
'            End If
'        Next errorIndex

    End If
End Sub


'Public Function FindNonEmptyCell(ByVal cell As Range) As Variant
'    Dim sheet As Worksheet
'    Dim row As Long
'    Dim column As Long
'    Dim value As Variant
'
'    ' �A�N�e�B�u�ȃZ����������V�[�g���擾
'    Set sheet = cell.Parent
'
'    ' ���݂̃Z���̍s�Ɨ���擾
'    row = cell.row
'    column = cell.column
'
'    ' ��łȂ��Z����������܂Ń��[�v
'    Do While row > 1 ' 1�s�ڂ܂łɐ���
'        value = sheet.Cells(row, column).value
'
'        If Not IsEmpty(value) And value <> "" Then
'            FindNonEmptyCell = value
'            Exit Function
'        End If
'
'        ' ���̍s�Ɉړ�
'        row = row - 1
'    Loop
'
'    FindNonEmptyCell = Null
'End Function

Private Function FindNonEmptyCell(rngCurrent As Range) As String

    If rngCurrent.Offset(0, TestTypePos).value <> "" Then
        FindNonEmptyCell = rngCurrent.Offset(0, TestTypePos).value
        Exit Function
    End If

    Dim row As Long: row = rngCurrent.row
    Dim value As String
    Dim index As Integer
    
    ' ��łȂ��Z����������܂Ń��[�v
    Do While row > 1 ' 1�s�ڂ܂łɐ���
        value = rngCurrent.Offset(-index, TestTypePos).value
        
        If Not IsEmpty(value) And value <> "" Then
            FindNonEmptyCell = value
            Exit Function
        End If
        
        ' ���̍s�Ɉړ�
        row = row - 1
        index = index + 1
    Loop
    
    FindNonEmptyCell = Null
End Function


Private Function GetTestFunctionName(c As Range) As String
    Dim TestType As String: TestType = FindNonEmptyCell(c)
    Dim testItem As String: testItem = c.Offset(0, -1).value
    GetTestFunctionName = TestType + "_" + testItem
End Function

Private Function getClassName() As String
    getClassName = Range("d4").value
    getClassName = UCase(Left(getClassName, 1)) & Mid(getClassName, 2)
End Function


' �L���ȗv�f�����J�E���g
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

'' �\���l���擾
'Private Function GetDisplayValue(value As Variant) As Variant
'    Dim result(1) As Variant
'
'    Select Case value
'        Case "���": result(0) = "LocalDate.now().minusDays(1).toString()": result(1) = 0
'        Case "����": result(0) = "LocalDate.now().toString()": result(1) = 0
'        Case "����": result(0) = "LocalDate.now().plusDays(1).toString()": result(1) = 0
'        Case "�����": result(0) = "LocalDate.now().plusDays(2).toString()": result(1) = 0
'        Case "���X���": result(0) = "LocalDate.now().plusDays(3).toString()": result(1) = 0
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
        Case "���": result = "LocalDate.now().minusDays(1).toString()"
        Case "����": result = "LocalDate.now().toString()"
        Case "����": result = "LocalDate.now().plusDays(1).toString()"
        Case "�����": result = "LocalDate.now().plusDays(2).toString()"
        Case "���X���": result = "LocalDate.now().plusDays(3).toString()"
        Case Else
            If LCase(value) = "null" Then
                result = "null"
            ElseIf IsDate(value) Then
                result = Format(value, "yyyy-mm-dd")
                result = Chr(34) & result & Chr(34)
            Else
                result = value
                result = Chr(34) & result & Chr(34)
            End If
    End Select

    convertValue = result
End Function

' �t�B�[���h�l���擾
Private Function GetCurrentFieldValues(rowIndex As Integer) As Variant
    Dim r1 As Range, r2 As Range
    Set r1 = startRange.Offset(rowIndex, 0)
    Set r2 = startRange.Offset(rowIndex, fieldCount - 1)
    GetCurrentFieldValues = sheet.Range(r1.Address & ":" & r2.Address).value
End Function

' �t�B�[���h�����擾
Private Function GetFieldName(columnIndex As Integer) As String

    GetFieldName = startRange.Offset(-1, columnIndex - startRange.row).value
End Function


Private Function getReturnHtml(rowIndex As Integer) As String
    'getReturnHtml = startRange.Offset(rowIndex - startRange.row, srcRangeSelection.Columns.count + 1).value
    getReturnHtml = startRange.Offset(0, srcRangeSelection.Columns.count + 1).value
End Function

Private Function isNormalTermination(rngCurrent As Range) As Boolean

    'Dim testcase As String: testcase = FindNonEmptyCell(startRange.Offset(c.row, -2))
    Dim testcase As String: testcase = FindNonEmptyCell(rngCurrent)
    If testcase = "����n" Then
      isNormalTermination = True
    Else
      isNormalTermination = False
    End If

End Function

Private Function getFieldErrors(rngCurrent As Range) As String

    
    Dim fields As String: fields = gerCurrentErrors(rngCurrent)
    
    getFieldErrors = Chr(34) + formTag + Chr(34) + "," + fields
    
End Function

Function gerCurrentErrors(rngCurrent As Range) As String
    
    Dim ret As String
    
    ' errorCode�̍s�iMin, Size�ȂǏ�����Ă���Z���͈́j���擾
    Dim currentErrorFieldRange As Range: Set currentErrorFieldRange = getFieldErrorRange(rngCurrent)
    
    ' �G���[���������iMin, Size�ȂǏ�����Ă���j�t�B�[���h����Ԃ�
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
    
    ' �G���[�����鍀�ڂ�Ԃ��B
    ' errorCode�̍s�iMin, Size�ȂǏ�����Ă���Z���͈́j��Ԃ��B
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
