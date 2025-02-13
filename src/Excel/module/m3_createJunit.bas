Option Explicit
Private writeCurrent As Integer

' ¥Ø¥Ã¥ÀºîÀ®ÍÑ
Private fileNumber As Integer
Private endpoint As String
Private functionName As String
Private methodType As String
Private formTagL As String
Private formTagU As String

' ¥»¥ë
Private srcRangeSelection As Range
Private srcRowSelection As Range
Private startRange As Range

' ½ÐÎÏ¥Õ¥¡¥¤¥ë¥Õ¥ë¥Ñ¥¹
Private filePath As String

' ¥ê¥¹¥È
Private minList As Collection ' ItemInfo
Private maxList As Collection ' ItemInfo
Private requiredList As Collection ' ItemInfo
Private enumList As Collection ' ItemInfo
Private defaultList As Collection ' ItemInfo

' ³µÍ×¡§²¼µ­¤Î¤è¤¦¤Ê¥Æ¥¹¥È¤òºîÀ®¤¹¤ë
'    @Nested
'    @DisplayName("/regist-review")
'    class registReview {
'        @Test
'        public void Àµ¾ï·Ï_ºÇ¾®() throws Exception {
' ­¡         mockMvc.perform(POST("/regist-review")
' ­¢             .param("restaurantId", "0")
' ¢¬             .param("userId", "aaa")
' ¢¬             .param("visitDate", "2025-02-12")
' ¢¬             .param("rating", "0")
' ¢¬             .param("comment", ""))
' ­£             .andExpect(status().isOk())
' ­¤             .andExpect(view().name(""))
' ­¥             .andExpect(model().attributeHasFieldErrors("reviewRegistForm","restaurantId","userId","rating","comment"))
' ­¦             .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","restaurantId","Min"))
' ¢¬             .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
' ¢¬         ¡¡  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
' ¢¬             .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));
'        }
'    }
Public Sub CreateTestCode()
    
    Call Initialize
    Dim i As Integer
    Dim testCaseValues As Variant
    testCaseValues = srcRangeSelection.value  ' ¥Æ¥¹¥È¥±¡¼¥¹¥Ç¡¼¥¿
    ' ¥¯¥é¥¹ÄêµÁ
    CreateClassDefine
    
    Dim c As Range
    
    For Each c In srcRowSelection
    'For i = LBound(testCaseValues, 1) To UBound(testCaseValues, 1)
        ' ¥Æ¥¹¥È´Ø¿ôºîÀ®¢·¥ê¥¯¥¨¥¹¥È
        ' ¥Æ¥¹¥È´Ø¿ô
        Call CreateFuncRequest(c)
        
        ' ¥ê¥¯¥¨¥¹¥È¥Ñ¥é¥á¡¼¥¿
        Call CreateRequestParam1(c)
        
        ' ¥Ç¥Õ¥©¥ë¥È¥â¥Ã¥¯
        Call CreateMock(c)
        
        ' post get¥á¥½¥Ã¥É
        Call CreateMetod(c)
        
        ' ¥ê¥¯¥¨¥¹¥È¤Î¥Ñ¥é¥á¡¼¥¿ÀßÄê¤òºîÀ®
        ' Call CreateRequestParam2(c)
        Call CreateRequestParam3(c)
        
        ' HTTP¥¹¥Æ¡¼¥¿¥¹
        Call CreateHttpStatus
        
        ' HTMLÌ¾
        Debug.Print c.Offset(0, -2).value & ":" & c.Offset(0, -1).value
        Call CreateReturnHtmlName(c)
        
        ' ¥¨¥é¡¼¾ðÊó
        Call CreateErrorInfo(c)
        
        ' ¥Æ¥¹¥È´Ø¿ôÄêµÁ½ªÎ»
        Print #fileNumber, "        }"
        Print #fileNumber, ""
    Next
    
    ' ¥¯¥é¥¹ÄêµÁ½ªÎ»
    Print #fileNumber, "    }"
    
    Call Terminate
End Sub

Private Sub Initialize()
    
    ' header¥»¥ë
    endpoint = Range("D2").value        '¥¨¥ó¥É¥Ý¥¤¥ó¥È
    methodType = Range("D3").value      'POST¤È¤«
    functionName = Range("D4").value    '¥¯¥é¥¹Ì¾
    
    'FormÌ¾
    formTagU = Range("D5").value
    formTagL = LCase(Left(formTagU, 1)) & Mid(formTagU, 2)
    

    ' input¥»¥ë
    Set srcRangeSelection = ActiveWindow.RangeSelection
    Set srcRowSelection = Range(srcRangeSelection.Columns(1).Address)
    Set startRange = ActiveCell
    
    '½ÐÎÏ¥Õ¥¡¥¤¥ë
    Dim folderPath As String: folderPath = ActiveWorkbook.Path & "\test\"
    ' ¥Õ¥©¥ë¥À¤¬Ìµ¤±¤ì¤ÐºîÀ®
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
    filePath = folderPath & getClassName & "Test.java"
    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    

End Sub

Private Sub Terminate()
    Close #fileNumber
    
    If MsgBox("ºîÀ®¤·¤¿¥Õ¥¡¥¤¥ë¤ò³«¤¯¡©", vbYesNo) = vbYes Then
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

' ³µÍ×¡§¥ê¥¯¥¨¥¹¥È¤ÎÁ÷¿®ÉôÊ¬¤òºîÀ®¤¹¤ë
'
' Îã¡¡¡§²¼µ­¤Î­¡¤òºîÀ®¤¹¤ë
' ­¡   @Test
' ­¡   public void Àµ¾ï·Ï_ºÇ¾®() throws Exception {
' ­¡       // ¥ê¥¯¥¨¥¹¥È
' ­¡       // ¥Ç¥Õ¥©¥ë¥È¥â¥Ã¥¯
' ­¡       // defaultMock();
' ­¡       mockMvc.perform(POST("/regist-review")
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
'      ¡¡  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
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
    
    
    Print #fileNumber, "            // ================== ¥ê¥¯¥¨¥¹¥È =================="
    Print #fileNumber, "            " + formTagU & " " & formTagL & " = new " & formTagU & "();"
    
    Set fields = Range(srcRangeSelection.Rows(rngCurrent.row - startRange.row + 1).Address)
    
    ' ¥ë¡¼¥×½èÍý
    Dim c As Range
    Dim rowRange As Range
    For Each c In fields
        ' ¥Ñ¥é¥á¡¼¥¿À¸À®
        Dim fieldName As String: fieldName = c.Offset(startRange.row - rngCurrent.row - 1, 0).value
        Dim fieldValue As String: fieldValue = convertValue(c.value)
        
        If c.value = "null" Then
            GoTo ContinueLoop
        End If
        paramLine = "            " & formTagL & ".set" & UCase(Left(fieldName, 1)) & Mid(fieldName, 2) & "(" & fieldValue & ");"
        ' ¥Õ¥¡¥¤¥ë½ñ¤­¹þ¤ß
        Print #fileNumber, "    " + paramLine

ContinueLoop:
    Next



End Sub


Private Sub CreateMock(c As Range)
    Print #fileNumber, "            // ================== ¥â¥Ã¥¯ =================="
    Print #fileNumber, "            defaultMock();"
    Print #fileNumber, ""
End Sub

Private Sub CreateMetod(c As Range)
    Print #fileNumber, "            // ================== ¼Â¹Ô =================="
    Print #fileNumber, "            mockMvc.perform(" & methodType & "(""" & endpoint & """)"
End Sub


' ³µÍ×¡§¥ê¥¯¥¨¥¹¥È¤Î¥Ñ¥é¥á¡¼¥¿ÉôÊ¬¤òºîÀ®¤¹¤ë
'
' Îã¡¡¡§²¼µ­¤Î­¢¤òºîÀ®¤¹¤ë
'      mockMvc.perform(POST("/regist-review")
' ­¢       .param("restaurantId", "0")
' ­¢       .param("userId", "aaa")
' ­¢       .param("visitDate", "2025-02-12")
' ­¢       .param("rating", "0")
' ­¢       .param("comment", ""))
'          .andExpect(status().isOk())
'          .andExpect(view().name(""))
'          .andExpect(model().attributeHasFieldErrors("reviewRegistForm","restaurantId","userId","rating","comment"))
'          .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","restaurantId","Min"))
'  ¡¡      .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
'      ¡¡  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
'  ¡¡      .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));

Private Sub CreateRequestParam2(rngCurrent As Range)
    Dim currentFieldValues As Variant
    Dim paramMaxCount As Integer
    Dim paramCount As Integer
    Dim j As Integer
    Dim paramLine As String
    Dim ret As Variant
    Dim fields As Range
    
    
    Set fields = Range(srcRangeSelection.Rows(rngCurrent.row - startRange.row + 1).Address)
    
    ' Í­¸ú¤ÊÍ×ÁÇ¿ô¤ò¥«¥¦¥ó¥È
    paramMaxCount = CountNonNullElements(fields)
    paramCount = 0
    
    ' ¥ë¡¼¥×½èÍý
    Dim c As Range
    Dim rowRange As Range
    
    For Each c In fields
        
        ' ¥Ñ¥é¥á¡¼¥¿À¸À®
        Dim fieldName As String: fieldName = c.Offset(startRange.row - rngCurrent.row - 1, 0).value
        Dim fieldValue As String: fieldValue = convertValue(c.value)
        If c.value = "null" Then
            GoTo ContinueLoop
        End If
        paramLine = "                .param(" & Chr(34) & fieldName & Chr(34) & ", " & fieldValue & ")"
        
        ' ºÇ¸å¤Î¥Ñ¥é¥á¡¼¥¿½èÍý
        paramCount = paramCount + 1
        If paramCount = paramMaxCount Then
            paramLine = paramLine & ")"
        End If
        
        ' ¥Õ¥¡¥¤¥ë½ñ¤­¹þ¤ß
        Print #fileNumber, "    " + paramLine

ContinueLoop:
    Next
End Sub

' //                     .flashAttr("taskRegistForm", form)) // ¥Õ¥©¡¼¥à¥ª¥Ö¥¸¥§¥¯¥È¤òÁ÷¿®
Private Sub CreateRequestParam3(rngCurrent As Range)
    
    ' ¥Õ¥¡¥¤¥ë½ñ¤­¹þ¤ß
    Print #fileNumber, "                .flashAttr(" & Chr(34) & formTagL & Chr(34) & " , " & formTagL & ")) // ¥Õ¥©¡¼¥à¥ª¥Ö¥¸¥§¥¯¥È"

End Sub

' ³µÍ×¡§¥ê¥¯¥¨¥¹¥È¤ÎHTTP¥¹¥Æ¡¼¥¿¥¹¤òºîÀ®¤¹¤ë
'
' Îã¡¡¡§²¼µ­¤Î­£¤òºîÀ®¤¹¤ë
'      mockMvc.perform(POST("/regist-review")
'          .param("restaurantId", "0")
'          .param("userId", "aaa")
'          .param("visitDate", "2025-02-12")
'          .param("rating", "0")
'          .param("comment", ""))
' ­£       .andExpect(status().isOk())
'          .andExpect(view().name(""))
'          .andExpect(model().attributeHasFieldErrors("reviewRegistForm","restaurantId","userId","rating","comment"))
'          .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","restaurantId","Min"))
' ¡¡       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
'      ¡¡  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
' ¡¡       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));
Private Sub CreateHttpStatus()
    Print #fileNumber, "                .andExpect(status().isOk())"
End Sub

' ³µÍ×¡§htmlÌ¾¤Î²Õ½ê¤òºîÀ®¤¹¤ë
'
' Îã¡¡¡§²¼µ­¤Î­¤¤òºîÀ®¤¹¤ë
'      mockMvc.perform(POST("/regist-review")
'          .param("restaurantId", "0")
'          .param("userId", "aaa")
'          .param("visitDate", "2025-02-12")
'          .param("rating", "0")
'          .param("comment", ""))
'          .andExpect(status().isOk())
' ­¤       .andExpect(view().name("task-regist-confirm"))
'          .andExpect(model().attributeHasFieldErrors("reviewRegistForm","restaurantId","userId","rating","comment"))
'          .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","restaurantId","Min"))
' ¡¡       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
'      ¡¡  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
' ¡¡       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));
Private Sub CreateReturnHtmlName(rngCurrent As Range)
    Print #fileNumber, "                .andExpect(view().name(""" & rngCurrent.Offset(0, srcRangeSelection.Columns.count + 1).value & """))"
End Sub


' ³µÍ×¡§¥¨¥é¡¼¸¡¾Ú¤ÎºîÀ®
'
' Îã¡¡¡§²¼µ­¤Î­¥¤È­¦¤òºîÀ®¤¹¤ë
'      mockMvc.perform(POST("/regist-review")
'          .param("restaurantId", "0")
'          .param("userId", "aaa")
'          .param("visitDate", "2025-02-12")
'          .param("rating", "0")
'          .param("comment", ""))
'          .andExpect(status().isOk())
'          .andExpect(view().name(""))
' ­¥       .andExpect(model().attributeHasFieldErrors("reviewRegistForm","restaurantId","userId","rating","comment"))
' ­¦       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","restaurantId","Min"))
' ¢¬       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
' ¢¬   ¡¡  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
' ¢¬       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));
Private Sub CreateErrorInfo(rngCurrent As Range)
    If isNormalTermination(rngCurrent) Then
        Print #fileNumber, "                .andExpect(model().hasNoErrors());"
    Else
        Dim fieldErrors As String
        fieldErrors = getFieldErrors(rngCurrent)
        ' ­¥¤òºîÀ®
        Print #fileNumber, "                .andExpect(model().attributeHasFieldErrors(" & fieldErrors & "))"
        
        ' ­¦¤òºîÀ®
        Dim rngFieldErrors As Range: Set rngFieldErrors = getFieldErrorRange(rngCurrent)
        Dim errorCount As Integer: errorCount = 0   ' ºÇ½ªÈ½Äê¤ËÍøÍÑ¡£¥¨¥é¡¼ºÇ¸å¤Ï";"¤ÇÄù¤á¤ë¡£
        Dim errorNum As Integer:   errorNum = CountNonEmptyElements(rngFieldErrors)
        Dim c As Range
        For Each c In rngFieldErrors
            If c.value <> "" Then
                ' ¥Õ¥£¡¼¥ë¥ÉÌ¾
                Dim errorField As String: errorField = c.Offset(startRange.row - rngCurrent.row - 1, 0).value
                ' ¥¨¥é¡¼¥³¡¼¥É
                Dim errorCode As String: errorCode = c.value
                ' ¹çÂÎ£±
                Dim fieldErrorCode As String: fieldErrorCode = """" & formTagL & """, """ & errorField & """, """ & errorCode & """"
                ' ¹çÂÎ£²
                Dim andExpect As String: andExpect = "                .andExpect(model().attributeHasFieldErrorCode(" & fieldErrorCode & "))"
                        
                ' ½ªÎ»È½Äê¡ÊºÇ¸å¤Ï";"¤òÉÕ¤±¤ë¡Ë
                If errorCount + 1 = errorNum Then
                    andExpect = andExpect & ";"
                End If
                
                ' ½ÐÎÏ
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
    
    ' ¶õ¤Ç¤Ê¤¤¥»¥ë¤¬¸«¤Ä¤«¤ë¤Þ¤Ç¥ë¡¼¥×
    Do While row > 1 ' 1¹ÔÌÜ¤Þ¤Ç¤ËÀ©¸Â
        value = rngCurrent.Offset(-index, TestTypePos).value
        
        If Not IsEmpty(value) And value <> "" Then
            FindNonEmptyCell = value
            Exit Function
        End If
        
        ' °ì¤Ä¾å¤Î¹Ô¤Ë°ÜÆ°
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


' Í­¸ú¤ÊÍ×ÁÇ¿ô¤ò¥«¥¦¥ó¥È
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

'' É½¼¨ÃÍ¤ò¼èÆÀ
'Private Function GetDisplayValue(value As Variant) As Variant
'    Dim result(1) As Variant
'
'    Select Case value
'        Case "ºòÆü": result(0) = "LocalDate.now().minusDays(1).toString()": result(1) = 0
'        Case "º£Æü": result(0) = "LocalDate.now().toString()": result(1) = 0
'        Case "ÌÀÆü": result(0) = "LocalDate.now().plusDays(1).toString()": result(1) = 0
'        Case "ÌÀ¸åÆü": result(0) = "LocalDate.now().plusDays(2).toString()": result(1) = 0
'        Case "ÌÀ¡¹¸åÆü": result(0) = "LocalDate.now().plusDays(3).toString()": result(1) = 0
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
'        Case "ºòÆü": result = "LocalDate.now().minusDays(1).toString()"
'        Case "º£Æü": result = "LocalDate.now().toString()"
'        Case "ÌÀÆü": result = "LocalDate.now().plusDays(1).toString()"
'        Case "ÌÀ¸åÆü": result = "LocalDate.now().plusDays(2).toString()"
'        Case "ÌÀ¡¹¸åÆü": result = "LocalDate.now().plusDays(3).toString()"
        Case "ºòÆü": result = "java.sql.Date.valueOf(LocalDate.now().minusDays(1).toString())"
        Case "º£Æü": result = "java.sql.Date.valueOf(LocalDate.now().toString())"
        Case "ÌÀÆü": result = "java.sql.Date.valueOf(LocalDate.now().plusDays(1).toString())"
        Case "ÌÀ¸åÆü": result = "java.sql.Date.valueOf(LocalDate.now().plusDays(2).toString())"
        Case "ÌÀ¡¹¸åÆü": result = "java.sql.Date.valueOf(LocalDate.now().plusDays(3).toString())"
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

' ¥Õ¥£¡¼¥ë¥ÉÃÍ¤ò¼èÆÀ
Private Function GetCurrentFieldValues(rowIndex As Integer) As Variant
    Dim r1 As Range, r2 As Range
    Set r1 = startRange.Offset(rowIndex, 0)
    Set r2 = startRange.Offset(rowIndex, fieldCount - 1)
    GetCurrentFieldValues = sheet.Range(r1.Address & ":" & r2.Address).value
End Function

' ¥Õ¥£¡¼¥ë¥ÉÌ¾¤ò¼èÆÀ
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
    If testcase = "Àµ¾ï·Ï" Then
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
    
    ' errorCode¤Î¹Ô¡ÊMin, Size¤Ê¤É½ñ¤«¤ì¤Æ¤¤¤ë¥»¥ëÈÏ°Ï¡Ë¤ò¼èÆÀ
    Dim currentErrorFieldRange As Range: Set currentErrorFieldRange = getFieldErrorRange(rngCurrent)
    
    ' ¥¨¥é¡¼¤¬¤¢¤Ã¤¿¡ÊMin, Size¤Ê¤É½ñ¤«¤ì¤Æ¤¤¤ë¡Ë¥Õ¥£¡¼¥ë¥ÉÌ¾¤òÊÖ¤¹
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
    
    ' ¥¨¥é¡¼¤¬¤¢¤ë¹àÌÜ¤òÊÖ¤¹¡£
    ' errorCode¤Î¹Ô¡ÊMin, Size¤Ê¤É½ñ¤«¤ì¤Æ¤¤¤ë¥»¥ëÈÏ°Ï¡Ë¤òÊÖ¤¹¡£
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
