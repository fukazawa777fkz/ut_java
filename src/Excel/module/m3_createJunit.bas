Option Explicit
Private writeCurrent As Integer

' ヘッダ作成用
Private fileNumber As Integer
Private endpoint As String
Private functionName As String
Private methodType As String
Private formTagL As String
Private formTagU As String

' セル
Private srcRangeSelection As Range
Private srcRowSelection As Range
Private startRange As Range

' 出力ファイルフルパス
Private filePath As String

' リスト
Private minList As Collection ' ItemInfo
Private maxList As Collection ' ItemInfo
Private requiredList As Collection ' ItemInfo
Private enumList As Collection ' ItemInfo
Private defaultList As Collection ' ItemInfo

' 概要：下記のようなテストを作成する
'    @Nested
'    @DisplayName("/regist-review")
'    class registReview {
'        @Test
'        public void 正常系_最小() throws Exception {
' ��         mockMvc.perform(POST("/regist-review")
' ��             .param("restaurantId", "0")
' ↑             .param("userId", "aaa")
' ↑             .param("visitDate", "2025-02-12")
' ↑             .param("rating", "0")
' ↑             .param("comment", ""))
' ��             .andExpect(status().isOk())
' ��             .andExpect(view().name(""))
' ��             .andExpect(model().attributeHasFieldErrors("reviewRegistForm","restaurantId","userId","rating","comment"))
' ��             .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","restaurantId","Min"))
' ↑             .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
' ↑         　  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
' ↑             .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));
'        }
'    }
Public Sub CreateTestCode()
    
    Call Initialize
    Dim i As Integer
    Dim testCaseValues As Variant
    testCaseValues = srcRangeSelection.value  ' テストケースデータ
    ' クラス定義
    CreateClassDefine
    
    Dim c As Range
    
    For Each c In srcRowSelection
    'For i = LBound(testCaseValues, 1) To UBound(testCaseValues, 1)
        ' テスト関数作成~リクエスト
        ' テスト関数
        Call CreateFuncRequest(c)
        
        ' リクエストパラメータ
        Call CreateRequestParam1(c)
        
        ' デフォルトモック
        Call CreateMock(c)
        
        ' post getメソッド
        Call CreateMetod(c)
        
        ' リクエストのパラメータ設定を作成
        ' Call CreateRequestParam2(c)
        Call CreateRequestParam3(c)
        
        ' HTTPステータス
        Call CreateHttpStatus
        
        ' HTML名
        Debug.Print c.Offset(0, -2).value & ":" & c.Offset(0, -1).value
        Call CreateReturnHtmlName(c)
        
        ' エラー情報
        Call CreateErrorInfo(c)
        
        ' テスト関数定義終了
        Print #fileNumber, "        }"
        Print #fileNumber, ""
    Next
    
    ' クラス定義終了
    Print #fileNumber, "    }"
    
    Call Terminate
End Sub

Private Sub Initialize()
    
    ' headerセル
    endpoint = Range("D2").value        'エンドポイント
    methodType = Range("D3").value      'POSTとか
    functionName = Range("D4").value    'クラス名
    
    'Form名
    formTagU = Range("D5").value
    formTagL = LCase(Left(formTagU, 1)) & Mid(formTagU, 2)
    

    ' inputセル
    Set srcRangeSelection = ActiveWindow.RangeSelection
    Set srcRowSelection = Range(srcRangeSelection.Columns(1).Address)
    Set startRange = ActiveCell
    
    '出力ファイル
    Dim folderPath As String: folderPath = ActiveWorkbook.Path & "\test\"
    ' フォルダが無ければ作成
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
    filePath = folderPath & getClassName & "Test.java"
    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    

End Sub

Private Sub Terminate()
    Close #fileNumber
    
    If MsgBox("作成したファイルを開く？", vbYesNo) = vbYes Then
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

' 概要：リクエストの送信部分を作成する
'
' 例　：下記の�，鮑鄒�する
' ��   @Test
' ��   public void 正常系_最小() throws Exception {
' ��       // リクエスト
' ��       // デフォルトモック
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
'      　  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
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
    
    
    Print #fileNumber, "            // ================== リクエスト =================="
    Print #fileNumber, "            " + formTagU & " " & formTagL & " = new " & formTagU & "();"
    
    Set fields = Range(srcRangeSelection.Rows(rngCurrent.row - startRange.row + 1).Address)
    
    ' ループ処理
    Dim c As Range
    Dim rowRange As Range
    For Each c In fields
        ' パラメータ生成
        Dim fieldName As String: fieldName = c.Offset(startRange.row - rngCurrent.row - 1, 0).value
        Dim fieldValue As String: fieldValue = convertValue(c.value)
        
        If c.value = "null" Then
            GoTo ContinueLoop
        End If
        paramLine = "            " & formTagL & ".set" & UCase(Left(fieldName, 1)) & Mid(fieldName, 2) & "(" & fieldValue & ");"
        ' ファイル書き込み
        Print #fileNumber, "    " + paramLine

ContinueLoop:
    Next



End Sub


Private Sub CreateMock(c As Range)
    Print #fileNumber, "            // ================== モック =================="
    Print #fileNumber, "            defaultMock();"
    Print #fileNumber, ""
End Sub

Private Sub CreateMetod(c As Range)
    Print #fileNumber, "            // ================== 実行 =================="
    Print #fileNumber, "            mockMvc.perform(" & methodType & "(""" & endpoint & """)"
End Sub


' 概要：リクエストのパラメータ部分を作成する
'
' 例　：下記の�△鮑鄒�する
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
'  　      .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
'      　  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
'  　      .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));

Private Sub CreateRequestParam2(rngCurrent As Range)
    Dim currentFieldValues As Variant
    Dim paramMaxCount As Integer
    Dim paramCount As Integer
    Dim j As Integer
    Dim paramLine As String
    Dim ret As Variant
    Dim fields As Range
    
    
    Set fields = Range(srcRangeSelection.Rows(rngCurrent.row - startRange.row + 1).Address)
    
    ' 有効な要素数をカウント
    paramMaxCount = CountNonNullElements(fields)
    paramCount = 0
    
    ' ループ処理
    Dim c As Range
    Dim rowRange As Range
    
    For Each c In fields
        
        ' パラメータ生成
        Dim fieldName As String: fieldName = c.Offset(startRange.row - rngCurrent.row - 1, 0).value
        Dim fieldValue As String: fieldValue = convertValue(c.value)
        If c.value = "null" Then
            GoTo ContinueLoop
        End If
        paramLine = "                .param(" & Chr(34) & fieldName & Chr(34) & ", " & fieldValue & ")"
        
        ' 最後のパラメータ処理
        paramCount = paramCount + 1
        If paramCount = paramMaxCount Then
            paramLine = paramLine & ")"
        End If
        
        ' ファイル書き込み
        Print #fileNumber, "    " + paramLine

ContinueLoop:
    Next
End Sub

' //                     .flashAttr("taskRegistForm", form)) // フォームオブジェクトを送信
Private Sub CreateRequestParam3(rngCurrent As Range)
    
    ' ファイル書き込み
    Print #fileNumber, "                .flashAttr(" & Chr(34) & formTagL & Chr(34) & " , " & formTagL & ")) // フォームオブジェクト"

End Sub

' 概要：リクエストのHTTPステータスを作成する
'
' 例　：下記の��を作成する
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
' 　       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
'      　  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
' 　       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));
Private Sub CreateHttpStatus()
    Print #fileNumber, "                .andExpect(status().isOk())"
End Sub

' 概要：html名の箇所を作成する
'
' 例　：下記の�い鮑鄒�する
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
' 　       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
'      　  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
' 　       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));
Private Sub CreateReturnHtmlName(rngCurrent As Range)
    Print #fileNumber, "                .andExpect(view().name(""" & rngCurrent.Offset(0, srcRangeSelection.Columns.count + 1).value & """))"
End Sub


' 概要：エラー検証の作成
'
' 例　：下記の�イ鉢Δ鮑鄒�する
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
' ↑       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","userId","Size"))
' ↑   　  .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","rating","Min"))
' ↑       .andExpect(model().attributeHasFieldErrorCode("reviewRegistForm","comment","Size"));
Private Sub CreateErrorInfo(rngCurrent As Range)
    If isNormalTermination(rngCurrent) Then
        Print #fileNumber, "                .andExpect(model().hasNoErrors());"
    Else
        Dim fieldErrors As String
        fieldErrors = getFieldErrors(rngCurrent)
        ' �イ鮑鄒�
        Print #fileNumber, "                .andExpect(model().attributeHasFieldErrors(" & fieldErrors & "))"
        
        ' �Δ鮑鄒�
        Dim rngFieldErrors As Range: Set rngFieldErrors = getFieldErrorRange(rngCurrent)
        Dim errorCount As Integer: errorCount = 0   ' 最終判定に利用。エラー最後は";"で締める。
        Dim errorNum As Integer:   errorNum = CountNonEmptyElements(rngFieldErrors)
        Dim c As Range
        For Each c In rngFieldErrors
            If c.value <> "" Then
                ' フィールド名
                Dim errorField As String: errorField = c.Offset(startRange.row - rngCurrent.row - 1, 0).value
                ' エラーコード
                Dim errorCode As String: errorCode = c.value
                ' 合体１
                Dim fieldErrorCode As String: fieldErrorCode = """" & formTagL & """, """ & errorField & """, """ & errorCode & """"
                ' 合体２
                Dim andExpect As String: andExpect = "                .andExpect(model().attributeHasFieldErrorCode(" & fieldErrorCode & "))"
                        
                ' 終了判定（最後は";"を付ける）
                If errorCount + 1 = errorNum Then
                    andExpect = andExpect & ";"
                End If
                
                ' 出力
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
    
    ' 空でないセルが見つかるまでループ
    Do While row > 1 ' 1行目までに制限
        value = rngCurrent.Offset(-index, TestTypePos).value
        
        If Not IsEmpty(value) And value <> "" Then
            FindNonEmptyCell = value
            Exit Function
        End If
        
        ' 一つ上の行に移動
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


' 有効な要素数をカウント
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

'' 表示値を取得
'Private Function GetDisplayValue(value As Variant) As Variant
'    Dim result(1) As Variant
'
'    Select Case value
'        Case "昨日": result(0) = "LocalDate.now().minusDays(1).toString()": result(1) = 0
'        Case "今日": result(0) = "LocalDate.now().toString()": result(1) = 0
'        Case "明日": result(0) = "LocalDate.now().plusDays(1).toString()": result(1) = 0
'        Case "明後日": result(0) = "LocalDate.now().plusDays(2).toString()": result(1) = 0
'        Case "明々後日": result(0) = "LocalDate.now().plusDays(3).toString()": result(1) = 0
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
'        Case "昨日": result = "LocalDate.now().minusDays(1).toString()"
'        Case "今日": result = "LocalDate.now().toString()"
'        Case "明日": result = "LocalDate.now().plusDays(1).toString()"
'        Case "明後日": result = "LocalDate.now().plusDays(2).toString()"
'        Case "明々後日": result = "LocalDate.now().plusDays(3).toString()"
        Case "昨日": result = "java.sql.Date.valueOf(LocalDate.now().minusDays(1).toString())"
        Case "今日": result = "java.sql.Date.valueOf(LocalDate.now().toString())"
        Case "明日": result = "java.sql.Date.valueOf(LocalDate.now().plusDays(1).toString())"
        Case "明後日": result = "java.sql.Date.valueOf(LocalDate.now().plusDays(2).toString())"
        Case "明々後日": result = "java.sql.Date.valueOf(LocalDate.now().plusDays(3).toString())"
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

' フィールド値を取得
Private Function GetCurrentFieldValues(rowIndex As Integer) As Variant
    Dim r1 As Range, r2 As Range
    Set r1 = startRange.Offset(rowIndex, 0)
    Set r2 = startRange.Offset(rowIndex, fieldCount - 1)
    GetCurrentFieldValues = sheet.Range(r1.Address & ":" & r2.Address).value
End Function

' フィールド名を取得
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
    If testcase = "正常系" Then
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
    
    ' errorCodeの行（Min, Sizeなど書かれているセル範囲）を取得
    Dim currentErrorFieldRange As Range: Set currentErrorFieldRange = getFieldErrorRange(rngCurrent)
    
    ' エラーがあった（Min, Sizeなど書かれている）フィールド名を返す
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
    
    ' エラーがある項目を返す。
    ' errorCodeの行（Min, Sizeなど書かれているセル範囲）を返す。
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
