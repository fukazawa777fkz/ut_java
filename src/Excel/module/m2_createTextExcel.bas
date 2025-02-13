Attribute VB_Name = "m2_createTextExcel"
Private writeCurrent As Integer

Private srcHeaderRange As Range
Private srcRangeSelection As Range

Private destHeaderRange As Range
Private destBodyRange As Range

' リスト
Private minList As Collection ' ItemInfo
Private maxList As Collection ' ItemInfo
Private requiredList As Collection ' ItemInfo
Private enumList As Collection ' ItemInfo
Private defaultList As Collection ' ItemInfo

Public Sub CreateUnitTestExcel()

    ' // 初期化
    Call Initialize
    
    ' // ヘッダを作成
    Call writeHeader
    
    ' // 試験表を作成
    Call createTable
    
    ' // 正常系を作成
    Call writeNormal
    
    ' // 異常系を作成
    Call writeAbnormal
    
End Sub


Private Sub Initialize()
    
    
    ' inputセル
    Set srcRangeSelection = ActiveWindow.RangeSelection
    Set srcHeaderRange = ActiveSheet.Range("C2")
    
    ' outputセル
    Worksheets.Add
    Set destHeaderRange = ActiveSheet.Range("C2")
    Set destBodyRange = ActiveSheet.Range("E8")
    
    Set minList = New Collection
    Set maxList = New Collection
    Set requiredList = New Collection
    Set enumList = New Collection
    Set defaultList = New Collection
    
    ' 出力位置
    writeCurrent = 1
    
    ' リスト
    Dim c As Range
    For Each c In srcRangeSelection
        
        Dim index As Integer
        Dim oItem As ItemInfo
        
        ' デフォルト
        Set oItem = New ItemInfo
        Call oItem.constructor(c, index, "")
        Call defaultList.Add(oItem)
        
        ' 必須リスト
        If c.Offset(0, srcPos.offsetRequired).value <> "" Then
            Set oItem = New ItemInfo
            Call oItem.constructor(c, index, getRequiredValue(c))
            Call requiredList.Add(oItem)
        End If
        
        ' enumリスト
        If c.Offset(0, srcPos.offsetEnum).value <> "" Then
            oItem = New ItemInfo
            Call oItem.constructor(c, index, "")
            enumList.Add (oItem)
        End If
        
        ' minリスト, maxリスト
        If c.Offset(0, srcPos.offsetMin).value <> "" And c.Offset(0, srcPos.offsetMax).value <> "" Then
            ' minリスト
            Set oItem = New ItemInfo
            Call oItem.constructor(c, index, c.Offset(0, srcPos.offsetMin).value)
            Call minList.Add(oItem)
        
            ' maxリスト
            Set oItem = New ItemInfo
            Call oItem.constructor(c, index, c.Offset(0, srcPos.offsetMax).value)
            Call maxList.Add(oItem)
        
        End If
        
        index = index + 1
    Next
      

    
End Sub

Private Function getRequiredValue(rng As Range)

    Dim typeInfo As String: typeInfo = rng.Offset(0, srcPos.offsetType)
    Dim max As String: max = rng.Offset(0, srcPos.offsetMax)
    Dim min As String: min = rng.Offset(0, srcPos.offsetMin)
    
    Dim ret As String
    If typeInfo = "String" Then
    
    Else
        ' // 最小があるなら最小値
        If min <> "" Then
            getRequiredValue = rng.Offset(0, srcPos.offsetMin).value
            Exit Function
        End If
        
        ' // 最大があるなら最大値
        If max <> "" Then
            getRequiredValue = rng.Offset(0, srcPos.offsetMax).value
            Exit Function
        End If
    End If
    
    ' 適当に6
    getRequiredValue = 6
    Exit Function
    
End Function

Private Sub writeHeader()
    destHeaderRange.Offset(0, 0).value = "エンドポイント"
    destHeaderRange.Offset(0, 1).value = srcHeaderRange.Offset(0, 5).value
    destHeaderRange.Offset(1, 0).value = "メソッド"
    destHeaderRange.Offset(1, 1).value = srcHeaderRange.Offset(1, 5).value
    destHeaderRange.Offset(2, 0).value = "機能名"
    destHeaderRange.Offset(2, 1).value = srcHeaderRange.Offset(2, 5).value
    destHeaderRange.Offset(3, 0).value = "入力フォーム"
    destHeaderRange.Offset(3, 1).value = getClassName()
End Sub

Private Sub createTable()
    
    destBodyRange.Offset(-2, 0).value = "入力"
    destBodyRange.Offset(-2, srcRangeSelection.count + 0).value = "期待値"
    destBodyRange.Offset(-1, srcRangeSelection.count + 0).value = "HTTPステータス"
    destBodyRange.Offset(-1, srcRangeSelection.count + 1).value = "HTML名"
    destBodyRange.Offset(-1, srcRangeSelection.count + 2).value = "errorCode"
    
    Dim c As Range
    Dim destCnt As Integer
    For Each c In srcRangeSelection
        destBodyRange.Offset(-1, destCnt).value = c.Offset(0, 0).value
        destBodyRange.Offset(0, destCnt).value = c.Offset(0, srcPos.offsetPhysical).value
        destBodyRange.Offset(0, destCnt + srcRangeSelection.count + FixedFiledNum).value = c.Offset(0, srcPos.offsetPhysical).value
        destCnt = destCnt + 1
    Next
End Sub

Private Sub writeBody()
End Sub

Private Function getClassName() As String
    getClassName = srcRangeSelection.Cells(1, 1).Offset(-1, srcPos.offsetPhysical)
    getClassName = UCase(Left(getClassName, 1)) & Mid(getClassName, 2)
End Function

Private Sub writeNormal()
    
    ' 正常系
    destBodyRange.Offset(writeCurrent, -2).value = "正常系"
    
    ' 最小
    destBodyRange.Offset(writeCurrent, -1).value = "最小"
    Call setFeildsValues(eTestType.NORMAL, minList, 0)
    writeCurrent = writeCurrent + 1
    
    ' 最大
    destBodyRange.Offset(writeCurrent, -1).value = "最大"
    Call setFeildsValues(eTestType.NORMAL, maxList, 0)
    writeCurrent = writeCurrent + 1
    
    ' 必須のみ
    destBodyRange.Offset(writeCurrent, -1).value = "必須のみ"
    Call setFeildsValues(eTestType.NORMAL, requiredList, 0, True)
    writeCurrent = writeCurrent + 1
    
    ' 空文字
    destBodyRange.Offset(writeCurrent, -1).value = "空文字"
    Call setFieldsDirect("")
    writeCurrent = writeCurrent + 1
    
End Sub

Sub writeAbnormal()

    ' 正常系
    destBodyRange.Offset(writeCurrent, -2).value = "異常系"
    
    ' 最小
    destBodyRange.Offset(writeCurrent, -1).value = "最小"
    Call setFeildsValues(eTestType.ABNORMAL, minList, -1)
    writeCurrent = writeCurrent + 1
    
    ' 最大
    destBodyRange.Offset(writeCurrent, -1).value = "最大"
    Call setFeildsValues(eTestType.ABNORMAL, maxList, 1)
    writeCurrent = writeCurrent + 1
    
    ' 必須
    destBodyRange.Offset(writeCurrent, -1).value = "null値"
    Call setFeildsValues(eTestType.ABNORMAL, requiredList, 0, True, True)
    writeCurrent = writeCurrent + 1
    
    ' 空文字
    destBodyRange.Offset(writeCurrent, -1).value = "空文字"
    Call setFieldsDirect("", True)
    writeCurrent = writeCurrent + 1
    
    ' 半角スペース
    destBodyRange.Offset(writeCurrent, -1).value = "半角スペース"
    Call setFieldsDirect(" ", False)
    writeCurrent = writeCurrent + 1
    
    ' 全角スペース
    destBodyRange.Offset(writeCurrent, -1).value = "全角スペース"
    Call setFieldsDirect("　", False)
    writeCurrent = writeCurrent + 1

End Sub

' * 概要：入力値を設定する
' * param：list 対象リスト
' * param：valueOffset 異常系の設定をする場合は値を設定しておく
' * param：isRequired 必須入力項目かどうか（その他の設定をするとき、必須でないものはnullに設定される）
' * param：isToggle 必須入力項目でないものをnullを設定するようになる

Private Sub setFeildsValues(TestType, list As Collection, valueOffset, Optional isRequired = False, Optional isToggle = False)


  ' nullの指定は、isToggleを元に設定する
    Dim index As Integer
    Dim ItemInfo As ItemInfo
    Dim rngObj As Range
    Dim errorCodeType As String
    Dim ws As Worksheet

    ' list を配列としてループ処理
    ' For index = LBound(list) To UBound(list)
    
    For index = 1 To list.count
        Set ItemInfo = list.Item(index)
        Set rngObj = destBodyRange.Offset(writeCurrent, ItemInfo.fileId)

        If isToggle = False Then
            If ItemInfo.typeInfo = "Integer" Or ItemInfo.typeInfo = "Int" Then
                rngObj.value = CDbl(ItemInfo.value) + valueOffset
                errorCodeType = "Max"
                If valueOffset < 0 Then
                    errorCodeType = "Min"
                End If
                Call setErrorCode(TestType, errorCodeType, rngObj, ItemInfo)
                
            ElseIf ItemInfo.typeInfo = "String" Then
                rngObj.Formula = "=REPT(""a"", " & (CDbl(ItemInfo.value) + valueOffset) & ")"
                Call setErrorCode(TestType, "Size", rngObj, ItemInfo)
                
            Else
                rngObj.value = CDbl(ItemInfo.value) + valueOffset
                Call setErrorCode(TestType, "Size", rngObj, ItemInfo)
            End If

            ' 空白セルの背景色を変更
            If rngObj.value = "" Then
                Set ws = rngObj.Worksheet
                Debug.Print ws.name
                Debug.Print rngObj.Address
                rngObj.Interior.Color = RGB(255, 255, 0) ' 黄色
            End If

        Else
            Select Case ItemInfo.typeInfo
                Case "Integer"
                    rngObj.value = "null"
                    Call setErrorCode(TestType, "NotNull", rngObj, ItemInfo)
                Case "Int"
                    rngObj.value = 0
                    Call setErrorCode(TestType, "Invalid", rngObj, ItemInfo)
                Case "String"
                    rngObj.value = "null"
                    Call setErrorCode(TestType, "NotNull", rngObj, ItemInfo)
                Case Else
                    rngObj.value = "null"
                    Call setErrorCode(TestType, "NotNull", rngObj, ItemInfo)
            End Select
        End If
    Next index

    ' 対象リストになかったフィールドを設定
    Call SetOtherFieldsValue(list, isRequired, isToggle, TestType, valueOffset)

End Sub


Sub setErrorCode(TestType, errorCodeType, rngObj, ItemInfo)

    If TestType = eTestType.NORMAL Then
        Exit Sub
    End If
    
    ' rngObj.Offset(0, this.numRows + FixedFiledNum).SetValue (errorCodeType)
    rngObj.Offset(0, srcRangeSelection.count + FixedFiledNum).value = errorCodeType

End Sub

Private Function isContain(list As Collection, physics As String) As Boolean
    Dim ItemInfo As ItemInfo
    For Each ItemInfo In list
        If ItemInfo.physics = physics Then
            isContain = True
            Exit Function
        End If
    Next
    isContain = False
End Function


Private Sub SetOtherFieldsValue(list As Collection, ByVal isRequired As Boolean, ByVal isToggle As Boolean, ByVal TestType As String, ByVal valueOffset As Double)
    Dim index As Integer
    Dim ItemInfo As ItemInfo
    Dim rngObj As Range
    Dim value As Variant
    Dim errorFlg As Boolean
    Dim errorCodeType As String
    Dim formattedDate As String
    Dim ws As Worksheet
    

    ' defaultList のループ
    For index = 1 To defaultList.count
    
        Set ItemInfo = defaultList.Item(index) 'Me.defaultList(index) ' クラスまたは辞書オブジェクトを想定
        
        Debug.Print ItemInfo.physics
        
        ' 対象リストに含まれていないか確認
        If (isContain(list, defaultList.Item(index).physics) = False) Then
            
            ' 対象セルの取得
            'Set rngObj = Me.stdRange.Offset(Me.current, itemInfo("fileId"))
            Set rngObj = destBodyRange.Offset(writeCurrent, ItemInfo.fileId)
           
            ' 条件による値設定
            If isRequired = True And isToggle = False Then
                rngObj.value = "null"
            Else
                value = ItemInfo.value
                errorFlg = False
                
                If value = "" And ItemInfo.min <> "" Then
                    value = ItemInfo.min + valueOffset
                    If value < ItemInfo.min Then
                        errorFlg = True
                    Else
                        value = ItemInfo.min
                    End If
                End If
                
                If value = "" And ItemInfo.max <> "" Then
                    value = ItemInfo.max + valueOffset
                    If value > ItemInfo.max Then
                        errorFlg = True
                    Else
                        value = ItemInfo.max
                    End If
                End If
                
                If value = "" Then
                    value = 2
                End If
                
                ' 型別処理
                If ItemInfo.typeInfo = "Integer" Or ItemInfo.typeInfo = "Int" Then
                    rngObj.value = value
                    errorCodeType = "Max"
                    If valueOffset < 0 Then
                        errorCodeType = "Min"
                    End If
                    If errorFlg Then
                        setErrorCode TestType, errorCodeType, rngObj, ItemInfo
                    End If
                
                ElseIf ItemInfo.typeInfo = "String" Then
                    rngObj.Formula = "=REPT(""z"", " & value & ")"
                
                ElseIf ItemInfo.typeInfo = "Date" Then
                    formattedDate = Format(Now, "yyyy/mm/dd")
                    rngObj.value = formattedDate
                
                Else
                    rngObj.value = value
                End If
                
                ' 値が空なら背景色変更＆デバッグログ
                If rngObj.value = "" Then
                    Debug.Print rngObj.Address
                    rngObj.Interior.Color = RGB(255, 255, 0) ' 黄色
                End If
            End If
        End If
    Next index
End Sub

Private Sub setFieldsDirect(paramValue As String, Optional ByVal isStringForced As Boolean = False)
    Dim rngObj As Range
    Dim value As Variant
    Dim formattedDate As String
    
    Dim ItemInfo As ItemInfo
    
    
    For Each ItemInfo In defaultList
        Set rngObj = destBodyRange.Offset(writeCurrent, ItemInfo.fileId)
        value = ItemInfo.value
        
        ' value が空の場合、min か max を代入
        If value = "" And ItemInfo.min <> "" Then
            value = ItemInfo.min
        End If
        If value = "" And ItemInfo.max <> "" Then
            value = ItemInfo.max
        End If
        
        ' データ型に応じた処理
        If ItemInfo.typeInfo = "Integer" Or ItemInfo.typeInfo = "Int" Then
            If value = "" Then value = 0
            rngObj.value = value
        
        ElseIf ItemInfo.typeInfo = "String" Then
            If isStringForced = True Then
                rngObj.value = paramValue
            Else
                If value <> "" Then
                    rngObj.Formula = "=REPT(""" & paramValue & """," & value & ")"
                Else
                    rngObj.value = paramValue
                End If
            End If
        
        ElseIf ItemInfo.typeInfo = "Date" Then
            ' 今日の日付を取得（yyyy/MM/dd）
            formattedDate = Format(Date, "yyyy/MM/dd")
            rngObj.value = formattedDate
        Else
            rngObj.value = paramValue
        End If
        
        ' セルが空なら背景色を黄色にする
        If rngObj.value = "" Then
            Debug.Print rngObj.Address ' Logger.log() の代わりにデバッグ出力
            rngObj.Interior.Color = RGB(255, 255, 0)
        End If
    Next
End Sub

