Attribute VB_Name = "m1_createDataClass"
Private rngAnnotation As Range

Private Sub init()
    Set rngAnnotation = Range("AA11:AC200")
End Sub

Sub create_data_class()

    Call init

    '出力ファイル
    Dim filePath As String: filePath = ActiveWorkbook.Path & "\" & getClassName & ".java"
    Dim fileNumber As Integer: fileNumber = FreeFile
    Open filePath For Output As #fileNumber

    ' // 関数ヘッダ作成
    Call createFuncHeder(fileNumber)
    
    ' // 関数本体作成
    Call createFuncBody(fileNumber)
    
    ' ファイルを閉じる
    Close #fileNumber

End Sub

' ' // 関数ヘッダ作成
 Private Sub createFuncHeder(fileNumber)

    Print #fileNumber, "/**"
    Print #fileNumber, " * " & ActiveCell.Offset(-1, -1).value
    Print #fileNumber, " *"
    
    Dim c As Range
    For Each c In ActiveWindow.RangeSelection
        Print #fileNumber, " * @property " & c.value
    Next
    Print #fileNumber, " */"
 End Sub

' // 関数本体作成
 Private Sub createFuncBody(fileNumber)
 
    Print #fileNumber, "@Data"
    Print #fileNumber, "public class " + getClassName + " {"

    Dim c As Range
    For Each c In ActiveWindow.RangeSelection
        
        ' // 論理名
        Print #fileNumber, "    // " + c.value
        
        ' // アノテーション
        Debug.Print c.value
        Call createAnotationForList(fileNumber, c)
        
        ' // メンバ定義
        Print #fileNumber, "    private " & c.Offset(0, srcPos.offsetType).value; " " + c.Offset(0, srcPos.offsetPhysical).value & ";"
        Print #fileNumber, ""
    Next
    
    Print #fileNumber, "}"

 End Sub

Private Function getClassName() As String
    getClassName = ActiveCell.Offset(-1, srcPos.offsetPhysical).value
    getClassName = UCase(Left(getClassName, 1)) & Mid(getClassName, 2)
    
End Function


' // アノテーション作成
Private Sub createAnotationForList(fileNumber, rngCurrent)
    Dim physics As String: physics = rngCurrent.Offset(0, srcPos.offsetPhysical).value
    Dim required As String: required = rngCurrent.Offset(0, srcPos.offsetRequired).value
    Dim min As String: min = rngCurrent.Offset(0, srcPos.offsetMin).value
    Dim max As String: max = rngCurrent.Offset(0, srcPos.offsetMax).value
    Dim stype As String: stype = rngCurrent.Offset(0, srcPos.offsetType).value
    
    Dim clAnnotation As New Collection
    
    ' // 必須項目
    If required = "有" Then
        Print #fileNumber, "    " + "@NotNull" + "(message=" + Chr(34) + "入力してください。" + Chr(34) + ")"
        Call clAnnotation.Add("@NotNull")
    End If

    If stype = "String" Then
        ' 指定あり（最小・最大）
        If min <> "" And max <> "" Then
            
            If min = max Then
                Print #fileNumber, "    " + "@Size(min=" + min + ",max=" + max + ",message=" + Chr(34) + min + "文字で" + "指定してください。" + Chr(34) + ")"
            Else
                Print #fileNumber, "    " + "@Size(min=" + min + ",max=" + max + ",message=" + Chr(34) + min + "文字から" + max + "文字で" + "指定してください。" + Chr(34) + ")"
            End If
            Call clAnnotation.Add("@Size")
        End If
        
        ' 指定あり（最小）
        If min <> "" And max = "" Then
            Print #fileNumber, "    " + "@Size(min=" + min + ",message=" + Chr(34) + min + "文字以上で" + "指定してください。" + Chr(34) + ")"
            Call clAnnotation.Add("@Size")
        End If
        
        ' 指定あり（最大）
        If min = "" And max <> "" Then
            Print #fileNumber, "    " + "@Size(max=" + max + ",message=" + Chr(34) + max + "文字以下で" + "指定してください。" + Chr(34) + ")"
            Call clAnnotation.Add("@Size")
        End If
    End If

    If stype = "Integer" Then
        If min <> "" And max <> "" Then
            Print #fileNumber, "    " + "@Min(value=" + min + ",message=" + Chr(34) + min + "-" + max + "で" + "指定してください。" + Chr(34) + ")"
            Print #fileNumber, "    " + "@Max(value=" + max + ",message=" + Chr(34) + min + "-" + max + "で" + "指定してください。" + Chr(34) + ")"
            Call clAnnotation.Add("@Min")
            Call clAnnotation.Add("@Max")
        End If
        
        If min <> "" And max = "" Then
            If min = 1 Then
                Print #fileNumber, "    " + "@Min(value=" + min + ",message=" + Chr(34) + "正の整数を入力してください。" + Chr(34) & ")"
                Call clAnnotation.Add("@Min")
            Else
                Print #fileNumber, "    " + "@Min(value=" + min + ",message=" + Chr(34) + min + "以上で" + "指定してください。" + Chr(34) + ")"
                Call clAnnotation.Add("@Min")
            End If
        End If
        
        If min = "" And max <> "" Then
            Print #fileNumber, "    " + "@Max(value=" + max + ",message=" + Chr(34) + max + "以下で" + "指定してください。" + Chr(34) + ")"
            Call clAnnotation.Add("@Max")
        End If
    End If
        
    ' アノテーションリストも追加（但し、既につけているアノテーションは無視）
    Dim c As Range
    For Each c In rngAnnotation
        If c.value = rngCurrent.value Then
            If clsContain(clAnnotation, c) = False Then
                Dim strAnnotaion As String: strAnnotaion = c.Offset(0, 1).value
                Dim strErrorMsg As String: strErrorMsg = c.Offset(0, 2).value
                Print #fileNumber, "    " + strAnnotaion + "(message=" + Chr(34) + strErrorMsg + Chr(34) + ")"
                Call clAnnotation.Add(strAnnotaion)
            End If
        End If
    Next
    
    Set clAnnotation = Nothing
End Sub


Private Function clsContain(cls As Collection, c As Range) As Boolean
    Dim t As Collection
    Dim i As Integer
    For i = 1 To cls.count
        If cls.Item(i) = c.Offset(0, 1).value Then
            clsContain = True
            Exit Function
        End If
    Next
    clsContain = False
End Function
