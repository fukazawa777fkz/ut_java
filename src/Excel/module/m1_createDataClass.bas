Attribute VB_Name = "m1_createDataClass"
Private rngAnnotation As Range

Private Sub init()
    Set rngAnnotation = Range("AA11:AC200")
End Sub

Sub create_data_class()

    Call init

    '�o�̓t�@�C��
    Dim filePath As String: filePath = ActiveWorkbook.Path & "\" & getClassName & ".java"
    Dim fileNumber As Integer: fileNumber = FreeFile
    Open filePath For Output As #fileNumber

    ' // �֐��w�b�_�쐬
    Call createFuncHeder(fileNumber)
    
    ' // �֐��{�̍쐬
    Call createFuncBody(fileNumber)
    
    ' �t�@�C�������
    Close #fileNumber

End Sub

' ' // �֐��w�b�_�쐬
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

' // �֐��{�̍쐬
 Private Sub createFuncBody(fileNumber)
 
    Print #fileNumber, "@Data"
    Print #fileNumber, "public class " + getClassName + " {"

    Dim c As Range
    For Each c In ActiveWindow.RangeSelection
        
        ' // �_����
        Print #fileNumber, "    // " + c.value
        
        ' // �A�m�e�[�V����
        Debug.Print c.value
        Call createAnotationForList(fileNumber, c)
        
        ' // �����o��`
        Print #fileNumber, "    private " & c.Offset(0, srcPos.offsetType).value; " " + c.Offset(0, srcPos.offsetPhysical).value & ";"
        Print #fileNumber, ""
    Next
    
    Print #fileNumber, "}"

 End Sub

Private Function getClassName() As String
    getClassName = ActiveCell.Offset(-1, srcPos.offsetPhysical).value
    getClassName = UCase(Left(getClassName, 1)) & Mid(getClassName, 2)
    
End Function


' // �A�m�e�[�V�����쐬
Private Sub createAnotationForList(fileNumber, rngCurrent)
    Dim physics As String: physics = rngCurrent.Offset(0, srcPos.offsetPhysical).value
    Dim required As String: required = rngCurrent.Offset(0, srcPos.offsetRequired).value
    Dim min As String: min = rngCurrent.Offset(0, srcPos.offsetMin).value
    Dim max As String: max = rngCurrent.Offset(0, srcPos.offsetMax).value
    Dim stype As String: stype = rngCurrent.Offset(0, srcPos.offsetType).value
    
    Dim clAnnotation As New Collection
    
    ' // �K�{����
    If required = "�L" Then
        Print #fileNumber, "    " + "@NotNull" + "(message=" + Chr(34) + "���͂��Ă��������B" + Chr(34) + ")"
        Call clAnnotation.Add("@NotNull")
    End If

    If stype = "String" Then
        ' �w�肠��i�ŏ��E�ő�j
        If min <> "" And max <> "" Then
            
            If min = max Then
                Print #fileNumber, "    " + "@Size(min=" + min + ",max=" + max + ",message=" + Chr(34) + min + "������" + "�w�肵�Ă��������B" + Chr(34) + ")"
            Else
                Print #fileNumber, "    " + "@Size(min=" + min + ",max=" + max + ",message=" + Chr(34) + min + "��������" + max + "������" + "�w�肵�Ă��������B" + Chr(34) + ")"
            End If
            Call clAnnotation.Add("@Size")
        End If
        
        ' �w�肠��i�ŏ��j
        If min <> "" And max = "" Then
            Print #fileNumber, "    " + "@Size(min=" + min + ",message=" + Chr(34) + min + "�����ȏ��" + "�w�肵�Ă��������B" + Chr(34) + ")"
            Call clAnnotation.Add("@Size")
        End If
        
        ' �w�肠��i�ő�j
        If min = "" And max <> "" Then
            Print #fileNumber, "    " + "@Size(max=" + max + ",message=" + Chr(34) + max + "�����ȉ���" + "�w�肵�Ă��������B" + Chr(34) + ")"
            Call clAnnotation.Add("@Size")
        End If
    End If

    If stype = "Integer" Then
        If min <> "" And max <> "" Then
            Print #fileNumber, "    " + "@Min(value=" + min + ",message=" + Chr(34) + min + "-" + max + "��" + "�w�肵�Ă��������B" + Chr(34) + ")"
            Print #fileNumber, "    " + "@Max(value=" + max + ",message=" + Chr(34) + min + "-" + max + "��" + "�w�肵�Ă��������B" + Chr(34) + ")"
            Call clAnnotation.Add("@Min")
            Call clAnnotation.Add("@Max")
        End If
        
        If min <> "" And max = "" Then
            If min = 1 Then
                Print #fileNumber, "    " + "@Min(value=" + min + ",message=" + Chr(34) + "���̐�������͂��Ă��������B" + Chr(34) & ")"
                Call clAnnotation.Add("@Min")
            Else
                Print #fileNumber, "    " + "@Min(value=" + min + ",message=" + Chr(34) + min + "�ȏ��" + "�w�肵�Ă��������B" + Chr(34) + ")"
                Call clAnnotation.Add("@Min")
            End If
        End If
        
        If min = "" And max <> "" Then
            Print #fileNumber, "    " + "@Max(value=" + max + ",message=" + Chr(34) + max + "�ȉ���" + "�w�肵�Ă��������B" + Chr(34) + ")"
            Call clAnnotation.Add("@Max")
        End If
    End If
        
    ' �A�m�e�[�V�������X�g���ǉ��i�A���A���ɂ��Ă���A�m�e�[�V�����͖����j
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
