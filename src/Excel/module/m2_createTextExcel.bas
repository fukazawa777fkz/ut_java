Attribute VB_Name = "m2_createTextExcel"
Private writeCurrent As Integer

Private srcHeaderRange As Range
Private srcRangeSelection As Range

Private destHeaderRange As Range
Private destBodyRange As Range

' ���X�g
Private minList As Collection ' ItemInfo
Private maxList As Collection ' ItemInfo
Private requiredList As Collection ' ItemInfo
Private enumList As Collection ' ItemInfo
Private defaultList As Collection ' ItemInfo

Public Sub CreateUnitTestExcel()

    ' // ������
    Call Initialize
    
    ' // �w�b�_���쐬
    Call writeHeader
    
    ' // �����\���쐬
    Call createTable
    
    ' // ����n���쐬
    Call writeNormal
    
    ' // �ُ�n���쐬
    Call writeAbnormal
    
End Sub


Private Sub Initialize()
    
    
    ' input�Z��
    Set srcRangeSelection = ActiveWindow.RangeSelection
    Set srcHeaderRange = ActiveSheet.Range("C2")
    
    ' output�Z��
    Worksheets.Add
    Set destHeaderRange = ActiveSheet.Range("C2")
    Set destBodyRange = ActiveSheet.Range("E8")
    
    Set minList = New Collection
    Set maxList = New Collection
    Set requiredList = New Collection
    Set enumList = New Collection
    Set defaultList = New Collection
    
    ' �o�͈ʒu
    writeCurrent = 1
    
    ' ���X�g
    Dim c As Range
    For Each c In srcRangeSelection
        
        Dim index As Integer
        Dim oItem As ItemInfo
        
        ' �f�t�H���g
        Set oItem = New ItemInfo
        Call oItem.constructor(c, index, "")
        Call defaultList.Add(oItem)
        
        ' �K�{���X�g
        If c.Offset(0, srcPos.offsetRequired).value <> "" Then
            Set oItem = New ItemInfo
            Call oItem.constructor(c, index, getRequiredValue(c))
            Call requiredList.Add(oItem)
        End If
        
        ' enum���X�g
        If c.Offset(0, srcPos.offsetEnum).value <> "" Then
            oItem = New ItemInfo
            Call oItem.constructor(c, index, "")
            enumList.Add (oItem)
        End If
        
        ' min���X�g, max���X�g
        If c.Offset(0, srcPos.offsetMin).value <> "" And c.Offset(0, srcPos.offsetMax).value <> "" Then
            ' min���X�g
            Set oItem = New ItemInfo
            Call oItem.constructor(c, index, c.Offset(0, srcPos.offsetMin).value)
            Call minList.Add(oItem)
        
            ' max���X�g
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
        ' // �ŏ�������Ȃ�ŏ��l
        If min <> "" Then
            getRequiredValue = rng.Offset(0, srcPos.offsetMin).value
            Exit Function
        End If
        
        ' // �ő傪����Ȃ�ő�l
        If max <> "" Then
            getRequiredValue = rng.Offset(0, srcPos.offsetMax).value
            Exit Function
        End If
    End If
    
    ' �K����6
    getRequiredValue = 6
    Exit Function
    
End Function

Private Sub writeHeader()
    destHeaderRange.Offset(0, 0).value = "�G���h�|�C���g"
    destHeaderRange.Offset(0, 1).value = srcHeaderRange.Offset(0, 5).value
    destHeaderRange.Offset(1, 0).value = "���\�b�h"
    destHeaderRange.Offset(1, 1).value = srcHeaderRange.Offset(1, 5).value
    destHeaderRange.Offset(2, 0).value = "�@�\��"
    destHeaderRange.Offset(2, 1).value = srcHeaderRange.Offset(2, 5).value
    destHeaderRange.Offset(3, 0).value = "���̓t�H�[��"
    destHeaderRange.Offset(3, 1).value = getClassName()
End Sub

Private Sub createTable()
    
    destBodyRange.Offset(-2, 0).value = "����"
    destBodyRange.Offset(-2, srcRangeSelection.count + 0).value = "���Ғl"
    destBodyRange.Offset(-1, srcRangeSelection.count + 0).value = "HTTP�X�e�[�^�X"
    destBodyRange.Offset(-1, srcRangeSelection.count + 1).value = "HTML��"
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
    
    ' ����n
    destBodyRange.Offset(writeCurrent, -2).value = "����n"
    
    ' �ŏ�
    destBodyRange.Offset(writeCurrent, -1).value = "�ŏ�"
    Call setFeildsValues(eTestType.NORMAL, minList, 0)
    writeCurrent = writeCurrent + 1
    
    ' �ő�
    destBodyRange.Offset(writeCurrent, -1).value = "�ő�"
    Call setFeildsValues(eTestType.NORMAL, maxList, 0)
    writeCurrent = writeCurrent + 1
    
    ' �K�{�̂�
    destBodyRange.Offset(writeCurrent, -1).value = "�K�{�̂�"
    Call setFeildsValues(eTestType.NORMAL, requiredList, 0, True)
    writeCurrent = writeCurrent + 1
    
    ' �󕶎�
    destBodyRange.Offset(writeCurrent, -1).value = "�󕶎�"
    Call setFieldsDirect("")
    writeCurrent = writeCurrent + 1
    
End Sub

Sub writeAbnormal()

    ' ����n
    destBodyRange.Offset(writeCurrent, -2).value = "�ُ�n"
    
    ' �ŏ�
    destBodyRange.Offset(writeCurrent, -1).value = "�ŏ�"
    Call setFeildsValues(eTestType.ABNORMAL, minList, -1)
    writeCurrent = writeCurrent + 1
    
    ' �ő�
    destBodyRange.Offset(writeCurrent, -1).value = "�ő�"
    Call setFeildsValues(eTestType.ABNORMAL, maxList, 1)
    writeCurrent = writeCurrent + 1
    
    ' �K�{
    destBodyRange.Offset(writeCurrent, -1).value = "null�l"
    Call setFeildsValues(eTestType.ABNORMAL, requiredList, 0, True, True)
    writeCurrent = writeCurrent + 1
    
    ' �󕶎�
    destBodyRange.Offset(writeCurrent, -1).value = "�󕶎�"
    Call setFieldsDirect("", True)
    writeCurrent = writeCurrent + 1
    
    ' ���p�X�y�[�X
    destBodyRange.Offset(writeCurrent, -1).value = "���p�X�y�[�X"
    Call setFieldsDirect(" ", False)
    writeCurrent = writeCurrent + 1
    
    ' �S�p�X�y�[�X
    destBodyRange.Offset(writeCurrent, -1).value = "�S�p�X�y�[�X"
    Call setFieldsDirect("�@", False)
    writeCurrent = writeCurrent + 1

End Sub

' * �T�v�F���͒l��ݒ肷��
' * param�Flist �Ώۃ��X�g
' * param�FvalueOffset �ُ�n�̐ݒ������ꍇ�͒l��ݒ肵�Ă���
' * param�FisRequired �K�{���͍��ڂ��ǂ����i���̑��̐ݒ������Ƃ��A�K�{�łȂ����̂�null�ɐݒ肳���j
' * param�FisToggle �K�{���͍��ڂłȂ����̂�null��ݒ肷��悤�ɂȂ�

Private Sub setFeildsValues(TestType, list As Collection, valueOffset, Optional isRequired = False, Optional isToggle = False)


  ' null�̎w��́AisToggle�����ɐݒ肷��
    Dim index As Integer
    Dim ItemInfo As ItemInfo
    Dim rngObj As Range
    Dim errorCodeType As String
    Dim ws As Worksheet

    ' list ��z��Ƃ��ă��[�v����
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

            ' �󔒃Z���̔w�i�F��ύX
            If rngObj.value = "" Then
                Set ws = rngObj.Worksheet
                Debug.Print ws.name
                Debug.Print rngObj.Address
                rngObj.Interior.Color = RGB(255, 255, 0) ' ���F
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

    ' �Ώۃ��X�g�ɂȂ������t�B�[���h��ݒ�
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
    

    ' defaultList �̃��[�v
    For index = 1 To defaultList.count
    
        Set ItemInfo = defaultList.Item(index) 'Me.defaultList(index) ' �N���X�܂��͎����I�u�W�F�N�g��z��
        
        Debug.Print ItemInfo.physics
        
        ' �Ώۃ��X�g�Ɋ܂܂�Ă��Ȃ����m�F
        If (isContain(list, defaultList.Item(index).physics) = False) Then
            
            ' �ΏۃZ���̎擾
            'Set rngObj = Me.stdRange.Offset(Me.current, itemInfo("fileId"))
            Set rngObj = destBodyRange.Offset(writeCurrent, ItemInfo.fileId)
           
            ' �����ɂ��l�ݒ�
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
                
                ' �^�ʏ���
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
                
                ' �l����Ȃ�w�i�F�ύX���f�o�b�O���O
                If rngObj.value = "" Then
                    Debug.Print rngObj.Address
                    rngObj.Interior.Color = RGB(255, 255, 0) ' ���F
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
        
        ' value ����̏ꍇ�Amin �� max ����
        If value = "" And ItemInfo.min <> "" Then
            value = ItemInfo.min
        End If
        If value = "" And ItemInfo.max <> "" Then
            value = ItemInfo.max
        End If
        
        ' �f�[�^�^�ɉ���������
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
            ' �����̓��t���擾�iyyyy/MM/dd�j
            formattedDate = Format(Date, "yyyy/MM/dd")
            rngObj.value = formattedDate
        Else
            rngObj.value = paramValue
        End If
        
        ' �Z������Ȃ�w�i�F�����F�ɂ���
        If rngObj.value = "" Then
            Debug.Print rngObj.Address ' Logger.log() �̑���Ƀf�o�b�O�o��
            rngObj.Interior.Color = RGB(255, 255, 0)
        End If
    Next
End Sub

