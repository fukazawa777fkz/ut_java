VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ItemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ItemInfo クラスモジュール (Class Module) を作成
Option Explicit

Public value As String
Public name As String
Public physics As String
Public typeInfo As String
Public required As String
Public min As String
Public max As String
Public enumValue As String
Public fileId As String

Public Sub constructor(rng As Range, fileId As Integer, value As String)
    Me.value = value
    Me.fileId = fileId
    
    Me.name = rng.Offset(0, 0)
    Me.physics = rng.Offset(0, srcPos.offsetPhysical)
    Me.typeInfo = rng.Offset(0, srcPos.offsetType)
    Me.required = rng.Offset(0, srcPos.offsetRequired)
    Me.min = rng.Offset(0, srcPos.offsetMin)
    Me.max = rng.Offset(0, srcPos.offsetMax)
    Me.enumValue = rng.Offset(0, srcPos.offsetEnum)

End Sub


