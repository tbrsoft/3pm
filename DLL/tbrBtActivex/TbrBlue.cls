VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TbrBtDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Sub TbrBT_PushObject Lib "tbrBlueC.dll" (ByVal addr As String, ByVal path As String)
Private Declare Function TbrBt_GetDataSent Lib "tbrBlueC.dll" (ByVal addr As String) As Long
'Estos son los campos de la estructura
Public uAddr As Variant
Public addr As String
Public name As String
Public PairStatus As Long
Public lFileLen As Long

Public Function DecToHex(ByVal DecNum As Double) As String
    Dim remainder As Integer
    Dim HexStr As String
    HexStr = ""
    Do While DecNum <> 0
        remainder = DecNum Mod 16
        If remainder <= 9 Then
            HexStr = Chr(Asc(remainder)) & HexStr
        Else
            HexStr = Chr(Asc("A") + remainder - 10) & HexStr
        End If
        DecNum = DecNum \ 16
    Loop
    If HexStr = "" Then HexStr = "0"
    DecToHex = HexStr
End Function

Public Function GetDataSent() As Long
    GetDataSent = TbrBt_GetDataSent(uAddr)
End Function

Public Function GetDataSentPorc() As Integer
    Dim aux As Long
    aux = TbrBt_GetDataSent(uAddr) * 100 / lFileLen
    If aux > 99 Then aux = 99
    GetDataSentPorc = aux
End Function


Public Function getAddress() As String
    Dim aux As String
    Dim auxDir As Integer
    Dim lenUAddr As Integer
    Dim I As Integer
    lenUAddr = Len(uAddr)
    For I = 5 To lenUAddr Step -1
    aux = aux + "00:"
    Next

    For I = lenUAddr To 1 Step -1
        auxDir = Asc(Mid(uAddr, I, 1))
        If Len(DecToHex(auxDir)) < 2 Then
            aux = aux + "0" + DecToHex(auxDir)
         
        Else
            aux = aux + DecToHex(auxDir)
        End If
        If I > 1 Then
             aux = aux + ":"
        End If
    Next
    getAddress = aux
        
End Function

Public Function getPairStatus() As String
    If PairStatus = 0 Then
        getPairStatus = "No"
    Else
        getPairStatus = "Si"
    End If
End Function

Public Sub push(path As String)
'reimplementar esta funcion para que nunca tenga problemas con el path del archivo a enviar
lFileLen = FileLen(path)
If Len(path) < 256 Then
     TbrBT_PushObject uAddr, path
Else
    MsgBox ("EL path es demasiado largo :(")
End If

End Sub

Public Sub setAddr(addr As Variant)
    uAddr = addr
End Sub
