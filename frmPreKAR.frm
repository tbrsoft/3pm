VERSION 5.00
Begin VB.Form frmPreKAR 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11265
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPreKAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function WaitOk(sCancion As String) As Long
    
    Dim sCancion2 As String
    sCancion2 = SystemF + "nowpl.mas"
    
    'VER SI ESTA ENCRIPTADO O NO!
    If LCase(Right(sCancion, 3)) = "mn1" Then
        'desencriptarlo!
        F1.MP3.doTem True, _
            "salimos a romper culos tbrkar " + _
            "024154646213546054613463136543613612361341646585154" + _
            "968899996665464668166", sCancion, sCancion2
    Else
        FSO.CopyFile sCancion, sCancion2, True
    End If
    
    Dim R As Long
    R = F1.MP3.DoOpenKar(sCancion2, picKAR, shKAR)
    
    If R = 1 Then
        WaitOk = 1
        Exit Function
    End If
    
    WaitOk = 0
    picKAR.Visible = False
    EsperoOkKar = True
End Function
