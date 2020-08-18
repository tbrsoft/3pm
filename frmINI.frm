VERSION 5.00
Begin VB.Form frmINI 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   Icon            =   "frmINI.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmINI.frx":030A
   ScaleHeight     =   6315
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Show
    Me.Refresh
    AP = App.Path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    'ver si ya estaba cargado
    If App.PrevInstance Then MsgBox "No se pueden abrir dos instancias de 3pm": End
    'caragar al inicio
    Dim carps As String
    carps = AP + App.EXEName + ".exe"
    'crear una clave en el registro para que se carge al inicio del arranque de windows
    SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "3pm", carps
    'eperar 4 segundos mostrando el logo
    t = Timer
    Do While Timer < t + 2
        DoEvents
    Loop
    'transformarse en fondo negro
    Me.Hide
    Me.Picture = LoadPicture
    Me.WindowState = vbMaximized
    Me.BackColor = vbBlack
    Me.Show
    frmINDEX.Show 1
    
End Sub
