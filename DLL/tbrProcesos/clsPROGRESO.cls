VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPROGRESO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Sub Class_Initialize()
    'lo dejo cargado...
    frmPorcesos.Show
    'y escondido
    HideProgreso
End Sub

Private Sub Class_Terminate()
    Unload frmPorcesos
End Sub

Public Sub ShowProgreso(Titulo As String, PORC As Long)
    'frmPorcesos.Visible = True
    frmPorcesos.Top = Screen.Height / 2 - frmPorcesos.Height / 2
    frmPorcesosLeft = Screen.Width / 2 - frmPorcesos.Width / 2
    frmPorcesos.Refresh
    frmPorcesos.ZOrder
    frmPorcesos.tbrPROGRESO1.ShowProceso Titulo, PORC
End Sub

Public Sub HideProgreso()
    'lo dejo fuera de la pantalla
    frmPorcesos.Top = -frmPorcesos.Height * 2
End Sub
