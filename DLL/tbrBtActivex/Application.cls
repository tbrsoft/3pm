VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Property Get btManager() As TbrBtManager
    Set btManager = btManagerLocal
End Property

Public Sub UsarBluetooth()
    UsarBluetoothLocal
End Sub

Public Sub SetWindowMsg(ByVal FormHWnd As Long)

    loOriginalWindowProcedureAddress = SetWindowLong(FormHWnd, GWL_WNDPROC, AddressOf WndProc)
End Sub

Public Sub ResetWindowMsg()
     Call SetWindowLong(FormHWnd, GWL_WNDPROC, loOriginalWindowProcedureAddress)
End Sub
