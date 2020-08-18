Attribute VB_Name = "ModTbrBlue"
Public Const WM_SETTEXT As Long = &HC 'Constant defining the value of the WM_SETTEXT message. Used in WndProc.
Public Const WM_USER As Long = &H400 'Constant defining the value of the WM_USER message. Used in WndProc.
Private Const WM_COPYDATA = &H4A 'Constant defining the value of the WM_COPYDATA message. Used in WndProc.
Public Const GWL_WNDPROC = -4 'Used in Form_Load to specify to SetWindowLong that we wish to overwrite the Window Procedure of the frmMain form.

'@@@@@@ API FUNCTIONS USED @@@@@@
'This is the key API function for subclassing. With this function you can subclass any window
'in the Operating System, provided that you have the proper handle and maybe permissions.
'This procedure is also used for countless other various things, but right here, we use it for subclassing.
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal HWND As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Another vital element in subclassing is this API function. It will call the default window procedure
'of a certain window. We provide a pointer to that window procedure along with the window handle
'and message data (msg, lparam, wparam). More details in WndProc.
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal HWND As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Used to grab a block of memory at a certain address. Very useful when we'll need to take a number
'of bytes from some memory location and put them in a Visual Basic structure or primitive data type.
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
'Measures the lenght of a string.
Public Declare Function lstrlen Lib "kernel32" (ByVal lpString As Long) As Long

Public sarrStrings() As String 'Array of strings used with the WM_COPYDATA message.



Public Type InquiryStructure
 vAddress As Variant 'Strings can be passed as Unicode via Variant and the _variant_t type from VC++.
 vName As Variant 'Strings can be passed as Unicode via Variant and the _variant_t type from VC++.
 lPairStatus As Long
End Type

Private Declare Function TbrBt_GetInqReport Lib "tbrBlueC.dll" () As InquiryStructure
Private Declare Sub TbrBT_PushObject Lib "tbrBlueC.dll" (ByVal addr As String, ByVal path As String)
Public Declare Sub TbrBT_RegisterCallBack Lib "tbrBlueC.dll" (ByVal AppName As String)

Public loOriginalWindowProcedureAddress As Long 'Used to memorize the address of the default window procedure of frmMain.
Public mBtManager As TbrBtManager 'por si no lo usa el cliente

Public Sub UsarBluetoothLocal()
    Set mBtManager = New TbrBtManager
End Sub

Public Property Get btManagerLocal()
    Set btManagerLocal = mBtManager
End Property
Private Sub GetInqReport()
    Dim aux As InquiryStructure
    aux = TbrBt_GetInqReport()
    mBtManager.BtM_CBK_inquiry aux
End Sub

Public Sub CBK_inquiry(ByRef p As InquiryStructure)
    mBtManager.BtM_CBK_inquiry p
End Sub


Public Function WndProc(ByVal HWND As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'The parameters are easy enough. hWnd is the handle of our frmMain window, uMsg is a message identifier.
'wParam and lParam are used to transmit message-specific data. Check MSDN to see what they mean
'for each message. We'll only care about WM_TEXT, WM_USER and WM_COPYDATA.
Dim sString As String 'To grab any string content from the message.
Dim sLen As Long 'For the lenght of the string.

    'If we get a WM_SETTEXT text message or WM_USER + 241, writing the text contained in the lParam in the
    'debug textbox. Using both these message simply to better demonstrate that system messages can also be used
    'for our own purposes just as well as a customly defined message such as WM_USER + 241.
    If (uMsg = WM_SETTEXT) Or (uMsg = WM_USER + 241) Then
        sLen = lstrlen(lParam) 'Getting the lenght of the lParam.
        sString = Space$(sLen) 'Creating a buffer... filling the sString variable with spaces.
        CopyMemory ByVal sString, ByVal lParam, sLen 'Grabbing the string from lParam and putting it in sString.
        'Now we successfully transported the data content in this string from VC++ to Visual Basic using
        'a subclassing trick. Not very usefull considering that there are other methods,
        'but this still has a highly educative value.
      '  MsgBox "Received: " & sString
        'We don't want to allow the default window procedure to process any WM_SETTEXT message 'cause
        'it would change the caption of the window. The WM_USER + 241 message can be forwarded to the
        'default window procedure. It's all the same if it is or not, since the window procedure isn't doing anything
        'with such a message anyway.
        Select Case sString
            Case "INQ_REPORT"
                GetInqReport
            Case "INQ_FINISH"
                mBtManager.BtM_CBK_inqFinsish
            Case "PUSH_SUCCESS"
                mBtManager.BtM_CBK_PushReport True, "Success"
            Case "PUSH_FAILURE"
                mBtManager.BtM_CBK_PushReport False, "Fallo General"
            Case "PUSH_CHECK_FAILURE"
                mBtManager.BtM_CBK_PushReport False, "Fallo Al comprobar el servicio"
            Case "STATUS_INCOMING_CONNECT"
            Case "STATUS_INCOMING_DISCONNECT"
            Case "STATUS_OUTGOING_CONNECT"
            Case "STATUS_OUTGOING_DISCONNECT"
            Case "STATUS_BLUETOOTH_STOPED"
            Case "STATUS_BLUETOOTH_STARTED"
            Case "STATUS_BT_INVALID_PARAMETER"
            
        End Select
        
        
        If (uMsg = WM_SETTEXT) Then Exit Function
    End If

   'Now that we did our stuff with the messages, we can call the default window procedure.
   'Try commenting this line and execute the Application. Don't forget to save your VB work first : ).
    WndProc = CallWindowProc(loOriginalWindowProcedureAddress, HWND, uMsg, wParam, lParam)
End Function




