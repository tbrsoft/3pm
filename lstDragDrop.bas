Attribute VB_Name = "lstDragDrop_bas"
'------------------------------------------------------------------
'Módulo para hacer Drag&Drop en los ListBox             (17/Jul/98)
'(para usar en el programa csRadio)
'
'©Guillermo 'guille' Som, 1998-99
'
'
'Microsoft Knowledge Base - Article ID: Q167746
'HOWTO: Arrange Order of List Items within ListBox Control
'
'This article contains sample code that illustrates how to arrange
'the order of items within ListBox Control using drag-and-drop.
'
'Para usarlo:
'
'    'Para cagar el icono:
'    List1.DragIcon = LoadPicture(App.Path & "\Files10.ico")
'
'    Private Sub List1_DragDrop(Source As Control, X As Single, Y As Single)
'        ListRowMove Source, DragIndex, ListRowCalc(Source, Y)
'    End Sub
'
'    Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'        'Si quieres usar el botón izquierdo
'        If Button = vbLeftButton Then
'        'Si quieres usar el botón derecho
'        'If Button = vbRightButton Then
'            DragIndex = ListRowCalc(List1, Y)
'            List1.Drag
'        End If
'    End Sub
'
'------------------------------------------------------------------
Option Explicit

Global DragIndex As Integer

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Long) As Long

Public Function ListRowCalc(lstTemp As Control, ByVal Y As Single) As Integer
#If Win16 Then
    Const WM_USER = &H400
    Const LB_GETITEMHEIGHT = (WM_USER + 34)
#Else
    Const LB_GETITEMHEIGHT = &H1A1
#End If
    'Determines the height of each item in ListBox control in pixels
    Dim ItemHeight As Long
    ItemHeight = SendMessage(lstTemp.hWnd, LB_GETITEMHEIGHT, 0, 0)
    ListRowCalc = min(((Y / Screen.TwipsPerPixelY) \ ItemHeight) + _
                  lstTemp.TopIndex, lstTemp.ListCount - 1)
    
    'Seleccionar el elemento a mover
    lstTemp.ListIndex = ListRowCalc
End Function

Private Function min(X As Integer, Y As Integer) As Integer
    If X > Y Then min = Y Else min = X
End Function

Public Sub ListRowMove(lstTemp As Control, ByVal OldRow As Integer, _
                ByVal NewRow As Integer)
    Dim SaveList As String, i As Integer
    
    If OldRow = NewRow Then Exit Sub
    SaveList = lstTemp.List(OldRow)
    If OldRow > NewRow Then
        For i = OldRow To NewRow + 1 Step -1
            lstTemp.List(i) = lstTemp.List(i - 1)
        Next i
    Else
        For i = OldRow To NewRow - 1
            lstTemp.List(i) = lstTemp.List(i + 1)
        Next i
    End If
    lstTemp.List(NewRow) = SaveList
    'Seleccionar el elemento dejado
    lstTemp.ListIndex = NewRow
End Sub

