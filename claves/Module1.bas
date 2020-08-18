Attribute VB_Name = "Module1"
Public vW As New clsWindowsVERSION

Sub main()
    Dim V As vWindows
    'esta es la primera y lo calcula, despues solo lo lee de la _
        propiedad version
    'queda como global el vW
    V = vW.GetVersion
    Form1.Show 1
End Sub
