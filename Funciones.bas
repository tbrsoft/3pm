Attribute VB_Name = "Funciones"
'--------------------------------------------
'GOBALES
Public AP As String
Public EstoyHaciendo As String
'define que estoy haciendo para saber quetiene que hacer los botones de movimiento
Public nDiscoElegido As Integer
'es solo de 1 a 4 para los 4 discos de la presentacion actual
'colores del lblGrupoDiscoElegido
Public Const BColorGrupo = &HFF8080
Public Const FColorGrupo = &H0&
Public Const BColorGrupoSel = &HFFFF&
Public Const FColorGrupoSel = &HFF&
'colores del lblNombreDiscoElegido
Public Const BColorNombre = &H800000
Public Const FColorNombre = &HFFFFFF
Public Const BColorNombreSel = &HFFFF&
Public Const FColorNombreSel = &HFF&

Public FSO As New Scripting.FileSystemObject
'-----------------------------------
'Conection y recordset
Public CN As New ADODB.Connection
Public rsDiscos As New ADODB.Recordset
Public rsTemas As New ADODB.Recordset
Public rsEstilo As New ADODB.Recordset


Public Sub MoverIZQ()
    Select Case EstoyHaciendo
        Case "ViendoDiscos"
            UnSelDisco nDiscoElegido
            nDiscoElegido = nDiscoElegido - 1
            If nDiscoElegido = 0 Then nDiscoElegido = 4
            ElegirDisco nDiscoElegido, frmCDS.SHsel
    End Select
            
End Sub

Public Sub MoverDER()
    Select Case EstoyHaciendo
        Case "ViendoDiscos"
            UnSelDisco nDiscoElegido
            nDiscoElegido = nDiscoElegido + 1
            If nDiscoElegido = 5 Then nDiscoElegido = 1
            ElegirDisco nDiscoElegido, frmCDS.SHsel
    End Select
            
End Sub


Public Sub ElegirDisco(n As Integer, SH As Shape)
    SH.Visible = False
    SH.Top = frmCDS.TapaCD(n).Top
    SH.Left = frmCDS.TapaCD(n).Left
    SH.Width = frmCDS.TapaCD(n).Width
    SH.Height = frmCDS.TapaCD(n).Height
    'retocar los lbls
    frmCDS.lblGDE(n).BackColor = BColorGrupoSel
    frmCDS.lblGDE(n).ForeColor = FColorGrupoSel
    frmCDS.lblNDE(n).BackColor = BColorNombreSel
    frmCDS.lblNDE(n).ForeColor = FColorNombreSel
    
    SH.Visible = True
End Sub

Public Sub UnSelDisco(n As Integer)
    'retocar los lbls
    frmCDS.lblGDE(n).BackColor = BColorGrupo
    frmCDS.lblGDE(n).ForeColor = FColorGrupo
    frmCDS.lblNDE(n).BackColor = BColorNombre
    frmCDS.lblNDE(n).ForeColor = FColorNombre
End Sub

Public Sub crgRSinCMB(rs As ADODB.Recordset, indiceFIELD As Integer, cmb As ComboBox)
    If rs.State = adStateClosed Then rs.Open
    If rs.RecordCount = 0 Then MsgBox "No hay registros": Exit Sub
    rs.MoveFirst
    cmb.Clear
    Do While Not rs.EOF
        If rs.Fields(indiceFIELD) <> "" Then cmb.AddItem rs.Fields(indiceFIELD)
        rs.MoveNext
    Loop
    cmb.ListIndex = 0
End Sub

Public Sub crgRSinLST(rs As ADODB.Recordset, LST As ListBox, Col1 As Integer, Col2 As Integer, Col3 As Integer)
    If rs.State = adStateClosed Then rs.Open
    If rs.RecordCount = 0 Then MsgBox "No hay registros": Exit Sub
    rs.MoveFirst
    LST.Clear
    Do While Not rs.EOF
        If Col2 = -1 Then LST.AddItem rs.Fields(Col1)
        If Col2 > -1 And Col3 = -1 Then LST.AddItem rs.Fields(Col1) + " - " + rs.Fields(Col2)
        If Col3 > -1 Then LST.AddItem rs.Fields(Col1) + " - " + rs.Fields(Col2) + " - " + rs.Fields(Col3)
        rs.MoveNext
    Loop
    LST.ListIndex = 0
End Sub

Public Function NewIdDisco() As String
    'genera un codigo alfanumérico de 7 cifras y comprueba que no este ya cargado
    '48-57 // 65-90 // 97-122
    Dim ID As String, CH As Integer, chACT As String
    Do
        c = 0: ID = ""
        Do While Len(ID) < 10
            Randomize Timer
            CH = Int(Rnd * 123)
            chstr = ""
            If CH > 47 And CH < 58 Then chACT = Chr(CH)
            If CH > 64 And CH < 91 Then chACT = Chr(CH)
            If CH > 96 And CH < 123 Then chACT = Chr(CH)
            ID = ID + chACT
        Loop
        'ver que no exita
        Dim YaEsta As Boolean
        With rsDiscos
            If .State = adStateClosed Then .Open
            If .RecordCount = 0 Then Exit Do
            If .RecordCount = -1 Then MsgBox "Error de recordset": Exit Function
            .MoveFirst
            YaEsta = False
            Do While Not .EOF
                If !IdDisco = ID Then YaEsta = True
                .MoveNext
            Loop
            If YaEsta = False Then Exit Do
        End With
    Loop
    NewIdDisco = ID
End Function

Public Function ArchivosInFolder(EXT As String, Carp As String) As Integer
    'Carp debe tener la barra invertida final
    Dim Arch As String, c As Integer
    c = 0
    Arch = Dir(Carp + "*." + EXT)
    Do While Arch <> ""
        c = c + 1
        Arch = Dir
    Loop
    ArchivosInFolder = c
End Function

Public Sub crgArchInLst(LST As ListBox, EXT As String, Carp As String)
    'Carp debe tener la barra invertida final
    Dim Arch As String, c As Integer
    c = 0
    Arch = Dir(Carp + "*." + EXT)
    Do While Arch <> ""
        Arch = Left(Arch, Len(Arch) - 4)
        LST.AddItem Arch
        c = c + 1
        Arch = Dir
    Loop
End Sub

Public Sub DuplicarLST(lstORIG As ListBox, lstDEST As ListBox)
    lstDEST.Clear
    Dim c As Integer
    c = 0
    Do While c < lstORIG.ListCount
        lstDEST.AddItem lstORIG.List(c)
        c = c + 1
    Loop
End Sub













