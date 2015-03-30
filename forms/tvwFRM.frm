VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form tvwFRM 
   Caption         =   "Objects"
   ClientHeight    =   6315
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4395
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   4395
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3150
      Top             =   1740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tvwFRM.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tvwFRM.frx":0A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tvwFRM.frx":0A70
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tvwFRM.frx":0ACE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog GuardarAbrir 
      Left            =   3210
      Top             =   510
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar barraTVW 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   5985
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1244
            MinWidth        =   1235
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1244
            MinWidth        =   1235
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1244
            MinWidth        =   1235
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar menuTVW 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "filtro_abrir"
            Object.ToolTipText     =   "Abrir filtro"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "filtro_guardar"
            Object.ToolTipText     =   "Guardar filtro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "filtro_aplicar"
            Object.ToolTipText     =   "Aplicar filtro"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "filtro_eliminar"
            Object.ToolTipText     =   "Eliminar filtro"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwStruct 
      Height          =   5985
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   10557
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList2"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3150
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Menu mnuNiveles 
      Caption         =   "mnuNiveles"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuAddNivel 
         Caption         =   "Nivel 0"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddNivel 
         Caption         =   "Nivel 1"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuAddNivel 
         Caption         =   "Nivle 2"
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddNivel 
         Caption         =   "Nivel 3"
         Enabled         =   0   'False
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddNivel 
         Caption         =   "Nivel 4"
         Enabled         =   0   'False
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddNivel 
         Caption         =   "Nivel 5"
         Enabled         =   0   'False
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddNivel 
         Caption         =   "Nivel 6"
         Enabled         =   0   'False
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddNivel 
         Caption         =   "Nivel 7"
         Enabled         =   0   'False
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddNivel 
         Caption         =   "Nivel 8"
         Enabled         =   0   'False
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddNivel 
         Caption         =   "Nivel 9"
         Enabled         =   0   'False
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddNivel 
         Caption         =   "Nivel 10"
         Enabled         =   0   'False
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddNivel 
         Caption         =   "Nivel 11"
         Enabled         =   0   'False
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddNivel 
         Caption         =   "Nivel 12"
         Enabled         =   0   'False
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddNivel 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu ModEli 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu ModEli 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   2
      End
   End
End
Attribute VB_Name = "tvwFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'para guardar el boton del mouse clickeado en el TVW
Dim intBotonMouse As Integer

'contadores de nivel, de campos filtrados y de menus
Dim contNivel, contFilt, contMenu As Integer

'guarda los id para levantar el filtro del tvw desde el ini
Dim strIdTVW As String

'graba cada vez que se expande o contrae un nodo
Dim strExpandNodo As String

'guarda filtros
Dim arrFiltros As Variant

'array que guarda los datos de cada nodo
Dim arrNodo() As Variant

'array que guarda los datos de dsiTVWconfig
Dim arrMenu() As Variant

Dim m_dsiNodosSeleccionados As Integer

'ID y vista para tvw
Dim m_IDopc As String
Dim m_vista As String
 
Public Property Let dsiNodosSeleccionados(intI As Integer)
  m_dsiNodosSeleccionados = intI
End Property

Public Property Get dsiNodosSeleccionados() As Integer
  dsiNodosSeleccionados = m_dsiNodosSeleccionados
End Property
 
Public Property Let dsiIDopc(ID As String)
  m_IDopc = ID
End Property

Public Property Get dsiIDopc() As String
  dsiIDopc = m_IDopc
End Property
 
Public Property Let dsiVista(ID As String)
  m_vista = ID
End Property

Public Property Get dsiVista() As String
  dsiVista = m_vista
End Property
 
Public Property Get dsiCondicionGuardar() As Boolean
    
  Dim blnB As Boolean
  Dim strT As String
  
  If strIdTVW <> "" Then 'si hay filtro
    
    'cambio puntero mouse
    Screen.MousePointer = vbHourglass
    
    'set filtro para que solo permita abrir los filtros de grilla actual
    strT = "Archivos de texto (*.sql)|" & "*.sql"
    Me.GuardarAbrir.Filter = strT
    
    'titulo de ventana
    Me.GuardarAbrir.DialogTitle = "Guardar Filtro como..."
  
    'set nombre de filtro
    strT = "*.sql"
    Me.GuardarAbrir.FileName = strT
    
    'abro cuadro de dialogo
    Me.GuardarAbrir.ShowSave
      
    'si cancelar salgo
    If Me.GuardarAbrir.FileName = "" Or Me.GuardarAbrir.FileName = "*.sql" Then
      'recupero puntero mouse
      Screen.MousePointer = vbDefault
      Exit Sub
    End If

    'guardo condicion
    blnB = WriteIni("filtro", "", strIdTVW, Me.GuardarAbrir.FileName)
  
    'aviso de operacion OK
    intRes = MsgBox("La operación se realizó con éxito.", vbInformation + vbOKOnly, "Atención...")
    
    'recupero puntero mouse
    Screen.MousePointer = vbDefault
        
  Else
    
    intRes = MsgBox("No existe ningun filtro para guardar", vbExclamation, "Atencion!")
  
  End If
   
End Property

Public Property Get dsiCondicionAbrir() As Boolean
    
  Dim blnB As Boolean
  Dim strT As String
      
  Dim strKey As String
  Dim intCuenta As Integer
  Dim arrKey, arrIdTvw As Variant
  
 'On Error GoTo Errores

  'set filtro para que solo permita abrir los filtros de grilla actual
  strT = "Archivos de texto (*.sql)|" & "*.sql"
  Me.GuardarAbrir.Filter = strT
    
  'set titulo de ventana
  Me.GuardarAbrir.DialogTitle = "Abrir Filtro..."
  
  'set tiene que seleccionar un nombre
  Me.GuardarAbrir.FileName = ""
  
  'abro cuadro de dialogo
  Me.GuardarAbrir.ShowOpen
      
  'si cancelar salgo
  If Me.GuardarAbrir.FileName = "" Then
    Exit Property
  End If

  'leo del ini el filtro que se eligio
  strKey = ReadIni("filtro", "", Me.GuardarAbrir.FileName)
  
  'cuando levanta del ini no trae las comillas simples de los extremos...
  If strKey <> "" And Left(strKey, 1) <> "'" Then
    strKey = "'" & strKey & "'"
  End If
  
  arrKey = separateText(strKey)
         
  'si no hay ningun nodo marcado (no hay filtro)
  If strIdTVW = "" Then
  
    'aplico filtro guardado
    For intCuenta = 1 To UBound(arrKey, 1)
    
      'le coloco el check al nodo y lo hago visible
      tvwStruct.Nodes.Item(arrKey(intCuenta)).Checked = True
      tvwStruct.Nodes.Item(arrKey(intCuenta)).EnsureVisible
      
      'sumo uno para la barra
      SelectNodos = SelectNodos + 1
      
    Next intCuenta
    
    strIdTVW = strKey
    
  Else
    
    'si el twv tiene algo filtrado pregunto si se desea agregar
    intRes = MsgBox("Desea agregar el filtro...", vbQuestion + vbYesNo, "Atención...")
      
    'si se confirma que se agrega a lo que ya tiene
    If intRes = vbYes Then
    
      For intCuenta = 1 To UBound(arrKey, 1)
      
        If tvwStruct.Nodes.Item(arrKey(intCuenta)).Checked = False Then
        
          'le coloco el check al nodo y lo hago visible
          tvwStruct.Nodes.Item(arrKey(intCuenta)).Checked = True
          tvwStruct.Nodes.Item(arrKey(intCuenta)).EnsureVisible
          
          'sumo uno para la barra
          SelectNodos = SelectNodos + 1
        
        Else
          tvwStruct.Nodes.Item(arrKey(intCuenta)).EnsureVisible
        End If
        
        'si no existe lo agrega
        If InStr(1, strIdTVW, arrKey(intCuenta), 1) = 0 Then
          strIdTVW = strIdTVW & ";" & arrKey(intCuenta)
        End If
                
      Next intCuenta
        
    Else 'si no se desea agregar al filtro anterior
      
      arrIdTvw = separateText(strIdTVW)
      
      'borro todos los check del tree
      For intCuenta = 1 To UBound(arrIdTvw, 1)
        tvwStruct.Nodes.Item(arrIdTvw(intCuenta)).Checked = False
      Next
      
      SelectNodos = 0
      
      'borro los filtros seleccionados
      strIdTVW = ""
 
      'agrego el nuevo filtro
      For intCuenta = 1 To UBound(arrKey, 1)
        
        'le coloco el check al nodo y lo hago visible
        tvwStruct.Nodes.Item(arrKey(intCuenta)).Checked = True
        tvwStruct.Nodes.Item(arrKey(intCuenta)).EnsureVisible
          
        'sumo uno para la barra
        SelectNodos = SelectNodos + 1
        
      Next intCuenta
       
      strIdTVW = strKey
      
      'recupero puntero mouse
      Screen.MousePointer = vbDefault
      
    End If
    
  End If
  
  'Call GuardarFiltros(strIdTVW)
  
  Form_Activate
  
  'muestro en la barra de estado
  barraTVW.Panels(1).AutoSize = sbrContents
  barraTVW.Panels(2).Text = "Total: " & tvwStruct.Nodes.Count & " "
  barraTVW.Panels(2).AutoSize = sbrContents
  barraTVW.Panels(3).Text = "Seleccionados: " & SelectNodos & " "
  barraTVW.Panels(3).AutoSize = sbrContents
  
  Exit Sub

Errores:

If Err.Number = 35601 Then 'elemento no encontrado en el treeview

  intRes = MsgBox("Uno de los elementos seleccionados ya no existe." & Chr(13) & Chr(13) & "Se cargara el filtro nuevamente.", vbApplicationModal + vbCritical + vbOKOnly, "Atención!")
  
  'borro el numero de nodo que ya no existe
  strKey = Replace(strKey, arrKey(intCuenta), "")
  'si queda dos ; juntas las remplaza por una sola
  strKey = Replace(strKey, ";;", ";")
  'si quedo ; a la derecha la saca
  If Right(strKey, 1) = ";" Then
    strKey = Left(strKey, Len(strKey) - 1)
  Else
    'si quedo ; a la izquierda la saca
    If Left(strKey, 1) = ";" Then
        strKey = Right(strKey, Len(strKey) - 1)
    End If
  End If
    
  'vuelvo a grabar el filtro
  blnB = WriteIni("filtro", "", strKey, Me.GuardarAbrir.FileName)
   
  'aplico el filtro
  blnB = Me.dsiCondicionAbrir
    
Else
  On Error GoTo 0
End If

End Property

Private Function BuscoIdKey() As String
'Funcion que busca en la tabla Structure el último id para agregar un nodo

  Dim rs As ADODB.Recordset
  
  'selecciono el último id de la tabla structure para el nuevo nodo
  strSQL = "select max(Id) as ultimo from dsiTVW"
  Set rs = SQLexec(strSQL)
  
  If rs.EOF Then
    BuscoIdKey = "'0'"
  Else
  'guardo el último id + 1 (para el siguiente)
  BuscoIdKey = "'" & rs!ultimo + 1 & "'"
  End If
  rs.Close
  
End Function

Private Sub AgregaNivel(strTitulo As String, IndexMenu As Integer) ', strWhere As String)
'agrega un nodo al tree, únicamente el ultimo nimvel (ej: pozo)
'se puede agregar a cualquier nivel
'cada nodo se agrega a la tabal dsiTVW y a su correspondiente tabla
'Los campos son: IdTVW(int:4),Id(inte:4),Nombre(nvarchar:50),
'                Referencia(nvarchar:50),Icono(nvarchar:50),Filtra(bit)

  Dim strSQL, strDato, strIdKey, strNodo, nodPadre As String
  Dim strPadre As Integer
  Dim nodo As Node 'agrega nodos
  Dim arrTag As Variant 'guarda los valores de Tag
  Dim imagen As ListImage 'agrega iconos al nodo
  Dim strSP As String 'indica el store
  
   
  strDato = Trim(InputBox("", strTitulo, , 3000, 3000))
  If strDato = "" Then Exit Sub
    'confirmacion
    intRes = MsgBox("Esta seguro que desea agregar " & strDato, vbQuestion + vbYesNo, "confirmación...")
   
  'si se confirma
  If intRes = vbYes Then
  
    'cambio puntero mouse
    Screen.MousePointer = vbHourglass
      
    'Busco el último id para guardar en key
    strIdKey = Replace(BuscoIdKey(), "'", "")
     
    'transformo el key en un int
    intPadre = Replace(tvwStruct.SelectedItem.Key, "'", "")
    
    'guardo en strSQL los datos para guardar en la tabla dsiTVW
    strSQL = Me.dsiIDopc & "," & strIdKey & "," & "'" & strDato & "'" & "," & "'" & arrMenu(IndexMenu, 3) & "'" & "," & Val(Replace(arrMenu(IndexMenu, 4), ".ico", "")) & "," & intPadre
    'agrego a strSQL los datos para guardar en la tabla de "pozos"
    strSQL = "exec " & arrMenu(IndexMenu, 6) & " " & strSQL
    
    'grabo en la tabla structure
    SQLexec (strSQL)
    
    'errores
    
    'le da las dimensiones, contiene tantas columnas como nodos y 5 filas
    '(1)Nivel, (2)Orden, (3)Descripción, (4)Icono, (5)Filtrar.
    ReDim Preserve arrNodo(5, tvwStruct.Nodes.Count + 1)
    
    'agrego nodo nuevo al tvw y armo IDStructura de MenuEstructura
    Set nodo = tvwStruct.Nodes.Add(tvwStruct.SelectedItem.Key, tvwChild, "'" & strIdKey & "'", strDato, Me.ImageList2.ListImages(Val(mnuAddNivel.Item(IndexMenu).Index)).Index)
      
    'lleno el array con los datos de cada nodo
    arrNodo(1, nodo.Index) = arrMenu(mnuAddNivel.Item(IndexMenu).Index, 1)
    arrNodo(2, nodo.Index) = arrMenu(mnuAddNivel.Item(IndexMenu).Index, 2)
    arrNodo(3, nodo.Index) = arrMenu(mnuAddNivel.Item(IndexMenu).Index, 3)
    arrNodo(4, nodo.Index) = arrMenu(mnuAddNivel.Item(IndexMenu).Index, 4)
    arrNodo(5, nodo.Index) = arrMenu(mnuAddNivel.Item(IndexMenu).Index, 5)

    
    'aviso de operacion OK
    intRes = MsgBox("La operación se realizó con éxito.", vbInformation + vbOKOnly, "Atención...")
          
    'se abre la rama donde se coloco el nodo
    tvwStruct.SelectedItem.Expanded = True
    'se selecciona el nuevo nodo
    tvwStruct.Nodes.Item(tvwStruct.Nodes.Count).Selected = True
    
    barraTVW.Panels(2).Text = "Total: " & tvwStruct.Nodes.Count & " "
    barraTVW.Panels(2).AutoSize = sbrContents
        
  End If
  
  'recupero el puntero del mouse
  Screen.MousePointer = vbDefault
  
  'de acuerdo al where recorre todas las grillas abiertas y le aplica el dataWhereGeneral
'  Dim frmOpen As Form
  
'  For Each frmOpen In Forms
'    If frmOpen.Name = "gridFRM" Then
'      If frmOpen.DataFiltraCampo <> "" Then
'        x = 0
'        Do
'          x = x + 1
'        Loop Until ("[" & arrFiltros(x, 1) & "]") = frmOpen.DataFiltraCampo
'        If arrFiltros(x, 3) <> "" Then
'          frmOpen.DataWhereTree = " [" & arrFiltros(x, 1) & "] " & arrFiltros(x, 3)
'        Else
'          frmOpen.DataWhereTree = ""
'        End If
'      End If
'      'si esta abierto lo activa
'      frmOpen.DataRefresh = True
'    End If
'  Next frmOpen

End Sub

Private Sub Form_Activate()

 Tree = True
 ' set botones menu invisibles
 ' MainMDI.tlbMenu.Buttons("graficos").Enabled = False
 ' MainMDI.tlbMenu.Buttons("insertar").Enabled = False
 ' MainMDI.tlbMenu.Buttons("ventanas").Enabled = True
 ' MainMDI.tlbMenu.Buttons("actualizar").Enabled = False
 ' MainMDI.tlbMenu.Buttons("editar").Enabled = False
 ' MainMDI.tlbMenu.Buttons("eliminar").Enabled = False
 ' MainMDI.tlbMenu.Buttons("buscar").Enabled = False
 ' MainMDI.tlbMenu.Buttons("rapido").Enabled = False
 ' MainMDI.tlbMenu.Buttons("avanzado").Enabled = False
 ' MainMDI.tlbMenu.Buttons("atras").Enabled = False
 ' MainMDI.tlbMenu.Buttons("excel").Enabled = False
 ' MainMDI.tlbMenu.Buttons("guardar").Enabled = True
 ' MainMDI.tlbMenu.Buttons("reportes").Enabled = False
 ' MainMDI.tlbMenu.Buttons("salir").Enabled = True
  
End Sub

Private Sub Form_Load()
    
  Dim rs As ADODB.Recordset
  Dim imagen As ListImage
  Dim nodo As Node
  Dim intCont As Integer
  Dim strFiltrados As String
  Dim arrFiltrados, arrKey, arrExpandNodo As Variant
  
  On Error GoTo 0
  
  'le indica al programa que el tvw está abierto
  Tree = True
  
  'pone en 0 para empezar a contar los nodos totales y los seleccionados
  SelectNodos = 0
  
  'cambio puntero mouse
  Screen.MousePointer = vbHourglass
  
  'pido los datos de dsiTVWconfig para llenar el menu
  strSQL = "select * from dsiTVWconfig where IDtvw = " & Me.dsiIDopc & " order by titulo"
  Set rs = SQLexec(strSQL)
    
  'check error
    
  'inicializo las variables para contar los niveles y los que se filtran
  contMenu = 0
  contFilt = 0
  
  'cargo los iconos en un ImageList y recorro y completo el menu contextual
  While Not rs.EOF
  
    If Not IsNull(rs!icono) Then
       Set imagen = ImageList2.ListImages.Add(, "'" & rs!ID & "'", LoadPicture(App.Path & "\iconos\well.ico"))
    End If
    
    'guardo la referencia de los campos que se filtran
'    If rs!filtrar Then
'      If strFiltrados = "" Then
'        strFiltrados = rs!Referencia
'      Else
'          strFiltrados = strFiltrados & "," & rs!Referencia
'      End If
'    End If
    
    'cuenta la cantidad de niveles
'    If contNivel < rs!nivel Then
'      contNivel = rs!nivel
'    End If
    
'    'cuenta la cantidad de campos que se filtran
'    If rs!filtrar Then
'      contFilt = contFilt + 1
'    End If
    
    rs.MoveNext
    
  Wend
  
  'llena treeView
  
  'puntero al primero
  rs.MoveFirst
    
  'le da las dimensiones, contiene tantas columnas como nodos y 5 filas
  '(1)Nombre, (2)Referencia, (3)Titulo, (4)Valor, (5)Icono
  ReDim arrNodo(5, rs.RecordCount)
  
  'recorro niveles
  While Not rs.EOF
    
    'agrego nodos padres
    Set nodo = tvwStruct.Nodes.Add(, , "'" & rs!ID & "'", rs!titulo, ImageList2.ListImages(1).Index)
    nodo.Tag = rs!valor
            
    'lleno el array con los datos de cada nodo
    arrNodo(1, nodo.Index) = rs!nombre
    arrNodo(2, nodo.Index) = rs!referencia
    arrNodo(3, nodo.Index) = rs!titulo
    arrNodo(4, nodo.Index) = rs!valor
    arrNodo(5, nodo.Index) = rs!icono
   
    'puntero proximo
    rs.MoveNext
    
  Wend
            
  'cierro
  SQLclose
          
  'separa los campos que se filtran
  arrFiltrados = separateText(strFiltrados, ",")
  
  'le doy la dimension a la matriz
  'en el 1 guardo el nombreCampo, 2 campos filtrados, 3 in [campos filtrados]
'  ReDim arrFiltros(1 To (UBound(arrFiltrados)), 1 To 3) As String
  
  'le agrego el nombre de campo que se filtra a la matriz
'  For x = 1 To UBound(arrFiltros)
'    arrFiltros(x, 1) = arrFiltrados(x)
'  Next
  
'*******************************************************
  'definicion de variables de ubicacion
  Dim strUbico As Variant
  Dim sngLeft, sngTop, sngWidth, sngHeight As Single

  'busco en ini si hay propiedades de ubicacion top, left, width, height
  strUbico = keyIniToArray(Me.dsiIDopc, "ubicacion")
  sngLeft = Val(arrayGetValue(strUbico, "left"))
  sngTop = Val(arrayGetValue(strUbico, "top"))
  sngWidth = Val(arrayGetValue(strUbico, "width"))
  sngHeight = Val(arrayGetValue(strUbico, "height"))

  'cambio propiedades de ubicacion segun ini
  If sngLeft <> 0 Or sngTop <> 0 Or sngWidth <> 0 Or sngHeight <> 0 Then
    Me.Left = sngLeft
    Me.Top = sngTop
    Me.Width = sngWidth
    Me.Height = sngHeight
  Else
  'cambio tamaño y hubico form segun tamaño de MDI
    Me.Left = sngLeft
    Me.Top = sngTop
    Me.Width = sngWidth
    Me.Height = sngHeight
  End If
'********************************************

  'propiedades para treeview
  tvwStruct.Indentation = 15
  tvwStruct.SingleSel = False
  tvwStruct.Checkboxes = True
 
'******************************
 
  'leo del ini los nodos expandidos
  strExpandNodo = ReadIni(Me.dsiIDopc, "NodosExpand")
  'si no está vacio, le agrego comillas simples
  If strExpandNodo <> "" Then
    strExpandNodo = "'" & strExpandNodo & "'"
  End If
  
  'separo el string con los nodos que se expanden
  arrExpandNodo = separateText(strExpandNodo, ",")
  
  'expando los nodos
'  For x = 1 To UBound(arrExpandNodo)
'   tvwStruct.Nodes.Item(arrExpandNodo(x)).Expanded = True
'  Next
     

'****
 
 'leo en ini los filtros guardados para el tvw
  strIdTVW = ReadIni(Me.dsiIDopc, "IDtvw")
  
  'cuando levanta del ini no trae las comillas simples de los extremos...
  If strIdTVW <> "" And Left(strIdTVW, 1) <> "'" Then
    strIdTVW = "'" & strIdTVW & "'"
  End If
  
  arrKey = separateText(strIdTVW)
  
  
  'le aplico el filtro al tvw
'  For intCont = 1 To UBound(arrKey, 1)
  
'    tvwStruct.Nodes.Item(arrKey(intCont)).Checked = True
'    tvwStruct.Nodes.Item(arrKey(intCont)).EnsureVisible
    
'    SelectNodos = SelectNodos + 1
    
'  Next intCont
  
  
  'Call GuardarFiltros(strIdTVW)
     
  'de acuerdo al where recorre todas las grillas abiertas y le aplica el dataWhereGeneral
'  Dim frmOpen As Form

'  For Each frmOpen In Forms
'    If frmOpen.Name = "gridFRM" Then
'      If frmOpen.DataFiltraCampo <> "" Then
'        x = 0
'        Do
'          x = x + 1
'        Loop Until ("[" & arrFiltros(x, 1) & "]") = frmOpen.DataFiltraCampo
'        If arrFiltros(x, 3) <> "" Then
'          frmOpen.DataWhereTree = " [" & arrFiltros(x, 1) & "] " & arrFiltros(x, 3)
'        Else
'          frmOpen.DataWhereTree = ""
'        End If
'      End If
'      'si esta abierto lo activa
'      frmOpen.DataRefresh = True
'    End If
'  Next frmOpen
     
 
  'muestro info en barra de estado para que actualice cuando
  barraTVW.Panels(1).AutoSize = sbrContents
  barraTVW.Panels(2).Text = "Total: " & tvwStruct.Nodes.Count & " "
  barraTVW.Panels(2).AutoSize = sbrContents
  barraTVW.Panels(3).Text = "Seleccionados: " & SelectNodos & " "
  barraTVW.Panels(3).AutoSize = sbrContents
 
  'recupero puntero mouse
  Screen.MousePointer = vbDefault
 
 
Exit Sub

Errores:

If Err.Number = 35601 Then 'elemento no encontrado en el treeview, al pretender checkearlo
'ocurre cuando se borran pozos que estaban chekeados, por ende eran parte de filtro gral
'o cuando se cambia de base de datos
'con esto vacio el filro y cargo nuevamente el TreeView sin check's
  
  intRes = MsgBox("El Filtro contiene elementos que ya no existen." & Chr(13) & Chr(13) & "Se va a inicializar ahora.", vbApplicationModal + vbCritical + vbOKOnly)
  'graba en ini dataWhereGeneral (vacio)
  intRes = WriteIni(Me.dsiIDopc, "Where", "")
  intRes = WriteIni(Me.dsiIDopc, "dataKey", "")
  'vacio las variables que se graban en el ini
  strIdTVW = ""
  
'  tvwStruct.Nodes.Clear 'borra los nodos
  Call Form_Load 'llamo al evento Load de este mismo Form

Else
  On Error GoTo 0
End If

End Sub

Private Sub Form_Resize()
  
  'ajusto tamaño de treeView segun FRM
  tvwStruct.Top = menuTVW.Height
  tvwStruct.Left = 0
  tvwStruct.Width = Me.ScaleWidth
  If (barraTVW.Top - menuTVW.Height) > 0 Then
    tvwStruct.Height = barraTVW.Top - menuTVW.Height
  End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

  Dim intCont As Integer

  'al cerrarce el tvw guardo toda la informacion para recuperarlo de igual manera
  strValor = "left;" & Me.Left & ";top;" & Me.Top & ";width;" & Me.Width & ";height;" & Me.Height
  intRes = WriteIni(Me.dsiIDopc, "ubicacion", strValor)
  
  'guardo el filtro al momento de cerrar
  intRes = WriteIni(Me.dsiIDopc, "IDtvw", strIdTVW)
    
'  For intCont = 1 To UBound(arrFiltros)
'    intRes = WriteIni(Me.dsiIDopc, arrFiltros(intCont, 1), arrFiltros(intCont, 3))
'  Next
  
  'guarda en el ini los nodos expandidos
  intRes = WriteIni(Me.dsiIDopc, "NodosExpand", strExpandNodo)
    
  'le indica al programa que el tvw está cerrado
  Tree = False
    
'  For Each frmOpen In Forms
'    If frmOpen.Name = "gridFRM" Then
'      'si esta abierto lo activa
'      frmOpen.dsiWhereTree = ""
'      frmOpen.DataRefresh = True
'    End If
'  Next frmOpen
  
End Sub

Private Sub mnuAddNivel_Click(Index As Integer)
 
  Call AgregaNivel(mnuAddNivel.Item(Index).Caption, mnuAddNivel.Item(Index).Index)

End Sub

Private Sub ModEli_Click(Index As Integer)
'Modifica o elimina un nodo del tvw

  Dim strDato, strSP, strPadre As String
  Dim frmOpen As Form
  Dim intCont As Integer
  Dim intKey As Integer
  
  'si modifica index=1
  If Index = 1 Then
    
    'si se selecciono modificar pregunto cual es el nuevo nombre
    strDato = Trim(InputBox("", "Modificar " & tvwStruct.SelectedItem.Text, tvwStruct.SelectedItem.Text, 3000, 3000))
    
    'si no se ingreso nada o se ingreso lo mismo salgo
    If strDato = "" Then Exit Sub
    If strDato Like tvwStruct.SelectedItem.Text Then Exit Sub
      
    'sino pido la confirmacion
    intRes = MsgBox("Esta seguro que desea modificar " & tvwStruct.SelectedItem.Text & " por " & strDato, vbQuestion + vbYesNo, "confirmación...")
  
  Else
    
    'si elimina
    'si el nodo tiene hijos no puede eliminarse
    If Not tvwStruct.SelectedItem.Children = 0 Then
      intRes = MsgBox("No puede eliminarse este nodo.", vbInformation + vbOKOnly, "Atención...")
      Exit Sub
    End If
    
    'si se selecciono eliminar pido confirmacion
    intRes = MsgBox("Esta seguro que desea eliminar " & tvwStruct.SelectedItem.Text, vbExclamation + vbYesNo, "Atención...")
  
  End If
  
  'si se confirma
  If intRes = vbNo Then Exit Sub
  
  'cambio puntero mouse
  Screen.MousePointer = vbHourglass
   
  If Index = 1 Then
    strSP = arrMenu(arrNodo(1, tvwStruct.SelectedItem.Index), 7)
  Else
    'si elimino
    strSP = arrMenu(arrNodo(1, tvwStruct.SelectedItem.Index), 8)
  End If

  'id del padre y le saco las comillas simples
  nodPadre = Replace(tvwStruct.SelectedItem.Parent.Key, "'", "")

  'cambio (true=1 y false=0) porque sino no lo acepta sql
  If arrNodo(5, tvwStruct.SelectedItem.Index) Then
    arrNodo(5, tvwStruct.SelectedItem.Index) = 1
  Else
    arrNodo(5, tvwStruct.SelectedItem.Index) = 0
  End If
  
  intKey = Replace(tvwStruct.SelectedItem.Key, "'", "")
    
  
  'guardo en strSQL los datos para guardar en la tabla Structure
  strSQL = Me.dsiIDopc & "," & intKey & "," & "'" & strDato & "'" & "," & "'" & arrNodo(3, tvwStruct.SelectedItem.Index) & "'" & "," & arrNodo(4, tvwStruct.SelectedItem.Index) & "," & nodPadre
   
  'agrego a strSQL los datos para guardar en la tabla de "pozos"
  strSQL = "exec " & strSP & " " & strSQL
    
  'grabo en la tabla structure
  SQLexec (strSQL)
          
  'errores

          
  'si el nodo eliminado o modificado tenia check elimino sus "registros"
  If tvwStruct.SelectedItem.Checked = True Then
  
    'guarda id para el ini
    Call GuardarIni_Check(tvwStruct.SelectedItem.Key, False)
    
    'si el nodo no pertenece al último nivel, busco si tiene hijos
    If arrNodo(1, Node.Index) < contNivel Then
     
      'llamo al procedimiento para hacer lo mismo con los hijos
      Call BuscoHijos_NodeCheck(tvwStruct.SelectedItem, False)
    
    End If
     
  End If
    
  If Index = 1 Then
    
    'si modifico solo cambio el nombre del nodo
    tvwStruct.SelectedItem.Text = strDato
      
    If tvwStruct.SelectedItem.Checked = True Then
      
      'guarda id para el ini
      Call GuardarIni_Check(tvwStruct.SelectedItem.Key, True)
       
      'llamo al procedimiento para hacer lo mismo con los hijos
      Call BuscoHijos_NodeCheck(tvwStruct.SelectedItem, True)
        
    End If
      
  Else
      
    Call tvwStruct_Collapse(tvwStruct.SelectedItem)
    tvwStruct.Nodes.Remove (tvwStruct.SelectedItem.Index)
   
  End If
    
  'aviso de operacion OK
  intRes = MsgBox("La operación se realizó con éxito.", vbInformation + vbOKOnly, "Atención...")
    
  'recupero puntero mouse
  Screen.MousePointer = vbDefault
    
  
  'de acuerdo al where recorre todas las grillas abiertas y le aplica el dataWhereGeneral
  For Each frmOpen In Forms
    
    If frmOpen.Name = "gridFRM" Then
      
      If frmOpen.DataFiltraCampo <> "" Then
        intCont = 0
        Do
          intCont = intCont + 1
        Loop Until ("[" & arrFiltros(intCont, 1) & "]") = frmOpen.DataFiltraCampo
        
        If arrFiltros(intCont, 3) <> "" Then
          frmOpen.DataWhereTree = " [" & arrFiltros(intCont, 1) & "] " & arrFiltros(intCont, 3)
        End If
      
      End If
      
      'si esta abierto lo activa
      frmOpen.DataRefresh = True
    
    End If
  
  Next frmOpen
  
End Sub

Private Sub menuTVW_ButtonClick(ByVal Button As MSComctlLib.Button)
'cuando hago click en uno de los botones de menu del tvw

  Dim blnB As Boolean
  Dim frmOpen As Form
  Dim intCont As Integer

  
  'segun el boton que se selecciona
  Select Case Button.Key
  
  Case "filtro_eliminar"
    
    'si es distinto de "" borra los filtros y saca los check
    If strIdTVW <> "" Then
      arrKey = separateText(strIdTVW)
      
      'saca todos los check
      For intCuenta = 1 To UBound(arrKey, 1)
        tvwStruct.Nodes.Item(arrKey(intCuenta)).Checked = False
      Next

      strIdTVW = ""
      SelectNodos = 0
      
      
      Call GuardarFiltros(strIdTVW)

    End If
    
  Case "filtro_aplicar"
  
    Call GuardarFiltros(strIdTVW)
  
    'de acuerdo al where recorre todas las grillas abiertas y le aplica el dsiCondicionTree
    For Each frmOpen In Forms
      
      If frmOpen.Name = "gridFRM" Then
        
        'limpio whereTree
        frmOpen.dsiWhereTree = ""
        
'       If frmOpen.DataFiltraCampo <> "" Then
          intCont = 0
'         Do
            intCont = intCont + 1
'         Loop Until ("[" & arrfiltros(intCont, 1) & "]") = frmOpen.DataFiltraCampo
          
          If arrFiltros(intCont, 3) <> "" Then
            frmOpen.dsiWhereTree = " [" & arrFiltros(intCont, 1) & "] " & arrFiltros(intCont, 3)
'         Else
'           frmOpen.dsiCondicionTree = ""
          End If
'       End If
        
        'si esta abierto lo activa
        blnB = frmOpen.dsiRefresh()
      
      End If
    
    Next frmOpen
    
  Case "filtro_guardar"
  
     blnB = Me.dsiCondicionGuardar
  
  Case "filtro_abrir"
  
    blnB = Me.dsiCondicionAbrir
        
  End Select
  
  'muestro info en barra de estado para que actualice cuando
  barraTVW.Panels(3).Text = "Seleccionados: " & SelectNodos & " "
  barraTVW.Panels(3).AutoSize = sbrContents

End Sub


Private Sub tvwStruct_Collapse(ByVal Node As MSComctlLib.Node)
'cuando un nodo se contrae

  'saco el Id de strExpandNodo
  strExpandNodo = Replace(strExpandNodo, Node.Key, "")
  strExpandNodo = Replace(strExpandNodo, ",,", ",")
  
  'saco las comas "," que estan demas en el string
  If Right(strExpandNodo, 1) = "," Then
    strExpandNodo = Left(strExpandNodo, Len(strExpandNodo) - 1)
  Else
    
    If Left(strExpandNodo, 1) = "," Then
      strExpandNodo = Right(strExpandNodo, Len(strExpandNodo) - 1)
    End If
  
  End If
  
End Sub

Private Sub tvwStruct_Expand(ByVal Node As MSComctlLib.Node)
'cuando se expande un nodo

  'si expando el nodo y strExpandNodo esta vacío guardo el Id
  If strExpandNodo = "" Then
    
    strExpandNodo = Node.Key
    
  Else
    
    'agrego a strExpandNodo el Id
    If InStr(1, strExpandNodo, Node.Key, 1) = 0 Then
      strExpandNodo = strExpandNodo & "," & Node.Key
    End If
    
  End If

End Sub

Private Sub tvwStruct_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

  'guarda el boton del mouse, para usarlo en el evento NodeClick
  intBotonMouse = Button
  
End Sub

Private Sub GuardarFiltros(strIdTVW As String)

  Dim Cont As Integer
  Dim contFor As Integer
  Dim arrKey As Variant
  
  arrKey = separateText(strIdTVW)
  
  'limpio los filtros anteriores
  For Cont = 1 To UBound(arrFiltros)
    arrFiltros(Cont, 2) = ""
  Next
  
  For contFor = 1 To UBound(arrKey, 1)
   
    'si el campo se filtra lo agrego al array de filtros
    If arrNodo(5, tvwStruct.Nodes(arrKey(contFor)).Index) Then
    
      Cont = 0
      
      'busca hasta que coincide el nombre del nivel
      Do
        Cont = Cont + 1
      Loop Until arrFiltros(Cont, 1) = arrNodo(3, tvwStruct.Nodes(arrKey(contFor)).Index)
      
      
      'si esta vacío lo agrega
      If arrFiltros(Cont, 2) = "" Then
        arrFiltros(Cont, 2) = "'" & tvwStruct.Nodes(arrKey(contFor)).Text & "'"
      Else
        'sino, verifica que no esté y lo agrega
        If InStr(1, arrFiltros(Cont, 2), "'" & tvwStruct.Nodes(arrKey(contFor)).Text & "'", 1) = 0 Then
          arrFiltros(Cont, 2) = arrFiltros(Cont, 2) & "," & "'" & tvwStruct.Nodes(arrKey(contFor)).Text & "'"
        End If
      End If
    
    End If
  Next
  
  'si exiten campos filtrados genera el string para el ini
  For Cont = 1 To UBound(arrFiltros)
    'armo el filtro con el IN
    If arrFiltros(Cont, 2) <> "" Then
      arrFiltros(Cont, 3) = " IN (" & arrFiltros(Cont, 2) & ")"
    Else
      arrFiltros(Cont, 3) = ""
    End If
  Next

End Sub

Private Sub GuardarIni_Check(ID As Variant, Estado As Boolean)
' cuando se hace un check o se saca, guarda o saca los id,

Dim Cont As Integer

  If Estado Then
  
    'si puso check y el strIdTVW esta vacío agrega el id
    If strIdTVW = "" Then
      strIdTVW = ID
        
     'agrego 1 a nodos seleccionados
     SelectNodos = SelectNodos + 1
    
    Else
      
      'si no existe lo agrega
      If InStr(1, strIdTVW, ID, 1) = 0 Then
         
         strIdTVW = strIdTVW & ";" & ID
        
        'agrego 1 a nodos seleccionados
        SelectNodos = SelectNodos + 1
      
      End If
    
    End If
 
  Else
    'si le estoy sacando el check
    'saco 1 a nodos seleccionados
    SelectNodos = SelectNodos - 1
  
    'si se saco un check, borro del strIdTVW el id del nodo
    strIdTVW = Replace(strIdTVW, ID, "")
    'si queda dos ; juntas las remplaza por una sola
    strIdTVW = Replace(strIdTVW, ";;", ";")
    
    'si quedo ; a la derecha la saca
    If Right(strIdTVW, 1) = ";" Then
      strIdTVW = Left(strIdTVW, Len(strIdTVW) - 1)
    Else
      
      'si quedo ; a la izquierda la saca
      If Left(strIdTVW, 1) = ";" Then
        strIdTVW = Right(strIdTVW, Len(strIdTVW) - 1)
      End If
    
    End If
 
  End If
  
End Sub

Private Sub BuscoHijos_NodeCheck(ByVal Node As MSComctlLib.Node, Estado As Boolean)
'cuando se hace o se saca un check le pone o saca el check a los hijos

  Dim intCont As Integer


  'si el nodo donde se hizo check tiene hijos, recorro los hijos
  If Node.Children > 0 Then
  
    'guarda en el indice el primer nodo hijo con la propiedad FirstSibling
    intCont = Node.Child.FirstSibling.Index
    
    'mientras intCont no sea el ultimo hijo (LastSibling) pasa al nodo siguiente (Next)
    While intCont <> Node.Child.LastSibling.Index
      
      'guarda id para el ini
      Call GuardarIni_Check(tvwStruct.Nodes(intCont).Key, Estado)
      
      'pone el check
      tvwStruct.Nodes(intCont).Checked = Estado
      
      'si el nodo no pertenece al último nivel, busco si tiene hijos
      If arrNodo(1, tvwStruct.Nodes(intCont).Index) < contNivel Then
      
        'vuelvo a llamar al procedimiento para hacer lo mismo con los hijos
        Call BuscoHijos_NodeCheck(tvwStruct.Nodes(intCont), Estado)
      
      End If
      
      'guarda en intContn el indice del siguiente nodo.
      intCont = tvwStruct.Nodes(intCont).Next.Index
      
    Wend
    
    'pone el check en el último hijo
    Node.Child.LastSibling.Checked = Estado

    'guarda id para el ini
    Call GuardarIni_Check(Node.Child.LastSibling.Key, Estado)
    
    'si el nodo no pertenece al último nivel, busco si tiene hijos
    If arrNodo(1, tvwStruct.Nodes(intCont).Index) < contNivel Then

      'vuelvo a llamar al procedimiento para hacer lo mismo con los hijos
      Call BuscoHijos_NodeCheck(tvwStruct.Nodes(intCont), Estado)
      
    End If
    
  End If

End Sub

Private Sub tvwStruct_NodeCheck(ByVal Node As MSComctlLib.Node)
'cuando hago un check en el nodo
 
  Dim Cont As Integer
  
  'cambio puntero mouse
  Me.tvwStruct.MousePointer = vbHourglass
  
  'si le estoy poniendo un check
  If Node.Checked Then

    'llamo al procedimiento que guarda el id para el ini
    Call GuardarIni_Check(Node.Key, True)
    
    'si el nodo no pertenece al último nivel, busco si tiene hijos
    If arrNodo(1, Node.Index) < contNivel Then
      
      'llamo al procedimiento para hacer lo mismo con los hijos
      Call BuscoHijos_NodeCheck(Node, True)
     
    End If
    
  Else
    
    'guarda id para el ini
    Call GuardarIni_Check(Node.Key, False)
        
    'si el nodo no pertenece al último nivel, busco si tiene hijos
    If arrNodo(1, Node.Index) < contNivel Then
      
      'llamo al procedimiento para hacer lo mismo con los hijos
      Call BuscoHijos_NodeCheck(Node, False)
      
    End If
    
  End If
  

  Form_Activate
  'muestro info en barra de estado para que actualice cuando
  barraTVW.Panels(1).Text = " " & Node.Text
  barraTVW.Panels(1).AutoSize = sbrContents
  barraTVW.Panels(3).Text = "Seleccionados: " & SelectNodos & " "
  barraTVW.Panels(3).AutoSize = sbrContents
  
  'recupero puntero mouse
  Me.tvwStruct.MousePointer = vbDefault

End Sub

Private Sub tvwStruct_NodeClick(ByVal Node As MSComctlLib.Node)
    
  'control de errores
  On Error GoTo controlNodo
  
  Dim rs As ADODB.Recordset
  Dim nodo As Node
  Dim arrHijo As Variant
      
  'separo select de nombre utilizado para condicion
  arrHijo = separateText(Node.Tag)
  
  'get Select
  strSQL = arrHijo(1)
  
  'get datos
  Set rs = SQLexec(strSQL)
  
  'check error
  
  'cambio puntero mouse
  Me.tvwStruct.MousePointer = vbHourglass
  
  'recorro rs
  Do While Not rs.EOF
      
    'agrego nodo
    Set nodo = tvwStruct.Nodes.Add(Node.Key, tvwChild, Node.Text & LTrim(str(rs!ID)), rs!nombre, ImageList2.ListImages(1).Index)
    nodo.Tag = rs!hijo & ";" & rs!condicion
              
    Debug.Print nodo.Tag
              
    'puntero siguiente
    rs.MoveNext
    
  Loop
  
  'expando padre
  Node.Expanded = True
    
  'recupero puntero mouse
  Me.tvwStruct.MousePointer = vbDefault
  
  'cierro
  SQLclose
  
'cuando se hace click en el boton derecho del mouse
  
'  Dim Cont As Integer
'  Dim arrMnuTag As Variant
  
'  'en la barra de estado aparece el nombre del nodo al que le hice click
'  barraTVW.Panels(1).Text = " " & Node.Text
  
'  'si no es boton derecho, salgo
'  If intBotonMouse <> 2 Then Exit Sub
    
'    'recorro la cantidad de niveles
'    For Cont = 0 To contMenu
       
'      'si el tag del menu es distinto de nada
'      If arrMenu(mnuAddNivel.Item(Cont).Index, 1) <> "" Then
        
'        'si el menu es igual al siguiente nivel lo habilita
'        If arrMenu(mnuAddNivel.Item(Cont).Index, 1) = arrNodo(1, Node.Index) + 1 Or arrMenu(mnuAddNivel.Item(Cont).Index, 1) = contNivel And arrNodo(1, Node.Index) <> contNivel Then
       
'          mnuAddNivel.Item(Cont).Enabled = True
        
'        Else
'          mnuAddNivel.Item(Cont).Enabled = False
'        End If
      
'      End If
    
'    Next

'    'habilito modificar y eliminar segun menu
'    ModEli.Item(1).Caption = "Modificar " & arrNodo(3, Node.Index) & " " & Node.Text
'    ModEli.Item(1).Enabled = True
'    ModEli.Item(2).Caption = "Eliminar " & arrNodo(3, Node.Index) & " " & Node.Text
'    ModEli.Item(2).Enabled = True

'    'levanta menu para agregar niveles
'    PopupMenu mnuNiveles, vbPopupMenuLeftAlign
    
'Control Errores
controlNodo:
  
  'recupero puntero mouse
  Me.tvwStruct.MousePointer = vbDefault
   
End Sub
