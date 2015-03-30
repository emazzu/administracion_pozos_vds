VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form costosFRM 
   Caption         =   "Costs"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   11535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Command1"
      Height          =   420
      Left            =   10260
      TabIndex        =   12
      Top             =   945
      Width           =   825
   End
   Begin VB.Frame Frame2 
      Height          =   4335
      Left            =   45
      TabIndex        =   11
      Top             =   3195
      Width           =   11400
      Begin VB.CommandButton cmd_guardar_det 
         Height          =   330
         Left            =   10485
         Picture         =   "costosFRM.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Editar una fila en detalle"
         Top             =   135
         Width           =   420
      End
      Begin VB.CommandButton cmd_cancelar_det 
         Height          =   330
         Left            =   10935
         Picture         =   "costosFRM.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Eliminar una fila en detalle"
         Top             =   135
         Width           =   420
      End
      Begin VB.CommandButton cmd_ELI_det 
         Height          =   330
         Left            =   9855
         Picture         =   "costosFRM.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Eliminar una fila en detalle"
         Top             =   135
         Width           =   420
      End
      Begin VB.CommandButton cmd_EDI_det 
         Height          =   330
         Left            =   9405
         Picture         =   "costosFRM.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Editar una fila en detalle"
         Top             =   135
         Width           =   420
      End
      Begin VB.CommandButton cmd_INS_det 
         Height          =   330
         Left            =   8955
         Picture         =   "costosFRM.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Insertar una fila en detalle"
         Top             =   135
         Width           =   420
      End
      Begin FPSpreadADO.fpSpread spdDet 
         Height          =   3750
         Left            =   45
         TabIndex        =   3
         Top             =   495
         Width           =   11310
         _Version        =   393216
         _ExtentX        =   19950
         _ExtentY        =   6615
         _StockProps     =   64
         BackColorStyle  =   1
         EditEnterAction =   5
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   501
         MaxRows         =   50
         SpreadDesigner  =   "costosFRM.frx":1BB2
         Appearance      =   1
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmd_filtro_rapido"
            Object.ToolTipText     =   "Haga clic aqui para visualizar el pozo seleccionado."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmd_insertar"
            Object.ToolTipText     =   "Insertar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmd_editar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmd_eliminar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmd_guardar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmd_cancelar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10395
      Top             =   1575
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "costosFRM.frx":544E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "costosFRM.frx":57E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "costosFRM.frx":5D82
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "costosFRM.frx":631C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "costosFRM.frx":68B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "costosFRM.frx":6E50
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   45
      TabIndex        =   0
      Top             =   315
      Width           =   10005
      Begin VB.ComboBox cboPozos 
         Height          =   315
         ItemData        =   "costosFRM.frx":73EA
         Left            =   135
         List            =   "costosFRM.frx":73EC
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   405
         Width           =   3075
      End
      Begin FPSpreadADO.fpSpread spdCab 
         Height          =   1995
         Left            =   135
         TabIndex        =   2
         Top             =   765
         Width           =   9735
         _Version        =   393216
         _ExtentX        =   17171
         _ExtentY        =   3519
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   6
         MaxRows         =   8
         ScrollBars      =   0
         SpreadDesigner  =   "costosFRM.frx":73EE
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Seleccione Well ID y haga clic en filtrar"
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   180
         Width           =   3255
      End
   End
End
Attribute VB_Name = "costosFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_Operacion As String
Public m_strCboEmpresas, m_strCboTrabajos As String
Private m_arrTD() As String

Public Property Let dsiOperacion(ByVal str As String)
  m_Operacion = str
End Property

Public Property Get dsiOperacion() As String
  dsiOperacion = m_Operacion
End Property

Public Property Let dsiCboEmpresas(ByVal str As String)
  m_strCboEmpresas = str
End Property

Public Property Get dsiCboEmpresas() As String
  dsiCboEmpresas = m_strCboEmpresas
End Property

Public Property Let dsiCboTrabajos(ByVal str As String)
  m_strCboTrabajos = str
End Property

Public Property Get dsiCboTrabajos() As String
  dsiCboTrabajos = m_strCboTrabajos
End Property

Private Sub Command1_Click()
  
  Me.spdDet.MaxRows = Me.spdDet.DataRowCnt + 1
  Me.spdDet.InsertRows Me.spdDet.DataRowCnt + 1, 1
  Me.spdDet.SetText 3, Me.spdDet.MaxRows, "dato"
    
End Sub


Private Sub cmd_cancelar_det_Click()
  
  'LOCK y paint fila nueva
  spdDet.col = 1
  spdDet.Col2 = spdDet.MaxCols
  spdDet.row = Me.spdDet.ActiveRow
  spdDet.Row2 = Me.spdDet.ActiveRow
  spdDet.BlockMode = True
  spdDet.Lock = True
  spdDet.BackColor = RGB(255, 255, 255)
  spdDet.BlockMode = False
    
  'CONFIG empresa como text
  Me.spdDet.col = conDETempresa
  Me.spdDet.CellType = CellTypeEdit
  
  'CONFIG trabajo como text
  Me.spdDet.col = conDETtrabajo
  Me.spdDet.CellType = CellTypeEdit
  
  'RECOVERY info antes de modificarla
  Call spdDet_fila_recuperar(Me.spdDet.ActiveRow)
  
End Sub

Private Sub cmd_EDI_det_Click()

  Dim strID, StrEmpresa, StrTrabajo As Variant
  
  'GET id detalle
'  Me.spdCab.GetText 2, 1, strID
      
  'CHECK si no exoste cabecera, exit
'  If strID = "" Then
'    MsgBox ("no se puede, primero debe guardar una cabecera.")
'    Exit Sub
'  End If
  
  'UNLOCK y paint fila nueva
  spdDet.col = 1
  spdDet.Col2 = spdDet.MaxCols
  spdDet.row = Me.spdDet.ActiveRow
  spdDet.Row2 = Me.spdDet.ActiveRow
  spdDet.BlockMode = True
  spdDet.Lock = False
  spdDet.BackColor = RGB(225, 247, 227)
  spdDet.BlockMode = False
    
  'SAVE info fila antes de modificarla, se usa en caso que el usuario decida cancelar
  Call spdDet_fila_guardar(Me.spdDet.ActiveRow)
      
  'SAVE texto empresa y trabajo
  Me.spdDet.GetText conDETempresa, Me.spdDet.ActiveRow, StrEmpresa
  Me.spdDet.GetText conDETtrabajo, Me.spdDet.ActiveRow, StrTrabajo
      
  'CONFIG y FILL empresas cboBox
  Me.spdDet.col = conDETempresa
  Me.spdDet.CellType = CellTypeComboBox
  Me.spdDet.TypeComboBoxList = Me.dsiCboEmpresas
  
  'CONFIG y FILL trabajos cboBox
  Me.spdDet.col = conDETtrabajo
  Me.spdDet.CellType = CellTypeComboBox
  Me.spdDet.TypeComboBoxList = Me.dsiCboTrabajos
  
  'PUT texto de empresa y trabajo
  Me.spdDet.SetText conDETempresa, Me.spdDet.ActiveRow, StrEmpresa
  Me.spdDet.SetText conDETtrabajo, Me.spdDet.ActiveRow, StrTrabajo
  
End Sub

Private Sub cmd_guardar_det_Click()

  'LOCK y paint fila nueva
  spdDet.col = 1
  spdDet.Col2 = spdDet.MaxCols
  spdDet.row = Me.spdDet.ActiveRow
  spdDet.Row2 = Me.spdDet.ActiveRow
  spdDet.BlockMode = True
  spdDet.Lock = True
  spdDet.BackColor = RGB(255, 255, 255)
  spdDet.BlockMode = False

End Sub

Private Sub cmd_INS_det_Click()
      
  Dim strID As Variant
  
  'GET id cabecera
  Me.spdCab.GetText 2, 1, strID
      
  'CHECK si no existe cabecera, exit
  If strID = "" Then
    MsgBox ("no se puede, primero debe guardar una cabecera.")
    Exit Sub
  End If
  
  'UNLOCK y paint fila nueva
  spdDet.col = 1
  spdDet.Col2 = spdDet.MaxCols
  spdDet.row = Me.spdDet.DataRowCnt + 1
  spdDet.Row2 = Me.spdDet.DataRowCnt + 1
  spdDet.BlockMode = True
  spdDet.Lock = False
  spdDet.BackColor = RGB(225, 247, 227)
  spdDet.BlockMode = False
    
  'GET info fila antes de modificarla, se usa en caso que el usuario decida cancelar
  Call spdDet_fila_guardar(Me.spdDet.DataRowCnt + 1)
    
  'CONFIG y FILL empresa comboBox
  Me.spdDet.col = conDETempresa
  Me.spdDet.CellType = CellTypeComboBox
  Me.spdDet.TypeComboBoxMaxDrop = 10
  Me.spdDet.TypeComboBoxList = Me.dsiCboEmpresas
      
  'CONFIGL trabajo comboBox
  Me.spdDet.col = conDETtrabajo
  Me.spdDet.CellType = CellTypeComboBox
  Me.spdDet.TypeComboBoxMaxDrop = 10
        
  'CONFIG formula para M
  Me.spdDet.col = conM
  Me.spdDet.Formula = "D#-E#"
      
End Sub

Private Sub cmdCalc_Click()

  'CALC costo - resuelvo formula
  Call formula_Calcula

End Sub


Private Sub Form_Load()
    
  Dim rs As ADODB.Recordset
  Dim strT, strL As String
  Dim intI As Integer
  Dim blnB As Boolean
  
  'check configuracion regional
  If Not checkConfigRegional() Then
    blnB = MsgBox("El sistema detecto que la configuración regional no es correcta." & vbCrLf & vbCrLf & _
           "Configurar el formato para números de esta forma: 123,456,789.00." & vbCrLf & vbCrLf & _
           "Configurar el formato para fechas  de esta forma: dd/MM/yyyy.", vbCritical + vbOKOnly, "Atención...")
    End
  End If
  
  'get parametros de conexion
  blnB = SQLgetParam()
  
  'check si parametros ok, get menu
  If blnB Then
    
  End If
    
    
    
  'FILL combo Box
  Call lleno_Pozos_conCosto
  Call lleno_Pozos
  Call lleno_Logging
  Call lleno_MudLogging
  Call lleno_WellSiteGeology
      
  'FILL empresas
  Call lleno_Empresas
  
'  'FILL trabajos
'  Call lleno_Trabajos
      
  'CLOSE conection
  SQLclose
  
  'LOCK cab y det
  Call spd_Bloquear(Me.spdCab, True)
  Call spd_Bloquear(Me.spdDet, True)
  
  'SET para que funcione la formula
  Me.spdDet.AllowUserFormulas = True
  Me.spdDet.AutoCalc = True
  Me.spdDet.FormulaSync = True
    
End Sub

Public Sub lleno_Pozos_conCosto()
  
  Dim strT, strComboTxt, strComboID As String
  Dim blnB As Boolean
  Dim rs As ADODB.Recordset
  Dim lngCantidadPozos, lngElemento As Long
  
  'BUILD vista
  strT = "select * from maeCostos_LISTA_pozos_conCosto_vw"
        
  'GET vista
  Set rs = SQLexec(strT)
          
  'CHECK errores
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    Exit Sub
  End If
                
  'CLEAR
  lngCantidadPozos = 0
  
  'CHECK si hay registros, reservo para guardar el TD
  If Not rs.EOF Then
    lngCantidadPozos = rs.RecordCount
  End If
    
  ReDim m_arrTD(lngCantidadPozos)
  
  'CLEAR
  Me.cboPozos.Clear
  lngElemento = 0
  
  'WHILE rs
  While Not rs.EOF
    
    'ADD a comboBox para busqueda
    Me.cboPozos.AddItem rs!nombre
    Me.cboPozos.ItemData(Me.cboPozos.NewIndex) = rs!ID
    m_arrTD(lngElemento) = rs!TD
    
    'ADD 1 a contador de elementos
    lngElemento = lngElemento + 1
    
    'NEXT rs
    rs.MoveNext
            
  Wend

End Sub


Public Sub lleno_Pozos()
  
  Dim strT, strComboTxt, strComboID As String
  Dim lngFila As Long
  Dim blnB As Boolean
  Dim rs As ADODB.Recordset
        
        
  'BUILD vista
  strT = "select * from maeCostos_LISTA_Pozos_vw"
        
  'GET vista
  Set rs = SQLexec(strT)
          
  'CHECK errores
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    Exit Sub
  End If
            
  'REDIM
  ReDim m_arrTD(rs.RecordCount)
  
  lngFila = 0
  
  'WHILE rs
  While Not rs.EOF
    
    'ADD texto
    strComboText = strComboText & rs!nombre & Chr(9)
    
    'ADD TD
    m_arrTD(lngFila) = IIf(IsNull(rs!TotalDepth), 0, rs!TotalDepth)
    
    'NEXT rs
    rs.MoveNext
            
    'ADD contador de filas
    lngFila = lngFila + 1
            
  Wend

  'FILL comboBox
  Me.spdCab.row = conCABpozo
  Me.spdCab.col = 2
  Me.spdCab.TypeComboBoxList = strComboText
                          
End Sub

Public Sub lleno_Logging()
  
  Dim strT, strComboTxt, strComboID As String
  Dim blnB As Boolean
  Dim rs As ADODB.Recordset
      
  'BUILD vista
  strT = "select * from maeCostos_LISTA_log_vw"
        
  'GET vista
  Set rs = SQLexec(strT)
          
  'CHECK errores
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    Exit Sub
  End If
            
  'WHILE rs
  While Not rs.EOF
    
    'ADD texto
    strComboText = strComboText & rs!nombre & Chr(9)
            
    'NEXT rs
    rs.MoveNext
            
  Wend
                          
  'FILL comboBox
  Me.spdCab.row = conCABlog
  Me.spdCab.col = 2
  Me.spdCab.TypeComboBoxList = strComboText
                          
End Sub

Public Sub lleno_MudLogging()

  Dim strT, strComboTxt, strComboID As String
  Dim blnB As Boolean
  Dim rs As ADODB.Recordset
      
  'BUILD vista
  strT = "select * from maeCostos_LISTA_mudLog_vw"
        
  'GET vista
  Set rs = SQLexec(strT)
          
  'CHECK errores
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    Exit Sub
  End If
            
  'WHILE rs
  While Not rs.EOF
    
    'ADD texto
    strComboTexto = strComboTexto & rs!nombre & Chr(9)
            
    'NEXT rs
    rs.MoveNext
            
  Wend
                          
  'FILL comboBox
  Me.spdCab.row = conCABmud
  Me.spdCab.col = 2
  Me.spdCab.TypeComboBoxList = strComboTexto

End Sub


Public Sub lleno_WellSiteGeology()

  Dim strT, strComboTxt  As String
  Dim blnB As Boolean
  Dim rs As ADODB.Recordset
      
  'BUILD vista
  strT = "select * from maeCostos_LISTA_site_vw"
        
  'GET vista
  Set rs = SQLexec(strT)
          
  'CHECK errores
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    Exit Sub
  End If
            
  'WHILE rs
  While Not rs.EOF
    
    'ADD texto
    strComboTxt = strComboTxt & rs!nombre & Chr(9)
            
    'NEXT rs
    rs.MoveNext
            
  Wend
                          
  'FILL comboBox
  Me.spdCab.row = conCABsite
  Me.spdCab.col = 2
  Me.spdCab.TypeComboBoxList = strComboTxt

End Sub


Public Sub lleno_Empresas()
  
  Dim strT, strComboTxt As String
  Dim blnB As Boolean
  Dim rs As ADODB.Recordset
        
  'BUILD vista
  strT = "select * from maeCostos_LISTA_empresas_vw"
        
  'GET vista
  Set rs = SQLexec(strT)
          
  'CHECK errores
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    Exit Sub
  End If
            
  'WHILE rs
  While Not rs.EOF
    
    'ADD texto
    strComboTxt = strComboTxt & rs!nombre & Chr(9)
            
    'NEXT rs
    rs.MoveNext
            
  Wend

  'SAVE comboBox
  Me.dsiCboEmpresas = strComboTxt
  
End Sub


Public Function lleno_Trabajos(ByVal str As String) As String
  
  Dim strT, strComboTxt As String
  Dim blnB As Boolean
  Dim rs As ADODB.Recordset
        
  'BUILD vista
  strT = "select * from maeCostos_LISTA_Trabajos_vw where empresa = '" & str & "'"
        
  'GET vista
  Set rs = SQLexec(strT)
          
  'CHECK errores
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    Exit Function
  End If
            
  'WHILE rs
  While Not rs.EOF
    
    'ADD texto
    strComboTxt = strComboTxt & rs!nombre & Chr(9)
            
    'NEXT rs
    rs.MoveNext
            
  Wend
  
  'RETURN
  lleno_Trabajos = strComboTxt
    
  'CLOSE
  SQLclose
  
End Function


'
'LOCK grilla
'
Public Sub spd_Bloquear(spdG As fpSpread, blnBloquear As Boolean)
  
  spdG.col = 1
  spdG.Col2 = spdG.MaxCols
  spdG.row = 1
  spdG.Row2 = spdG.MaxRows
  spdG.BlockMode = True
  spdG.Lock = blnBloquear
  spdG.BlockMode = False
      
  'LOCK celdas de cabeza siempre
  If spdG.Name = "spdCab" Then
    
    spdG.col = 2
    spdG.row = 1
    spdG.Lock = True
    
    spdG.row = 4
    spdG.Lock = True
    
    spdG.row = 5
    spdG.Lock = True
    
  End If
      
End Sub


'
'CLEAR grilla cab
'
Public Sub spdCab_Blanquear(spdG As fpSpread)
  
  spdG.ClearRange 2, 1, 2, spdG.MaxRows, True
  spdG.ClearRange 5, 1, 5, spdG.MaxRows, True
  
End Sub

'
'SAVE cab
'
Public Sub spdCab_Guardar(spdG As fpSpread)
      
  Dim sngDfwa, sngTopeZone, sngPerforation, sngBasicCharge, sngDiscAdd, sngDiscLog, sngQuote As Variant
  Dim intID, dtmFecha, strCurrency, strT, strPozo, strEleLog, strMudLog, strWelsite As Variant
  Dim rs As ADODB.Recordset
    
  'GET datos e ID en caso de LISTA
  Me.spdCab.GetText 2, 1, intID
  Me.spdCab.GetText 2, 2, dtmFecha
  Me.spdCab.GetText 2, 3, strPozo
  Me.spdCab.GetText 2, 6, strEleLog
  Me.spdCab.GetText 2, 7, strMudLog
  Me.spdCab.GetText 2, 8, strWelsite
  Me.spdCab.GetText 5, 1, sngDfwa
  Me.spdCab.GetText 5, 2, sngTopeZone
  Me.spdCab.GetText 5, 3, sngPerforation
  Me.spdCab.GetText 5, 4, sngBasicCharge
  Me.spdCab.GetText 5, 5, sngDiscAdd
  Me.spdCab.GetText 5, 6, sngDiscLog
  Me.spdCab.GetText 5, 7, intCurrency
  Me.spdCab.GetText 5, 8, sngQuote
  
  'CHECK operacion
  Select Case Me.dsiOperacion
  
  Case "INS"
    strT = "costos_INS_sp "
    
  Case "EDI"
    strT = "costos_EDI_sp "
  
  Case "ELI"
    strT = "costos_ELI_sp "
    
  End Select
  
  'BUILD query
  strT = "exec " & strT & Val(intID) & ",'" & dateToIso(dtmFecha) & "','" & strPozo & "','" & _
         strEleLog & "','" & strMudLog & "','" & strWelsite & "'," & sngDfwa & "," & sngTopeZone & "," & _
         sngPerforation & "," & sngBasicCharge & "," & sngDiscAdd & "," & sngDiscLog & ",'" & _
         intCurrency & "'," & sngQuote
  
  'SAVE
  SQLexec (strT)
        
  'CHECK error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    Exit Sub
  End If
    
  'CHECK operacion
  Select Case Me.dsiOperacion
  
  
  Case "INS"
    
    'GET record completo despues de insertar
    strT = "select * from costos_vw where id = " & "(select max(id) from costos_vw)"
    Set rs = SQLexec(strT)
    
    'CHECK error
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
      Exit Sub
    End If
        
    'COPY datos de rs a grilla
    Call spdCab_copiar_RS_a_spd(rs)
            
        
  Case "EDI"
    
    'GET record completo despues de editar
    strT = "select * from costos_vw where id = " & intID
    Set rs = SQLexec(strT)
    
    'CHECK error
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
      Exit Sub
    End If
    
    'COPY datos de rs a grilla
    Call spdCab_copiar_RS_a_spd(rs)
    
       
  Case "ELI"
                
    'CLEAR cab
    Call spdCab_Blanquear(Me.spdCab)
            
        
  End Select
      
  'LOCK cab
  Call spd_Bloquear(Me.spdCab, True)
      
  'FILL combo Box
  Call lleno_Pozos_conCosto
  
  'SET pozo actual
  'pendiente
  
  
  'CLOSE
  SQLclose
    
End Sub


'
'COPY datos de recordset a grilla cab
'
Public Sub spdCab_copiar_RS_a_spd(rs As ADODB.Recordset)
  
  'COPY datos
  spdCab.SetText 2, 1, rs!ID
  spdCab.SetText 2, 2, rs!Date
  spdCab.SetText 2, 3, rs![well ID]
  spdCab.SetText 2, 4, rs!uwi
  spdCab.SetText 2, 5, rs!Area
  spdCab.SetText 2, 6, rs![Electric logging]
  spdCab.SetText 2, 7, rs![Mud logging company]
  spdCab.SetText 2, 8, rs![Well site geology company]
  
  spdCab.SetText 5, 1, rs!Dfwa
  spdCab.SetText 5, 2, rs![Tope Zone]
  spdCab.SetText 5, 3, rs!Perforations
  spdCab.SetText 5, 4, rs![Basic charge cost]
  spdCab.SetText 5, 5, rs![Discount additional]
  spdCab.SetText 5, 6, rs![Discount logging]
  spdCab.SetText 5, 7, rs!Currency
  spdCab.SetText 5, 8, rs!Quote
  
End Sub


'
'GET costo
'
Public Sub costo_get()
    
  Dim strT As String
  Dim rsCab As ADODB.Recordset
    
  'BUILD query
  strT = "select * from costos_vw where IDpozo = " & Me.cboPozos.ItemData(Me.cboPozos.ListIndex)
  
  'GET rs
  Set rsCab = SQLexec(strT)
        
  'CHECK error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    Exit Sub
  End If
    
'  'BUILD query
'  strT = "select * from costos_vw where IDpozo = '" & Me.cboPozos & "'"
'
'  'GET rs
'  Set rsCab = SQLexec(strT)
'
'  'CHECK error
'  If Not SQLparam.CnErrNumero = -1 Then
'    SQLError
'    SQLclose
'    Exit Sub
'  End If
        
  'COPY datos de rs a grilla cab
  Call spdCab_copiar_RS_a_spd(rsCab)
        
'  'COPY datos de rs a grilla det
'  Call spdDet_copiar_RS_a_spd(rsDet)
        
  'CLOSE
  SQLclose
  
End Sub



Private Sub spdDet_ComboSelChange(ByVal col As Long, ByVal row As Long)
  
  Dim StrEmpresa As Variant
  Dim strTrabajos As Variant
    
  'CHECK cambio columna empresa
  If col = 1 Then
    
    'GET empresa
    Me.spdDet.GetText conDETempresa, Me.spdDet.ActiveRow, StrEmpresa
    
    'GET trabajos según empresa
    strTrabajos = lleno_Trabajos(StrEmpresa)
    
    'FILL comboBox
    Me.spdDet.row = Me.spdDet.ActiveRow
    Me.spdDet.col = conDETtrabajo
    Me.spdDet.TypeComboBoxList = strTrabajos
    
  End If
  
  'CHECK cambio columna trabajo
  If col = 2 Then
    
    'GET empresa y trabajo
    Me.spdDet.GetText conDETempresa, Me.spdDet.ActiveRow, StrEmpresa
    Me.spdDet.GetText conDETtrabajo, Me.spdDet.ActiveRow, StrTrabajo
    
    'GET formula y valores segun empresa y trabajo
    Call formula_Get(StrEmpresa, StrTrabajo)
    
    'FILL grilla
    Me.spdDet.SetText conTipo, Me.spdDet.ActiveRow, typeFormula.strTipo
    Me.spdDet.SetText conXdia, Me.spdDet.ActiveRow, typeFormula.sngXdia
    Me.spdDet.SetText conXpozo, Me.spdDet.ActiveRow, typeFormula.sngXpozo
    Me.spdDet.SetText conXprof, Me.spdDet.ActiveRow, typeFormula.sngXprofundidad
    Me.spdDet.SetText conXregistroMen, Me.spdDet.ActiveRow, typeFormula.sngXregistroMen
    Me.spdDet.SetText conXregistroMas, Me.spdDet.ActiveRow, typeFormula.sngXtegistroMas
    Me.spdDet.SetText conDesc, Me.spdDet.ActiveRow, typeFormula.sngXdescPje
    Me.spdDet.SetText conFormula, Me.spdDet.ActiveRow, typeFormula.strFormula
        
    'FILL td
    Me.spdDet.SetText conFrom, Me.spdDet.ActiveRow, m_arrTD(Me.cboPozos.ListIndex)
    
    
  End If
    
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  
  'CHECK opción
  Select Case Button.Key
  
  
  'filtro rapido
  Case "cmd_filtro_rapido"
      
    'GET costo
    Call costo_get
    
    
  'insertar
  Case "cmd_insertar"
    
    'UNLOCK cab
    Call spd_Bloquear(Me.spdCab, False)
    
    'CLEAR cab
    Call spdCab_Blanquear(Me.spdCab)
      
    'ASSIGN operacion
    Me.dsiOperacion = "INS"
    
    
  'editar
  Case "cmd_editar"
      
    'LOCK cab y det
    Call spd_Bloquear(Me.spdCab, False)
    
    'ASSIGN operacion
    Me.dsiOperacion = "EDI"
    
    
  'eliminar
  Case "cmd_eliminar"
        
    'ASSIGN operacion
    Me.dsiOperacion = "ELI"
    
    
  'guardar
  Case "cmd_guardar"
    
    'CHECK si operacion ELI
    If Me.dsiOperacion = "ELI" Then
      
      intI = MsgBox("Esta seguro que desea eliminar los costos para el pozo " & Me.cboPozos & " ?", vbQuestion + vbYesNo, "Atención...")
        
      'CHECK si cancela
      If intI = 7 Then
        Exit Sub
      End If
      
    End If
    
    'SAVE cab
    Call spdCab_Guardar(Me.spdCab)
    
    
  End Select
  

End Sub

'
'SAVE informacion fila grilla detalle
'
Sub spdDet_fila_guardar(ByVal lngFila As Long)
  
  Me.spdDet.GetText conDETempresa, lngFila, typeCostoDet.StrEmpresa
  Me.spdDet.GetText conDETtrabajo, lngFila, typeCostoDet.StrTrabajo
  Me.spdDet.GetText conDays, lngFila, typeCostoDet.intDays
  Me.spdDet.GetText conFrom, lngFila, typeCostoDet.intFrom
  Me.spdDet.GetText conTo, lngFila, typeCostoDet.intTo
  Me.spdDet.GetText conM, lngFila, typeCostoDet.intM
  Me.spdDet.GetText conVSPS, lngFila, typeCostoDet.intVSPS
  Me.spdDet.GetText conCheck, lngFila, typeCostoDet.intCheck
  Me.spdDet.GetText conPressure, lngFila, typeCostoDet.intPressure
  Me.spdDet.GetText conCost, lngFila, typeCostoDet.sngCost
  Me.spdDet.GetText conCostDesc, lngFila, typeCostoDet.sngCostDesc
  Me.spdDet.GetText conComment, lngFila, typeCostoDet.strComment
  
End Sub


'
'RECOVERY informacion fila grilla detalle
'
Sub spdDet_fila_recuperar(ByVal lngFila As Long)
  
  Me.spdDet.SetText conDETempresa, lngFila, typeCostoDet.StrEmpresa
  Me.spdDet.SetText conDETtrabajo, lngFila, typeCostoDet.StrTrabajo
  Me.spdDet.SetText conDays, lngFila, typeCostoDet.intDays
  Me.spdDet.SetText conFrom, lngFila, typeCostoDet.intFrom
  Me.spdDet.SetText conTo, lngFila, typeCostoDet.intTo
  Me.spdDet.SetText conM, lngFila, typeCostoDet.intM
  Me.spdDet.SetText conVSPS, lngFila, typeCostoDet.intVSPS
  Me.spdDet.SetText conCheck, lngFila, typeCostoDet.intCheck
  Me.spdDet.SetText conPressure, lngFila, typeCostoDet.intPressure
  Me.spdDet.SetText conCost, lngFila, typeCostoDet.sngCost
  Me.spdDet.SetText conCostDesc, lngFila, typeCostoDet.sngCostDesc
  Me.spdDet.SetText conComment, lngFila, typeCostoDet.strComment
  
End Sub

'
'formula Get
'
Public Sub formula_Get(ByVal StrEmpresa As String, ByVal StrTrabajo As String)
    
  Dim strT As String
  Dim rs As ADODB.Recordset
    
  'BUILD query
  strT = "select * from maeCostos_LISTA_formula_vw where empresa = '" & StrEmpresa & "' and trabajo = '" & StrTrabajo & "'"
  
  'GET rs
  Set rs = SQLexec(strT)
        
  'CHECK error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    Exit Sub
  End If
    
  'CHECK si hay registros
  If Not rs.EOF Then
    
    typeFormula.strTipo = rs!tipo
    typeFormula.dtmFecha = rs!fecha
    typeFormula.sngXdia = rs!xDia
    typeFormula.sngXpozo = rs!xPozo
    typeFormula.sngXprofundidad = rs!xProfundidad
    typeFormula.sngXregistroMen = rs!xRegistroMen
    typeFormula.sngXtegistroMas = rs!xRegistroMas
    typeFormula.sngXdescPje = rs!descPje
    typeFormula.strFormula = rs!Formula
    
  End If
  
  'CLOSE
  SQLclose
  
End Sub

'
'formula calcula
'
Public Sub formula_Calcula()
    
  Dim strFormula, strDescuento As Variant
  Dim intFila As Integer
  
  'WHILE filas
  For intFila = 1 To Me.spdDet.DataRowCnt
  
    'GET formula
    Me.spdDet.GetText conFormula, intFila, strFormula
          
    'GET valores ingresados a mano por usuario
    Me.spdDet.GetText conDays, intFila, typeCostoDet.intDays
    Me.spdDet.GetText conFrom, intFila, typeCostoDet.intFrom
    Me.spdDet.GetText conTo, intFila, typeCostoDet.intTo
    Me.spdDet.GetText conM, intFila, typeCostoDet.intM
    Me.spdDet.GetText conVSPS, intFila, typeCostoDet.intVSPS
    Me.spdDet.GetText conCheck, intFila, typeCostoDet.intCheck
    Me.spdDet.GetText conPressure, intFila, typeCostoDet.intPressure
    
    'GET costos
    Me.spdDet.GetText conXdia, intFila, typeFormula.sngXdia
    Me.spdDet.GetText conXpozo, intFila, typeFormula.sngXpozo
    Me.spdDet.GetText conXprof, intFila, typeFormula.sngXprofundidad
    Me.spdDet.GetText conXregistroMen, intFila, typeFormula.sngXregistroMen
    Me.spdDet.GetText conXregistroMas, intFila, typeFormula.sngXtegistroMas
    Me.spdDet.GetText conDesc, intFila, typeFormula.sngXdescPje
    
    'BUILD formula, reemplazo variables por valores
    strFormula = Replace(strFormula, "[DAYS]", Val(typeCostoDet.intDays))
    strFormula = Replace(strFormula, "[COST_DAYS]", typeFormula.sngXdia)
    
    strFormula = Replace(strFormula, "[TD]", typeCostoDet.intFrom)
    strFormula = Replace(strFormula, "[COST_TD]", typeFormula.sngXprofundidad)
    strFormula = Replace(strFormula, "[M]", typeCostoDet.intM)
    strFormula = Replace(strFormula, "[COST_M]", typeFormula.sngXprofundidad)
    strFormula = Replace(strFormula, "[FIXED_COST]", typeFormula.sngXpozo)
    strFormula = Replace(strFormula, "[REGISTER-]", typeFormula.sngXregistroMen)
    strFormula = Replace(strFormula, "[REGISTER+]", typeFormula.sngXtegistroMas)
    strFormula = Replace(strFormula, "[NUMBER]", Val(typeCostoDet.intDays))
    
    'aplico costo
    Me.spdDet.row = intFila
    Me.spdDet.col = conCost
    Me.spdDet.Formula = strFormula
    
    'aplico costo con descuento
    strFormula = "J#*[DISCOUNT]/100"
    strFormula = Replace(strFormula, "[DISCOUNT]", typeFormula.sngXdescPje)
            
    'aplico costo
    Me.spdDet.col = conCostDesc
    Me.spdDet.Formula = strFormula
            
  Next
  
End Sub
