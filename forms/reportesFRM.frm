VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form reportesFRM 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10185
   Icon            =   "reportesFRM.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8490
   ScaleWidth      =   10185
   WindowState     =   2  'Maximized
   Begin VB.Frame frame 
      Height          =   1365
      Left            =   5790
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      Begin VB.ListBox lstLista 
         Height          =   735
         ItemData        =   "reportesFRM.frx":038A
         Left            =   90
         List            =   "reportesFRM.frx":038C
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   150
         Width           =   1725
      End
      Begin VB.CommandButton cmdLstCancelar 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   990
         TabIndex        =   6
         Top             =   990
         Width           =   825
      End
      Begin VB.CommandButton cmdLstAceptar 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   150
         TabIndex        =   5
         Top             =   990
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      Height          =   300
      Left            =   4590
      TabIndex        =   3
      Top             =   480
      Width           =   1080
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   300
      Left            =   4590
      TabIndex        =   2
      Top             =   120
      Width           =   1080
   End
   Begin FPSpreadADO.fpSpread SPD 
      Height          =   1365
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4275
      _Version        =   393216
      _ExtentX        =   7541
      _ExtentY        =   2408
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      SpreadDesigner  =   "reportesFRM.frx":038E
      ClipboardOptions=   3
   End
   Begin CRVIEWERLibCtl.CRViewer rptVisor 
      Height          =   2430
      Left            =   90
      TabIndex        =   1
      Top             =   1590
      Width           =   7665
      lastProp        =   500
      _cx             =   13520
      _cy             =   4286
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "ReportesFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_IDmenu As String            'save IDmenu
Private m_IDopc As String             'save IDopcion
Private m_Titulo As String            'save titulo

Dim m_dsiReport() As String           'save ID, nombre, source de reporte y subreportes
Dim m_dsiReportCant As Integer        'save cantidad de reportes y subreportes

Dim m_dsiCondicion As String          'save condicion en string
Dim m_dsiCondicionArr As Variant      'save condicion separada en array

Dim m_dsiFormulas() As String         'save formulas
Dim m_dsiFormulasCant As Integer      'save cantidad de formulas

Dim m_dsiColumnasItem() As Variant    'items de la lista en los cuales se hizo click
Dim m_dsiParamConfig() As String      'save parametros
Dim m_dsiParamCant As Integer         'save cantidad de parametros

Dim m_dsiSPant As String              'save store procedure anterior
Dim m_dsiSPpos  As String             'save store procedure posterior
Dim m_dsiSPparam As String            'save parametros para store procedure

Dim m_intDataFiltroItems As Integer   'contador reporte y subreportes

Public Property Let dsiIDmenu(ID As String)
  m_IDmenu = ID
End Property

Public Property Get dsiIDmenu() As String
  dsiIDmenu = m_IDmenu
End Property

Public Property Let dsiIDopc(ID As String)
  m_IDopc = ID
End Property

Public Property Get dsiIDopc() As String
  dsiIDopc = m_IDopc
End Property

Public Property Let dsiTitulo(ID As String)
  m_Titulo = ID
  
  'put titulo form
  Me.Caption = m_Titulo
  
End Property

Public Property Get dsiTitulo() As String
  dsiTitulo = m_Titulo
End Property

Public Property Let dsiSPant(strSP As String)
  m_dsiSPant = strSP
End Property

Public Property Get dsiSPant() As String
  dsiSPant = m_dsiSPant
End Property

Public Property Let dsiSPpos(strSP As String)
  m_dsiSPpos = strSP
End Property

Public Property Get dsiSPpos() As String
  dsiSPpos = m_dsiSPpos
End Property

Public Property Get dsiSPparam() As String
  dsiSPparam = m_dsiSPparam
End Property

'
'CAMBIA APARIENCIA A GRILLA HORIZONTAL
'
Public Property Get dsiGrillaCambiarApariencia() As Boolean
  
  SPD.UnitType = UnitTypeTwips                  'trabajar en twips

  SPD.Appearance = AppearanceFlat               'apariencia 3D
  SPD.BorderStyle = BorderStyleNone             'tipo de borde: sin borde
  
  SPD.ColHeadersAutoText = DispBlank            'titulos de columnas en blanco
  SPD.ColHeadersShow = True                     'muestra encabezado columnas
  SPD.RowHeadersShow = True                     'muestra encabezado de filas
  
  SPD.CursorStyle = CursorStyleArrow            'stilo cursor
  SPD.CursorType = CursorTypeDefault            'tipo cursor
  
  SPD.AutoSize = False                           'automaticamente ajusta ancho grilla
  SPD.DAutoSizeCols = DAutoSizeColsMax          'tipo de ajuste 2: al dato mas ancho
  
  SPD.UserColAction = UserColActionDefault      'cuando hace click en header pinta columna o fila
  SPD.FontSize = 9                              'tamaño letra
'  SPD.RowHeight(0) = 250                        'altura fila de titulos
  SPD.MoveActiveOnFocus = False                 '
  SPD.Protect = False                           'exporta a excel sin proteccion
  
  SPD.BackColorStyle = BackColorStyleUnderGrid  'estilo
  SPD.GridShowHoriz = True                      'muestra grilla horizontal
  SPD.GridShowVert = True                       'muestra grilla vertical
  SPD.GridColor = RGB(200, 200, 200)            'color muy suave
  SPD.NoBorder = True                           'sin borde fin zona de datos
  
  SPD.ScrollBars = ScrollBarsBoth               'ambas barras de desplazamiento
  SPD.ScrollBarExtMode = True                   'cuando sean necesarias
  SPD.VScrollSpecial = False                     'barra especial
      
  intRes = SPD.SetOddEvenRowColor(RGB(245, 245, 245), RGB(60, 60, 60), RGB(245, 245, 245), RGB(60, 60, 60))
  SPD.SelBackColor = RGB(204, 230, 255)         'fondo del area seleccionada amarillo
  SPD.GrayAreaBackColor = RGB(245, 245, 245)
  
  SPD.VirtualMode = False                        ' ajusta rows al tamaño del recordset
  'spd.VirtualRows = 300                         ' rows a leer del virtual buffer
  'spd.VirtualScrollBuffer = True                ' scroll vertical lee de tantas rows del buffer
  
  'setea para mostrar tooltip en las celdas donde no se ve toda la info
  SPD.TextTip = TextTipFixed
  SPD.TextTipDelay = 250
  
End Property

'
'agrega nombre de reporte y dataSource asociado a cada reporte
'
Function dsiReportAdd(strID, strNB As String, strDS As String)
  
  'add 1 a contador de reportes y subreportes
  m_dsiReportCant = m_dsiReportCant + 1
  
  'redim array
  ReDim Preserve m_dsiReport(3, m_dsiReportCant)
  
  'save reporte
  m_dsiReport(1, m_dsiReportCant) = strID
  m_dsiReport(2, m_dsiReportCant) = strNB
  m_dsiReport(3, m_dsiReportCant) = strDS
  
End Function

'guarda parametros en un array
'
Public Property Get dsiParamAdd(strNombre, strTipo, strLista, blnNoWhere)
  
  'add 1 a contador de parametros
  m_dsiParamCant = m_dsiParamCant + 1
  
  'redim array
  ReDim Preserve m_dsiParamConfig(4, m_dsiParamCant)
    
  'save parametro
  m_dsiParamConfig(1, m_dsiParamCant) = strNombre
  m_dsiParamConfig(2, m_dsiParamCant) = strTipo
  m_dsiParamConfig(3, m_dsiParamCant) = IIf(IsNull(strLista), "", strLista)
  m_dsiParamConfig(4, m_dsiParamCant) = IIf(IsNull(blnNoWhere), 0, blnNoWhere)
    
End Property

'vaciar formulas
'
Public Property Get dsiFormulasLimpiar()
  
  m_dsiFormulasCant = 0
  ReDim Preserve m_dsiFormulas(2, m_dsiFormulasCant)
  
End Property

'agregar formulas
'
Public Property Get dsiFormulasAgregar(ByVal strFormula As String, ByVal strValor As String)
  
  'add 1 a contador de formulas
  m_dsiFormulasCant = m_dsiFormulasCant + 1
  
  'redim array
  ReDim Preserve m_dsiFormulas(2, m_dsiFormulasCant)
  
  'save formula
  m_dsiFormulas(1, m_dsiFormulasCant) = strFormula
  m_dsiFormulas(2, m_dsiFormulasCant) = strValor
  
End Property

'
'get formula nombre
'
Public Property Get dsiFormulasGetNombre(intID As Integer)
  
  dsiFormulasGetNombre = m_dsiFormulas(1, intID)

End Property

'
'get formula valor
'
Public Property Get dsiFormulasGetValor(intID As Integer)
  
  dsiFormulasGetValor = m_dsiFormulas(2, intID)
  
End Property

'congirura grilla a partir de los parametros
'
Public Property Get dsiGrillaConfig()
  
  Dim lngAlturaGrilla As Long
  Dim sngWidth As Single
  Dim blnB As Boolean
  Dim intFila As Integer
  
  'set limites a grilla
  Me.SPD.MaxRows = m_dsiParamCant
  Me.SPD.MaxCols = 7
  
  'locate grilla
  Me.SPD.Top = 0
  Me.SPD.Left = 0
  Me.SPD.Width = Me.ScaleWidth
  Me.SPD.Height = Me.SPD.RowHeight(0) + 200 + (Me.SPD.RowHeight(1) * Me.SPD.MaxRows)
    
  'adjust boton cancelar
  cmdCancelar.Left = Me.ScaleWidth - cmdCancelar.Width
  cmdCancelar.Top = Me.SPD.Height
    
  'adjust boton aceptar
  cmdAceptar.Left = Me.ScaleWidth - cmdCancelar.Width - cmdAceptar.Width - 300
  cmdAceptar.Top = Me.SPD.Height
  
  'locate crystal
  rptVisor.Top = Me.SPD.Height + 200
  rptVisor.Left = 0
  rptVisor.Width = Me.ScaleWidth
  rptVisor.Height = Me.ScaleHeight - Me.SPD.Height - 200
      
  'put titulos de columna a grilla
  Me.SPD.SetText 1, 0, "Parametro"
  Me.SPD.SetText 2, 0, "Operador"
  Me.SPD.SetText 3, 0, "Criterio"
  Me.SPD.SetText 4, 0, ""
    
  'set altura fila con titulos
  Me.SPD.RowHeight(0) = 250
  
  'set columnas ocultas
  Me.SPD.ColWidth(5) = 0                'tipo de dato
  Me.SPD.ColWidth(6) = 0                'lista con select
  Me.SPD.ColWidth(7) = 0                'no lo incluye en la condicion
      
  'inicializo
  intFila = 0
    
  'recorro parametros
  For intFila = 1 To m_dsiParamCant
      
    'set puntero fila
    Me.SPD.Row = intFila
      
    'put nombre parametro
    Me.SPD.Col = 1
    Me.SPD.CellType = CellTypeStaticText
    Me.SPD.Text = m_dsiParamConfig(1, intFila)
      
    'set puntero columna operadores
    Me.SPD.Col = 2
    Me.SPD.CellType = CellTypeComboBox
    Me.SPD.TypeHAlign = TypeHAlignCenter
    Me.SPD.Lock = True
              
    'put operadores segun tipo de dato
    Select Case LCase(m_dsiParamConfig(2, intFila))
            
    Case "fecha"
      Me.SPD.TypeComboBoxList = "[]"
            
    Case "texto", "numero", "sino"
      Me.SPD.TypeComboBoxList = "="
      
    'Case "period"
    '    Me.SPD.CellType = CellTypeDate
    '    Me.SPD.TypeDateFormat = TypeDateFormatDDMMYY
          
    End Select
          
    'default operador en posicion 1
    Me.SPD.TypeComboBoxCurSel = 0
      
      
    'set columna criterio
    Me.SPD.Col = 3
    Me.SPD.CellType = CellTypeEdit
    Me.SPD.TypeMaxEditLen = 5000
      
    'set columna boton para que cuando se hace click levanta lista con datos
    Me.SPD.Col = 4
    Me.SPD.CellType = CellTypeButton
    Me.SPD.TypeButtonText = "..."
        
    'set columna tipo de parametro: fecha, numero, texto, sino
    Me.SPD.SetText 5, intFila, m_dsiParamConfig(2, intFila)
          
    'set columna LISTA
    Me.SPD.Col = 6
    Me.SPD.TypeMaxEditLen = 500
    Me.SPD.SetText 6, intFila, m_dsiParamConfig(3, intFila)
          
    'set columna para que el parametro no lo incluya en el where
    Me.SPD.SetText 7, intFila, m_dsiParamConfig(4, intFila)
      
  Next
    
  'set celda activa
  Me.SPD.SetActiveCell 3, 1
  
End Property

Private Sub cmdAceptar_Click()
   
  'build condicion
  blnB = Me.dsiCondicionArmar()
   
  'check si tiene que procesar store antes
  If Me.dsiSPant <> "" Then
    
    'exec store
    SQLexec ("exec " & Me.dsiSPant & " " & Me.dsiSPparam)
    
    'check errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
    End If
      
  End If
   
  'preview report
  blnB = Me.dsiRefresh()
   
End Sub


Private Sub cmdCancelar_Click()
      
  Dim intI As Integer
      
  'check si tiene que procesar store posterior
  If Me.dsiSPpos <> "" Then
      
    'message
    intI = MsgBox("El informe tiene una tarea pendiente, desea procesarla ??" & vbCrLf & vbCrLf & _
           "Haga click en Si, solo si el informe se imprimio correctamente.", vbQuestion + vbYesNo, "Atención")
      
    'check si confirmo store procedure
    If intI = 6 Then
      
      'exec store
      SQLexec ("exec " & Me.dsiSPpos & " " & Me.dsiSPparam)
          
      'check errores
      If Not SQLparam.CnErrNumero = -1 Then
        SQLError
        Exit Sub
      End If
      
    End If
      
  End If
      
  'close cn
  SQLclose
  
  'close form
  Unload Me
  
End Sub



Private Sub cmdlstAceptar_Click()
  
  Dim intI As Integer
  Dim strT As String
  
  'inicializo
  strT = ""
  
  'recorro item seleccionados
  For intI = 1 To Me.lstLista.SelCount
    
    'si item <> -1 agrego al string
    If Me.dsiColumnasItem(intI) <> -1 Then
      strT = strT & Me.lstLista.List(Me.dsiColumnasItem(intI)) & ","
    End If
    
  Next
    
  'delete ultima coma
  If strT <> "" Then
    strT = Left(strT, Len(strT) - 1)
  End If
    
  'put criterio
  Me.SPD.SetText 3, Me.SPD.ActiveRow, strT
    
  'set visible lista
  Me.frame.Visible = False
  
End Sub

Private Sub cmdlstCancelar_Click()
  
  'set visible lista
  Me.frame.Visible = False
  
End Sub


Private Sub Form_Load()
  
  Dim rs As ADODB.Recordset
  Dim strT As String
  Dim blnB, blnFlag As Boolean
  Dim intI As Integer
    
  'get config
  strSQL = "select o.id, o.nombre, c.tipo, c.columna, c.valor " & _
           "from dsiOPC o, dsiOPCconfig c " & _
           "where o.id = c.IDopc and " & _
           "o.IDmenu = " & SQLparam.IDmenu & " And IDopc = " & Me.dsiIDopc & " " & _
           "order by orden"
  
  'exec
  Set rs = SQLexec(strSQL)
      
  'chequeo errores
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    End
  End If
        
  'inicializo
  blnFlag = 0
        
  'while config
  While Not rs.EOF
    
    'check tipo de configuracion
    Select Case LCase(rs!tipo)
    
    'check si store antes
    Case Is = "spant"
      
      Me.dsiSPant = IIf(IsNull(rs!valor), "", rs!valor)
      
    'check si store despues
    Case Is = "sppos"
      
      Me.dsiSPpos = IIf(IsNull(rs!valor), "", rs!valor)
      
    'default report y subreport
    Case Else
      
      'check flag primera vez
      If blnFlag = 0 Then
             
        'add config reporte
        'blnB = dsiReportAdd(rs!ID, App.Path & "\reportes\" & rs!columna, "select * from " & rs!valor)
        
        'cambio 18/09/2006
        blnB = dsiReportAdd(rs!ID, SQLparam.ReportesPath & "\" & rs!columna, "select * from " & rs!valor)
            
        'put flag
        blnFlag = 1
            
      Else
        
        'add config subreportes
        blnB = dsiReportAdd(rs!ID, rs!columna, "select * from " & rs!valor)
        
      End If
      
    End Select
    
    'move puntero
    rs.MoveNext
    
  Wend
      
  'get parametros
  strSQL = "select * " & _
           "from dsiOPCparam " & _
           "where IDmenu = " & SQLparam.IDmenu & " And IDopc = " & Me.dsiIDopc
  
  'exec
  Set rs = SQLexec(strSQL)
      
  'chequeo errores
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    End
  End If
        
  'inicializo
  m_dsiParamCant = 0
  
  'while parametros
  While Not rs.EOF
    
    'add parametros
      blnB = Me.dsiParamAdd(rs!nombre, rs!tipo, rs!lista, rs!Formula)
            
    'move puntero
    rs.MoveNext
    
  Wend
      
  'personalizar grilla
  blnB = Me.dsiGrillaCambiarApariencia()
  
  'config grilla
  blnB = Me.dsiGrillaConfig()
  
  'no permite ordenar filas
  SPD.UserColAction = UserColActionDefault
  
  'agrego que con ENTER baje a la proxima fila
  SPD.EditEnterAction = EditEnterActionDown
  
  'elimino header de filas
  SPD.RowHeadersShow = False
  
  'elimino barra de desplazamiento
  SPD.ScrollBars = ScrollBarsNone
  
  'set dato a ingresar que se reemplace
  SPD.EditModeReplace = True
  
  'cambio colo fondo a frm
  Me.BackColor = SPD.GrayAreaBackColor
    
End Sub


Private Sub Form_Resize()
  Dim lngGrilla, lngBotonAceptar As Long
  
  'ajusto top
  Me.SPD.Left = 150
    
  'ajusto ancho grilla
  Me.SPD.Width = IIf(Me.ScaleWidth - 300 > 0, Me.ScaleWidth - 300, 0)
    
  'ajusto columnas grilla
  Dim sngWidth As Single
  Me.SPD.TwipsToColWidth Me.SPD.Width - 1050, sngWidth
  Me.SPD.ColWidth(1) = sngWidth * 0.44                  'nombre parametro
  Me.SPD.ColWidth(2) = sngWidth * 0.08                  'operador
  Me.SPD.ColWidth(3) = sngWidth * 0.44                  'criterio
  Me.SPD.ColWidth(4) = sngWidth * 0.03                  'boton para comboBox
  
  'ajusto boton aceptar
  cmdAceptar.Left = Me.ScaleWidth - Me.cmdAceptar.Width - 100
  cmdAceptar.Top = Me.ScaleTop + 50
  
  'ajusto boton cancelar
  cmdCancelar.Left = Me.ScaleWidth - Me.cmdAceptar.Width - 100
  cmdCancelar.Top = Me.cmdAceptar + Me.cmdAceptar.Height + 100
  
  'guardo altura de grilla y aceptar
  lngGrilla = Me.SPD.Height
  lngBotonAceptar = cmdAceptar.Height
    
  'crystal top
  rptVisor.Top = Me.SPD.Top + Me.SPD.Height + 100
  
  'ajusto crystal
  Me.rptVisor.Left = 0
  Me.rptVisor.Width = Me.ScaleWidth
  
  'para que no de error y que el height del crystal no sea negativo
  If Me.ScaleHeight - Me.SPD.Height - 100 > 50 Then
    rptVisor.Height = IIf(Me.ScaleHeight - Me.SPD.Height - 100 > 0, Me.ScaleHeight - Me.SPD.Height - 100, 0)
  End If
  
  'check si state normal, hago activa la celda 1,1 y luego 3,1
  'esto es porque cuando minimizo y luego maximizo la grilla queda corrida
  If Me.WindowState = 0 Or Me.WindowState = 2 Then
    
    'set celda activa
    Me.SPD.SetActiveCell 1, 1
    Me.SPD.SetActiveCell 3, 1
    
  End If
  
End Sub

Property Get dsiRefresh()
  
  Dim rptRpt() As CRAXDRT.Report
  Dim appRpt As New CRAXDRT.Application
  Dim arrSubRpt, arrFormula As Variant
  Dim rs() As ADODB.Recordset
  Dim intCuenta As Integer
  Dim intFila As Integer
  Dim blnB As Boolean
  Dim strT As String
          
  'change mouse
  Screen.MousePointer = vbHourglass
        
  'while propiedades reportes y subreportes
  For intFila = 1 To m_dsiReportCant
      
    'reservo un elemento de un array para referenciar un nuevo Reporte o Subreporte
    ReDim Preserve rptRpt(intFila)
      
    'check fila 1 reporte principal, sino subreportes
    If intFila = 1 Then
        
      'open reporte principal
      Set rptRpt(intFila) = appRpt.OpenReport(Me.dsiReport(2, intFila))
       
    Else
        
      'open subreportes asociados
      Set rptRpt(intFila) = rptRpt(1).OpenSubreport((Me.dsiReport(2, intFila)))
                
    End If
      
    'reservo un elemento de un array para referenciar el recordset asociado al reporte o subreporte
    ReDim Preserve rs(intFila)
      
    'antes de abrir el rs asociado al reporte, check si hay caracteristicas adicionales
    'DataIDReporte:             ID reporte como fue definido
    'Me.DataReport(intFila, 1): ID del reporte
    'Me.DataReport(intFila, 2): nombre del reporte
    'Me.DataReport(intFila, 3): vista asociada al reporte
    'rptRpt(intFila):           referencia al reporte
    'rs(intFila):               referencia al rs
    'Me.DataWhere:              string con condicion
    'Me.DataWhereArr:           array con condicion por separado
    blnB = Me.dsiRptAdicional(Me.dsiReport(1, intFila), Me.dsiReport(2, intFila), Me.dsiReport(3, intFila), rptRpt(intFila), rs(intFila), Me.dsiCondicion, Me.dsiCondicionArr)
      
    'check si no hay caracteristicas adicionales, abro rs con la condicion seleccionada
    If blnB = False Then
        
      'si existe vista asociado al reporte la abro
      If Not IsEmpty(Me.dsiReport(3, intFila)) Then
          
        'save source
        strT = Me.dsiReport(3, intFila)
          
        'si el usuario realizo un filtro lo agrego
        If Me.dsiCondicion <> "" Then
          strT = strT & " where " & Me.dsiCondicion
        End If
            
        'abro recordset si existe
        Set rs(intFila) = SQLexec(strT)
          
        'check errores
        If Not SQLparam.CnErrNumero = -1 Then
          SQLError                            'show error
          Screen.MousePointer = vbDefault     'change mouse
          Exit Property                       'exit property
        End If
        
      End If
          
    End If
      
    'set rs al reporte
    rptRpt(intFila).Database.SetDataSource rs(intFila)
      
  Next
            
            
  'while formulas para pasarle al crystal
  For intCuenta = 1 To m_dsiFormulasCant
      
    rptRpt(1).FormulaFields.GetItemByName(Me.dsiFormulasGetNombre(intCuenta)).Text = "'" & Replace(Me.dsiFormulasGetValor(intCuenta), "'", "") & "'"
    
  Next
       
  'le paso la condicion al crystal entre comillas simples, sino da error
    
  'while formulas reporte
  For intCuenta = 1 To rptRpt(1).FormulaFields.Count
    
    'check si existe formula @condicion
    If rptRpt(1).FormulaFields(intCuenta).Name = "{@parametrosFiltrados}" Then
      
      'put condicion
      rptRpt(1).FormulaFields.GetItemByName("parametrosFiltrados").Text = "'" & Replace(Me.dsiCondicion, "'", "") & "'"
      
    End If
    
  Next
      
  'set report a view
  rptVisor.ReportSource = rptRpt(1)
  
  'show view
  Me.rptVisor.ViewReport
    
  'mouse default
  Screen.MousePointer = vbDefault
    
End Property

'agrega item seleccionado en Lista
'
Public Property Get dsiColumnasItemAgregar(Item As Integer, blnB As Boolean) As Boolean
  
  Dim intI As Integer
    
  'agrego elemento
  If blnB Then
    
    'agrego un elemento
    ReDim Preserve m_dsiColumnasItem(Me.lstLista.SelCount)
    
    'agrego item
    m_dsiColumnasItem(Me.lstLista.SelCount) = Item
    
  'elimino elemento
  Else
    
    'busco si el item en el cual se hizo click ya existe
    For intI = 1 To UBound(m_dsiColumnasItem)
      
      'busco item, cuando encuentro, put 0, como si lo estuviera borrando
      If m_dsiColumnasItem(intI) = Item Then
        m_dsiColumnasItem(intI) = -1
      End If
        
    Next
    
  End If
    
End Property

Public Property Get dsiColumnasItem(intI As Integer) As String
  dsiColumnasItem = m_dsiColumnasItem(intI)
End Property


Property Get dsiReport(intFila As Integer, intCol As Integer) As String
  dsiReport = m_dsiReport(intFila, intCol)
End Property


Property Get dsiCondicionArr() As Variant
  dsiCondicionArr = m_dsiCondicionArr
End Property


Property Get dsiCondicion() As String
  dsiCondicion = m_dsiCondicion
End Property


Private Sub lstLista_ItemCheck(Item As Integer)
  
  Dim blnB As Boolean
  
  'agrego o elimino segun click o unClick
  blnB = Me.dsiColumnasItemAgregar(Item, Me.lstLista.Selected(Item))
  
End Sub



Private Sub SPD_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
  Dim blnB As Boolean
  Dim strT As String
  Dim varColumna, varLista As Variant
  Dim rs As ADODB.Recordset
  
  'get source
  Me.SPD.GetText 6, Me.SPD.ActiveRow, varLista
  
  'check si source para mostrar en lista
  If varLista = "" Then
    blnB = MsgBox("No hay datos para mostrar, el criterio debe ser ingresado.", vbInformation + vbOKOnly, "Atención...")
    Exit Sub
  End If
  
  'move frame
  Me.frame.Move Me.SPD.Left + Me.SPD.ColWidth(1) + Me.SPD.ColWidth(2) + Me.SPD.ColWidth(3) - 2480, Me.SPD.Top + Me.SPD.RowHeight(0) + 20, 2500, (MainMDI.ScaleHeight / 2) + 250
      
  'move lista
  Me.lstLista.Move 20, 20, 2500 - 40, (MainMDI.ScaleHeight / 2)
  
  'move botones aceptar cancelar
  Me.cmdLstCancelar.Move 20 + Me.lstLista.Width - Me.cmdLstCancelar.Width - 40, 20 + Me.lstLista.Height + 70
  Me.cmdLstAceptar.Move Me.cmdLstCancelar.Left - Me.cmdLstAceptar.Width - 200, 20 + Me.lstLista.Height + 70
  
  'cambio mouse
  Me.MousePointer = vbHourglass
    
  'armo query
  strT = varLista
      
  'get rs
  Set rs = SQLexec(strT)
    
  'check error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError                                    'muestro error
    Me.MousePointer = vbDefault                 'cambio mouse
    Exit Sub                                    'exit function
  End If
    
  'vacio lista
  lstLista.Clear
  
  'recorro rs
  While Not rs.EOF
    
    'agrego dato a lista
    Me.lstLista.AddItem IIf(IsNull(rs(0)), "", rs(0))
                
    'puntero proximo
    rs.MoveNext
    
  Wend
  
  'cambio mouse
  Me.MousePointer = vbDefault
  
  'set visible lista
  Me.frame.Visible = True
  
  'set foco en lst
  Me.lstLista.SetFocus
  
End Sub


'armo condicion final
'
Public Property Get dsiCondicionArmar()
  
  Dim blnB As Boolean
  Dim intI, intIcant As Integer
  Dim blnColumnaTexto As Boolean
  Dim strParcial, strParcialFormula, strParcialStore As String
  Dim strFinal, strGeneral, strVista, strColumna As String
  Dim varColumna, varOperador, varCriterio, varTipo, varNoWhere, arrCriterio As Variant
  Dim rs As ADODB.Recordset
  
  'inicializo
  strFinal = ""
  strParcialStore = ""
  
  'inicializo formulas
  blnB = Me.dsiFormulasLimpiar()
  
  'while parametros
  For intI = 1 To Me.SPD.MaxRows
    
    'get columna, operador, criterio
    Me.SPD.GetText 1, intI, varColumna
    Me.SPD.GetText 2, intI, varOperador
    Me.SPD.GetText 3, intI, varCriterio
    Me.SPD.GetText 5, intI, varTipo
    Me.SPD.GetText 7, intI, varNoWhere
        
    'save si existe filtro
    If varCriterio <> "" Then
      
      'inicializo
      strParcial = ""
      strParcialFormula = ""
      
      'set columna de tipo texto
      blnColumnaTexto = False
      
      'separate criterio en array para saber si hay mas de un criterio
      arrCriterio = separateText(varCriterio, ",")
                  
      'while criterios
      For intIcant = 1 To UBound(arrCriterio)
                                  
        'case tipo de celda
        Select Case LCase(varTipo)
          
        'booleano
        Case "sino"
            
          'move true = 1, false = 0, sino valor
          Select Case LCase(arrCriterio(intIcant))
          Case "true"
            arrCriterio(intIcant) = 1
          Case "false"
            arrCriterio(intIcant) = 0
          Case Else
            arrCriterio(intIcant) = varCriterio
          End Select
            
          'add condicion parcial
          strParcial = strParcial & arrCriterio(intIcant) & ","
                    
          'add condicion para formula
          strParcialFormula = strParcialFormula & arrCriterio(intIcant) & ","
          
        'numerico
        Case "numero"
          
          'add condicion parcial
          strParcial = strParcial & arrCriterio(intIcant) & ","
          
          'add condicion para formula
          strParcialFormula = strParcialFormula & arrCriterio(intIcant) & ","
          
        'fecha
        Case "fecha"
          
          'put comillas y format dateTOiso
          arrCriterio(intIcant) = "'" & dateToIso(arrCriterio(intIcant)) & "'"
            
          'add condicion parcial
          strParcial = strParcial & arrCriterio(intIcant) & ","
            
          'add condicion para formula
          strParcialFormula = strParcialFormula & arrCriterio(intIcant) & ","
            
        'texto
        Case "texto"
            
          'si = like, sino not like
          If LCase(varOperador) = "=" Then
            
            'put comillas
            arrCriterio(intIcant) = "'" & arrCriterio(intIcant) & "'"
            
            'add condicion parcial
            strParcial = strParcial & "[" & varColumna & "] like " & arrCriterio(intIcant) & " or "
            
            'add condicion para formula
            strParcialFormula = strParcialFormula & arrCriterio(intIcant) & ","
            
          End If
            
          'set columna string
          blnColumnaTexto = True
          
        End Select
                
      Next
                             
      'check si hay una (,) en la ultima posicion, delete
      If Mid(strParcial, Len(strParcial), 1) = "," Then
      
        strParcial = Left(strParcial, Len(strParcial) - 1)
        strParcialFormula = Left(strParcialFormula, Len(strParcialFormula) - 1)
      
      End If
      
      'check si columna tipo texto
      If blnColumnaTexto Then
        
        'check si hay un (or) en la ultima posicion, delete
        If Mid(strParcial, Len(strParcial) - 3, 4) = " or " Then
          
          strParcial = Left(strParcial, Len(strParcial) - 4)
          strParcialFormula = Left(strParcialFormula, Len(strParcialFormula) - 1)
        
        End If
        
        'check si hay un (and) en la ultima posicion, delete
        If Mid(strParcial, Len(strParcial) - 4, 5) = " and " Then
          
          strParcial = Left(strParcial, Len(strParcial) - 5)
          strParcialFormula = Left(strParcialFormula, Len(strParcialFormula) - 1)
        
        End If
      
      End If
      
      'check si operador [] y cantidad de criterios 2, OK, sino, error
      If LCase(varOperador) = "[]" Then
        
        If UBound(arrCriterio) = 2 Then
          strParcial = "[" & varColumna & "]" & " between " & arrCriterio(1) & " and " & arrCriterio(2)
        Else
          blnB = MsgBox("Para el operador [] (entre), se debe ingresar 2 valores separados por una coma." & vbCrLf & vbCrLf & _
                 "ejemplo: Texto: xx,xx. Número: 99,99. Fecha: dd/mm/yyyy,dd/mm/yyyy." & vbCrLf & vbCrLf & _
                 "Cuando se ingresen los 2 valores, " & varColumna & "  será tomada en cuenta.", vbInformation + vbOKOnly, "Atención...")
        End If
        
      End If
      
      'check si operador =
      If LCase(varOperador) = "=" Then
        
        'check si tipo texto
        If blnColumnaTexto Then
          strParcial = "(" & strParcial & ")"
        Else
          strParcial = "[" & varColumna & "] in " & "(" & strParcial & ")"
        End If
        
      End If
      
      'check si criterio no debe formar parte de la condicion, es formula, sino condicion
      If varNoWhere Then
        
        'add formula
        blnB = Me.dsiFormulasAgregar(varColumna, strParcialFormula)
          
      Else
        
        'add condicion
        strFinal = strFinal & strParcial & " and "
        
        'add parametros para store procedure
        strParcialStore = strParcialStore & strParcialFormula & ","
            
      End If
      
    End If
        
  Next
  
  'check si hay un (and) en la ultima posicion, delete
  If Len(strFinal) <> 0 Then
      
    'si en las ultimas posiciones hay un and
    If Mid(strFinal, Len(strFinal) - 4, 5) = " and " Then
      strFinal = Left(strFinal, Len(strFinal) - 5)
    End If
    
  End If
  
  'check si hay una (,) en la ultima posicion, delete
  If Len(strParcialStore) <> 0 Then
      
    'si en las ultimas posiciones hay un and
    If Mid(strParcialStore, Len(strParcialStore), 1) = "," Then
      strParcialStore = Left(strParcialStore, Len(strParcialStore) - 1)
    End If
    
  End If
  
  'save condicion
  m_dsiCondicion = strFinal
      
  'save parametros para store
  m_dsiSPparam = strParcialStore
      
End Property




'
'SIRVE PARA ESPECIFICAR CARACTERISTICAS ADICIONALES A LOS REPORTES
'
'strID:         ID reporte o subreporte como esta definido en menuOpciones
'strNBReport:   string con nombre del rpt
'strDataSource: string con vista asociada al rpt
'rpt:           referencia al reporte o subreportes
'rs:            puntero al recordset que todavia no fue abierto
'strWhere:      string con where
'strWhereArr:   array con where separado para poder acceder a cada condicion por separado
'
Public Property Get dsiRptAdicional(strID As String, strNBReport As String, strDataSource As String, rpt As CRAXDRT.Report, rs As ADODB.Recordset, strWhere As String, strWhereArr As Variant) As Boolean
  Dim strWhereActual As String
  Dim strIntervalo As String
  
  'valor predeterminado
  dsiRptAdicional = False
  
'  Select Case LCase(strID)
    
  'ejemplo
  '
'  Case "205"
      
'    'armo where
'    strWhereActual = ""
'    If strWhere <> "" Then
'      strWhereActual = " where " & strWhere
'    End If
      
    'predeterminado 3 por si no se selecciono intervalo o se ingreso algo distinto de 3,6,9
'    strSQL = " select * from graficosBaseIIP_avg3_vw " & strWhereActual
      
'    If arrayGetValue(strWhereArr, "intervalo") <> "" Then
      
      'determino intervalo
'      Select Case arrayGetValue(strWhereArr, "intervalo")
      
'      Case 2
'        strSQL = " select * from graficosBaseIIP_avg2_vw" & strWhereActual
'      Case 3
'        strSQL = " select * from graficosBaseIIP_avg3_vw" & strWhereActual
'      Case 5
'        strSQL = " select * from graficosBaseIIP_avg5_vw" & strWhereActual
'      Case 6
'        strSQL = " select * from graficosBaseIIP_avg6_vw" & strWhereActual
'      Case 12
'        strSQL = " select * from graficosBaseIIP_avg12_vw" & strWhereActual
'      End Select
      
'    End If
    
'    'abro rs y se lo paso a rpt
'    Set rs = SQLexec(strSQL)
'    rpt.Database.SetDataSource rs
'
'    dsiRptAdicional = True
'
'  End Select

End Property

