VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpspr60.ocx"
Begin VB.Form procesosFRM 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9390
   Icon            =   "procesosFRM.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2715
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frame 
      Height          =   1365
      Left            =   5790
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      Begin VB.ListBox lstLista 
         Height          =   735
         ItemData        =   "procesosFRM.frx":0CCA
         Left            =   90
         List            =   "procesosFRM.frx":0CCC
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   150
         Width           =   1725
      End
      Begin VB.CommandButton cmdLstCancelar 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   990
         TabIndex        =   5
         Top             =   990
         Width           =   825
      End
      Begin VB.CommandButton cmdLstAceptar 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   990
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      Height          =   300
      Left            =   4590
      TabIndex        =   2
      Top             =   480
      Width           =   1080
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   300
      Left            =   4590
      TabIndex        =   1
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
      SpreadDesigner  =   "procesosFRM.frx":0CCE
      ClipboardOptions=   3
   End
End
Attribute VB_Name = "procesosFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_IDmenu As String            'save IDmenu
Private m_IDopc As String             'save IDopcion
Private m_Titulo As String            'save titulo

Dim m_dsiCondicion As String          'save condicion en string
Dim m_dsiCondicionArr As Variant      'save condicion separada en array

Dim m_dsiFormulas() As String         'save formulas
Dim m_dsiFormulasCant As Integer      'save cantidad de formulas

Dim m_dsiColumnasItem() As Variant    'items de la lista en los cuales se hizo click
Dim m_dsiParamConfig() As String      'save parametros
Dim m_dsiParamCant As Integer         'save cantidad de parametros

Dim m_dsiSP As String                 'save store procedure
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

Public Property Let dsiSP(strSP As String)
  m_dsiSP = strSP
End Property

Public Property Get dsiSP() As String
  dsiSP = m_dsiSP
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
    Me.SPD.row = intFila
      
    'put nombre parametro
    Me.SPD.col = 1
    Me.SPD.CellType = CellTypeStaticText
    Me.SPD.Text = m_dsiParamConfig(1, intFila)
      
    'set puntero columna operadores
    Me.SPD.col = 2
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
    Me.SPD.col = 3
    Me.SPD.CellType = CellTypeEdit
    Me.SPD.TypeMaxEditLen = 5000
      
    'set columna boton para que cuando se hace click levanta lista con datos
    Me.SPD.col = 4
    Me.SPD.CellType = CellTypeButton
    Me.SPD.TypeButtonText = "..."
        
    'set columna tipo de parametro: fecha, numero, texto, sino
    Me.SPD.SetText 5, intFila, m_dsiParamConfig(2, intFila)
          
    'set columna LISTA
    Me.SPD.col = 6
    Me.SPD.TypeMaxEditLen = 500
    Me.SPD.SetText 6, intFila, m_dsiParamConfig(3, intFila)
          
    'set columna para que el parametro no lo incluya en el where
    Me.SPD.SetText 7, intFila, m_dsiParamConfig(4, intFila)
      
  Next
    
  'set celda activa
  Me.SPD.SetActiveCell 3, 1
  
End Property

Private Sub cmdAceptar_Click()
  
  Dim blnB As Boolean
   
  'check si no existe sp, mensaje y exit
  If Me.dsiSP = "" Then
    blnB = MsgBox("Esta intentando ejecutar un procedimiento que no existe.", vbCritical + vbOKOnly, "Atención...")
    Exit Sub
  End If
      
  'build condicion
  blnB = Me.dsiCondicionArmar()
    
  'exec store
  SQLexec ("exec " & Me.dsiSP & " " & Me.dsiSPparam)
    
  'check errores
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    Exit Sub
  End If
      
  blnB = MsgBox("El proceso finalizo correctamente.", vbInformation + vbOKOnly, "Atención...")
      
End Sub


Private Sub cmdCancelar_Click()
      
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
      
    Me.dsiSP = IIf(IsNull(rs!valor), "", rs!valor)
      
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
  Me.SPD.ColWidth(1) = sngWidth * 0.43                  'nombre parametro
  Me.SPD.ColWidth(2) = sngWidth * 0.1                   'operador
  Me.SPD.ColWidth(3) = sngWidth * 0.43                  'criterio
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
    
  'check si state normal, hago activa la celda 1,1 y luego 3,1
  'esto es porque cuando minimizo y luego maximizo la grilla queda corrida
  If Me.WindowState = 0 Then
    
    'set celda activa
    Me.SPD.SetActiveCell 1, 1
    Me.SPD.SetActiveCell 3, 1
    
  End If
  
End Sub


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



Private Sub SPD_ButtonClicked(ByVal col As Long, ByVal row As Long, ByVal ButtonDown As Integer)
    
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

