VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form gridFRM 
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   Icon            =   "gridFRM.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3105
   ScaleWidth      =   6360
   Begin MSComDlg.CommonDialog comArchivos 
      Left            =   5640
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpreadADO.fpSpread spdGrid 
      Height          =   1650
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   5085
      _Version        =   393216
      _ExtentX        =   8969
      _ExtentY        =   2910
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
      SpreadDesigner  =   "gridFRM.frx":038A
   End
   Begin FPSpreadADO.fpSpread spdE 
      Height          =   945
      Left            =   120
      TabIndex        =   1
      Top             =   1980
      Visible         =   0   'False
      Width           =   5100
      _Version        =   393216
      _ExtentX        =   8996
      _ExtentY        =   1667
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
      SpreadDesigner  =   "gridFRM.frx":0588
   End
   Begin MSComDlg.CommonDialog comDestino 
      Left            =   5640
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "gridFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_IDmenu As String                  'ID menu
Private m_IDopc As String                   'ID opcion
Private m_Titulo As String                  'titulo opcion

Private m_Operacion As String               'operacion INS, EDI, ELI, IMP
Private m_Formulario As String
Private m_FormularioAnt As String
Private m_spINSpermiso As Boolean           'store para INS
Private m_spEDIpermiso As Boolean           'store para EDI
Private m_spELIpermiso As Boolean           'store para ELI
Private m_spIMPpermiso As Boolean           'store para IMP
Private m_Esquema As ADODB.Fields           'estructura de la vista

Private m_ComboBoxPrimeraVez As Boolean     'para llenar los comboBox una sola vez

Private m_Config() As String                'configuración de la grilla
Private m_ConfigCant As Integer             'cantidad de filas de configuración

Private m_GrillaIniFila As String
Private m_GrillaIniCol As String

Private m_ColumnasDecimales() As Variant
Private m_ColumnasAncho() As Variant
Private m_ColumnasOcultas() As Variant
Private m_ColumnasCongeladas As String

Private m_Condicion() As Variant
Private m_CondicionCant As Integer

Private m_Where As String
Private m_OrderBy As String

Private m_AltoAux As Long
Private m_AnchoAux As Long
Private m_FormularioActivado As Boolean

Const SW_SHOWNORMAL = 1
'** 08/01/08 --
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Property Let dsiFormularioActivado(blnB As Boolean)
  m_FormularioActivado = blnB
End Property

Public Property Get dsiFormularioActivado() As Boolean
  dsiFormularioActivado = m_FormularioActivado
End Property

Public Property Get dsiFormularioDoble() As Boolean
  
  'si cantidad de filas * altura de cada fila > al 60% del tamaño de pantalla, es doble
  dsiFormularioDoble = Me.dsiEsquema.Count * 250 > Screen.Height * 0.6

End Property

Public Property Let dsiFormulario(strT As String)
  m_Formulario = strT
End Property

Public Property Get dsiFormulario() As String
  dsiFormulario = m_Formulario
End Property

Public Property Let dsiFormularioAnt(strT As String)
  m_FormularioAnt = strT
End Property

Public Property Get dsiFormularioAnt() As String
  dsiFormularioAnt = m_FormularioAnt
End Property

Public Property Get dsiHerramientas(blnB As Boolean) As Boolean
  
  MainMDI.tlbHerra.Buttons("cmd_forma_vista").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_fila_pri").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_fila_ant").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_fila_sig").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_fila_ult").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_actualizar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_insertar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_editar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_eliminar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_buscar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_filtro_abr").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_filtro_gua").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_filtro_rap").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_filtro_ava").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_filtro_atr").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_filtro_eli").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_columna_orden_asc").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_columna_orden_desc").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_columna_ajustar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_columna_mas_decimales").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_columna_menos_decimales").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_columna_ocultar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_columna_mostrar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_columna_congelar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_columna_descongelar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_copiar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_totalizar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_exportar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_importar").Enabled = blnB

End Sub

Public Property Get dsiHerramientasComfirmar(blnB As Boolean) As Boolean
  
  MainMDI.tlbHerra.Buttons("cmd_guardar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_cancelar").Enabled = blnB

End Property

Public Property Get dsiFormularioCambiar() As Boolean
    
  'visualizar form en forma de planilla
  If (m_Formulario = "F" Or m_Formulario = "") Then
        
    'hago visible grilla para planilla
    Me.spdGrid.Visible = True
    Me.spdE.Visible = False
    
    'cuando cambio la forma de ver el form de formulario a planilla
    'recupera el tamaño origal de form, simpre y cuando este visible
    If Me.spdGrid.Visible Then
      Me.Height = Me.dsiAltoAnt
      Me.Width = Me.dsianchoAnt
    End If
        
    m_Formulario = "P"
      
  ElseIf m_Formulario = "P" Then
        
    'guardo tamaño de original de form, esto lo hago porque con
    'esta forma de ver los datos, se modifica el tamaño del form
    Me.dsiAltoAnt = Me.Height
    Me.dsianchoAnt = Me.Width
        
    'check si comboBox todavia no fueron configurados, lo hago, sino sigo de largo
    If Not Me.dsiComboBoxPrimeraVez Then
          
      'dibuja pantalla
      blnB = Me.dsiFormularioDibujar()
      
      'set celdas obligatorias, bloqueadas, listas
      blnB = Me.dsiFormularioDatosConfig()
      
      'set comboBox configurados por primera vez
      Me.dsiComboBoxPrimeraVez = True
      
    End If
      
    'set ajustar form
    blnB = Me.dsiFormularioAjustar()
      
    'set pasa el valor de las celdas de spdGrid a las celdas de spdE
    blnB = Me.dsiFormularioDatosPutPaF()
     
    'hago visible grilla para formulario
    Me.spdGrid.Visible = False
    Me.spdE.Visible = True
     
    m_Formulario = "F"
    
  End If
  
End Property

Public Property Let dsiComboBoxPrimeraVez(ByVal blnB As Boolean)
  m_ComboBoxPrimeraVez = blnB
End Property

Public Property Get dsiComboBoxPrimeraVez() As Boolean
  dsiComboBoxPrimeraVez = m_ComboBoxPrimeraVez
End Property

Public Property Let dsiAltoAnt(ByVal lngL As Long)
  m_AltoAux = lngL
End Property

Public Property Get dsiAltoAnt() As Long
  dsiAltoAnt = m_AltoAux
End Property

Public Property Let dsianchoAnt(ByVal lngL As Long)
  m_AnchoAux = lngL
End Property

Public Property Get dsianchoAnt() As Long
  dsianchoAnt = m_AnchoAux
End Property

Public Property Let dsiOperacion(ByVal str As String)
  m_Operacion = str
End Property

Public Property Get dsiOperacion() As String
  dsiOperacion = m_Operacion
End Property

Public Property Get dsiFormularioDibujar() As Boolean
  
  Dim intCol, intFila, intDecimales As Integer
  Dim varTitulo, varDato As Variant
  Dim intEsq As Integer
  Dim sngTotal, sngAnchoMaximo As Single
  
  'cambio caption del form edit
  'frmEdit.Caption = frmBase.Caption & " - " & UCase(strOperacion)
      
  'set cantidad columnas
  Me.spdE.MaxCols = 6
      
  'set ancho de columnas
  Me.spdE.ColWidth(1) = 2000
  Me.spdE.ColWidth(2) = 2000
  Me.spdE.ColWidth(3) = 0
  Me.spdE.ColWidth(4) = 0
  Me.spdE.ColWidth(5) = 0
  Me.spdE.ColWidth(6) = 0
      
  'si es formulario doble
  If Me.dsiFormularioDoble() Then
    
    'set cantidad filas
    Me.spdE.MaxRows = CInt(Format(Me.dsiEsquema.Count / 2, "##"))
    
    'set ancho columnas
    Me.spdE.ColWidth(4) = 2000
    Me.spdE.ColWidth(5) = 2000
    Me.spdE.ColWidth(6) = 0
    
  Else
    
    'set cantidad filas
    Me.spdE.MaxRows = Me.dsiEsquema.Count
    
  End If
      
  'set modifica grilla en BackGround
  Me.spdE.ReDraw = False
    
  'set valores iniciales
  intCol = 1
  intFila = 1
  intEsq = 0
    
  'recorre esquema
  While intEsq <= Me.dsiEsquema.Count - 1
    
    'si fila = cantidad de filas de grilla, la pantalla es doble
    If intEsq = Me.spdE.MaxRows Then
      intCol = 4
      intFila = 1
    End If
        
    'get puntero a fila y columna activa en spdGrid
    Me.spdGrid.row = Me.spdGrid.ActiveRow
    Me.spdGrid.col = intEsq + 1
    
    'check si tipo celda numerica, get cantidad de decimales
    intDecimales = 0
    If Me.dsiEsquema(intEsq).Type = conMoney Or Me.dsiEsquema(intEsq).Type = conSmallMoney Or Me.dsiEsquema(intEsq).Type = conReal Or Me.dsiEsquema(intEsq).Type = conFloat Or Me.dsiEsquema(intEsq).Type = conNumeric Or Me.dsiEsquema(intEsq).Type = conDecimal Then
    
      'check si decimales en 0, check si fila 0
      If Me.spdGrid.TypeNumberDecPlaces = 0 And Me.spdGrid.row = 0 Then
        intDecimales = 2
      Else
        intDecimales = Me.spdGrid.TypeNumberDecPlaces
      End If
    End If
      
    'set puntero en fila en spdE
    Me.spdE.row = intFila
    
    'set puntero en columna para titulos
    Me.spdE.col = intCol
    
    'set titulo de columna
    Me.spdE.SetText intCol, intFila, Me.dsiEsquema(intEsq).Name
        
    'set altura fila
    Me.spdE.RowHeight(intFila) = 255
        
    'set color fondo gris
    Me.spdE.BackColor = RGB(240, 240, 240)
          
    'set color texto gris oscuro
    'cambio 05/04/2005, ahora texto negro
    'Me.spdE.ForeColor = RGB(131, 131, 131)
      
    'tamaño de letra para titulos
    Me.spdE.FontSize = 10
      
    'set titulos estaticos
    Me.spdE.CellType = CellTypeStaticText
    
    'set puntero en columna para datos
    Me.spdE.col = intCol + 1
      
    'set backColor
    Me.spdE.BackColor = RGB(255, 255, 255)
      
    'set foreColor para datos
    Me.spdE.ForeColor = RGB(105, 105, 105)
          
    'tamaño de letra para ingreso de datos
    Me.spdE.FontSize = 9
      
    'set celda lock
    Me.spdE.Lock = True
                      
    'case tipo de celda
    Select Case Me.dsiEsquema(intEsq).Type
          
      'Bit
      Case conBit
        
        Me.spdE.CellType = CellTypeCheckBox
        Me.spdE.Value = 0
          
      'enteros
      Case conSmallInt, conInt, conTinyInt
        
        Me.spdE.CellType = CellTypeNumber
        Me.spdE.TypeNumberDecPlaces = 0
        Me.spdE.TypeHAlign = TypeHAlignRight
                   
      'decimal
      Case conMoney, conSmallMoney, conReal, conFloat, conNumeric, conDecimal
            
        Me.spdE.CellType = CellTypeNumber
        Me.spdE.TypeNumberDecPlaces = intDecimales
        Me.spdE.TypeNumberShowSep = False
        Me.spdE.TypeHAlign = TypeHAlignRight
      
      'fecha
      Case conSmallDateTime, conDateTime
      
        Me.spdE.CellType = CellTypeDate
        Me.spdE.TypeDateFormat = TypeDateFormatDDMMYY
        Me.spdE.TypeDateCentury = True
        Me.spdE.TypeDateMin = "01011000"
        Me.spdE.Text = "06/17/00"
                                        
      'string
      Case conChar, conNchar, conVarchar, conText, conNVarchar, conNtext
        
        'set tipo edit
        Me.spdE.CellType = CellTypeEdit
        
        'set limite de digitos segun SQL
        If Me.dsiEsquema(intEsq).Type <> conNtext Then
          Me.spdE.TypeMaxEditLen = Me.dsiEsquema(intEsq).DefinedSize
        Else
          Me.spdE.TypeMaxEditLen = 5000        'para tipo de dato nText
          Me.spdE.RowHeight(intFila) = 1000
          Me.spdE.TypeEditMultiLine = True
        End If
        
        'set le agrega puntos suspensivos.... cuando no se ve todo el texto
        Me.spdE.TypeEllipses = True

'cambio 27/01/2005
'        'set altura fila
'        If Me.dsiEsquema(intEsq).DefinedSize > 50 Then
'          Me.spdE.RowHeight(intFila) = 1000
'          Me.spdE.TypeEditMultiLine = True
'        End If
        
      End Select
        
    'fila proxima y columna proxima
    intFila = intFila + 1
    intEsq = intEsq + 1
        
  Wend
          
  'esto solo se procesa para formularios doble y esquemas impares,
  'es para pintar la ultima fila de la derecha
  For intI = intFila To Me.spdE.MaxRows
        
    'set puntero a fila columna para titulos
    Me.spdE.row = intI
    Me.spdE.col = 4
        
    'set color fondo y celda estatica
    Me.spdE.BackColor = RGB(250, 250, 250)
    Me.spdE.CellType = CellTypeStaticText
    
    'set puntero a columna para datos
    Me.spdE.col = 5
          
    'set color fondo y bloqueo celda
    Me.spdE.BackColor = RGB(238, 238, 238)
    Me.spdE.Lock = True
        
  Next
        
  'el ancho de columnas lo establece automaticamente a la columna mas ancha
  'si la columna mas ancho es menor que 2000, set default 2000
  
  'defa ancho en 2000
  sngAnchoMaximo = 2000
  
  'get ancho maximo formulario simple
  If Me.spdE.MaxTextColWidth(1) + 50 > 2000 Then
    sngAnchoMaximo = Me.spdE.MaxTextColWidth(1) + 50
  End If
  
  'si formulario doble
  If Me.dsiFormularioDoble() Then
    
    'get ancho formulario doble, si es > a ancho formulario simple
    If Me.spdE.MaxTextColWidth(4) + 50 > sngAnchoMaximo Then
      sngAnchoMaximo = Me.spdE.MaxTextColWidth(4) + 50
    End If
    
    'set ancho columnas formulario doble
    Me.spdE.ColWidth(1) = sngAnchoMaximo
    Me.spdE.ColWidth(2) = sngAnchoMaximo
    Me.spdE.ColWidth(4) = sngAnchoMaximo
    Me.spdE.ColWidth(5) = sngAnchoMaximo
    
  Else
    
    'set ancho columnas formulario simple
    Me.spdE.ColWidth(1) = sngAnchoMaximo
    Me.spdE.ColWidth(2) = sngAnchoMaximo
    
  End If
    
  'set ubicacion de grilla
  Me.spdE.Left = 50
  Me.spdE.Top = 50
              
  'set altura de grilla
  sngTotal = 0
  For intFila = 0 To Me.spdE.MaxRows
    sngTotal = sngTotal + Me.spdE.RowHeight(intFila) + 20
  Next
  Me.spdE.Height = sngTotal
  
  'set ancho de grilla
  sngTotal = 0
  For intFila = 0 To Me.spdE.MaxCols
    sngTotal = sngTotal + Me.spdE.ColWidth(intFila) + 18
  Next
  Me.spdE.Width = sngTotal
          
  'set modifica grilla en ForeGround
  Me.spdE.ReDraw = True
          
End Property

Public Property Get dsiFormularioAjustar() As Boolean

  'set alto y ancho de formulario, antes de modificar tamaño, cambio estado
  'si llega a estar minimizado o maximizado genera error
  Me.WindowState = 0
  Me.Height = Me.spdE.Top + Me.spdE.Height + 550
  Me.Width = Me.spdE.Width + 230

End Property


'transfiere los datos de la grilla en forma de planilla a la grilla en forma de formulario
'
Public Property Get dsiFormularioDatosPutPaF() As Boolean
  
  Dim intCol, intFila  As Integer
  Dim varDato As Variant
  Dim intEsq As Integer
  
  'set modifica grilla en BackGround
  Me.spdE.ReDraw = False
  
  'set valores iniciales
  intCol = 2
  intFila = 1
  intEsq = 0
  
  'recorre esquema
  While intEsq <= Me.dsiEsquema.Count - 1
    
    'si fila = cantidad de filas de grilla, la pantalla es doble
    If intEsq = Me.spdE.MaxRows Then
      intCol = 5
      intFila = 1
    End If
    
    'get dato de spdGrid
    Me.spdGrid.GetText intEsq + 1, Me.spdGrid.ActiveRow, varDato
      
    'set dato en spdE
    Me.spdE.SetText intCol, intFila, varDato
    
    'fila proxima y columna proxima
    intFila = intFila + 1
    intEsq = intEsq + 1
        
  Wend
        
  'set modifica grilla en ForeGround
  Me.spdE.ReDraw = True
    
End Property

'transfiere los datos de un recordset a spdE (grilla de edicion de datos)
'
Public Property Get dsiFormularioDatosPutRaF(rs As ADODB.Recordset) As Boolean
  
  Dim intCol, intFila  As Integer
  Dim varDato As Variant
  Dim intEsq As Integer
  
  'set valores iniciales
  intCol = 2
  intFila = 1
  intEsq = 0
  
  'recorre esquema
  While intEsq <= Me.dsiEsquema.Count - 1
    
    'si fila = cantidad de filas de grilla, la pantalla es doble
    If intEsq = Me.spdE.MaxRows Then
      intCol = 5
      intFila = 1
    End If
    
    'set dato de spdE
    Me.spdE.SetText intCol, intFila, rs(intEsq)
    
    'fila proxima y columna proxima
    intFila = intFila + 1
    intEsq = intEsq + 1
        
  Wend
    
End Property

'transfiere los datos de la grilla en forma de formulario a la grilla en forma de planilla
'
Public Property Get dsiFormularioDatosPutFaP() As Boolean
  
  Dim intCol, intFila  As Integer
  Dim varDato As Variant
  Dim intEsq As Integer
  
  'set valores iniciales
  intCol = 2
  intFila = 1
  intEsq = 0
  
  'recorre esquema
  While intEsq <= Me.dsiEsquema.Count - 1
    
    'si fila = cantidad de filas de grilla, la pantalla es doble
    If intEsq = Me.spdE.MaxRows Then
      intCol = 5
      intFila = 1
    End If
    
    'get dato de spdE
    Me.spdE.GetText intCol, intFila, varDato
    
    'get dato en spdGrid
    Me.spdGrid.SetText intEsq + 1, Me.spdGrid.ActiveRow, varDato
      
    'fila proxima y columna proxima
    intFila = intFila + 1
    intEsq = intEsq + 1
        
  Wend
    
End Property

'arma un string con la grilla spdE para pasarle la info al store procedure
'
Public Property Get dsiFormularioDatosGet() As String
  
  Dim intCol, intFila, intComboInd As Integer
  Dim varNombre, varValor  As Variant
  Dim intEsq As Integer
  Dim strValor, strCombo, strParam, strObliga, strParcial, strFinal As String
  Dim blnIDNumerico As Boolean
  
  'set valores iniciales
  intCol = 2
  intFila = 1
  intEsq = 0
  
  'recorre esquema
  While intEsq <= Me.dsiEsquema.Count - 1
    
    'si fila = cantidad de filas de grilla, la pantalla es doble
    If intEsq = Me.spdE.MaxRows Then
      intCol = 5
      intFila = 1
    End If
    
    'get nombre
    Me.spdE.GetText intCol - 1, intFila, varNombre
    
    'get valor
    Me.spdE.GetText intCol, intFila, varValor
        
    'get columna NOACTUALIZO
    strValor = dsiConfigGetValor("NOACTUALIZO", varNombre)
    strObliga = dsiConfigGetValor("OBLIGATORIO", varNombre)
    
    'check si columna no esta dentro de las NOACTUALIZO
    If strValor = "" Then
        
      'case tipo de dato para formatear y alinear
      Select Case Me.dsiEsquema(varNombre).Type
          
      'bit - si esta vacio se devuelve: 0
      Case conBit
      
        'put resultado parcial
        strParcial = IIf(varValor = "", "null", varValor)
      
      'numeros - si esta vacio se devuelve: null
      Case conSmallInt, conInt, conTinyInt, conMoney, conSmallMoney, conReal, conFloat, conNumeric, conDecimal
          
        'put resultado parcial
        strParcial = IIf(IsEmpty(varValor), "null", varValor)
          
      'fecha - si esta vacio se devuelve: ''
      Case conSmallDateTime, conDateTime
      
        'put resultado parcial
'        strParcial = "'" & dateToIso(varValor) & "'"
        strParcial = IIf(varValor = "", "null", "'" & dateToIso(varValor) & "'")
      
      'string
      'si esta vacio y es texto comun se devuelve: null
      'si esta vacio y es ComboBox se devuelve: 0
      Case conChar, conNchar, conVarchar, conText, conNVarchar, conNtext
      
        'get valores combo
        strCombo = dsiConfigGetValor("lista", varNombre)
        strParam = dsiConfigGetParam("lista", varNombre)

        'ckeck si es un comboBox
        If strCombo <> "" Then
        
          'puntero de grilla en fila
          Me.spdE.row = intFila
          
          'puntero de grilla en columna en donde se encuentra el ID
          Me.spdE.col = intCol + 1
          
          'determino si ID a devolver es numerico o texto
          'set fuerzo puntero al ultimo elemento de comboBox
          Me.spdE.TypeComboBoxCurSel = Me.spdE.TypeComboBoxCount - 1
          
          'check si tipo de dato numerico, uso flags
          If IsNumeric(Me.spdE.Text) Then
            blnIDNumerico = True
          Else
            blnIDNumerico = False
          End If
          
          'puntero de grilla en columna en donde se encuentra el texto
          Me.spdE.col = intCol
          
          'check si no se selecciono ningun item, puede devuolver un 0 o null, segun
          If Me.spdE.TypeComboBoxCurSel = -1 Then
            
            'check comboBox editable
            If LCase(strParam) = "editable" Then
                
              'check si existe texto, devuelvo texto, sino null
              If Me.spdE.Text <> "" Then
                strParcial = "'" & Me.spdE.Text & "'"
              Else
                strParcial = "null"
              End If
              
            'comboBox no editable
            Else
                  
              'si columna obligatorio, devuelvo null, sino 0
              If strObliga <> "" Then
                strParcial = "null"
              Else
                strParcial = IIf(blnIDNumerico, "0", "''")
              End If
              
            End If
              
          Else
                  
            'get posicion del item seleccionado
            intComboInd = Me.spdE.TypeComboBoxCurSel
            
            'set puntero en columna en donde se encuentra el ID
            Me.spdE.col = intCol + 1
          
            'put posicion del item seleccionado
            Me.spdE.TypeComboBoxCurSel = intComboInd
            
            'get valor
            varValor = Me.spdE.Text
            
            'put resultado parcial
            strParcial = IIf(blnIDNumerico, varValor, "'" & varValor & "'")
            
          End If
      
        'si es un texto comun
        Else
          strParcial = IIf(varValor = "", "null", "'" & varValor & "'")
        End If
            
      End Select
      
      'armando el sql final
      strFinal = strFinal & strParcial & ","
        
    End If
            
    'fila proxima y columna proxima
    intFila = intFila + 1
    intEsq = intEsq + 1
        
  Wend
  
  'devuelvo resultado final, elimino la coma que esta al final del string
  dsiFormularioDatosGet = Left(strFinal, Len(strFinal) - 1)
  
End Property

Public Property Get dsiFormularioDatosINS() As Boolean
        
  'guardo en que modo esta el formulario antes de INS
  Me.dsiFormularioAnt = Me.dsiFormulario
        
  'si formulario en modo planilla, fuerzo y cambio a modo formulario
  If Me.dsiFormulario = "P" Then
    blnB = Me.dsiFormularioCambiar()
  End If
      
  'set limpio pantalla
  blnB = Me.dsiFormularioDatosClear()
  
  'set desbloqueo celdas
  blnB = Me.dsiFormularioDatosBloquear(False)
  
  'bloqueo barra herramientas
  blnB = Me.dsiHerramientas(False)
  blnB = Me.dsiHerramientasComfirmar(True)
    
  'set operacion
  Me.dsiOperacion = "INS"

End Property

Public Property Get dsiFormularioDatosEDI() As Boolean
    
  'guardo en que modo esta el formulario antes de INS
  Me.dsiFormularioAnt = Me.dsiFormulario
  
  'si formulario en modo planilla, fuerzo y cambio a modo formulario
  If Me.dsiFormulario = "P" Then
    blnB = Me.dsiFormularioCambiar()
  End If
   
  'set metodo config
  blnB = Me.dsiFormularioDatosBloquear(False)
  
  'bloqueo barra herramientas
  blnB = Me.dsiHerramientas(False)
  blnB = Me.dsiHerramientasComfirmar(True)
  
  'set operacion
  Me.dsiOperacion = "EDI"

End Property

Public Property Get dsiFormularioDatosELI() As Boolean
    
  'guardo en que modo esta el formulario antes de INS
  Me.dsiFormularioAnt = Me.dsiFormulario
  
  'si formulario en modo planilla, fuerzo y cambio a modo formulario
  If Me.dsiFormulario = "P" Then
    blnB = Me.dsiFormularioCambiar()
  End If
   
  'bloqueo barra herramientas
  blnB = Me.dsiHerramientas(False)
  blnB = Me.dsiHerramientasComfirmar(True)
  
  'set operacion
  Me.dsiOperacion = "ELI"
  
End Property

Public Property Get dsiFormularioDatosBloquear(blnB As Boolean) As Boolean
  Dim intCol, intFila, intEsq, intF As Integer
  Dim strValor As String
  Dim varColumna As Variant
      
  'set valores iniciales
  intCol = 1
  intFila = 1
  intEsq = 0
  
  'recorre esquema
  While intEsq <= Me.dsiEsquema.Count - 1
    
    'si fila = cantidad de filas de grilla, la pantalla es doble
    If intEsq = Me.spdE.MaxRows Then
      intCol = 4
      intFila = 1
    End If
    
    'puntero a fila y columna
    Me.spdE.row = intFila
    Me.spdE.col = intCol + 1
          
    'Lock o unLock celda
    spdE.Lock = blnB
    
    'check si hice Lock, set color Lock
    If blnB = True Then
      Me.spdE.ForeColor = RGB(155, 155, 155)
    Else  'unLock
           
      'set color unLock
      Me.spdE.ForeColor = RGB(0, 0, 0)
      
      'get nombre columna
      Me.spdE.GetText intCol, intFila, varColumna
      
      'get config columna
      strValor = dsiConfigGetValor("bloqueado", varColumna)
      
      'check si columna fue config como Lock, set lock y color
      If strValor <> "" Then
        Me.spdE.Lock = True
        Me.spdE.ForeColor = RGB(155, 155, 155)
      End If
      
    End If
    
    'fila proxima y columna proxima
    intFila = intFila + 1
    intEsq = intEsq + 1
        
  Wend
        
End Property

Public Property Get dsiFormularioDatosGuardar() As Boolean
  Dim strT, strSP, strSource, strIDnombre As String
  Dim blnB As Boolean
  Dim rs As ADODB.Recordset
  Dim intIDposicion As Integer
  Dim varIDvalor As Variant
      
  'get store procedure que se utilizara para registrar cambios
  strSP = Me.dsiConfigGetValor("sp" & Me.dsiOperacion)
        
  'ckeck si no existe store procedure, salgo
  If strSP = "" Then
    Exit Property
  End If
        
  'get info para save
  strT = Me.dsiFormularioDatosGet()
            
  'save info
  SQLexec ("exec " & strSP & " " & strT)
        
  'check errores
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    Exit Sub
  End If
        
  'get vista
  strSource = Me.dsiConfigGetValor("VISTA")
        
  'get nombre columna ID
  strIDnombre = Me.dsiConfigGetValor("CLAVE")
  
  'check sino encuentra el nombre de la columna ID, default ID
  If strIDnombre = "" Then
    strIDnombre = "id"
  End If
  
  'search fila en donde se encuentra ID
  intIDposicion = Me.spdE.SearchCol(1, 0, -1, strIDnombre, SearchFlagsNone)
  
  'get valor columna ID
  Me.spdE.GetText 2, intIDposicion, varIDvalor
  
  'put cambios segun operación
  Select Case Me.dsiOperacion
        
  Case "INS"
    
    'get record que acabo de insertar
    strT = "select * from " & strSource & " where " & strIDnombre & " = (select max(" & strIDnombre & ") from " & strSource & ")"
    Set rs = SQLexec(strT)
    
    'copy datos de recordset a spdE
    blnB = Me.dsiFormularioDatosPutRaF(rs)
    
    'delete puntero de fila anterior al insert
    blnB = Me.dsiFilaPintaNo(Me.spdGrid.ActiveCol, Me.spdGrid.ActiveRow)
    
    'insert fila en grilla planilla
    Me.spdGrid.MaxRows = Me.spdGrid.MaxRows + 1
    Me.spdGrid.InsertRows Me.spdGrid.ActiveRow, 1
    
    'set puntero a fila actual
    blnB = Me.dsiFilaPintaSi(Me.spdGrid.ActiveCol, Me.spdGrid.ActiveRow)
    
    'copy datos de spdE a spdGrid
    blnB = Me.dsiFormularioDatosPutFaP()
    
  Case "EDI"
    
    'get id utilizado por SQL
    strT = "select * from " & strSource & " where " & strIDnombre & " = " & varIDvalor
    Set rs = SQLexec(strT)
    
    'copy datos de recordset a spdE
    blnB = Me.dsiFormularioDatosPutRaF(rs)
    
    'put datos a spdGrid modo planilla
    blnB = Me.dsiFormularioDatosPutFaP()
    
  Case "ELI"
              
    'elimino fila de grilla spdGrid, pero como no encontre la forma hago lo siguiente
    'tengo que agregar 1 fila y luego borrar 2
    Me.spdGrid.MaxRows = Me.spdGrid.MaxRows + 1
    Me.spdGrid.InsertRows Me.spdGrid.ActiveRow, 1
    Me.spdGrid.DeleteRows Me.spdGrid.ActiveRow, 2
    Me.spdGrid.MaxRows = Me.spdGrid.MaxRows - 2
    
    'set pintar fila actual
    blnB = Me.dsiFilaPintaSi(Me.spdGrid.ActiveCol, Me.spdGrid.ActiveRow)
                
  End Select
        
  'cierro cn
  SQLclose
        
  'bloqueo celdas
  blnB = Me.dsiFormularioDatosBloquear(True)
  
  'si modo F antes de realizar operación, P fuerzo modo
  If Me.dsiFormularioAnt = "P" Then
    blnB = Me.dsiFormularioCambiar()
  End If
  
  'set operacion
  Me.dsiOperacion = "VER"
      
  'habilito barra de herramientas
  blnB = Me.dsiHerramientas(True)
  blnB = Me.dsiHerramientasComfirmar(False)
  
End Property

Public Property Get dsiFormularioDatosCancelar() As Boolean
    
  'bloqueo botones guardar, cancelar
  blnB = Me.dsiHerramientasComfirmar(False)
    
  'bloqueo celdas
  blnB = Me.dsiFormularioDatosBloquear(True)
  
  'recupero los datos anteriores, en una de esas se modificaron celdas y no se guardo
  blnB = Me.dsiFormularioDatosPutPaF()
  
  'si modo antes de realizar operación P fuerzo modo
  If Me.dsiFormularioAnt = "P" Then
    blnB = Me.dsiFormularioCambiar()
  End If
  
  'set operacion
  Me.dsiOperacion = "VER"
  
  'habilito barra de herramientas
  blnB = Me.dsiHerramientas(True)
  blnB = Me.dsiHerramientasComfirmar(False)
  
End Property


Public Property Get dsiFormularioDatosConfig() As Boolean
  
  Dim intCol, intFila, intEsq, intF As Integer
  Dim varDato As Variant
  Dim strDato, strValor, strParam, strT, strComboID, strComboTexto As String
  Dim rs As ADODB.Recordset
  Dim blnDefectoError As Boolean
  Dim arrComboBox, arrItem As Variant
  Dim blnB As Boolean
    
  'si formulario en modo de planilla, set modo de un registro x vez
'  If Me.dsiFormulario = "P" Then
'    blnB = Me.dsiFormularioCambiar()
'  End If

'check si ya esta en modo formulario, salgo.
'cambio: 16/05/2005
  If Me.dsiFormulario = "F" Then
    Exit Property
  End If
    
  'set modifica grilla en BackGround
  Me.spdE.ReDraw = False
  
  'set valores iniciales
  intCol = 1
  intFila = 1
  intEsq = 0
  
  'recorre esquema
  While intEsq <= Me.dsiEsquema.Count - 1
    
    'si fila = cantidad de filas de grilla, la pantalla es doble
    If intEsq = Me.spdE.MaxRows Then
      intCol = 4
      intFila = 1
    End If
    
    'get titulo columna
    Me.spdE.GetText intCol, intFila, varDato
        
    'puntero a fila y columna
    Me.spdE.row = intFila
    Me.spdE.col = intCol + 1
          
    'default config sin errores
    blnDefectoError = False
        
    'convierto a string
    strDato = CStr(varDato)
    
    'configuro celda LISTA-------------------------------------------------------------
    strValor = dsiConfigGetValor("lista", strDato)
    strParam = dsiConfigGetParam("lista", strDato)
    
    'si existe lista
    If strValor <> "" Then
            
      'set puntero celda para id
      Me.spdE.col = intCol + 2
      
      'set celda de tipo combo box
      spdE.CellType = CellTypeComboBox
                 
      'set puntero celda para texto
      Me.spdE.col = intCol + 1
                
      'set celda de tipo combo box
      spdE.CellType = CellTypeComboBox
                
      'set comboBox editable, si en parametro dice EDITABLE
      If LCase(strParam) = "editable" Then
        spdE.TypeComboBoxEditable = True
      End If
      spdE.TypeComboBoxhWnd = 4
          
      'set busqueda avanzada
      spdE.TypeComboBoxAutoSearch = TypeComboBoxAutoSearchMultipleChar
         
      'set muestra 20 lineas
      spdE.TypeComboBoxMaxDrop = 20
              
      'set ajuste de comboBox al dato mas ancho
      spdE.TypeComboBoxWidth = -1
            
      'clear datos para combo
      strComboID = ""
      strComboTexto = ""
      
      'si contiene un select ejecuto y lleno comboBox
      If InStr(LCase(strValor), "select") Then
        
        'ejecuto select
        Set rs = SQLexec(strValor)
          
        'si no hay errores, lleno comboBox
        If SQLparam.CnErrNumero = -1 Then
            
          'change puntero mouse
          Screen.MousePointer = vbHourglass
          
          'recorro rs
          While Not rs.EOF
              
            'si cantidad de columnas es 1
            'columna1=ID de combo, columna2=texto de combo
            If rs.Fields.Count = 1 Then
                          
              strComboID = strComboID & rs(0) & Chr(9)
              strComboTexto = strComboTexto & rs(0) & Chr(9)
            
            'si cantidad de columnas es 2
            'columna1=ID de combo, columna2=texto de combo
            Else
              
              strComboID = strComboID & rs(0) & Chr(9)
              strComboTexto = strComboTexto & rs(1) & Chr(9)
            
            End If
            
            'puntero al proximo
            rs.MoveNext
            
          Wend
                        
          'change puntero mouse
          Screen.MousePointer = vbDefault
              
        Else
          
          blnDefectoError = True                  'marca de select con error
          
          'set vacio comboBox para texto
          spdE.TypeComboBoxList = ""
          
          'set vacio comboBox para id
          Me.spdE.col = intCol + 2
          spdE.TypeComboBoxList = ""
          
          'puntero a columna para dato
          Me.spdE.col = intCol + 1
          
        End If
          
      'esto es para valores de comboBox a mano como por ejemplo:
      'enero;enero o enero,1;febrero,2 o enero,ene;febrero,feb
      Else
        
        'separo los valores para llenar comboBox
        arrComboBox = separateText(strValor, ";")
        
        'recorro array con valores para comboBox
        For intF = 1 To UBound(arrComboBox, 1)
          
          'separo item para saber si tiene id y texto
          arrItem = separateText(arrComboBox(intF), ",")
          
          'si cantidad de columnas es 1
          'columna1=ID de combo, columna1=texto de combo
          If UBound(arrItem) = 1 Then
                          
            strComboID = strComboID & arrItem(1) & Chr(9)
            strComboTexto = strComboTexto & arrItem(1) & Chr(9)
            
          'si cantidad de columnas es 2
          'columna1=ID de combo, columna 2=texto de combo
          Else
             
            strComboID = strComboID & arrItem(1) & Chr(9)
            strComboTexto = strComboTexto & arrItem(2) & Chr(9)
            
          End If
        
        Next
        
      End If
      
      'set lleno comboBox con texto
      spdE.TypeComboBoxList = strComboTexto
          
      'set lleno comboBox con id
      Me.spdE.col = intCol + 2
      spdE.TypeComboBoxList = strComboID
            
      'puntero a columna para dato
      Me.spdE.col = intCol + 1
      
    End If
    
    'configuro celda OBLIGATORIA----------------------------------------------------------
    strValor = dsiConfigGetValor("obligatorio", strDato)
    
    If strValor <> "" Then
      spdE.BackColor = RGB(216, 237, 223)
    End If
            
    'configuro celda BLOQUEADA------------------------------------------------------------
    strValor = dsiConfigGetValor("bloqueado", strDato)
    
    If strValor <> "" Then
      Me.spdE.ForeColor = RGB(155, 155, 155)
    End If
    
'cambio: 16/03/2005
'      spdE.Lock = True
'    Else
'      spdE.Lock = False
    
    
    'bloqueo celda
    'cambio: 16/05/2005
    spdE.Lock = True
            
    'fila proxima y columna proxima
    intFila = intFila + 1
    intEsq = intEsq + 1
        
  Wend
        
  'set modifica grilla en ForeGround
  Me.spdE.ReDraw = True
  
End Property

Public Property Get dsiFormularioDatosClear() As Boolean
  
  Dim intCol, intFila, intEsq As Integer
  Dim varColumna As Variant
  Dim strColumna, strValor As String
    
  'set modifica grilla en BackGround
  Me.spdE.ReDraw = False
  
  'set valores iniciales
  intCol = 1
  intFila = 1
  intEsq = 0
  
  'recorre esquema
  While intEsq <= Me.dsiEsquema.Count - 1
    
    'si fila = cantidad de filas de grilla, la pantalla es doble
    If intEsq = Me.spdE.MaxRows Then
      intCol = 4
      intFila = 1
    End If
    
    'get titulo columna
    Me.spdE.GetText intCol, intFila, varColumna
        
    'puntero a fila y columna
    Me.spdE.row = intFila
    Me.spdE.col = intCol + 1
          
    'convierto a string
    strColumna = CStr(varColumna)
          
    'configuro valor MANTENER-------------------------------------------------------------
    strValor = dsiConfigGetValor("mantener", strColumna)
        
    'check si celda mantiene valor
    If strValor = "" Then
      Me.spdE.Text = ""
    End If
            
    'configuro valor DEFECTO--------------------------------------------------------------
    strValor = dsiConfigGetValor("defecto", strColumna)
    
    'existe defecto y operacion INS
    If strValor <> "" Then
      
      'si contiene un select ejecuto y get valor, sino muestro valor directamente
      If InStr(LCase(strValor), "select") Then
        
        'ejecuto select
        Set rs = SQLexec(strValor)
          
        'si no hay errores, muestro valor, sino pinto celda en rojo para indicar error
        If SQLparam.CnErrNumero = -1 Then
          Me.spdE.Text = rs(0)
        Else
          Me.spdE.BackColor = RGB(255, 128, 128)
        End If
          
      Else
        Me.spdE.Text = strValor
      End If
      
    End If
            
    'fila proxima y columna proxima
    intFila = intFila + 1
    intEsq = intEsq + 1
        
  Wend
        
  'set modifica grilla en ForeGround
  Me.spdE.ReDraw = True
  
End Property

Public Property Let dsiGrillaIniFila(strT As String)
  m_GrillaIniFila = strT
End Property

Public Property Get dsiGrillaIniFila() As String
  dsiGrillaIniFila = m_GrillaIniFila
End Property

Public Property Let dsiGrillaIniCol(strT As String)
  m_GrillaIniCol = strT
End Property

Public Property Get dsiGrillaIniCol() As String
  dsiGrillaIniCol = m_GrillaIniCol
End Property

Public Property Get dsiUbicacionGrillaNo()
    
  'limpio ubicacion grilla de barra de estado
  MainMDI.staBarra1.Panels(2) = ""
  MainMDI.staBarra1.Panels(2).ToolTipText = ""
  
End Property

Public Property Get dsiSumAvgMinMaxNo()
    
  'clear barraStado
  MainMDI.staBarra1.Panels(3) = ""
  MainMDI.staBarra1.Panels(3).ToolTipText = ""
  
End Property

Public Property Get dsiSumAvgMinMaxSi()
  Dim blnB As Boolean
      
  blnB = show_SumAvgMinMax()
  
End Property


Public Property Get dsiUbicacionGrillaSi()
    
  'set puntero interno en fila columna
  Me.spdGrid.row = Me.spdGrid.ActiveRow
  Me.spdGrid.col = Me.spdGrid.ActiveCol
  
  'muestra en barra de estado de frm MDI
  MainMDI.staBarra1.Panels(2) = " fila: " & Me.spdGrid.ActiveRow & " de " & Me.spdGrid.MaxRows & " , columna: " & Me.spdGrid.ActiveCol & " de " & Me.spdGrid.MaxCols & " , valor: " & Me.spdGrid.Text & " "
  MainMDI.staBarra1.Panels(2).ToolTipText = " fila: " & Me.spdGrid.ActiveRow & " de " & Me.spdGrid.MaxRows & " , columna: " & Me.spdGrid.ActiveCol & " de " & Me.spdGrid.MaxCols & " , valor: " & Me.spdGrid.Text & " "
  
End Property

Public Property Get dsiFilaPri() As Boolean
  
  Dim blnB As Boolean
  
  'chequeo si hay filas y que el puntero no este en la primera
  If Me.spdGrid.MaxRows > 0 And Me.spdGrid.ActiveRow <> 1 Then
        
    blnB = dsiFilaPintaNo(Me.spdGrid.ActiveCol, Me.spdGrid.ActiveRow)
    Me.spdGrid.SetActiveCell Me.spdGrid.ActiveCol, 1
    blnB = Me.dsiUbicacionGrillaSi()
    blnB = dsiFilaPintaSi(Me.spdGrid.ActiveCol, Me.spdGrid.ActiveRow)
        
  End If
  
  'set foco
  If Me.spdE.Visible Then
    blnB = Me.dsiFormularioDatosPutPaF()
  End If
  
End Property

Public Property Get dsiFilaSig() As Boolean
  
  Dim blnB As Boolean
  
  'chequeo si hay filas y que el puntero no este en la ultima
  If Me.spdGrid.MaxRows > 0 And Me.spdGrid.ActiveRow <> Me.spdGrid.MaxRows Then
    
    blnB = dsiFilaPintaNo(Me.spdGrid.ActiveCol, Me.spdGrid.ActiveRow)
    Me.spdGrid.SetActiveCell Me.spdGrid.ActiveCol, Me.spdGrid.ActiveRow + 1
    blnB = Me.dsiUbicacionGrillaSi()
    blnB = dsiFilaPintaSi(Me.spdGrid.ActiveCol, Me.spdGrid.ActiveRow)
  
  End If

  'set foco
  If Me.spdE.Visible Then
    blnB = Me.dsiFormularioDatosPutPaF()
  End If

End Property

Public Property Get dsiFilaAnt() As Boolean

  Dim blnB As Boolean
  
  'chequeo si hay filas y que el puntero no este en la primera
  If Me.spdGrid.MaxRows > 0 And Me.spdGrid.ActiveRow <> 1 Then
    
    blnB = dsiFilaPintaNo(Me.spdGrid.ActiveCol, Me.spdGrid.ActiveRow)
    Me.spdGrid.SetActiveCell Me.spdGrid.ActiveCol, Me.spdGrid.ActiveRow - 1
    blnB = Me.dsiUbicacionGrillaSi()
    blnB = dsiFilaPintaSi(Me.spdGrid.ActiveCol, Me.spdGrid.ActiveRow)
  
  End If

  'set foco
  If Me.spdE.Visible Then
    blnB = Me.dsiFormularioDatosPutPaF()
  End If

End Property

Public Property Get dsiFilaUlt() As Boolean

  Dim blnB As Boolean
  
  'chequeo si hay filas y que el puntero no este en la ultima
  If Me.spdGrid.MaxRows > 0 And Me.spdGrid.ActiveRow <> Me.spdGrid.MaxRows Then
    
    blnB = dsiFilaPintaNo(Me.spdGrid.ActiveCol, Me.spdGrid.ActiveRow)
    Me.spdGrid.SetActiveCell Me.spdGrid.ActiveCol, Me.spdGrid.MaxRows
    blnB = Me.dsiUbicacionGrillaSi()
    blnB = dsiFilaPintaSi(Me.spdGrid.ActiveCol, Me.spdGrid.ActiveRow)
  
  End If

  'set foco
  If Me.spdE.Visible Then
    blnB = Me.dsiFormularioDatosPutPaF()
  End If

End Property

Public Property Get dsiFilaPintaNo(lngCol, lngRow) As Boolean

  'set color default de grilla a celda para abandonar
  Me.spdGrid.BlockMode = True
  Me.spdGrid.row = lngRow
  Me.spdGrid.col = lngCol
  Me.spdGrid.col2 = -1
  Me.spdGrid.BackColor = RGB(245, 245, 245)
  Me.spdGrid.ForeColor = RGB(60, 60, 60)
  
End Property

Public Property Get dsiFilaPintaSi(lngCol, lngRow) As Boolean

  'set color de seleccion a celda nueva
  Me.spdGrid.BlockMode = True
  Me.spdGrid.row = lngRow
  Me.spdGrid.row2 = lngRow
  Me.spdGrid.col2 = 1
  Me.spdGrid.col2 = -1
  Me.spdGrid.BackColor = RGB(220, 220, 220)
  Me.spdGrid.ForeColor = RGB(60, 60, 60)
  Me.spdGrid.BlockMode = False
    
End Property

Public Property Get dsiCopiarSeleccion() As Boolean
  
  Me.ActiveControl.ClipboardCopy
  
End Property

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
End Property

Public Property Get dsiTitulo() As String
  dsiTitulo = m_Titulo
End Property

Public Property Set dsiEsquema(ByVal fld As ADODB.Fields)
  Set m_Esquema = fld
End Property

Public Property Get dsiEsquema() As ADODB.Fields
  Set dsiEsquema = m_Esquema
End Property

Public Property Let dsiSpIMPpermiso(ByVal bln As Boolean)
  m_spIMPpermiso = bln
End Property

Public Property Get dsiSpIMPpermiso() As Boolean
  dsiSpIMPpermiso = m_spIMPacceso
End Property

Public Property Let dsiSpINSpermiso(ByVal bln As Boolean)
  m_spINSpermiso = bln
End Property

Public Property Get dsiSpINSpermiso() As Boolean
  dsiSpINSpermiso = m_spINSacceso
End Property

Public Property Let dsiSpEDIpermiso(ByVal bln As Boolean)
  m_spEDIpermiso = bln
End Property

Public Property Get dsiSpEDIpermiso() As Boolean
  dsiSpEDIpermiso = m_spEDIpermiso
End Property

Public Property Let dsiSpELIpermiso(ByVal bln As Boolean)
  m_spELIpermiso = bln
End Property

Public Property Get dsiSpELIpermiso() As Boolean
  dsiSpELIpermiso = m_spELIpermiso
End Property

Public Property Let dsiWhere(ByVal str As String)
  m_Where = str
End Property

Public Property Get dsiWhere() As String
  dsiWhere = m_Where
End Property

Public Property Get dsiWhereMostrarSi() As Boolean

  'muestra en barra de estado de frm MDI
  MainMDI.staBarra2.Panels(1) = " General: " & MainMDI.dsiWhereGen("nombre") & ", Grilla: " & Me.dsiWhere & " "
  MainMDI.staBarra2.Panels(1).ToolTipText = " General: " & MainMDI.dsiWhereGen("nombre") & ", Grilla: " & Me.dsiWhere & " "

End Property

Public Property Get dsiWhereMostrarNo() As Boolean

  'muestra en barra de estado de frm MDI
  MainMDI.staBarra2.Panels(1) = ""
  MainMDI.staBarra2.Panels(1).ToolTipText = ""

End Property

Public Property Let dsiOrderBy(ByVal str As String)
  m_OrderBy = str
End Property

Public Property Get dsiOrderByMostrarSi() As Boolean

  'muestra en barra de estado de frm MDI
  MainMDI.staBarra1.Panels(1) = " orden: " & Me.dsiOrderBy & " "
  MainMDI.staBarra1.Panels(1).ToolTipText = " orden: " & Me.dsiOrderBy & " "

End Property

Public Property Get dsiOrderByMostrarNo() As Boolean

  'limpio barra de estado
  MainMDI.staBarra1.Panels(1) = ""
  MainMDI.staBarra1.Panels(1).ToolTipText = ""

End Property

Public Property Get dsiOrderBy() As String
  dsiOrderBy = m_OrderBy
End Property

Public Property Let dsiConfigAgregar(arr As Variant)
  
  m_ConfigCant = m_ConfigCant + 1
  ReDim Preserve m_Config(4, m_ConfigCant)
  
  m_Config(0, m_ConfigCant) = IIf(IsNull(arr(0)), "", arr(0))
  m_Config(1, m_ConfigCant) = IIf(IsNull(arr(1)), "", arr(1))
  m_Config(2, m_ConfigCant) = IIf(IsNull(arr(2)), "", arr(2))
  m_Config(3, m_ConfigCant) = IIf(IsNull(arr(3)), "", arr(3))
  
End Property

Public Property Get dsiConfigGet() As Variant
  dsiConfigGet = m_Config
End Property

Public Property Get dsiConfigGetFila(intI As Integer) As Variant
  dsiConfigGetFila = Array(m_Config(0, intI), m_Config(1, intI), m_Config(2, intI), m_Config(3, intI))
End Property

Public Property Get dsiConfigGetValor(strTipo As String, Optional ByVal strColumna As String) As String
  
  Dim intI As Integer
  
  'valor default que devuelve
  dsiConfigGetValor = ""
  
  'si parametro null, fuerzo ""
  If IsEmpty(strColumna) Then
    strColumna = ""
  End If
  
  'recorro config
  For intI = 0 To Me.dsiConfigCant
        
    'cuando encuentro devuelvo
    If LCase(strTipo) = LCase(m_Config(0, intI)) And LCase(strColumna) = LCase(m_Config(1, intI)) Then
      dsiConfigGetValor = m_Config(2, intI)
      Exit For
    End If
      
  Next

End Property

Public Property Get dsiConfigGetParam(strTipo As String, Optional ByVal strColumna As String) As String
  
  Dim intI As Integer
  
  'valor default que devuelve
  dsiConfigGetParam = ""
  
  'si parametro null, fuerzo ""
  If IsEmpty(strColumna) Then
    strColumna = ""
  End If
  
  'recorro config
  For intI = 0 To Me.dsiConfigCant
        
    'cuando encuentro devuelvo
    If LCase(strTipo) = LCase(m_Config(0, intI)) And LCase(strColumna) = LCase(m_Config(1, intI)) Then
      dsiConfigGetParam = m_Config(3, intI)
      Exit For
    End If
      
  Next

End Property

Public Property Get dsiConfigCant() As Integer
  dsiConfigCant = m_ConfigCant
End Property

Public Property Let dsiCondicionAgregar(str As String)
  
  Dim intI As Integer
  
  'agrego un elemento al array y guardo nueva condicion
  m_CondicionCant = m_CondicionCant + 1
  
  ReDim Preserve m_Condicion(m_CondicionCant)
  m_Condicion(m_CondicionCant) = str

  'guardo condicion completa en propiedad where
  str = ""
  For intI = 1 To m_CondicionCant
    str = str & m_Condicion(intI) & " and "
  Next
     
  'elimino el ultimo end
  If m_CondicionCant > 0 Then
    str = Left(str, Len(str) - 5)
  End If
  
  Me.dsiWhere = str

End Property

Public Property Get dsiCondicionAtras() As Boolean
    
  Dim strT As String
  Dim intI As Integer
    
  'chequeo si existen condiciones
  If m_CondicionCant > 0 Then
    
    'elimino elemento del array en donde se guardan las condiciones
    m_CondicionCant = m_CondicionCant - 1
    ReDim Preserve m_Condicion(m_CondicionCant)
    
    'guardo condicion completa en propiedad where
    strT = ""
    For intI = 1 To Me.dsiCondicionCant
      strT = strT & m_Condicion(intI) & " and "
    Next
      
    'elimino el ultimo end
    If m_CondicionCant > 0 Then
      strT = Left(strT, Len(strT) - 5)
    End If
    
    Me.dsiWhere = strT
    
    'refresh
    blnB = Me.dsiRefresh()
        
    'muestra en barra de estado
    blnB = Me.dsiWhereMostrarSi()
  
  End If
    
  'set foco
  If Me.spdE.Visible Then
    blnB = Me.dsiFormularioDatosPutPaF()
  End If
    
End Property

Public Property Get dsiCondicionEliminar() As Boolean
    
  Dim strT As String
  Dim intI As Integer
    
  'chequeo si existen condiciones
  If m_CondicionCant > 0 Then
    
    'elimino elemento del array en donde se guardan las condiciones
    m_CondicionCant = 0
    ReDim Preserve m_Condicion(m_CondicionCant)
    
    'guardo condicion completa en propiedad where
    strT = ""
      
    Me.dsiWhere = strT
    
    'refresh
    blnB = Me.dsiRefresh()
    
    'muestra en barra de estado
    blnB = Me.dsiWhereMostrarSi()
    
  End If
    
  'set foco
  If Me.spdE.Visible Then
    blnB = Me.dsiFormularioDatosPutPaF()
  End If
    
End Property

Public Property Get dsiCondicionCant() As Integer
  dsiCondicionCant = m_CondicionCant
End Property

Public Property Let dsiCondicion(arr As Variant)
  m_Condicion = arr
  
  Dim intI As Integer
  
  'si existen condiciones las agrego
  If UBound(arr) <> -1 Then
        
    For intI = 1 To UBound(arr)
      Me.dsiCondicionAgregar = arr(intI)
    Next
    
  End If
  
End Property

Public Property Get dsiCondicionGet(intIDcolumna As Integer) As String
  
  'default
  dsiCondicionGet = ""
  
  'si array vacio exit
  If UBound(m_Condicion) = -1 Then
    Exit Property
  End If
  
  'si indice mayor que maximo exit
  If intIDcolumna > UBound(m_Condicion) Then
    Exit Property
  End If
   
  'devuelvo cantidad decimales
  dsiCondicionGet = m_Condicion(intIDcolumna)

End Property

Public Property Get dsiCondicionRapido() As Boolean
  
  'declaro variables
  Dim strColName As Variant
  Dim varValue As Variant
  Dim strWhere As String
  Dim lngL As Long
  Dim blnB As Boolean
    
  'toma nombre de columna y valor de celda seleccionada
  lngL = Me.spdGrid.GetText(Me.spdGrid.ActiveCol, 0, strColName)
  lngL = Me.spdGrid.GetText(Me.spdGrid.ActiveCol, Me.spdGrid.ActiveRow, varValue)
    
  'set row y col para conocer tipo
  Me.spdGrid.row = Me.spdGrid.ActiveRow
  Me.spdGrid.col = Me.spdGrid.ActiveCol
    
  'verifico tipo de dato
  Select Case Me.spdGrid.CellType
  
  Case CellTypeCheckBox, CellTypeNumber, CellTypeCurrency
    
    If IsEmpty(varValue) Then
      strWhere = " is null"
    Else
      strWhere = " = " & varValue
    End If
  
  Case CellTypeDate
    strWhere = " = '" & dateToIso(varValue) & "'"
  
  Case CellTypeEdit
    strWhere = " like '" & varValue & "'"
  
  End Select
    
  'agrego ultimo filtro a propiedad dataFiltro
  Me.dsiCondicionAgregar = "[" & strColName & "]" & strWhere

  'refresh
  blnB = Me.dsiRefresh()

  'muestro en barra de estado
  blnB = Me.dsiWhereMostrarSi()

  'set foco
  If Me.spdE.Visible Then
    blnB = Me.dsiFormularioDatosPutPaF()
  End If

End Property

Public Property Get dsiCondicionGuardar() As Boolean
    
  Dim blnB As Boolean
  Dim strT As String
  Dim intI As Integer
    
  'set filtro para que solo permita abrir los filtros de grilla actual
  strT = "Archivos de texto (" & Me.dsiIDopc & "*.sql)|" & Me.dsiIDopc & "*.sql"
  Me.comDestino.Filter = strT
    
  'titulo de ventana
  Me.comDestino.DialogTitle = "Guardar condición..."
  
  'set nombre de filtro, cambio 01/07/2005
'  strT = Me.dsiIDopc & " - " & Replace(Me.dsiWhere, "'", "") & ".sql"
  Me.comDestino.FileName = ""
    
  'abro cuadro de dialogo
  Me.comDestino.ShowSave
      
  'si cancelar salgo
  If Me.comDestino.FileName = "" Then
    Exit Sub
  End If
      
  'add IDopcion para poder identificar grilla
  intI = InStr(Me.comDestino.FileName, Me.comDestino.FileTitle)                                 'separo path de nombre de filtro
  strT = Left(Me.comDestino.FileName, intI - 1) & Me.dsiIDopc & " - " & Me.comDestino.FileTitle 'armo nombre path + IDopc + nombre
    
  'guardo condicion
  blnB = WriteIni("condicion", Me.dsiIDopc, Me.dsiWhere, strT)
  
  'set foco
  If Me.spdE.Visible Then
    blnB = Me.dsiFormularioDatosPutPaF()
  End If
  
End Property

Public Property Get dsiCondicionAbrir() As Boolean
    
  Dim blnB As Boolean
  Dim strT As String
      
  'set filtro para que solo permita abrir los filtros de grilla actual
  strT = "Archivos de texto (" & Me.dsiIDopc & "*.sql)|" & Me.dsiIDopc & "*.sql"
  Me.comDestino.Filter = strT
    
  'set titulo de ventana
  Me.comDestino.DialogTitle = "Abrir condición..."
  
  'set tiene que seleccionar un nombre
  Me.comDestino.FileName = ""
  
  'abro cuadro de dialogo
  Me.comDestino.ShowOpen
      
  'si cancelar salgo
  If Me.comDestino.FileName = "" Then
    Exit Property
  End If
      
  'get condicion
  strT = ReadIni("condicion", Me.dsiIDopc, Me.comDestino.FileName)
  
  'si existen condiciones, las borro
  If m_CondicionCant > 0 Then
    
    'elimino elemento del array en donde se guardan las condiciones
    m_CondicionCant = 0
    ReDim Preserve m_Condicion(m_CondicionCant)
    
  End If
      
  'agrego condicion
  Me.dsiCondicionAgregar = strT
      
  'refresh
  blnB = Me.dsiRefresh()
  
  'muestra en barra de estado
  blnB = Me.dsiWhereMostrarSi()
    
  'set foco
  If Me.spdE.Visible Then
    blnB = Me.dsiFormularioDatosPutPaF()
  End If
  
End Property

Public Property Let dsiColumnasCongeladas(ByVal str As String)
  m_ColumnasCongeladas = str
End Property

Public Property Get dsiColumnasCongeladas() As String
  dsiColumnasCongeladas = m_ColumnasCongeladas
End Property

Public Property Let dsiColumnasDecimales(arr As Variant)
  m_ColumnasDecimales = arr
End Property

Public Property Get dsiColumnasDecimalesGet(intIDcol As Long) As Integer
    
  Dim intI As Integer
    
  'default
  dsiColumnasDecimalesGet = -1
  
  'si array vacio exit
  If UBound(m_ColumnasDecimales) = -1 Then
    Exit Property
  End If
  
  intI = 1
  For intI = 1 To UBound(m_ColumnasDecimales) - 1 Step 2
    If m_ColumnasDecimales(intI) = intIDcol Then
      dsiColumnasDecimalesGet = m_ColumnasDecimales(intI + 1)
      Exit For
    End If
  Next
    
End Property

Public Property Get dsiColumnasDecimalesMenos() As Boolean
  
  Dim col, row, col2, row2 As Variant
  Dim lngL As Long
    
  'si modo formulario no se puede ajustar ancho
  If Me.dsiFormulario = "F" Then
    Exit Property
  End If
    
  'cambios en backGround
  Me.spdGrid.ReDraw = False
  
  'set propiedades en bloque habilitado
  Me.spdGrid.BlockMode = True
  
  'tomo numeros de filas y columnas seleccionadas
  Me.spdGrid.GetSelection 1, col, row, col2, row2
  
  'los guarda en las propiedades para cambiar en bloque
  Me.spdGrid.col = col
  Me.spdGrid.row = 1
  Me.spdGrid.col2 = col2
  Me.spdGrid.row2 = -1
  
  'recorro columnas seleccionadas
  For lngL = col To col2
    
    Me.spdGrid.col = lngL
    Me.spdGrid.col2 = lngL
    
    'aplico si por lo menos tiene 1 decimal
    If Me.spdGrid.TypeNumberDecPlaces >= 1 Then
    
      'agrego decimales
      Me.spdGrid.TypeNumberDecPlaces = Me.spdGrid.TypeNumberDecPlaces - 1
    
      'averiguo el valor mas ancho de la columna y lo aplico, le sumo un 10% mas
      dblAncho = spdGrid.MaxTextColWidth(lngL)
      spdGrid.ColWidth(lngL) = dblAncho + (dblAncho * 0.1)
  
    End If
  
  Next

  'cambios en foreGround
  Me.spdGrid.ReDraw = False

End Property

Public Property Get dsiColumnasDecimalesMas() As Boolean
  
  Dim col, row, col2, row2 As Variant
  Dim lngL As Long
    
  'si modo formulario no se puede ajustar ancho
  If Me.dsiFormulario = "F" Then
    Exit Property
  End If
    
  'cambios en backGround
  Me.spdGrid.ReDraw = False
  
  'set propiedades en bloque habilitado
  Me.spdGrid.BlockMode = True
  
  'tomo numeros de filas y columnas seleccionadas
  Me.spdGrid.GetSelection 1, col, row, col2, row2
  
  'los guarda en las propiedades para cambiar en bloque
  Me.spdGrid.col = col
  Me.spdGrid.row = row
  Me.spdGrid.col2 = col2
  Me.spdGrid.row2 = -1
  
  Dim intCol As Integer
  
  'recorro columnas seleccionadas
  For lngL = col To col2
    
    Me.spdGrid.col = lngL
    Me.spdGrid.col2 = lngL
    
    'limite 10 decimales
    If Me.spdGrid.TypeNumberDecPlaces < 10 Then
    
      'agrego decimales
      Me.spdGrid.TypeNumberDecPlaces = Me.spdGrid.TypeNumberDecPlaces + 1
    
      'averiguo el valor mas ancho de la columna y lo aplico, le sumo un 10% mas
      dblAncho = spdGrid.MaxTextColWidth(lngL)
      spdGrid.ColWidth(lngL) = dblAncho + (dblAncho * 0.1)
  
    End If
  
  Next
  
  'set de propiedades en bloque deshabilitado
  Me.spdGrid.BlockMode = False

  'cambios en foreGround
  Me.spdGrid.ReDraw = False

End Property

Public Property Let dsiColumnasAncho(arr As Variant)
  m_ColumnasAncho = arr
End Property

Public Property Get dsiColumnasAnchoGet(intIDcol As Long) As Single
  
  Dim intI As Integer
    
  'default
  dsiColumnasAnchoGet = -1
  
  'si array vacio exit
  If UBound(m_ColumnasAncho) = -1 Then
    Exit Property
  End If
  
  intI = 1
  For intI = 1 To UBound(m_ColumnasAncho) - 1 Step 2
    If m_ColumnasAncho(intI) = intIDcol Then
      dsiColumnasAnchoGet = m_ColumnasAncho(intI + 1)
      Exit For
    End If
  Next
  
End Property

Public Property Get dsiColumnasAnchoAjustar() As Boolean
  
  Dim col, row, col2, row2 As Variant
  Dim lngL As Long
  Dim SPD As fpSpread
    
  'set referencia a grilla activa
  If Me.spdGrid.Visible = True Then
    Set SPD = Me.spdGrid
  Else
    Set SPD = Me.spdE
  End If
  
  'si modo formulario no se puede ajustar ancho
  If Me.dsiFormulario = "F" Then
    Exit Property
  End If
    
  'cambios en backGround
  SPD.ReDraw = False
  
  'posibilidad de aplicar config en bloque
  SPD.BlockMode = True
  
  'get seleccion
  SPD.GetSelection 1, col, row, col2, row2
  
  'los guarda en las propiedades para cambiar en bloque
  col = IIf(col = -1, 1, col)
  col2 = IIf(col2 = -1, SPD.MaxCols, col2)
  
  'recorro columnas seleccionadas
  For lngL = col To col2
    
    SPD.col = lngL
    SPD.col2 = lngL
    
    'averiguo el valor mas ancho de la columna y lo aplico, le sumo un 10% mas
    dblAncho = SPD.MaxTextColWidth(lngL)
    SPD.ColWidth(lngL) = dblAncho + (dblAncho * 0.1)
  
  Next
  
  'set de propiedades en bloque deshabilitado
  SPD.BlockMode = False
  
  'cambios en backGround
  SPD.ReDraw = True
    
End Property

Public Property Let dsiColumnasOcultas(arr As Variant)
  m_ColumnasOcultas = arr
End Property

Public Property Get dsiColumnasOcultasGet(intIDcol As Long) As Boolean
  
  Dim intI As Integer
    
  'default
  dsiColumnasOcultasGet = False
  
  'si array vacio exit
  If UBound(m_ColumnasOcultas) = -1 Then
    Exit Property
  End If
  
  intI = 1
  For intI = 1 To UBound(m_ColumnasOcultas)
    If m_ColumnasOcultas(intI) = intIDcol Then
      dsiColumnasOcultasGet = True
      Exit For
    End If
  Next

End Property

Public Property Get dsiColumnasOcultasOcultar() As Boolean
  
  Dim col, row, col2, row2 As Variant
  Dim lngL As Long
  
  'si modo formulario no se puede ajustar ancho
  If Me.dsiFormulario = "F" Then
    Exit Property
  End If
  
  'cambios en backGround
  Me.spdGrid.ReDraw = False
  
  'posibilidad de aplicar config en bloque
  Me.spdGrid.BlockMode = True
  
  'get seleccion
  Me.spdGrid.GetSelection 1, col, row, col2, row2
  
  'los guarda en las propiedades para cambiar en bloque
  col = IIf(col = -1, 1, col)
  col2 = IIf(col2 = -1, Me.spdGrid.MaxCols, col2)
  
  'recorro columnas seleccionadas
  For lngL = col To col2
    
    Me.spdGrid.col = lngL
    Me.spdGrid.col2 = lngL
    
    'oculto columna
    Me.spdGrid.col = lngL
    Me.spdGrid.ColHidden = True
  
  Next
  
  'set de propiedades en bloque deshabilitado
  Me.spdGrid.BlockMode = False
  
  'cambios en backGround
  Me.spdGrid.ReDraw = True
  
End Property


Public Property Get dsiColumnasOcultasMostrar() As Boolean
  
  Dim lngL As Long
  
  'si modo formulario no se puede ajustar ancho
  If Me.dsiFormulario = "F" Then
    Exit Property
  End If
  
  'cambios en backGround
  Me.spdGrid.ReDraw = False
    
  'si hay columnas ocultas permito volver a verlas
  For lngL = 1 To Me.spdGrid.MaxCols
    
    'puntero a columna
    Me.spdGrid.col = lngL
    
    'si columna oculta la vuelvo a mostrar
    If Me.spdGrid.ColHidden = True Then
    
      Me.spdGrid.ColHidden = False
      
      'averiguo el valor mas ancho de la columna y lo aplico, le sumo un 10% mas
      dblAncho = spdGrid.MaxTextColWidth(lngL)
      spdGrid.ColWidth(lngL) = dblAncho + (dblAncho * 0.1)
      
    End If
  
  Next
  
  'cambios en foreGround
  Me.spdGrid.ReDraw = True
  
End Property

Public Property Get dsiColumnasCongelar() As Boolean
  
  Dim lngFilaAnt, lngColAnt As Long
    
  'si modo formulario no se puede ajustar ancho
  If Me.dsiFormulario = "F" Then
    Exit Property
  End If
    
  'guardo posicion celda activa antes de congelar
  lngFilaAnt = Me.spdGrid.ActiveRow
  lngColAnt = Me.spdGrid.ActiveCol
    
  'congelo
  Me.spdGrid.ColsFrozen = Me.spdGrid.ActiveCol
  
  'recupero posicion de fila columna
  Me.spdGrid.SetActiveCell lngColAnt, lngFilaAnt
  
  'set foco
  If Me.spdGrid.Visible = True Then
    Me.spdGrid.SetFocus
  Else
    Me.spdE.SetFocus
  End If
  
End Property

Public Property Get dsiColumnasDescongelar() As Boolean
  
  Dim lngFilaAnt, lngColAnt As Long
    
  'si modo formulario no se puede ajustar ancho
  If Me.dsiFormulario = "F" Then
    Exit Property
  End If
    
  'guardo posicion celda activa antes de congelar
  lngFilaAnt = Me.spdGrid.ActiveRow
  lngColAnt = Me.spdGrid.ActiveCol
    
  'descongelo
  Me.spdGrid.ColsFrozen = 0
  
  'recupero posicion de fila columna
  Me.spdGrid.SetActiveCell lngColAnt, lngFilaAnt
  
End Property


Public Property Get dsiColumnasOrdenar(strTipo) As Boolean
  
  Dim strRowName As Variant
  Dim strType As String
  Dim blnB As Boolean
  Dim intI As Integer
  
  'tomo id de tipo de columna
  intI = Me.dsiEsquema(Me.spdGrid.ActiveCol - 1).Type
    
  'si columna es de tipo text o Ntext, no permito ordenar, no lo soporta SQL server
  If intI = conText Or intI = conNtext Then
     Exit Property
  End If
  
  'cambio puntero mouse
  Screen.MousePointer = vbHourglass
  
  'tomo titulo columna en donde se hizo click
  intRes = Me.spdGrid.GetText(Me.spdGrid.ActiveCol, 0, strRowName)
  
  'set color default en grilla en fila para abandonar
  blnB = Me.dsiFilaPintaNo(Me.spdGrid.ActiveCol, Me.spdGrid.ActiveRow)

  'tipo de orden
  If strTipo = "asc" Then
    SortKeyOrder = Array(1, 1)
    Me.dsiOrderBy = "[" & strRowName & "] asc"
  Else
    SortKeyOrder = Array(2, 2)
    Me.dsiOrderBy = "[" & strRowName & "] desc"
  End If
  
  'set columna a ordenar, orden ascendente, todas las filas
  SortKeys = Array(Me.spdGrid.ActiveCol, Me.spdGrid.ActiveCol)
  Me.spdGrid.Sort 1, 1, -1, -1, SS_SORT_BY_ROW, SortKeys, SortKeyOrder
  
  'set color en grilla en fila activa
  blnB = Me.dsiFilaPintaSi(1, Me.spdGrid.ActiveRow)
  
  'muestro en barra de estado
  blnB = Me.dsiUbicacionGrillaSi()
  blnB = Me.dsiOrderByMostrarSi()
  
  'recupero puntero mouse
  Screen.MousePointer = vbDefault
  
End Property

'cambiar apariencia a grilla
Public Property Get dsiCambiaApariencia() As Boolean
  
  Me.spdGrid.UnitType = UnitTypeTwips                  'trabajar en twips
  
  Me.spdGrid.Appearance = AppearanceFlat               'apariencia 3D
  Me.spdGrid.BorderStyle = BorderStyleNone             'tipo de borde: sin borde
  
  Me.spdGrid.ColHeadersAutoText = DispBlank            'titulos de columnas en blanco
  Me.spdGrid.ColHeadersShow = True                     'muestra encabezado columnas
  Me.spdGrid.RowHeadersShow = True                     'muestra encabezado de filas
  
  Me.spdGrid.CursorStyle = CursorStyleArrow            'stilo cursor
  Me.spdGrid.CursorType = CursorTypeDefault            'tipo cursor
  
  Me.spdGrid.AutoSize = False                          'automaticamente ajusta ancho grilla
  Me.spdGrid.DAutoSizeCols = DAutoSizeColsNone         'tipo de ajuste 2: al dato mas ancho
  
  Me.spdGrid.UserColAction = UserColActionDefault      'cuando hace click en header pinta columna o fila
  Me.spdGrid.FontSize = 9                              'tamaño letra
  Me.spdGrid.RowHeight(0) = 450                        'altura fila de titulos
  Me.spdGrid.MoveActiveOnFocus = False                 '
  Me.spdGrid.Protect = False                           'exporta a excel sin proteccion
  
  Me.spdGrid.BackColorStyle = BackColorStyleUnderGrid  'estilo
  Me.spdGrid.GridShowHoriz = True                      'muestra grilla horizontal
  Me.spdGrid.GridShowVert = True                       'muestra grilla vertical
  Me.spdGrid.GridColor = RGB(200, 200, 200)            'color muy suave
  Me.spdGrid.NoBorder = True                           'sin borde fin zona de datos
  
  Me.spdGrid.ScrollBars = ScrollBarsBoth               'ambas barras de desplazamiento
  Me.spdGrid.ScrollBarExtMode = True                   'cuando sean necesarias
  Me.spdGrid.VScrollSpecial = False                     'barra especial
      
  Me.spdGrid.SetOddEvenRowColor RGB(245, 245, 245), RGB(60, 60, 60), RGB(245, 245, 245), RGB(60, 60, 60)
  Me.spdGrid.SelBackColor = RGB(204, 230, 255)         'fondo del area seleccionada
  Me.spdGrid.GrayAreaBackColor = RGB(245, 245, 245)
  
  Me.spdGrid.VirtualMode = False                        ' ajusta rows al tamaño del recordset
  'spd.VirtualRows = 300                         ' rows a leer del virtual buffer
  'spd.VirtualScrollBuffer = True                ' scroll vertical lee de tantas rows del buffer
  
  'setea para mostrar tooltip en las celdas donde no se ve toda la info
  Me.spdGrid.TextTip = TextTipFixedFocusOnly
  Me.spdGrid.TextTipDelay = 250
  
End Property

'cambiar apariencia a grilla
Public Property Get dsiCambiaAparienciaE() As Boolean
  
  Dim blnB As Boolean
  
  'estilo y tipo de cursor
  Me.spdE.CursorStyle = CursorStyleArrow
  Me.spdE.CursorType = CursorTypeDefault
  
  'grilla en twips
  Me.spdE.UnitType = UnitTypeTwips
  
  'fondo del area seleccionada amarillo
  Me.spdE.SelBackColor = RGB(204, 230, 255)
    
  'set apariencia de borde 3d
  Me.spdE.Appearance = Appearance3D
      
  'set ancho de titulos de filas y columnas
  Me.spdE.RowHeadersAutoText = DispBlank
  Me.spdE.ColHeadersAutoText = DispBlank
  Me.spdE.RowHeight(0) = 105
  Me.spdE.ColWidth(0) = 105
      
  'barra de desplazamiento vertical solo cuando la necesita
  Me.spdE.ScrollBars = ScrollBarsNone
      
  'set muestro cuadriculado horizontal y vertical y color de fondo debajo de grilla
  Me.spdE.BackColorStyle = BackColorStyleUnderGrid
  Me.spdE.GridShowHoriz = True
  Me.spdE.GridShowVert = True
          
  'set cuando enter voy fila siguiente
  Me.spdE.EditEnterAction = EditEnterActionDown
          
  'set cuando ingreso valor arriba de otro lo reemplaza
  Me.spdE.EditModeReplace = True
          
  'set lineas de fondo
  Me.spdE.GridColor = RGB(192, 192, 192)
      
  'set area fuera de la grilla blanco
  Me.spdE.GrayAreaBackColor = RGB(255, 255, 255)
      
  'estilo borde grilla no
  Me.spdE.BorderStyle = BorderStyleFixedSingle
      
  Me.spdE.NoBorder = True
      
  'setea para mostrar tooltip en las celdas donde no se ve toda la info
  Me.spdE.TextTip = TextTipFixed
  Me.spdE.TextTipDelay = 250
  
End Property

'ajustar grilla a frm
Public Property Get dsiAjusta() As Boolean
      
  Me.spdGrid.ReDraw = False
  Me.spdGrid.Top = ScaleTop
  Me.spdGrid.Left = ScaleLeft
  Me.spdGrid.Height = Me.ScaleHeight
  Me.spdGrid.Width = Me.ScaleWidth
  Me.spdGrid.ReDraw = True

End Property

'refresh primera vez
Public Property Get dsiRefreshPrimeraVez() As Boolean
  
  Dim rs As ADODB.Recordset
  Dim strSource, strTop, strWhere, strOrderBy, strFiltraGen As String
  Dim intI As Integer
  Dim strT As String
    
  'puntero mouse reloj
  Screen.MousePointer = vbHourglass
   
  'tomo nombre de tabla o vista
  strSource = Me.dsiConfigGetValor("vista")
      
  'get cantidad Filas
  strTop = Me.dsiConfigGetValor("cantidadFilas")
      
  'check Cantidad filas blanco, get standard
  If strTop = "" Then
    strTop = SQLparam.CantidadFilas
  End If
    
  'si CantidadFilas no esta en el INI, chequeo si esta en la configuracion de la Opcion
  
  'si tampoco esta en la configuracion de la Opcion, tomo un default de 100
  If strTop = "" Then
    strTop = 100
  End If
  
  'si no hay filtro leo tantas filas como FilasMaximo este configurado
  If Me.dsiWhere = "" And MainMDI.dsiWhereGen("condicion") = "" Then
    strSource = "select top " & strTop & " * from " & strSource
  Else
    strSource = "select * from " & strSource
  End If
      
  'clear condicion
  strWhere = ""
  
  'check si grilla se le debe aplicar filtro general
  strFiltraGen = Me.dsiConfigGetValor("filtra")
  If strFiltraGen <> "" Then
  
    'primer where: si existe where General lo agrego al query
    If MainMDI.dsiWhereGen("condicion") <> "" Then
      strWhere = " where " & MainMDI.dsiWhereGen("condicion")
    End If
  
  End If
    
  'segundo where: si existe where en grilla lo agrego al query
  If Me.dsiWhere <> "" Then
    
    If strWhere <> "" Then
      strWhere = strWhere & " and " & Me.dsiWhere
    Else
      strWhere = " where " & Me.dsiWhere
    End If
  
  End If
  
  'check si existe orden, build OrderBy
  If Me.dsiOrderBy <> "" Then
    strOrderBy = " order by " & Me.dsiOrderBy
  End If
  
  'abro rs
  Set rs = SQLexec(strSource & strWhere & strOrderBy)
          
  'chequeo error
  If Not SQLparam.CnErrNumero = -1 Then
    
    SQLError                              'rutina error
    Me.dsiOrderBy = ""                    '10/10/2008 clear order by, por si columna ya no existe
    Screen.MousePointer = vbDefault       'recupero mouse
    Exit Property
  
  End If
      
  'cierro
  SQLclose
      
  'set propiedad fields - estructura de la vista
  Set Me.dsiEsquema = rs.Fields
      
  'rs a grilla
  Set Me.spdGrid.DataSource = rs
    
  'limites de grilla
  Me.spdGrid.MaxRows = rs.RecordCount
  Me.spdGrid.MaxCols = rs.Fields.Count
  
  'bloqueo grilla completa para que no puedan modificar info
  Me.spdGrid.row = -1
  Me.spdGrid.col = -1
  Me.spdGrid.Lock = True
  Me.spdGrid.Protect = True
        
  'recupero puntero mouse
  Screen.MousePointer = vbDefault
  
End Property

Public Property Get dsiRefresh(Optional ByVal blnNOTcnClose As Boolean) As Boolean
  
  Dim rs As ADODB.Recordset
  Dim strDataSource, strDataTop, strDataWhere, strDataOrder As String
  Dim intI As Integer
  Dim strT As String
  Dim lngFilaAnt, lngColAnt As Long
    
  'puntero mouse reloj
  Screen.MousePointer = vbHourglass
   
  'guardo fila y columna activa antes del refresh
  lngFilaAnt = Me.spdGrid.ActiveRow
  lngColAnt = Me.spdGrid.ActiveCol
   
  'tomo nombre de tabla o vista
  strSource = Me.dsiConfigGetValor("vista")
      
  'tomo Cantidad de filas a mostrar
  strTop = SQLparam.CantidadFilas
  
  'si CantidadFilas no esta en el INI, chequeo si esta en la configuracion de la Opcion
  If strTop = "" Then
    strTop = Me.dsiConfigGetValor("cantidadFilas")
  End If
  
  'si tampoco esta en la configuracion de la Opcion, tomo un default de 100
  If strTop = "" Then
    strTop = 100
  End If
  
  'si no hay filtro leo tantas filas como FilasMaximo este configurado
  If Me.dsiWhere = "" And MainMDI.dsiWhereGen("condicion") = "" Then
    strSource = "select top " & strTop & " * from " & strSource
  Else
    strSource = "select * from " & strSource
  End If
    
  'clear condicion
  strWhere = ""
  
  'check si grilla se le debe aplicar filtro general
  strFiltraGen = Me.dsiConfigGetValor("filtra")
  If strFiltraGen <> "" Then
  
    'primer where: si existe where General lo agrego al query
    If MainMDI.dsiWhereGen("condicion") <> "" Then
      strWhere = " where " & MainMDI.dsiWhereGen("condicion")
    End If
  
  End If
  
  'segundo where: si existe where lo agrego al query
  If Me.dsiWhere <> "" Then
    
    If strWhere <> "" Then
      strWhere = strWhere & " and " & Me.dsiWhere
    Else
      strWhere = " where " & Me.dsiWhere
    End If
  
  End If
    
  'check si existe orden, build OrderBy
  If Me.dsiOrderBy <> "" Then
    strOrderBy = " order by " & Me.dsiOrderBy
  End If
  
  'abro rs
  Set rs = SQLexec(strSource & strWhere & strOrderBy)
    
  'chequeo error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    Screen.MousePointer = vbDefault   'recupero puntero mouse
    Exit Property
  End If
      
  'si es true no se cierra conexion
  If Not blnNOTcnClose Then
    'cierro
    SQLclose
  End If
      
  'set propiedad fields - estructura de la tabla/vista
  Set Me.dsiEsquema = rs.Fields
      
  'deshabilito para no vuelva a configurar el typo de celda
  Me.spdGrid.DAutoCellTypes = False
    
  'le paso el recordset a la grilla
  Set Me.spdGrid.DataSource = rs
  
  'habilito para no vuelva a configurar el typo de celda
  Me.spdGrid.DAutoCellTypes = True
    
  'set limites
  Me.spdGrid.MaxRows = rs.RecordCount
  Me.spdGrid.MaxCols = rs.Fields.Count
  
  'defino variables
  Dim intCol As Integer
  Dim strNombreCol  As String
    
  'bloqueo grilla completa para que puedan editar info
  Me.spdGrid.row = -1
  Me.spdGrid.col = -1
  Me.spdGrid.Lock = True
  Me.spdGrid.Protect = True
        
  'set recupero celda activa
  Me.spdGrid.SetActiveCell lngColAnt, lngFilaAnt
    
  'pinto linea
  blnB = Me.dsiFilaPintaSi(Me.spdGrid.ActiveCol, Me.spdGrid.ActiveRow)
      
  'muestro info en barra de estado
  blnB = dsiUbicacionGrillaSi()
  
  'check si formulario en modo F, copy datos de spdGrid a spdE
  If Me.dsiFormulario = "F" Then
    blnB = Me.dsiFormularioDatosPutPaF()
  End If
  
  'recupero puntero mouse
  Screen.MousePointer = vbDefault
    
End Property


Private Sub Form_Activate()
  
  'set celda activa - sin uso a partir de 15/03/2005
  'Me.spdGrid.SetActiveCell Val(Me.dsiGrillaIniCol), Val(Me.dsiGrillaIniFila)
  
  'muestro info en barra de estado
  blnB = Me.dsiUbicacionGrillaSi()
  blnB = Me.dsiOrderByMostrarSi()
  blnB = Me.dsiWhereMostrarSi()
  blnB = Me.dsiSumAvgMinMaxSi()
    
  'pinto fila activa - sin uso a partir de 15/03/2005
  'blnB = Me.dsiFilaPintaSi(Val(Me.dsiGrillaIniCol), Val(Me.dsiGrillaIniFila))
      
  If Not Me.dsiFormularioActivado() Then
    
    blnB = Me.dsiFilaPintaSi(1, 1)
    Me.dsiFormularioActivado = True
    
  End If
      
  'check operacion, para saber si habilito barra o no
  If Me.dsiOperacion = "VER" Then
    blnB = Me.dsiHerramientas(True)
  Else
    blnB = Me.dsiHerramientas(False)
    blnB = Me.dsiHerramientasComfirmar(True)
  End If
      
End Sub

Private Sub Form_Load()
  
  Dim rs As ADODB.Recordset
  Dim strT As String
  Dim intI As Integer
  Dim blnB As Boolean
  Dim arrUbica As Variant
  
  'chequeo si existe tabla en donde se guarda configuracion
  strT = "select * from sysobjects where name = 'dsiOPCconfig'"
  Set rs = SQLexec(strT)
  
  'chequeo errores
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    Exit Sub
  End If
      
  'chequeo que haya configuracion
  If rs.EOF Then
    intI = MsgBox("No se encontro la tabla en donde se guarda configuración para la opción seleccionada.", vbCritical + vbOKOnly, "atención...")
    SQLclose
    Exit Sub
  End If
    
  'si encontro
  If Not rs.EOF Then
  
    'leo configuracion
    strT = "select * from dsiOPCconfig where IDmenu = " & Me.dsiIDmenu & " and IDopc = " & Me.dsiIDopc
    Set rs = SQLexec(strT)
    
    'chequeo errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
      Exit Sub
    End If
    
    'chequeo que haya configuracion
    If rs.EOF Then
      intI = MsgBox("No se encontro configuración para la opción seleccionada.", vbCritical + vbOKOnly, "atención...")
      SQLclose
      Exit Sub
    End If
    
    'recorro
    While Not rs.EOF
           
      'guardo configuracion
      Me.dsiConfigAgregar = Array(rs!tipo, rs!columna, rs!valor, rs!parametro)
      
      'puntero al proximo
      rs.MoveNext
      
    Wend
  
  End If
  
  'leo otras configuraciones
  Me.Caption = Me.dsiIDopc & " - " & Me.dsiTitulo
  arrUbica = keyIniToArray(Me.dsiIDopc, "ubicacionFrm")
  Me.dsiCondicion = keyIniToArray(Me.dsiIDopc, "where")
  Me.dsiOrderBy = ReadIni(Me.dsiIDopc, "orderBy")
  Me.dsiColumnasCongeladas = ReadIni(Me.dsiIDopc, "columnasCongeladas")
  Me.dsiColumnasAncho = keyIniToArray(Me.dsiIDopc, "ColumnasAncho")
  Me.dsiColumnasOcultas = keyIniToArray(Me.dsiIDopc, "ColumnasOcultas")
  Me.dsiColumnasDecimales = keyIniToArray(Me.dsiIDopc, "ColumnasDecimales")
  
  'permisos full
  Me.dsiSpINSpermiso = True
  Me.dsiSpEDIpermiso = True
  Me.dsiSpELIpermiso = True
      
  'si existe ubicacion
  If UBound(arrUbica) <> -1 Then
      
    'asigno state
    Me.WindowState = arrUbica(2)
  
    'si ubicacion es manual
    If arrUbica(2) = 0 Then
            
      'si width y height es 0 abro default mitad de pantalla
      If arrUbica(8) = 0 And arrUbica(10) = 0 Then
        Me.Move arrUbica(4), arrUbica(6), arrUbica(8) / 2, arrUbica(10) / 2
      Else
         Me.Move arrUbica(4), arrUbica(6), arrUbica(8), arrUbica(10)
      End If
      
    End If
      
    'guardo coordenadas iniciales de grilla
    Me.dsiGrillaIniFila = arrUbica(12)
    Me.dsiGrillaIniCol = arrUbica(14)
      
  End If
      
  'cambia apariencia spdGrid
  blnB = Me.dsiCambiaApariencia()
    
  'cambia apariencia spdE
  blnB = Me.dsiCambiaAparienciaE()
    
  'lee rs primera vez
  blnB = Me.dsiRefreshPrimeraVez()
      
  'set visualizar form como planilla default
  blnB = Me.dsiFormularioCambiar()
          
  'set operacion
  Me.dsiOperacion = "VER"
            
  '-----------------
'  Me.mnuDet.MenuItems.Add 0, 1, smiNone, "Detalle"
'  Me.mnuDet.MenuItems.Add 1, , smiNone, "Opción 1"
'  Me.mnuDet.MenuItems.Add 1, , smiNone, "Opción 2"
            
End Sub

Private Sub Form_Resize()
    
  Dim bln As Boolean
  bln = Me.dsiAjusta()
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  Dim intI As Integer
  Dim lngL As Long
  Dim strT As String
  Dim blnB As Boolean
    
  'save configuracion en INI------------------------------------------------------------------------
  
  'save ubicacion frm y grilla
  strT = "state;" & Me.WindowState & _
         ";left;" & Me.Left & ";top;" & Me.Top & ";width;" & Me.Width & ";height;" & Me.Height & _
         ";fila;" & Me.spdGrid.ActiveRow & ";col;" & Me.spdGrid.ActiveCol
  blnB = WriteIni(Me.dsiIDopc, "ubicacionFrm", strT)
  
  'armo where
  strT = ""
  For intI = 1 To Me.dsiCondicionCant
    strT = strT & Me.dsiCondicionGet(intI) & ";"
  Next
    
  'elimino ultima coma
  If strT <> "" Then
    strT = Left(strT, Len(strT) - 1)
  End If
  
  'save where
  blnB = WriteIni(Me.dsiIDopc, "where", strT)

  blnB = WriteIni(Me.dsiIDopc, "orderBy", Me.dsiOrderBy)
  
  'save columnas congeladas
  strT = ""
  strT = Me.spdGrid.ColsFrozen
  intRes = WriteIni(Me.dsiIDopc, "ColumnasCongeladas", strT)
  
  'ancho de columnas
  strT = ""
  For lngL = 1 To Me.spdGrid.MaxCols
    Me.spdGrid.col = lngL
    strT = strT & lngL & ";" & Me.spdGrid.ColWidth(lngL) & ";"
  Next
  
  'elimino la ultima coma
  If strT <> "" Then
    strT = Left(strT, Len(strT) - 1)
  End If
  
  'save ancho columnas
  blnB = WriteIni(Me.dsiIDopc, "ColumnasAncho", strT)
  
  'columnas ocultas
  strT = ""
  For lngL = 1 To Me.spdGrid.MaxCols
    Me.spdGrid.col = lngL
    If Me.spdGrid.ColHidden = True Then
      strT = strT & lngL & ";"
    End If
  Next
  
  'elimino la ultima coma
  If strT <> "" Then
    strT = Left(strT, Len(strT) - 1)
  End If
  
  'save columnas ocultas
  blnB = WriteIni(Me.dsiIDopc, "ColumnasOcultas", strT)
  
  'cantidad decimales
  strT = ""
  For lngL = 1 To Me.spdGrid.MaxCols
    Me.spdGrid.col = lngL
    If Me.spdGrid.CellType = CellTypeNumber Then
      strT = strT & lngL & ";" & Me.spdGrid.TypeNumberDecPlaces & ";"
    ElseIf Me.spdGrid.CellType = CellTypeCurrency Then
          strT = strT & lngL & ";" & Me.spdGrid.TypeCurrencyDecPlaces & ";"
        End If
  Next
  
  'elimino la ultima coma
  If strT <> "" Then
    strT = Left(strT, Len(strT) - 1)
  End If
  
  'save cantidad decimales
  blnB = WriteIni(Me.dsiIDopc, "ColumnasDecimales", strT)
      
  'save forma de ver un formulario
  blnB = WriteIni(Me.dsiIDopc, "Modo", Me.dsiFormulario)
      
  'save operacion
  blnB = WriteIni(Me.dsiIDopc, "Operacion", Me.dsiOperacion)
      
  'limpio barra de estado
  blnB = Me.dsiUbicacionGrillaNo()
  blnB = Me.dsiOrderByMostrarNo()
  blnB = Me.dsiWhereMostrarNo()
  blnB = Me.dsiSumAvgMinMaxNo()
  
  'set herramientas
  blnB = dsiHerramientas(False)
  blnB = dsiHerramientasComfirmar(False)
  
End Sub

Private Sub spdE_KeyPress(KeyAscii As Integer)
  
'CHECK si se tipeo un enter, lo anulo
If KeyAscii = Asc(vbCrLf) Then
  KeyAscii = 0
End If
  
End Sub

Private Sub spdE_LeaveCell(ByVal col As Long, ByVal row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
  Dim blnB As Boolean
  Dim intI As Integer
  Dim varDato As Variant
  
  'si el modo es VER, mantiene actualizada la ubicacion
  'de celda para poder filtrar u ordenar en modo ver
  If Me.dsiOperacion = "VER" Or Me.dsiOperacion = "EDI" Then
  
    'cuando la grilla pierde el foco, pasa por aca, y pone a newCol y newRow en -1
    'y me caga cuando muestro la ubicacion de celda en pantalla por eso este IF
    If NewCol <> -1 And NewRow <> -1 Then
            
      'defa formulario simple
      intI = 0
      
      'formulario doble
      If NewCol = 4 Or NewCol = 5 Then
        intI = Me.spdE.MaxRows
      End If
          
      'get totilo de columna
      Me.spdE.GetText intI, NewRow, varDato
          
      'busco ubicacion columna
      'a = Me.spdGrid.SearchRow(Me.spdGrid.ActiveRow, 1, -1, varDato, SearchFlagsNone)
      
      'set puntero a fila columna
      Me.spdGrid.SetActiveCell NewRow + intI, Me.spdGrid.ActiveRow
        
      'muestra ubicacion en barra de estado
      blnB = dsiUbicacionGrillaSi()
      
    End If
  
  End If
    
End Sub

Private Sub spdGrid_DataColConfig(ByVal col As Long, ByVal DataField As String, ByVal DataType As Integer)
    
  Dim intDecimales As Integer
  Dim sngAncho As Single
  Dim blnB As Boolean
  Dim lngL As Long
      
  'set columna
  Me.spdGrid.col = col
  Me.spdGrid.row = -1
  
  'si hay por lo menos 1 fila, configuro
  If Me.spdGrid.MaxRows > 0 Then
            
    'set no muestra separador de miles
    Me.spdGrid.TypeNumberShowSep = False
        
    'activo columna con ........ cuando no se ve todo el texto
    Me.spdGrid.TypeEllipses = True
    
    'activo limite de columna tipo texto a 5000
    Me.spdGrid.TypeMaxEditLen = 5000
      
    'tomo cantidad de decimales configurados por usuario
    intDecimales = Me.dsiColumnasDecimalesGet(col)
  
    'si existen decimales
    If intDecimales <> -1 Then
          
      'set row
      Me.spdGrid.row = 1
      
      'si decimales actuales son diferentes a los configurado por el usuario
      If Me.spdGrid.TypeNumberDecPlaces <> intDecimales Then
        
        'marco todas las filas
        Me.spdGrid.row = -1
        
        'set decimales nuevos
        Me.spdGrid.TypeNumberDecPlaces = intDecimales
        
      End If
       
    End If
        
  End If
        
  'tomo ancho de columna configurado por usuario
  sngAncho = Me.dsiColumnasAnchoGet(col)
  
  'si existen ancho
  If sngAncho <> -1 Then
          
    'set row
    'Me.spdGrid.row = 1
      
    'si ancho actual es diferente al  configurado por el usuario
    If Me.spdGrid.ColWidth(col) <> sngAncho Then
        
      'marco todas las filas
      Me.spdGrid.row = -1
        
      'set decimales nuevos
      Me.spdGrid.ColWidth(col) = sngAncho
  
    End If
    
  End If

  'tomo columna oculta configurado por usuario
  blnB = Me.dsiColumnasOcultasGet(col)
  
  'si oculta
  If blnB = True Then
              
    Me.spdGrid.ColHidden = True
    
  End If

  'tomo columna congeladas y aplico
  lngL = Val(Me.dsiColumnasCongeladas())
  
  If lngL <> 0 Then
    Me.spdGrid.ColsFrozen = lngL
  End If

End Sub

'** 08/01/08 -- NUEVO PARA GUARDAR UN VINCULO EN LA GRILLA
Private Sub spdE_DblClick(ByVal col As Long, ByVal row As Long)

  Dim varNombre As Variant
 
  'levanto el nombre de la columna
  Me.spdE.GetText 1, row, varNombre
    
  If Me.dsiFormulario = "F" And Me.dsiConfigGetValor("VINCULO", varNombre) <> "" And col = 2 Then
  
    With comArchivos
      .FileName = vbNullString
      .DialogTitle = "Seleccionar archivo"
      .Filter = "Todos los tipos *.*|*.*"
  
      'abre
      .ShowOpen

      If .FileName <> vbNullString Then
        Me.spdE.SetText col, row, .FileName
        
      End If
    End With
  End If
  
End Sub

'** 08/01/08 -- NUEVO PARA ABRIR UN VINCULO
Private Sub spdGrid_Click(ByVal col As Long, ByVal row As Long)
  
  Dim blnB As Boolean
 
  'hace algo si el dbClick fue arriba de fila o columna 0, sino escapo
  If row > 0 Then
    'llamo para abrir archivo
    Call spdAbrirArchivos(col, row, Me.spdGrid)
  Else
    Exit Sub
  End If
    
  'ajusto ancho
  blnB = Me.dsiColumnasAnchoAjustar()
  
End Sub

'** 08/01/08 -- NUEVO PARA ABRIR UN VINCULO DESDE LA GRILLA
Private Sub spdAbrirArchivos(ByVal col As Long, ByVal row As Long, NewSpd As fpSpread)
  
  Dim blnCol As Boolean
  Dim strRuta As String
  Dim varNombre As Variant
 
  If Me.dsiFormulario = "P" Then
    'levanto el nombre de la columna
    NewSpd.GetText col, 0, varNombre
    strRuta = NewSpd.Text
    blnCol = True
  Else
    NewSpd.GetText 1, row, varNombre
    If col = 2 Then
      blnCol = True
      strRuta = NewSpd.Text
    Else
      blnCol = False
    End If
  End If
  
  'SI HACE DOBLE CLICK EN UNA CELDA DE VINCULO Y NO ESTA VACIO ABRE EL ARCHIVO
  If Me.dsiConfigGetValor("VINCULO", varNombre) <> "" And strRuta <> "" And blnCol Then
    ShellExecute Me.hwnd, "open", strRuta, "", App.Path, SW_SHOWNORMAL
    NewSpd.ForeColor = &H80000002
  Else
    Exit Sub
  End If

End Sub

Private Sub spdGrid_DblClick(ByVal col As Long, ByVal row As Long)
  
  Dim blnB As Boolean
    
  'hace algo si el dbClick fue arriba de fila o columna 0, sino escapo
  If row > 0 Then
    Exit Sub
  End If
    
  'ajusto ancho
  blnB = Me.dsiColumnasAnchoAjustar()
      
End Sub

Private Sub spdGrid_KeyUp(KeyCode As Integer, Shift As Integer)

  Dim blnB As Boolean
  
  blnB = show_SumAvgMinMax()

End Sub

Private Sub spdGrid_LeaveCell(ByVal col As Long, ByVal row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
        
  Dim blnB As Boolean
  
  'cuando la grilla pierde el foco, pasa por aca, y pone a newCol y newRow en -1
  'y me caga cuando muestro la ubicacion de celda en pantalla por eso este IF
  If NewCol <> -1 And NewRow <> -1 Then
            
    'set puntero a fila columna
    Me.spdGrid.SetActiveCell NewCol, NewRow
        
    'marca fila como activa
    blnB = dsiFilaPintaNo(col, row)
    blnB = dsiFilaPintaSi(NewCol, NewRow)
    
    'muestra ubicacion en barra de estado
    blnB = dsiUbicacionGrillaSi()
    
  End If
  
End Sub

'Private Sub spdGrid_RightClick(ByVal ClickType As Integer, ByVal col As Long, ByVal row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
  
'      x = Me.mnuDet.ClientToScreenX(Me.hWnd, x)
'      y = Me.mnuDet.ClientToScreenY(Me.hWnd, y)

'  Me.mnuDet.PopupMenu Me.mnuDet.MenuItems.Key2ID("Detalle"), x, y, 0
  
'End Sub

Private Function show_SumAvgMinMax() As Boolean

  Dim dblSUM, dblAVG, dblMIN, dblMAX As Double
  Dim lngFila, lngCol, lngCAN As Long
  Dim c1, r1, c2, r2 As Variant
  
  'inicializo
  dblSUM = 0
  dblAVG = 0
  dblMIN = 0
  dblMAX = 0
  lngCAN = 0
  
  'get rango seleccion
  Me.spdGrid.GetSelection 0, c1, r1, c2, r2
     
  'while filas
  For lngFila = r1 To r2
       
    'while columnas
    For lngCol = c1 To c2
      
      'set puntero interono en fila columna
      Me.spdGrid.row = lngFila
      Me.spdGrid.col = lngCol
      
      'check si numerico, acumulo y sumo
      If IsNumeric(Me.spdGrid.Value) Then
        
        'add 1 a contador de valores
        lngCAN = lngCAN + 1
        
        'fuezo valor minimo
        If lngCAN = 1 Then
          dblMIN = Me.spdGrid.Value
        End If
        
        'calc sum
        dblSUM = dblSUM + Me.spdGrid.Value
        
        'calc avg
        dblAVG = dblSUM / lngCAN
        
        'save minimo
        If dblMIN > Me.spdGrid.Value Then
          dblMIN = Me.spdGrid.Value
        End If
        
        'save maximo
        If dblMAX < Me.spdGrid.Value Then
          dblMAX = Me.spdGrid.Value
        End If
        
      End If
        
    Next
  
  Next
    
  'show sum, avg, min, max
  strT = "Sum=" & dblSUM & ", Avg=" & Round(dblAVG, 3) & ", Min=" & dblMIN & ", Max=" & dblMAX
  MainMDI.staBarra1.Panels(3).Text = strT
  MainMDI.staBarra1.Panels(3).ToolTipText = strT

End Function

Private Sub spdGrid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  Dim blnB As Boolean
  
  blnB = show_SumAvgMinMax()
      
End Sub

