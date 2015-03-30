VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpspr60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form filtrarGenFRM 
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3930
   Icon            =   "filtrarGenFRM.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3585
   ScaleWidth      =   3930
   Begin MSComDlg.CommonDialog comDestino 
      Left            =   2880
      Top             =   1980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frame 
      BorderStyle     =   0  'None
      Height          =   1245
      Left            =   150
      TabIndex        =   1
      Top             =   1770
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   825
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   840
         Width           =   825
      End
      Begin VB.ListBox lstLista 
         Height          =   735
         ItemData        =   "filtrarGenFRM.frx":038A
         Left            =   90
         List            =   "filtrarGenFRM.frx":038C
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   60
         Width           =   1725
      End
   End
   Begin FPSpreadADO.fpSpread spdF 
      Height          =   1575
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   3405
      _Version        =   393216
      _ExtentX        =   6006
      _ExtentY        =   2778
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
      SpreadDesigner  =   "filtrarGenFRM.frx":038E
   End
End
Attribute VB_Name = "filtrarGenFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_IDmenu As String
Private m_IDopc As String
Private m_Titulo As String

Private m_Esquema As ADODB.Fields

Private m_dsiColumnasFiltradasCant As Integer
Private m_dsiColumnasFiltradas() As Variant
Private m_dsiColumnasOperadores() As Variant
Private m_dsiColumnasCriterios() As Variant
Private m_dsiColumnasItem() As Variant
Private m_dsiCondicion As String
Private m_dsiCondicionFinal As String

Private m_Config() As String
Private m_ConfigCant As Integer

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

Public Property Let dsiTitulo(str As String)
  m_Titulo = str
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

Public Property Get dsiColumnasFiltradasVaciar() As Boolean
  m_dsiColumnasFiltradasCant = 0
End Property

Public Property Get dsiColumnasFiltradas(intI As Integer) As String
  dsiColumnasFiltradas = m_dsiColumnasFiltradas(intI)
End Property

Public Property Get dsiColumnasFiltradasCant() As Integer
  dsiColumnasFiltradasCant = m_dsiColumnasFiltradasCant
End Property

Public Property Get dsiColumnasFiltradasAgregar(varColumna As Variant) As Boolean
      
  'add columnas filtradas
  m_dsiColumnasFiltradasCant = m_dsiColumnasFiltradasCant + 1
    
  'redim array columnas filtradas
  ReDim Preserve m_dsiColumnasFiltradas(m_dsiColumnasFiltradasCant)
  
  'set columna filtrada
  m_dsiColumnasFiltradas(m_dsiColumnasFiltradasCant) = varColumna
  
End Property

Public Property Get dsiColumnasFiltradasEliminar() As Boolean
  
  'delete columnas filtradas
  m_dsiColumnasFiltradasCant = m_dsiColumnasFiltradasCant - 1
  
  'redim array columnas filtradas
  ReDim Preserve m_dsiColumnasFiltradas(m_dsiColumnasFiltradasCant)
  
End Property

Public Property Get dsiColumnasOperadoresAgregar(varOperador As Variant) As Boolean
      
  'redim array columnas datos
  ReDim Preserve m_dsiColumnasOperadores(m_dsiColumnasFiltradasCant)
  
  'set columna datos
  m_dsiColumnasOperadores(m_dsiColumnasFiltradasCant) = varOperador
  
End Property

Public Property Get dsiColumnasOperadores(intI As Integer) As String
  dsiColumnasOperadores = m_dsiColumnasOperadores(intI)
End Property

Public Property Get dsiColumnasCriteriosAgregar(varCriterio As Variant) As Boolean
      
  'redim array columnas datos
  ReDim Preserve m_dsiColumnasCriterios(m_dsiColumnasFiltradasCant)
  
  'set columna datos
  m_dsiColumnasCriterios(m_dsiColumnasFiltradasCant) = varCriterio
  
End Property

Public Property Get dsiColumnasCriterios(intI As Integer) As String
  dsiColumnasCriterios = m_dsiColumnasCriterios(intI)
End Property

Public Property Get dsiColumnasItem(intI As Integer) As String
  dsiColumnasItem = m_dsiColumnasItem(intI)
End Property

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

'agrega condicion previa
Public Property Let dsiCondicion(ByVal strCondPrevia As String)
  
  Dim strT, strCondFinal, strVista, strColumna As String
  Dim blnB As Boolean
  Dim lngPozos As Long
  
  'add condicion previa
  m_dsiCondicion = strCondPrevia
  
  'si condicion previa vacia, elimino condicion final y exit
  If m_dsiCondicion = "" Then
    blnB = dsiCondicionEliminar()
    Exit Property
  End If
  
  'set condicion final
    
  'get vista
  strVista = dsiConfigGetValor("vista")
    
  'get columna
  strColumna = dsiConfigGetValor("filtra")
    
  'cambio mouse
  Screen.MousePointer = vbHourglass
  
  'armo string
  strT = "select [" & strColumna & "] from " & strVista & " where " & strCondPrevia
  
  'get rs
  Set rs = SQLexec(strT)
    
  'check error
  If Not SQLparam.CnErrNumero = -1 Then
    
    SQLError                        'muestro error
    SQLclose                        'cierro cn
    Screen.MousePointer = vbDefault 'cambio mouse
    Exit Property                   'exit
   
  End If
      
  'inicializo
  strCondFinal = ""
  
  'recorro rs
  While Not rs.EOF
    
    'armo string
    strCondFinal = strCondFinal & rs(0) & ","
                
    'puntero proximo
    rs.MoveNext
    
  Wend
  
  'get cantidad objetos filtrados
  lngPozos = rs.RecordCount
  
  'recupero mouse
  Screen.MousePointer = vbDefault
  
  'si condicion final vacia, add comillas simples para forzar una condicion vacia
  If strCondFinal = "" Then
    strCondFinal = "0,"
  End If
  
  'delete coma en ultima posicion
  strCondFinal = Left(strCondFinal, Len(strCondFinal) - 1)
  
  'add condicion final
  m_dsiCondicionFinal = strCondFinal
  
  'arma condicion final
  strCondFinal = "[" & strColumna & "] in (" & strCondFinal & ")"
  
  'guardo condicion en mainMDI
  blnB = MainMDI.dsiWhereGenAgregar(strColumna, strCondPrevia & " (" & lngPozos & ")", strCondFinal)
    
End Property
  

'armo condicion previa
'
Public Property Get dsiRefresh()
  
  Dim blnB As Boolean
  Dim intI, intIcant As Integer
  Dim lngPozos As Long
  Dim blnColumnaTexto As Boolean
  Dim strParcial, strFinal, strGeneral, strVista, strColumna As String
  Dim varColumna, varOperador, varCriterio, varTipoColumna, arrCriterio As Variant
  Dim rs As ADODB.Recordset
  
  blnB = Me.dsiCondicionVaciar()          'vacio condicion
  blnB = Me.dsiColumnasFiltradasVaciar()  'vacio columnas filtradas
  strFinal = ""                           'vacio condicion final
  
  'recorro columnas
  For intI = 1 To Me.spdF.MaxRows
    
    'get columna, operador, criterio
    Me.spdF.GetText 1, intI, varColumna
    Me.spdF.GetText 2, intI, varOperador
    Me.spdF.GetText 3, intI, varCriterio
    Me.spdF.GetText 5, intI, varTipoColumna
        
    'save si existe filtro
    If varOperador <> "" And varCriterio <> "" Then
      
      'add columnas filtradas
      blnB = Me.dsiColumnasFiltradasAgregar(varColumna)
      blnB = Me.dsiColumnasOperadoresAgregar(varOperador)
      blnB = Me.dsiColumnasCriteriosAgregar(varCriterio)
            
      'set para condicion parcial
      strParcial = ""
      
      'set columna de tipo texto
      blnColumnaTexto = False
      
      'get array con criterio, para saber si hay mas de un criterio
      arrCriterio = separateText(varCriterio, ",")
                  
      'recorro cantidad de criterios
      For intIcant = 1 To UBound(arrCriterio)
                                  
        'case tipo de celda
        Select Case varTipoColumna
          
        'Bit
        Case conBit
          
          'move true = 1, false = 0, sino valor
          Select Case LCase(arrCriterio(intIcant))
          Case "true"
            arrCriterio(intIcant) = 1
          Case "false"
            arrCriterio(intIcant) = 0
          Case Else
            arrCriterio(intIcant) = varCriterio
          End Select
          
          'set condicion parcial
          strParcial = strParcial & arrCriterio(intIcant) & ","
          
        'enteros y decimales
        Case conSmallInt, conInt, conTinyInt, conMoney, conSmallMoney, conReal, conFloat, conNumeric, conDecimal
          
          'set condicion parcial
          strParcial = strParcial & arrCriterio(intIcant) & ","
        
        'fecha
        Case conSmallDateTime, conDateTime
        
          'put comillas y format dateTOiso
          arrCriterio(intIcant) = "'" & dateToIso(arrCriterio(intIcant)) & "'"
          
          'set condicion parcial
          strParcial = strParcial & arrCriterio(intIcant) & ","
          
        'string
        Case conChar, conNchar, conVarchar, conText, conNVarchar, conNtext
          
          'si = like, sino not like
          If LCase(varOperador) = "=" Then
            strParcial = strParcial & "[" & varColumna & "] like '" & arrCriterio(intIcant) & "' or "
          Else
            strParcial = strParcial & "[" & varColumna & "] not like '" & arrCriterio(intIcant) & "' and "
          End If
          
          'set columna string
          blnColumnaTexto = True
          
        End Select
                
      Next
                        
      'si encuentra una coma en la ultima posicion, delete
      If Mid(strParcial, Len(strParcial), 1) = "," Then
        strParcial = Left(strParcial, Len(strParcial) - 1)
      End If
      
      'check si columna tipo texto
      If blnColumnaTexto Then
      
        'si encuentra un or en la ultima posicion, delete
        If Mid(strParcial, Len(strParcial) - 3, 4) = " or " Then
          strParcial = Left(strParcial, Len(strParcial) - 4)
        End If
        
        'si encuentra un and en la ultima posicion, delete
        If Mid(strParcial, Len(strParcial) - 4, 5) = " and " Then
          strParcial = Left(strParcial, Len(strParcial) - 5)
        End If
      
      End If
      
      'put menor que y 1 solo criterio, OK, sino, error
      If LCase(varOperador) = "<" And UBound(arrCriterio) = 1 Then
        strParcial = "[" & varColumna & "]" & " < " & " " & strParcial
        strFinal = strFinal & strParcial & " and "
      End If
      
      'put menor o igual que y 1 solo criterio, OK, sino, error
      If LCase(varOperador) = "<=" And UBound(arrCriterio) = 1 Then
        strParcial = "[" & varColumna & "]" & " <= " & strParcial
        strFinal = strFinal & strParcial & " and "
      End If
      
      'put mayor que y 1 solo criterio, OK, sino, error
      If LCase(varOperador) = ">" And UBound(arrCriterio) = 1 Then
        strParcial = "[" & varColumna & "]" & " > " & strParcial
        strFinal = strFinal & strParcial & " and "
      End If
      
      'put mayor o igual que y 1 solo criterio, OK, sino, error
      If LCase(varOperador) = ">=" And UBound(arrCriterio) = 1 Then
        strParcial = "[" & varColumna & "]" & " >= " & strParcial
        strFinal = strFinal & strParcial & " and "
      End If
      
      'put entre y 2 criterios, OK, sino, error
      If LCase(varOperador) = "[]" And UBound(arrCriterio) = 2 Then
        strParcial = "[" & varColumna & "]" & " between " & arrCriterio(1) & " and " & arrCriterio(2)
        strFinal = strFinal & strParcial & " and "
      End If
      
      'put igual
      If LCase(varOperador) = "=" Then
      
        'si columna texto
        If blnColumnaTexto Then
          strParcial = "(" & strParcial & ")"
          strFinal = strFinal & strParcial & " and "
        Else
          strParcial = "[" & varColumna & "] in " & "(" & strParcial & ")"
          strFinal = strFinal & strParcial & " and "
        End If
      
      End If
      
      'put distinto
      If LCase(varOperador) = "<>" Then
          
        'si columna texto
        If blnColumnaTexto Then
          strParcial = "(" & strParcial & ")"
          strFinal = strFinal & strParcial & " and "
        Else
          strParcial = "[" & varColumna & "] not in " & "(" & strParcial & ")"
          strFinal = strFinal & strParcial & " and "
        End If
        
      End If
      
    End If
        
  Next
    
  'check si hay filtro, elimino ultimo and
  If Len(strFinal) <> 0 Then
      
    'si en las ultimas posiciones hay un and
    If Mid(strFinal, Len(strFinal) - 4, 5) = " and " Then
      strFinal = Left(strFinal, Len(strFinal) - 5)
    End If
    
  End If
  
  'add condicion previa
  Me.dsiCondicion = strFinal
      
End Property

Public Property Get dsiCondicionEliminar() As Boolean
   
  Dim intI  As Integer
  
  'inicializo columnas filtradas
  m_dsiColumnasFiltradasCant = 0
  
  'redim array
  ReDim m_dsiColumnasFiltradas(0)
  ReDim m_dsiColumnasOperadores(0)
  ReDim m_dsiColumnasCriterios(0)
   
  'recorro grilla
  For intI = 1 To Me.spdF.MaxRows
    
    'delete operador y criterio de grilla
    Me.spdF.SetText 2, intI, "="
    Me.spdF.SetText 3, intI, ""
        
  Next
   
  'elimino filtro gen de mainMDI
  blnB = MainMDI.dsiWhereGenEliminar()
  
End Property

Public Property Get dsiCondicionVaciar() As Boolean
  m_dsiCondicion = ""
End Property

Public Property Get dsiFiltroAgrilla() As Boolean
  
  Dim blnB As Boolean
  Dim intI As Integer
  Dim lngL As Long
  
  'recorro columnas filtradas
  For intI = 1 To Me.dsiColumnasFiltradasCant
      
    'busco columna fitrada
    lngL = Me.spdF.SearchCol(1, 0, -1, Me.dsiColumnasFiltradas(intI), SearchFlagsCaseSensitive)
          
    'si encontro nombre de columna
    If lngL <> -1 Then
        
      'put condición en grilla
      Me.spdF.SetText 2, lngL, Me.dsiColumnasOperadores(intI)
      Me.spdF.SetText 3, lngL, Me.dsiColumnasCriterios(intI)
        
    End If
          
  Next
  
End Property


Public Property Get dsiHerramientas(blnB As Boolean) As Boolean
  
  MainMDI.tlbHerra.Buttons("cmd_forma_vista").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_fila_pri").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_fila_ant").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_fila_sig").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_fila_ult").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_insertar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_editar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_eliminar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_buscar").Enabled = blnB
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
  MainMDI.tlbHerra.Buttons("cmd_exportar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_importar").Enabled = blnB

End Sub

Public Property Get dsiHerramientasFiltro(blnB As Boolean) As Boolean

  MainMDI.tlbHerra.Buttons("cmd_actualizar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_filtro_abr").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_filtro_gua").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_filtro_eli").Enabled = blnB

End Property

Public Property Get dsiHerramientasComfirmar(blnB As Boolean) As Boolean
  
  MainMDI.tlbHerra.Buttons("cmd_guardar").Enabled = blnB
  MainMDI.tlbHerra.Buttons("cmd_cancelar").Enabled = blnB

End Property


Public Property Get dsiCambiaApariencia() As Boolean
  
  Dim blnB As Boolean
   
  'set grilla funciona en Twips
  Me.spdF.UnitType = UnitTypeTwips                  'trabajar en twips

  Me.spdF.CursorStyle = CursorStyleArrow            'stilo cursor
  Me.spdF.CursorType = CursorTypeDefault            'tipo cursor
  
  'altura header
  spdF.RowHeight(0) = 300
    
  'estilo borde
  spdF.BorderStyle = BorderStyleNone
    
  'set no muestra nombre de header ni en filas ni en columnas
  spdF.ColHeadersAutoText = DispBlank
  spdF.ColHeadersAutoText = DispBlank
    
  'reemplaza informacion si existe
  spdF.EditModeReplace = True
    
  'quita barra de desplazamientos y habilita tab
  spdF.ScrollBars = ScrollBarsVertical
  spdF.ProcessTab = True

  'para que se vean solo las separaciones verticales
  spdF.BackColorStyle = BackColorStyleOverHorzGridOnly
    
  'habilita las barras de desplazamiento si las necesita
  spdF.ScrollBarExtMode = True
      
  Me.BackColor = RGB(216, 237, 223)               'fondo formulario
  
  Me.spdF.ShadowText = &H404040                   'texto de encabezados
  Me.spdF.GrayAreaBackColor = RGB(245, 245, 245)  'fondo area grilla verde
  Me.spdF.SetOddEvenRowColor RGB(245, 245, 245), RGB(60, 60, 60), RGB(252, 252, 252), RGB(60, 60, 60)
'  Me.spdF.GrayAreaBackColor = RGB(216, 237, 223) 'fondo area grilla verde
  
  'establece numero maximo filas y columnas
  spdF.MaxRows = 0
  spdF.MaxCols = 5
    
  'muestra combo cuando la celda esta activa
  'spdF.ButtonDrawMode = 1 'cell only
  
  spdF.EditModePermanent = True
     
  'columna 0
  spdF.ColWidth(0) = 80
     
  'columna para nombres de columnas del esquema
  spdF.SetText 1, 0, "Columna"
  spdF.ColWidth(1) = 1000
      
  spdF.Col = 1
  spdF.CellType = CellTypeStaticText
 'set celda de tipo label
      
  'columna Operadores
  spdF.SetText 2, 0, "Operador"
  spdF.ColWidth(2) = 1000
  
  'set celda de tipo Buttom
  spdF.Col = 2
  spdF.CellType = CellTypeComboBox
  
  'Populate Criteria header
  spdF.SetText 3, 0, "Criterio"
  spdF.ColWidth(3) = 2000
  
  'Populate Criteria header
  spdF.SetText 4, 0, ""
  spdF.ColWidth(4) = 250
  
  'set celda de tipo Buttom
  spdF.Col = 4
  spdF.CellType = CellTypeButton
  spdF.TypeButtonText = "..."
    
  'esconde columna de tipos de dato y columna string de fecha
  spdF.Col = 5
  spdF.ColHidden = True
  
  'setea para mostrar tooltip en las celdas donde no se ve toda la info
  spdF.TextTip = TextTipFixedFocusOnly
  spdF.TextTipDelay = 250
  
End Property

Public Property Get dsiGetDatos()

  Dim blnB As Boolean
  Dim rs As ADODB.Recordset
  Dim varColumna As Variant
  Dim strSQL As String
  Dim strVista As String
    
  'cambio mouse
  Me.MousePointer = vbHourglass
    
  'get vista
  strVista = dsiConfigGetValor("vista")
    
  'get nombrecolumna activa
  Me.spdF.GetText 1, Me.spdF.ActiveRow, varColumna
    
  'armo string
  strSQL = "select [" & varColumna & "] from " & strVista & " group by [" & varColumna & "]"
  
  'get rs
  Set rs = SQLexec(strSQL)
    
  'check error
  If Not SQLparam.CnErrNumero = -1 Then
    
    SQLError                        'muestro error
    SQLclose                        'cierro cn
    Me.MousePointer = vbDefault     'cambio mouse
    Exit Property                   'exit
  
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
  
  'cierro
  SQLclose
  
  'cambio mouse
  Me.MousePointer = vbDefault
  
End Property

Private Sub cmdAceptar_Click()
    
  Dim intI As Integer
  Dim strT As String
  
  strT = ""
  
  'recorro item seleccionados
  For intI = 1 To Me.lstLista.SelCount
    
    'si item <> -1 agrego al string
    If Me.dsiColumnasItem(intI) <> -1 Then
      strT = strT & Me.lstLista.List(Me.dsiColumnasItem(intI)) & ","
    End If
    
  Next
    
  'elimina ultima coma
  If strT <> "" Then
    strT = Left(strT, Len(strT) - 1)
  End If
    
  'get nombre de columna
  Me.spdF.SetText 3, Me.spdF.ActiveRow, strT
    
  'set visible lista
  Me.frame.Visible = False
  
  'set enable grilla
  Me.spdF.Enabled = True
  
End Sub

Private Sub cmdCancelar_Click()
  
  'set visible lista
  Me.frame.Visible = False
  
  'set enable grilla
  Me.spdF.Enabled = True
  
End Sub

Private Sub Form_Activate()
  
  Dim blnB As Boolean
  
  'habilito / deshabilito herramientas
  blnB = Me.dsiHerramientas(False)
  blnB = Me.dsiHerramientasComfirmar(False)
  blnB = Me.dsiHerramientasFiltro(True)
  blnB = MainMDI.dsiWhereMostrarSi()
  
End Sub

Private Sub Form_Load()
  
  Dim blnB As Boolean
  Dim intI As Integer
  Dim strUbico, strColumnas, strOperadores, strCriterios  As Variant
  Dim sngLeft, sngTop, sngWidth, sngHeight As Single
  
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
           
    'get configuracion
    Me.dsiConfigAgregar = Array(rs!tipo, rs!columna, rs!valor, rs!parametro)
      
    'puntero al proximo
    rs.MoveNext
      
  Wend
  
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
  
  'cambia apariencia grilla
  blnB = Me.dsiCambiaApariencia()
    
  'set cantidad columnas filtradas
  m_dsiColumnasFiltradasCant = 0
    
  'get columnas
  blnB = Me.dsiGetColumnas()
        
  'get ultimo filtro
  strColumnas = keyIniToArray(Me.dsiIDopc, "columnasFiltradas")
  strOperadores = keyIniToArray(Me.dsiIDopc, "columnasOperadores")
        
  'recorro columnas filtradas
  For intI = 1 To UBound(strColumnas)
    
    'agrego columnas
    blnB = Me.dsiColumnasFiltradasAgregar(strColumnas(intI))
    
    'agrego operadores
    blnB = Me.dsiColumnasOperadoresAgregar(strOperadores(intI))
    
    'get criterio x columna
    strCriterios = keyIniToArray(Me.dsiIDopc, strColumnas(intI))
    
    'agrego criterios
    blnB = Me.dsiColumnasCriteriosAgregar(strCriterios(1))
      
  Next
        
  'set ultimo filtro a grilla
  blnB = Me.dsiFiltroAgrilla()
        
  'cambio apariencia columna 3 - criterio
  Me.spdF.Col = 3
  Me.spdF.Row = -1
    
  'activo columna con ........ cuando no se ve todo el texto
  Me.spdF.TypeEllipses = True
  
  'activo limite de columna tipo texto a 5000
  Me.spdF.TypeMaxEditLen = 5000
    
  'setea para mostrar tooltip en las celdas donde no se ve toda la info
  Me.spdF.TextTip = TextTipFixedFocusOnly
  Me.spdF.TextTipDelay = 250
        
  'put titulo
  Me.Caption = Me.dsiTitulo
        
End Sub

Private Sub Form_Resize()
    
  spdF.Left = IIf(Me.ScaleLeft < 0, 0, Me.ScaleLeft)
  spdF.Top = IIf(Me.ScaleTop < 0, 0, Me.ScaleTop)
  spdF.Height = IIf(Me.ScaleHeight < 0, 0, Me.ScaleHeight)
  spdF.Width = IIf(Me.ScaleWidth < 0, 0, Me.ScaleWidth)
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  Dim intI As Integer
  Dim strT As String
  Dim blnB As Boolean
  Dim strColumnas As String
  
  'save ubicación frm
  strValor = "left;" & Me.Left & ";top;" & Me.Top & ";width;" & Me.Width & ";height;" & Me.Height
  blnB = WriteIni(Me.dsiIDopc, "ubicacion", strValor)
  
  'armo string con columnas filtradas
  strT = ""
  For intI = 1 To Me.dsiColumnasFiltradasCant()
    strT = strT & Me.dsiColumnasFiltradas(intI) & ";"
  Next
  
  'elimino ultima coma
  If strT <> "" Then
    strT = Left(strT, Len(strT) - 1)
  End If
    
  'save columnas filtradas
  blnB = WriteIni(Me.dsiIDopc, "columnasFiltradas", strT)
  
  'armo string con columnas operadores
  strT = ""
  For intI = 1 To Me.dsiColumnasFiltradasCant()
    strT = strT & Me.dsiColumnasOperadores(intI) & ";"
  Next
  
  'elimino ultima coma
  If strT <> "" Then
    strT = Left(strT, Len(strT) - 1)
  End If
    
  'save columnas filtradas
  blnB = WriteIni(Me.dsiIDopc, "columnasOperadores", strT)
  
  'armo string con columnas filtradas
  For intI = 1 To Me.dsiColumnasFiltradasCant()
    'save columnas criterios
    blnB = WriteIni(Me.dsiIDopc, Me.dsiColumnasFiltradas(intI), Me.dsiColumnasCriterios(intI))
  Next
    
  'set herramientas
  blnB = Me.dsiHerramientasFiltro(False)
  
End Sub

Public Property Get dsiGetColumnas() As Boolean
  
  Dim blnB As Boolean
  Dim sngAnchoMaximo As Single
  Dim rs As ADODB.Recordset
  Dim fld As ADODB.Field
  Dim strVista As String
    
  'cambio puntero mouse
  Screen.MousePointer = vbHourglass
  
  'get vista
  strVista = dsiConfigGetValor("vista")
  
  'get rs
  strSQL = "select top 1 * from " & strVista
  Set rs = SQLexec(strSQL)
    
  'check error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    Exit Sub
  End If
    
  'add columnas a grilla
  For Each fld In rs.Fields
    
    'add fila a grilla
    Me.spdF.MaxRows = Me.spdF.MaxRows + 1
        
    'put texto
    Me.spdF.SetText 1, Me.spdF.MaxRows, fld.Name
                    
    'put tipo de dato
    Me.spdF.SetText 5, Me.spdF.MaxRows, fld.Type
                    
    'set puntero a fila y columna operador
    Me.spdF.Row = Me.spdF.MaxRows
    Me.spdF.Col = 2
                
    'put operadores segun tipo de dato
    Select Case fld.Type
    
    Case conChar, conNchar, conVarchar, conText, conNVarchar, conNtext
      Me.spdF.TypeComboBoxList = "<elimina>" & vbTab & "=" & vbTab & "<>"
    
    Case conDateTime, conSmallDateTime, conSmallInt, conInt, conTinyInt, conReal, conFloat, conMoney, conSmallMoney, conNumeric, conDecimal
      Me.spdF.TypeComboBoxList = "<elimina>" & vbTab & "=" & vbTab & "<>" & vbTab & "[]" & vbTab & "<" & vbTab & "<=" & vbTab & ">" & vbTab & ">="
    
    Case conBit
      Me.spdF.TypeComboBoxList = "<elimina>" & vbTab & "="
    
    End Select
            
    'default =
    Me.spdF.TypeComboBoxCurSel = 1
            
    'set puntero a columna criterio
    Me.spdF.Col = 3
    
    'set maximo string
    Me.spdF.TypeMaxEditLen = 5000
  
  Next
  
  'cierro
  SQLclose
      
  'set ancho columna
  Me.spdF.ColWidth(1) = Me.spdF.MaxTextColWidth(1) + 50
          
  'set ancho columna
  Me.spdF.ColWidth(2) = Me.spdF.MaxTextColWidth(2) + 300
          
  'recupero puntero mouse
  Screen.MousePointer = vbDefault
    
End Property

Public Property Get dsiCondicionGuardar() As Boolean
    
  Dim blnB As Boolean
  Dim strT As String
  Dim intI As Integer
  Dim strColumnas As String
    
  'set filtro para archivos *.sql
  strT = "Archivos de texto (" & Me.dsiIDopc & "*.sql)|" & Me.dsiIDopc & "*.sql"
  Me.comDestino.Filter = strT
    
  'titulo de ventana
  Me.comDestino.DialogTitle = "Guardar condición general..."
  
  'set nombre de filtro, cambio 01/07/2005
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
'  blnB = WriteIni("condicion", Me.dsiIDopc, Me.dsi, strT)
  
  'armo string con columnas filtradas
  strT = ""
  For intI = 1 To Me.dsiColumnasFiltradasCant()
    strT = strT & Me.dsiColumnasFiltradas(intI) & ";"
  Next
  
  'elimino ultima coma
  If strT <> "" Then
    strT = Left(strT, Len(strT) - 1)
  End If
    
  'save columnas filtradas
  blnB = WriteIni(Me.dsiIDopc, "columnasFiltradas", strT)
  
  'armo string con columnas operadores
  strT = ""
  For intI = 1 To Me.dsiColumnasFiltradasCant()
    strT = strT & Me.dsiColumnasOperadores(intI) & ";"
  Next
  
  'elimino ultima coma
  If strT <> "" Then
    strT = Left(strT, Len(strT) - 1)
  End If
    
  'save columnas filtradas
  blnB = WriteIni(Me.dsiIDopc, "columnasOperadores", strT)
  
  'armo string con columnas filtradas
  For intI = 1 To Me.dsiColumnasFiltradasCant()
    'save columnas criterios
    blnB = WriteIni(Me.dsiIDopc, Me.dsiColumnasFiltradas(intI), Me.dsiColumnasCriterios(intI))
  Next
  
'  'set foco
'  If Me.spdE.Visible Then
'    blnB = Me.dsiFormularioDatosPutPaF()
'  End If
  
End Property


Private Sub lstLista_ItemCheck(Item As Integer)
    
  'agrego o elimino segun click o unClick
  blnB = Me.dsiColumnasItemAgregar(Item, Me.lstLista.Selected(Item))
  
End Sub

Private Sub spdF_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
     
  Dim blnB As Boolean
  
  'set disable grilla
  Me.spdF.Enabled = False
  
  'move frame
  Me.frame.Move (Me.spdF.Left + Me.spdF.ColWidth(0) + Me.spdF.ColWidth(1) + Me.spdF.ColWidth(2)) + 50, (Me.spdF.Top + Me.spdF.RowHeight(0)) + 20, Me.spdF.ColWidth(3), (MainMDI.ScaleHeight / 2) + 250
      
  'move lista
  Me.lstLista.Move Me.spdF.Left, Me.spdF.Top, Me.frame.Width - 10, (MainMDI.ScaleHeight / 2)
  
  'move botones aceptar cancelar
  Me.cmdCancelar.Move (Me.spdF.Left - 70) + Me.frame.Width - Me.cmdCancelar.Width, (Me.frame.Height - 270)
  Me.cmdAceptar.Move (Me.spdF.Left - 70) + Me.frame.Width - Me.cmdCancelar.Width - Me.cmdAceptar.Width - 50, (Me.frame.Height - 270)
  
  'lleno grilla
  blnB = dsiGetDatos()
  
  'set visible lista
  Me.frame.Visible = True
  
End Sub

Private Sub spdF_ComboSelChange(ByVal Col As Long, ByVal Row As Long)

  Dim varOperador As Variant
  
  'get oeprador
  Me.spdF.GetText Col, Row, varOperador
  
  'si selecciono eliminar, elimino condicion
  If LCase(varOperador) = "<elimina>" Then
    
    'set puntero a fila columna
    Me.spdF.Row = Row
    Me.spdF.Col = Col
    
    'set desSelecciono comboBox
    Me.spdF.TypeComboBoxCurSel = 1
    
    'elimino criterio
    Me.spdF.SetText 3, Row, ""
    
  End If

End Sub
