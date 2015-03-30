VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpspr60.ocx"
Begin VB.Form filtrarFRM 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pesonalizando filtro..."
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAyuda 
      Height          =   885
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2280
      Width           =   6495
   End
   Begin VB.CheckBox chk_filtro_agregar 
      Caption         =   "Agregar condición a la existente"
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   3240
      Width           =   2715
   End
   Begin VB.Frame Frame1 
      Height          =   2145
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   6495
      Begin FPSpreadADO.fpSpread spdFiltro 
         Height          =   1965
         Left            =   30
         TabIndex        =   0
         Top             =   120
         Width           =   6405
         _Version        =   393216
         _ExtentX        =   11298
         _ExtentY        =   3466
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
         SpreadDesigner  =   "filtrarFRM.frx":0000
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   270
      Left            =   4590
      TabIndex        =   1
      Top             =   3240
      Width           =   915
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   270
      Left            =   5640
      TabIndex        =   2
      Top             =   3240
      Width           =   915
   End
End
Attribute VB_Name = "filtrarFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbsql As String
Dim querybuilt As Boolean
Dim querysaved As Boolean
Dim viewingsql As Boolean
Dim savepressed As Boolean
Dim rowdataadded As Boolean

'Variables for operator combos, depending on data type.  Limit choices.
Dim strOperTxt As String
Dim strOperBln As String
Dim strOperNum As String
Dim strOperDtm As String

'Spreadsheet column constants
Const conFieldCol = 1
Const conFieldTypeCol = 2
Const conOperCol = 3
Const conCritCol = 4
Const conAndOrCol = 5

Dim m_gridFrm As gridFRM
Dim m_EsquemaOrdenado() As String

Public Property Get dsiAyudaFiltroLimpiar() As String
  m_dsiAyudaFiltro = ""
End Property

Public Property Set dsiGridFrm(ByVal frm As gridFRM)
  Set m_gridFrm = frm
End Property

Public Property Get dsiGridFrm() As gridFRM
  Set dsiGridFrm = m_gridFrm
End Property

Public Property Get dsiCambiaApariencia() As Boolean
  
  Dim blnB As Boolean
  
  'altura header
  spdFiltro.RowHeight(0) = 13
  
  'estilo borde
  spdFiltro.BorderStyle = BorderStyleNone
    
  'reemplaza informacion si existe
  spdFiltro.EditModeReplace = True
    
  'quita barra de desplazamientos y habilita tab
  spdFiltro.ScrollBars = ScrollBarsNone
  spdFiltro.ProcessTab = True

  'esconde numeros de filas
  spdFiltro.RowHeadersShow = False
    
  'esconde lineas verticales
  spdFiltro.GridShowVert = False
  
  'habilita las barras de desplazamiento si las necesita
  spdFiltro.ScrollBarExtMode = True
      
  Me.Frame1.BackColor = RGB(216, 237, 223)    'fondo frame
  Me.BackColor = RGB(216, 237, 223)           'fondo formulario
  Me.txtAyuda.BackColor = RGB(245, 245, 245)  'fondo ayuda
  Me.txtAyuda.ForeColor = RGB(60, 60, 60)     'texto ayuda
  Me.chk_filtro_agregar.BackColor = RGB(216, 237, 223)
  
  Me.spdFiltro.ShadowText = &H404040           'texto de encabezados
  Me.spdFiltro.GrayAreaBackColor = RGB(216, 237, 223)
  Me.spdFiltro.SetOddEvenRowColor RGB(245, 245, 245), RGB(60, 60, 60), RGB(252, 252, 252), RGB(60, 60, 60)
  
  'establece numero maximo filas y columnas
  spdFiltro.MaxRows = 0
  spdFiltro.MaxCols = 5
    
  'muestra combo cuando la celda esta activa
  'spdFiltro.ButtonDrawMode = 1 'cell only
  
  spdFiltro.EditModePermanent = True
     
  'columna para nombres de columnas del esquema
  spdFiltro.row = 0
  spdFiltro.col = conFieldCol
  spdFiltro.Text = "Columna"
  spdFiltro.ColWidth(conFieldCol) = 18
    
  'columna Operadores
  spdFiltro.row = 0
  spdFiltro.col = conOperCol
  spdFiltro.Text = "Operator"
  spdFiltro.ColWidth(conOperCol) = 8
    
  'columna And Or
  spdFiltro.row = 0
  spdFiltro.col = conAndOrCol
  spdFiltro.Text = "And"
  spdFiltro.ColWidth(conAndOrCol) = 7
    
  'Populate Criteria header
  spdFiltro.row = 0
  spdFiltro.col = conCritCol
  spdFiltro.Text = "Criterio"
  spdFiltro.ColWidth(conCritCol) = 20

  'esconde columna de tipos de dato y columna string de fecha
  spdFiltro.col = conFieldTypeCol
  spdFiltro.ColHidden = True

End Property

Public Property Get dsiAgregarFila() As Boolean
  
  Dim strOper As String
  
  Static i, RowCnt As Integer
  Dim strFieldNames As String, strFieldValues As String
  Dim ret As Boolean
  Dim intI As Integer
      
  'agrego una fila y puntero a fila nueva
  Me.spdFiltro.MaxRows = Me.spdFiltro.MaxRows + 1
  Me.spdFiltro.row = Me.spdFiltro.MaxRows
  
  'set columna
  Me.spdFiltro.col = conFieldCol
      
  'como primer item del combo muestra clear row
  strFieldNames = "<Eliminar fila>" & vbTab
  strFieldValues = "<Eliminar fila>" & vbTab
  
  'lleno combo de grilla con nombres de columnas ordenadas alfabeticamente
  For intI = 0 To Me.dsiEsquemaCantidad
    strFieldNames = strFieldNames & Me.dsiEsquemaOrdenado(intI) & vbTab
    strFieldValues = strFieldValues & Me.dsiGridFrm.dsiEsquema(Me.dsiEsquemaOrdenado(intI)).Type & vbTab
  Next
    
  'set columna nombre de columnas
  spdFiltro.col = conFieldCol                    'nombre de campo
  
  spdFiltro.CellType = CellTypeComboBox          'estableciendo celda como comboBox
  spdFiltro.TypeComboBoxList = strFieldNames     'asigno lista de campo
  spdFiltro.TypeComboBoxMaxDrop = 15             'largo de la ventana
  spdFiltro.TypeComboBoxAutoSearch = 2           'busca en modo multi caracter
  spdFiltro.col = conFieldTypeCol                'tipo de dato
  spdFiltro.CellType = CellTypeComboBox          'estableciendo celda como comboBox
  spdFiltro.TypeComboBoxList = strFieldValues    'asigno lista de tipos de dato
     
  'set columna operador
  spdFiltro.col = conOperCol
  spdFiltro.CellType = CellTypeComboBox
  spdFiltro.TypeComboBoxList = OperAll

  'tipo de celda
  spdFiltro.col = conAndOrCol
  spdFiltro.CellType = CellTypeComboBox
  spdFiltro.TypeComboBoxList = "and" & vbTab & "or"
      
  'Change Edit header
  spdFiltro.col = conEditCol
  spdFiltro.row = 0
  spdFiltro.ColWidth(conEditCol) = 4.5
  spdFiltro.Text = " "
    
  'esconde columna de tipos de dato y columna string de fecha
  spdFiltro.col = conFieldTypeCol
  spdFiltro.ColHidden = True

End Property


Public Property Get dsiEsquemaOrdenar() As Boolean
    
  Dim fld As ADODB.Field
  Dim intI, intAUX1, strAUX2 As Integer
        
  intI = -1
        
  'recorro esquema y guardo en array
  For Each fld In Me.dsiGridFrm.dsiEsquema
    intI = intI + 1
    ReDim Preserve m_EsquemaOrdenado(intI)
    m_EsquemaOrdenado(intI) = fld.Name
  Next
  
  'ordeno columnas
  For intAUX1 = 0 To UBound(m_EsquemaOrdenado)
    
    For intAUX2 = intAUX1 + 1 To UBound(m_EsquemaOrdenado)
      
      If m_EsquemaOrdenado(intAUX1) >= m_EsquemaOrdenado(intAUX2) Then
        strAUX1 = m_EsquemaOrdenado(intAUX1)
        m_EsquemaOrdenado(intAUX1) = m_EsquemaOrdenado(intAUX2)
        m_EsquemaOrdenado(intAUX2) = strAUX1
      End If
    
    Next
  
  Next
  
End Property

Public Property Get dsiEsquemaOrdenado(intI) As String
  dsiEsquemaOrdenado = m_EsquemaOrdenado(intI)
End Property

Public Property Get dsiEsquemaCantidad() As Integer
  dsiEsquemaCantidad = UBound(m_EsquemaOrdenado)
End Property

Function GetCriteriaValue(fpcriteria As String, ByVal intDataType As Integer) As String
'Return the Criteria value
Dim ret As Integer

        'Make sure not a ' in the text (bad sql)
        ret = InStr(1, fpcriteria, "'")
        If ret > 0 Then
            'Found : fix by addin and extra ' so sql will understand
            GetCriteriaValue = Left(fpcriteria, ret) & "'" & Right(fpcriteria, Len(fpcriteria) - ret)
            Exit Function
        End If
        
        'Check for relative date value
        If intDataType = conSmallDateTime Or intDataType = conDateTime Then
            If fpcriteria <> "" Then
                'Relative date entered
                GetCriteriaValue = "'" & dateToIso(spdFiltro.Text) & "'"
                Exit Function
            End If
        End If
        
        GetCriteriaValue = fpcriteria
    
End Function

Function AddConditions(ByVal strField As String, ByVal strCondition As String, ByRef strCriterio As String, ByVal intDataType As Integer) As String
  Dim AddQuotes As Boolean
    
  'Do not add ' '  around anything other than text type
  Select Case intDataType
  
  Case conChar, conNchar, conVarchar, conText, conNVarchar, conNtext        ' texto
    If LCase(strCriterio) = "null" Then
      AddQuotes = False
    Else
      AddQuotes = True
    End If
  
  Case conBit                                                               ' boolean
    If strCriterio = "0" Then
      strCriterio = 0
    Else
      strCriterio = 1
    End If
    AddQuotes = False
          
  Case Else                                                                 ' otros
     AddQuotes = False
  
  End Select
    
  Select Case LCase(strCondition)
  
  Case "menor que"
    If AddQuotes Then
      AddConditions = "[" & strField & "] < '" & strCriterio & "'"
    Else
      AddConditions = "[" & strField & "] < " & strCriterio
    End If
            
  Case "menor o igual que"
    If AddQuotes Then
      AddConditions = "[" & strField & "] <= '" & strCriterio & "'"
    Else
      AddConditions = "[" & strField & "] <= " & strCriterio
    End If
            
  Case "igual"
    If AddQuotes Then
      AddConditions = "[" & strField & "] like '" & strCriterio & "'"
    Else
      AddConditions = "[" & strField & "] = " & strCriterio
    End If
            
  Case "distinto"
    If AddQuotes Then
      AddConditions = "[" & strField & "] not like '" & strCriterio & "'"
    Else
      AddConditions = "[" & strField & "] <> " & strCriterio
    End If
        
  Case "mayor que"
    If AddQuotes Then
      AddConditions = "[" & strField & "] > '" & strCriterio & "'"
    Else
      AddConditions = "[" & strField & "] > " & strCriterio
    End If
            
  Case "mayor o igual que"
    If AddQuotes Then
      AddConditions = "[" & strField & "] >= '" & strCriterio & "'"
    Else
      AddConditions = "[" & strField & "] >= " & strCriterio
    End If
        
  Case "comienza con"
    If AddQuotes Then
      AddConditions = "[" & strField & "] like '" & strCriterio & "%'"
    Else
      AddConditions = "[" & strField & "] like '" & strCriterio & "%'"
    End If
        
  Case "termina con"
    If AddQuotes Then
      AddConditions = "[" & strField & "] like '%" & strCriterio & "'"
    Else
      AddConditions = "[" & strField & "] like '%" & strCriterio & "'"
    End If
            
  Case "contiene"
    If AddQuotes Then
      AddConditions = "[" & strField & "] like '%" & strCriterio & "%'"
    Else
      AddConditions = "[" & strField & "] like '%" & strCriterio & "%'"
    End If
            
  End Select
    
  querybuilt = True
    
End Function




Private Sub cmdAceptar_Click()

  Dim blnB, blnColumnaTexto As Boolean
  Dim intI, intIcant As Integer
  Dim strParcial, strFinal, strColumna As String
  Dim varColumna, varOperador, varCriterio, varTipoColumna, arrCriterio As Variant
  
  strFinal = ""                           'vacio condicion final
  
  'recorro filas
  For intI = 1 To Me.spdFiltro.DataRowCnt
    
    'get columna, operador, criterio
    Me.spdFiltro.GetText 1, intI, varColumna
    Me.spdFiltro.GetText 2, intI, varTipoColumna
    Me.spdFiltro.GetText 3, intI, varOperador
    Me.spdFiltro.GetText 4, intI, varCriterio
        
    'check si fila completa
    If varColumna <> "" And varOperador <> "" And varCriterio <> "" Then
      
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
      
      'put "menor que" con 1 solo criterio, OK, sino, error
      If LCase(varOperador) = "<" And UBound(arrCriterio) = 1 Then
        strParcial = "[" & varColumna & "]" & " < " & " " & strParcial
        strFinal = strFinal & strParcial & " and "
      End If
      
      'put "menor o igual que" con 1 solo criterio, OK, sino, error
      If LCase(varOperador) = "<=" And UBound(arrCriterio) = 1 Then
        strParcial = "[" & varColumna & "]" & " <= " & strParcial
        strFinal = strFinal & strParcial & " and "
      End If
      
      'put "mayor que" con 1 solo criterio, OK, sino, error
      If LCase(varOperador) = ">" And UBound(arrCriterio) = 1 Then
        strParcial = "[" & varColumna & "]" & " > " & strParcial
        strFinal = strFinal & strParcial & " and "
      End If
      
      'put "mayor o igual que" con 1 solo criterio, OK, sino, error
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
  
  'guardo valor antes de bajar frm
  blnAgregarAcondActual = Me.chk_filtro_agregar
  
  'descargo frm
  Unload Me
  
  'si existe filtro, aplico
  If strFinal <> "" Then
    
    'si no se hizo click en agregar a condición a la existente, elimino filtro actual
    If blnAgregarAcondActual = 0 Then
      blnB = Me.dsiGridFrm.dsiCondicionEliminar()
    End If
    
    'set agrego condicion
    Me.dsiGridFrm.dsiCondicionAgregar = strFinal
    
    'muestro en barra de estado
    blnB = Me.dsiGridFrm.dsiWhereMostrarSi()
    
    'aplica filtro
    blnB = Me.dsiGridFrm.dsiRefresh()
    
  End If
  
  blnAceptar = True
  blnCancelar = False
  
End Sub

Private Sub cmdCancelar_Click()
  
  blnAceptar = False
  blnCancelar = True
  
  Unload Me

End Sub

Private Sub Form_Load()
       
  Dim blnB As Boolean
       
  'referencio propiedad necesaria con esquema
  Set Me.dsiGridFrm = MainMDI.ActiveForm
           
  'pongo titulo a formulario
  Me.Caption = "Filtrando " & Me.dsiGridFrm.Caption & "..."
           
  'guardo esquema ordenado
  blnB = Me.dsiEsquemaOrdenar()
           
  'cambia apariencia grilla
  blnB = Me.dsiCambiaApariencia()
    
  'agrega primer fila
  blnB = Me.dsiAgregarFila()
    
End Sub


Private Sub spdFiltro_ComboCloseUp(ByVal col As Long, ByVal row As Long, ByVal SelChange As Integer)
  
  'set fila columna seleccionada
  Me.spdFiltro.row = row
  Me.spdFiltro.row = col
  
  'si selecciono And o Or y es la ultima fila, agrego otra condición
  If col = conAndOrCol And Me.spdFiltro.ActiveRow = Me.spdFiltro.MaxRows Then
    
    blnB = Me.dsiAgregarFila()
    Exit Sub
    
  End If
  
  'si selecciona <Eliminar fila>
  If col = conFieldCol And SelChange = 0 Then
      
    'si hay una sola fila, limpio, sino elimino
    If Me.spdFiltro.MaxRows = 1 Then
                  
      'limpio fila
      Me.spdFiltro.ClearRange 2, row, -1, row, True
      
    Else
      
      If row = Me.spdFiltro.MaxRows Then
          
          'set fila anterior como activa
          Me.spdFiltro.SetActiveCell col, row - 1
      Else
      
          'set fila siguiente como activa
          Me.spdFiltro.SetActiveCell col, row + 1
          
      End If
      
      'elimino fila
      Me.spdFiltro.DeleteRows row, 1
      Me.spdFiltro.MaxRows = Me.spdFiltro.MaxRows - 1
          
    End If
    
    Exit Sub
  
  End If

End Sub

Private Sub spdFiltro_ComboSelChange(ByVal col As Long, ByVal row As Long)
  
  Dim intTipo As Integer
  Dim blnB As Boolean
  Dim strT As String
    
  'set fila columna actual
  Me.spdFiltro.row = row
  Me.spdFiltro.col = col
  
  'si seleccion de operador, display help
  If col = conOperCol Then
        
    Select Case Me.spdFiltro.Text
          
    Case "=", "<>"
      Me.txtAyuda = "Usted puede tipear una lista de valores, ejemplo: valor1,valor2,valor3,valor4,etc. " & _
                    "Para columnas de tipo texto, junto a la lista de valores se pueden utilizar comodines, ejemplo: valor1% para texto que comience con valor1, " & _
                    "%valor1% para texto que contenga valor1 y %valor1 para texto que finalice valor1."
    
    Case "[]"
      Me.txtAyuda = "Operador (entre), usted puede tipear solo 2 valores, ejemplo: valor1,valor2."
      
    Case "<", "<=", ">", ">="
      Me.txtAyuda = "Usted puede tipear tipear solo un valor, ejemplo: valor1."
      
    End Select
        
  End If
    
  'solo set operadores y criterios cuando se selecciona una columna
  If col <> conFieldCol Then
    Exit Sub
  End If
  
'  'set fila columna actual
'  Me.spdFiltro.row = row
'  Me.spdFiltro.col = col
  
  'si selecciono eliminar fila, salgo
  If Me.spdFiltro.Text = "<Eliminar fila>" Then
    Exit Sub
  End If
  
  'get tipo de dato de columna seleccionada
  intTipo = Me.dsiGridFrm.dsiEsquema(Me.spdFiltro.Text).Type
    
  'cambio tipo de celda para ingresar criterio
  Select Case intTipo
    
  Case conChar, conNchar, conVarchar, conText, conNVarchar, conNtext
        
    strT = "=" & vbTab & "<>"
    
  Case conDateTime, conSmallDateTime
    
    strT = "=" & vbTab & "<>" & vbTab & "[]" & vbTab & "<" & vbTab & "<=" & vbTab & ">" & vbTab & ">="
    
  Case conSmallInt, conInt, conTinyInt, conReal, conFloat, conMoney, conSmallMoney, conNumeric, conDecimal
    
    strT = "=" & vbTab & "<>" & vbTab & "[]" & vbTab & "<" & vbTab & "<=" & vbTab & ">" & vbTab & ">="
    
  Case conBit
    
    strT = "="
  
  End Select
  
  'put tipo de dato
  Me.spdFiltro.col = conFieldTypeCol
  Me.spdFiltro.Text = intTipo
  
  'put posibles operadores segun tipo de dato, cambio 30/06/2005
  Me.spdFiltro.col = conOperCol
  Me.spdFiltro.CellType = CellTypeComboBox
  Me.spdFiltro.TypeComboBoxEditable = False
  Me.spdFiltro.TypeComboBoxList = strT
    
  'set tipo celda texto para criterio, cambio 30/06/2005
  Me.spdFiltro.col = conCritCol
  Me.spdFiltro.CellType = CellTypeEdit
    
  'put condicion and
  Me.spdFiltro.col = conAndOrCol
  Me.spdFiltro.CellType = CellTypeComboBox
  Me.spdFiltro.TypeComboBoxEditable = False
  Me.spdFiltro.TypeComboBoxList = "and"                   'cambio 30/06/2005
'  Me.spdFiltro.TypeComboBoxList = "and" & Chr(9) & "or"

End Sub

