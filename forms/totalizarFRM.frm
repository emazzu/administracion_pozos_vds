VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpspr60.ocx"
Begin VB.Form totalizarFRM 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Totalizando..."
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4380
      Top             =   1590
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "totalizarFRM.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "totalizarFRM.frx":077A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "totalizarFRM.frx":0EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "totalizarFRM.frx":166E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "totalizarFRM.frx":1A08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "totalizarFRM.frx":2182
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   847
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "diseno"
            Object.ToolTipText     =   "Visualizar Diseño"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "query"
            Object.ToolTipText     =   "Visualizar Query"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "grilla"
            Object.ToolTipText     =   "Visualizar Grilla"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "totalizar"
            Object.ToolTipText     =   "Totalizar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copiar"
            Object.ToolTipText     =   "Copiar Seleccion"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin FPSpreadADO.fpSpread spdRS 
      Height          =   1485
      Left            =   600
      TabIndex        =   1
      Top             =   3570
      Visible         =   0   'False
      Width           =   3435
      _Version        =   393216
      _ExtentX        =   6059
      _ExtentY        =   2619
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
      SpreadDesigner  =   "totalizarFRM.frx":28FC
   End
   Begin VB.TextBox txtQuery 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2550
      Visible         =   0   'False
      Width           =   3435
   End
   Begin FPSpreadADO.fpSpread spdCon 
      Height          =   1560
      Left            =   30
      TabIndex        =   0
      Top             =   780
      Width           =   3525
      _Version        =   393216
      _ExtentX        =   6218
      _ExtentY        =   2752
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
      SpreadDesigner  =   "totalizarFRM.frx":2AD0
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccionar columnas para Query"
      Height          =   225
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Width           =   2535
   End
End
Attribute VB_Name = "totalizarFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_gridFrm As gridFRM


Public Property Set dsiGridFrm(frm As gridFRM)
  Set m_gridFrm = frm
End Property
    
Public Property Get dsiGridFrm() As gridFRM
  Set dsiGridFrm = m_gridFrm
End Property

Public Property Get dsiGrillaConfig() As Boolean
  
  Dim strT, strF, strO As String
  Dim intCan, intCol As Integer
  
  Me.spdCon.SetText 0, 1, "Columna"
  Me.spdCon.SetText 0, 2, "Nombre"
  Me.spdCon.SetText 0, 3, "Funcion"
  Me.spdCon.SetText 0, 4, "Total"
  Me.spdCon.SetText 0, 5, "Orden"
  
  'add opcion eliminar
  strT = "(eliminar columna)" & Chr(9)
  
  'inicializo
  intCan = 0
  
  'while columnas
  For Each fld In Me.dsiGridFrm.dsiEsquema
    
    'add nombre de columnas a string
    strT = strT & "[" & fld.Name & "]" & Chr(9)
    
    'add acumulador de columnas
    intCan = intCan + 1
    
  Next
        
  'set limites grilla
'  Me.spdCon.MaxRows = 5
  Me.spdCon.MaxCols = intCan
        
  'add nombre de funciones a string
  strF = "Group By" & Chr(9) & "Sum" & Chr(9) & "Avg" & Chr(9) & "Min" & Chr(9) & "Max" & Chr(9) & "Count"
      
  'add nombre de orden a string
  strO = "Asc" & Chr(9) & "Desc" & Chr(9) & "(sin orden)"
      
  'while cantidad columnas
  For intCol = 1 To intCan
        
    'set puntero grilla en fila columna nombres
    Me.spdCon.row = 1
    Me.spdCon.col = intCol
    Me.spdCon.ColWidth(intCol) = 1300
      
    'set tipo de celda comboBox
    Me.spdCon.CellType = CellTypeComboBox
    Me.spdCon.TypeComboBoxWidth = 1
        
    'put valores a comboBox
    Me.spdCon.TypeComboBoxList = strT
    
    'set puntero grilla en fila columna funcion
    Me.spdCon.row = 3
    Me.spdCon.col = intCol
    
    'set tipo de celda comboBox
    Me.spdCon.CellType = CellTypeComboBox
    Me.spdCon.TypeComboBoxEditable = True
    Me.spdCon.TypeComboBoxWidth = 1
    
    'set puntero grilla en fila columna total
    Me.spdCon.row = 4
    Me.spdCon.col = intCol
    
    'set tipo de celda comboBox
    Me.spdCon.CellType = CellTypeComboBox
    
    'put valores a comboBox
    Me.spdCon.TypeComboBoxList = strF
    
    'set puntero grilla en fila columna orden
    Me.spdCon.row = 5
    Me.spdCon.col = intCol
    
    'set tipo de celda comboBox
    Me.spdCon.CellType = CellTypeComboBox
    
    'put valores a comboBox
    Me.spdCon.TypeComboBoxList = strO
    
  Next
        
End Property

'cambiar apariencia a grilla
Public Property Get dsiCambiaApariencia() As Boolean
  
  Me.spdRS.UnitType = UnitTypeTwips                  'trabajar en twips
  
  Me.spdRS.Appearance = Appearance3D                 'apariencia 3D
  Me.spdRS.BorderStyle = BorderStyleFixedSingle      'tipo de borde: sin borde
  
  Me.spdRS.ColHeadersAutoText = DispBlank            'titulos de columnas en blanco
  Me.spdRS.ColHeadersShow = True                     'muestra encabezado columnas
  Me.spdRS.RowHeadersShow = True                     'muestra encabezado de filas
  
  Me.spdRS.CursorStyle = CursorStyleArrow            'stilo cursor
  Me.spdRS.CursorType = CursorTypeDefault            'tipo cursor
  
  Me.spdRS.AutoSize = False                          'automaticamente ajusta ancho grilla
  Me.spdRS.DAutoSizeCols = DAutoSizeColsNone         'tipo de ajuste 2: al dato mas ancho
  
  Me.spdRS.UserColAction = UserColActionDefault      'cuando hace click en header pinta columna o fila
  Me.spdRS.FontSize = 9                              'tamaño letra
  Me.spdRS.RowHeight(0) = 450                        'altura fila de titulos
  Me.spdRS.MoveActiveOnFocus = False                 '
  Me.spdRS.Protect = False                           'exporta a excel sin proteccion
  
  Me.spdRS.BackColorStyle = BackColorStyleUnderGrid  'estilo
  Me.spdRS.GridShowHoriz = True                      'muestra grilla horizontal
  Me.spdRS.GridShowVert = True                       'muestra grilla vertical
  Me.spdRS.GridColor = RGB(200, 200, 200)            'color muy suave
  Me.spdRS.NoBorder = True                           'sin borde fin zona de datos
  
  Me.spdRS.ScrollBars = ScrollBarsBoth               'ambas barras de desplazamiento
  Me.spdRS.ScrollBarExtMode = True                   'cuando sean necesarias
  Me.spdRS.VScrollSpecial = False                     'barra especial
      
  Me.spdRS.SetOddEvenRowColor RGB(245, 245, 245), RGB(60, 60, 60), RGB(245, 245, 245), RGB(60, 60, 60)
  Me.spdRS.SelBackColor = RGB(204, 230, 255)         'fondo del area seleccionada
'  Me.spdRS.GrayAreaBackColor = RGB(245, 245, 245)
  
  Me.spdRS.VirtualMode = False                        ' ajusta rows al tamaño del recordset
  'spd.VirtualRows = 300                         ' rows a leer del virtual buffer
  'spd.VirtualScrollBuffer = True                ' scroll vertical lee de tantas rows del buffer
  
  'setea para mostrar tooltip en las celdas donde no se ve toda la info
  Me.spdRS.TextTip = TextTipFixedFocusOnly
  Me.spdRS.TextTipDelay = 250
  
End Property



Private Sub Form_Load()
  
  Dim fld As Field
  Dim blnB As Boolean
        
  'referencio grilla
  Set Me.dsiGridFrm = MainMDI.ActiveForm
      
  'pongo titulo a formulario
  Me.Caption = "Totalizando " & Me.dsiGridFrm.Caption & "..."
      
  'cambia apariencia
  spdCon.Appearance = Appearance3D
  spdCon.BorderStyle = BorderStyleFixedSingle
  spdCon.RowHeadersAutoText = DispBlank
  
  'grilla en twips
  spdCon.UnitType = UnitTypeTwips
    
  'header muy finito
  spdCon.RowHeight(0) = 50
  spdCon.ColWidth(0) = 800
            
  'pinta fila activa
'  spdCon.OperationMode = OperationModeRow
            
  'muestra lineas horizontales y verticales
  spdCon.BackColorStyle = BackColorStyleUnderGrid
  spdCon.GridShowHoriz = True
  spdCon.GridShowVert = True
  spdCon.GridColor = RGB(200, 200, 200)
          
  'set barra desplazamiento vertical
  spdCon.ScrollBars = ScrollBarsHorizontal
            
  'muestra combo cuando la celda esta activa
  spdCon.ButtonDrawMode = 1 'cell only
            
  'set color
  blnB = spdCon.SetOddEvenRowColor(RGB(245, 245, 245), RGB(60, 60, 60), RGB(245, 245, 245), RGB(60, 60, 60))
  Me.spdCon.SelBackColor = RGB(204, 230, 255)
            
  'change apariencia
  blnB = Me.dsiCambiaApariencia()
            
  'config grilla
  blnB = Me.dsiGrillaConfig()
  
End Sub

Private Sub Form_Resize()
  
  'check si grilla config visible, entonces set ancho y alto
  If Me.spdCon.Visible = True Then
    Me.spdCon.Top = Me.Toolbar1.Height + 50
    Me.spdCon.Left = 50
    Me.spdCon.Width = IIf(Me.ScaleWidth - Me.spdCon.Left - 50 < 0, 0, Me.ScaleWidth - Me.spdCon.Left - 50)
    Me.spdCon.Height = IIf(Me.ScaleHeight - Me.spdCon.Top - 50 < 0, 0, Me.ScaleHeight - Me.spdCon.Top - 50)
  End If
  
  'check si grilla rs visible, entonces set ancho y alto
  If Me.spdRS.Visible = True Then
    Me.spdRS.Top = Me.Toolbar1.Height + 50
    Me.spdRS.Left = 50
    Me.spdRS.Width = IIf(Me.ScaleWidth - Me.spdRS.Left - 50 < 0, 0, Me.ScaleWidth - Me.spdRS.Left - 50)
    Me.spdRS.Height = IIf(Me.ScaleHeight - Me.spdRS.Top - 50 < 0, 0, Me.ScaleHeight - Me.spdRS.Top - 50)
  End If
  
  'check si query visible, entonces set ancho y alto
  If Me.txtQuery.Visible = True Then
    Me.txtQuery.Top = Me.Toolbar1.Height + 50
    Me.txtQuery.Left = 50
    Me.txtQuery.Width = IIf(Me.ScaleWidth - Me.txtQuery.Left - 50 < 0, 0, Me.ScaleWidth - Me.txtQuery.Left - 50)
    Me.txtQuery.Height = IIf(Me.ScaleHeight - Me.txtQuery.Top - 50 < 0, 0, Me.ScaleHeight - Me.txtQuery.Top - 50)
  End If
  
End Sub


Private Sub spdCon_ComboSelChange(ByVal col As Long, ByVal row As Long)

  Dim varNombre, varTotal As Variant
  Dim strN, strD, strT As String
    
  Select Case row
  
  'si click en nombre de columna, build columna
  Case 1
    
    'get nombre columna
    Me.spdCon.GetText col, 1, varNombre
        
    'check si (eliminar columna), elimino
    If varNombre = "(eliminar columna)" Then
      
      Me.spdCon.DeleteCols col, 1
      
    'set nombre columna, Funcion, Group by, (sin orden) default
    Else
      
      Me.spdCon.SetText col, 2, varNombre
      Me.spdCon.SetText col, 4, "Group By"
      Me.spdCon.SetText col, 5, "(sin orden)"
      
      strN = "" & _
      "Abs (@columna)" & Chr(9) & "Exp (@columna)" & Chr(9) & "Power (@columna, 2)" & Chr(9) & _
      "Round (@columna, 3, 0)" & Chr(9) & "Sqrt (@columna)"
           
      strT = "" & _
      "Left (@columna, 1)" & Chr(9) & "Right (@columna, 1)" & Chr(9) & _
      "Substring (@columna, 1, 1)"
      
      strD = "" & _
      "Day (@columna)" & Chr(9) & "Month (@columna)" & Chr(9) & "Year (@columna)" & Chr(9) & _
      "Convert (nvarchar(6), @columna, 112)" & Chr(9) & "Convert (nvarchar(7), @columna, 111)" & Chr(9) & _
      "Convert (nvarchar, @columna - day(dateadd(m, 1, @columna)), 111)" & Chr(9) & _
      "Datepart (qq, @columna)" & Chr(9) & _
      "Datepart (wk, @columna)" & Chr(9) & "Datepart (dw, @columna)"
          
      'set puntero grilla en fila columna funcion
      Me.spdCon.row = 3
      Me.spdCon.col = col
      
      'put valores a comboBox
      Me.spdCon.TypeComboBoxList = Replace(strN & Chr(9) & strT & Chr(9) & strD, "@columna", varNombre)
      
    End If
    
  'si click en total, set nombre columna
  Case 4
    
    Me.spdCon.GetText col, 1, varNombre
    Me.spdCon.GetText col, 4, varTotal
    Me.spdCon.SetText col, 2, "[" & varTotal & "_" & _
                              Replace(Replace(varNombre, "[", ""), "]", "") & "]"
    
  End Select
  
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  
  Dim intCol, intCan As Integer
  Dim strSelect, strFrom, strFuncion, strGroupBy, strOrderBy As String
  Dim strFiltraGen, strWhere As String
  Dim varColumna, varNombre, varFuncion, varTotal, varOrden As Variant
  Dim rs As ADODB.Recordset
  
  'check opcion seleccionada
  Select Case Button.Key
  
  Case "diseno"
        
      'set visible
      Me.txtQuery.Visible = False
                  
      'set visible
      Me.spdRS.Visible = False
      
      'set visible
      Me.spdCon.Visible = True
                              
      'fuerzo ajuste de objetos
      Call Form_Resize
        
  Case "query"
  
      'set visible
      Me.spdRS.Visible = False
      
      'set visible
      Me.spdCon.Visible = False
                              
      'set visible
      Me.txtQuery.Visible = True
                              
      'fuerzo ajuste de objetos
      Call Form_Resize
        
  Case "grilla"
  
      'set visible
      Me.spdCon.Visible = False
                              
      'set visible
      Me.txtQuery.Visible = False
                              
      'set visible
      Me.spdRS.Visible = True
                              
      'fuerzo ajuste de objetos
      Call Form_Resize
        
  Case "totalizar"
    
    'while columnas seleccionadas
    For intCol = 1 To Me.spdCon.DataColCnt
      
      'get valor de columnas
      Me.spdCon.GetText intCol, 1, varColumna
      Me.spdCon.GetText intCol, 2, varNombre
      Me.spdCon.GetText intCol, 3, varFuncion
      Me.spdCon.GetText intCol, 4, varTotal
      Me.spdCon.GetText intCol, 5, varOrden
      
      'build select----
      
      'check si existe GROUP BY
      If InStr(varTotal, "Group By") <> 0 Then
        
        'check si existe funcion
        If varFuncion <> "" Then
          strSelect = strSelect & varFuncion & " AS " & varNombre & ", "
        Else
          strSelect = strSelect & varColumna & ", "
        End If
        
      'no existe GROUP BY
      Else
        
        'check si existe funcion
        If varFuncion <> "" Then
          strSelect = strSelect & varTotal & "(" & varFuncion & ") AS " & varNombre & ", "
        Else
          strSelect = strSelect & varTotal & "(" & varColumna & ") AS " & varNombre & ", "
        End If
        
      End If
         
      'get nombre de tabla o vista
      strFrom = Me.dsiGridFrm.dsiConfigGetValor("vista")
      
      'build group by----
      
      'check si existe GROUP BY
      If InStr(varTotal, "Group By") <> 0 Then
              
        'check si existe funcion
        If varFuncion <> "" Then
          strGroupBy = strGroupBy & varFuncion & ", "
        Else
          strGroupBy = strGroupBy & varColumna & ", "
        End If
        
      End If
                
      'build order by
      If varOrden <> "(sin orden)" Then
        strOrderBy = strOrderBy & intCol & " " & varOrden & ", "
      End If
          
    Next
       
    'check si contiene algo, elimino ultima coma y add palabra reservada
    If Len(strSelect) > 0 Then
      strSelect = "SELECT " & Left(strSelect, Len(strSelect) - 2)
    End If
       
    'check si contiene algo, elimino ultima coma y add palabra reservada
    If Len(strGroupBy) > 0 Then
      strGroupBy = "GROUP BY " & Left(strGroupBy, Len(strGroupBy) - 2)
    End If
       
    'check si contiene algo, elimino ultima coma y add palabra reservada
    If Len(strOrderBy) > 0 Then
      strOrderBy = "ORDER BY " & Left(strOrderBy, Len(strOrderBy) - 2)
    End If
       
    'check si contiene algo, add palabra reservada
    If Len(strFrom) > 0 Then
      strFrom = "FROM " & strFrom
    End If
       
    'check si grilla se le debe aplicar filtro general
    strFiltraGen = Me.dsiGridFrm.dsiConfigGetValor("filtra")
    If strFiltraGen <> "" Then
  
      'primer where: si existe where General lo agrego al query
      If MainMDI.dsiWhereGen("condicion") <> "" Then
        strWhere = "WHERE " & MainMDI.dsiWhereGen("condicion")
      End If
  
    End If
  
    'segundo where: si existe where lo agrego al query
    If Me.dsiGridFrm.dsiWhere <> "" Then
    
      If strWhere <> "" Then
        strWhere = strWhere & " and " & Me.dsiGridFrm.dsiWhere
      Else
        strWhere = "WHERE " & Me.dsiGridFrm.dsiWhere
      End If
  
    End If
    
    'show query
    txtQuery = strSelect & vbCrLf & strFrom & vbCrLf & strWhere & IIf(strWhere = "", "", vbCrLf) & strGroupBy & vbCrLf & strOrderBy
    
    'check si query vacio, salgo
    If Replace(txtQuery, vbCrLf, "") = "" Then
      Exit Sub
    End If
    
    'puntero mouse reloj
    Screen.MousePointer = vbHourglass
    
    'abro rs
    Set rs = SQLexec(txtQuery)
        
    'chequeo error
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      Screen.MousePointer = vbDefault   'recupero puntero mouse
      Exit Sub
    End If
          
    'cierro
    SQLclose
    
    'check si rs existe, set rs a grilla
    If Not rs.EOF Then
    
      'set rs as grilla
      Set Me.spdRS.DataSource = rs
          
      'set limite a grilla
      Me.spdRS.MaxRows = rs.RecordCount
                  
      'set visible
      Me.spdCon.Visible = False
                  
      'set visible
      Me.txtQuery.Visible = False
                  
      'set visible
      Me.spdRS.Visible = True
                              
      'fuerzo ajuste de objetos
      Call Form_Resize

    End If
            
    'recupero puntero mouse
    Screen.MousePointer = vbDefault
    
  Case "copiar"
    
    Me.spdRS.ClipboardCopy
    
  End Select
  
End Sub
