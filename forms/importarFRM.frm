VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpspr60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form importarFRM 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Importando..."
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleMode       =   0  'User
   ScaleWidth      =   9628.834
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2445
      Left            =   0
      TabIndex        =   4
      Top             =   495
      Width           =   9915
      Begin VB.TextBox txtOperacion 
         Height          =   285
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2070
         Width           =   9690
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H80000002&
         Height          =   1860
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "importarFRM.frx":0000
         Top             =   180
         Width           =   9690
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8820
      Top             =   6390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "importarFRM.frx":01F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "importarFRM.frx":0972
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "importarFRM.frx":10EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "importarFRM.frx":1866
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "importarFRM.frx":1F60
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog comDestino 
      Left            =   8325
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpreadADO.fpSpread spdD 
      Height          =   2400
      Left            =   0
      TabIndex        =   1
      Top             =   5340
      Width           =   9900
      _Version        =   393216
      _ExtentX        =   17462
      _ExtentY        =   4233
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
      SpreadDesigner  =   "importarFRM.frx":26DA
   End
   Begin FPSpreadADO.fpSpread spdO 
      Height          =   2400
      Left            =   0
      TabIndex        =   0
      Top             =   2910
      Width           =   9900
      _Version        =   393216
      _ExtentX        =   17462
      _ExtentY        =   4233
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
      SpreadDesigner  =   "importarFRM.frx":28AE
   End
   Begin MSComctlLib.Toolbar tlbHerra 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9885
      _ExtentX        =   17436
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
            Key             =   "cmd_abrir"
            Object.ToolTipText     =   "Abrir archivo de donde se van a importar los datos"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmd_copiar_auto"
            Object.ToolTipText     =   "Copiar columnas en forma automática"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmd_copiar_manual"
            Object.ToolTipText     =   "Copiar una columna en forma manual"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmd_chequear"
            Object.ToolTipText     =   "Chequear información"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmd_importar"
            Object.ToolTipText     =   "Importar información"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CommandButton Command2 
         Caption         =   "GRA"
         Height          =   255
         Left            =   11070
         TabIndex        =   3
         Top             =   30
         Visible         =   0   'False
         Width           =   495
      End
   End
End
Attribute VB_Name = "importarFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_gridFrm As gridFRM
Dim m_arrComboBoxCant As Integer
Dim m_arrComboBox() As String
Dim m_arrErrores() As String
Dim m_dsiChequeada As Boolean

Public Property Set dsiGridFrm(frm As gridFRM)
  Set m_gridFrm = frm
End Property
    
Public Property Get dsiGridFrm() As gridFRM
  Set dsiGridFrm = m_gridFrm
End Property

Public Property Get dsiComboBoxAgregar(ByVal strColumna As String, ByVal strValor As String, ByVal strID As String)
  
  'add 1 a contador de cantidad de items
  m_arrComboBoxCant = m_arrComboBoxCant + 1
  
  'redim array
  ReDim Preserve m_arrComboBox(3, m_arrComboBoxCant) As String
  
  'put datos
  m_arrComboBox(1, m_arrComboBoxCant) = strColumna 'nombre de columna
  m_arrComboBox(2, m_arrComboBoxCant) = strValor   'descripcion
  m_arrComboBox(3, m_arrComboBoxCant) = strID      'id
  
End Property

Public Property Get dsiComboBoxBuscar(ByVal strColumna As String, ByVal strValor As String, ByVal strLista As String) As String
  
  Dim blnB As Boolean
  Dim strT, strID As String
  Dim lngL As Long
  Dim rs As ADODB.Recordset
  Dim arrSel(), arrEle As Variant
  Dim intI As Integer
  
  'valor default devuelto
  dsiComboBoxBuscar = "null"
    
  'paso 1: busco en array
  For lngL = 1 To m_arrComboBoxCant
      
    'check si encontro ID
    If strColumna = m_arrComboBox(1, lngL) And strValor = m_arrComboBox(2, lngL) Then
        dsiComboBoxBuscar = m_arrComboBox(3, lngL)
        Exit For
    End If
    
  Next
      
  'paso 2: si paso 1: no encontro, busco en LISTA
  If dsiComboBoxBuscar = "null" Then
      
    'check si LISTA contiene la palabra SELECT
    'ejemplo de select: select id, nombre from maeAreas
    'ejemplo de valores a mano: PROVEN,PROVEN;PROBABLE,PROBABLE;POSSIBLE,POSSIBLE;,;
    If InStr(LCase(strLista), "select") Then
      
      'paso1: descompongo string en array, tomo en cuanta el espacio, queda un array igual a la linea abajo
      '(1):select (2):id, (3):nombre (4):from (5):maeAreas (6):order (7):by (8):nombre
      arrSel = separateText(strLista, " ")
      
      'armo string para buscar descripcion en la tabla
      strT = arrSel(1) & " " & Replace(arrSel(2), ",", "") & " from " & arrSel(5) & " where " & arrSel(3) & " = '" & strValor & "'"
      
      'busco en tabla
      Set rs = SQLexec(strT)
      
      'chequeo errores
      If Not SQLparam.CnErrNumero = -1 Then
        SQLError
        SQLclose
        Exit Property
      End If
                  
      'check si encontro algo
      If (Not rs.EOF) Then
        strID = rs(0)
      Else
        strID = "null"
      End If
                              
    Else
      
      'paso1: descompongo string en array, tomo en cuanta el punto y coma, queda un array igual a la linea abajo
      '(1):PROVEN,PROVEN (2):PROBABLE,PROBABLE, (3):POSSIBLE,POSSIBLE (4):, etc.
      arrSel = separateText(strLista, ";")
      
      'while array
      For intI = 1 To UBound(arrSel)
        
        'descompongo en dato y id, (1):PROVEN (2):PROVEN
        arrEle = separateText(arrSel(intI), ",")
        
        'check si encontro
        If InStr(arrEle(1), strValor) <> 0 Then
          strID = arrEle(2)
          Exit For
        Else
          strID = "null"
        End If
        
      Next
      
    End If
      
    'check si no es numerico, agrego comillas simples
    If Not IsNumeric(strID) Then
      strID = "'" & strID & "'"
    End If
    
    'save devulevo ID
    dsiComboBoxBuscar = strID
      
    'agrego ID encontrado al array, asi la próxima vez no tienen que buscar en la tabla
    blnB = Me.dsiComboBoxAgregar(strColumna, strValor, strID)
      
  End If
    
End Property

Public Property Get dsiSpdAparienciaO() As Boolean

  'cambia apariencia
  spdO.BorderStyle = BorderStyleNone
  
  'grilla en twips
  spdO.UnitType = UnitTypeTwips
  
  'header muy finito
  spdO.RowHeight(0) = 80
            
  'muestra lineas horizontales y verticales
  spdO.BackColorStyle = BackColorStyleUnderGrid
  spdO.GridShowHoriz = True
  spdO.GridShowVert = True
  spdO.GridColor = RGB(200, 200, 200)
          
  'set color
  blnB = spdO.SetOddEvenRowColor(RGB(245, 245, 245), RGB(60, 60, 60), RGB(245, 245, 245), RGB(60, 60, 60))
  Me.spdO.SelBackColor = RGB(220, 220, 220)
  
  'set color area fuera de grilla
  Me.spdO.GrayAreaBackColor = RGB(245, 245, 245)
    
  'congelo fila 1 para nombres de columna
  Me.spdO.RowsFrozen = 1
  
End Property

Public Property Get dsiSpdAparienciaD() As Boolean

  'cambia apariencia
  spdD.BorderStyle = BorderStyleNone
  
  'grilla en twips
  spdD.UnitType = UnitTypeTwips
  
  'header muy finito
  spdD.RowHeight(0) = 80
            
  'muestra lineas horizontales y verticales
  spdD.BackColorStyle = BackColorStyleUnderGrid
  spdD.GridShowHoriz = True
  spdD.GridShowVert = True
  spdD.GridColor = RGB(200, 200, 200)
          
  'set color
  blnB = spdD.SetOddEvenRowColor(RGB(245, 245, 245), RGB(60, 60, 60), RGB(245, 245, 245), RGB(60, 60, 60))
  Me.spdD.SelBackColor = RGB(220, 220, 220)
  
  'set color area fuera de grilla
  Me.spdD.GrayAreaBackColor = RGB(245, 245, 245)
    
  'congelo fila 1 para nombres de columna
  Me.spdD.RowsFrozen = 2
    
End Property

Public Property Get dsiColumnasConfig() As Boolean
  
  Dim intI As Integer
  Dim strT, strValor As String
  
  'set limites grilla destino
  'columna 1 reservada para mostrar el estado de la importacion de cada fila
  Me.spdD.MaxRows = Me.spdO.DataRowCnt + 1
  Me.spdD.MaxCols = Me.dsiGridFrm.dsiEsquema.Count + 1
    
  'dim array para errores de importacion
  ReDim m_arrErrores(Me.spdD.MaxRows)
    
  'set tipo edit, centrado, letra bold, bloqueado
  Me.spdD.col = 1
  Me.spdD.row = -1
  Me.spdD.CellType = CellTypeEdit
  Me.spdD.TypeHAlign = TypeHAlignCenter
  Me.spdD.FontBold = True
  Me.spdD.Lock = True
        
  'recorro esquema
  For intI = 0 To Me.dsiGridFrm.dsiEsquema.Count - 1
            
    'set alcance para change tipo de dato
    Me.spdD.col = intI + 2
    Me.spdD.row = -1
                
    'check tipo de celda
    Select Case Me.dsiGridFrm.dsiEsquema(intI).Type
          
    'Bit
    Case conBit
      Me.spdD.CellType = CellTypeCheckBox
      
      strT = "(0/1)"
          
    'enteros
    Case conSmallInt, conInt, conTinyInt
      Me.spdD.CellType = CellTypeNumber
      Me.spdD.TypeNumberDecPlaces = 0
      
      strT = "(entero)"
                    
    'decimal
    Case conMoney, conSmallMoney, conReal, conFloat, conNumeric, conDecimal
      Me.spdD.CellType = CellTypeNumber
      Me.spdD.TypeNumberDecPlaces = 2
      
      strT = "(decimal)"
            
    'fecha
    Case conSmallDateTime, conDateTime
      Me.spdD.CellType = CellTypeDate
      Me.spdD.TypeDateFormat = TypeDateFormatDDMMYY
      
      strT = "(dd/mm/yyyy)"
        
    'string
    Case conChar, conNchar, conVarchar, conText, conNVarchar, conNtext
      Me.spdD.CellType = CellTypeEdit
              
      'check si ntext
      If Me.dsiGridFrm.dsiEsquema(intI).Type = conNtext Then
        Me.spdD.TypeMaxEditLen = 5000
        strT = "(texto) 5000"
      Else
        Me.spdD.TypeMaxEditLen = Me.dsiGridFrm.dsiEsquema(intI).DefinedSize
        strT = "(texto) " & Me.dsiGridFrm.dsiEsquema(intI).DefinedSize
      End If
        
    End Select
              
    'set lock toda la grilla para que no se puedan modifcar datos a mano
    Me.spdD.Lock = True
        
    'check columna OBLIGATORIA
    strValor = Me.dsiGridFrm.dsiConfigGetValor("obligatorio", Me.dsiGridFrm.dsiEsquema(intI).Name)
      
    'si columna obligatoria, celda verde
    If strValor <> "" Then
      Me.spdD.BackColor = RGB(216, 237, 223)
    End If
        
'    'check columna BLOQUEADA
'    strValor = Me.dsiGridFrm.dsiConfigGetValor("bloqueado", Me.dsiGridFrm.dsiEsquema(intI).Name)
      
'    'si columna bloqueada, celda gris atenuado
'    If strValor <> "" Then
'      Me.spdD.ForeColor = RGB(155, 155, 155)
'    End If
        
    'set nombre de columna
    Me.spdD.row = 1
    Me.spdD.CellType = CellTypeStaticText
    Me.spdD.Text = Me.dsiGridFrm.dsiEsquema(intI).Name
      
    'set tipo de dato de columna
    Me.spdD.row = 2
    Me.spdD.CellType = CellTypeStaticText
    Me.spdD.Text = strT
        
  Next  'cuenta columnas
                        
  'set nombre de columna 1
  Me.spdD.col = 1
  Me.spdD.row = 1
  Me.spdD.CellType = CellTypeStaticText
  Me.spdD.Text = "Estado"
  Me.spdD.FontBold = False
  Me.spdD.TypeHAlign = TypeHAlignCenter
          
  'set tipo de dato columna 1
  Me.spdD.col = 1
  Me.spdD.row = 2
  Me.spdD.CellType = CellTypeStaticText
  Me.spdD.Text = ""
                        
End Property

'
'COPIA columna en forma automatica
'si nombre de columna en origen coincide en destino, copia info
'
Public Property Get dsiColumnaCopiar_auto()
  
  Dim lngColumnaO, lngColumnaD As Long
  Dim strColumnaNombre As Variant
  
  'CHECK si no filas en grilla origen, SALGO
  If Me.spdO.DataRowCnt = 0 Or Me.spdD.DataRowCnt = 0 Then
    
    blnB = MsgBox("No hay información para procesar...", vbExclamation, "atención...")
    Exit Property
    
  End If
    
  'WHILE columnas grilla origen
  For lngColumnaO = 1 To Me.spdO.DataColCnt
        
    'GET nombre de columna
    Me.spdO.GetText lngColumnaO, 1, strColumnaNombre
    
    'FIND columna en grilla destino
    lngColumnaD = Me.spdD.SearchRow(1, 1, Me.spdD.DataColCnt, strColumnaNombre, SearchFlagsNone)
      
    'CHECK si encontro columna
    If lngColumnaD > 0 Then
      
      'CALL copiar columna
      Call dsiColumnaCopiar_una(lngColumnaO, lngColumnaD)
      
    End If
      
  Next
  
End Property

'
'COPIA columna en forma manual
'
Public Property Get dsiColumnaCopiar_manual()
  
  'CHECK si no filas en grilla origen, SALGO
  If Me.spdO.DataRowCnt = 0 Or Me.spdD.DataRowCnt = 0 Then
    
    blnB = MsgBox("No hay información para procesar...", vbExclamation, "atención...")
    Exit Property
    
  End If
    
  'CALL copiar columna activa
  Call dsiColumnaCopiar_una(Me.spdO.ActiveCol, Me.spdD.ActiveCol)
   
     
End Property

Public Property Get dsiImportar()
  
  Dim lngFila As Long
  Dim intCol, intTipo, intOK, intERROR As Integer
  Dim strColumna, strLista, strID, strParcial, strFinal, strSP, strSPprimero, strSPultimo As String
  Dim strNoActualizo As String
  Dim blnB As Boolean
  Dim varColumna, varValor As Variant
    
  'CHECK si no se puede importar, SALGO
  If Not m_dsiChequeada Then
  
    blnB = MsgBox("Antes de poder importar información debe hacer clic en chequear información...", vbExclamation, "atención...")
    Exit Property
  
  End If
    
  'puntero mouse reloj
  Screen.MousePointer = vbHourglass
  
  '-------------------------------------------------------------------------------
  '04/09/2009 Si existe el tipo SPIMP_PRIMERO, ejecuta un store antes de importar
  '
  '
  'GET store procedure primera vez
  strSPprimero = Me.dsiGridFrm.dsiConfigGetValor("SPIMP_PRIMERO")
     
  'CHECK si existe store primero
  If strSPprimero <> "" Then
    
    'EXEC store unica vez antes de importar
    SQLexec ("exec " & strSPprimero)
    
    'check errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
      Exit Sub
    End If
  
  End If
  '
  '-------------------------------------------------------------------------------

  
  'dim array para errores de importacion
  ReDim m_arrErrores(Me.spdD.MaxRows)
  
  'inicializo
  intOK = 0
  intERROR = 0
  
  'recorro filas
  For lngFila = 3 To Me.spdD.MaxRows
          
    'inicializo
    strParcial = ""
    m_arrErrores(lngFila) = ""
          
    'recorro columnas
    For intCol = 2 To Me.spdD.MaxCols
        
      'set puntero a fila coluna
      Me.spdD.col = intCol
      Me.spdD.row = lngFila
        
      'get nombre de columna
      strColumna = Me.dsiGridFrm.dsiEsquema(intCol - 2).Name
                          
      'get tipo de columna
      intTipo = Me.dsiGridFrm.dsiEsquema(strColumna).Type
                          
      'get valor de la celda
      Me.spdD.GetText intCol, lngFila, varValor
                    
      'get columna NOACTUALIZO
      strNoActualizo = Me.dsiGridFrm.dsiConfigGetValor("NOACTUALIZO", strColumna)
         
      'check si columna no esta dentro de las NOACTUALIZO
      If strNoActualizo = "" Then
                    
        'get LISTA
        strLista = Me.dsiGridFrm.dsiConfigGetValor("lista", strColumna)
            
        'check celda tipo LISTA
        If strLista <> "" And varValor <> "" Then
          
          'busco ID
          strID = Me.dsiComboBoxBuscar(strColumna, varValor, strLista)
                
          'set string parcial
          strParcial = strParcial + strID & ","
           
        Else
          
          'set string parcial
          strParcial = strParcial + DataConvert(intTipo, varValor) & ","
          
        End If
            
      End If
            
    Next
    
    'delete ultima coma
    strFinal = Left(strParcial, Len(strParcial) - 1)
    
    'get store procedure para importar
    strSP = Me.dsiGridFrm.dsiConfigGetValor("spIMP")
        
    'save info
    SQLexec ("exec " & strSP & " " & strFinal)
            
    'set puntero celda estado
    Me.spdD.col = 1
    
    'check errores
    If Not SQLparam.CnErrNumero = -1 Then
      
      'add numero y texto del error
      m_arrErrores(lngFila) = m_arrErrores(lngFila) + str(SQLparam.CnErrNumero) + ": " & SQLparam.CnErrTexto & " "
      
      'set celda indicando error
      Me.spdD.Text = "error"
      Me.spdD.ForeColor = RGB(255, 0, 0)
      
      'sum cantidad de filas con error
      intERROR = intERROR + 1
      
    Else
      
      'set celda indicando ok
      Me.spdD.Text = "ok"
      Me.spdD.ForeColor = RGB(0, 196, 0)
      
      'sum cantidad de filas con exito
      intOK = intOK + 1
    
    End If
    
  Next
    
  '-------------------------------------------------------------------------------
  '17/09/2009 Si existe el tipo SPIMP_ULTIMO, ejecuta un store antes de importar
  '
  '
  'GET store procedure primera vez
  strSPultimo = Me.dsiGridFrm.dsiConfigGetValor("SPIMP_ULTIMO")
     
  'CHECK si existe store primero
  If strSPultimo <> "" Then
    
    'EXEC store unica vez antes de importar
    SQLexec ("exec " & strSPultimo)
    
    'check errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
      Exit Sub
    End If
  
  End If
  '
  '-------------------------------------------------------------------------------
  
  'cierro
  SQLclose
  
  'recupero puntero mouse
  Screen.MousePointer = vbDefault
  
  blnB = MsgBox("Fin de la importación." & vbCrLf & vbCrLf & "Filas procesadas (" & LTrim(str(intOK)) & ")" & vbCrLf & "Filas con error (" & LTrim(str(intERROR)) & ")", vbInformation, "atención...")
  
End Property

Public Property Get dsiOrigenAbrir()

  Dim blnB As Boolean
  Dim strT As String
    
  Dim listcount As Integer, handle As Integer
  Dim List(10) As String
    
  'si formato excel filtro xls
  Me.comDestino.Filter = "Archivos de Excel (*.xls)|*.xls"
    
  'titulo de ventana
  Me.comDestino.DialogTitle = "Seleccionar archivo..."
  
  Me.comDestino.FileName = ""
  
  'abro cuadro de dialogo
  Me.comDestino.ShowOpen
      
  'si cancelar salgo
  If Me.comDestino.FileName = "" Then
    Exit Property
  End If
    
  'mouse reloj
  Screen.MousePointer = vbHourglass
  
  'importa excel origen
  blnB = Me.spdO.GetExcelSheetList(Me.comDestino.FileName, List, listcount, "", handle, True)
  blnB = Me.spdO.ImportExcelSheet(handle, 0)
  
  'mouse defa
  Screen.MousePointer = vbDefault
    
  'set activa cambios backGround
  Me.spdO.Redraw = False
  
  'change apariencia grilla
  blnB = Me.dsiSpdAparienciaO()
  
  'set limites grilla origen
  Me.spdO.MaxRows = Me.spdO.DataRowCnt
  Me.spdO.MaxCols = Me.spdO.DataColCnt
  
  'set activa cambios en foreGround
  Me.spdO.Redraw = True
  
  'config columnas destino
  blnB = Me.dsiColumnasConfig()
  
  'CLEAR
  m_dsiChequeada = False
  
End Property




Private Sub Form_Load()
  
  Dim blnB As Boolean
  
  'referencio gridFRM
  Set Me.dsiGridFrm = MainMDI.ActiveForm
    
  'pongo titulo a formulario
  Me.Caption = "Importando " & Me.dsiGridFrm.Caption & "..."
    
  'change apariencia grilla
  blnB = Me.dsiSpdAparienciaO()
  blnB = Me.dsiSpdAparienciaD()
    
  'inicializo cantidad de items en arrComboBox
  m_arrComboBoxCant = 0
  
End Sub

Private Sub Form_Resize()
  
'  Me.spdO.Width = Me.ScaleWidth
'  Me.spdD.Width = Me.ScaleWidth
'
'  If Me.ScaleHeight - Me.spdO.Height - Me.tlbHerra.Height - 40 > 0 Then
'    Me.spdD.Height = Me.ScaleHeight - Me.spdO.Height - Me.tlbHerra.Height - 40
'  End If
    
End Sub

Private Sub spdD_DblClick(ByVal col As Long, ByVal row As Long)
  
  Dim blnB As Boolean
  
  blnB = MsgBox(m_arrErrores(row), vbCritical + vbOKOnly, "atención...")
  
End Sub

Private Sub spdO_Click(ByVal col As Long, ByVal row As Long)
  
  Me.spdO.ToolTipText = "Debe hacer click en una columna de la grilla Destino, y luego hacer click en la barra de herramientas, icono Copiar."
  
End Sub

Private Sub tlbHerra_ButtonClick(ByVal Button As MSComctlLib.Button)
  
  Dim blnB As Boolean
  
  Select Case LCase(Button.Key)
    
    Case "cmd_abrir"
      
      blnB = Me.dsiOrigenAbrir()
      
    Case "cmd_copiar_auto"
      
      blnB = Me.dsiColumnaCopiar_auto()
      
    Case "cmd_copiar_manual"
      
      blnB = Me.dsiColumnaCopiar_manual()
      
    Case "cmd_chequear"
      
      blnB = Me.dsiChequear()
      
    Case "cmd_importar"
      
      blnB = Me.dsiImportar()
      
    End Select
      
End Sub

Public Property Get dsiChequear()
  
  Dim lngFila As Long
  Dim intIDposicion, intIDtipo As Integer
  Dim strIDnombre As String
  Dim varColumna, varValor, varParaValidar As Variant
  Dim strOperacion As String
  Dim blnColumnasOK As Boolean
  
  'CHECK si no filas en grilla origen, SALGO
  If Me.spdO.DataRowCnt = 0 Or Me.spdD.DataRowCnt = 0 Then
    
    blnB = MsgBox("No hay información para procesar...", vbExclamation, "atención...")
    Exit Property
    
  End If
  
  'FLAGS info chequeada, defecto OK
  m_dsiChequeada = True
  
  'GET nombre de columna ID
  strIDnombre = Me.dsiGridFrm.dsiConfigGetValor("CLAVE")
      
  'GET posicion de columna ID
  intIDposicion = Me.spdD.SearchRow(1, 1, -1, strIDnombre, SearchFlagsValue)
    
  'GET tipo de columna ID
  Me.spdD.row = 3
  Me.spdD.col = intIDposicion
  intIDtipo = Me.spdD.CellType
      
  'SET defecto insertar
  strOperacion = "INS"
    
  'SET defecto ID numerico
  varParaValidar = 0
  
  'CHECK si columna ID es texto
  If intIDtipo = 11 Then
    varParaValidar = ""
  End If
  
  'WHILE filas
  For lngFila = 3 To Me.spdD.DataRowCnt
      
    'GET valor ID
    Me.spdD.GetText intIDposicion, lngFila, varValor
      
    'CHECK si columna contiene datos
    If Val(varValor) <> varParaValidar Then
      strOperacion = "EDI"
    End If
      
  Next
    
  'SHOW cartel segun operacion
  If strOperacion = "INS" Then
    
    Me.txtOperacion.BackColor = RGB(225, 247, 227)
    Me.txtOperacion = "Usted va a realizar una operación para AGREGAR datos inexistente..."
        
    'CHECK columnas obligatorias
     blnColumnasOK = Me.dsiColumnasObligatorias()
    
    'CHECK si no paso validacion de columnas, no se permite importar
    If Not blnColumnasOK Then
      m_dsiChequeada = False
    End If
        
  Else
    
    Me.txtOperacion.BackColor = RGB(255, 215, 174)
    Me.txtOperacion = "Usted va a realizar una operación de ACTUALIZACION de datos existente..."
    
  End If
      
  'SET reforzada
  Me.txtOperacion.FontBold = True
      

End Property
  
'CHECK columnas obligatorias
'RETURN 1:OK, 0:ERROR
'
Public Property Get dsiColumnasObligatorias() As Boolean
  
  Dim lngFila As Long
  Dim intCol As Integer
  Dim strColumnaNombre, varValor As Variant
  Dim strObligatoria As String
  Dim intI As Integer
  
  'RETURN defecto en OK
  dsiColumnasObligatorias = 1
  
  'WHILE filas
  For lngFila = 3 To Me.spdD.DataRowCnt
    
    For intCol = 2 To Me.spdD.DataColCnt
          
      'GET columna nombre
      Me.spdD.GetText intCol, 1, strColumnaNombre
          
      'GET columna valor
      Me.spdD.GetText intCol, lngFila, varValor
          
      'check columna OBLIGATORIA
      strObligatoria = Me.dsiGridFrm.dsiConfigGetValor("obligatorio", strColumnaNombre)
        
      'CHECK si columna olbigatoria y celda vacia, no permito importar, y SHOW mensaje
      If strObligatoria <> "" And varValor = "" Then
        
        intI = MsgBox("Hay columnas olbigatorias que no contienen información...", vbCritical, "Atención...")
        
        'RETURN ERROR
        dsiColumnasObligatorias = 0
        
        'EXIT
        Exit Property
        
      End If
        
    Next
    
  Next
        
End Property

Public Function dsiColumnaCopiar_una(ByVal intO As Integer, ByVal intD As Integer) As Boolean
  
  Dim intUno, intDos, intTres As Integer
  Dim arr(), strNombreCol As Variant
    
  'genero colores diferentes para cada columna
  intUno = Int(Rnd() * 100 + 151)
  intDos = Int(Rnd() * 100 + 151)
  intTres = Int(Rnd() * 100 + 151)
    
  'redim cantidad de elementos para guardar datos de origen
  ReDim arr(Me.spdO.DataRowCnt) As Variant
  
  'set alcance grilla origen
  Me.spdO.col = intO
  Me.spdO.col2 = intO
  Me.spdO.row = 2
  Me.spdO.row2 = Me.spdO.MaxRows
  
  'change en bloque activate
  Me.spdO.BlockMode = True
  
  'change color
  Me.spdO.BackColor = RGB(intUno, intDos, intTres)
  
  'change en bloque deactivate
  Me.spdO.BlockMode = False
  
  '18/11/2009
  'FILL array con grilla origen
  'blnB = Me.spdO.GetArray(intO, 2, arr)
  
  '18/11/2009
  'Se modifico un metodo de la grilla por este codigo
  'para evitar que importa el signo negativo correctamente
  For a = 2 To Me.spdO.DataRowCnt
    Me.spdO.GetText intO, a, arr(a)
    Debug.Print intO, arr(a)
  Next
  
  'set limites destino
  Me.spdD.col = intD
  Me.spdD.col2 = intD
  Me.spdD.row = 3
  Me.spdD.row2 = Me.spdD.MaxRows
  
  'change en bloque activate
  Me.spdD.BlockMode = True
  
  '18/11/2009
  'Para celdas columnas numericas, tomo cantidad de decimales y ancho
  'Y se lo aplico a columnas destino
  Me.spdD.TypeNumberDecPlaces = Me.spdO.TypeNumberDecPlaces
  Me.spdD.ColWidth(intD) = Me.spdO.ColWidth(intO)
  
  '18/11/2009
  'FILL grilla destino con array
  'blnB = Me.spdD.SetArray(intD, 3, arr)
  
  '18/11/2009
  'Se modifico un metodo de la grilla por este codigo
  'para evitar que importa el signo negativo correctamente
  For a = 2 To Me.spdO.DataRowCnt
    Me.spdD.SetText intD, a + 1, arr(a)
    Debug.Print intD, arr(a)
  Next
  
  'paint columna destino
  Me.spdD.BackColor = RGB(intUno, intDos, intTres)
  
  'change en bloque deactivate
  Me.spdD.BlockMode = False
  
  'delete seleccion en origen y destino
  Me.spdO.ClearSelection
  Me.spdD.ClearSelection
  
End Function

