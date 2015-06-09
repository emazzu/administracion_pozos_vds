VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ExportarFRM 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Exportando..."
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   1185
      Left            =   7740
      TabIndex        =   10
      Top             =   -45
      Width           =   1680
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   330
         Left            =   315
         TabIndex        =   7
         Top             =   675
         Width           =   1100
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "&Exportar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   300
         TabIndex        =   6
         Top             =   270
         Width           =   1100
      End
   End
   Begin MSComDlg.CommonDialog comDestino 
      Left            =   945
      Top             =   4725
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Height          =   510
      Left            =   2340
      TabIndex        =   9
      Top             =   -45
      Width           =   5400
      Begin VB.CheckBox chkTitulo 
         Caption         =   "Mostrar nombre de columnas"
         Enabled         =   0   'False
         Height          =   240
         Left            =   270
         TabIndex        =   1
         Top             =   180
         Width           =   2445
      End
   End
   Begin VB.Frame Frame2 
      Height          =   825
      Left            =   2340
      TabIndex        =   8
      Top             =   315
      Width           =   1800
      Begin VB.OptionButton opcTXT 
         Caption         =   "Texto"
         Enabled         =   0   'False
         Height          =   240
         Left            =   270
         TabIndex        =   3
         Top             =   495
         Width           =   915
      End
      Begin VB.OptionButton opcEXC 
         Caption         =   "Excel"
         Enabled         =   0   'False
         Height          =   195
         Left            =   270
         TabIndex        =   2
         Top             =   270
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   4140
      TabIndex        =   4
      Top             =   315
      Width           =   3600
      Begin VB.ComboBox cboSepa 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "exportarFRM.frx":0000
         Left            =   855
         List            =   "exportarFRM.frx":0010
         TabIndex        =   11
         Top             =   405
         Width           =   1770
      End
      Begin VB.Label Label1 
         Caption         =   "Separador de columnas"
         Height          =   285
         Left            =   900
         TabIndex        =   12
         Top             =   180
         Width           =   1770
      End
   End
   Begin VB.ListBox lstEsquema 
      Height          =   5325
      IntegralHeight  =   0   'False
      Left            =   45
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   30
      Width           =   2265
   End
   Begin FPSpreadADO.fpSpread spdAux 
      Height          =   4245
      Left            =   2340
      TabIndex        =   5
      Top             =   1170
      Width           =   7080
      _Version        =   393216
      _ExtentX        =   12488
      _ExtentY        =   7488
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
      SpreadDesigner  =   "exportarFRM.frx":0034
   End
End
Attribute VB_Name = "ExportarFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_gridFrm As gridFRM
Dim m_columnasCant As Integer

Public Property Set dsiGridFrm(frm As gridFRM)
  Set m_gridFrm = frm
End Property
    
Public Property Get dsiGridFrm() As gridFRM
  Set dsiGridFrm = m_gridFrm
End Property

Public Property Let dsiColumnasCant(intI As Integer)
  
  Let m_columnasCant = intI
  
End Property
    
Public Property Get dsiColumnasCant() As Integer
  
  Let dsiColumnasCant = m_columnasCant
  
End Property

Public Property Get dsiColumnaAgregar(intI As Integer) As Boolean

  Dim arrCol As Variant
  Dim intTipoDato As Integer
  Dim intDecimales As Integer
        
  'mouse reloj arena
  Screen.MousePointer = vbHourglass
        
  'dimensiono array con cantidad de filas
  ReDim arrCol(Me.dsiGridFrm.spdGrid.MaxRows) As Variant
    
  'set cantidad filas grilla destino
  Me.spdAux.MaxRows = Me.dsiGridFrm.spdGrid.MaxRows
      
  'guardo columna de grilla origen en array
  Me.dsiGridFrm.spdGrid.GetArray intI, 0, arrCol
            
  'get tomo formato de celda origen
  Me.dsiGridFrm.spdGrid.col = intI
  intTipoDato = Me.dsiGridFrm.spdGrid.CellType
  intDecimales = Me.dsiGridFrm.spdGrid.TypeNumberDecPlaces
            
  'activo columna con ........ cuando no se ve todo el texto
  Me.spdAux.TypeEllipses = True
    
  'activo limite de columna tipo texto a 5000
  Me.spdAux.TypeMaxEditLen = 5000
            
  'setea para mostrar tooltip en las celdas donde no se ve toda la info
  Me.spdAux.TextTip = TextTipFixedFocusOnly
  Me.spdAux.TextTipDelay = 250
  
  'copio columna de array en destino
  Me.spdAux.SetArray Me.dsiColumnasCant, 0, arrCol
  
  'set aplico formato a celda destino
  Me.spdAux.row = -1
  Me.spdAux.col = Me.dsiColumnasCant
  Me.spdAux.CellType = intTipoDato
  Me.spdAux.TypeNumberDecPlaces = intDecimales
  
  'set ID columna igual a nombre de columna
'  Me.spdAux.col = Me.dsiColumnasCant
  Me.spdAux.ColID = arrCol(0)
    
  'set activo celda copiada
  Me.spdAux.SetActiveCell Me.dsiColumnasCant, 1
    
   'mouse flecha
   Screen.MousePointer = vbDefault
    
End Property

Public Property Get dsiColumnaEliminar(strColNombre As String) As Boolean
  
  Dim lngL As Long
          
  'averiguo numero de columna a partir del nombre
  lngL = Me.spdAux.GetColFromID(strColNombre)
          
  'elimino columna
  Me.spdAux.DeleteCols lngL, 1
  Me.spdAux.MaxCols = Me.spdAux.MaxCols - 1
        
End Property

Public Property Get dsiColumnaMostrarNombre(blnB As Boolean) As Boolean
  
  Dim lngL As Long
  Dim varT As Variant
  
  If blnB Then
    
    'agrega fila como fila 1 para nombres de columnas
    Me.spdAux.MaxRows = Me.spdAux.MaxRows + 1
    Me.spdAux.InsertRows 1, 1
    
    'set tipo de dato texto y copio nombres de columnas
    For lngL = 1 To Me.spdAux.DataColCnt
      Me.spdAux.GetText lngL, 0, varT
      Me.spdAux.row = 1
      Me.spdAux.col = lngL
      Me.spdAux.CellType = CellTypeEdit
      Me.spdAux.SetText lngL, 1, varT
    Next
     
    'set color a fila con titulos
    Me.spdAux.row = 1
    Me.spdAux.col = -1
    Me.spdAux.BackColor = RGB(191, 215, 189)
    
    'congela fila con titulos, haciendola siempre visible
    Me.spdAux.RowsFrozen = 1
     
  Else
    
    'elimina fila con nombre de columnas
    Me.spdAux.DeleteRows 1, 1
    Me.spdAux.MaxRows = Me.spdAux.MaxRows - 1
    
    'congela fila con titulos, haciendola siempre visible
    Me.spdAux.RowsFrozen = 0
    
  End If

End Property

Private Sub chkTitulo_Click()
    
  Dim blnB As Boolean
    
  'muestra o no nombre de columna
  If chkTitulo.Value Then
    blnB = Me.dsiColumnaMostrarNombre(True)
  Else
    blnB = Me.dsiColumnaMostrarNombre(False)
  End If
  
End Sub

Private Sub cmdCancelar_Click()
  
  Unload Me

End Sub

Private Sub cmdExportar_Click()
  
  Dim blnB As Boolean
  Dim strT As String
    
  'si formato excel filtro xls, sino txt
  If Me.opcEXC = True Then
    Me.comDestino.Filter = "Archivos de Excel|*.xls"
  Else
    Me.comDestino.Filter = "Archivos de texto (*.txt)|*.txt"
  End If
    
  'titulo de ventana
  Me.comDestino.DialogTitle = "Exportando..."
  
  Me.comDestino.FileName = ""
  
  'abro cuadro de dialogo
  Me.comDestino.ShowSave
      
  'si cancelar salgo
  If Me.comDestino.FileName = "" Then
    Exit Sub
  End If
    
  'mouse reloj
  Screen.MousePointer = vbHourglass
  
  'exportando a excel
  If Me.opcEXC.Value Then
    
    'esto es para no generar una planilla bloqueada y exporto
    Me.spdAux.Protect = False
    blnB = Me.spdAux.ExportToExcel(Me.comDestino.FileName, "", "")
  
  'exportando a texto
  ElseIf Me.cboSepa = "tab" Then
          
    'exporto con tab
    blnB = Me.spdAux.SaveTabFile(Me.comDestino.FileName)
      
      Else
                
        'segun separador seleccionado, genero string y exporto
        strT = Switch(Me.cboSepa = "coma", ",", Me.cboSepa = "punto y coma", ";", Me.cboSepa = "space", " ")
        blnB = Me.spdAux.ExportToTextFile(Me.comDestino.FileName, "", strT, "", ExportToTextFileCreateNewFile, "")
        
  End If
  
  'mouse defa
  Screen.MousePointer = vbDefault
    
  'status
  If blnB Then
    blnB = MsgBox("La exportación se realizó con éxito.", vbInformation + vbOKOnly, "atención...")
  Else
    blnB = MsgBox("La exportación fallo.", vbCritical + vbOKOnly, "atención...")
  End If
    
End Sub


Private Sub Form_Load()
  
  Dim fld As Field
  Dim blnB As Boolean
        
  'referencio grilla
  Set Me.dsiGridFrm = MainMDI.ActiveForm
      
  'pongo titulo a formulario
  Me.Caption = "Exportando " & Me.dsiGridFrm.Caption & "..."
      
  'cambia apariencia
  spdAux.Appearance = AppearanceFlat
  spdAux.BorderStyle = BorderStyleNone
  
  'grilla en twips
  spdAux.UnitType = UnitTypeTwips
  
  'header muy finito
  spdAux.RowHeight(0) = 50
            
  'pinta fila activa
  spdAux.OperationMode = OperationModeRow
            
  'muestra lineas horizontales y verticales
  spdAux.BackColorStyle = BackColorStyleUnderGrid
  spdAux.GridShowHoriz = True
  spdAux.GridShowVert = True
  spdAux.GridColor = RGB(200, 200, 200)
          
  'set color
  blnB = spdAux.SetOddEvenRowColor(RGB(245, 245, 245), RGB(60, 60, 60), RGB(245, 245, 245), RGB(60, 60, 60))
  Me.spdAux.SelBackColor = RGB(220, 220, 220)
          
  'agrego opcion todos
  lstEsquema.AddItem "(todos)"
  
  'agrego nombres de columnas
  For Each fld In Me.dsiGridFrm.dsiEsquema
    lstEsquema.AddItem fld.Name
  Next
    
  'set cantidad columnas filtradas
  Me.dsiColumnasCant = 0
  
End Sub

Private Sub Form_Resize()
  
  'set altura de lista y grilla
  Me.lstEsquema.Height = IIf(Me.ScaleHeight - Me.lstEsquema.Top - 50 < 0, 0, Me.ScaleHeight - Me.lstEsquema.Top - 50)
  Me.spdAux.Height = IIf(Me.ScaleHeight - Me.spdAux.Top - 50 < 0, 0, Me.ScaleHeight - Me.spdAux.Top - 50)
  
  'set ancho de grilla y frame
  Me.spdAux.Width = IIf(Me.ScaleWidth - Me.spdAux.Left - 50 < 0, 0, Me.ScaleWidth - Me.spdAux.Left - 50)
    
  'set ancho frame5
  Me.Frame5.Width = IIf(Me.ScaleWidth - Me.Frame5.Left - 50 < 0, 0, Me.ScaleWidth - Me.Frame5.Left - 50)
    
End Sub

Private Sub lstEsquema_ItemCheck(Item As Integer)
    
  Dim blnB As Boolean
  Dim intI As Integer
    
  'chequeo si se hizo click en (todos)
  If lstEsquema = "(todos)" Then
        
    'cambios de grilla backGround
    Me.spdAux.ReDraw = False
    
    'limpio grilla, si seleccion todos y ya existe algun dato
    If Me.dsiColumnasCant > 0 Then
      
      'clear cantidad columnas
      Me.spdAux.ClearRange 1, 1, -1, -1, False
      Me.dsiColumnasCant = 0
      
    End If
    
    'fuerzo click / unclick  en todos los items
    For intI = 1 To lstEsquema.listcount - 1
      
      'verifico si se hizo un check
      If lstEsquema.Selected(Item) = True Then
      
        'add 1 cantidad columnas
        Me.dsiColumnasCant = Me.dsiColumnasCant + 1
        
        'add columna
        lstEsquema.Selected(intI) = True
        blnB = Me.dsiColumnaAgregar(intI)
          
      Else
        
        'delete columna
        lstEsquema.Selected(intI) = False
        blnB = Me.dsiColumnaEliminar(lstEsquema.ItemData(intI))
          
        'delete 1 cantidad columnas
        Me.dsiColumnasCant = Me.dsiColumnasCant - 1

      End If
      
    Next
      
    'cambios de grilla foreGround
    Me.spdAux.ReDraw = True
    
  'click en un item solo
  Else
      
    'verifico si se hizo un check
    If lstEsquema.Selected(Item) = True Then
      
      'add 1 cantidad columnas
      Me.dsiColumnasCant = Me.dsiColumnasCant + 1
      
      'add columna
      blnB = Me.dsiColumnaAgregar(Item)
        
    Else
      
      'delete columna
      blnB = Me.dsiColumnaEliminar(lstEsquema)
      
      'delete 1 cantidad columnas
      Me.dsiColumnasCant = Me.dsiColumnasCant - 1
      
    End If
      
  End If
      
  'habilito / deshabilito opciones
  If Me.dsiColumnasCant > 0 Then
    
    Me.opcEXC.Enabled = True
    Me.opcEXC = True
    Me.opcTXT.Enabled = True
    Me.chkTitulo.Enabled = True
    Me.cmdExportar.Enabled = True
    
  Else
    
    Me.opcEXC.Enabled = False
    Me.opcEXC = True
    Me.opcTXT.Enabled = False
    Me.chkTitulo.Value = False
    Me.chkTitulo.Enabled = False
    Me.cmdExportar.Enabled = False
  
  End If
      
End Sub

Private Sub opcEXC_Click()
  
  Me.cboSepa.Enabled = False
  Me.cboSepa = ""
    
End Sub

Private Sub opcTXT_Click()
  
  'defa coma
  Me.cboSepa.Enabled = True
  Me.cboSepa = "coma"
  
End Sub

