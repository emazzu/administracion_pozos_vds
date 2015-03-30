VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form gridConfigFRM 
   Caption         =   "Configuration of Options"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Configuración"
      Height          =   6810
      Left            =   4815
      TabIndex        =   4
      Top             =   90
      Width           =   6540
      Begin FPSpreadADO.fpSpread fpSpread3 
         Height          =   870
         Left            =   3510
         TabIndex        =   25
         Top             =   2655
         Width           =   2625
         _Version        =   393216
         _ExtentX        =   4630
         _ExtentY        =   1535
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
         SpreadDesigner  =   "gridConfigFRM.frx":0000
      End
      Begin FPSpreadADO.fpSpread fpSpread2 
         Height          =   870
         Left            =   3510
         TabIndex        =   24
         Top             =   1620
         Width           =   2670
         _Version        =   393216
         _ExtentX        =   4710
         _ExtentY        =   1535
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
         SpreadDesigner  =   "gridConfigFRM.frx":01D4
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   1140
         Left            =   3510
         TabIndex        =   23
         Top             =   315
         Width           =   2625
         _Version        =   393216
         _ExtentX        =   4630
         _ExtentY        =   2011
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
         SpreadDesigner  =   "gridConfigFRM.frx":03A8
      End
      Begin VB.TextBox txtCierre 
         Height          =   330
         Left            =   135
         TabIndex        =   13
         Top             =   5220
         Width           =   2985
      End
      Begin VB.TextBox txtFiltra 
         Height          =   330
         Left            =   135
         TabIndex        =   12
         Top             =   4635
         Width           =   2985
      End
      Begin VB.TextBox txtCantidadFilas 
         Height          =   330
         Left            =   135
         TabIndex        =   11
         Top             =   4050
         Width           =   2985
      End
      Begin VB.TextBox txtSPimp 
         Height          =   330
         Left            =   135
         TabIndex        =   10
         Top             =   3465
         Width           =   2985
      End
      Begin VB.TextBox txtSPeli 
         Height          =   330
         Left            =   135
         TabIndex        =   9
         Top             =   2880
         Width           =   2985
      End
      Begin VB.TextBox txtSPedi 
         Height          =   330
         Left            =   135
         TabIndex        =   8
         Top             =   2295
         Width           =   2985
      End
      Begin VB.TextBox txtSPins 
         Height          =   330
         Left            =   135
         TabIndex        =   7
         Top             =   1710
         Width           =   2985
      End
      Begin VB.TextBox txtClave 
         Height          =   330
         Left            =   135
         TabIndex        =   6
         Top             =   1125
         Width           =   2985
      End
      Begin VB.TextBox txtVista 
         Height          =   330
         Left            =   135
         TabIndex        =   5
         Top             =   540
         Width           =   2985
      End
      Begin VB.Label Label9 
         Caption         =   "Cierre"
         Height          =   240
         Left            =   180
         TabIndex        =   22
         Top             =   4995
         Width           =   2895
      End
      Begin VB.Label Label8 
         Caption         =   "Filtro General"
         Height          =   240
         Left            =   180
         TabIndex        =   21
         Top             =   4410
         Width           =   2895
      End
      Begin VB.Label Label7 
         Caption         =   "Cantidad de Filas"
         Height          =   240
         Left            =   180
         TabIndex        =   20
         Top             =   3825
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Store Procedure para Importar"
         Height          =   240
         Left            =   180
         TabIndex        =   19
         Top             =   3240
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Store Procedure para Eliminar"
         Height          =   240
         Left            =   180
         TabIndex        =   18
         Top             =   2655
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Store Procedure para Editar"
         Height          =   240
         Left            =   180
         TabIndex        =   17
         Top             =   2070
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Store Procedure para Insertar"
         Height          =   240
         Left            =   180
         TabIndex        =   16
         Top             =   1485
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Clave"
         Height          =   240
         Left            =   180
         TabIndex        =   15
         Top             =   900
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Vista"
         Height          =   240
         Left            =   180
         TabIndex        =   14
         Top             =   315
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccionar las columnas a configurar"
      Height          =   3300
      Left            =   90
      TabIndex        =   2
      Top             =   3600
      Width           =   4695
      Begin FPSpreadADO.fpSpread spdV 
         Height          =   2985
         Left            =   90
         TabIndex        =   3
         Top             =   225
         Width           =   4515
         _Version        =   393216
         _ExtentX        =   7964
         _ExtentY        =   5265
         _StockProps     =   64
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "gridConfigFRM.frx":057C
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar la opción a configurar"
      Height          =   3480
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4695
      Begin FPSpreadADO.fpSpread spdO 
         Height          =   3165
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   4515
         _Version        =   393216
         _ExtentX        =   7964
         _ExtentY        =   5583
         _StockProps     =   64
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "gridConfigFRM.frx":0750
      End
   End
End
Attribute VB_Name = "gridConfigFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'cambiar apariencia a grilla
Public Property Get dsiCambiaApariencia() As Boolean
  
  Me.spdO.UnitType = UnitTypeTwips                  'trabajar en twips
  Me.spdO.OperationMode = OperationModeRow          'modo seleccion multiple
    
  Me.spdO.Appearance = AppearanceFlat               'apariencia 3D
  
  Me.spdO.ColHeadersAutoText = DispBlank            'titulos de columnas en blanco
  Me.spdO.ColHeadersShow = True                     'muestra encabezado columnas
  Me.spdO.RowHeadersShow = True                     'muestra encabezado de filas
  
  Me.spdO.CursorStyle = CursorStyleArrow            'stilo cursor
  Me.spdO.CursorType = CursorTypeDefault            'tipo cursor
  
  Me.spdO.AutoSize = False                          'automaticamente ajusta ancho grilla
  Me.spdO.DAutoSizeCols = DAutoSizeColsMax          'tipo de ajuste 2: al dato mas ancho
  
  Me.spdO.UserColAction = UserColActionDefault      'cuando hace click en header pinta columna o fila
  Me.spdO.FontSize = 9                              'tamaño letra
  
  Me.spdO.RowHeight(0) = 300                        'set altura fila con titulos
  Me.spdO.ColWidth(0) = 50                          'set ancho columna 0
  Me.spdO.MoveActiveOnFocus = False                 '
  Me.spdO.Protect = False                           'exporta a excel sin proteccion
  
  Me.spdO.BackColorStyle = BackColorStyleUnderGrid  'estilo
  Me.spdO.GridShowHoriz = True                      'muestra grilla horizontal
  Me.spdO.GridShowVert = True                       'muestra grilla vertical
  Me.spdO.GridColor = RGB(200, 200, 200)            'color muy suave
  Me.spdO.NoBorder = True                           'sin borde fin zona de datos
  
  Me.spdO.ScrollBars = ScrollBarsBoth               'ambas barras de desplazamiento
  Me.spdO.ScrollBarExtMode = False                  'cuando sean necesarias
  Me.spdO.VScrollSpecial = crollBarsVertical        'barra especial
      
  Me.spdO.SetOddEvenRowColor RGB(245, 245, 245), RGB(60, 60, 60), RGB(245, 245, 245), RGB(60, 60, 60)
  Me.spdO.SelBackColor = RGB(204, 230, 255)         'fondo del area seleccionada
  Me.spdO.GrayAreaBackColor = RGB(245, 245, 245)    'set area sobrante
  Me.spdO.LockBackColor = RGB(230, 230, 230)        'set color celdas bloqueadas
  
  Me.spdO.VirtualMode = False                       ' ajusta rows al tamaño del recordset
  Me.spdO.VirtualRows = 300                         ' rows a leer del virtual buffer
  Me.spdO.VirtualScrollBuffer = True                ' scroll vertical lee de tantas rows del buffer
  
  'setea para mostrar tooltip en las celdas donde no se ve toda la info
  Me.spdO.TextTip = TextTipFixedFocusOnly
  Me.spdO.TextTipDelay = 250
    
End Property


Private Sub Form_Load()
  
  Dim rs As ADODB.Recordset
  Dim strT As String
  Dim blnB As Boolean
  
  'get grupos y usuarios SQL
  strT = "select IDmenu, id as IDopción, Nombre from dsiOPC where tipo not in ('grp')"
    
  'exec
  Set rs = SQLexec(strT)
      
  'chequeo errores
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    Exit Sub
  End If
        
  'set datasource
  Set Me.spdO.DataSource = rs.DataSource
      
  'set milites grilla
  Me.spdO.MaxRows = Me.spdO.DataRowCnt
  Me.spdO.MaxCols = Me.spdO.DataColCnt
        
  'customize grilla
  blnB = dsiCambiaApariencia()
  
  'hidden columna 1
  Me.spdO.col = 1
  Me.spdO.ColHidden = True
  
End Sub
