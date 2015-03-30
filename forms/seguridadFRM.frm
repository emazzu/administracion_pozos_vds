VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form seguridadFRM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administrador de Seguridad..."
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9780
   Icon            =   "seguridadFRM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Seleccionar las opciones sobre las cuales se aplicarán los permisos y luego click al boton derecho del mouse"
      Height          =   5385
      Left            =   60
      TabIndex        =   2
      Top             =   780
      Width           =   9675
      Begin FPSpreadADO.fpSpread SPD 
         Height          =   5085
         Left            =   60
         TabIndex        =   3
         Top             =   240
         Width           =   9555
         _Version        =   393216
         _ExtentX        =   16854
         _ExtentY        =   8969
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
         SpreadDesigner  =   "seguridadFRM.frx":1CCA
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccionar Usuario o Grupo"
      Height          =   645
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   9675
      Begin VB.CheckBox chkSeleccionAnular 
         Caption         =   "Eliminar selección de opciones"
         Height          =   255
         Left            =   6000
         TabIndex        =   6
         Top             =   300
         Width           =   2535
      End
      Begin VB.CheckBox chkSeleccionarTodas 
         Caption         =   "Seleccionar todas las opciones"
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   300
         Width           =   2535
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Salir"
         Height          =   300
         Left            =   8670
         TabIndex        =   4
         Top             =   210
         Width           =   900
      End
      Begin VB.ComboBox cboUsuarioGrupo 
         Height          =   315
         Left            =   60
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   210
         Width           =   3225
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnu_ag 
         Caption         =   "Agregar"
         Begin VB.Menu mnu_ag_todos 
            Caption         =   "Todos"
         End
         Begin VB.Menu mnu_ag_visualizar 
            Caption         =   "Visualizar"
         End
         Begin VB.Menu mnu_ag_Insertar 
            Caption         =   "Insertar"
         End
         Begin VB.Menu mnu_ag_editar 
            Caption         =   "Editar"
         End
         Begin VB.Menu mnu_ag_eliminar 
            Caption         =   "Eliminar"
         End
         Begin VB.Menu mnu_ag_importar 
            Caption         =   "Importar"
         End
      End
      Begin VB.Menu mnu_el 
         Caption         =   "Eliminar"
         Begin VB.Menu mnu_el_todos 
            Caption         =   "Todos"
         End
         Begin VB.Menu mnu_el_visualizar 
            Caption         =   "Visualizar"
         End
         Begin VB.Menu mnu_el_Insertar 
            Caption         =   "Insertar"
         End
         Begin VB.Menu mnu_el_editar 
            Caption         =   "Editar"
         End
         Begin VB.Menu mnu_el_eliminar 
            Caption         =   "Eliminar"
         End
         Begin VB.Menu mnu_el_importar 
            Caption         =   "Importar"
         End
      End
   End
End
Attribute VB_Name = "seguridadFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'cambiar apariencia a grilla
Public Property Get dsiCambiaApariencia() As Boolean
  
  Me.SPD.UnitType = UnitTypeTwips                  'trabajar en twips
  Me.SPD.OperationMode = OperationModeMulti         'modo seleccion multiple
    
  Me.SPD.Appearance = AppearanceFlat               'apariencia 3D
  Me.SPD.BorderStyle = BorderStyleNone             'tipo de borde: sin borde
  
  Me.SPD.ColHeadersAutoText = DispBlank            'titulos de columnas en blanco
  Me.SPD.ColHeadersShow = True                     'muestra encabezado columnas
  Me.SPD.RowHeadersShow = True                     'muestra encabezado de filas
  
  Me.SPD.CursorStyle = CursorStyleArrow            'stilo cursor
  Me.SPD.CursorType = CursorTypeDefault            'tipo cursor
  
  Me.SPD.AutoSize = False                          'automaticamente ajusta ancho grilla
  Me.SPD.DAutoSizeCols = DAutoSizeColsNone         'tipo de ajuste 2: al dato mas ancho
  
  Me.SPD.UserColAction = UserColActionDefault      'cuando hace click en header pinta columna o fila
  Me.SPD.FontSize = 9                              'tamaño letra
  
  Me.SPD.RowHeight(0) = 300                        'set altura fila con titulos
  Me.SPD.ColWidth(0) = 50                          'set ancho columna 0
  Me.SPD.MoveActiveOnFocus = False                 '
  Me.SPD.Protect = False                           'exporta a excel sin proteccion
  
  Me.SPD.BackColorStyle = BackColorStyleUnderGrid  'estilo
  Me.SPD.GridShowHoriz = True                      'muestra grilla horizontal
  Me.SPD.GridShowVert = True                       'muestra grilla vertical
  Me.SPD.GridColor = RGB(200, 200, 200)            'color muy suave
  Me.SPD.NoBorder = True                           'sin borde fin zona de datos
  
  Me.SPD.ScrollBars = ScrollBarsBoth               'ambas barras de desplazamiento
  Me.SPD.ScrollBarExtMode = False                  'cuando sean necesarias
  Me.SPD.VScrollSpecial = crollBarsVertical        'barra especial
      
  Me.SPD.SetOddEvenRowColor RGB(245, 245, 245), RGB(60, 60, 60), RGB(245, 245, 245), RGB(60, 60, 60)
  Me.SPD.SelBackColor = RGB(204, 230, 255)         'fondo del area seleccionada
  Me.SPD.GrayAreaBackColor = RGB(245, 245, 245)    'set area sobrante
  Me.SPD.LockBackColor = RGB(230, 230, 230)        'set color celdas bloqueadas
  
  Me.SPD.VirtualMode = False                        ' ajusta rows al tamaño del recordset
  'spd.VirtualRows = 300                         ' rows a leer del virtual buffer
  'spd.VirtualScrollBuffer = True                ' scroll vertical lee de tantas rows del buffer
  
  'setea para mostrar tooltip en las celdas donde no se ve toda la info
  Me.SPD.TextTip = TextTipFixedFocusOnly
  Me.SPD.TextTipDelay = 250
    
End Property

'build permisos
'
Property Get dsiPermisosSet(ByVal strSEL, strINS, strEDI, strELI, strIMP As String, ByVal strPermiso As String)
  
  Dim intI As Variant
  Dim intC, intC2, intF, intF2 As Variant
  Dim strT As String
  Dim blnB As Boolean
    
  'check si no se seleccionaron opciones, exit
  If Me.SPD.SelectionCount = 0 Then
    Exit Property
  End If
  
  'while cantidad selecciones
  For intI = 0 To Me.SPD.SelectionCount - 1
      
    'get fila seleccionada
    Me.SPD.GetSelection intI, intC, intF, intC2, intF2
      
    'get IDopcion
    Me.SPD.GetText 1, intF, strIDopcion
      
    'GRANT o REVOKE select
    If strSEL Then
      blnB = dsiPermisosSELexec(9, intF, strPermiso, "SEL")
    End If
            
    'GRANT o REVOKE insert
    If strINS Then
      blnB = dsiPermisosExec(10, intF, strPermiso, "INS")
    End If
      
    'GRANT o REVOKE update
    If strEDI Then
      blnB = dsiPermisosExec(11, intF, strPermiso, "EDI")
    End If
      
    'GRANT o REVOKE delete
    If strELI Then
      blnB = dsiPermisosExec(12, intF, strPermiso, "ELI")
    End If
      
    'GRANT o REVOKE import
    If strIMP Then
      blnB = dsiPermisosExec(13, intF, strPermiso, "IMP")
    End If
      
  Next
    
End Property


'agrega o elimina permisos de visualizacion SEL
'
'se trata en forma separada ya que este evento lo que hace
'es agregar o eliminar la opcion seleccionada a la tabla dsiOPCpermis
'
Property Get dsiPermisosSELexec(ByVal intCol As Integer, ByVal intFila As Integer, ByVal strPermiso As String, ByVal strOperacion As String)
      
  Dim strT As String
  Dim varIDopc, varSEL, varINS, varEDI, varELI, varIMP As Variant
  Dim blnB As Boolean
       
  'get IDopc
  Me.SPD.GetText 1, intFila, varIDopc
  
  'get marca celda visualizar
  Me.SPD.GetText 5, intFila, varSEL
  
  'check si GRANT o REVOKE
  If strPermiso = "GRANT" Then
    
    'check si no existe visualizar, agrego
    If varSEL = False Then
      
      'add permiso
      strT = "dsiOPCpermis_INS_sp " & SQLparam.IDmenu & "," & varIDopc & ",'" & Me.cboUsuarioGrupo & "'"
      SQLexec (strT)
      
      'check errores
      If Not SQLparam.CnErrNumero = -1 Then
        SQLError
        Exit Property
      End If
      
      'update celda visualizar
      Me.SPD.SetText 5, intFila, 1
      
    End If
    
  Else
    
    'check si existe visualizar, elimino
    If varSEL = True Then
      
      'get permisos segun grilla
      Me.SPD.GetText 6, intFila, varINS
      Me.SPD.GetText 7, intFila, varEDI
      Me.SPD.GetText 8, intFila, varELI
      Me.SPD.GetText 9, intFila, varIMP
      
      'check si existen permisos
      If varINS = 1 Or varEDI = 1 Or varELI = 1 Or varIMP = 1 Then
        blnB = MsgBox("Para eliminar permisos de Visualización, primero debe eliminar los otros permisos.", vbCritical + vbOKOnly, "Atención...")
        Exit Property
      End If
      
      'delete permiso
      strT = "exec dsiOPCpermis_ELI_sp " & SQLparam.IDmenu & "," & varIDopc & ",'" & Me.cboUsuarioGrupo & "'"
      SQLexec (strT)
      
      'check errores
      If Not SQLparam.CnErrNumero = -1 Then
        SQLError
        Exit Property
      End If
      
      'update celda visualizar
      Me.SPD.SetText 5, intFila, 0
      
    End If
    
  End If
  
End Property


'agrega o elimina permisos de INS, EDI, ELI, IMP
'
'se trata en forma separada ya que este evento lo que hace
'es por un lado agregar o eliminar permisos en SQL y por otro
'actualiza la tabla tabla dsiOPCpermis para sincronizarlo con SQL
'
Property Get dsiPermisosExec(ByVal intCol As Integer, ByVal intFila As Integer, ByVal strPermiso As String, ByVal strOperacion As String)
      
  Dim strT As String
  Dim varGrant, varStore, varIDopc, varSEL, varINS, varEDI, varELI, varIMP As Variant
  Dim blnB As Boolean
        
  'get IDopc
  Me.SPD.GetText 1, intFila, varIDopc
  
  'get store
  Me.SPD.GetText intCol, intFila, varStore
  
  'check si existe store
  If varStore <> "" Then
                
    'build query para permisos
    strT = strPermiso & " EXEC ON " & varStore & " TO [" & Me.cboUsuarioGrupo & "]"
        
    'exec transac
    SQLexec (strT)
        
    'check errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      Exit Property
    End If
        
    'update celda
    Me.SPD.SetText intCol - 4, intFila, IIf(strPermiso = "GRANT", 1, 0)
    
    'get permisos segun grilla
    Me.SPD.GetText 6, intFila, varINS
    Me.SPD.GetText 7, intFila, varEDI
    Me.SPD.GetText 8, intFila, varELI
    Me.SPD.GetText 9, intFila, varIMP
    
    'check si valor celda = 2, put 0, el 2 es cuando no hay store asociado a una celda
    varINS = IIf(varINS = 2, 0, varINS)
    varEDI = IIf(varEDI = 2, 0, varEDI)
    varELI = IIf(varELI = 2, 0, varELI)
    varIMP = IIf(varIMP = 2, 0, varIMP)
    
    'build query para sincronizar dsiOPCpermis con los permisos de SQL
    strT = "exec dsiOPCpermis_EDI_sp " & SQLparam.IDmenu & "," & varIDopc & ",'" & Me.cboUsuarioGrupo & "'," & _
           varINS & "," & varEDI & "," & varELI & "," & varIMP
           
    'exec transac
    SQLexec (strT)
        
    'check errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      Exit Property
    End If
     
  End If
    
End Property


Private Sub cboUsuarioGrupo_Click()
  
  Dim rs As ADODB.Recordset
  Dim strT As String
  Dim intI As Integer
  Dim varDato As Variant
  
  'get permisos
  strT = "" & _
  "select id as IDopción, Nombre, upper(tipo) as Tipo, usuario, " & _
  "case when dsiOPCpermis.IDopc is null then convert(bit, 0) else convert(bit, 1) end as Visualizar," & _
  "INS as Insertar, EDI as Editar, ELI as Eliminar, IMP as Importar," & _
  "(select valor from dsiOPCconfig where IDmenu = dsiOPC.IDmenu and IDopc = dsiOPC.id and tipo = 'SPINS') as SPINS," & _
  "(select valor from dsiOPCconfig where IDmenu = dsiOPC.IDmenu and IDopc = dsiOPC.id and tipo = 'SPEDI') as SPEDI," & _
  "(select valor from dsiOPCconfig where IDmenu = dsiOPC.IDmenu and IDopc = dsiOPC.id and tipo = 'SPELI') as SPELI," & _
  "(select valor from dsiOPCconfig where IDmenu = dsiOPC.IDmenu and IDopc = dsiOPC.id and tipo = 'SPIMP') as SPIMP " & _
  "from dsiOPC left join dsiOPCpermis on dsiOPC.IDmenu = dsiOPCpermis.IDmenu and dsiOPC.id = dsiOPCpermis.IDopc and " & _
  "dsiOPCpermis.usuario = '" & Me.cboUsuarioGrupo & "' and dsiOPC.IDmenu = " & SQLparam.IDmenu
   
  'exec
  Set rs = SQLexec(strT)
      
  'chequeo errores
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
  End If
      
  'set rs a grilla
  Set Me.SPD.DataSource = rs
      
  'set limite a grilla
  Me.SPD.MaxRows = Me.SPD.DataRowCnt
  Me.SPD.MaxCols = Me.SPD.DataColCnt
  
  'change tipo de celda par los que no tienen asociado un store procedure
  For intI = 1 To Me.SPD.MaxRows
        
    'set puntero row
    Me.SPD.row = intI
        
    'get store ins
    Me.SPD.GetText 10, intI, varDato
    
    'check si existe
    If varDato = "" Then
      Me.SPD.col = 6
      Me.SPD.TypeCheckType = TypeCheckTypeThreeState
      Me.SPD.Value = 2
    End If
    
    'get store edi
    Me.SPD.GetText 11, intI, varDato
    
    'check si existe
    If varDato = "" Then
      Me.SPD.col = 7
      Me.SPD.TypeCheckType = TypeCheckTypeThreeState
      Me.SPD.Value = 2
    End If
    
    'get store eli
    Me.SPD.GetText 12, intI, varDato
    
    'check si existe
    If varDato = "" Then
      Me.SPD.col = 8
      Me.SPD.TypeCheckType = TypeCheckTypeThreeState
      Me.SPD.Value = 2
    End If
    
    'get store imp
    Me.SPD.GetText 13, intI, varDato
    
    'check si existe
    If varDato = "" Then
      Me.SPD.col = 9
      Me.SPD.TypeCheckType = TypeCheckTypeThreeState
      Me.SPD.Value = 2
    End If
    
  Next

End Sub



Private Sub chkSeleccionAnular_Click()
  
  'check si hay filas para seleccionar
  If Me.SPD.DataRowCnt = 0 Then
    Exit Sub
  End If
  
  'set todas las filas
  Me.SPD.row = -1
    
  'set seleccion completa
  Me.SPD.SelModeSelected = False
    
  'unCheck seleccionar todo
  Me.chkSeleccionAnular.Value = 0
  
End Sub

Private Sub chkSeleccionarTodas_Click()
  
  'check si hay filas para seleccionar
  If Me.SPD.DataRowCnt = 0 Then
    Exit Sub
  End If
  
  'set todas las filas
  Me.SPD.row = -1
    
  'set seleccion completa
  Me.SPD.SelModeSelected = True
    
  'unCheck seleccionar todo
  Me.chkSeleccionarTodas.Value = 0
    
End Sub

Private Sub cmdCancelar_Click()
  
  'exit
  Unload Me
  
End Sub

Private Sub Form_Load()
  
  Dim rs As ADODB.Recordset
  Dim strT As String
  Dim blnB As Boolean
  
  'get grupos y usuarios SQL
  strT = "select name from sysusers where (isntgroup = 1 or isntuser = 1 or issqluser = 1) and uid not in (1,2,3)"
    
  'exec
  Set rs = SQLexec(strT)
      
  'chequeo errores
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    Exit Sub
  End If
        
  'while grupos y usuarios
  While Not rs.EOF
    
    'add usuarios a comboBox
    Me.cboUsuarioGrupo.AddItem rs!Name
    
    rs.MoveNext
    
  Wend
    
  'custimize grilla
  blnB = dsiCambiaApariencia()
    
  'hidden columnas
  Me.SPD.col = 4
  Me.SPD.ColHidden = True
    
  'hidden columnas
  Me.SPD.col = 10
  Me.SPD.ColHidden = True
  Me.SPD.col = 11
  Me.SPD.ColHidden = True
  Me.SPD.col = 12
  Me.SPD.ColHidden = True
  Me.SPD.col = 13
  Me.SPD.ColHidden = True
    
End Sub

Private Sub mnu_ag_editar_Click()
  Dim blnB As Boolean
  
  blnB = Me.dsiPermisosSet(False, False, True, False, False, "GRANT")
  
End Sub

Private Sub mnu_ag_eliminar_Click()
  Dim blnB As Boolean
  
  blnB = Me.dsiPermisosSet(False, False, False, True, False, "GRANT")
  
End Sub

Private Sub mnu_ag_importar_Click()
  Dim blnB As Boolean
  
  blnB = Me.dsiPermisosSet(False, False, False, False, True, "GRANT")
  
End Sub

Private Sub mnu_ag_Insertar_Click()
  Dim blnB As Boolean
  
  blnB = Me.dsiPermisosSet(False, True, False, False, False, "GRANT")
  
End Sub

Private Sub mnu_ag_todos_Click()
  
  blnB = Me.dsiPermisosSet(True, False, False, False, False, "GRANT")
  blnB = Me.dsiPermisosSet(False, True, False, False, False, "GRANT")
  blnB = Me.dsiPermisosSet(False, False, True, False, False, "GRANT")
  blnB = Me.dsiPermisosSet(False, False, False, True, False, "GRANT")
  blnB = Me.dsiPermisosSet(False, False, False, False, True, "GRANT")
  
End Sub

Private Sub mnu_ag_visualizar_Click()
  Dim blnB As Boolean
  
  blnB = Me.dsiPermisosSet(True, False, False, False, False, "GRANT")
  
End Sub

Private Sub mnu_el_editar_Click()
  Dim blnB As Boolean
  
  blnB = Me.dsiPermisosSet(False, False, True, False, False, "REVOKE")
  
End Sub

Private Sub mnu_el_eliminar_Click()
  Dim blnB As Boolean
  
  blnB = Me.dsiPermisosSet(False, False, False, True, False, "REVOKE")
  
End Sub

Private Sub mnu_el_importar_Click()
  Dim blnB As Boolean
  
  blnB = Me.dsiPermisosSet(False, False, False, False, True, "REVOKE")
  
End Sub

Private Sub mnu_el_Insertar_Click()
  Dim blnB As Boolean
  
  blnB = Me.dsiPermisosSet(False, True, False, False, False, "REVOKE")
  
End Sub

Private Sub mnu_el_todos_Click()
  
  blnB = Me.dsiPermisosSet(True, False, False, False, False, "REVOKE")
  blnB = Me.dsiPermisosSet(False, True, False, False, False, "REVOKE")
  blnB = Me.dsiPermisosSet(False, False, True, False, False, "REVOKE")
  blnB = Me.dsiPermisosSet(False, False, False, True, False, "REVOKE")
  blnB = Me.dsiPermisosSet(False, False, False, False, True, "REVOKE")

End Sub

Private Sub mnu_el_visualizar_Click()
  Dim blnB As Boolean
  
  blnB = Me.dsiPermisosSet(True, False, False, False, False, "REVOKE")
  
End Sub

Private Sub SPD_RightClick(ByVal ClickType As Integer, ByVal col As Long, ByVal row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
  
  'check si no se seleccionaron opciones, exit
  If Me.SPD.SelectionCount = 0 Then
    Exit Sub
  End If
  
  'open menu emergente
  PopupMenu mnu, vbPopupMenuCenterAlign
  
End Sub
