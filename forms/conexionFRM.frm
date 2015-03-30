VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form conexionFRM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parametros de Conexión"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4305
   Icon            =   "conexionFRM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog comOrigen 
      Left            =   90
      Top             =   5490
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtAyuda 
      Alignment       =   2  'Center
      Height          =   885
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   4665
      Width           =   4125
   End
   Begin FPSpreadADO.fpSpread SPD 
      Height          =   4590
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4125
      _Version        =   393216
      _ExtentX        =   7276
      _ExtentY        =   8096
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      EditEnterAction =   2
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      MaxRows         =   13
      RowHeaderDisplay=   0
      ScrollBars      =   0
      SpreadDesigner  =   "conexionFRM.frx":0CCA
      Appearance      =   1
      TextTip         =   1
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   300
      Left            =   1950
      TabIndex        =   2
      Top             =   5625
      Width           =   1080
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      Height          =   300
      Left            =   3120
      TabIndex        =   1
      Top             =   5625
      Width           =   1080
   End
End
Attribute VB_Name = "conexionFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
  
  Dim strTm, strN, strINInombre As String
  Dim rs As ADODB.Recordset
  Dim blnB As Boolean
  Dim varDato, varSI As Variant
    
  'change mouse
  Screen.MousePointer = vbHourglass
    
  'set datos de conexion
  Me.SPD.GetText 2, 1, varDato
  SQLparam.IDmenu = varDato
  
  Me.SPD.GetText 2, 2, varDato
  SQLparam.Proveedor = varDato
  
  Me.SPD.GetText 2, 3, varDato
  SQLparam.ServerSeguridad = varDato
  
  Me.SPD.GetText 2, 4, varDato
  SQLparam.ServerDatos = varDato
  
  Me.SPD.GetText 2, 5, varDato
  SQLparam.BaseDEdatos = varDato
  
  Me.SPD.GetText 2, 6, varSI
  SQLparam.SeguridadIntegrada = IIf(Val(varSI) = 1, "true", "false")
  
  Me.SPD.GetText 2, 7, varDato
  SQLparam.TiempoEspera = varDato
  
  Me.SPD.GetText 2, 8, varDato
  SQLparam.Usuario = IIf(Val(varSI) = 1, "", varDato)
  
  Me.SPD.GetText 2, 9, varDato
  SQLparam.UsuarioClave = IIf(Val(varSI) = 1, "", varDato)
  
  Me.SPD.GetText 2, 10, varDato
  SQLparam.CantidadFilas = varDato
  
  Me.SPD.GetText 2, 13, varDato
  SQLparam.ReportesPath = varDato
  
  'check conexion
  strT = "select top 1 * from sysUsers"
  Set rs = SQLexec(strT)
  
  'check errores
  If Not SQLparam.CnErrNumero = -1 Then
    
    'message error
    SQLError
    
  Else
        
    'get UserName: 26/09/2006
    strINInombre = NTuserName()
    
    'build nombre de INI: 26/09/2006
    strN = App.Path & "\" & strINInombre & ".ini"
    
    'save parametros de conexion
    blnB = WriteIni("conexion", "IDmenu", SQLparam.IDmenu, strN)
    blnB = WriteIni("conexion", "Proveedor", SQLparam.Proveedor, strN)
    blnB = WriteIni("conexion", "ServerSeguridad", SQLparam.ServerSeguridad, strN)
    blnB = WriteIni("conexion", "ServerDatos", SQLparam.ServerDatos, strN)
    blnB = WriteIni("conexion", "BaseDEdatos", SQLparam.BaseDEdatos, strN)
    blnB = WriteIni("conexion", "SeguridadIntegrada", SQLparam.SeguridadIntegrada, strN)
    blnB = WriteIni("conexion", "TiempoEspera", SQLparam.TiempoEspera, strN)
    blnB = WriteIni("conexion", "Usuario", SQLparam.Usuario, strN)
    blnB = WriteIni("conexion", "UsuarioClave", SQLparam.UsuarioClave, strN)
    blnB = WriteIni("conexion", "CantidadFilas", SQLparam.CantidadFilas, strN)
    blnB = WriteIni("conexion", "ReportesPath", SQLparam.ReportesPath, strN)
    
    'message OK
    blnB = MsgBox("Los parametros de conexión son correctos.", vbInformation + vbOKOnly, "Atención...")
      
    'exit form
    Unload Me
    
  End If
   
  'change mouse
  Screen.MousePointer = vbDefault
   
End Sub

Private Sub cmdCancelar_Click()
  
  'exit
  Unload Me
  
End Sub

Private Sub Form_Load()
  
  Dim strG As String
  
  'check si array contiene algo
  If Not IsEmpty(SQLparam.GrupoConectado) Then
    
    'while grupos a los cuales pertenece el usuario
    For intI = 1 To UBound(SQLparam.GrupoConectado)
    
      'add grupo
      strG = strG & SQLparam.GrupoConectado(intI) & vbCrLf
    
    Next
    
  End If
  
  'get datos de conexion
  Me.SPD.SetText 2, 1, IIf(SQLparam.IDmenu = "", 0, SQLparam.IDmenu)
  Me.SPD.SetText 2, 2, IIf(SQLparam.Proveedor = "", "SQLOLEDB", SQLparam.Proveedor)
  Me.SPD.SetText 2, 3, SQLparam.ServerSeguridad
  Me.SPD.SetText 2, 4, SQLparam.ServerDatos
  Me.SPD.SetText 2, 5, SQLparam.BaseDEdatos
  Me.SPD.SetText 2, 6, IIf(SQLparam.SeguridadIntegrada = "true", 1, 0)
  Me.SPD.SetText 2, 7, IIf(SQLparam.TiempoEspera = "", 60, CVar(SQLparam.TiempoEspera))
  Me.SPD.SetText 2, 8, SQLparam.Usuario
  Me.SPD.SetText 2, 9, SQLparam.UsuarioClave
  Me.SPD.SetText 2, 10, IIf(SQLparam.CantidadFilas = "", "1000", SQLparam.CantidadFilas)
  Me.SPD.SetText 2, 11, SQLparam.UsuarioConectado
  Me.SPD.SetText 2, 12, strG
  Me.SPD.SetText 2, 13, SQLparam.ReportesPath
  
  'set celda activa
  Me.SPD.SetActiveCell 2, 1
  
End Sub



Private Sub SPD_DblClick(ByVal col As Long, ByVal row As Long)
  
  'check si row <> 13, exit
  If row <> 13 Then
    Exit Sub
  End If
  
  'si formato excel filtro xls
  Me.comOrigen.Filter = "Archivos de Crystal Report (*.rpt) | *.rpt"
  
  'titulo de ventana
  Me.comOrigen.DialogTitle = "Seleccionar archivo..."
  
  Me.comOrigen.FileName = ""
  
  'abro cuadro de dialogo
  Me.comOrigen.ShowOpen
      
  'si cancelar salgo
  If Me.comOrigen.FileName = "" Then
    Exit Sub
  End If
  
  'put path a celda
  Me.SPD.SetText 2, row, Left(Me.comOrigen.FileName, InStrRev(Me.comOrigen.FileName, "\") - 1)
   
End Sub

Private Sub SPD_LeaveCell(ByVal col As Long, ByVal row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
  
  Select Case NewRow
  
  Case 1
    Me.txtAyuda = ""
  
  Case 2
    Me.txtAyuda = ""
  
  Case 3
    Me.txtAyuda = "Debe ingresar el nombre del Dominio o del Servidor en donde estan registrados los Usuarios."
  
  Case 4
    Me.txtAyuda = "Debe ingresar el nombre del Servidor en donde esta instalado el SQL Server con los datos."
  
  Case 5
    Me.txtAyuda = "Debe ingresar el Nombre de la Base de Datos."
  
  Case 6
    Me.txtAyuda = ""
  
  Case 7
    Me.txtAyuda = "Si la conexión es local, el parametro debe ser 60. " & vbCrLf & "Si la conexión es a un punto remoto debe ser aproximadamente 240 o mayor."
  
  Case 8
    Me.txtAyuda = "Debe ingresar el Usuario, solo en caso que no haya seleccionado Seguridad Integrada."
  
  Case 9
    Me.txtAyuda = "Debe ingresar la Contraseña, solo en caso que no haya seleccionado Seguridad Integrada."
  
  Case 10
    Me.txtAyuda = "Se refiere a la cantidad de filas que cada opción muestra por default, cuando no hay filtros seleccionados. Si esta conectado a un punto remoto el valor debe ser pequeño."
  
  Case 11
    Me.txtAyuda = ""
  
  Case 12
    Me.txtAyuda = ""
  
  Case 13
    Me.txtAyuda = "Debe ingresar el Path en donde se encuentran los reportes." & vbCrLf & "Puede hacer doble click para seleccionarlo."
  
  End Select
  
End Sub
