VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form rpt_legajos_FRM 
   Caption         =   "Generación de Legajos a formato Acrobat"
   ClientHeight    =   1140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   1140
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   285
      Left            =   4350
      TabIndex        =   3
      Top             =   780
      Width           =   855
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Top             =   780
      Width           =   855
   End
   Begin VB.ComboBox cboPozos 
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   5115
   End
   Begin CRVIEWERLibCtl.CRViewer rptVisor 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   810
      Visible         =   0   'False
      Width           =   645
      lastProp        =   500
      _cx             =   1138
      _cy             =   450
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccionar Well ID"
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1485
   End
End
Attribute VB_Name = "rpt_legajos_FRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
  
  End
  
End Sub

Private Sub cmdExportar_Click()
  
  Dim rs, rs1 As ADODB.Recordset
  Dim blnB As Boolean
  Dim intI As Integer
  
  Dim strHeader, strTest, strProduction, strProductionGraf, strT, strCondicion As String
  
  'declaro objeto reporte y recordset
  Dim rptRpt As CRAXDRT.Report
  Dim appRpt As New CRAXDRT.Application
  Dim arrSubRpt, arrFormula As Variant
  
  'check si selecciono algun pozo
  If Me.cboPozos.ListIndex = -1 Then
    blnB = MsgBox("Debe seleccionar alguna opción.", vbCritical + vbOKOnly, "atención...")
    Exit Sub
  End If
  
  'inicializo
  strCondicion = ""
  
  'check si selecciono algun pozo, armo condicion
  If Me.cboPozos <> "(todos)" Then
    strCondicion = " where [well ID] = '" & Me.cboPozos & "'"
  End If
  
  'get destinos
  strHeader = ReadIni("destino", "header", App.Path & "\" & "rpt_legajos.ini")
  strTest = ReadIni("destino", "test", App.Path & "\" & "rpt_legajos.ini")
  strProduction = ReadIni("destino", "production", App.Path & "\" & "rpt_legajos.ini")
  strProductionGraf = ReadIni("destino", "productionGraf", App.Path & "\" & "rpt_legajos.ini")
    
  'mouse reloj
  Screen.MousePointer = vbHourglass
    
  'get rs con pozos operados por vintage
  strT = "select [Well ID] from IN_Legajos_vw" & strCondicion
  Set rs = SQLexec(strT)
    
  'chequeo error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    End
  End If
    
  'recorro pozos
  While Not rs.EOF
    
    'open pozos----------------------------------------------------------------------------------
    strSQL = "select * from IN_pozos_vw where [well ID] = '" & rs![well id] & "'"
    Set rs1 = SQLexec(strSQL)
    
    'chequeo error
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
      End
    End If
        
    'open report
    Set rptRpt = appRpt.OpenReport(App.Path & "\reportes\" & "in_pozos.rpt")
      
    'set rs a report
    rptRpt.Database.SetDataSource rs1
                        
    'set export PDF automaticamente
    rptRpt.ExportOptions.PDFExportAllPages = True
    rptRpt.ExportOptions.DestinationType = crEDTDiskFile
    rptRpt.ExportOptions.FormatType = crEFTPortableDocFormat
    rptRpt.EnableParameterPrompting = False
    rptRpt.ExportOptions.DiskFileName = strHeader & Replace(rs![well id], "/", "") & ".pdf"
    rptRpt.Export False
    
    'set report a Visor
    Me.rptVisor.ReportSource = rptRpt
    Me.rptVisor.ViewReport
    
    'get ensayos---------------------------------------------------------------------------------
    strSQL = "select * from IN_ensayos_vw where [well ID] = '" & rs![well id] & "'"
    Set rs1 = SQLexec(strSQL)
    
    'chequeo error
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
      End
    End If
        
    'open report
    Set rptRpt = appRpt.OpenReport(App.Path & "\reportes\" & "in_ensayos.rpt")
      
    'set rs a report
    rptRpt.Database.SetDataSource rs1
                        
    'set export PDF automaticamente
    rptRpt.ExportOptions.PDFExportAllPages = True
    rptRpt.ExportOptions.DestinationType = crEDTDiskFile
    rptRpt.ExportOptions.FormatType = crEFTPortableDocFormat
    rptRpt.EnableParameterPrompting = False
    rptRpt.ExportOptions.DiskFileName = strTest & Replace(rs![well id], "/", "") & ".pdf"
    rptRpt.Export False
    
    'set report a Visor
    Me.rptVisor.ReportSource = rptRpt
    Me.rptVisor.ViewReport
    
    'get produccion------------------------------------------------------------------------------
    strSQL = "select * from IN_produccion_vw where [well ID] = '" & rs![well id] & "'"
    Set rs1 = SQLexec(strSQL)
    
    'chequeo error
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
      End
    End If
       
    'open report
    Set rptRpt = appRpt.OpenReport(App.Path & "\reportes\" & "in_produccion.rpt")
      
    'set rs a report
    rptRpt.Database.SetDataSource rs1
                        
    'set export PDF automaticamente
    rptRpt.ExportOptions.PDFExportAllPages = True
    rptRpt.ExportOptions.DestinationType = crEDTDiskFile
    rptRpt.ExportOptions.FormatType = crEFTPortableDocFormat
    rptRpt.EnableParameterPrompting = False
    rptRpt.ExportOptions.DiskFileName = strProduction & Replace(rs![well id], "/", "") & ".pdf"
    rptRpt.Export False
    
    'set report a Visor
    Me.rptVisor.ReportSource = rptRpt
    Me.rptVisor.ViewReport
    
    'get produccion grafico---------------------------------------------------------------------
    strSQL = "select * from IN_GRAFproduccion_vw where [well ID] = '" & rs![well id] & "'"
    Set rs1 = SQLexec(strSQL)
    
    'chequeo error
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
      End
    End If
        
    'open report
    Set rptRpt = appRpt.OpenReport(App.Path & "\reportes\" & "in_graf_produccion.rpt")
      
    'set rs a report
    rptRpt.Database.SetDataSource rs1
                        
    'set export PDF automaticamente
    rptRpt.ExportOptions.PDFExportAllPages = True
    rptRpt.ExportOptions.DestinationType = crEDTDiskFile
    rptRpt.ExportOptions.FormatType = crEFTPortableDocFormat
    rptRpt.EnableParameterPrompting = False
    rptRpt.ExportOptions.DiskFileName = strProductionGraf & Replace(rs![well id], "/", "") & ".pdf"
    rptRpt.Export False
    
    'set report a Visor
    Me.rptVisor.ReportSource = rptRpt
    Me.rptVisor.ViewReport
        
    'close
    rs1.Close
        
    'puntero proximo
    rs.MoveNext
    
  Wend
    
  'close
  SQLclose
    
  'mouse default
  Screen.MousePointer = vbDefault
  
  blnB = MsgBox("La exportación se realizó con éxito.", vbInformation + vbOKOnly, "atención...")
    
End Sub

Private Sub Form_Load()
    
  Dim rs As ADODB.Recordset
  Dim strT As String
  Dim blnB As Boolean
    
  'leo parametros de conexion
  blnB = SQLgetParam()
  If Not blnB Then
    blnB = MsgBox("Los parametro son incorrectos: " & SQLparam.nombreINI, vbCritical + vbOKOnly, "atención...")
    SQLclose
    End
  End If
  
  'check roles
'  If Not SQLparam.Role = "" Then
'    SQLparam.RoleClave = "petroleo!15092004"
'  End If
    
  'geo pozos perforados por vintage
  strT = "select * from IN_Legajos_vw order by 1"
  Set rs = SQLexec(strT)
  
  'chequeo error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    End
  End If
      
  'add (todos) a comboBox
  cboPozos.AddItem "(todos)"
      
  'recorro
  While Not rs.EOF
        
    'add pozo a comboBox
    cboPozos.AddItem rs![well id]
        
    'siguiente
    rs.MoveNext
    
  Wend
  
End Sub
