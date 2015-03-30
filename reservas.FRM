VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0002E558-0000-0000-C000-000000000046}#1.0#0"; "OWC11.DLL"
Begin VB.Form reservas_FRM 
   Caption         =   "Reservas"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14745
   Icon            =   "reservas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   14745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   7935
      Left            =   4995
      TabIndex        =   21
      Top             =   1305
      Width           =   9690
      Begin OWC11.ChartSpace ChartSpace1 
         Height          =   7665
         Left            =   90
         OleObjectBlob   =   "reservas.frx":038A
         TabIndex        =   22
         Top             =   180
         Width           =   9510
      End
   End
   Begin MSComDlg.CommonDialog ComOrigen 
      Left            =   90
      Top             =   8775
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   585
      Top             =   8685
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "reservas.frx":105C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "reservas.frx":13F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "reservas.frx":1790
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "reservas.frx":1B2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "reservas.frx":1EC4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmWellID 
      Height          =   1275
      Left            =   90
      TabIndex        =   15
      Top             =   3825
      Visible         =   0   'False
      Width           =   4905
      Begin VB.CommandButton cmdBuscarWellID 
         Height          =   285
         Left            =   4320
         Picture         =   "reservas.frx":225E
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   405
         Width           =   510
      End
      Begin VB.ComboBox cmdWellID 
         Height          =   315
         Left            =   135
         TabIndex        =   18
         Top             =   405
         Width           =   4110
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   285
         Left            =   3015
         TabIndex        =   17
         Top             =   855
         Width           =   1185
      End
      Begin VB.CommandButton cmdAgregarPozo 
         Caption         =   "&Agregar"
         Height          =   285
         Left            =   1755
         TabIndex        =   16
         Top             =   855
         Width           =   1185
      End
      Begin VB.Label Label5 
         Caption         =   "Well ID"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   180
         Width           =   2085
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   14745
      _ExtentX        =   26009
      _ExtentY        =   741
      ButtonWidth     =   3281
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar pozos"
            Key             =   "imp_pozos"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   9
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "500m"
                  Text            =   "500"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "550m"
                  Text            =   "550"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "600m"
                  Text            =   "600"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "650m"
                  Text            =   "650"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "700m"
                  Text            =   "700"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "750m"
                  Text            =   "750"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "800m"
                  Text            =   "800"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "850m"
                  Text            =   "850"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "900m"
                  Text            =   "900"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agregar un pozo"
            Key             =   "agregar_pozo"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar un pozo"
            Key             =   "eliminar_pozo"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Importar producción"
            Key             =   "imp_prod_actual"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "pozo_actual"
                  Text            =   "Pozo actual desde Capitulo IV"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "todos_pozos"
                  Text            =   "Todos los Pozos desde Capitulo IV"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "pozo_actual_excel"
                  Text            =   "Pozo actual desde Excel"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Graficar curva tipo"
            Key             =   "graficar_curva"
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "36m"
                  Text            =   "36 meses"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "120m"
                  Text            =   "120 meses"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame frmProd 
      Height          =   5415
      Left            =   90
      TabIndex        =   10
      Top             =   3825
      Width           =   4905
      Begin FPSpreadADO.fpSpread spdProd 
         Height          =   4965
         Left            =   180
         TabIndex        =   11
         Top             =   360
         Width           =   4605
         _Version        =   393216
         _ExtentX        =   8123
         _ExtentY        =   8758
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
         SpreadDesigner  =   "reservas.frx":25E8
      End
      Begin VB.Label Label6 
         Caption         =   "Producción"
         Height          =   195
         Left            =   225
         TabIndex        =   12
         Top             =   135
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   90
      TabIndex        =   7
      Top             =   1305
      Width           =   4905
      Begin FPSpreadADO.fpSpread spdCercanos 
         Height          =   2085
         Left            =   180
         TabIndex        =   8
         Top             =   360
         Width           =   4605
         _Version        =   393216
         _ExtentX        =   8123
         _ExtentY        =   3678
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
         SpreadDesigner  =   "reservas.frx":27BC
      End
      Begin VB.Label Label4 
         Caption         =   "Pozos Cercanos"
         Height          =   195
         Left            =   225
         TabIndex        =   9
         Top             =   135
         Width           =   1275
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5805
      TabIndex        =   3
      Top             =   855
      Width           =   2445
   End
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Width           =   14595
      Begin VB.CommandButton cmdTamano 
         Caption         =   "Maximizar gráfico"
         Height          =   330
         Left            =   13140
         TabIndex        =   23
         Top             =   360
         Width           =   1365
      End
      Begin VB.ComboBox cboUniqueID 
         Height          =   315
         Left            =   135
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   405
         Width           =   2445
      End
      Begin VB.CommandButton cmdBuscarUniqueID 
         Height          =   285
         Left            =   2610
         Picture         =   "reservas.frx":2990
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   405
         Width           =   510
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3375
         TabIndex        =   2
         Top             =   405
         Width           =   2220
      End
      Begin VB.Label Label3 
         Caption         =   "Northing"
         Height          =   195
         Left            =   5850
         TabIndex        =   5
         Top             =   180
         Width           =   2085
      End
      Begin VB.Label Label2 
         Caption         =   "Easting"
         Height          =   195
         Left            =   3420
         TabIndex        =   4
         Top             =   180
         Width           =   2085
      End
      Begin VB.Label Label1 
         Caption         =   "Unique ID"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   180
         Width           =   2085
      End
   End
End
Attribute VB_Name = "reservas_FRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngLeft, lngWidth As Long
Dim blnMax As Boolean

'fill grilla con produccion segun pozo seleccionado
'
Private Sub graficarCurva(ByVal strMeses As String)
  Dim strT, strWellIDaux As String
  Dim intI, intPozos As Integer
  Dim rsCur, rsCer As ADODB.Recordset
    
  'dim arrays
  Dim asSeriesNames(1)
  Dim aiValues()
    
  'check si no selecciono ubicacion, exit
  If Me.cboUniqueID.ListIndex < 0 Then
    
    intI = MsgBox("No ha seleccionado un Unique ID aún.", vbCritical + vbApplicationModal, "Atención...")
    Exit Sub
    
  End If
    
  'puntero mouse reloj
  Screen.MousePointer = vbHourglass
    
  strT = "select * from ubicacionesCurvasGra_vw where uniqueID = '" & Me.cboUniqueID.Text & "' order by orden"
    
  'get datos
  Set rsCur = SQLexec(strT)
  
  'check error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    
    'recupero puntero mouse
    Screen.MousePointer = vbDefault
    
    End
  End If
    
  strT = "select * from ubicacionesNorma_vw where IDcercano = " & Me.cboUniqueID.ItemData(Me.cboUniqueID.ListIndex) & " order by wellID, orden"
    
  'get datos
  Set rsCer = SQLexec(strT)
  
  'check error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    
    'recupero puntero mouse
    Screen.MousePointer = vbDefault
    
    End
  End If
        
  '
  'comienzo de sistema de graficación
  '
  
  'initialize
  intI = 0
  
  'while valores curva tipo
  Do While Not rsCur.EOF And intI < Val(strMeses)
           
    'check si primera pasada
    If intI = 0 Then
    
      'asign nombre a serie
      asSeriesNames(0) = rsCur!nombre
      
    End If
       
    'add 1 a contador
    intI = intI + 1
    
    'redim array
    ReDim Preserve aiValues(intI)
    
    'asign valor
    aiValues(intI) = rsCur!OILdia
       
    'siguiente
    rsCur.MoveNext
       
  Loop
      
  'clear grafico
  ChartSpace1.Clear
      
  'build grafico dinamicamente
  Set chConstants = ChartSpace1.Constants
  
  'Add grafico
  Set chtnewchart = ChartSpace1.Charts.Add
      
  'set color fondo blanco
  chtnewchart.PlotArea.Interior.Color = "white"
      
  'Specify tipo de chart
  chtnewchart.Type = chConstants.chChartTypeLineMarkers
  
  'asign nombre de serie
  chtnewchart.SetData chConstants.chDimSeriesNames, chConstants.chDataLiteral, asSeriesNames
  
  'asign valores a serie1 curva tipo
  chtnewchart.SeriesCollection(0).SetData chConstants.chDimValues, chConstants.chDataLiteral, aiValues
      
  'set tañamo de marca de referencia
  chtnewchart.SeriesCollection(0).Marker.Size = 2
  
  'set color para linea, verde
  chtnewchart.SeriesCollection(0).Line.Color = RGB(0, 100, 0)
  
  'set color blanco interior
  chtnewchart.Interior.Color = RGB(255, 255, 255)
    
  'initialize
  intPozos = 0
  strWellIDaux = ""
    
  'while produccion pozos cercanos
  Do While Not rsCer.EOF
           
    'clear array
    For intI = 0 To Val(strMeses)
      aiValues(intI) = Null
    Next
    
    'initialize
    intI = 0
       
    'add 1 a contador pozos
    intPozos = intPozos + 1
    
    'asign nombre serie
    asSeriesNames(0) = rsCer!WellID
       
    'save wellID anterior
    strWellIDaux = rsCer!WellID
       
    'while not EOF y mismo wellID
    Do While Not rsCer.EOF
       
      'check si wellID actual <> wellID anterior
      If rsCer!WellID <> strWellIDaux Then
        Exit Do
      End If
       
      'add 1 a contador
      intI = intI + 1
            
      'check contador de valores <= meses a graficar
      If intI < Val(strMeses) Then
        
        'asign valor
        aiValues(intI) = rsCer!OILmes
        
      End If
       
      'siguiente
      rsCer.MoveNext
       
    Loop
    
    'add serie nueva
    chtnewchart.SeriesCollection.Add
    
    'asign nombre de serie
    chtnewchart.SeriesCollection(intPozos).SetData chConstants.chDimSeriesNames, chConstants.chDataLiteral, asSeriesNames
    
    'asign valores a serie2...3...4...etc. para pozos cercanos
    chtnewchart.SeriesCollection(intPozos).SetData chConstants.chDimValues, chConstants.chDataLiteral, aiValues
      
    'set tañamo de marca de referencia
    chtnewchart.SeriesCollection(intPozos).Marker.Size = 2
    
'    'set color para linea, verde
'    Me.ChartSpace1.Charts(0).SeriesCollection(intPozos).Line.Color = RGB(0, 100, 0)
    
  Loop
    
  'set habilita logatitmo, lineas horizontales, color gris, unidad 10, titulo, tamaño letra
  chtnewchart.Axes(chConstants.chAxisPositionValue).Scaling.Type = chScaleTypeLogarithmic
  chtnewchart.Axes(chConstants.chAxisPositionValue).HasMinorGridlines = True
  chtnewchart.Axes(chConstants.chAxisPositionValue).MajorGridlines.Line.Color = RGB(220, 220, 220)
  chtnewchart.Axes(chConstants.chAxisPositionValue).MinorGridlines.Line.Color = RGB(220, 220, 220)
  chtnewchart.Axes(chConstants.chAxisPositionValue).MajorUnit = 10
  chtnewchart.Axes(chConstants.chAxisPositionValue).MinorUnit = 1
  chtnewchart.Axes(chConstants.chAxisPositionValue).HasTitle = True
  chtnewchart.Axes(chConstants.chAxisPositionValue).Title.Font.Size = 9
  chtnewchart.Axes(chConstants.chAxisPositionValue).Title.Font.Bold = True
  chtnewchart.Axes(chConstants.chAxisPositionValue).Title.Caption = "Rate (m3opd)"
  
  'set habilita logatitmo, lineas horizontales, color gris, unidad 10, titulo, tamaño letra
  chtnewchart.Axes(chConstants.chAxisPositionCategory).HasTitle = True
  chtnewchart.Axes(chConstants.chAxisPositionCategory).Title.Font.Size = 9
  chtnewchart.Axes(chConstants.chAxisPositionCategory).Title.Font.Bold = True
  chtnewchart.Axes(chConstants.chAxisPositionCategory).Title.Caption = "Time months"
  
  'check si 36 meses, set espacio 1, sino 5
  If Val(strMeses) = 36 Then
    chtnewchart.Axes(chConstants.chAxisPositionCategory).TickLabelSpacing = 1
  Else
    chtnewchart.Axes(chConstants.chAxisPositionCategory).TickLabelSpacing = 5
  End If
      
  'set leyenda series
  chtnewchart.HasLegend = True
  chtnewchart.Legend.Position = chLegendPositionTop
      
  'close
  SQLclose
  
  'recupero puntero mouse
  Screen.MousePointer = vbDefault
      
End Sub


'fill grilla con pozos cercanos, segun ubicacion seleccionada en comboBox
'
Private Sub fillPozos(ByVal intIDubicacion As Integer)
  Dim rs As ADODB.Recordset
    
  'build query
  strT = "select * from ubiBuscaCercanos_vw where IDubicacion = " & str(intIDubicacion) & " order by [oil (m3/d)] desc"
    
  'get datos
  Set rs = SQLexec(strT)
  
  'check error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    End
  End If
        
  'clear
  Me.spdCercanos.ClearRange 1, 1, Me.spdCercanos.MaxRows, Me.spdCercanos.MaxCols, True
  
  'assign rs a grilla
  Me.spdCercanos.DataSource = rs.DataSource
    
  'close
  SQLclose
     
End Sub

'fill grilla con produccion segun pozo seleccionado
'
Private Sub fillPozoPM(ByVal intIDcercano As Integer, ByVal intIDpozo As Integer)
  Dim rs As ADODB.Recordset
    
  'puntero mouse reloj
  Screen.MousePointer = vbHourglass
    
  'build query
  strT = "select * from ubiBuscaCercanosPM_vw where IDcercano = " & str(intIDcercano) & " and IDpozo = " & intIDpozo & " order by date desc"
    
  'get datos
  Set rs = SQLexec(strT)
  
  'check error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    
    'recupero puntero mouse
    Screen.MousePointer = vbDefault
    
    End
  End If
        
  'activa cambio en background
  Me.spdProd.ReDraw = False
        
  'clear
  Me.spdProd.ClearRange 1, 1, Me.spdProd.MaxRows, Me.spdProd.MaxCols, True
  
  'assign rs a grilla
  Me.spdProd.DataSource = rs.DataSource
  
  'activa cambio en foreground
  Me.spdProd.ReDraw = True
  
  'close
  SQLclose
     
  'recupero puntero mouse
  Screen.MousePointer = vbDefault
      
End Sub


'llena comboBox con ubicaciones, permite busqueda
'
Private Sub fillUbicaciones(ByVal strCriterio As String)
    
  'build condicion
  If strCriterio = "" Then
    
    strCriterio = "%"
    
  Else
    
    strCriterio = "%" & strCriterio & "%"
    
  End If
    
  'build query
  strT = "select id, uniqueID, IDcurva from ubicaciones where uniqueID like '" & strCriterio & "' order by uniqueID"
    
  'get ubicaciones
  Set rs = SQLexec(strT)
  
  'chequeo error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    End
  End If
        
  'clear combo
  Me.cboUniqueID.Clear
  
  'recorro
  While Not rs.EOF
        
    'add texto y id a combo
    Me.cboUniqueID.AddItem rs!uniqueID
    Me.cboUniqueID.ItemData(Me.cboUniqueID.NewIndex) = rs!ID
          
    'siguiente
    rs.MoveNext
    
  Wend
   
  'close
  SQLclose
   
End Sub

'get pozos Cercanos segun ubicacion seleccionada en comboBox
'
Private Sub cboUniqueID_Click()
  Dim blnB As Boolean
    
  'check si no selecciono item, exit
  If Me.cboUniqueID.ListIndex < 0 Then
    Exit Sub
  End If
    
  'fill cercanos
  Call fillPozos(Me.cboUniqueID.ItemData(Me.cboUniqueID.ListIndex))
  
End Sub


'importa pozos cercanos segun distancia
'
Private Sub importarPozos(ByVal strDistancia As String)
  Dim intI As Integer
  
  'check si no selecciono ubicacion, exit
  If Me.cboUniqueID.ListIndex < 0 Then
    
    intI = MsgBox("No ha seleccionado Unique ID aún.", vbCritical + vbApplicationModal, "Atención...")
    Exit Sub
    
  End If
    
  'check si existen pozos
  If Me.spdCercanos.DataRowCnt > 0 Then
    
    intI = MsgBox("Se encontraron pozos ya importados, si desea continuar, los mismos serán reemplazados.", vbQuestion + vbYesNo, "Atención...")
    If intI = 7 Then
      Exit Sub
    End If
    
  End If
  
  'build
  strT = "exec ubiImportarPozos_INS_sp " & Me.cboUniqueID.ItemData(Me.cboUniqueID.ListIndex) & "," & strDistancia
    
  'exec
  Set rs = SQLexec(strT)
  
  'chequeo error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    End
  End If
         
  'call cercanos
  Call fillPozos(Me.cboUniqueID.ItemData(Me.cboUniqueID.ListIndex))

End Sub


'importa producción mensual T: para todos los pozos, A: para pozo actual
'
Private Sub importarProdCapituloIV(ByVal strCriterio As String)
  
  Dim intRow, intRow1, intRow2, intI As Integer
  Dim varIDcercano, varIDpozo As Variant
  Dim rs As ADODB.Recordset
    
  'check si puede importar
  If Me.spdCercanos.DataRowCnt = 0 Then
    
    intI = MsgBox("No hay pozos dsiponibles.", vbCritical + vbApplicationModal, "Atención...")
    Exit Sub
    
  End If
    
  'puntero mouse reloj
  Screen.MousePointer = vbHourglass
    
  'set limites
  If strCriterio = "T" Then
    intRow1 = 1
    intRow2 = Me.spdCercanos.DataRowCnt
  Else
    intRow1 = Me.spdCercanos.ActiveRow
    intRow2 = Me.spdCercanos.ActiveRow
  End If
  
  'while pozos cercanos
  For intRow = intRow1 To intRow2
    
    'get IDubicacion
    Me.spdCercanos.GetText 5, intRow, varIDcercano
    
    'get IDpozo
    Me.spdCercanos.GetText 6, intRow, varIDpozo
        
    'build
    strT = "exec ubiImportarProdCapituloIV_INS_sp " & varIDcercano & "," & varIDpozo
    
    'exec
    Set rs = SQLexec(strT)
    
    'chequeo error
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
      
      'recupero puntero mouse
      Screen.MousePointer = vbDefault
      
      End
    End If
        
  Next
    
  'close
  SQLclose
        
  'recupero puntero mouse
  Screen.MousePointer = vbDefault
          
  'show mensaje
  intI = MsgBox("La importación se realizó con éxito.", vbInformation + vbApplicationModal, "Atención...")
          
  'fill grilla con produccion para pozo cercano actual
  Call fillPozoPM(varIDcercano, varIDpozo)
          
End Sub

'importa producción mensual para todos los pozos desde axcel
'
Private Sub importarProdExcel()
  Dim intRow, intI  As Integer
  Dim strWellID As String
  Dim dtmFecha As Date
  Dim sngVolumen As Single
  Dim rs As ADODB.Recordset
  Dim excApp  As Excel.Application
  Dim excInfo As Excel.Workbook
  Dim varIDcercano, varIDpozo As Variant
  
  'check si puede importar
  If Me.spdCercanos.DataRowCnt = 0 Then
    
    intI = MsgBox("No hay pozos dsiponibles.", vbCritical + vbApplicationModal, "Atención...")
    Exit Sub
    
  End If
  
  
  
  'si formato excel filtro xls
  Me.ComOrigen.Filter = "Archivos de Excel (*.xls)|*.xls"
    
  'titulo de ventana
  Me.ComOrigen.DialogTitle = "Seleccionar archivo..."
  
  Me.ComOrigen.FileName = ""
  
  'abro cuadro de dialogo
  Me.ComOrigen.ShowOpen
      
  'si cancelar salgo
  If Me.ComOrigen.FileName = "" Then
    Exit Sub
  End If
    
  'open excel
  Set excApp = New Excel.Application
  Set excInfo = excApp.Workbooks.Open(Me.ComOrigen.FileName)
      
  'check error de apertura
      
      
      
  'puntero mouse reloj
  Screen.MousePointer = vbHourglass
      
  'get IDubicacion
  Me.spdCercanos.GetText 5, Me.spdCercanos.ActiveRow, varIDcercano
      
  'get IDpozo
  Me.spdCercanos.GetText 6, Me.spdCercanos.ActiveRow, varIDpozo
      
  'build, delete pozo
  strT = "exec ubiImportarProdExcel_ELI_sp " & varIDcercano & "," & varIDpozo
              
  'exec
  Set rs = SQLexec(strT)
        
  'chequeo error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
          
    'recupero puntero mouse
    Screen.MousePointer = vbDefault
          
    End
  End If
      
      
      
  'initialize
  intRow = 2
  
  'while excel cuando no encuentra wellID, corta
  Do While excInfo.Sheets.Application.Cells(intRow, 1).Text <> ""
          
    'save wellID, fecha, volumen
    strWellID = excInfo.Sheets.Application.Cells(intRow, 1).Text
    dtmFecha = excInfo.Sheets.Application.Cells(intRow, 2).Text
    sngVolumen = excInfo.Sheets.Application.Cells(intRow, 3).Text
    
    'check si null, asign 0
    If IsNull(sngVolumen) Then
      sngVolumen = 0
    End If
      
    'check si wellID de excel = a WellID de grilla pozos
    intI = Me.spdCercanos.SearchCol(1, Me.spdCercanos.ActiveRow - 1, Me.spdCercanos.ActiveRow, strWellID, SearchFlagsNone)
        
    'check si encontro WellID en grilla pozos
    If intI > 0 Then
            
      'build, insert pozo
      strT = "exec ubiImportarProdExcel_INS_sp " & varIDcercano & "," & varIDpozo & ",'" & dateToIso(dtmFecha) & "'," & sngVolumen
            
      'exec
      Set rs = SQLexec(strT)
      
      'chequeo error
      If Not SQLparam.CnErrNumero = -1 Then
        SQLError
        SQLclose
        
        'recupero puntero mouse
        Screen.MousePointer = vbDefault
        
      End
      End If
            
    End If
      
    'add 1 contador fila
    intRow = intRow + 1
      
  Loop
  
  'close Workbook y Excel
  excInfo.Close
  excApp.Quit
   
  'recupero puntero mouse
  Screen.MousePointer = vbDefault
  
  intI = MsgBox("La importación se realizó con éxito..", vbInformation + vbApplicationModal, "Atención...")
  
  'fill grilla con produccion
  Call fillPozoPM(varIDcercano, varIDpozo)
          
End Sub


'agrega pozo en forma manual
'
Private Sub cmdAgregarPozo_Click()
  Dim intI As Integer
  
  'check si no selecciono ubicacion, exit
  If Me.cmdWellID.ListIndex < 0 Then
    
    intI = MsgBox("No ha seleccionado un WellID aún.", vbCritical + vbApplicationModal, "Atención...")
    Exit Sub
    
  End If
    
  'build
  strT = "exec ubiImportarPozosManual_INS_sp " & Me.cboUniqueID.ItemData(Me.cboUniqueID.ListIndex) & "," & Me.cmdWellID.ItemData(Me.cmdWellID.ListIndex)
    
  'exec
  Set rs = SQLexec(strT)
  
  'chequeo error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    End
  End If
             
  'fill cercanos
  Call fillPozos(Me.cboUniqueID.ItemData(Me.cboUniqueID.ListIndex))
    
  'close
  SQLclose
  
  'hide
  Me.frmWellID.Visible = False
  
  'show
  Me.frmProd.Visible = True
     
End Sub

Private Sub cmdBuscarUniqueID_Click()
  
  'get ubicaciones
  Call fillUbicaciones(Me.cboUniqueID)
      
End Sub


Private Sub Command1_Click()

End Sub

'llena comboBox con ubicaciones, permite busqueda
'
Private Sub cmdBuscarWellID_Click()
  Dim strCriterio As String
  Dim rs As ADODB.Recordset
    
  'build condicion
  If Me.cmdWellID = "" Then
    
    strCriterio = "%"
  
  Else
    
    strCriterio = "%" & Me.cmdWellID & "%"
  
  End If
    
  'build query
  strT = "select id, nombre from pozos where nombre like '" & strCriterio & "' order by nombre"
    
  'get ubicaciones
  Set rs = SQLexec(strT)
  
  'chequeo error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    End
  End If
        
  'clear combo
  Me.cmdWellID.Clear
  
  'recorro
  While Not rs.EOF
        
    'add texto y id a combo
    Me.cmdWellID.AddItem rs!nombre
    Me.cmdWellID.ItemData(Me.cmdWellID.NewIndex) = rs!ID
          
    'siguiente
    rs.MoveNext
    
  Wend
   
  'close
  SQLclose
   
End Sub

Private Sub cmdCancelar_Click()
  
  'hide
  Me.frmWellID.Visible = False
  
  'show
  Me.frmProd.Visible = True
    
    
End Sub

Private Sub cmdTamano_Click()
  
  If Not blnMax Then
    
    'save tamaño
    lngLeft = Me.Frame3.Left
    lngWidth = Me.Frame3.Width
      
    'change tamaño frame
    Me.Frame3.Left = Me.ScaleLeft + 100
    Me.Frame3.Width = Me.Width - 290
    
    'change tamaño grafico
    Me.ChartSpace1.Left = Me.Frame3.Left
    Me.ChartSpace1.Width = Me.Frame3.Width - 190
    
    'change label
    Me.cmdTamano.Caption = "Minimizar gráfico"
    
    'change bandera
    blnMax = True
    
  Else
    
    'change tamaño frame
    Me.Frame3.Left = lngLeft
    Me.Frame3.Width = lngWidth
    
    'change tamaño grafico
    Me.ChartSpace1.Width = Me.Frame3.Width - 190
    
    'change label
    Me.cmdTamano.Caption = "Maximizar gráfico"
    
    'change bandera
    blnMax = False
    
    
  End If
  
End Sub

Private Sub Form_Load()

  Dim rs As ADODB.Recordset
  Dim strT, strL As String
  Dim intI As Integer
  Dim blnB As Boolean
  
  'check configuracion regional
  If Not checkConfigRegional() Then
    blnB = MsgBox("El sistema detecto que la configuración regional no es correcta." & vbCrLf & vbCrLf & _
           "Configurar el formato para números de esta forma: 123,456,789.00." & vbCrLf & vbCrLf & _
           "Configurar el formato para fechas  de esta forma: dd/MM/yyyy.", vbCritical + vbOKOnly, "Atención...")
    End
  End If
  
  'get parametros de conexion
  blnB = SQLgetParam()
  
  'check si parametros ok, get menu
  If blnB Then
    
  End If
      
  'get ubicaciones
  Call fillUbicaciones("")
          
  Me.spdCercanos.OperationMode = OperationModeRow
  Me.spdCercanos.UnitType = UnitTypeTwips
  Me.spdCercanos.RowHeight(0) = 250
  
  Me.spdProd.UnitType = UnitTypeTwips
  Me.spdProd.RowHeight(0) = 250
  
            
End Sub

Private Sub spdCercanos_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
  Dim varIDcercano, varIDpozo As Variant
  
  If Me.spdCercanos.DataRowCnt = 0 Then
    Exit Sub
  End If
  
  'get IDcercano y IDubicacion
  Me.spdCercanos.GetText 5, NewRow, varIDcercano
  Me.spdCercanos.GetText 6, NewRow, varIDpozo
  
  'get produccion para pozo cercano actual
  Call fillPozoPM(varIDcercano, varIDpozo)
  
End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim blnB As Boolean

'check opción seleccionada------------------------------------------
Select Case Button.Key
  

'agregar pozo manual, habilitar
Case "agregar_pozo"
  
  'hide
  Me.frmWellID.Visible = True
  
  'show
  Me.frmProd.Visible = False
  
  
  
  
Case "eliminar_pozo"
  
  Dim intI As Integer
  Dim varIDcercano, varIDpozo As Variant
  
  'check si selecciono pozo
  If Me.spdCercanos.ActiveRow < 0 Or Me.spdCercanos.ActiveRow > Me.spdCercanos.DataRowCnt Then
      
    intI = MsgBox("No ha seleccionado un pozo aún.", vbCritical + vbApplicationModal, "Atención...")
    Exit Sub
        
  End If
  
  'get IDubicacion y IDpozo
  Me.spdCercanos.GetText 5, Me.spdCercanos.ActiveRow, varIDcercano
  Me.spdCercanos.GetText 6, Me.spdCercanos.ActiveRow, varIDpozo
  
  'show adventencia
  intI = MsgBox("Desea eliminar toda la información correspondiente al pozo seleccionado.", vbQuestion + vbYesNo, "Atención...")
  If intI = 6 Then
           
    'build
    strT = "exec ubiImportarPozosManual_ELI_sp " & varIDcercano & "," & varIDpozo
    
    'exec
    Set rs = SQLexec(strT)
    
    'chequeo error
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
      End
    End If
      
    'close
    SQLclose
        
    'delete fila en grilla
    Me.spdCercanos.InsertRows Me.spdCercanos.ActiveRow, 1
    Me.spdCercanos.DeleteRows Me.spdCercanos.ActiveRow, 2
    Me.spdCercanos.MaxRows = Me.spdCercanos.MaxRows - 2
      
    'get IDcercano y IDubicacion
    Me.spdCercanos.GetText 5, Me.spdCercanos.ActiveRow, varIDcercano
    Me.spdCercanos.GetText 6, Me.spdCercanos.ActiveRow, varIDpozo
    
    'get produccion para pozo cercano actual
    Call fillPozoPM(varIDcercano, varIDpozo)
        
  End If
    
End Select

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  
  'check opción
  Select Case LCase(ButtonMenu.Key)
    
  
  
  'importa produccion pozo actual
  Case "pozo_actual"
    
    'importa produccion para pozo actual desde Capitulo IV
    Call importarProdCapituloIV("A")
    
    
    
  'importa produccion pozo todas desde Capitulo IV
  Case "todos_pozos"
    
    'importa produccion para todos los pozos
    Call importarProdCapituloIV("T")
    
    
    
  'importa produccion pozo todas
  Case "pozo_actual_excel"
    
    'importa produccion desde Excel
    Call importarProdExcel
    
    
    
  'importa produccion pozo tod
  Case "500m", "550m", "600m", "650m", "700m", "750m", "800m", "850m", "900m"
    
    'importa pozos cercanos segun distancia
    importarPozos (ButtonMenu.Text)
    
  'graficar curvas tipo
  Case "36m", "120m"
    
    'importa pozos cercanos segun distancia
    graficarCurva (ButtonMenu.Text)
    
  End Select
    
    
End Sub
