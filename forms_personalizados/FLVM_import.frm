VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form LF_VM_FRM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flow Log and Valves Movement - Export / Import"
   ClientHeight    =   9720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15840
   Icon            =   "FLVM_import.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   15840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   2265
      Left            =   90
      TabIndex        =   9
      Top             =   45
      Width           =   7530
      Begin VB.TextBox cmd_Destino 
         Height          =   330
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1800
         Width           =   7170
      End
      Begin VB.TextBox txt_Pozo 
         Height          =   330
         Left            =   180
         TabIndex        =   0
         Text            =   "Ingrese el Well Name del Pozo, para el cual, desea generar los Template"
         Top             =   225
         Width           =   5370
      End
      Begin VB.CommandButton cmd_valves_movement_generate 
         Caption         =   "Generate Excel for Valves Movement"
         Height          =   600
         Left            =   2610
         Picture         =   "FLVM_import.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   2940
      End
      Begin VB.CommandButton cmd_flow_log_generate 
         Caption         =   "Generate Excel for Flow Log"
         Height          =   600
         Left            =   180
         Picture         =   "FLVM_import.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   720
         Width           =   2310
      End
      Begin MSComDlg.CommonDialog comOrigen 
         Left            =   6705
         Top             =   855
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "La plantilla generada, la puede encontrar en la carpeta siguiente"
         Height          =   285
         Left            =   225
         TabIndex        =   13
         Top             =   1575
         Width           =   6765
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7260
      Left            =   90
      TabIndex        =   8
      Top             =   2340
      Width           =   15630
      Begin FPSpreadADO.fpSpread spdSQL 
         Height          =   6990
         Left            =   135
         TabIndex        =   6
         Top             =   180
         Width           =   15360
         _Version        =   393216
         _ExtentX        =   27093
         _ExtentY        =   12330
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
         SpreadDesigner  =   "FLVM_import.frx":0A9E
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2265
      Left            =   7695
      TabIndex        =   7
      Top             =   45
      Width           =   8025
      Begin VB.TextBox cmd_Origen 
         Height          =   330
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
         Width           =   7665
      End
      Begin VB.CommandButton cmd_Salir 
         Caption         =   "&Salir"
         Height          =   330
         Left            =   6390
         TabIndex        =   10
         Top             =   180
         Width           =   1455
      End
      Begin VB.CommandButton cmd_Importar 
         Caption         =   "&Importar"
         Height          =   600
         Left            =   6390
         Picture         =   "FLVM_import.frx":0C72
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmd_movement_valves 
         Caption         =   "Open Valves Movement Excel for Import"
         Height          =   600
         Left            =   3015
         Picture         =   "FLVM_import.frx":0FFC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   3210
      End
      Begin VB.CommandButton cmd_open_flow_log 
         Caption         =   "Open Flow Log Excel for Import"
         Height          =   600
         Left            =   180
         Picture         =   "FLVM_import.frx":1386
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   2670
      End
   End
End
Attribute VB_Name = "LF_VM_FRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngLeft, lngWidth As Long
Dim blnMax As Boolean

Dim intFilTitu, lngFmaxima, lngCmaxima As Integer
Dim colForm, colQpHk, colCeme, colRese, colComienzo, filComienzo, colEntreCadaOperacion As Integer


Private Sub cmd_flow_log_Click()

End Sub

Private Sub cmd_flow_log_generate_Click()


  Dim blnB, blnRowPainted, blnFilaConMerge As Boolean
  Dim strT As String
  Dim lngRow, lngCol, lngColVisible As Long
  Dim varDato, varDatoColRig, varDatoColRigOficial As Variant
  
  Dim str_Date_Master, str_Date_Assembly As String
  Dim int_Id_Well, int_Row_Start, int_Row_End, int_Madril_Aux, int_Merge_Start, int_Merge_End As Integer
  
  
  Dim Obj As Object
  Dim Libro As Object
  Dim Hoja As Object
  
  
    '--------------------------------------------------------------------------------------
    'WELL NAME - CHEQUEO EXISTENCIA
    '--------------------------------------------------------------------------------------
    
    'GET Date Master
    strT = "select WellName FROM pozosPropios_vw WHERE WellName = '" + txt_Pozo + "'"
    Set rs = SQLexec(strT)
    
    'CHECK errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
    End If
    
    'CHECK si encontro algo
    If rs.EOF Then
                
        blnB = MsgBox("El Pozo ingresado no existe..", vbExclamation, "atención...")
        
        Exit Sub
        
    End If
  
  
  
  'mouse reloj
  Screen.MousePointer = vbHourglass
  
    'CREATE excel
    Set Obj = CreateObject("Excel.application")
    

    Set Libro = Obj.Workbooks.Open(App.Path + "\TEMPLATE\Flowlog_YYYYMMDD_POZO.xlsx")
    Set Hoja = Libro.Sheets(1)
    
    
    'ENABLE excel visible
    Obj.Visible = False
        
    'DISABLE proteccion
    Hoja.UnProtect Password:="Sinopec2015"
        
        
    '--------------------------------------------------------------------------------------
    'FLOW LOG GET DATE MASTER
    '--------------------------------------------------------------------------------------
    
    'GET Date Master
    strT = "select * FROM Injection_Template_GET_MAX_Date_Master_vw WHERE Well_Name = '" + Me.txt_Pozo + "'"
    Set rs = SQLexec(strT)
    
    'CHECK errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
    End If
    
    'CHECK si encontro algo
    If rs.EOF Then
                
        blnB = MsgBox("Date Master, no se encuentra en Targets - Main, para el Pozo: " + var_WellID + ". No puede continuar con la generacion de la Plantilla.", vbExclamation, "atención...")
        
        'mouse defa
        Screen.MousePointer = vbDefault
        
        Exit Sub
        
    End If
    
    'SAVE Date Master, para usarlo en los procesos siguientes
    str_Date_Master = dateToIso(rs!Date_Master)
        
    'SAVE ID Well, para usarlo en los procesos siguientes
    int_Id_Well = rs!Id_Well
    
        
    '--------------------------------------------------------------------------------------
    'FLOW LOG - GENERAL
    '--------------------------------------------------------------------------------------
    
    'SET puntero en fila donde se encuentra la info para MAIN
    lngRow = 12
            
            
    'GET SUM Qwi Target
    strT = "select * FROM Injection_Template_GET_SUM_Qwi_Target_vw WHERE ID_Well = " + str(int_Id_Well) + " AND " + "Date_Master = '" + str_Date_Master + "'"
    Set rs = SQLexec(strT)
    
    'CHECK errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
    End If
    
    'CHECK si encontro algo
    If rs.EOF Then
                
        blnB = MsgBox("Qwi Target, no se encuentra en Targets - Main, para el Pozo: " + var_WellID + ". No puede continuar con la generacion de la Plantilla.", vbExclamation, "atención...")
        
        'mouse defa
        Screen.MousePointer = vbDefault
        
        Exit Sub
        
    End If
    
    
    'PUT - Well Name
    Hoja.cells(lngRow, 3) = Me.txt_Pozo
    
    'PUT - Caudal Target
    Hoja.cells(lngRow, 6) = rs!Qwi_Target
    
    
    '--------------------------------------------------------------------------------------
    'FLOW LOG - AUXILIARY
    '--------------------------------------------------------------------------------------
    
    'SET puntero en fila donde se encuentra la info para AUXILIARY
    lngRow = 37
    
    
    'GET Date Master
    strT = "select * FROM Injection_Template_GET_MAX_Packers_vw WHERE Id_Well = " + str(int_Id_Well) + " AND " + "Date_Master = '" + str_Date_Master + "' ORDER BY Packer_Depth"
    Set rs = SQLexec(strT)
    
    'CHECK errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
    End If
    
    'CHECK si encontro algo
    If rs.EOF Then
                
        blnB = MsgBox("Los Packers, no se encuentran en Targets - Main, para el Pozo: " + var_WellID + ". No puede continuar con la generacion de la Plantilla.", vbExclamation, "atención...")
        
        'mouse defa
        Screen.MousePointer = vbDefault
        
        Exit Sub
        
    End If
    
    
    'WHILE Mandrels
    Do While Not rs.EOF
        
        'BUILD string con el rango, en donde debo seleccionar
        strT = Replace(str(lngRow) + ":" + str(lngRow), " ", "")
        
        'SELECT - COPY fila completa
        Hoja.Rows(strT).Select
'        Obj.CutCopyMode = False
        Obj.Selection.Copy
        
        'BUILD string con el rango, en donde debo copiar
        strT = Replace(str(lngRow + 1) + ":" + str(lngRow + 1), " ", "")
        
        Hoja.Rows(strT).Select
         
        'PASTE fila completa
        Obj.Selection.Insert
        'Shift:=xlDown
        
        'ASSING informacion a las celdas
        Hoja.cells(lngRow, 3) = rs!Packer_Number
        Hoja.cells(lngRow, 4) = rs!Packer_Depth
        
        lngRow = lngRow + 1
        
        'MOVE al siguiente registro
        rs.MoveNext
        
    Loop
    
    
    
    '--------------------------------------------------------------------------------------
    'FLOW LOG - MAIN
    '--------------------------------------------------------------------------------------
    
    'SET puntero en fila donde se encuentra la info para MAIN
    lngRow = 26
    
    'SAVE fila de comienzo, para cuando finalizo la generaciòn de Main, armo el MERGE en la columna E - Caudal Admisión Mandril [m3/d]
    int_Row_Start = lngRow
    
    
    'GET Date Master
    strT = "select * FROM Injection_Template_GET_MAX_Mandrels_vw WHERE Id_Well = " + str(int_Id_Well) + " AND " + "Date_Master = '" + str_Date_Master + "' ORDER BY [Top]"
    Set rs = SQLexec(strT)
    
    'CHECK errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
    End If
    
    'CHECK si encontro algo
    If rs.EOF Then
                
        blnB = MsgBox("Los Mandrels, no se encuentra en Targets - Main, para el Pozo: " + var_WellID + ". No puede continuar con la generacion de la Plantilla.", vbExclamation, "atención...")
        
        'mouse defa
        Screen.MousePointer = vbDefault
        
        Exit Sub
        
    End If
    
    
    'WHILE Mandrels
    Do While Not rs.EOF
        
        'BUILD string con el rango, en donde debo seleccionar
        strT = Replace(str(lngRow) + ":" + str(lngRow), " ", "")
        
        'SELECT - COPY fila completa
        Hoja.Rows(strT).Select
'        Obj.CutCopyMode = False
        Obj.Selection.Copy
        
        'BUILD string con el rango, en donde debo copiar
        strT = Replace(str(lngRow + 1) + ":" + str(lngRow + 1), " ", "")
        
        Hoja.Rows(strT).Select
         
        'PASTE fila completa
        Obj.Selection.Insert
        'Shift:=xlDown
        
        'ASSING informacion a las celdas
        Hoja.cells(lngRow, 3) = rs!Mandrel_Number
        Hoja.cells(lngRow, 4) = rs!Mandrel_Depth
        Hoja.cells(lngRow, 6) = rs!Reservoir
        Hoja.cells(lngRow, 7) = rs!Top
        Hoja.cells(lngRow, 8) = rs!Base
        
        
        lngRow = lngRow + 1
        
        'MOVE al siguiente registro
        rs.MoveNext
        
    Loop
    
    
    'SAVE fila de fin, para cuando finalizo la generaciòn de Main, armo el merge en la columna E - Caudal Admisión Mandril [m3/d]
    int_Row_End = lngRow
    
    'MERGE en la columna E - Caudal Admisión Mandril [m3/d], mientras se repite el nùmero de mandril, MERGE de las filas
    '---------------------------------------------
    
    'ASSIGN fila de comienzo
    lngRow = int_Row_Start
    
    'ASSING fila de comienzo para MERGE
    int_Merge_Start = lngRow
    
    'SAVE Nùmero de Mandril en un auxiliar, para poder saber si se repite
    int_Madril_Aux = Hoja.cells(lngRow, 3)
    
    'WHILE Mandrels
    Do While lngRow <= int_Row_End
        
        If Hoja.cells(lngRow, 3) <> int_Madril_Aux Then
        
            'MERGE
            
            'BUILD string con el rango, en donde debo seleccionar
            strT = "E" + Replace(str(int_Merge_Start) + ":" + "E" + str(int_Merge_End), " ", "")
            
            'SELECT - celdas
            Hoja.Range(strT).Select
            
            'MERGE
            Obj.Selection.Merge
            
            'ASSING fila de comienzo para MERGE
            int_Merge_Start = lngRow
            
            'SAVE Nùmero de Mandril en un auxiliar, para poder saber si se repite
            int_Madril_Aux = Hoja.cells(lngRow, 3)
        
        End If
        
        'ASSIGN valor de fila fin para MERGE
        int_Merge_End = lngRow
        
        'ADD 1 a fila
        lngRow = lngRow + 1
                
    Loop
    '---------------------------------------------
    
    
    'DELETE el ultimo Copy, para que en el excel que se genera, no tenga nada seleccionado
    Obj.CutCopyMode = False
    
    
    'SELECT fila 12, para que cuando guarde el excel, quede el puntero en donde deben comenzar a ingresar datos
    Hoja.cells(12, 3).Select
    
    
    
      
    
    'ENABLE proteccion
    Hoja.Protect Password:="Sinopec2015"
    
    
    
    'SAVE template, con Fecha Actual y Nombre de Pozo
    Obj.ActiveWorkbook.SaveAs FileName:=Me.cmd_Destino + "Flowlog_" + dateToIso(Now()) + "_" + UCase(txt_Pozo) + ".xlsx"

  
    'CLOSE objetos
    Libro.Close
'    Set Obj = Nothing
    Obj.Quit
    Set Libro = Nothing
    Set Hoja = Nothing

  'mouse defa
  Screen.MousePointer = vbDefault
  

End Sub

Private Sub cmd_Importar_Click()

    Dim lngF As Long
    Dim varTabla, varQuery As Variant
    Dim strT, strSP As String
    
    
    'WHILE grilla
    For lngF = 1 To Me.spdSQL.DataRowCnt
        
        'GET tabla y Query
        Me.spdSQL.GetText 1, lngF, varTabla
        Me.spdSQL.GetText 2, lngF, varQuery
        
        'BUILD query
        strT = "exec " & strSP & varQuery
        
        'busco en tabla
        Set rs = SQLexec(strT)

        'chequeo errores
        If Not SQLparam.CnErrNumero = -1 Then
          SQLError
          SQLclose
        End If
        
    Next
    
    blnB = MsgBox("La importación, se realizo con éxito.", vbExclamation, "atención...")


End Sub

Private Sub cmd_movement_valves_Click()

  Dim blnB, blnRowPainted, blnFilaConMerge As Boolean
  Dim strT As String
  Dim lngRow, lngCol, lngColVisible As Long
  Dim varDato, varDatoColRig, varDatoColRigOficial, strManiobra As Variant
  
  Dim Obj As Object
  Dim Libro As Object
  Dim Hoja As Object
  
  Me.comOrigen.Filter = "Archivos de Excel|Valves*.xlsx"
    
  'titulo de ventana
  Me.comOrigen.DialogTitle = "Proceso de importaciòn..."
  
  Me.comOrigen.FileName = ""
  
  'abro cuadro de dialogo
  Me.comOrigen.ShowSave
      
  'si cancelar salgo
  If Me.comOrigen.FileName = "" Then
    Exit Sub
  End If
  
  'mouse reloj
  Screen.MousePointer = vbHourglass
  
  'CREATE excel
  Set Obj = CreateObject("Excel.application")
  Set Libro = Obj.Workbooks.Open(Me.comOrigen.FileName)
  Set Hoja = Libro.Sheets(1)
  
  
  'DISABLE proteccion
  Hoja.UnProtect Password:="Sinopec2015"
  
  
  'ENABLE excel visible
  Obj.Visible = False
  

    'DELETE filas, por las dudas que haya basura, con alguna importaciòn anterior
    Me.spdSQL.DeleteRows 1, 100


    '--------------------------------------------------------------------------------------
    'VALVES MOVEMENT - DELETE
    '--------------------------------------------------------------------------------------
    
    'SET algunas variables utiles
    lngRow = 13
    strT = ""
    
    var_WellID = Hoja.cells(lngRow, 3)
    strT = strT + "'" + var_WellID + "', "
    
    var_Date = Hoja.cells(lngRow, 5)
    strT = strT + "'" + dateToIso(var_Date) + "'"
    
    'ADD Query a Grilla
    Me.spdSQL.SetText 1, Me.spdSQL.DataRowCnt + 1, "Delete"
    Me.spdSQL.SetText 2, Me.spdSQL.DataRowCnt, "Injection_Valves_Movement_ELI_sp " + strT
    
    
    
    '--------------------------------------------------------------------------------------
    'FLOW LOG GET DATE MASTER
    '--------------------------------------------------------------------------------------
    
    Dim var_Date_Master As Variant
    
    'GET Date Master
    strT = "select max([Date Master]) as Date_Master FROM Injection_Targets_Main_vw WHERE [Well Name] = '" + var_WellID + "'"
    Set rs = SQLexec(strT)
    
    'CHECK errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
    End If
    
    'CHECK si encontro algo
    If Not rs.EOF Then
                
        If IsNull(rs!Date_Master) Then
            
            blnB = MsgBox("Date Master, no se encuentra en Targets - Main, para el Pozo: " + var_WellID + ". No puede continuar con el proceso de importaciòn.", vbExclamation, "atención...")
            
            'mouse defa
            Screen.MousePointer = vbDefault
            
            Exit Sub
        
        End If
        
    End If
    
    'SAVE Date Master, para usar en Main y Auxiliary
    var_Date_Master = rs!Date_Master
    

    
    '--------------------------------------------------------------------------------------
    'VALVES MOVEMENT GENERAL
    '--------------------------------------------------------------------------------------
    
    'SET algunas variables utiles
    blnB_1 = True
    lngRow = 13
    strT = ""
    
    var_WellID = Hoja.cells(lngRow, 3)
    strT = strT + "'" + var_WellID + "', "
    
    var_Date = Hoja.cells(lngRow, 5)
    strT = strT + "'" + dateToIso(var_Date) + "', "
    
    'CHANGE Color Blanco - antes de validad
    Hoja.cells(lngRow, 5).Interior.Color = RGB(255, 255, 255)
    
    'CHECK, si fecha de ejecuciòn vacia
    If var_Date = "" Then
        Hoja.cells(lngRow, 5).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If
    
    strT = strT + "'" + dateToIso(var_Date_Master) + "', "
    
    var_Date_Program = Hoja.cells(lngRow, 4)
    strT = strT + "'" + dateToIso(var_Date_Program) + "', "
    
    'CHANGE Color Blanco - antes de validad
    Hoja.cells(lngRow, 4).Interior.Color = RGB(255, 255, 255)
    
    'CHECK, si fecha program vacia
    If var_Date_Program = "" Then
        Hoja.cells(lngRow, 4).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If
    
    
    'CHECK, si Fecha Ejecuciòn <= Date Master
    If dateToIso(var_Date) <= dateToIso(var_Date_Master) Then
        Hoja.cells(lngRow, 5).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If
    
    
    varText = Hoja.cells(lngRow, 7)
    strT = strT + str(varText) + ", "

    'CHANGE Color Blanco - antes de validad
    Hoja.cells(lngRow, 7).Interior.Color = RGB(255, 255, 255)

    If varText = "" Then
        Hoja.cells(lngRow, 7).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If

    'CHANGE Color Blanco - antes de validad
    Hoja.cells(lngRow, 8).Interior.Color = RGB(255, 255, 255)

    varText = Hoja.cells(lngRow, 8)
    strT = strT + str(varText) + ", "

    If varText = "" Then
        Hoja.cells(lngRow, 8).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If
    
    varText = Hoja.cells(lngRow, 10)
    strT = strT + str(varText) + ", "
    
    'CHANGE Color Blanco - antes de validad
    Hoja.cells(lngRow, 10).Interior.Color = RGB(255, 255, 255)
    
    If varText = "" Then
        Hoja.cells(lngRow, 10).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If
    
    varText = Hoja.cells(lngRow, 11)
    strT = strT + str(varText) + ", "
    
    'CHANGE Color Blanco - antes de validad
    Hoja.cells(lngRow, 11).Interior.Color = RGB(255, 255, 255)
    
    If varText = "" Then
        Hoja.cells(lngRow, 11).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If
    
    
    varText = Hoja.cells(lngRow, 9)
    strT = strT + str(varText) + ", "
    
    'CHANGE Color Blanco - antes de validad
    Hoja.cells(lngRow, 9).Interior.Color = RGB(255, 255, 255)
    
    If varText = "" Then
        Hoja.cells(lngRow, 9).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If
    
    
    'Comments, no validar
    varText = Hoja.cells(lngRow + 4, 3)
    If varText = "" Then
        strT = strT + "'" + varText + "'"
    ElseIf IsNumeric(varText) Then
            strT = strT + "'" + str(varText) + "'"
        Else
            strT = strT + "'" + varText + "'"
        End If
    
    'ADD Query a Grilla
    Me.spdSQL.SetText 1, Me.spdSQL.DataRowCnt + 1, "General"
    Me.spdSQL.SetText 2, Me.spdSQL.DataRowCnt, "Injection_Valves_Movement_General_INS_sp " + strT
  
      
  
    '--------------------------------------------------------------------------------------
    'VALVES MOVEMENT MAIN
    '--------------------------------------------------------------------------------------
    strCabecera = ""
    
    'SET puntero en fila donde tomo datos que identifican al registro
    lngRow = 13
        
    'GET datos adicionales, Well Name, Date Program, Date
    var_WellID = Hoja.cells(lngRow, 3)
    strCabecera = strCabecera + "'" + var_WellID + "', "
    
    strCabecera = strCabecera + "'" + dateToIso(var_Date) + "', "
    
    strCabecera = strCabecera + "'" + dateToIso(var_Date_Master) + "', "
    
    'FIND celda en donde comienza ESTANQUEIDAD, si no encuentra, busca hasta fila 500 y corta
    lngRowAd = 1
    varText = Hoja.cells(lngRowAd, 3)
    
    Do Until varText = "Mandril" Or lngRowAd = 500
        lngRowAd = lngRowAd + 1
        varText = Hoja.cells(lngRowAd, 3)
    Loop
    
    'CHECK, si no encontro las celdas correspondientes a ADMISIONES, cancelo proceso
    If lngRowAd = 500 Then
        Exit Sub
    End If
    
    'SET fila comienzo de detalle
    lngRow = lngRowAd + 2
    
    'WHILE para armar store en FLow Log - MAIN
    Do
        
        'GET valor de fila, columna 2
        varText = Hoja.cells(lngRow, 3)
        
        'CHECK si finalizaron las filas para ADMISIONES, EXIT
        If varText = "" Then
            Exit Do
        End If
        
        'SET algunas variables utiles
        blnB_3 = True
        strT = ""
        
        'ADD cabecera
        strT = strCabecera
        
        'CHANGE Color Blanco - antes de validad
        Hoja.cells(lngRow, 3).Interior.Color = RGB(255, 255, 255)
        
        'ADD detalle
        varText = Hoja.cells(lngRow, 3)
        strT = strT + str(varText) + ", "
        
        If varText = "" Then
            Hoja.cells(lngRow, 3).Interior.Color = RGB(250, 152, 84)
            blnB_1 = False
        End If
        
        varText = Hoja.cells(lngRow, 4)
        strT = strT + str(varText) + ", "
        
        'CHANGE Color Blanco - antes de validad
        Hoja.cells(lngRow, 4).Interior.Color = RGB(255, 255, 255)
        
        If varText = "" Then
            Hoja.cells(lngRow, 4).Interior.Color = RGB(250, 152, 84)
            blnB_1 = False
        End If
        
        'Maniobra
        varText = Hoja.cells(lngRow, 5)
        strT = strT + "'" + varText + "', "
        
        'SAVE Maniobra, para validar otras columnas que dependen de esta
        strManiobra = varText
    
        'CHANGE Color Blanco - antes de validad
        Hoja.cells(lngRow, 5).Interior.Color = RGB(255, 255, 255)
        
        If varText = "" Then
            Hoja.cells(lngRow, 5).Interior.Color = RGB(250, 152, 84)
            blnB_1 = False
        End If
        
        varText = Hoja.cells(lngRow, 6)
        strT = strT + "'" + varText + "', "
        
        'CHANGE Color Blanco - antes de validad
        Hoja.cells(lngRow, 6).Interior.Color = RGB(255, 255, 255)
        
        If varText = "" Then
'            Hoja.Cells(lngRow, 6).Interior.Color = RGB(250, 152, 84)
'            blnB_1 = False
        End If
        
    
        varText = Hoja.cells(lngRow, 7)
        strT = strT + str(varText) + ", "
        
        'CHANGE Color Blanco - antes de validad
        Hoja.cells(lngRow, 7).Interior.Color = RGB(255, 255, 255)
        
        If varText = "" Then
'            Hoja.Cells(lngRow, 7).Interior.Color = RGB(250, 152, 84)
'            blnB_1 = False
        End If
        
        varText = Hoja.cells(lngRow, 8)
        If varText = "" Then
            strT = strT + "'" + varText + "',"
        ElseIf IsNumeric(varText) Then
                strT = strT + str(varText) + ","
            Else
                strT = strT + "'" + varText + "',"
            End If
        
        'CHANGE Color Blanco - antes de validad
        Hoja.cells(lngRow, 8).Interior.Color = RGB(255, 255, 255)
        
        If varText = "" Then
'            Hoja.Cells(lngRow, 8).Interior.Color = RGB(250, 152, 84)
'            blnB_1 = False
        End If
        
        varText = Hoja.cells(lngRow, 9)
        strT = strT + IIf(IsNumeric(varText), str(varText), "null") + ", "
        
        'CHANGE Color Blanco - antes de validad
        Hoja.cells(lngRow, 9).Interior.Color = RGB(255, 255, 255)
        
        If (strManiobra = "Reemplazo" And varText = "") Or (strManiobra <> "Reemplazo" And varText <> "") Then
            Hoja.cells(lngRow, 9).Interior.Color = RGB(250, 152, 84)
            blnB_1 = False
        End If
            
        'CHANGE Color Blanco - antes de validad
        Hoja.cells(lngRow, 10).Interior.Color = RGB(255, 255, 255)
            
        varText = Hoja.cells(lngRow, 10)
        strT = strT + IIf(IsNumeric(varText), str(varText), "null") + ", "
        
        If (strManiobra = "Reemplazo" And varText = "") Or (strManiobra <> "Reemplazo" And varText <> "") Then
            Hoja.cells(lngRow, 10).Interior.Color = RGB(250, 152, 84)
            blnB_1 = False
        End If
        
        'CHANGE Color Blanco - antes de validad
        Hoja.cells(lngRow, 11).Interior.Color = RGB(255, 255, 255)
        
        varText = Hoja.cells(lngRow, 11)
        strT = strT + IIf(IsNumeric(varText), str(varText), "null") + ", "
        
        If (strManiobra = "Reemplazo" And varText = "") Or (strManiobra <> "Reemplazo" And varText <> "") Then
            Hoja.cells(lngRow, 11).Interior.Color = RGB(250, 152, 84)
            blnB_1 = False
        End If
        
        'CHANGE Color Blanco - antes de validad
        Hoja.cells(lngRow, 12).Interior.Color = RGB(255, 255, 255)
        
        varText = Hoja.cells(lngRow, 12)
        strT = strT + IIf(IsNumeric(varText), str(varText), "null") + ", "
        
        If (strManiobra = "Reemplazo" And varText = "") Or (strManiobra <> "Reemplazo" And varText <> "") Then
            Hoja.cells(lngRow, 12).Interior.Color = RGB(250, 152, 84)
            blnB_1 = False
        End If
        
        varText = Hoja.cells(lngRow, 13)
        strT = strT + IIf(IsNumeric(varText), str(varText), "null") + ", "
        
        'CHANGE Color Blanco - antes de validad
        Hoja.cells(lngRow, 13).Interior.Color = RGB(255, 255, 255)
        
        If (strManiobra = "Reemplazo" And varText = "") Or (strManiobra <> "Reemplazo" And varText <> "") Then
            Hoja.cells(lngRow, 13).Interior.Color = RGB(250, 152, 84)
            blnB_1 = False
        End If
        
        varText = Hoja.cells(lngRow, 14)
        strT = strT + IIf(IsNumeric(varText), str(varText), "null") + ", "
        
        'CHANGE Color Blanco - antes de validad
        Hoja.cells(lngRow, 14).Interior.Color = RGB(255, 255, 255)
        
        If (strManiobra = "Reemplazo" And varText = "") Or (strManiobra <> "Reemplazo" And varText <> "") Then
            Hoja.cells(lngRow, 14).Interior.Color = RGB(250, 152, 84)
            blnB_1 = False
        End If
        
        varText = Hoja.cells(lngRow, 15)
        strT = strT + "'" + varText + "', "
        
        'CHANGE Color Blanco - antes de validad
        Hoja.cells(lngRow, 15).Interior.Color = RGB(255, 255, 255)
        
        If varText = "" Then
            Hoja.cells(lngRow, 15).Interior.Color = RGB(250, 152, 84)
            blnB_1 = False
        End If
        
        varText = Hoja.cells(lngRow, 16)
        strT = strT + str(varText) + ", "
        
        'CHANGE Color Blanco - antes de validad
        Hoja.cells(lngRow, 16).Interior.Color = RGB(255, 255, 255)
        
        If varText = "" Then
            Hoja.cells(lngRow, 16).Interior.Color = RGB(250, 152, 84)
            blnB_1 = False
        End If
        
        varText = Hoja.cells(lngRow, 17)
        If IsNumeric(varText) Then
            strT = strT + "'" + str(varText) + "', "
        Else
            strT = strT + "'" + varText + "', "
        End If
        
        'CHANGE Color Blanco - antes de validad
        Hoja.cells(lngRow, 17).Interior.Color = RGB(255, 255, 255)
        
        If varText = "" Then
            Hoja.cells(lngRow, 17).Interior.Color = RGB(250, 152, 84)
            blnB_1 = False
        End If
        
        'CHECK las ùltimas 4 columnas de tipo LISTA (X = true), no se validan
        varText = Hoja.cells(lngRow, 18)
        strT = strT + str(IIf(LCase(varText) = "x", 1, 0)) + ", "
        
        varText = Hoja.cells(lngRow, 19)
        strT = strT + str(IIf(LCase(varText) = "x", 1, 0)) + ", "
        
        varText = Hoja.cells(lngRow, 20)
        strT = strT + str(IIf(LCase(varText) = "x", 1, 0)) + ", "
        
        varText = Hoja.cells(lngRow, 21)
        strT = strT + str(IIf(LCase(varText) = "x", 1, 0))
        
        
        'ADD Query a Grilla
        Me.spdSQL.SetText 1, Me.spdSQL.DataRowCnt + 1, "Main"
        Me.spdSQL.SetText 2, Me.spdSQL.DataRowCnt, "Injection_Valves_Movement_Main_INS_sp " + strT
        strT = ""
        
        'ADD 1 - contador de filas
        lngRow = lngRow + 1
    
    Loop
    
  
    'mouse defa
    Screen.MousePointer = vbDefault


    'DEFAULT, habilita el botòn de exportar
    Me.cmd_Importar.Enabled = True

    'CHECK si detecta validaciones que no se cumplieron, emite mensaje, Dehabilita la capacidad de Importar
    If blnB_1 Then
        blnB = MsgBox("El excel que desea importar, cumple con las validaciones, haga clic en el boton Importar", vbExclamation, "atención...")
        Me.cmd_Importar.Enabled = True
    Else
        blnB = MsgBox("El excel que desea importar, no cumple con las validaciones requeridas, debe abrirlo, modificar las celdas en naranja y volver a importarlo.", vbExclamation, "atención...")
        Me.cmd_Importar.Enabled = False
    End If
    
    'ENABLE proteccion
    Hoja.Protect Password:="Sinopec2015"
    
    'Esta linea, es para que no pregunte si desea guardar los cambios que se le hicieron para la validacion
    Obj.displayAlerts = False
    
    '   CHECK si la validaciòn fallo, guardamos excel con el mismo nombre, celdas en naranja, las que se deben modificar
    '   DISABLE para que muestre mensaje de confirmacion de guardar, si archivo existe, lo fuerza directamente
    Obj.ActiveWorkbook.SaveAs FileName:=Me.comOrigen.FileName
    
    'CLOSE objetos
    Libro.Close
'    Set Obj = Nothing
    Obj.Quit
    Set Libro = Nothing
    Set Hoja = Nothing


    'CHECK si no paso validacion, EXIT
    If Not blnB_1 Then
        Exit Sub
    End If
  
  
    'CHECK si existe Pozo y Fecha importado con anterioridad
    strT = "select count(*) as Existe FROM Injection_Valves_Movement_General_vw WHERE [Well Name] = '" + var_WellID + "' AND Date = '" + dateToIso(var_Date) + "'"
    Set rs = SQLexec(strT)
    
    'chequeo errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
    End If
    
    'CHECK si trajo algùn registro
    If Not rs.EOF Then
    
        If rs!Existe = 1 Then
            blnB = MsgBox("Se ha encontrado informacion, para el Pozo: " + var_WellID + ", con Fecha: " + str(var_Date) + ", si continua con la importación, la informaciòn sera reemplazada.", vbExclamation, "atención...")
        End If
        
    End If
  

End Sub

Private Sub cmd_open_flow_log_Click()


  Dim blnB, blnRowPainted, blnFilaConMerge As Boolean
  Dim strT As String
  Dim lngRow, lngCol, lngColVisible As Long
  Dim varDato, varDatoColRig, varDatoColRigOficial As Variant
  
  Dim Obj As Object
  Dim Libro As Object
  Dim Hoja As Object
  
  Me.comOrigen.Filter = "Archivos de Excel|Flow*.xlsx"
    
  'titulo de ventana
  Me.comOrigen.DialogTitle = "Proceso de importaciòn..."
  
  Me.comOrigen.FileName = ""
  
  'abro cuadro de dialogo
  Me.comOrigen.ShowSave
      
  'si cancelar salgo
  If Me.comOrigen.FileName = "" Then
    Exit Sub
  End If
  
  'mouse reloj
  Screen.MousePointer = vbHourglass
  
  'CREATE excel
  Set Obj = CreateObject("Excel.application")
  Set Libro = Obj.Workbooks.Open(Me.comOrigen.FileName)
  Set Hoja = Libro.Sheets(1)
  
  
  'DISABLE proteccion
  Hoja.UnProtect Password:="Sinopec2015"
  
  
  'ENABLE excel visible
  Obj.Visible = False
  
  
    'DELETE filas, por las dudas que haya basura, con alguna importaciòn anterior
    Me.spdSQL.DeleteRows 1, 100

    '--------------------------------------------------------------------------------------
    'FLOW LOG DELETE
    '--------------------------------------------------------------------------------------
    
    'SET algunas variables utiles
    blnB_1 = True
    lngRow = 12
    strT = ""
    
    var_WellID = Hoja.cells(lngRow, 3)
    strT = strT + "'" + var_WellID + "', "
    
    var_Date = Hoja.cells(lngRow, 5)
    strT = strT + "'" + dateToIso(var_Date) + "'"
    
    'ADD Query a Grilla
    Me.spdSQL.SetText 1, Me.spdSQL.DataRowCnt + 1, "Delete"
    Me.spdSQL.SetText 2, Me.spdSQL.DataRowCnt, "Injection_Flow_Log_ELI_sp " + strT
    
        
    '--------------------------------------------------------------------------------------
    'FLOW LOG DATE MASTER
    '--------------------------------------------------------------------------------------
    
    Dim var_Date_Master As Variant
    
    'GET Date Master
    strT = "select max([Date Master]) as Date_Master FROM Injection_Targets_Main_vw WHERE [Well Name] = '" + var_WellID + "'"
    Set rs = SQLexec(strT)
    
    'CHECK errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
    End If
    
    'CHECK si encontro algo
    If Not rs.EOF Then
                
        If IsNull(rs!Date_Master) Then
            
            blnB = MsgBox("Date Master, no se encuentra en Targets - Main, para el Pozo: " + var_WellID + ". No puede continuar con el proceso de importaciòn.", vbExclamation, "atención...")
            
            'mouse defa
            Screen.MousePointer = vbDefault
            
            Exit Sub
        
        End If
        
    End If
    
    'SAVE Date Master, para usar en Main y Auxiliary
    var_Date_Master = rs!Date_Master
    
    
    
    '--------------------------------------------------------------------------------------
    'FLOW LOG GENERAL
    '--------------------------------------------------------------------------------------
    
    'SET algunas variables utiles
    blnB_1 = True
    lngRow = 12
    strT = ""
    
    var_WellID = Hoja.cells(lngRow, 3)
    strT = strT + "'" + var_WellID + "', "
    
    var_Date = Hoja.cells(lngRow, 5)
    strT = strT + "'" + dateToIso(var_Date) + "', "
    
    'CHANGE Color Blanco - antes de validad
    Hoja.cells(lngRow, 5).Interior.Color = RGB(255, 255, 255)
    
    If var_Date = "" Then
        Hoja.cells(lngRow, 5).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If
    
    strT = strT + "'" + dateToIso(var_Date_Master) + "', "
    
    var_Date_Program = Hoja.cells(lngRow, 4)
    strT = strT + "'" + dateToIso(var_Date_Program) + "', "
    
    'CHANGE Color Blanco - antes de validad
    Hoja.cells(lngRow, 4).Interior.Color = RGB(255, 255, 255)
    
    'CHECK, si fecha vacia
    If var_Date_Program = "" Then
        Hoja.cells(lngRow, 4).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If
    
    
    'CHECK, si Fecha Ejecuciòn <= Date Master
    If dateToIso(var_Date) <= dateToIso(var_Date_Master) Then
        Hoja.cells(lngRow, 5).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If

    
    varText = Hoja.cells(lngRow, 7)
    strT = strT + str(varText) + ", "

    'CHANGE Color Blanco - antes de validad
    Hoja.cells(lngRow, 7).Interior.Color = RGB(255, 255, 255)

    If varText = "" Then
        Hoja.cells(lngRow, 7).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If
    
    varText = Hoja.cells(lngRow, 8)
    strT = strT + str(varText) + ", "

    'CHANGE Color Blanco - antes de validad
    Hoja.cells(lngRow, 8).Interior.Color = RGB(255, 255, 255)

    If varText = "" Then
        Hoja.cells(lngRow, 8).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If
    
    varText = Hoja.cells(lngRow, 9)
    strT = strT + str(varText) + ", "
    
    'CHANGE Color Blanco - antes de validad
    Hoja.cells(lngRow, 9).Interior.Color = RGB(255, 255, 255)
    
    If varText = "" Then
        Hoja.cells(lngRow, 9).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If
    
    varText = Hoja.cells(lngRow, 10)
    strT = strT + str(varText) + ", "
    
    'CHANGE Color Blanco - antes de validad
    Hoja.cells(lngRow, 10).Interior.Color = RGB(255, 255, 255)
    
    If varText = "" Then
        Hoja.cells(lngRow, 10).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If
    
    varText = Hoja.cells(lngRow, 11)
    strT = strT + str(varText) + ", "
    
    'CHANGE Color Blanco - antes de validad
    Hoja.cells(lngRow, 11).Interior.Color = RGB(255, 255, 255)
    
    If varText = "" Then
        Hoja.cells(lngRow, 11).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If
    
    varText = Hoja.cells(lngRow, 12)
    strT = strT + str(varText) + ", "
    
    'CHANGE Color Blanco - antes de validad
    Hoja.cells(lngRow, 12).Interior.Color = RGB(255, 255, 255)
    
    If varText = "" Then
        Hoja.cells(lngRow, 12).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If
    
    varText = Hoja.cells(lngRow, 13)
    strT = strT + str(varText) + ", "
    
    'CHANGE Color Blanco - antes de validad
    Hoja.cells(lngRow, 13).Interior.Color = RGB(255, 255, 255)
    
    If varText = "" Then
        Hoja.cells(lngRow, 13).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If
    
    
    'FIND celda en donde comienza ESTANQUEIDAD, si no encuentra, busca hasta fila 500 y corta
    lngRowEs = 1
    varText = Hoja.cells(lngRow, 3)
    
    Do Until varText = "ESTANQUEIDAD" Or lngRowEs = 500
        lngRowEs = lngRowEs + 1
        varText = Hoja.cells(lngRowEs, 3)
    Loop
    
    'CHECK, si no encontro las celdas correspondientes a ESTANQUEIDAD, cancelo proceso
    If lngRowEs = 500 Then
        Exit Sub
    End If
    
    varText = Hoja.cells(lngRowEs + 5, 7)
    If varText = "" Then
        strT = strT + "null" + ", "
    Else
        strT = strT + str(varText) + ", "
    End If
    
    varText = Hoja.cells(lngRowEs + 5, 8)
    strT = strT + "'" + varText + "', "
    
    'CHANGE Color Blanco - antes de validad
    Hoja.cells(lngRowEs + 5, 8).Interior.Color = RGB(255, 255, 255)

    If varText = "" Then
        Hoja.cells(lngRowEs + 5, 8).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If
    
    'Comments - Previene que si se ingresa un nùmero como comentario, no de error.
    varText = Hoja.cells(lngRow + 4, 3)
    If varText = "" Then
        strT = strT + "'" + varText + "'"
    ElseIf IsNumeric(varText) Then
            strT = strT + "'" + str(varText) + "'"
        Else
            strT = strT + "'" + varText + "'"
        End If
    
    
    'ADD Query a Grilla
    Me.spdSQL.SetText 1, Me.spdSQL.DataRowCnt + 1, "General"
    Me.spdSQL.SetText 2, Me.spdSQL.DataRowCnt, "Injection_Flow_Log_General_INS_sp " + strT
  
  
  
  
    '--------------------------------------------------------------------------------------
    'FLOW LOG MAIN
    '--------------------------------------------------------------------------------------
    strCabecera = ""
    
    'SET puntero en fila donde tomo datos que identifican al registro
    lngRow = 12
        
    'GET datos adicionales, Well Name, Date Program, Date
    var_WellID = Hoja.cells(lngRow, 3)
    strCabecera = strCabecera + "'" + var_WellID + "', "
    
    strCabecera = strCabecera + "'" + dateToIso(var_Date) + "', "
    
    strCabecera = strCabecera + "'" + dateToIso(var_Date_Master) + "', "
    
    
    'FIND celda en donde comienza ESTANQUEIDAD, si no encuentra, busca hasta fila 500 y corta
    lngRowAd = 1
    varText = Hoja.cells(lngRowAd, 3)
    
    Do Until varText = "ADMISIONES" Or lngRowAd = 500
        lngRowAd = lngRowAd + 1
        varText = Hoja.cells(lngRowAd, 3)
    Loop
    
    'CHECK, si no encontro las celdas correspondientes a ADMISIONES, cancelo proceso
    If lngRowAd = 500 Then
        Exit Sub
    End If
    
    'SET fila comienzo de detalle
    lngRow = lngRowAd + 4
    
    'WHILE para armar store en FLow Log - MAIN
    Do
        
        'GET valor de fila, columna 2
        varText = Hoja.cells(lngRow, 3)
        
        'CHECK si finalizaron las filas para ADMISIONES, EXIT
        If varText = "" Then
            Exit Do
        End If
        
        'SET algunas variables utiles
        strT = ""
        
        'ADD cabecera
        strT = strCabecera
        
        'ADD detalle
        'Mandril Number
        varText = Hoja.cells(lngRow, 3)
        strT = strT + str(varText) + ", "
        
        'Mandril Depth
        varText = Hoja.cells(lngRow, 4)
        strT = strT + str(varText) + ", "
        
        'Reservoir
        varText = Hoja.cells(lngRow, 6)
        strT = strT + "'" + varText + "',"
    
        'Tope
        varText = Hoja.cells(lngRow, 7)
        strT = strT + str(varText) + ","
        
        'Base
        varText = Hoja.cells(lngRow, 8)
        strT = strT + str(varText) + ","
        
        'Caudal Admision Capa (es el calculado)
        varText = Hoja.cells(lngRow, 13)
        
        If varText = "" Then
            strT = strT + "null" + ", "
        Else
            strT = strT + str(varText) + ", "
        End If
        
        'CHANGE Color Blanco - antes de validad
        Hoja.cells(lngRow, 13).Interior.Color = RGB(255, 255, 255)
        
        If varText = "" Then
            Hoja.cells(lngRow, 13).Interior.Color = RGB(250, 152, 84)
            blnB_1 = False
        End If
        
        'Comment
        varText = Hoja.cells(lngRow, 14)
        strT = strT + "'" + varText + "'"
        
        'ADD Query a Grilla
        Me.spdSQL.SetText 1, Me.spdSQL.DataRowCnt + 1, "Main"
        Me.spdSQL.SetText 2, Me.spdSQL.DataRowCnt, "Injection_Flow_Log_Main_INS_sp " + strT
        strT = ""
        
        'ADD 1 - contador de filas
        lngRow = lngRow + 1
    
    Loop
  
    
    'CHECK, si TOTAL MANDRILES = TOTAL CAPAS 1 = TOTAL CAPAS 2
    varText1 = Hoja.cells(lngRow + 3, 5)
    varText2 = Hoja.cells(lngRow + 3, 10)
    varText3 = Hoja.cells(lngRow + 3, 13)
  
    'CHANGE Color Blanco - antes de validad
    Hoja.cells(lngRow + 3, 5).Interior.Color = RGB(255, 255, 255)
    Hoja.cells(lngRow + 3, 10).Interior.Color = RGB(255, 255, 255)
    Hoja.cells(lngRow + 3, 13).Interior.Color = RGB(255, 255, 255)
  
    'Si alguna de las 3 columnas no es igual, las marcamos
    If Val(varText1) <> Val(varText2) Or Val(varText1) <> Val(varText3) Then
        Hoja.cells(lngRow + 3, 5).Interior.Color = RGB(250, 152, 84)
        Hoja.cells(lngRow + 3, 10).Interior.Color = RGB(250, 152, 84)
        Hoja.cells(lngRow + 3, 13).Interior.Color = RGB(250, 152, 84)
        blnB_1 = False
    End If
  
  
    '--------------------------------------------------------------------------------------
    'FLOW LOG AUXILIARY
    '--------------------------------------------------------------------------------------
    strCabecera = ""
    
    'SET puntero en fila donde tomo datos que identifican al registro
    lngRow = 12
        
   'GET datos adicionales, Well Name, Date Program, Date
    var_WellID = Hoja.cells(lngRow, 3)
    strCabecera = strCabecera + "'" + var_WellID + "', "
    
    strCabecera = strCabecera + "'" + dateToIso(var_Date) + "', "
    
    strCabecera = strCabecera + "'" + dateToIso(var_Date_Master) + "', "
    
    
    'FIND celda en donde comienza ESTANQUEIDAD, si no encuentra, busca hasta fila 500 y corta
    lngRowEs = 1
    varText = Hoja.cells(lngRowEs, 3)
    
    Do Until varText = "ESTANQUEIDAD" Or lngRowEs = 500
        lngRowEs = lngRowEs + 1
        varText = Hoja.cells(lngRowEs, 3)
    Loop
    
    'CHECK, si no encontro las celdas correspondientes a ADMISIONES, cancelo proceso
    If lngRowEs = 500 Then
        Exit Sub
    End If
    
    'SET fila comienzo de detalle
    lngRow = lngRowEs + 6
    
    'WHILE para armar store en FLow Log - AUXILIARY
    Do
        
        'GET valor de fila, columna 2
        varText = Hoja.cells(lngRow, 3)
        
        'CHECK si no encuentra nada, EXIT
        If varText = "" Then
            Exit Do
        End If
        
        'SET algunas variables utiles
        strT = ""
        
        'ADD cabecera
        strT = strCabecera
        
        'ADD detalle
        varText = Hoja.cells(lngRow, 3)
        strT = strT + str(varText) + ", "
        
        varText = Hoja.cells(lngRow, 4)
        strT = strT + str(varText) + ", "
        
        varText = Hoja.cells(lngRow, 5)
        strT = strT + "'" + varText + "'"
        
        'CHANGE Color Blanco - antes de validad
        Hoja.cells(lngRow, 5).Interior.Color = RGB(255, 255, 255)

        If varText = "" Then
            Hoja.cells(lngRow, 5).Interior.Color = RGB(250, 152, 84)
            blnB_1 = False
        End If

        
        'ADD Query a Grilla
        Me.spdSQL.SetText 1, Me.spdSQL.DataRowCnt + 1, "Auxiliary"
        Me.spdSQL.SetText 2, Me.spdSQL.DataRowCnt, "Injection_Flow_Log_Auxiliary_INS_sp " + strT
        strT = ""
        
        'ADD 1 - contador de filas
        lngRow = lngRow + 1
        
    Loop
    
    
    'mouse defa
    Screen.MousePointer = vbDefault


    'CHECK si detecta validaciones que no se cumplieron, emite mensaje, Dehabilita la capacidad de Importar
    If blnB_1 Then
        blnB = MsgBox("El excel que desea importar, cumple con las validaciones, haga clic en el boton Importar", vbExclamation, "atención...")
        Me.cmd_Importar.Enabled = True
    Else
        blnB = MsgBox("El excel que desea importar, no cumple con las validaciones requeridas, debe abrirlo, modificar las celdas en naranja y volver a importarlo.", vbExclamation, "atención...")
        Me.cmd_Importar.Enabled = False
    End If
  
    'ENABLE proteccion
    Hoja.Protect Password:="Sinopec2015"
    
    'Esta linea, es para que no pregunte si desea guardar los cambios que se le hicieron para la validacion
    Obj.displayAlerts = False
    
    '   CHECK si la validaciòn fallo, guardamos excel con el mismo nombre, celdas en naranja, las que se deben modificar
    '   DISABLE para que muestre mensaje de confirmacion de guardar, si archivo existe, lo fuerza directamente
    Obj.ActiveWorkbook.SaveAs FileName:=Me.comOrigen.FileName
  
    'CLOSE objetos
    Libro.Close
'    Set Obj = Nothing
    Obj.Quit
    Set Libro = Nothing
    Set Hoja = Nothing


    'CHECK si no paso validacion, EXIT
    If Not blnB_1 Then
        Exit Sub
    End If
        
    
    'CHECK si existe Pozo y Fecha importado con anterioridad
    strT = "select count(*) as Existe FROM Injection_Flow_Log_General_vw WHERE [Well Name] = '" + var_WellID + "' AND Date = '" + dateToIso(var_Date) + "'"
    Set rs = SQLexec(strT)
    
    'chequeo errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
    End If
    
    
    'CHECK si trajo algùn registro
    If Not rs.EOF Then
    
        If rs!Existe = 1 Then
            blnB = MsgBox("Se ha encontrado informacion, para el Pozo: " + var_WellID + ", con Fecha: " + str(var_Date) + ", si continua con la importación, la informaciòn sera reemplazada.", vbExclamation, "atención...")
        End If
        
    End If
        
 

End Sub

Private Sub cmd_Salir_Click()
  
  Unload Me
  
End Sub

  
  
Private Sub cmd_valves_movement_generate_Click()

  
  
  Dim blnB, blnRowPainted, blnFilaConMerge As Boolean
  Dim strT As String
  Dim lngRow, lngCol, lngColVisible As Long
  Dim varDato, varDatoColRig, varDatoColRigOficial As Variant
  
  Dim str_Date_Master, str_Date_Assembly As String
  Dim int_Id_Well As Integer
  
  
  Dim Obj As Object
  Dim Libro As Object
  Dim Hoja As Object
  
  
    '--------------------------------------------------------------------------------------
    'WELL NAME - CHEQUEO EXISTENCIA
    '--------------------------------------------------------------------------------------
    
    'GET Date Master
    strT = "select WellName FROM pozosPropios_vw WHERE WellName = '" + txt_Pozo + "'"
    Set rs = SQLexec(strT)
    
    'CHECK errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
    End If
    
    'CHECK si encontro algo
    If rs.EOF Then
                
        blnB = MsgBox("El Pozo ingresado no existe..", vbExclamation, "atención...")
        
        Exit Sub
        
    End If
  
  
  
  'mouse reloj
  Screen.MousePointer = vbHourglass
  
  
    'CREATE excel
    Set Obj = CreateObject("Excel.application")
    Set Libro = Obj.Workbooks.Open(App.Path + "\TEMPLATE\ValvesMovement_YYYYMMDD_POZO.xlsx")
    Set Hoja = Libro.Sheets(1)
    
    
    'ENABLE excel visible
    Obj.Visible = False
        
    'DISABLE proteccion
    Hoja.UnProtect Password:="Sinopec2015"
        
        
    '--------------------------------------------------------------------------------------
    'FLOW LOG GET DATE MASTER
    '--------------------------------------------------------------------------------------
    
    'GET Date Master
    strT = "select * FROM Injection_Template_GET_MAX_Date_Master_vw WHERE Well_Name = '" + Me.txt_Pozo + "'"
    Set rs = SQLexec(strT)
    
    'CHECK errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
    End If
    
    'CHECK si encontro algo
    If rs.EOF Then
                
        blnB = MsgBox("Date Master, no se encuentra en Targets - Main, para el Pozo: " + var_WellID + ". No puede continuar con la generacion de la Plantilla.", vbExclamation, "atención...")
        
        'mouse defa
        Screen.MousePointer = vbDefault
        
        Exit Sub
        
    End If
    
    'SAVE Date Master, para usarlo en los procesos siguientes
    str_Date_Master = dateToIso(rs!Date_Master)
        
    'SAVE ID Well, para usarlo en los procesos siguientes
    int_Id_Well = rs!Id_Well
    
        
    '--------------------------------------------------------------------------------------
    'FLOW LOG - GENERAL
    '--------------------------------------------------------------------------------------
    
    'SET puntero en fila donde se encuentra la info para MAIN
    lngRow = 13
            
            
    'GET SUM Qwi Target
    strT = "select * FROM Injection_Template_GET_SUM_Qwi_Target_vw WHERE ID_Well = " + str(int_Id_Well) + " AND " + "Date_Master = '" + str_Date_Master + "'"
    Set rs = SQLexec(strT)
    
    'CHECK errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
    End If
    
    'CHECK si encontro algo
    If rs.EOF Then
                
        blnB = MsgBox("Qwi Target, no se encuentra en Targets - Main, para el Pozo: " + var_WellID + ". No puede continuar con la generacion de la Plantilla.", vbExclamation, "atención...")
        
        'mouse defa
        Screen.MousePointer = vbDefault
        
        Exit Sub
        
    End If
    
    
    'PUT - Well Name
    Hoja.cells(lngRow, 3) = Me.txt_Pozo
    
    'PUT - Caudal Target
    Hoja.cells(lngRow, 6) = rs!Qwi_Target
    
    
    
    '--------------------------------------------------------------------------------------
    'FLOW LOG - MAIN
    '--------------------------------------------------------------------------------------
    
    'SET puntero en fila donde se encuentra la info para MAIN
    lngRow = 25
    
    
    'GET Date Master
    strT = "select * FROM Injection_Template_GET_MAX_Mandrels_VM_vw WHERE Id_Well = " + str(int_Id_Well) + " ORDER BY Mandrel_Depth"
    Set rs = SQLexec(strT)
    
    'CHECK errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
    End If
    
    'CHECK si encontro algo
    If rs.EOF Then
                
        blnB = MsgBox("Los Mandrels, no se encuentra en Targets - Main, para el Pozo: " + var_WellID + ". No puede continuar con la generacion de la Plantilla.", vbExclamation, "atención...")
        
        'mouse defa
        Screen.MousePointer = vbDefault
        
        Exit Sub
        
    End If
    
    
    'WHILE Mandrels
    Do While Not rs.EOF
        
        'BUILD string con el rango, en donde debo seleccionar
        strT = Replace(str(lngRow) + ":" + str(lngRow), " ", "")
        
        'SELECT - COPY fila completa
        Hoja.Rows(strT).Select
'        Obj.CutCopyMode = False
        Obj.Selection.Copy
        
        'BUILD string con el rango, en donde debo copiar
        strT = Replace(str(lngRow + 1) + ":" + str(lngRow + 1), " ", "")
        
        Hoja.Rows(strT).Select
         
        'PASTE fila completa
        Obj.Selection.Insert
        'Shift:=xlDown
        
        'ASSING informacion a las celdas
        Hoja.cells(lngRow, 3) = rs!Mandrel_Number
        Hoja.cells(lngRow, 4) = rs!Mandrel_Depth
        Hoja.cells(lngRow, 6) = rs!New_Supplier
        Hoja.cells(lngRow, 7) = rs!New_Diameter
        Hoja.cells(lngRow, 8) = rs!New_Qwi_Theoretical
        
        
        '-------------------------------------------------------------------------------------------------------
        '
        'BUID string con el rango para generar RESUMEN DE MOVIMIENTO DE VALVULAS, segùn la cantidad de mandriles
        Hoja.Range(Hoja.cells(lngRow + 6, 4), Hoja.cells(lngRow + 8, 4)).Select
        
        'COPY
        Obj.Selection.Copy
        
        'PASTE - Utilizo el mismo contador de filas, para las columnas, asi no tengo que definir nada nuevo
        Hoja.cells(lngRow + 6, lngRow - 21).Select
        Obj.activesheet.Paste
        
        'ASSIGN - nombre de mandril
        Hoja.cells(lngRow + 6, lngRow - 21) = rs!Mandrel_Number
        
        '
        '-------------------------------------------------------------------------------------------------------
        
        
        'NEXT avanzo puntero a fila siguiente
        lngRow = lngRow + 1
        
        'MOVE al siguiente registro
        rs.MoveNext
        
    Loop
    
    'DELETE el ultimo Copy, para que en el excel que se genera, no tenga nada seleccionado
    Obj.CutCopyMode = False
    
    
    'SELECT fila 12, para que cuando guarde el excel, quede el puntero en donde deben comenzar a ingresar datos
    Hoja.cells(12, 3).Select
    
      
    
    'ENABLE proteccion
    Hoja.Protect Password:="Sinopec2015"
    
    
    
    'SAVE template, con Fecha Actual y Nombre de Pozo
    Obj.ActiveWorkbook.SaveAs FileName:=Me.cmd_Destino + "ValvesMovement_" + dateToIso(Now()) + "_" + UCase(txt_Pozo) + ".xlsx"

  
    'CLOSE objetos
    Libro.Close
'    Set Obj = Nothing
    Obj.Quit
    Set Libro = Nothing
    Set Hoja = Nothing

  'mouse defa
  Screen.MousePointer = vbDefault



End Sub



Private Sub Form_Load()

  Dim rs As ADODB.Recordset
  Dim strT, strL As String
  Dim intI As Integer
  Dim blnB As Boolean
  
  'get parametros de conexion
  blnB = SQLgetParam()
  
  'check si parametros ok, get menu
  If blnB Then
  End If
    
  'Width columna 1, en donde se guardan los Queries
  Me.spdSQL.ColWidth(1) = 6
    
  'Width columna 1, en donde se guardan los Queries
  Me.spdSQL.ColWidth(2) = 115
    
  'HIDE titulo de columna 1
  Me.spdSQL.ColHeadersShow = False
    
  Me.spdSQL.Col = 2
  Me.spdSQL.TypeMaxEditLen = 500
  Me.spdSQL.FontName = "Courier New"
    
  Me.spdSQL.SetOddEvenRowColor RGB(245, 245, 245), RGB(60, 60, 60), RGB(245, 245, 245), RGB(60, 60, 60)
  Me.spdSQL.SelBackColor = RGB(204, 230, 255)         'fondo del area seleccionada
  Me.spdSQL.GrayAreaBackColor = RGB(245, 245, 245)
    
  Me.spdSQL.BackColorStyle = BackColorStyleUnderGrid  'estilo
  Me.spdSQL.GridShowHoriz = True                      'muestra grilla horizontal
  Me.spdSQL.GridShowVert = True                       'muestra grilla vertical
  Me.spdSQL.GridColor = RGB(200, 200, 200)            'color muy suave
  Me.spdSQL.NoBorder = True                           'sin borde fin zona de datos
  
  
  
    'ASSIGN origen y destino para los template de Excel
    Me.cmd_Origen = App.Path + "\TEMPLATE\"
    Me.cmd_Destino = Environ$("USERPROFILE") + "\"

  
End Sub

Private Sub txt_Pozo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.txt_Pozo.SelStart = 0
    Me.txt_Pozo.SelLength = Len(Me.txt_Pozo)

End Sub
