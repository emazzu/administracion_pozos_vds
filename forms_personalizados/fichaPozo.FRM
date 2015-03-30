VERSION 5.00
Object = "{0002E558-0000-0000-C000-000000000046}#1.1#0"; "OWC11.DLL"
Begin VB.Form fichaPozo_FRM 
   Caption         =   "Ficha de Pozo"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14745
   Icon            =   "fichaPozo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   14745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   90
      TabIndex        =   5
      Top             =   -45
      Width           =   14595
      Begin VB.TextBox cmdDestino 
         Height          =   330
         Left            =   90
         TabIndex        =   2
         Text            =   "C:\DATA"
         Top             =   1170
         Width           =   10275
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "&Exportar"
         Height          =   330
         Left            =   10485
         TabIndex        =   3
         Top             =   810
         Width           =   1005
      End
      Begin VB.TextBox txtPozos 
         Height          =   690
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   450
         Width           =   10275
      End
      Begin VB.CommandButton cmd_Salir 
         Caption         =   "&Salir"
         Height          =   330
         Left            =   13635
         TabIndex        =   4
         Top             =   405
         Width           =   870
      End
      Begin VB.CommandButton cmdFichaPozo 
         Caption         =   "&Aceptar"
         Height          =   330
         Left            =   10485
         TabIndex        =   1
         Top             =   450
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Ingrese una lista de Pozos. Si usted desea visualizar más de uno, debe separarlos por una coma"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   225
         Width           =   7170
      End
   End
   Begin VB.Frame Frame3 
      Height          =   7620
      Left            =   90
      TabIndex        =   7
      Top             =   1485
      Width           =   14595
      Begin OWC11.Spreadsheet spdExp 
         Height          =   3840
         Left            =   12150
         OleObjectBlob   =   "fichaPozo.frx":038A
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1680
      End
      Begin OWC11.Spreadsheet spd 
         Height          =   7350
         Left            =   90
         OleObjectBlob   =   "fichaPozo.frx":0D0A
         TabIndex        =   6
         Top             =   165
         Width           =   14415
      End
   End
End
Attribute VB_Name = "fichaPozo_FRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngLeft, lngWidth As Long
Dim blnMax As Boolean

Dim intFilTitu, lngFmaxima, lngCmaxima As Integer
Dim colForm, colQpHk, colCeme, colRese, colComienzo, filComienzo, colEntreCadaOperacion As Integer


Private Sub cmd_Salir_Click()
  
  Unload Me
  
End Sub



Private Sub cmdExportar_Click()

  Dim blnB As Boolean
  Dim strT As String
  
  'mouse reloj
  Screen.MousePointer = vbHourglass
  
  'DELETE sheets oculta utilizada para Exportar
  For Each Sh In spdExp.Sheets
    
    If Sh.Name <> "Sheet1" Then
      Sh.Delete
    End If
        
  Next
  
  'EXPORT cada Sheets como un nuevo Excel
  For Each Sh In spd.Sheets
    
    If Sh.Name <> "Sheet1" Then
      
      'COPY configuracion
      spd.Sheets(Sh.Name).Cells.Copy
      
      'PASTE configuracion
      spdExp.Sheets("Sheet1").Cells.Paste
      
      'EXPORT Sheets
      spdExp.Export Me.cmdDestino & "\" & Sh.Name & ".xls", ssExportActionNone, ssExportAsAppropriate
    End If
        
  Next
  
  
  'mouse defa
  Screen.MousePointer = vbDefault
    
  blnB = MsgBox("La exportación se realizó con éxito.", vbInformation + vbOKOnly, "atención...")

End Sub

Private Sub cmdFichaPozo_Click()

  Dim rs, rsFm, rsEV As ADODB.Recordset
  Dim strT, strDesdeHastaAUX As String
  Dim dtmFechaAUX As Date
  Dim intI As Integer
  Dim lngF, lngC  As Long
  Dim arrColor() As String
  Dim arrPozos() As String
  Dim strFmDesde, strFmHasta, strFm, strRs, arrFm() As String
  Dim intCant As Integer
  
  '--------------------------------------------------------------------------------------------------
  '
  ' 1.  GET profundidades_operaciones, se reserva 1 fila para cada ensayo y 11 columnas para cada operacion
  '     arma para cada profundidad, todas las operaciones, luego cambia a otra profundidad
  '
  
  'CHECK si se ingreso algo
  If Me.txtPozos.Text = "" Then
    Exit Sub
  End If
  
  'CHANGE mouse reloj
  Screen.MousePointer = vbHourglass
  
  'SPLIT lista de pozos
  '
  
  'CHECK si existe un enter al final
  If Right(Me.txtPozos.Text, 2) = Chr(13) + Chr(10) Then
    Me.txtPozos.Text = Left(Me.txtPozos.Text, Len(Me.txtPozos.Text) - 2)
  End If
  
  'CHECK si existe una coma
  If InStr(Me.txtPozos.Text, ",") <> 0 Then
    
    arrPozos = Split(Me.txtPozos, ",")
    
  'CHECK si existe un enter
  ElseIf InStr(Me.txtPozos.Text, Chr(13) + Chr(10)) <> 0 And _
     InStr(Me.txtPozos.Text, Chr(13) + Chr(10)) <> Len(Me.txtPozos.Text) - 1 Then
    
    arrPozos = Split(Me.txtPozos, Chr(13) + Chr(10))
    
  'CHECK un pozo
  Else
  
    arrPozos = Split(Me.txtPozos)
  
  End If
  
  'DELETE sheets excluyendo la 1
  For Each Sh In spd.Sheets
    
    If Sh.Name <> "Sheet1" Then
      Sh.Delete
    End If
        
  Next
  
  'DISABLE dibujo mientras construye
  spd.ScreenUpdating = False
  
  'WHILE pozos
  intCant = 0
  Do While intCant <= UBound(arrPozos)
      
      'COPY configuracion
      spd.Sheets("Sheet1").Cells.Copy
      
      'ADD sheet con el nombre del pozo a graficar
      spd.Sheets.Add
      spd.ActiveSheet.Name = UCase(Trim(arrPozos(intCant)))
      
      'PASTE configuracion
      spd.ActiveSheet.Cells.Paste
      
      'SELECT fila 1, columna 1
      spd.ActiveCell.Cells(1, 1).Select
      
      'DELETE y CLEAR grilla
      spd.Range("A1", "ZZZ500").Delete
      
      'build query - profundidades
      strT = "select * from fichaPozo_profundidades_operaciones_vw where WellID = '" & UCase(Trim(arrPozos(intCant))) & "' order by desde, fecha, operacion"
      Set rs = SQLexec(strT)
      
      'chequeo error
      If Not SQLparam.CnErrNumero = -1 Then
        SQLError
        SQLclose
        End
      End If
        
      'build query - formaciones OPEN WORKS
      strT = "select * from OW_Formations where left(PICK_SURF_NAME, 1) = 'f' and WELL_NAME_FREE = '" & UCase(Trim(arrPozos(intCant))) & "'"
      Set rsFm = SQLexec(strT)
      
      'chequeo error
      If Not SQLparam.CnErrNumero = -1 Then
        SQLError
        SQLclose
        End
      End If
        
      'clear contador de celdas en grilla
      lngF = filComienzo
      lngC = colComienzo
        
      'WHILE rs
      Do While Not rs.EOF
        
        'SAVE rs!desde_hasta anterior
        strDesdeHastaAUX = rs!desde_hasta
        
        'ADD 1 a contador fila
        lngF = lngF + 2
        
        'CLEAR contador columnas
        lngC = colComienzo
        
        'WHILE rs
        Do While Not rs.EOF
          
          'CHECK si valor rs!desde_hasta actual <> a anterior
          If rs!desde_hasta <> strDesdeHastaAUX Then
            Exit Do
          End If
          
          'SAVE rs!fecha anterior
          dtmFechaAUX = rs!fecha
          
          'ADD encabezados
          spd.Cells(2, lngC) = rs!fecha & " - " & rs!operacion
          spd.Cells(2, lngC).Font.Bold = True
          spd.Rows(2).RowHeight = 20
          
          spd.Cells(4, 1) = "Form"
          spd.Columns(1).ColumnWidth = 4
    
          spd.Cells(4, 2) = "Qp/Hk"
          spd.Cells(5, 2) = "(lts/h m)"
          spd.Columns(2).ColumnWidth = 7
          spd.Cells(4, 2).HorizontalAlignment = xlCenter
          spd.Cells(5, 2).HorizontalAlignment = xlCenter
          
          spd.Cells(4, 3) = "Cem"
          spd.Columns(3).ColumnWidth = 3.5
          
          spd.Cells(4, 4) = "Reservoir"
          spd.Columns(4).ColumnWidth = 8
          
          'CHANGE ancho columnas para colores pozo
          spd.Columns(lngC + colRese - 4).ColumnWidth = 0.8
          spd.Columns(lngC + colRese - 3).ColumnWidth = 0.8
          spd.Columns(lngC + colRese - 2).ColumnWidth = 0.8
          
          spd.Cells(4, lngC + colRese - 1) = "Test#"
          spd.Columns(lngC + colRese - 1).ColumnWidth = 4.5
          
          spd.Cells(4, lngC + colRese) = "From"
          spd.Cells(5, lngC + colRese) = "(m)"
          spd.Columns(lngC + colRese).ColumnWidth = 6.5
          spd.Columns(lngC + colRese).NumberFormat = "0.0"
          spd.Cells(4, lngC + colRese).HorizontalAlignment = xlCenter
          spd.Cells(5, lngC + colRese).HorizontalAlignment = xlCenter
          
          spd.Cells(4, lngC + colRese + 1) = "-"
          spd.Cells(5, lngC + colRese + 1) = ""
          spd.Columns(lngC + colRese + 1).ColumnWidth = 1
          spd.Cells(4, lngC + colRese + 1).HorizontalAlignment = xlCenter
          spd.Cells(5, lngC + colRese + 1).HorizontalAlignment = xlCenter
          
          spd.Cells(4, lngC + colRese + 2) = "To"
          spd.Cells(5, lngC + colRese + 2) = "(m)"
          spd.Columns(lngC + colRese + 2).ColumnWidth = 6.5
          spd.Columns(lngC + colRese + 2).NumberFormat = "0.0"
          spd.Cells(4, lngC + colRese + 2).HorizontalAlignment = xlCenter
          spd.Cells(5, lngC + colRese + 2).HorizontalAlignment = xlCenter
          
          spd.Cells(4, lngC + colRese + 3) = "Fluid"
          
          spd.Cells(4, lngC + colRese + 4) = "QL"
          spd.Cells(5, lngC + colRese + 4) = "(lts/h)"
          spd.Columns(lngC + colRese + 4).ColumnWidth = 6
          spd.Columns(lngC + colRese + 4).NumberFormat = "0"
          spd.Cells(4, lngC + colRese + 4).HorizontalAlignment = xlCenter
          spd.Cells(5, lngC + colRese + 4).HorizontalAlignment = xlCenter
          
          spd.Cells(4, lngC + colRese + 5) = "BS&W"
          spd.Cells(5, lngC + colRese + 5) = "(%)"
          spd.Columns(lngC + colRese + 5).ColumnWidth = 6
          spd.Columns(lngC + colRese + 5).NumberFormat = "0"
          spd.Cells(4, lngC + colRese + 5).HorizontalAlignment = xlCenter
          spd.Cells(5, lngC + colRese + 5).HorizontalAlignment = xlCenter
          
          spd.Cells(4, lngC + colRese + 6) = "FL"
          spd.Cells(5, lngC + colRese + 6) = "(m)"
          spd.Columns(lngC + colRese + 6).ColumnWidth = 6
          spd.Columns(lngC + colRese + 6).NumberFormat = "0,0"
          spd.Cells(4, lngC + colRese + 6).HorizontalAlignment = xlCenter
          spd.Cells(5, lngC + colRese + 6).HorizontalAlignment = xlCenter
          
          spd.Cells(4, lngC + colRese + 7) = "QP"
          spd.Cells(5, lngC + colRese + 7) = "(lts/h)"
          spd.Columns(lngC + colRese + 7).ColumnWidth = 6
          spd.Columns(lngC + colRese + 7).NumberFormat = "0.0"
          spd.Cells(4, lngC + colRese + 7).HorizontalAlignment = xlCenter
          spd.Cells(5, lngC + colRese + 7).HorizontalAlignment = xlCenter
          
          'WIDTH columna que separa operaciones}
          spd.Columns(lngC + colRese + 8).ColumnWidth = 3
          
          
          'WHILE rs
          Do While Not rs.EOF
            
            'CHECK si profundidades y fechas distintas EXIT WHILE
            If rs!fecha <> dtmFechaAUX Or rs!desde_hasta <> strDesdeHastaAUX Then
              Exit Do
            End If
                    
            'ADD profundidad
            spd.Cells(lngF, lngC + colRese) = rs!desde
            spd.Cells(lngF, lngC + colRese + 1) = "/"
            spd.Cells(lngF, lngC + colRese + 2) = rs!hasta
            
            'CENTER celda profundidad
            spd.Cells(lngF, lngC + colRese).HorizontalAlignment = xlCenter
            spd.Cells(lngF, lngC + colRese + 1).HorizontalAlignment = xlCenter
            spd.Cells(lngF, lngC + colRese + 2).HorizontalAlignment = xlCenter
                    
            'ADD profundidad y fecha para busqueda en rs y poder insertar la info
            spd.Cells(lngF, lngC + colRese - 1) = rs!desde_hasta & " - " & rs!fecha
            
            'FIND formacion
            'La produndidad desde_hasta del punzado debe caer dentro de los limites de la formacion
            
            'CLEAR
            strFmDesde = ""
            strFmHasta = ""
            
            'FILTER limite inferior
            rsFm.Filter = "PICK_DEPTH < " & CSng(rs!desde)
            rsFm.Sort = "PICK_DEPTH desc"
            
            'CHECK not eof
            If Not rsFm.EOF Then
              strFmDesde = rsFm!PICK_SURF_NAME
            End If
            
            'FILTER limite superior
            rsFm.Filter = "PICK_DEPTH > " & CSng(rs!desde)
            rsFm.Sort = "PICK_DEPTH asc"
            
            'CHECK not eof
            If Not rsFm.EOF Then
              strFmHasta = rsFm!PICK_SURF_NAME
            End If
            
            'CHECK si formacion desde = hasta, ADD nombre de formacion
            If Replace(strFmDesde, "_T", "") = Replace(strFmHasta, "_B", "") Then
              
              'SEPARATE formacion y reservorio
              arrFm = Split(strFmDesde, "-")
              strFm = ""
              strRs = ""
            
              'CHECK si array contiene 1 elemento
              If UBound(arrFm) = 1 Then
                strFm = arrFm(0)
                strRs = arrFm(1)
              End If
              
              'CHECK si array contiene 2 elemento
              If UBound(arrFm) = 2 Then
                strFm = arrFm(0) & arrFm(1)
                strRs = arrFm(2)
              End If
                        
              'PUT formacion y reservocio
              spd.Cells(lngF, colForm) = strFm
              spd.Cells(lngF, colRese) = strRs
              
            End If
            
            'ADD contador de operaciones
            lngC = lngC + colEntreCadaOperacion
            
            
            'NEXT rs
            rs.MoveNext
            
          Loop
          
        Loop
        
      Loop
      
     
      'SET filas y columnas maximas utilizadas
      lngFmaxima = lngF
      lngCmaxima = lngC
      
      
      '--------------------------------------------------------------------------------------------------
      '
      ' 2.  GET ensayos - operaciones, completa la informacion de cada operacion
      '
    
      'build query - operaciones
      strT = "select * from fichaPozo_eventos_vw where wellID = '" & UCase(Trim(arrPozos(intCant))) & "' order by desde, fecha, operacion, testNum"
      Set rs = SQLexec(strT)
    
      'chequeo error
      If Not SQLparam.CnErrNumero = -1 Then
        SQLError
        SQLclose
        End
      End If
      
      'CLEAR contador de celdas en grilla
      lngF = filComienzo
      lngC = colComienzo
      
      'WHILE operaciones
      Do While lngC < lngCmaxima
        
        'WHILE profundidades
        Do While lngF <= lngFmaxima
                
          'CHECK si celda con datos
          If Me.spd.Cells(lngF, lngC + colRese - 1) <> "" Then
                    
            'FILTER eventos para que correspodan a cada operacion
            rs.Filter = "desde_hasta_fecha = '" & Me.spd.Cells(lngF, lngC + colRese - 1) & "'"
            rs.Sort = "TestNum"
                    
            'CHECK si hay registros
            If Not rs.EOF Then
              
              'MOVE inicio
              rs.MoveFirst
              
              'CLEAR contador operaciones
              intI = 0
              
              'WHILE operaciones
              While Not rs.EOF
                
                'ADD test#, Resultado, QL, BS&W, FL, QP
                spd.Cells(lngF, lngC + colRese - 1) = rs!testNum
                spd.Cells(lngF, lngC + colRese + 3) = rs!resultado
                spd.Cells(lngF, lngC + colRese + 4) = rs!QL
                spd.Cells(lngF, lngC + colRese + 5) = rs!BSWensayo
                spd.Cells(lngF, lngC + colRese + 6) = rs!FLensayo
                spd.Cells(lngF, lngC + colRese + 7) = rs!QP
                spd.Cells(lngF, colQpHk) = rs!Qp_Hk
                
                'CHECK si tema = TAPON
                If rs!tema = "TAPON" Then
                  
                  'SET negrita
                  spd.Cells(lngF, lngC + colRese + 3).Font.Bold = True
                  
'                  'CALC rango
                  strR1 = Left(Replace(spd.Columns(colRese + 1).Address, "$", ""), InStr(Replace(spd.Columns(colRese + 1).Address, "$", ""), ":") - 1) & lngF
                  strR2 = Left(Replace(spd.Columns(colRese + 3).Address, "$", ""), InStr(Replace(spd.Columns(colRese + 3).Address, "$", ""), ":") - 1) & lngF
                  
                  'DRAW bordes
                  spd.Range(strR1, strR2).Borders(xlEdgeLeft).LineStyle = 1
                  spd.Range(strR1, strR2).Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
                  spd.Range(strR1, strR2).Borders(xlEdgeLeft).Weight = 3

                  spd.Range(strR1, strR2).Borders(xlEdgeTop).LineStyle = 1
                  spd.Range(strR1, strR2).Borders(xlEdgeTop).Color = RGB(0, 0, 0)
                  spd.Range(strR1, strR2).Borders(xlEdgeTop).Weight = 3

                  spd.Range(strR1, strR2).Borders(xlEdgeBottom).LineStyle = 1
                  spd.Range(strR1, strR2).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
                  spd.Range(strR1, strR2).Borders(xlEdgeBottom).Weight = 3

                  spd.Range(strR1, strR2).Borders(xlEdgeRight).LineStyle = 1
                  spd.Range(strR1, strR2).Borders(xlEdgeRight).Color = RGB(0, 0, 0)
                  spd.Range(strR1, strR2).Borders(xlEdgeRight).Weight = 3

                  spd.Range(strR1, strR2).Borders(xlInsideVertical).LineStyle = 1
                  spd.Range(strR1, strR2).Borders(xlInsideVertical).Color = RGB(0, 0, 0)
                  spd.Range(strR1, strR2).Borders(xlInsideVertical).Weight = 3
                  
                End If
                
                'CHECK si tema = PACKER
                If rs!tema = "PACKER" Then
                  
                  'SET negrita
                  spd.Cells(lngF, lngC + colRese + 3).Font.Bold = True
                  
                  'CALC rango celda izquierda
                  strR1 = Left(Replace(spd.Columns(colRese + 1).Address, "$", ""), InStr(Replace(spd.Columns(colRese + 1).Address, "$", ""), ":") - 1) & lngF
                  strR2 = Left(Replace(spd.Columns(colRese + 1).Address, "$", ""), InStr(Replace(spd.Columns(colRese + 1).Address, "$", ""), ":") - 1) & lngF
                  
                  'DRAW bordes
                  spd.Range(strR1, strR2).Borders(xlEdgeLeft).LineStyle = 1
                  spd.Range(strR1, strR2).Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
                  spd.Range(strR1, strR2).Borders(xlEdgeLeft).Weight = 3

                  spd.Range(strR1, strR2).Borders(xlEdgeTop).LineStyle = 1
                  spd.Range(strR1, strR2).Borders(xlEdgeTop).Color = RGB(0, 0, 0)
                  spd.Range(strR1, strR2).Borders(xlEdgeTop).Weight = 3

                  spd.Range(strR1, strR2).Borders(xlEdgeBottom).LineStyle = 1
                  spd.Range(strR1, strR2).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
                  spd.Range(strR1, strR2).Borders(xlEdgeBottom).Weight = 3

                  spd.Range(strR1, strR2).Borders(xlEdgeRight).LineStyle = 1
                  spd.Range(strR1, strR2).Borders(xlEdgeRight).Color = RGB(0, 0, 0)
                  spd.Range(strR1, strR2).Borders(xlEdgeRight).Weight = 3

                  'CALC rango celda derecha
                  strR1 = Left(Replace(spd.Columns(colRese + 3).Address, "$", ""), InStr(Replace(spd.Columns(colRese + 3).Address, "$", ""), ":") - 1) & lngF
                  strR2 = Left(Replace(spd.Columns(colRese + 3).Address, "$", ""), InStr(Replace(spd.Columns(colRese + 3).Address, "$", ""), ":") - 1) & lngF
                  
                  'DRAW bordes
                  spd.Range(strR1, strR2).Borders(xlEdgeLeft).LineStyle = 1
                  spd.Range(strR1, strR2).Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
                  spd.Range(strR1, strR2).Borders(xlEdgeLeft).Weight = 3

                  spd.Range(strR1, strR2).Borders(xlEdgeTop).LineStyle = 1
                  spd.Range(strR1, strR2).Borders(xlEdgeTop).Color = RGB(0, 0, 0)
                  spd.Range(strR1, strR2).Borders(xlEdgeTop).Weight = 3

                  spd.Range(strR1, strR2).Borders(xlEdgeBottom).LineStyle = 1
                  spd.Range(strR1, strR2).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
                  spd.Range(strR1, strR2).Borders(xlEdgeBottom).Weight = 3

                  spd.Range(strR1, strR2).Borders(xlEdgeRight).LineStyle = 1
                  spd.Range(strR1, strR2).Borders(xlEdgeRight).Color = RGB(0, 0, 0)
                  spd.Range(strR1, strR2).Borders(xlEdgeRight).Weight = 3

                  'CALC rango celda del medio hasta fila 7
                  strR1 = Left(Replace(spd.Columns(colRese + 2).Address, "$", ""), InStr(Replace(spd.Columns(colRese + 2).Address, "$", ""), ":") - 1) & 7
                  strR2 = Left(Replace(spd.Columns(colRese + 2).Address, "$", ""), InStr(Replace(spd.Columns(colRese + 2).Address, "$", ""), ":") - 1) & lngF
                  
                  'CLEAR color celda del medio
                  spd.Range(strR1, strR2).Clear
                  
                  'DRAW bordes
                  spd.Range(strR1, strR2).Borders(xlEdgeLeft).LineStyle = 1
                  spd.Range(strR1, strR2).Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
                  spd.Range(strR1, strR2).Borders(xlEdgeLeft).Weight = 3

                  spd.Range(strR1, strR2).Borders(xlEdgeTop).LineStyle = 1
                  spd.Range(strR1, strR2).Borders(xlEdgeTop).Color = RGB(0, 0, 0)
                  spd.Range(strR1, strR2).Borders(xlEdgeTop).Weight = 3

                  spd.Range(strR1, strR2).Borders(xlEdgeBottom).LineStyle = 1
                  spd.Range(strR1, strR2).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
                  spd.Range(strR1, strR2).Borders(xlEdgeBottom).Weight = 3

                  spd.Range(strR1, strR2).Borders(xlEdgeRight).LineStyle = 1
                  spd.Range(strR1, strR2).Borders(xlEdgeRight).Color = RGB(0, 0, 0)
                  spd.Range(strR1, strR2).Borders(xlEdgeRight).Weight = 3
                  
                End If
                
                'CHANGE format
                spd.Cells(lngF, colQpHk).NumberFormat = "0.0"
                
                'CHECK si cement
                If rs!cement Then
                  
                  'PAINT cemento
                  spd.Cells(lngF, colCeme).Interior.Color = RGB(175, 175, 175)
                  
                End If
                
                'CENTER profundidad y fecha, luego de completar la info en esta celda
                'se borra la profundidad y fecha, ya que es una celda reservada para Test #
                spd.Cells(lngF, lngC + colRese - 1).HorizontalAlignment = xlCenter
                
                'ASSIGN color
                strT = rs!Color
                
                'CHECK si null
                If IsNull(rs!Color) Then
                  strT = "255, 255, 255"
                End If
                
                'SEPARATE color
                arrColor = Split(strT, ",")
                            
                'CHECK si not null
                If rs!Color <> "" Then
                
                  'PAINT celda
                  spd.Cells(lngF, lngC + colRese - 4).Interior.Color = RGB(arrColor(0), arrColor(1), arrColor(2))
                  spd.Cells(lngF, lngC + colRese - 3).Interior.Color = RGB(arrColor(0), arrColor(1), arrColor(2))
                  spd.Cells(lngF, lngC + colRese - 2).Interior.Color = RGB(arrColor(0), arrColor(1), arrColor(2))
                
                End If
                
                'ADD contador filas
                lngF = lngF + 1
                
    '            'SET limite maximo de filas
    '            lngFmaxima = lngFmaxima + 1
                
                'CHECK operaciones > 1
                If Me.spd.Cells(lngF, lngC + colRese - 1) <> "" Then
                  
                  'INSERT fila para mantener espacio con profundidad siguiente
                  spd.Rows(lngF).Insert Shift:=xlDown
                  
                  'CLEAR celda para borrar color que arrastra luego de insertar fila nueva
                  spd.Cells(lngF, lngC + colRese - 4).Clear
                  spd.Cells(lngF, lngC + colRese - 3).Clear
                  spd.Cells(lngF, lngC + colRese - 2).Clear
                  
    '              'ADD contador filas
    '              lngF = lngF + 1
                  
                  'SET limite maximo de filas
                  lngFmaxima = lngFmaxima + 1
                  
                End If
                
                'MOVE proximo
                rs.MoveNext
                
              Wend
              
            End If 'CHECK si hay registros
            
          End If 'CHECK celda con datos
          
          'DELETE celda profundidad y fecha
          spd.Cells(lngF, lngC + colRese - 1) = ""
          
          'ADD contador filas
          lngF = lngF + 1
          
        Loop
        
        'AUTOFIT resultado de ensayos, se realiza en esta lugar para que ajuste luego de insertar la informacion
        spd.Columns(lngC + colRese + 3).AutoFit
        
        'ADD contador columnas
        lngC = lngC + colEntreCadaOperacion
        
        'CLEAR contador filas
        lngF = filComienzo
        
      Loop
    
      'close
      SQLclose
      
      'CALL dibuja formato
      Call cmd_FichaPozo_Formato(UCase(Trim(arrPozos(intCant))))
        
      'ADD contador de pozos
      intCant = intCant + 1
      
  Loop
  
  'ENABLE dibujo una vez finalizado
  spd.ScreenUpdating = True
  
  
  'CHANGE mouse estandar
  Screen.MousePointer = vbDefault
  
End Sub

Private Sub cmd_FichaPozo_Formato(strPozo As String)

  Dim ssConstants
  Dim intC, intF, intI As Integer
  Dim strR1, strR2 As String
  
  Set ssConstants = spd.Constants
      
  
  '------------------------------------------------------------------------------------------------------
  'DRAW borde completo
  '
  
  'GET limites filas y columnas
  intF = spd.ActiveSheet.UsedRange.Row - 1 + spd.ActiveSheet.UsedRange.Rows.Count
  intC = spd.ActiveSheet.UsedRange.Column - 1 + spd.ActiveSheet.UsedRange.Columns.Count
  
  'CALC rango
  strR1 = "A1"
  strR2 = Left(Replace(spd.Columns(intC).Address, "$", ""), InStr(Replace(spd.Columns(intC).Address, "$", ""), ":") - 1) & intF

  'DRAW arriba
  spd.Range(strR1, strR2).Borders(xlEdgeTop).LineStyle = 1
  spd.Range(strR1, strR2).Borders(xlEdgeTop).Color = RGB(0, 0, 0)
  spd.Range(strR1, strR2).Borders(xlEdgeTop).Weight = 3
  
  'DRAW abajo
  spd.Range(strR1, strR2).Borders(xlEdgeBottom).LineStyle = 1
  spd.Range(strR1, strR2).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
  spd.Range(strR1, strR2).Borders(xlEdgeBottom).Weight = 3
  
  'DRAW izquierda
  spd.Range(strR1, strR2).Borders(xlEdgeLeft).LineStyle = 1
  spd.Range(strR1, strR2).Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
  spd.Range(strR1, strR2).Borders(xlEdgeLeft).Weight = 3
  
  'DRAW derecha
  spd.Range(strR1, strR2).Borders(xlEdgeRight).LineStyle = 1
  spd.Range(strR1, strR2).Borders(xlEdgeRight).Color = RGB(0, 0, 0)
  spd.Range(strR1, strR2).Borders(xlEdgeRight).Weight = 3
  
  
  '------------------------------------------------------------------------------------------------------
  'DRAW borde operation
  '
  
  'CALC rango
  strR1 = "E2"
  strR2 = Left(Replace(spd.Columns(intC).Address, "$", ""), InStr(Replace(spd.Columns(intC).Address, "$", ""), ":") - 1) & 2

  'DRAW arriba
  spd.Range(strR1, strR2).Borders(xlEdgeTop).LineStyle = 1
  spd.Range(strR1, strR2).Borders(xlEdgeTop).Color = RGB(0, 0, 0)
  spd.Range(strR1, strR2).Borders(xlEdgeTop).Weight = 2
  
  'DRAW abajo
  spd.Range(strR1, strR2).Borders(xlEdgeBottom).LineStyle = 1
  spd.Range(strR1, strR2).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
  spd.Range(strR1, strR2).Borders(xlEdgeBottom).Weight = 2
  
  'DRAW izquierda
  spd.Range(strR1, strR2).Borders(xlEdgeLeft).LineStyle = 1
  spd.Range(strR1, strR2).Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
  spd.Range(strR1, strR2).Borders(xlEdgeLeft).Weight = 2
  
'  'DRAW derecha
'  spd.Range(strR1, strR2).Borders(xlEdgeRight).LineStyle = 1
'  spd.Range(strR1, strR2).Borders(xlEdgeRight).Color = RGB(0, 0, 0)
'  spd.Range(strR1, strR2).Borders(xlEdgeRight).Weight = 2
  
  
  '------------------------------------------------------------------------------------------------------
  'DRAW borde encabezado
  '
  
  'CALC rango
  strR1 = "A4"
  strR2 = Left(Replace(spd.Columns(intC).Address, "$", ""), InStr(Replace(spd.Columns(intC).Address, "$", ""), ":") - 1) & 5
  
  'DRAW arriba
  spd.Range(strR1, strR2).Borders(xlEdgeTop).LineStyle = 1
  spd.Range(strR1, strR2).Borders(xlEdgeTop).Color = RGB(0, 0, 0)
  spd.Range(strR1, strR2).Borders(xlEdgeTop).Weight = 1.5
  
  'DRAW abajo
  spd.Range(strR1, strR2).Borders(xlEdgeBottom).LineStyle = 1
  spd.Range(strR1, strR2).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
  spd.Range(strR1, strR2).Borders(xlEdgeBottom).Weight = 1.5
  
  
  '------------------------------------------------------------------------------------------------------
  'DRAW borde columna cementation
  '
  
  'CALC rango
  strR1 = "C4"
  strR2 = "C" & intF
  
  'DRAW izquierda
  spd.Range(strR1, strR2).Borders(xlEdgeLeft).LineStyle = 1
  spd.Range(strR1, strR2).Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
  spd.Range(strR1, strR2).Borders(xlEdgeLeft).Weight = 1.5

  'DRAW derecha
  spd.Range(strR1, strR2).Borders(xlEdgeRight).LineStyle = 1
  spd.Range(strR1, strR2).Borders(xlEdgeRight).Color = RGB(0, 0, 0)
  spd.Range(strR1, strR2).Borders(xlEdgeRight).Weight = 1.5
  
  
  '------------------------------------------------------------------------------------------------------
  'DRAW borde columnas well
  '
  
  'WHILE operaciones
  intI = 5
  Do While intI < intC
    
    'CALC rango
    strR1 = Left(Replace(spd.Columns(intI).Address, "$", ""), InStr(Replace(spd.Columns(intI).Address, "$", ""), ":") - 1) & 1
    strR2 = Left(Replace(spd.Columns(intI).Address, "$", ""), InStr(Replace(spd.Columns(intI).Address, "$", ""), ":") - 1) & intF
    
    'DRAW izquierda
    spd.Range(strR1, strR2).Borders(xlEdgeLeft).LineStyle = 1
    spd.Range(strR1, strR2).Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
    spd.Range(strR1, strR2).Borders(xlEdgeLeft).Weight = 2.5
    
    'CALC rango
    strR1 = Left(Replace(spd.Columns(intI + 2).Address, "$", ""), InStr(Replace(spd.Columns(intI + 2).Address, "$", ""), ":") - 1) & 4
    strR2 = Left(Replace(spd.Columns(intI + 2).Address, "$", ""), InStr(Replace(spd.Columns(intI + 2).Address, "$", ""), ":") - 1) & intF
    
    'DRAW derecha
    spd.Range(strR1, strR2).Borders(xlEdgeRight).LineStyle = 1
    spd.Range(strR1, strR2).Borders(xlEdgeRight).Color = RGB(0, 0, 0)
    spd.Range(strR1, strR2).Borders(xlEdgeRight).Weight = 2.5
        
    'ADD contador de operaciones
    intI = intI + colEntreCadaOperacion
    
  Loop
    
    
  '------------------------------------------------------------------------------------------------------
  'INSERT titulo reporte + nombre de pozo
  '
  
  'INSERT y CHANGE apariencia fila 1 para titulos
  spd.Rows(1).Insert
  spd.Rows(1).RowHeight = 25
  spd.Cells(1, 1) = "Ficha de Pozo"
  spd.Cells(1, 1).Font.Size = 16
  spd.Cells(1, 1).Font.Bold = True
  
  spd.Cells(1, 5) = strPozo
  spd.Cells(1, 5).Font.Size = 22
  spd.Cells(1, 5).Font.Bold = True
  
  
  
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
      
''  'get ubicaciones
''  Call fill_pozos("")
          
  'CLEAR ubicacion filas
  intFilTitu = 2
  
  'CLEAR ubicacion columnas
  colForm = 1
  colQpHk = 2
  colCeme = 3
  colRese = 4
  colComienzo = 5
  filComienzo = 5
  colEntreCadaOperacion = 13
  
End Sub


