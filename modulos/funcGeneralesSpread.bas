Attribute VB_Name = "funcGeneralesSpread"
'
' CAMBIA APARIENCIA A GRILLA HORIZONTAL
'
Function spdCustomize(ByRef spd As fpSpread) As Boolean

  spd.Appearance = AppearanceFlat               'apariencia 3D
  spd.BorderStyle = BorderStyleNone             'tipo de borde: sin borde
  
  spd.ColHeadersAutoText = DispBlank            'titulos de columnas en blanco
  spd.ColHeadersShow = True                     'muestra encabezado columnas
  spd.RowHeadersShow = True                     'muestra encabezado de filas
  
  spd.CursorStyle = CursorStyleArrow            'stilo cursor
  spd.CursorType = CursorTypeDefault            'tipo cursor
  
  spd.AutoSize = False                           'automaticamente ajusta ancho grilla
  spd.DAutoSizeCols = DAutoSizeColsMax          'tipo de ajuste 2: al dato mas ancho
  
  spd.UserColAction = UserColActionDefault      'cuando hace click en header pinta columna o fila
  spd.FontSize = 9                              'tamaño letra
  spd.RowHeight(0) = 20                         'altura fila de titulos
  spd.MoveActiveOnFocus = False                 '
  spd.Protect = False                           'exporta a excel sin proteccion
  
  spd.BackColorStyle = BackColorStyleUnderGrid  'estilo
  spd.GridShowHoriz = True                      'muestra grilla horizontal
  spd.GridShowVert = True                       'muestra grilla vertical
  spd.GridColor = RGB(200, 200, 200)            'color muy suave
  spd.NoBorder = True                           'sin borde fin zona de datos
  
  spd.ScrollBars = ScrollBarsBoth               'ambas barras de desplazamiento
  spd.ScrollBarExtMode = True                   'cuando sean necesarias
  spd.VScrollSpecial = False                     'barra especial
      
  intRes = spd.SetOddEvenRowColor(RGB(245, 245, 245), RGB(60, 60, 60), RGB(245, 245, 245), RGB(60, 60, 60))
  spd.SelBackColor = RGB(204, 230, 255)         'fondo del area seleccionada amarillo
  spd.GrayAreaBackColor = RGB(245, 245, 245)
  
  spd.VirtualMode = False                        ' ajusta rows al tamaño del recordset
  'spd.VirtualRows = 300                         ' rows a leer del virtual buffer
  'spd.VirtualScrollBuffer = True                ' scroll vertical lee de tantas rows del buffer
  
  'setea para mostrar tooltip en las celdas donde no se ve toda la info
  spd.TextTip = TextTipFixed
  spd.TextTipDelay = 250
  
End Function

'
'LLENA UN COMBO DE UNA GRILLA VERTICAL CON VARIAS POSIBILIDADES
'FUNCIONA CON UNA TABLA LLAMADA menuOpciones columna ComboBox
'
'1. pozo;select pozo, id from pozos     ' devuelve id
'2. equipo;select equipo from equipos   ' devuelve texto
'3. UNO,0,DOS,0,TRES,0,CUATRO,0         ' devuelve texto
'4. UNO,1,DOS,2,TRES,3,CUATRO,4         ' devuelve id
'
Function spdDataToCbo(ByRef spd As fpSpread, ByVal strColumna As String, ByVal strDataSelect As String) As Boolean
  Dim intInd  As Integer
  Dim strList, strItemData As String
  Dim rs As ADODB.Recordset
  Dim strArrayDatosSeparados As Variant
      
  'separo columnas de select
  strArrayDatosSeparados = separateText(strDataSelect, ";")
    
  'busca ubicacion del nombre de la columna y en el siguiente esta el select o lista a mano
  'ejemplo: select Idxxx, Nombrexxxx from xxxxxx o Uno,1,Dos,2,Tres,3
  For intInd = 1 To UBound(strArrayDatosSeparados) - 1 Step 2
    If LCase(strArrayDatosSeparados(intInd)) = "[" & LCase(strColumna) & "]" Then
      Exit For
    End If
  Next
    
  'chequeo que haya encontrado la palabra select
  If InStr(LCase(strArrayDatosSeparados(intInd + 1)), "select") <> 0 Then
  
    'abro recordset
    Set rs = SQLexec(strArrayDatosSeparados(intInd + 1))
  
    'chequeo errores
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      Exit Function
    End If
  
    'busco cantidad de columnas del rs para
    'saber si viene solo texto o texto e indice
    Select Case rs.Fields.Count
    
    Case 1  ' solo texto
    
      'recorro recordset
      strList = ""
      strItemData = ""
      While Not rs.EOF
        'armo descripciones e indices
        strList = strList & rs(0) & vbTab
        strItemData = strItemData & "0" & vbTab
        
        'avanzo puntero
        rs.MoveNext
      
      Wend
      
    Case 2      ' texto e indice
           
      'recorro recordset
      strList = ""
      strItemData = ""
      While Not rs.EOF
        
        'busco en cual columna viene el texto y en cual el indice
        Select Case rs(0).Type
        Case conChar, conNchar, conVarchar, conText, conNVarchar, conText
          strList = strList & rs(0) & vbTab
          strItemData = strItemData & rs(1) & vbTab
        Case Else
          strList = strList & rs(1) & vbTab
          strItemData = strItemData & rs(0) & vbTab
        End Select
      
        'avanzo puntero
        rs.MoveNext
      
      Wend
      rs.Close
    
    End Select
    
  'lista de datos a escrita a mano: Uno,1,Dos,2,Tres,3 o Uno,0,Dos,0,Tres,0,Cuatro,0
  Else
    
    'separo los elementos de la lista
    Dim strArraySeparoLista As Variant
    strArraySeparoLista = separateText(strArrayDatosSeparados(intInd + 1), ",")
      
    'recorro array
    strList = ""
    strItemData = ""
    For intInd = 1 To UBound(strArraySeparoLista) - 1 Step 2
        
      strList = strList & strArraySeparoLista(intInd) & vbTab
      strItemData = strItemData & strArraySeparoLista(intInd + 1) & vbTab
      
    Next
    
  End If
    
  'agrego la lista al comboBox de la grilla, columna 2, fila actual
  spd.col = 2
  spd.TypeComboBoxList = strList
    
  'agrego ItemData al combobox de la grilla columna 3 invisible, fila actual
  spd.col = 3
  spd.TypeComboBoxList = strItemData
    
  'set vuelvo a columna 2
  spd.col = 2
    
End Function

