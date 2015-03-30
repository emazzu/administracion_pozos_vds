Attribute VB_Name = "funcSpread"
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
'ARMA GRILLA VERTICAL EN BASE A LA GRILLA
'HORIZONTAL PARA PODER EDITAR LOS DATOS
'FUNCIONA CON UNA TABLA LLAMADA menuOpciones
'Y DIBUJA TODO AUTOMATICAMENTYE AJUSTANDO FRM
'
Function spdEdit(ByVal frmBase As Form, ByRef frmEdit As Form, ByVal strOperacion As String, Optional blnAjustable As Boolean) As Boolean
  Dim a As fpSpread
  Dim varDefault As Variant
  Dim intDecimales As Variant
  Dim intCantDecimales As Integer
   
  'separo en un array los valores por default
  'ejemplo: [telefono],4361-0000,[pais],Argentina
  varDefault = separateText(frmBase.DataValorDefault)
  
  'cambio caption del form edit
  frmEdit.Caption = frmBase.Caption & " - " & UCase(strOperacion)
  
  'set apariencia de borde
  frmEdit.spdEdit.Appearance = Appearance3D
      
  'set no muestra encabezado fila y columna
  frmEdit.spdEdit.ColHeadersShow = False
  frmEdit.spdEdit.RowHeadersShow = False
  
  'barra de desplazamiento vertical solo cuando la necesita
  frmEdit.spdEdit.ScrollBars = ScrollBarsVertical
  frmEdit.spdEdit.ScrollBarExtMode = True
      
  'set muestro grilla horizontal vertical y color de fondo debajo de grilla
  frmEdit.spdEdit.BackColorStyle = BackColorStyleUnderGrid
  frmEdit.spdEdit.GridShowHoriz = True
  frmEdit.spdEdit.GridShowVert = True
          
  'set enter fila siguiente
  frmEdit.spdEdit.EditEnterAction = EditEnterActionDown
          
  'set cuando ingreso valor arriba de otro lo reemplaza
  frmEdit.spdEdit.EditModeReplace = True
          
  'set lineas de fondo
  frmEdit.spdEdit.GridColor = RGB(192, 192, 192)
      
  'set area fuera de la grilla blanco
  frmEdit.spdEdit.GrayAreaBackColor = RGB(255, 255, 255)
      
  'estilo borde grilla no
  frmEdit.spdEdit.BorderStyle = BorderStyleFixedSingle
      
  frmEdit.spdEdit.NoBorder = True
      
  'setea para mostrar tooltip en las celdas donde no se ve toda la info
  frmEdit.spdEdit.TextTip = TextTipFixed
  frmEdit.spdEdit.TextTipDelay = 250
      
  'maximo de columnas
  frmEdit.spdEdit.MaxCols = 3
      
  'set ancho columna 1
  'frmEdit.spdEdit.ColWidth(1) = (Screen.Height * 0.001649)
  frmEdit.spdEdit.ColWidth(1) = 25
  
  'set ancho columna 2
  'frmEdit.spdEdit.ColWidth(2) = (Screen.Height * 0.00217)
  frmEdit.spdEdit.ColWidth(2) = 25
      
  'set oculto columna 3 para indice de combo
  frmEdit.spdEdit.col = 3
  frmEdit.spdEdit.ColHidden = True
      
  'dim variables cuenta columnas y toma titulo
  Dim intCol, intUltimaFilaAgregada As Integer
  Dim sngAlturaTotalFilas As Single
  Dim varTitulo, varDato As Variant
                
  intUltimaFilaAgregada = 0
  sngAlturaTotalFilas = 0
                
  'recorro columnas para armar grilla
  For intCol = 1 To frmBase.spdGrid.MaxCols
        
    'set columna 1 donde se encuentran los titulos de las columnas
    frmEdit.spdEdit.col = 1
        
    'get nombre de columna de grilla base
    intRes = frmBase.spdGrid.GetText(intCol, 0, varTitulo)
     
    'tomo cantidad de decimales de columna actual si es de tipo numerica
    frmBase.spdGrid.row = frmBase.spdGrid.ActiveRow
    frmBase.spdGrid.col = intCol
    
    'tomo decimales columna actual
    If frmBase.spdGrid.CellType = CellTypeNumber Then
      intDecimales = frmBase.spdGrid.TypeNumberDecPlaces
    End If
  
    'busco columna que no este definida como no se muestran en edit
    If InStr(LCase(frmBase.DataNoMuestraEnEdit), "[" & LCase(varTitulo) & "]") = 0 Then
        
      intUltimaFilaAgregada = intUltimaFilaAgregada + 1
        
      'inserto fila
      frmEdit.spdEdit.InsertRows intUltimaFilaAgregada, 1
          
      'seteo maximo fila
      frmEdit.spdEdit.MaxRows = intUltimaFilaAgregada
        
      'set nombre de columna en grilla edit
      frmEdit.spdEdit.SetText 1, intUltimaFilaAgregada, varTitulo
       
      'set dato en grilla edit
      frmEdit.spdEdit.SetText 2, intUltimaFilaAgregada, varDato
        
      'set puntero en fila y columna actual para cambiar propiedades
      frmEdit.spdEdit.col = 1
      frmEdit.spdEdit.row = intUltimaFilaAgregada
        
      'set altura de fila porcentaje de la altura maxima de screen
      frmEdit.spdEdit.RowHeight(intUltimaFilaAgregada) = (Screen.Height * 0.0012)
        
      'set alineacion vertical
      frmEdit.spdEdit.TypeVAlign = TypeVAlignCenter

      frmEdit.spdEdit.RowHeight(intUltimaFilaAgregada) = 12
      
      'set color fondo gris para titulos
      frmEdit.spdEdit.BackColor = RGB(240, 240, 240)
          
      'set color texto gris para titulos
      frmEdit.spdEdit.ForeColor = RGB(131, 131, 131)
      
      'tamaño de letra para titulos
      frmEdit.spdEdit.FontSize = 10
      
      'set titulos estaticos
      frmEdit.spdEdit.CellType = CellTypeStaticText
      
      'set columna 2 donde se encuentran los datos a actualizar
      frmEdit.spdEdit.col = 2
          
      'set color de fondo para ingreso de datos
      frmEdit.spdEdit.BackColor = RGB(255, 255, 255)
          
      'set color de texto para ingreso de datos
      frmEdit.spdEdit.ForeColor = RGB(0, 0, 0)
          
      'tamaño de letra para ingreso de datos
      frmEdit.spdEdit.FontSize = 9
          
      'case tipo de celda para formatear y alinear
      Select Case frmBase.DataFields(varTitulo).Type
          
      'Bit
      Case conBit
        
        frmEdit.spdEdit.CellType = CellTypeCheckBox
        frmEdit.spdEdit.Value = 0
          
      'enteros
      Case conSmallInt, conInt, conTinyInt
        
        frmEdit.spdEdit.CellType = CellTypeNumber
        frmEdit.spdEdit.TypeNumberDecPlaces = 0
        frmEdit.spdEdit.TypeHAlign = TypeHAlignRight
        frmEdit.spdEdit.Value = 0
                    
      'decimal
      Case conMoney, conSmallMoney, conReal, conFloat, conNumeric, conDecimal
            
        frmEdit.spdEdit.CellType = CellTypeNumber
        frmEdit.spdEdit.TypeHAlign = TypeHAlignRight
        frmEdit.spdEdit.Value = 0
          
        'set decimales
        frmEdit.spdEdit.TypeNumberDecPlaces = intDecimales
                    
      'fecha
      Case conSmallDateTime, conDateTime
      
        frmEdit.spdEdit.CellType = CellTypeDate
        frmEdit.spdEdit.Text = "06/17/00"
                                        
      'string
      Case conChar, conNchar, conVarchar, conText, conNVarchar, conText
        
        'set ComboBox si es una columna que se muestra en un comboBox
        'tambien armo una fila de tipo comboBox en la columan 3 para
        'guardar los index de cada combo
        If InStr(LCase(frmBase.DataComboBox), "[" & LCase(varTitulo) & "]") <> 0 Then
          frmEdit.spdEdit.CellType = CellTypeComboBox
          frmEdit.spdEdit.TypeComboBoxEditable = False
          frmEdit.spdEdit.col = 3
          frmEdit.spdEdit.CellType = CellTypeComboBox
          frmEdit.spdEdit.TypeComboBoxEditable = False
          frmEdit.spdEdit.col = 2
          'lleno combo con datos
          intRes = spdDataToCbo(frmEdit.spdEdit, varTitulo, frmBase.DataComboBox)
        Else
          'sino se muestra solo un texto
          frmEdit.spdEdit.CellType = CellTypeEdit
          
          'set limite 500 caracteres
          frmEdit.spdEdit.TypeMaxEditLen = 500
          
          'set columna con .... cuando no se ve todo el texto
          frmEdit.spdEdit.TypeEllipses = True
          
        End If
                                       
      End Select
         
      'set valores por default, pueden ser fijos o pueden venir en un select
      If InStr(LCase(frmBase.DataValorDefault), "[" & LCase(varTitulo) & "]") <> 0 Then
        frmEdit.spdEdit.Text = arrayGetValue(varDefault, "[" & LCase(varTitulo) & "]")
      End If
          
      'si operacion es un U de update o D de delete busco el valor de la
      'columna en grilla horizontal luego lo guardo en grilla vertical
      If strOperacion = "editar" Or strOperacion = "eliminar" Or strOperacion = "consultar" Then
        intRes = frmBase.spdGrid.GetText(intCol, frmBase.spdGrid.ActiveRow, varDato)
        intRes = frmEdit.spdEdit.SetText(2, intCol, varDato)
      End If
         
      'set color texto negro columna 2
      frmEdit.spdEdit.ForeColor = RGB(0, 0, 0)
          
      'set alineacion vertical columna 2
      frmEdit.spdEdit.TypeVAlign = TypeVAlignCenter
          
      'set lock y Backcolor cuando columna es No Permite Edit
      If InStr(LCase(frmBase.DataSoloLecturaEnEdit), "[" & LCase(varTitulo) & "]") <> 0 Then
        frmEdit.spdEdit.LockBackColor = RGB(220, 220, 220)
        frmEdit.spdEdit.Lock = True
      End If
    
      'se Backcolor cuando columna es obligatoria
      If InStr(LCase(frmBase.DataObligatorioEnEdit), "[" & LCase(varTitulo) & "]") <> 0 Then
        frmEdit.spdEdit.BackColor = RGB(204, 230, 255)
      End If
    
      'set lock si operacion es Consulta
      If strOperacion = "consultar" Then
        frmEdit.spdEdit.Lock = True
      End If
    
      'sumo altura de fila actual para determinar altura total de la grilla
      sngAlturaTotalFilas = sngAlturaTotalFilas + frmEdit.spdEdit.RowHeight(frmEdit.spdEdit.MaxRows) + 0.28
    
    End If
    
  Next
      
  'si es formulario ajustable
  If blnAjustable Then
    
    'ancho columna 1 con titulos ajusta al mas ancho automaticamente
    'si el campo mas ancho es menor que 25 por default es 25
    Dim dblAnchoMaximo As Double
    dblAnchoMaximo = frmEdit.spdEdit.MaxTextColWidth(1)
  
    If dblAnchoMaximo < 25 Then
      dblAnchoMaximo = 25
    End If
  
    frmEdit.spdEdit.ColWidth(1) = dblAnchoMaximo
      
    'ancho columna 2 en donde se ingresan los datos igual
    'al ancho de titulos para que quede una grilla pareja
    frmEdit.spdEdit.ColWidth(2) = dblAnchoMaximo
      
    'ancho columna 3 para indice de combo oculta
    frmEdit.spdEdit.col = 3
    frmEdit.spdEdit.ColHidden = True
          
    'ancho de grilla dinamico es la suma del ancho de la
    'columna 1 + 2 pero primero debo convertir a twips
    Dim lngAnchoGrilla As Long
    frmEdit.spdEdit.ColWidthToTwips (frmEdit.spdEdit.ColWidth(1) + frmEdit.spdEdit.ColWidth(2)), lngAnchoGrilla
    'frmEdit.spdEdit.Width = lngAnchoGrilla + 100
    frmEdit.spdEdit.Width = lngAnchoGrilla + 50
        
    'alto de grilla dinamico es la suma de la altura de
    'todas las filas, pero primero debo convertir a twips
    'le agrego un 6 por ciento porque si no funciona mal
    Dim lngAltoGrilla As Long
    frmEdit.spdEdit.RowHeightToTwips 1, sngAlturaTotalFilas, lngAltoGrilla
    frmEdit.spdEdit.Height = lngAltoGrilla + lngAltoGrilla * 0.06
    
    'ancho formulario
    frmEdit.Width = frmEdit.spdEdit.Width + 300
    
    'alto del formulario
    frmEdit.Height = frmEdit.spdEdit.Height + 1000
    
    'ubico grilla en form
    frmEdit.spdEdit.Left = 100
    frmEdit.spdEdit.Top = 100
    
    'cambio tamaño a botones aceptar
    frmEdit.cmdAceptar.Width = lngAnchoGrilla / 2 + 30
    frmEdit.cmdAceptar.Height = 300
  
    'cambio tamaño a botones cancelar
    frmEdit.cmdCancelar.Width = lngAnchoGrilla / 2 + 30
    frmEdit.cmdCancelar.Height = 300
      
    'cambio ubicacion de botones aceptar
    frmEdit.cmdAceptar.Left = 100
    frmEdit.cmdAceptar.Top = frmEdit.spdEdit.Height + 200
    
    'cambio ubicacion de botones cancelar
    frmEdit.cmdCancelar.Left = frmEdit.cmdAceptar.Left + frmEdit.cmdAceptar.Width + 30
    frmEdit.cmdCancelar.Top = frmEdit.spdEdit.Height + 200
  
  End If
  
  'selda activa fila 1 columna 2
  frmEdit.spdEdit.SetActiveCell 2, 1
  
  'muestra form
  frmEdit.Show vbModal

End Function

'
'ESTA FUNCION SE UTILIZA PARA CUANDO TENGO UNA GRILLA PERSONALIZADA, O SEA NO AUTOMATICA
'SE LE PASA LA GRILLA HORIZONTAL BASE Y EL NOMBRE DE LA GRILLA PERSONALIZADA
'PINTA LOS OBLIGATORIOS, BLOQUEA LO QUE SE HAYA DEFINIDO ASI Y SE ENCARGA DE TRANSFERIRLE
'LOS DATOS EN CASO DE EDICION Y ELIMINACION
'
Function spdEditSet(ByVal frmBase As Form, ByRef spdEdit As fpSpread, ByVal strOperacion As String) As Boolean
  Dim varTitulo, varDato As Variant
  Dim lngEncontroFila As Long
                
  'set columna 2 donde se encuentran los datos
  spdEdit.col = 2
                
  'recorro columnas
  For intCol = 1 To frmBase.spdGrid.MaxCols
    
    'tomo nombre de columna de grilla base
    intRes = frmBase.spdGrid.GetText(intCol, 0, varTitulo)
        
    'busco en grilla edit la posicion del nombre de la columna de grilla base
    lngEncontroFila = spdEdit.SearchCol(1, 1, frmBase.spdGrid.MaxCols, varTitulo, SearchFlagsNone)
    
    'si encontro nombre de columna
    If lngEncontroFila <> -1 Then
    
      'si editar o eliminar tomo valor grilla horizontal y pongo en grilla vertical
      If (strOperacion = "editar" Or strOperacion = "eliminar" Or strOperacion = "consultar") Then
      
        intRes = frmBase.spdGrid.GetText(intCol, frmBase.spdGrid.ActiveRow, varDato)
        spdEdit.SetText 2, lngEncontroFila, varDato
    
      End If
      
      'set puntero a fila actual
      spdEdit.row = lngEncontroFila
    
      'set lock y Backcolor cuando columna es No Permite Edit
      If InStr(LCase(frmBase.DataSoloLecturaEnEdit), "[" & LCase(varTitulo) & "]") <> 0 Then
        spdEdit.LockBackColor = RGB(240, 240, 240)
        spdEdit.Lock = True
      End If
    
      'se Backcolor cuando columna es obligatoria
      If InStr(LCase(frmBase.DataObligatorioEnEdit), "[" & LCase(varTitulo) & "]") <> 0 Then
        spdEdit.BackColor = RGB(225, 241, 255)
      End If
    
      'set lock si operacion es Consulta
      If strOperacion = "consultar" Then
        spdEdit.Lock = True
      End If
    
    End If
   
  Next
  
End Function

'
'ESTA FUNCION SE UTILIZA PARA CUANDO TENGO UNA GRILLA PERSONALIZADA, O SEA NO AUTOMATICA
'SE DEVUELVEN LOS VALORES SEPARADOS POR COMA EN UN STRING PASADO COMO ARGUMENTO PARA LUEGO
'SE LOS PUEDA DAR A UN STORE PROCEDURE Y ESTE ACTUALICE UNA TABLA
'
Function spdEditGet(ByVal spdEdit As fpSpread, ByRef str As String) As Boolean
  Dim intFila As Integer
  Dim varTitulo, varDato As Variant
  Dim lngEncontroFila As Long
                
  'recorro columnas
  For intFila = 1 To spdEdit.MaxRows
    
    'tomo nombre de columna de grilla base
    intRes = spdEdit.GetText(1, intFila, varTitulo)
        
    'set fila 2 datos
    spdEdit.GetText 2, intFila, varDato
        
    'set fila columna actual
    spdEdit.row = intFila
    spdEdit.col = 2
        
    If spdEdit.CellType <> CellTypeStaticText Then
        
      'pongo coma
      If str <> "" Then
        str = str & ","
      End If
        
      'armo string segun tipo de celda
      Select Case spdEdit.CellType
      
      Case CellTypeCheckBox, CellTypeNumber
        str = str & Val(varDato)
       
      Case CellTypeDate
        str = str & "'" & dateToIso(varDato) & "'"
      
      Case CellTypeEdit
        str = str & "'" & varDato & "'"
      
      Case CellTypeComboBox
        'si es un comboBox
         
        Dim intCantidadItem As Integer
        Dim varIndice As Variant
        Dim blnIndNumerico As Boolean
         
        'puntero columna 3 en donde se encuentra el indice asociado
        spdEdit.col = 3
            
        'chequeo si tengo que devolver dato numerico o texto
        'si es numerico es porque el comboBox tiene indice asociado
        blnIndNumerico = False
        For a = intCantidadItem To spdEdit.TypeComboBoxCount
          spdEdit.TypeComboBoxCurSel = intCantidadItem
          If spdEdit.CellType = CellTypeNumber Then
            If spdEdit.Text <> 0 Then
              blnIndNumerico = True
            End If
          End If
        Next
            
        'puntero en donde se encuentra el texto del combo
        spdEdit.col = 2
            
        'si se selecciono algun item
        If spdEdit.TypeComboBoxCurSel <> -1 Then
            
          'si hay que velolver numerico
          If blnIndNumerico Then
              
            Dim intIndice As Integer
              
            'puntero a columna 3 para asignarle la posicion
            'del item seleccionado al comboBox de la columna
            '3 que mantiene el identificador numerico del texto
            spdEdit.col = 2
            intIndice = spdEdit.TypeComboBoxCurSel
            spdEdit.col = 3
            spdEdit.TypeComboBoxCurSel = intIndice
              
            intRes = spdEdit.GetText(3, intFila, varDato)
            str = str & varDato
              
          'devuelvo texto
          Else
            str = str & "'" & varDato & "'"
          End If
            
        'no se selecciono ningun item
        Else
            
          'si hay que devolver numero devuelvo -1 sino blanco
          str = str & IIf(blnIndNumerico, "-1", "''")
            
        End If
      
      End Select
        
    End If
        
  Next
  
End Function


'
'PASA VALORES DE GRILLA VERTICAL (EDIT) A GRILLA HORIZONTAL (BASE)
'
Function spdEditToBase(ByVal frmBase As gridFRM, ByRef varNombreClave As Variant, varValorClave As Variant) As Boolean
  
  'dim variables cuenta columnas y toma titulo
  Dim rs As ADODB.Recordset
  Dim fld As ADODB.Field
  Dim varDato As Variant
                
  'busco registro
  strSQL = frmBase.dsi & " where " & varNombreClave & " = " & varValorClave
  Set rs = adoGetRS(strSQL)
  
  'chequeo que haya traido algo
  If Not rs.EOF Then
  
    'recorro nombres de columnas
    For Each fld In rs.Fields
        
      'busco segun nombreColumna en grilla base para saber ubicacion de columna
      'intRes = frmBase.spdGrid.SearchRow(0, 0, -1, fld.Name, SearchFlagsCaseSensitive)
      
      'reemplazando linea anterior porque cuando la grilla esta vacio no funciona
      'este codigo busca el nombre de la columna del recordset en la grilla
      Dim varNombreCol As Variant
      Dim intCol As Integer
      intRes = -1
      For intCol = 0 To frmBase.spdGrid.MaxCols
        frmBase.spdGrid.GetText intCol, 0, varNombreCol
        If fld.Name = varNombreCol Then
          intRes = intCol
          Exit For
        End If
      Next
      
      'si encontro celda reemplazo su valor
      If Not intRes = -1 Then
        
        'si el dato es BIT tengo que hacer esto porque la grilla no toma un true o false solo 0 o -1
        If fld.Type = conBit Then
          varDato = fld.Value * -1
        Else
          varDato = fld.Value
        End If
        
        frmBase.spdGrid.SetText intRes, frmBase.spdGrid.ActiveRow, varDato
      
      End If
    
    Next
    
  End If
  
End Function

'
'ESTA FUNCION SE UTILIZA PARA CUANDO TENGO UNA GRILLA PERSONALIZADA, O SEA NO AUTOMATICA
'SE LE PASA LA GRILLA HORIZONTAL BASE Y EL NOMBRE DE LA GRILLA PERSONALIZADA
'Y SE ENCARGA DE LIMPIAR LOS DATOS DE PANTALLA
'
Function spdEditClear(ByVal frmBase As Form, ByRef spdEdit As fpSpread, ByVal strOperacion As String) As Boolean
  Dim varTitulo, varDato As Variant
                
  'set columna 2 donde se encuentran los datos
  spdEdit.col = 2
                
  'recorro columnas
  For intRow = 1 To spdEdit.MaxRows
    
    'tomo titulo de columna de la grilla
    spdEdit.GetText 1, intRow, varTitulo
    
    'si no encuentro nombre de columna en la propiedad de noBlanquear, blanqueo
    If InStr(LCase(frmBase.DataNoBlanquear), "[" & LCase(varTitulo) & "]") = 0 Then
      
      spdEdit.col = 2
      spdEdit.row = intRow
      
      'si es numerico o bit fuerzo un cero
      If spdEdit.CellType = CellTypeCurrency Or spdEdit.CellType = CellTypeNumber Or spdEdit.CellType = CellTypeCheckBox Then
        spdEdit.SetText 2, intRow, 0
      Else
        spdEdit.SetText 2, intRow, ""
      End If
      
    End If
      
  Next
    
  'pongo foco en grilla
  spdEdit.SetFocus
  
  'me paro en la primera fila directamente para ingresar datos
  spdEdit.SetActiveCell 2, 1
  
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

'
'PASA VALORES DE GRILLA VERTICAL A GRILLA
'HORIZONTAL UNA VEZ EDITADOS LOS DATOS
'
Function spdEditSetToSpdBase(ByVal frmBase As Form, ByRef frmEdit As Form, ByVal strOperacion As String) As String
  
  'dim variables cuenta columnas y toma titulo
  Dim intRow As Integer
  Dim varTitulo, varDato As Variant
  Dim strResulParcial  As String
                
  'recorro filas para armar grilla
  strResulParcial = ""
  For intRow = 1 To frmEdit.spdEdit.MaxRows
        
    'tomo nombre de columna de grilla edit
    intRes = frmEdit.spdEdit.GetText(1, intRow, varTitulo)
        
    'busco columna no esta dentro de las no permiten edicion
    'If InStr(LCase(frmBase.DataSoloLecturaEnEdit), "[" & LCase(varTitulo) & "]") = 0 Then
        
      'tomo dato de columna de grilla edit
      intRes = frmEdit.spdEdit.GetText(2, intRow, varDato)
        
      'case tipo de dato para formatear y alinear
      Select Case frmBase.DataFields(varTitulo).Type
          
      'Bit, enteros, decimales
      Case conBit, conSmallInt, conInt, conTinyInt, conMoney, conSmallMoney, conReal, conFloat, conNumeric, conDecimal
          
        'si es un valor null al SQL le tengo que pasar un NULL
        strResulParcial = IIf(IsEmpty(varDato), "null", varDato)
          
      'fecha
      Case conSmallDateTime, conDateTime
      
        strResulParcial = "'" & dateToIso(varDato) & "'"
      
      'string
      Case conChar, conNchar, conVarchar, conText, conNVarchar, conNtext
      
        'puntero de grilla en fila y columna
        frmEdit.spdEdit.row = intRow
        frmEdit.spdEdit.col = 2
        
        'si es un comboBox
        If frmEdit.spdEdit.CellType = CellTypeComboBox Then
        
          Dim intCantidadItem As Integer
          Dim varIndice As Variant
          Dim blnIndNumerico As Boolean
        
          'puntero columna 3 en donde se encuentra el indice asociado
          frmEdit.spdEdit.col = 3
            
          'chequeo si tengo que devolver dato numerico o texto
          'si es numerico es porque el comboBox tiene indice asociado
          blnIndNumerico = False
          For a = intCantidadItem To frmEdit.spdEdit.TypeComboBoxCount
            frmEdit.spdEdit.TypeComboBoxCurSel = intCantidadItem
            If frmEdit.spdEdit.Text <> 0 Then
              blnIndNumerico = True
            End If
          Next
            
          'puntero en donde se encuentra el texto del combo
          frmEdit.spdEdit.col = 2
            
          'si se selecciono algun item
          If frmEdit.spdEdit.TypeComboBoxCurSel <> -1 Then
          
            'si hay que velolver numerico
            If blnIndNumerico Then
            
              Dim intIndice As Integer
              
              'puntero a columna 3 para asignarle la posicion
              'del item seleccionado al comboBox de la columna
              '3 que mantiene el identificador numerico del texto
              frmEdit.spdEdit.col = 2
              intIndice = frmEdit.spdEdit.TypeComboBoxCurSel
              frmEdit.spdEdit.col = 3
              frmEdit.spdEdit.TypeComboBoxCurSel = intIndice
            
              intRes = frmEdit.spdEdit.GetText(3, intRow, varDato)
              strResulParcial = varDato
            
            'devuelvo texto
            Else
              strResulParcial = "'" & varDato & "'"
            End If
            
          'no se selecciono ningun item
          Else
          
            'si hay que devolver numero devuelvo -1 sino blanco
            strResulParcial = IIf(blnIndNumerico, "-1", "''")
            
          End If
        
        'si es texto comun
        Else
          strResulParcial = "'" & varDato & "'"
        End If
        
      End Select
    
      'armando el sql final
      strResulFinal = strResulFinal & strResulParcial & ","
    
    'End If
    
  Next
  
  'set string final
  spdEditSetToSpdBase = "'" & strOperacion & "'," & Left(strResulFinal, Len(strResulFinal) - 1)

End Function

'
' VALIDA FILAS DE LA GRILLA TODAS SON OBLIGATORIAS EXCEPTO LAS QUE ESTAN EN LA
' COLUMNA NoObligatorio TRABAJA CON TABLA menuOpciones COLUMNA NOOBLIGATORIO
'
Function spdValidateData(ByVal frmBase As Form, ByVal spd As fpSpread) As Boolean
  
  spdValidateData = True
  
  'dim variables cuenta columnas y toma titulo
  Dim intFila As Integer
  Dim varTitulo, varDato As Variant
  Dim strLeyenda As String
                
  'recorro filas para armar grilla
  strLeyenda = ""
  For intFila = 1 To spd.MaxRows
        
    'get nombre de columna de grilla edit
    intRes = spd.GetText(1, intFila, varTitulo)
        
    'set fila columna
    spd.row = intFila
    spd.col = 2
    
    'si es static la paso por algo, en la mayoria de los casos
    'se va a dar en grillas perzonalizadas en filas titulo con Span
    If Not spd.CellType = CellTypeStaticText Then
        
      'busco columna, si es obligatoria
      If InStr(LCase(frmBase.DataObligatorioEnEdit), "[" & LCase(varTitulo) & "]") <> 0 Then
        
        'get dato de columna de grilla edit
        intRes = spd.GetText(2, intFila, varDato)
                
        'case tipo de celda para formatear y alinear
        Select Case frmBase.DataFields(varTitulo).Type
          
        'Bit, enteros, decimales
        Case conBit
      
        'nose valida porque por default es 0, false
      
        'enteros y decimales
        Case conSmallInt, conInt, conTinyInt, conMoney, conSmallMoney, conReal, conFloat, conNumeric, conDecimal
          
        ' nose valida porque por default es 0
          
        'fecha
        Case conSmallDateTime, conDateTime
      
          If varDato = "" Then
            strLeyenda = strLeyenda & varTitulo & vbCrLf
            spdValidateData = False
          End If
      
        'string
        Case conChar, conNchar, conVarchar, conText, conNVarchar, conNtext
        
          If varDato = "" Then
            strLeyenda = strLeyenda & varTitulo & vbCrLf
            spdValidateData = False
          End If
        
        End Select
    
      End If
      
    End If
    
  Next
  
  If strLeyenda <> "" Then
    intRes = MsgBox(strLeyenda & vbCrLf & "Información obligatoria.", vbApplicationModal + vbCritical + vbOKOnly, frmBase.Caption)
  End If
  
End Function

'
' DEVUELVE EL VALOR DE FILA ACTUAL Y NOMBRE DE COLUMNA PASADO COMO ARGUMENTO
'
Function spdGetValue(spdGrilla As fpSpread, strNombreColumna As String) As Variant
  Dim intPosicion As Integer
  Dim intColAnt As Integer
  
  'valor default
  spdGetValue = ""
  
  'busco posicion de la columna
  intPosicion = spdGrilla.SearchRow(0, 0, -1, strNombreColumna, 0)
  
  'si encontro devuelvo su valor
  If intPosicion <> -1 Then
    
    'guardo columna actual
    intColAnt = spdGrilla.col
    
    'puntero a columna encontrada
    spdGrilla.col = intPosicion
  
    'case tipo de celda para formatear y alinear
    Select Case spdGrilla.CellType
          
    'fecha
    Case CellTypeDate
      
       spdGetValue = "'" & dateToIso(spdGrilla.Text) & "'"
      
    'string
    Case CellTypeComboBox, CellTypeEdit
        
       spdGetValue = "'" & spdGrilla.Text & "'"
        
    'numero
    Case CellTypeCheckBox, CellTypeCurrency, CellTypeNumber
       
       spdGetValue = spdGrilla.Text
    
    End Select
  
    'recupero columna anterior
    spdGrilla.col = intColAnt
  
  End If
  
End Function

'
' GENERA ARCHIVO PARA IMPORTAR TRABAJA CON FORMULARIO gridFRM, importExportFRM Y
'  CON TABLA MENUOPCIONES CON LA DEFINICION DE CAMPOS OBLIGATORIOS Y AUTOMATICOS
'
Function spdGeneraArchivoParaImportar(gridFRM) As Boolean

  'verifico existencia de archivo
  If Dir(App.Path & "\Importar\" & gridFRM.Caption & "_estructura.xls") <> "" Then
    intRes = MsgBox("El archivo: " & App.Path & "\Importar\" & gridFRM.Caption & "_estructura.Xls" & Chr(13) & Chr(13) & "Ya existe. Sobreescribe ?", vbApplicationModal + vbQuestion + vbYesNo)
    If intRes = 7 Then 'click boton NO, cancela generacion estructura de importacion
      Exit Function
    End If
  End If
         
  'cambio puntero mouse
  Screen.MousePointer = vbHourglass
        
  'cargo auxiliar
  Load importExportFRM
    
  'determino cantidad de filas y columnas para grilla a generar
  importExportFRM.spdAuxiliar.MaxRows = 2
  importExportFRM.spdAuxiliar.MaxCols = 0
          
  'recorro grilla activa
  Dim intCuenta As Integer
  Dim varNombreColumna As Variant
  Dim strValor As String
  Dim strObligatoria As String
  Dim strAutomatico As String
  For intCol = 1 To gridFRM.spdGrid.MaxCols
        
    'tomo nombre de columna
    gridFRM.spdGrid.GetText intCol, 0, varNombreColumna
        
    'busco columna que no este definida como NoMuestraEnEdit
    If InStr(LCase(gridFRM.DataNoMuestraEnEdit), "[" & LCase(varNombreColumna) & "]") = 0 Then
                
      'agrego columna a la grilla a generar
      importExportFRM.spdAuxiliar.MaxCols = importExportFRM.spdAuxiliar.MaxCols + 1
                
      'puntero a fila y columna
      importExportFRM.spdAuxiliar.row = 1
      importExportFRM.spdAuxiliar.col = importExportFRM.spdAuxiliar.MaxCols
                
      'busco columna que este definida como obligatorio
      If InStr(LCase(gridFRM.DataObligatorioEnEdit), "[" & LCase(varNombreColumna) & "]") <> 0 Then
        strObligatoria = "OBLIGATORIO"
        importExportFRM.spdAuxiliar.BackColor = RGB(225, 241, 255)
        importExportFRM.spdAuxiliar.row = 2
        importExportFRM.spdAuxiliar.BackColor = RGB(225, 241, 255)
        importExportFRM.spdAuxiliar.row = 1
      Else
        strObligatoria = ""
      End If
        
      'busco columna que este definida como automatico
      If InStr(LCase(gridFRM.DataSoloLecturaEnEdit), "[" & LCase(varNombreColumna) & "]") <> 0 Then
        strAutomatico = "AUTOMATICO"
        importExportFRM.spdAuxiliar.BackColor = RGB(225, 225, 225)
        importExportFRM.spdAuxiliar.row = 2
        importExportFRM.spdAuxiliar.BackColor = RGB(225, 225, 225)
        importExportFRM.spdAuxiliar.row = 1
      Else
        strAutomatico = ""
      End If
                
      'consulto tipo de celda para formatear y alinear
      Select Case gridFRM.DataFields(varNombreColumna).Type
          
      'Bit
      Case conBit
        strValor = strValor & "(0/1)"
        
      'enteros
      Case conSmallInt, conInt, conTinyInt
        strValor = "(entero numerico)"
                    
      'decimal
      Case conMoney, conSmallMoney, conReal, conFloat, conNumeric, conDecimal
        strValor = "(decimal numerico)"
            
      'fecha
      Case conSmallDateTime, conDateTime
        strValor = "(fecha dd/mm/yyyy)"
        
      'string
      Case conChar, conNchar, conVarchar, conText, conNVarchar, conText
        strValor = "(texto) " & gridFRM.DataFields(varNombreColumna).DefinedSize
        
      End Select
        
      'le paso el valor a la grilla fila columna
      importExportFRM.spdAuxiliar.SetText intCol, 1, varNombreColumna
      importExportFRM.spdAuxiliar.FontBold = True
      importExportFRM.spdAuxiliar.SetText intCol, 2, strValor & " " & strObligatoria & " " & strAutomatico
        
    End If  'noMuestraEnEdit
      
  Next  'cuenta columnas
                    
  'saca proteccion para excel
  importExportFRM.spdAuxiliar.Protect = False
                  
  'genero el excel con estructura
  intRes = importExportFRM.spdAuxiliar.ExportToExcel(App.Path & "\Importar\" & gridFRM.Caption & "_estructura", "default", "")
          
  'cierro auxiliar
  Unload importExportFRM
          
  'recupero puntero mouse
  Screen.MousePointer = vbDefault
        
  If intRes = -1 Then 'ok
    intRes = MsgBox("Se generó archivo para importar para: " & gridFRM.Caption & ", en:" & Chr(13) & Chr(13) & App.Path & "\Import\" & gridFRM.Caption & "_estructura.Xls", vbApplicationModal + vbInformation + vbOKOnly)
  Else ' intRes = 0
    intRes = MsgBox("La estructura no se pudo generar." & Chr(13) & Chr(13) & "Verificar que el archivo no se encuentre abierto.", vbApplicationModal + vbCritical + vbOKOnly)
  End If

End Function
'
' FUNCION IMPORTA DATOS A UNA TABLA DESDE UN EXCEL GENERADO
' CON LA FUNCION DE GENERACION DE ARCHIVO PARA IMPORTAR
'
Function spdImportarDesdeExcel(gridFRM As gridFRM) As Boolean

  Dim X As Integer, listcount As Integer, handle As Integer
  Dim Y, Z As Boolean
  Dim rs As New ADODB.Recordset
  Dim cn As ADODB.Connection
  Dim strArrayDatosSeparados As Variant
  Dim rsID(5) As New ADODB.Recordset
  Dim CantRsID, intAux As Integer
  Dim intFila, intColumna, intFilasImportadas, intFilasNoImportadas, intIDBuscado As Long
  Dim blnRegistroOk As Boolean
  Dim strParcial, strTotal, strCampoRel, strNombreStore As String
  Dim strErrorImport, strCamposConErr As String
  Dim varNombreColumna, varTipoColumna As Variant
  
  If Dir(App.Path & "\Importar\" & gridFRM.Caption & "_Estructura.xls") = "" Then
    intRes = MsgBox("No se encontro el archivo: " & App.Path & "\Importar\" & gridFRM.Caption & "_estructura.Xls", vbApplicationModal + vbApplicationModal + vbExclamation + vbOKOnly)
    Exit Function
  End If
        
  'On Error GoTo Errores_SQL
  
  'armo linea de error de importacion, falta solo el nro de linea del excel
  strErrorImport = "'" & gridFRM.Caption & "_estructura.Xls" & "','" & dateToIso(Date) & "','" & Time & "',"
    
  'cargo form auxiliar
  Load importExportFRM
  
  'verifica que realmente sea un archivo de excel
  X = importExportFRM.spdAuxiliar.IsExcelFile(App.Path & "\Importar\" & gridFRM.Caption & "_estructura.Xls")
  If X <> 1 Then  'no es excel, o el arch esta abierto
    
    'cierro auxiliar
    Unload importExportFRM
    intRes = MsgBox("No se puede abrir archivo de Excel." & Chr(13) & Chr(13) & "Verificar que no se encuentre abierto.", vbApplicationModal + vbCritical + vbOKOnly)
    Exit Function
  
  End If
  
  'levanto archivo de excel a la grilla
  Y = importExportFRM.spdAuxiliar.GetExcelSheetList(App.Path & "\Importar\" & gridFRM.Caption & "_estructura.Xls", "", listcount, "", handle, True)
  Z = importExportFRM.spdAuxiliar.ImportExcelSheet(handle, 0)
  If Z = False Then
    Unload importExportFRM
    intRes = MsgBox("La importación no se pudo realizar.", vbApplicationModal + vbCritical + vbOKOnly)
    Exit Function
  End If
  
  'verifico que la grilla tenga mas de 2 filas
  ' la 1ra. fila es El nombre de columna y la 2da. es tipo de dato
  If importExportFRM.spdAuxiliar.DataRowCnt = 2 Then
    Unload importExportFRM
    intRes = MsgBox("La hoja de Excel sólo tiene las filas de nombres y tipos de Campo.", vbApplicationModal + vbCritical + vbOKOnly)
    Exit Function
  End If
  
  'chequeo que todas las columnas del excel importado
  'existan en la prodiedad fields del formulario actual
  
  'cambio puntero mouse
  Screen.MousePointer = vbHourglass
  
  'para saber si hay o no tablas relacionadas
  If MainMDI.ActiveForm.DataComboBox <> "" Then
    'separo columnas de select
    'posicion impar: ID, posicion par: select *....
    strArrayDatosSeparados = separateText(MainMDI.ActiveForm.DataComboBox, ";")
    
    'cargo los recordset con las tablas que tienen ID
    For intRes = 2 To UBound(strArrayDatosSeparados) Step 2
      CantRsID = CantRsID + 1
      strSQL = strArrayDatosSeparados(intRes)
      Set rsID(CantRsID) = adoGetRS(strSQL)
    Next intRes
  End If
  
  'abro conexion
  Set cn = adoOpenCn()
  
  'recorro la grilla a partir de la fila 3 que es donde comienzan los datos
  For intFila = 3 To importExportFRM.spdAuxiliar.DataRowCnt 'cant filas grilla
    
    strTotal = "": strParcial = ""
    strCamposConErr = "** "
    blnRegistroOk = True
    
    'recorro cantidad de columnas de la grilla importada
    For intColumna = 1 To importExportFRM.spdAuxiliar.DataColCnt
        
      'tomo nombre de columna para saber si existe en las columnas
      'del recordsety tambien para verificar el tipo de dato
      importExportFRM.spdAuxiliar.GetText intColumna, 1, varNombreColumna
      importExportFRM.spdAuxiliar.GetText intColumna, 2, varTipoColumna
        
      'puntero a fila y columna de la grilla
      importExportFRM.spdAuxiliar.row = intFila
      importExportFRM.spdAuxiliar.col = intColumna
        
      'verifico tipo de datos entre grilla y recordset
      Select Case gridFRM.DataFields(varNombreColumna).Type
          
        'Bit
        Case conBit
          
          If Trim(importExportFRM.spdAuxiliar.Value) = "" Then
            importExportFRM.spdAuxiliar.Value = 0
          End If
            
          If importExportFRM.spdAuxiliar.Value < 0 Or importExportFRM.spdAuxiliar.Value > 1 Then
            blnRegistroOk = False
          End If
            
          If blnRegistroOk Then
            strParcial = importExportFRM.spdAuxiliar.Value
          Else
            strCamposConErr = strCamposConErr & varNombreColumna & " **"
          End If
          
        'enteros
        Case conSmallInt, conInt, conTinyInt
          
          If Trim(importExportFRM.spdAuxiliar.Value) = "" Then
            importExportFRM.spdAuxiliar.Value = 0
          End If
          
          If Not IsNumeric(importExportFRM.spdAuxiliar.Value) Then
            blnRegistroOk = False
          End If
          
          If blnRegistroOk Then
            importExportFRM.spdAuxiliar.Value = Int(importExportFRM.spdAuxiliar.Value)
            strParcial = importExportFRM.spdAuxiliar.Value
          Else
            strCamposConErr = strCamposConErr & varNombreColumna & " **"
          End If

        'decimal
        Case conMoney, conSmallMoney, conReal, conFloat, conNumeric, conDecimal
          
          If importExportFRM.spdAuxiliar.Value = "" Then
            importExportFRM.spdAuxiliar.Value = 0
          End If
          
          If Not IsNumeric(importExportFRM.spdAuxiliar.Value) Then
            blnRegistroOk = False
          End If
          
          If blnRegistroOk Then
            strParcial = importExportFRM.spdAuxiliar.Value
          Else
            strCamposConErr = strCamposConErr & varNombreColumna & " **"
          End If
        
        'string
        Case conChar, conNchar, conVarchar, conText, conNVarchar, conText
          
          If blnRegistroOk Then
            strParcial = "'" & importExportFRM.spdAuxiliar.Text & "'"
          End If
        
        'fecha
        Case conSmallDateTime, conDateTime
        
          'si esta vacio cargo la fecha a "19000101"
          If Trim(importExportFRM.spdAuxiliar.Value) = "" Then
            importExportFRM.spdAuxiliar.Text = "01/01/1900"
          End If
          
          'verificamos un formato de fecha tradicional
          If Not IsDate(importExportFRM.spdAuxiliar.Text) Then
            blnRegistroOk = False
          End If
      
          If blnRegistroOk Then
            strParcial = "'" & dateToIso(importExportFRM.spdAuxiliar.Text) & "'"
          Else
            strCamposConErr = strCamposConErr & varNombreColumna & " **"
          End If

      End Select
          
            
      'si hay tablas relacionadas con la que estoy importando..
      If MainMDI.ActiveForm.DataComboBox <> "" Then
      
        'verifico si el nombre del campo del recorset rs es = strArrayDatosSeparados(NRO IMPAR)
        'que viene de menuopciones. De ser asi quiere decir que es un campo relacionado
        'de otra tabla, por lo cual en ese campo va un ID.
        'el dato que viene lo busco en el array de recordset que arme para obtener el ID
        intIDBuscado = 0
        
        'el for comienza en 1 ya que en los impares estan los nombres de los campos
        For intRes = 1 To UBound(strArrayDatosSeparados) Step 2
          
          strCampoRel = Replace(strArrayDatosSeparados(intRes), "[", "")
          strCampoRel = Replace(strCampoRel, "]", "")
          If LCase(gridFRM.DataFields(intColumna - 1).Name) = LCase(strCampoRel) Then
            
            'busco el dato en el array de recordset
            For intAux = 1 To CantRsID
                
              'me fijo en cual recordset busco
              'el array de recordset esta armado: ID, DATO
              If LCase(rsID(intAux)(1).Name) = LCase(strCampoRel) Then
                  
                If rsID(intAux).RecordCount > 0 Then
                  rsID(intAux).MoveFirst
                End If
                    
                  While Not rsID(intAux).EOF
                    If LCase(rsID(intAux)(1)) = LCase(importExportFRM.spdAuxiliar.Value) Then
                      intIDBuscado = rsID(intAux)(0)
                      Exit For
                    Else
                      rsID(intAux).MoveNext
                    End If
                  Wend
                  
              End If
              
            Next intAux
            
            If intIDBuscado > 0 Then
              strParcial = intIDBuscado
            Else
              blnRegistroOk = False
              strCamposConErr = strCamposConErr & varNombreColumna & " **"
            End If
          
          End If
        
        Next intRes
        
      End If
    
      If blnRegistroOk Then
        If strTotal = "" Then
          strTotal = strParcial
        Else
          strTotal = strTotal & "," & strParcial
        End If
      End If
      
    Next intColumna
    
    'el registro se puede insertar
    If blnRegistroOk Then
      blnBool = True
      cn.Execute gridFRM.DataSpINS & " 'insertar'," & strTotal
      
      If blnBool = True Then 'entonces la grabacion fue exitosa
        intFilasImportadas = intFilasImportadas + 1

      Else
        'grabo reg en la tabla ErroresDeImportacion, problema c/duplicacion
        intFilasNoImportadas = intFilasNoImportadas + 1
        cn.Execute "dsinfo_Errores_INS_sp " & strErrorImport & str$(intFila) & ",'Revisar, posible registro duplicado'"
      End If
    Else
      'grabo reg en la tabla ErroresDeImportacion, problema c/tipo de dato
      intFilasNoImportadas = intFilasNoImportadas + 1
      cn.Execute "dsinfo_Errores_INS_sp " & strErrorImport & str$(intFila) & ",'" & strCamposConErr & "'"
    End If
    
  Next intFila
  
  'cierro recordset's y form
  'si habia tablas relacionadas se cierran los recordset's
  If gridFRM.DataComboBox <> "" Then
    For intRes = 1 To CantRsID
      rsID(intRes).Close
    Next intRes
  End If
  'rs.Close
  
  'cierro conection
  cn.Close  'cierra conexion con la db
  
  Unload importExportFRM 'ACTIVAR LINEA ***************************
  
  'recupero puntero mouse
  Screen.MousePointer = vbDefault
  
  'Refresh grilla
  gridFRM.DataRefresh = True

  intRes = MsgBox(gridFRM.Caption & ":" & Chr(13) & Chr(13) & "(" & Val(intFilasImportadas) & ") Filas Importadas. " & Chr(13) & "(" & Val(intFilasNoImportadas) & ") Filas No Importadas.", vbApplicationModal + vbInformation + vbOKOnly)

End Function
'
'FUNCION EXPORTA DATOS DE UNA GRILLA A UN EXCEL
'-1:error, 0:no hay valores a exportar, 1:todo bien
'
Function spdExportarAExcel(gridFRM As gridFRM, strFile As String) As Integer

  If gridFRM.spdGrid.MaxRows = 0 Then
    spdExportarAExcel = 0
    Exit Function
  End If
        
  'mouse reloj
  If Dir(App.Path & "\iconos\_24x24_excel.ico") <> "" Then
    Screen.MouseIcon = LoadPicture(App.Path & "\iconos\_24x24_excel.ico")
    Screen.MousePointer = vbCustom
  Else
    Screen.MousePointer = vbHourglass
  End If
    
  'inserto 1 linea arriba de todo para ponerle los titulos de las columnas
  gridFRM.spdGrid.MaxRows = gridFRM.spdGrid.MaxRows + 1
  gridFRM.spdGrid.InsertRows 1, 1
  
  'cambio altura
  gridFRM.spdGrid.RowHeight(1) = 15
  
  For intRes = 1 To gridFRM.spdGrid.MaxCols
    gridFRM.spdGrid.GetText intRes, 0, titulo
    gridFRM.spdGrid.row = 1
    gridFRM.spdGrid.row2 = 1
    gridFRM.spdGrid.col = intRes
    gridFRM.spdGrid.col2 = intRes
    gridFRM.spdGrid.CellType = CellTypeEdit
    gridFRM.spdGrid.BackColor = RGB(205, 238, 254)
    gridFRM.spdGrid.SetText intRes, 1, titulo
  Next
  
  'saca proteccion para excel
  gridFRM.spdGrid.Protect = False
  
  'exporto
  intRes = gridFRM.spdGrid.ExportToExcel(strFile, "", "")
 
  'elimino fila 1 que servia para exportar los nombres de columnas
  gridFRM.spdGrid.DeleteRows 1, 1
  gridFRM.spdGrid.MaxRows = gridFRM.spdGrid.MaxRows - 1
    
  'mouse default
  Screen.MousePointer = vbDefault
    
  'exporto OK
  If intRes = -1 Then
    spdExportarAExcel = 1
  Else
    spdExportarAExcel = -1
  End If

End Function


