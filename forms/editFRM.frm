VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpspr60.ocx"
Begin VB.Form editFRM 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "form1"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   3855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   330
      Left            =   2130
      TabIndex        =   1
      Top             =   2055
      Width           =   1050
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   870
      TabIndex        =   0
      Top             =   2055
      Width           =   1095
   End
   Begin FPSpreadADO.fpSpread spdE 
      Height          =   1635
      Left            =   150
      TabIndex        =   2
      Top             =   300
      Width           =   3030
      _Version        =   393216
      _ExtentX        =   5345
      _ExtentY        =   2884
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
      SpreadDesigner  =   "editFRM.frx":0000
   End
End
Attribute VB_Name = "editFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngAnchoInicial, lngAlturaInicial As Long
Dim m_gridFrm As gridFRM


Public Property Set dsiGridFrm(ByVal frm As gridFRM)
  Set m_gridFrm = frm
End Property

Public Property Get dsiGridFrm() As gridFRM
  Set dsiGridFrm = m_gridFrm
End Property

Public Property Get dsiPantallaDoble() As Boolean
  dsiPantallaDoble = Me.dsiGridFrm.dsiEsquema.Count * 225 > Screen.Height * 0.6
End Property

Public Property Get dsiDatosDibujar() As Boolean
  
  Dim intI, intCol As Integer
  
  Dim a As fpSpread
  Dim varDefault As Variant
  Dim intDecimales As Variant
  Dim intCantDecimales As Integer
   
  'separo en un array los valores por default
  'ejemplo: [telefono],4361-0000,[pais],Argentina
  'varDefault = separateText(frmBase.DataValorDefault)
  
  'cambio caption del form edit
  'frmEdit.Caption = frmBase.Caption & " - " & UCase(strOperacion)
      
  'set cantidad filas, segun si es necesario doble pantalla 2 no
  If Me.dsiPantallaDoble() Then
    Me.spdE.MaxRows = Round(Me.dsiGridFrm.dsiEsquema.Count / 2)
  Else
    Me.spdE.MaxRows = Me.dsiGridFrm.dsiEsquema.Count
  End If
  
  'set cantidad columnas
  Me.spdE.MaxCols = 6
    
  'set ancho de 1 columna
  Me.spdE.ColWidth(1) = 2000
  Me.spdE.ColWidth(2) = 2000
  Me.spdE.ColWidth(3) = 0
  Me.spdE.ColWidth(4) = 0
  Me.spdE.ColWidth(5) = 0
  Me.spdE.ColWidth(6) = 0
      
  'calculando si la pantalla se divide en 2, porque son muchas columnas
  If Me.dsiPantallaDoble() Then
    
    'set ancho columnas
    Me.spdE.ColWidth(4) = 2000
    Me.spdE.ColWidth(5) = 2000
    Me.spdE.ColWidth(6) = 0
    
  End If
      
  'dim variables cuenta columnas y toma titulo
  Dim intUltimaFilaAgregada As Integer
  Dim sngAlturaTotalFilas As Single
  Dim varTitulo, varDato As Variant
                
  intUltimaFilaAgregada = 0
  sngAlturaTotalFilas = 0
  
  'si la pantalla se divide en 2
  If Me.dsiPantallaDoble() Then
  
    'set titulos a doble pantalla
    For intI = 0 To Format(Me.dsiGridFrm.dsiEsquema.Count / 2, "##") - 1
      Me.spdE.SetText 1, intI + 1, Me.dsiGridFrm.dsiEsquema(intI).Name
      Me.spdE.SetText 4, intI + 1, Me.dsiGridFrm.dsiEsquema(intI + Format(Me.dsiGridFrm.dsiEsquema.Count / 2, "##")).Name
    Next
    
  Else
    
    'set titulos a una sola pantalla
    For intI = 0 To Me.dsiGridFrm.dsiEsquema.Count - 1
      Me.spdE.SetText 1, intI + 1, Me.dsiGridFrm.dsiEsquema(intI).Name
    Next
  
  End If
      
  'el ancho de columnas lo establece automaticamente a la columna mas ancha
  'si la columna mas ancho es menor que 2000, set default 2000
  Dim dblAnchoMaximo As Long
  dblAnchoMaximo = Me.spdE.MaxTextColWidth(1)
  
  If dblAnchoMaximo < 2000 Then
    dblAnchoMaximo = 2000
  End If
  
  Me.spdE.ColWidth(1) = dblAnchoMaximo
  Me.spdE.ColWidth(2) = dblAnchoMaximo
  
  If Me.dsiPantallaDoble() Then
    Me.spdE.ColWidth(4) = dblAnchoMaximo
    Me.spdE.ColWidth(5) = dblAnchoMaximo
  End If
              
  'set posicion de grilla
  Me.spdE.Left = 50
  Me.spdE.Top = 50
              
  'altura de grilla
  tot = 0
  For z = 0 To Me.spdE.MaxRows
    tot = tot + Me.spdE.RowHeight(z) + (Me.spdE.RowHeight(z) * 0.09)
  Next
  
  Me.spdE.Height = tot
  
  'ancho de grilla
  tot = 0
  For z = 0 To Me.spdE.MaxCols
    tot = tot + Me.spdE.ColWidth(z) + (Me.spdE.RowHeight(z) * 0.1)
  Next
  
  Me.spdE.Width = tot
          
  'set altura, anchura formulario
  Me.Height = Me.spdE.Top + Me.spdE.Height + 500
  Me.Width = Me.spdE.Width + 500
    
  'set tamaño a botones aceptar
  Me.cmdAceptar.Width = tot / 2 + 30
  Me.cmdAceptar.Height = 300
  
  'cambio tamaño a botones cancelar
  Me.cmdCancelar.Width = tot / 2 + 30
  Me.cmdCancelar.Height = 300
      
  'set ubicacion de botones aceptar
  Me.cmdAceptar.Left = 100
  Me.cmdAceptar.Top = Me.spdE.Height + 200
    
  'set ubicacion de botones cancelar
  Me.cmdCancelar.Left = Me.cmdAceptar.Left + Me.cmdAceptar.Width + 30
  Me.cmdCancelar.Top = Me.spdE.Height + 200
  
End Property

Public Property Get dsiDatosSet() As Boolean
  
    'tomo cantidad de decimales de columna actual si es de tipo numerica
'    frmBase.spdGrid.row = frmBase.spdGrid.ActiveRow
'    frmBase.spdGrid.col = intCol
    
    'tomo decimales columna actual
'    If frmBase.spdGrid.CellType = CellTypeNumber Then
'      intDecimales = frmBase.spdGrid.TypeNumberDecPlaces
'    End If
  
    'busco columna que no este definida como no se muestran en edit
'    If InStr(LCase(frmBase.DataNoMuestraEnEdit), "[" & LCase(varTitulo) & "]") = 0 Then
        
'      intUltimaFilaAgregada = intUltimaFilaAgregada + 1
        
      'inserto fila
'      frmEdit.spdE.InsertRows intUltimaFilaAgregada, 1
          
      'seteo maximo fila
'      frmEdit.spdE.MaxRows = intUltimaFilaAgregada
        
      'set nombre de columna en grilla edit
 '     frmEdit.spdE.SetText 1, intUltimaFilaAgregada, varTitulo
       
      'set dato en grilla edit
 '     frmEdit.spdE.SetText 2, intUltimaFilaAgregada, varDato
        
      'set puntero en fila y columna actual para cambiar propiedades
 '     frmEdit.spdE.col = 1
 '     frmEdit.spdE.row = intUltimaFilaAgregada
        
      'set altura de fila porcentaje de la altura maxima de screen
 '     frmEdit.spdE.RowHeight(intUltimaFilaAgregada) = (Screen.Height * 0.0012)
        
      'set alineacion vertical
 '     frmEdit.spdE.TypeVAlign = TypeVAlignCenter

 '     frmEdit.spdE.RowHeight(intUltimaFilaAgregada) = 12
      
      'set color fondo gris para titulos
 '     frmEdit.spdE.BackColor = RGB(240, 240, 240)
          
      'set color texto gris para titulos
 '     frmEdit.spdE.ForeColor = RGB(131, 131, 131)
      
      'tamaño de letra para titulos
 '     frmEdit.spdE.FontSize = 10
      
      'set titulos estaticos
 '     frmEdit.spdE.CellType = CellTypeStaticText
      
      'set columna 2 donde se encuentran los datos a actualizar
 '     frmEdit.spdE.col = 2
          
      'set color de fondo para ingreso de datos
 '     frmEdit.spdE.BackColor = RGB(255, 255, 255)
          
      'set color de texto para ingreso de datos
 '     frmEdit.spdE.ForeColor = RGB(0, 0, 0)
          
      'tamaño de letra para ingreso de datos
 '     frmEdit.spdE.FontSize = 9
          
      'case tipo de celda para formatear y alinear
 '     Select Case frmBase.DataFields(varTitulo).Type
          
      'Bit
 '     Case conBit
        
 '       frmEdit.spdE.CellType = CellTypeCheckBox
 '       frmEdit.spdE.Value = 0
          
      'enteros
 '     Case conSmallInt, conInt, conTinyInt
        
 '       frmEdit.spdE.CellType = CellTypeNumber
 '       frmEdit.spdE.TypeNumberDecPlaces = 0
 '       frmEdit.spdE.TypeHAlign = TypeHAlignRight
 '       frmEdit.spdE.Value = 0
                    
      'decimal
 '     Case conMoney, conSmallMoney, conReal, conFloat, conNumeric, conDecimal
            
 '       frmEdit.spdE.CellType = CellTypeNumber
 '       frmEdit.spdE.TypeHAlign = TypeHAlignRight
 '       frmEdit.spdE.Value = 0
 '
        'set decimales
 '       frmEdit.spdE.TypeNumberDecPlaces = intDecimales
                    
      'fecha
 '     Case conSmallDateTime, conDateTime
      
 '       frmEdit.spdE.CellType = CellTypeDate
 '       frmEdit.spdE.Text = "06/17/00"
                                        
      'string
 '     Case conChar, conNchar, conVarchar, conText, conNVarchar, conText
        
        'set ComboBox si es una columna que se muestra en un comboBox
        'tambien armo una fila de tipo comboBox en la columan 3 para
        'guardar los index de cada combo
 '       If InStr(LCase(frmBase.DataComboBox), "[" & LCase(varTitulo) & "]") <> 0 Then
 '         frmEdit.spdE.CellType = CellTypeComboBox
 '         frmEdit.spdE.TypeComboBoxEditable = False
 '         frmEdit.spdE.col = 3
 '         frmEdit.spdE.CellType = CellTypeComboBox
 '         frmEdit.spdE.TypeComboBoxEditable = False
 '         frmEdit.spdE.col = 2
          'lleno combo con datos
 '         intRes = spdDataToCbo(frmEdit.spdE, varTitulo, frmBase.DataComboBox)
 '       Else
          'sino se muestra solo un texto
 '         frmEdit.spdE.CellType = CellTypeEdit
          
          'set limite 500 caracteres
 '         frmEdit.spdE.TypeMaxEditLen = 500
          
          'set columna con .... cuando no se ve todo el texto
 '         frmEdit.spdE.TypeEllipses = True
          
 '       End If
                                       
 '     End Select
         
 '     'set valores por default, pueden ser fijos o pueden venir en un select
 '     If InStr(LCase(frmBase.DataValorDefault), "[" & LCase(varTitulo) & "]") <> 0 Then
 '       frmEdit.spdE.Text = arrayGetValue(varDefault, "[" & LCase(varTitulo) & "]")
 '     End If
          
      'si operacion es un U de update o D de delete busco el valor de la
      'columna en grilla horizontal luego lo guardo en grilla vertical
 '     If strOperacion = "editar" Or strOperacion = "eliminar" Or strOperacion = "consultar" Then
 '       intRes = frmBase.spdGrid.GetText(intCol, frmBase.spdGrid.ActiveRow, varDato)
 '       intRes = frmEdit.spdE.SetText(2, intCol, varDato)
 '     End If
         
      'set color texto negro columna 2
 '     frmEdit.spdE.ForeColor = RGB(0, 0, 0)
          
      'set alineacion vertical columna 2
 '     frmEdit.spdE.TypeVAlign = TypeVAlignCenter
          
      'set lock y Backcolor cuando columna es No Permite Edit
 '     If InStr(LCase(frmBase.DataSoloLecturaEnEdit), "[" & LCase(varTitulo) & "]") <> 0 Then
 '       frmEdit.spdE.LockBackColor = RGB(220, 220, 220)
 '       frmEdit.spdE.Lock = True
 '     End If
    
      'se Backcolor cuando columna es obligatoria
 '     If InStr(LCase(frmBase.DataObligatorioEnEdit), "[" & LCase(varTitulo) & "]") <> 0 Then
 '       frmEdit.spdE.BackColor = RGB(204, 230, 255)
 '     End If
    
      'set lock si operacion es Consulta
 '     If strOperacion = "consultar" Then
 '       frmEdit.spdE.Lock = True
 '     End If
    
      'sumo altura de fila actual para determinar altura total de la grilla
'      sngAlturaTotalFilas = sngAlturaTotalFilas + frmEdit.spdE.RowHeight(frmEdit.spdE.MaxRows) + 0.28
    
 '   End If
    
 ' Next
      
  'si es formulario ajustable
'  If blnAjustable Then
    
    'ancho columna 1 con titulos ajusta al mas ancho automaticamente
    'si el campo mas ancho es menor que 25 por default es 25
'    Dim dblAnchoMaximo As Double
'    dblAnchoMaximo = frmEdit.spdE.MaxTextColWidth(1)
  
'    If dblAnchoMaximo < 25 Then
'      dblAnchoMaximo = 25
'    End If
  
'    frmEdit.spdE.ColWidth(1) = dblAnchoMaximo
      
    'ancho columna 2 en donde se ingresan los datos igual
    'al ancho de titulos para que quede una grilla pareja
'    frmEdit.spdE.ColWidth(2) = dblAnchoMaximo
      
    'ancho columna 3 para indice de combo oculta
'    frmEdit.spdE.col = 3
'    frmEdit.spdE.ColHidden = True
          
    'ancho de grilla dinamico es la suma del ancho de la
    'columna 1 + 2 pero primero debo convertir a twips
'    Dim lngAnchoGrilla As Long
'    frmEdit.spdE.ColWidthToTwips (frmEdit.spdE.ColWidth(1) + frmEdit.spdE.ColWidth(2)), lngAnchoGrilla
    'frmEdit.spdE.Width = lngAnchoGrilla + 100
'    frmEdit.spdE.Width = lngAnchoGrilla + 50
        
    'alto de grilla dinamico es la suma de la altura de
    'todas las filas, pero primero debo convertir a twips
    'le agrego un 6 por ciento porque si no funciona mal
'    Dim lngAltoGrilla As Long
'    frmEdit.spdE.RowHeightToTwips 1, sngAlturaTotalFilas, lngAltoGrilla
'    frmEdit.spdE.Height = lngAltoGrilla + lngAltoGrilla * 0.06
    
    'ancho formulario
'    frmEdit.Width = frmEdit.spdE.Width + 300
    
    'alto del formulario
'    frmEdit.Height = frmEdit.spdE.Height + 1000
    
    'ubico grilla en form
'    frmEdit.spdE.Left = 100
'    frmEdit.spdE.Top = 100
    
    'cambio tamaño a botones aceptar
'    frmEdit.cmdAceptar.Width = lngAnchoGrilla / 2 + 30
'    frmEdit.cmdAceptar.Height = 300
  
    'cambio tamaño a botones cancelar
'    frmEdit.cmdCancelar.Width = lngAnchoGrilla / 2 + 30
'    frmEdit.cmdCancelar.Height = 300
      
    'cambio ubicacion de botones aceptar
'    frmEdit.cmdAceptar.Left = 100
'    frmEdit.cmdAceptar.Top = frmEdit.spdE.Height + 200
    
    'cambio ubicacion de botones cancelar
'    frmEdit.cmdCancelar.Left = frmEdit.cmdAceptar.Left + frmEdit.cmdAceptar.Width + 30
'    frmEdit.cmdCancelar.Top = frmEdit.spdE.Height + 200
  
'  End If
  
  'selda activa fila 1 columna 2
'  frmEdit.spdE.SetActiveCell 2, 1
  
  'muestra form
'  frmEdit.Show vbModal
  
End Property


'Private Sub cmdAceptar_Click()

'  If Not spdValidateData(MainMDI.ActiveForm, Me.spdE) Then Exit Sub

'  'confirmacion
'  USOgen.LngN = MsgBox("Esta seguro que desea confirmar.", vbQuestion + vbYesNo, "confirmación...")
  
'  If USOgen.LngN = vbYes Then
    
'    'determino operacion para saber cual es el SP que debe procesar SQL
'    Dim strSP As String
    
'    Select Case LCase(Me.Operacion)
'    Case "insertar"
'      strSP = MainMDI.ActiveForm.DataSpINS
    
'    Case "editar"
'      strSP = MainMDI.ActiveForm.DataSpEDI
    
'    Case "eliminar"
'      strSP = MainMDI.ActiveForm.DataSpELI
    
'    End Select
        
    'procesa operacion en SQL
'    USOgen.strT = spdESetToSpdBase(MainMDI.ActiveForm, Me, Me.Operacion)
'    USOgen.strT = "exec " & strSP & " " & USOgen.strT
'    SQLexec (USOgen.strT)
    
'    'si grabo OK
'    If SQLparam.CnErrNumero = -1 Then
      
'      'declaro
'      Dim varNombreClave, varValorClave As Variant
'      Dim rs As ADODB.Recordset
      
'      Select Case Me.Operacion
      
'      Case "insertar"
              
'        'tomo nombre de columna ID
'        USOgen.LngN = Me.spdE.GetText(1, 1, varNombreClave)
        
'        'tomo valor de columna ID
'        USOgen.strT = Replace(MainMDI.ActiveForm.DataSource, "*", "max(" & varNombreClave & ") as clave")
'        Set rs = SQLexec(USOgen.strT)
             
'        If Not rs.EOF Then
          
'          Select Case rs("clave").Type
          
'          Case conChar, conNchar, conVarchar, conText, conNVarchar
'            varValorClave = "'" & rs!clave & "'"
          
'          Case conSmallInt, conInt, conTinyInt
'            varValorClave = rs!clave
            
'          End Select
        
'        End If
                
'        'set le digo a grilla que va a haber una fila mas
'        MainMDI.ActiveForm.spdGrid.MaxRows = MainMDI.ActiveForm.spdGrid.MaxRows + 1
        
'        'inserto fila, si es la primera, agrego en fila 1 y sino agrego en fila activa
'        If MainMDI.ActiveForm.spdGrid.MaxRows = 1 Then
'          MainMDI.ActiveForm.spdGrid.InsertRows 1, 1
'        Else
'          MainMDI.ActiveForm.spdGrid.InsertRows MainMDI.ActiveForm.spdGrid.ActiveRow, 1
'        End If
        
'        'bloqueo fila insertada
'        MainMDI.ActiveForm.spdGrid.row = MainMDI.ActiveForm.spdGrid.ActiveRow
'        MainMDI.ActiveForm.spdGrid.row2 = MainMDI.ActiveForm.spdGrid.ActiveRow
'        MainMDI.ActiveForm.spdGrid.col = 1
'        MainMDI.ActiveForm.spdGrid.col2 = -1
'        MainMDI.ActiveForm.spdGrid.Lock = True
        
'        'lleno fila en blanco con la info que agregue
'        USOgen.LngN = spdEToBase(MainMDI.ActiveForm, varNombreClave, varValorClave)
        
'        'si hay mas de una fila aplico a fila activa color default
'        If MainMDI.ActiveForm.spdGrid.MaxRows <> 1 Then
        
'         'set color default de grilla en celda para abandonar
'          MainMDI.ActiveForm.spdGrid.row = MainMDI.ActiveForm.spdGrid.ActiveRow + 1
'          MainMDI.ActiveForm.spdGrid.row2 = MainMDI.ActiveForm.spdGrid.ActiveRow + 1
'          MainMDI.ActiveForm.spdGrid.col = -1
'          MainMDI.ActiveForm.spdGrid.BackColor = RGB(245, 245, 245)
'          MainMDI.ActiveForm.spdGrid.ForeColor = RGB(60, 60, 60)
          
'        End If
        
'        'set color de seleccion a celda nueva
'        MainMDI.ActiveForm.spdGrid.row = MainMDI.ActiveForm.spdGrid.ActiveRow
'        MainMDI.ActiveForm.spdGrid.row2 = MainMDI.ActiveForm.spdGrid.ActiveRow
'        MainMDI.ActiveForm.spdGrid.col = -1
'        MainMDI.ActiveForm.spdGrid.BackColor = RGB(220, 220, 220)
'        MainMDI.ActiveForm.spdGrid.ForeColor = RGB(60, 60, 60)
        
'        'limpia datos encabezado grilla Edit
'        USOgen.LngN = spdEClear(MainMDI.ActiveForm, Me.spdE, Me.Operacion)
        
'      Case "editar"
        
'        'tomo nombre y valor de columna ID
'        USOgen.LngN = Me.spdE.GetText(1, 1, varNombreClave)
'        USOgen.LngN = Me.spdE.GetText(2, 1, varValorClave)
        
'        'paso valores de grilla edit a grilla Base en fila actual
'        USOgen.LngN = spdEToBase(MainMDI.ActiveForm, varNombreClave, varValorClave)
        
'      Case "eliminar"
              
'        'elimino fila actual de grilla base, por ahora tengo que agregar una fila y borrar 2 sino no funciona
'        MainMDI.ActiveForm.spdGrid.MaxRows = MainMDI.ActiveForm.spdGrid.MaxRows + 1
'        MainMDI.ActiveForm.spdGrid.InsertRows MainMDI.ActiveForm.spdGrid.ActiveRow, 1
'        MainMDI.ActiveForm.spdGrid.DeleteRows MainMDI.ActiveForm.spdGrid.ActiveRow, 2
'        MainMDI.ActiveForm.spdGrid.MaxRows = MainMDI.ActiveForm.spdGrid.MaxRows - 2
        
'        'set color de seleccion a celda nueva
'        MainMDI.ActiveForm.spdGrid.row = MainMDI.ActiveForm.spdGrid.ActiveRow
'        MainMDI.ActiveForm.spdGrid.row2 = MainMDI.ActiveForm.spdGrid.ActiveRow
'        MainMDI.ActiveForm.spdGrid.col = -1
'        MainMDI.ActiveForm.spdGrid.BackColor = RGB(220, 220, 220)
'        MainMDI.ActiveForm.spdGrid.ForeColor = RGB(60, 60, 60)
        
'      End Select
      
'      'conexion
'      SQLclose
      
'      'aviso de operacion OK
'      USOgen.LngN = MsgBox("La operación de " & Me.Operacion & " se realizó con éxito.", vbInformation + vbOKOnly, "atención...")
        
'      USOgen.blnYES = True
'      USOgen.blnNO = False
      
'      'fuerzo cerrar form
'      If Me.Operacion = "editar" Or Me.Operacion = "eliminar" Then
'        cmdCancelar_Click
'      End If
      
'    Else
'      SQLError
'    End If ' operacion exito
    
'  End If 'confirmo
  

'End Sub

Private Sub cmdCancelar_Click()
  
'  USOgen.blnYES = False
'  USOgen.blnNO = True
  Unload Me

End Sub

Private Sub Form_Activate()
  
  'guardo ancho y alto inicial de form para no permitir achicar menos de eso
  lngAnchoInicial = Me.Width
  lngAlturaInicial = Me.Height
  
End Sub

Private Sub Form_Load()
  
  Dim blnB As Boolean
  
  'referencia a gridFRM
  Set Me.dsiGridFrm = MainMDI.ActiveForm
  
  'grilla en twips
  Me.spdE.UnitType = UnitTypeTwips
  
  'fondo del area seleccionada amarillo
  Me.spdE.SelBackColor = RGB(204, 230, 255)
    
  'set apariencia de borde 3d
  Me.spdE.Appearance = Appearance3D
    
  'barra de desplazamiento vertical solo cuando la necesita
  Me.spdE.ScrollBars = ScrollBarsNone
      
  'set muestro cuadriculado horizontal y vertical y color de fondo debajo de grilla
  Me.spdE.BackColorStyle = BackColorStyleUnderGrid
  Me.spdE.GridShowHoriz = True
  Me.spdE.GridShowVert = True
          
  'set cuando enter voy fila siguiente
  Me.spdE.EditEnterAction = EditEnterActionDown
          
  'set cuando ingreso valor arriba de otro lo reemplaza
  Me.spdE.EditModeReplace = True
          
  'set lineas de fondo
  Me.spdE.GridColor = RGB(192, 192, 192)
      
  'set area fuera de la grilla blanco
  Me.spdE.GrayAreaBackColor = RGB(255, 255, 255)
      
  'estilo borde grilla no
  Me.spdE.BorderStyle = BorderStyleFixedSingle
      
  Me.spdE.NoBorder = True
      
  'setea para mostrar tooltip en las celdas donde no se ve toda la info
  Me.spdE.TextTip = TextTipFixed
  Me.spdE.TextTipDelay = 250
    
  'inicializo valores de ancho y alto default para el form
  lngAnchoInicial = 0
  lngAlturaInicial = 0
    
  blnB = Me.dsiDatosDibujar()
    
End Sub

Private Sub Form_Resize()
  
'  Dim sngAnchoColumna As Single
'  Dim lngAnchoColumna As Long
  
  'no permite que en ancho sea menor al default
'  If Me.Width < lngAnchoInicial Then
'    Me.Width = lngAnchoInicial
'    Exit Sub
'  End If
  
  'no permite cambiar la altura del form
'  If Me.Height <> lngAlturaInicial And lngAlturaInicial <> 0 Then
'    Me.Height = lngAlturaInicial
'  End If
  
  'ancho grilla igual a ancho de formulario - 300
'  Me.spdE.Width = Me.Width - 300
  
  'convierto ancho grilla a twips
'  Me.spdE.TwipsToColWidth Me.spdE.Width, sngAnchoColumna
  
  'ancho de columna 2 igual a ancho de grilla menos columna 1
'  Me.spdE.ColWidth(2) = sngAnchoColumna - Me.spdE.ColWidth(1) - 0.6
  
  'convierto ancho de columna 2 a width
'  Me.spdE.ColWidthToTwips Me.spdE.ColWidth(2), lngAnchoColumna
  
  'ajusto boton cancelar
'  Me.cmdCancelar.Width = lngAnchoColumna
  
End Sub


