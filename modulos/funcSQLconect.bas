Attribute VB_Name = "funcSQLconect"

'parametros para conexion SQL
Type SQLparam
  nombreINI As String
  nombreDAT As String
  IDmenu As String
  Proveedor As String
  ServerSeguridad As String
  ServerDatos As String
  BaseDEdatos As String
  SeguridadIntegrada As String
  TiempoEspera As String
  Usuario As String
  UsuarioClave As String
  UsuarioConectado As String
  GrupoConectado  As Variant
  CantidadFilas As String
  ReportesPath As String
  Role As String
  RoleClave As String
  cn As New ADODB.Connection
  CnErrNumero As Long
  CnErrTexto As String
  Tag As String
End Type

Global SQLparam As SQLparam

'constantes para tipos de dato SQL
Global Const conChar = 129
Global Const conNchar = 130
Global Const conVarchar = 200
Global Const conText = 201
Global Const conNVarchar = 202
Global Const conNtext = 203
Global Const conDateTime = 135
Global Const conSmallDateTime = 135
Global Const conSmallInt = 2
Global Const conInt = 3
Global Const conTinyInt = 17
Global Const conReal = 4
Global Const conFloat = 5
Global Const conMoney = 6
Global Const conSmallMoney = 6
Global Const conNumeric = 131
Global Const conDecimal = 131
Global Const conBit = 11

'leo parametros de conexion de un INI
Function SQLgetParam() As Boolean
    
  Dim strT, strINInombre As String
  Dim intI As Integer
    
  'default return true
  SQLgetParam = True
  
  'inicializo
  strT = ""
  
  'get UserName: 26/09/2006
  strINInombre = NTuserName()
  
  'build nombre de INI y DAT: 26/09/2006
  SQLparam.nombreINI = App.Path & "\" & strINInombre & ".ini"
  SQLparam.nombreDAT = App.Path & "\" & strINInombre & ".dat"
  
  SQLparam.IDmenu = ReadIni("conexion", "idMenu", SQLparam.nombreINI)
  SQLparam.Proveedor = ReadIni("conexion", "Proveedor", SQLparam.nombreINI)
  SQLparam.ServerSeguridad = ReadIni("conexion", "ServerSeguridad", SQLparam.nombreINI)
  SQLparam.ServerDatos = ReadIni("conexion", "ServerDatos", SQLparam.nombreINI)
  SQLparam.BaseDEdatos = ReadIni("conexion", "BaseDEdatos", SQLparam.nombreINI)
  SQLparam.SeguridadIntegrada = ReadIni("conexion", "SeguridadIntegrada", SQLparam.nombreINI)
  SQLparam.TiempoEspera = ReadIni("conexion", "TiempoEspera", SQLparam.nombreINI)
  SQLparam.Usuario = ReadIni("conexion", "Usuario", SQLparam.nombreINI)
  SQLparam.UsuarioClave = ReadIni("conexion", "UsuarioClave", SQLparam.nombreINI)
  SQLparam.CantidadFilas = ReadIni("conexion", "CantidadFilas", SQLparam.nombreINI)
  SQLparam.ReportesPath = ReadIni("conexion", "ReportesPath", SQLparam.nombreINI)
    
  'check si encuentro datos de conexion, get usuario y grupo
  If SQLparam.Proveedor <> "" Or SQLparam.ServerSeguridad <> "" Or SQLparam.ServerDatos <> "" Or SQLparam.BaseDEdatos <> "" Or SQLparam.SeguridadIntegrada <> "" Then
    
    'get usuario conectado
    SQLparam.UsuarioConectado = SQLgetUsuario()
    
    'get grupos al que pertenece el usuario conectado
    SQLparam.GrupoConectado = SQLgetGrupo()
    
    'build lista de grupos al cual pertenece el usuario conectado
    If Not IsEmpty(SQLparam.GrupoConectado) Then
      For intI = 1 To UBound(SQLparam.GrupoConectado)
        strT = strT & "'" & SQLparam.GrupoConectado(intI) & "',"
      Next
    End If
    
    'check si encontro algo, delete ultima coma
    If Len(strT) > 0 Then
      strT = Left(strT, Len(strT) - 1)
    End If
    
    'save lista de grupos separadas por coma
    If strT = "" Then
      SQLparam.Tag = "'" & strT & "'"
    Else
      SQLparam.Tag = strT
    End If
    
  End If
  
  'check si no encontro usuario, return false
  If SQLparam.UsuarioConectado = "" Then
    SQLgetParam = False
  End If
  
End Function

'devuelve usuario conectado
'
Public Function SQLgetUsuario() As String
  
  Dim strT As String
  Dim rs As New ADODB.Recordset
    
  'dafault devuelve ningun usuario
  SQLgetUsuario = ""
  
  'busco usuario
  strT = "select system_user"
  Set rs = SQLexec(strT)
    
  'chequeo error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    Exit Function
  End If
          
  'si encuentro devuelvo nombre
  If Not rs.EOF Then
    SQLgetUsuario = rs(0)
  End If
    
End Function


'devuelve un array con el nombre de usuario si es seguridad sql
'devuelve un array con los nombres de grupos a los que pertenece el usuario
'
Public Function SQLgetGrupo() As Variant
  
  Dim strT As String
  Dim rsG, rsM As New ADODB.Recordset
  Dim intI As Integer
  Dim arrAUX As Variant
    
  'return array vacio default
  arrAUX = Array()
    
  'seguridad SQL
  If SQLparam.SeguridadIntegrada = False Then
      
    'busco usuario
    strT = "select name " & _
           "from sysusers " & _
           "Where isSqlUser = 1 and name = '" & SQLparam.Usuario & "'"
    Set rs = SQLexec(strT)
    
    'chequeo error
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      Exit Function
    End If
          
    'si encuentro devuelvo nombre
    If Not rs.EOF Then
      
      'redim array
      ReDim Preserve arrAUX(1)
      
      'put nombre usuario
      arrAUX(1) = rs!Name
      
    End If
            
  End If
            
  'seguridad integrada
  If SQLparam.SeguridadIntegrada = True Then

    'get lista de grupos
    strT = "select name " & _
           "From sysusers " & _
           "Where (isNtGroup = 1 or isNtUser = 1) and uid not in (1,2)"
    
    'execute
    Set rsG = SQLexec(strT)
        
    'check error
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      Exit Function
    End If
        
    'inicializo
    intI = 0
    
    'while lista de grupos
    Do While Not rsG.EOF
      
      'check si usuario conectado pertenece al grupo
      strT = "select is_member('" & rsG!Name & "') as Miembro"
            
      'execute
      Set rsM = SQLexec(strT)
        
      'check error
      If Not SQLparam.CnErrNumero = -1 Then
        SQLError
        Exit Function
      End If
      
      'check si pertenece al grupo, add array
      If Not rsM.EOF And rsM!miembro = 1 Then
        
        'acum
        intI = intI + 1
        
        'redim array
        ReDim Preserve arrAUX(intI)
        
        'put nombre usuario
        arrAUX(intI) = rsG!Name
        
      End If
            
      'moveNext puntero
      rsG.MoveNext
      
    Loop
    
  End If
    
  'devuelvo array
  SQLgetGrupo = arrAUX
    
End Function

'leo nombre del role asociado al usuario de conexion
Public Function SQLgetRole() As String
  
  Dim strT As String
  Dim rs As New ADODB.Recordset
  
  'default devuelve ningun role
  SQLgetRole = ""
  
  'chequeo si existe la tabla menuRoles para no generar un error
  strT = "select name from sysObjects where name = 'menuRoles'"
  Set rs = SQLexec(strT)
  
  'chequeo errores
  If Not SQLparam.CnErrNumero = -1 Then
    Exit Function
  End If
    
  'si no existe fin
  If rs.EOF Then
    Exit Function
  End If
    
  'busco role correspondiente a usuario conectado
  strT = "select role from menuRoles where usuario = '" & SQLparam.UsuarioConectado & "'"
  Set rs = SQLexec(strT)
  
  'chequeo errores
  If Not SQLparam.CnErrNumero = -1 Then
    Exit Function
  End If
          
  'si no encontro usuario fin
  If rs.EOF Then
    Exit Function
  End If
    
  'devuelvo role
  SQLgetRole = rs!Role
      
End Function


'ejecuta un comando SQL
Public Function SQLexec(ByVal strSQL As String, Optional ByVal strProvider As String, Optional ByVal strServerName As String, Optional ByVal strDatabaseName As String, Optional ByVal blnIntegratedSecurity As Boolean, Optional ByVal intTimeOut As Integer, Optional ByVal strUserID As String, Optional ByVal strPassword As String) As ADODB.Recordset
  
  Dim strT As String
  Dim rs As New ADODB.Recordset
  
  'control de errores
  On Error GoTo controlError
    
  'default sin error
  SQLparam.CnErrNumero = True
  SQLparam.CnErrTexto = ""
    
  'si conexion cerrada, abro
  If SQLparam.cn.State = adStateClosed Then
    
    'si default vacio, asigno false
    If SQLparam.SeguridadIntegrada = "" Then SQLparam.SeguridadIntegrada = "False"
    
    'si parametros vacios, asigno default
    If strProvider = "" Then strProvider = SQLparam.Proveedor
    If strServerName = "" Then strServerName = SQLparam.ServerDatos
    If strDatabaseName = "" Then strDatabaseName = SQLparam.BaseDEdatos
    If blnIntegratedSecurity = False Then blnIntegratedSecurity = CBool(SQLparam.SeguridadIntegrada)
    If intTimeOut = 0 Then intTimeOut = Val(SQLparam.TiempoEspera)
    If strUserID = "" Then strUserID = SQLparam.Usuario
    If strPassword = "" Then strPassword = SQLparam.UsuarioClave
    
    'proveedor y server
    SQLparam.cn.Provider = strProvider
    SQLparam.cn.Properties("Data Source") = strServerName
    
    'seguridad integrada
    If blnIntegratedSecurity = True Then
      SQLparam.cn.Properties("Integrated Security") = "SSPI"
    Else
      SQLparam.cn.Properties("User Id") = strUserID
      SQLparam.cn.Properties("Password") = strPassword
    End If
    
    'time out
    If intTimeOut <> 0 Then
      SQLparam.cn.ConnectionTimeout = intTimeOut
      SQLparam.cn.CommandTimeout = intTimeOut
    End If
    
    'abro conexion
    SQLparam.cn.Open
    
    'base de datos default
    SQLparam.cn.DefaultDatabase = strDatabaseName
    
    'si usamos seguridad por roles de aplicacion, la activo
    If SQLparam.Role <> "" Then
      strT = "exec sp_setappRole '" & SQLparam.Role & "', '" & SQLparam.RoleClave & "'"
      SQLparam.cn.Execute strT
    End If
    
  End If
      
  'check si hay un select, es rs, sino exec
  If InStr(LCase(strSQL), "select") Then
    Set rs.ActiveConnection = SQLparam.cn
    rs.CursorLocation = adUseClient       'cursor cliente
    rs.CursorType = adOpenStatic          'tipo estatico
    rs.Source = strSQL                    'origen igual a query
    rs.Open , , , , adCmdText             'abro le indico que le estoy pasando un string con el Query
    Set SQLexec = rs                      'devuelvo recordset
    rs.ActiveConnection = Nothing         'desconecto recordset con coneccion
  Else
    SQLparam.cn.Execute strSQL
    Set SQLexec = Nothing
  End If
  
  Exit Function                     'exit funcion
  
  'control errores
controlError:
  SQLparam.CnErrNumero = Err.Number
  SQLparam.CnErrTexto = Err.Description

End Function

'cierra conexion
Public Function SQLclose()
  
  'si conexion abierta, cierro
  If SQLparam.cn.State = adStateOpen Then
  
    SQLparam.cn.Close
    Set SQLparam.cn = Nothing
    
  End If
  
End Function

'
'control de errores de ADO
Function SQLError() As Boolean
  
  Dim intN As Integer
  Dim blnB As Boolean
  
  'default
  SQLError = True
  
  Select Case SQLparam.CnErrNumero
      
  Case 20               'no hay errores
      
'  Case -2147217873      'clave primaria repetida
'    intN = MsgBox("Esta intentando agregar información que ya existe.", vbCritical + vbOKOnly, "atención...")
      
'  Case -2147217911      'no tiene permisos
'    intN = MsgBox("Esta intentando realizar una operación, para la cual no tiene permisos necesarios.", vbCritical + vbOKOnly, "atención...")
      
  Case Else             'otros errores
    blnB = MsgBox("Error: " & " " & SQLparam.CnErrNumero & " " & SQLparam.CnErrTexto, vbCritical, "Atención...")
          
  End Select

End Function

