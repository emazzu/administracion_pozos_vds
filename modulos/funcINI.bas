Attribute VB_Name = "funcINI"
'
' API para leer y grabar INIS
'
Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'
'API lee Donain User: 26/09/2006
'
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'
' leo formato de un ini y devuelvo array
' 2 filas y cantidad de columnas como info halla
'
Public Function keyIniToArray(ByVal strHead As String, ByVal strKey As String, Optional ByVal strINIname As String) As Variant
  Dim strDato As String
  Dim intInd, intColumn As Integer
    
  ' nombre de ini default
  If strINIname = "" Then
    strINIname = SQLparam.nombreDAT
  End If
    
  ' leo informacion de INI
  strDato = ReadIni(strHead, strKey, strINIname)

  ' si no encontro ini o clave
  ' devuelvo array vacio y exit
  If strDato = "" Then
    keyIniToArray = Array()
    Exit Function
  End If

  ' devuelvo array con formato
  keyIniToArray = separateText(strDato)

End Function

' strSección , se refiere a lo que va entre corchetes en el <.ini>
' strClave , lo que quieres leer
' Por ejemplo:  de uno llamado <Configuracion.ini>
' [Seccion1]  --> strSección
' MiNombre=JJ --> strClave
'
Public Function ReadIni(strHead As String, strKey As String, Optional strINIname As String) As String
    
    ' nombre de ini default
  If strINIname = "" Then
    strINIname = SQLparam.nombreDAT
  End If
  
  'Los parámetros son:
  'vDefault:      Valor opcional que devolverá
  '               si no se encuentra la clave.
  Dim lpString As String
  Dim LTmp As Long
  Dim sRetVal As String
    
  'Si no se especifica el valor por defecto,
  'asignar incialmente una cadena vacía
  If IsMissing(vDefault) Then
    lpString = ""
  Else
    lpString = vDefault
  End If
    
  sRetVal = String$(50000, 0)
  LTmp = GetPrivateProfileString(strHead, strKey, _
            "", sRetVal, Len(sRetVal), strINIname)
    
  If LTmp = 0 Then
    ReadIni = ""
  Else
    ReadIni = Left(sRetVal, LTmp)
  End If

End Function


'
' guarda una clave en un INI
'
Public Function WriteIni(ByVal strHead As String, ByVal strKey As String, ByVal strValue As String, Optional ByVal strINIname As String)
  Dim LTmp As Long

  ' nombre de ini default
  If strINIname = "" Then
    strINIname = SQLparam.nombreDAT
  End If
    
  LTmp = WritePrivateProfileString(strHead, strKey, strValue, strINIname)

End Function

'
'get and return Domain Name: 26/09/2006
'
Public Function NTuserName() As String
  
  Dim m_myBuf As String * 25
  Dim m_Val As Long
  
  m_Val = GetUserName(m_myBuf, 25)
  NTuserName = Left(m_myBuf, InStr(m_myBuf, Chr(0)) - 1)
  
End Function

