VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claCostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'
'LOCK grilla
'
Public Sub spd_Bloquear(spdG As fpSpread, blnBloquear As Boolean)
  
  spdG.col = 1
  spdG.col2 = spdG.MaxCols
  spdG.row = 1
  spdG.row2 = spdG.MaxRows
  spdG.BlockMode = True
  spdG.Lock = blnBloquear
  spdG.BlockMode = False
      
  'LOCK celdas de cabeza siempre
  If spdG.Name = "spdCab" Then
    
    spdG.col = 2
    spdG.row = 1
    spdG.Lock = True
    
    spdG.row = 4
    spdG.Lock = True
    
    spdG.row = 5
    spdG.Lock = True
    
  End If
      
End Sub

'
'CLEAR grilla cab
'
Public Sub spdCab_Blanquear(spdG As fpSpread)
  
  spdG.ClearRange 2, 2, 2, spdG.MaxRows, True
  spdG.ClearRange 5, 1, 5, spdG.MaxRows, True
  
End Sub

'
'SAVE cab
'
Public Sub spdCab_Guardar(spdG As fpSpread)
      
  Dim dtmFecha As Date
  Dim intID, intIDpozo, intIDeleLog, intIDudLog, intIDwelsite As Integer
  Dim sngDfwa, sngTopeZone, sngPerforation, sngBasicCharge, sngDiscAdd, sngDiscLog, sngQuote As Single
  Dim strCurrency, strT As String
  Dim cCostos As New claCostos
  
  'GET datos e ID en caso de LISTA
  spdG.col = 2
  
  spdG.row = 1
  intID = Val(spdG.Text)
  spdG.row = 2
  dtmFecha = spdG.Text
  spdG.row = 3
  intIDpozo = spdG.TypeComboBoxCurSel
  spdG.row = 6
  intIDeleLog = spdG.TypeComboBoxCurSel
  spdG.row = 7
  intIDudLog = spdG.TypeComboBoxCurSel
  spdG.row = 8
  intIDwelsite = spdG.TypeComboBoxCurSel
  spdG.col = 5
  
  spdG.row = 1
  sngDfwa = Val(spdG.Value)
  spdG.row = 2
  sngTopeZone = Val(spdG.Value)
  spdG.row = 3
  sngPerforation = Val(spdG.Value)
  spdG.row = 4
  sngBasicCharge = Val(spdG.Value)
  spdG.row = 5
  sngDiscAdd = Val(spdG.Value)
  spdG.row = 6
  sngDiscLog = Val(spdG.Value)
  spdG.row = 7
  intCurrency = spdG.Text
  spdG.row = 8
  sngQuote = Val(spdG.Value)
  
  'PUT ID en LISTA segun elemento seleccionado
  spdG.col = 3
  
  spdG.row = 3
  spdG.TypeComboBoxCurSel = Val(intIDpozo)
  spdG.row = 6
  spdG.TypeComboBoxCurSel = Val(intIDeleLog)
  spdG.row = 7
  spdG.TypeComboBoxCurSel = Val(intIDudLog)
  spdG.row = 8
  spdG.TypeComboBoxCurSel = Val(intIDwelsite)
  
  'GET ID de LISTA
  spdG.row = 3
  intIDpozo = spdG.TypeComboBoxCurSel
  spdG.row = 6
  intIDeleLog = spdG.TypeComboBoxCurSel
  spdG.row = 7
  intIDudLog = spdG.TypeComboBoxCurSel
  spdG.row = 8
  intIDwelsite = spdG.TypeComboBoxCurSel
  
  'CHECK operacion
  Select Case cCostos.strOP
  
  Case "INS"
    strT = "costos_INS_sp "
    
  Case "EDI"
    strT = "costos_EDI_sp "
  
  Case "ELI"
    strT = "costos_ELI_sp "
    
  End Select
  
  'BUILD query
  strT = strT & intID & ",'" & dateToIso(dtmFecha) & "'," & intIDpozo & "," & _
         intIDeleLog & "," & intIDudLog & "," & intIDwelsite & "," & sngDfwa & "," & sngTopeZone & "," & _
         sngPerforation & "," & sngBasicCharge & "," & sngDiscAdd & "," & sngDiscLog & ",'" & _
         intCurrency & "'," & sngQuote
  
  'SAVE
  SQLexec ("exec " & strT)
        
  'CHECK error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    Exit Sub
  End If
    
  'CLOSE
  SQLclose
    
    
End Sub

