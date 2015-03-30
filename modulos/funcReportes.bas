Attribute VB_Name = "funcReportes"
'
'SIRVE PARA ESPECIFICAR CARACTERISTICAS ADICIONALES A LOS REPORTES
'
'strID:         ID reporte o subreporte como esta definido en menuOpciones
'strNBReport:   string con nombre del rpt
'strDataSource: string con vista asociada al rpt
'rpt:           referencia al reporte o subreportes
'rs:            puntero al recordset que todavia no fue abierto
'strWhere:      string con where
'strWhereArr:   array con where separado para poder acceder a cada condicion por separado
'
Public Function rptAdicional(strID As String, strNBReport As String, strDataSource As String, rpt As CRAXDRT.Report, rs As ADODB.Recordset, strWhere As String, strWhereArr As Variant) As Boolean
  Dim strWhereActual As String
  Dim strIntervalo As String
  
  'valor predeterminado
  rptAdicional = False
  
  Select Case LCase(strID)
    
  'Equipo Pulling Mvg Avg
  '
  Case "205"
      
    'armo where
    strWhereActual = ""
    If strWhere <> "" Then
      strWhereActual = " where " & strWhere
    End If
      
    'predeterminado 3 por si no se selecciono intervalo o se ingreso algo distinto de 3,6,9
    strSQL = " select * from graficosBaseIIP_avg3_vw " & strWhereActual
      
    If arrayGetValue(strWhereArr, "intervalo") <> "" Then
      
      'determino intervalo
      Select Case arrayGetValue(strWhereArr, "intervalo")
      
      Case 2
        strSQL = " select * from graficosBaseIIP_avg2_vw" & strWhereActual
      Case 3
        strSQL = " select * from graficosBaseIIP_avg3_vw" & strWhereActual
      Case 5
        strSQL = " select * from graficosBaseIIP_avg5_vw" & strWhereActual
      Case 6
        strSQL = " select * from graficosBaseIIP_avg6_vw" & strWhereActual
      Case 12
        strSQL = " select * from graficosBaseIIP_avg12_vw" & strWhereActual
      End Select
      
    End If
    
    'abro rs y se lo paso a rpt
    Set rs = SQLexec(strSQL)
    rpt.Database.SetDataSource rs
  
    rptAdicional = True
    
  End Select

End Function

