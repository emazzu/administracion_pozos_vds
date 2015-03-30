Attribute VB_Name = "modCostos"
'
'constantes columnas grilla encabezado
Public Const conCABpozo = 3
Public Const conCABlog = 6
Public Const conCABmud = 7
Public Const conCABsite = 8

'constantes columnas grilla detalle
Public Const conDETempresa = 1
Public Const conDETtrabajo = 2
Public Const conDays = 3
Public Const conFrom = 4
Public Const conTo = 5
Public Const conM = 6
Public Const conVSPS = 7
Public Const conCheck = 8
Public Const conPressure = 9
Public Const conCost = 10
Public Const conCostDesc = 11
Public Const conComment = 12
Public Const conTipo = 13
Public Const conXdia = 14
Public Const conXpozo = 15
Public Const conXprof = 16
Public Const conXregistroMen = 17
Public Const conXregistroMas = 18
Public Const conDesc = 19
Public Const conFormula = 20

'estructura Detalle
'
Type costoDet
  StrEmpresa As Variant
  StrTrabajo As Variant
  intDays As Variant
  intFrom As Variant
  intTo As Variant
  intM As Variant
  intVSPS As Variant
  intCheck As Variant
  intPressure As Variant
  sngCost As Variant
  sngCostDesc As Variant
  strComment As Variant
End Type

Global typeCostoDet As costoDet

'estructura formula
'
Type Formula
  strTipo As Variant
  dtmFecha As Variant
  sngXdia As Variant
  sngXpozo As Variant
  sngXprofundidad As Variant
  sngXregistroMen As Variant
  sngXtegistroMas As Variant
  sngXdescPje As Variant
  strFormula As Variant
End Type

Global typeFormula As Formula

