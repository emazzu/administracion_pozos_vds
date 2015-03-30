VERSION 5.00
Object = "{1937DC2F-517A-11D1-9F5D-00C04FC32E79}#2.0#0"; "FlpGrf.ocx"
Begin VB.Form graficarFRM 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Graficando..."
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox tipo 
      Height          =   315
      ItemData        =   "graficarFRM.frx":0000
      Left            =   870
      List            =   "graficarFRM.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   90
      Width           =   795
   End
   Begin VB.CheckBox titulo 
      Caption         =   "titulo"
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   645
   End
   Begin FLPGRFLibCtl.FlpGrf Graf 
      Height          =   5145
      Left            =   90
      TabIndex        =   0
      Top             =   660
      Width           =   8535
      _Version        =   131072
      _ExtentX        =   15055
      _ExtentY        =   9075
      _StockProps     =   33
      ColorTableSize  =   34
      ColorTable0     =   0
      ColorTable1     =   4210752
      ColorTable2     =   64
      ColorTable3     =   16384
      ColorTable4     =   4194304
      ColorTable5     =   16448
      ColorTable6     =   4194368
      ColorTable7     =   4210688
      ColorTable8     =   8421504
      ColorTable9     =   128
      ColorTable10    =   32768
      ColorTable11    =   8388608
      ColorTable12    =   32896
      ColorTable13    =   8388736
      ColorTable14    =   8421376
      ColorTable15    =   12632256
      ColorTable16    =   255
      ColorTable17    =   65280
      ColorTable18    =   16711680
      ColorTable19    =   65535
      ColorTable20    =   16711935
      ColorTable21    =   16776960
      ColorTable22    =   16777215
      ColorTable23    =   3381759
      ColorTable24    =   3394815
      ColorTable25    =   52275
      ColorTable26    =   16763955
      ColorTable27    =   26316
      ColorTable28    =   39372
      ColorTable29    =   39219
      ColorTable30    =   13408512
      ColorTable31    =   14737632
      ColorTable32    =   33023
      ColorTable33    =   -2147483633
      HatchBits10     =   170
      HatchBits11     =   85
      HatchBits12     =   170
      HatchBits13     =   85
      HatchBits14     =   170
      HatchBits15     =   85
      HatchBits16     =   170
      HatchBits17     =   85
      HatchBits20     =   223
      HatchBits21     =   175
      HatchBits22     =   119
      HatchBits23     =   175
      HatchBits24     =   223
      HatchBits25     =   255
      HatchBits26     =   255
      HatchBits27     =   255
      HatchBits30     =   119
      HatchBits31     =   171
      HatchBits32     =   221
      HatchBits33     =   186
      HatchBits34     =   119
      HatchBits35     =   234
      HatchBits36     =   221
      HatchBits37     =   174
      HatchBits40     =   125
      HatchBits41     =   187
      HatchBits42     =   215
      HatchBits43     =   238
      HatchBits44     =   215
      HatchBits45     =   187
      HatchBits46     =   125
      HatchBits47     =   254
      FontCount       =   1
      FontTable0      =   "Arial,8,N"
      LayerObjectCount0=   2
      ObjectShape0    =   16
      ObjectNumber0   =   3
      GraphState_BkColor0=   31
      GraphState_axisNum0=   0
      GraphState_columnNum0=   0
      AxisState_StartPercent0Axis0=   0
      AxisState_EndPercent0Axis0=   0
      AxisState_StartPercent0Axis1=   0
      AxisState_EndPercent0Axis1=   0
      AxisState_StartPercent0Axis2=   0
      AxisState_EndPercent0Axis2=   0
      ColumnState_lineType0Column0=   1
      ColumnState_Side0Column0=   0
      ColumnState_Search0Column0=   0
      ColumnState_lineType0Column1=   3
      ColumnState_Side0Column1=   1
      ColumnState_color0Column1=   10
      ColumnState_style0Column1=   1
      ColumnState_Search0Column1=   0
      ColumnState_BarGradientType0Column1=   1
      Column_legend0Column1=   "1"
      ColumnState_lineType0Column2=   4
      ColumnState_Side0Column2=   1
      ColumnState_color0Column2=   32
      ColumnState_style0Column2=   1
      ColumnState_Search0Column2=   0
      ColumnState_BarGradientType0Column2=   1
      Column_legend0Column2=   "2"
      ColumnState_Search0Column3=   0
      ColumnState_Search0Column4=   0
      ColumnState_Search0Column5=   0
      ColumnState_Search0Column6=   0
      Data_UserModeSaved0=   0   'False
      DataState_rows0 =   10
      DataState_columns0=   3
      Data_Params0R1C0=   0
      Data_VT0R1C0    =   4
      Data_Variant0R1C0=   10
      Data_Params0R1C1=   0
      Data_VT0R1C1    =   4
      Data_Variant0R1C1=   72
      Data_Params1R1C2=   0
      Data_VT1R1C2    =   4
      Data_Variant1R1C2=   71
      Data_Params2R2C0=   0
      Data_VT2R2C0    =   4
      Data_Variant2R2C0=   20
      Data_Params0R2C1=   0
      Data_VT0R2C1    =   4
      Data_Variant0R2C1=   52
      Data_Params1R2C2=   0
      Data_VT1R2C2    =   4
      Data_Variant1R2C2=   51
      Data_Params2R3C0=   0
      Data_VT2R3C0    =   4
      Data_Variant2R3C0=   30
      Data_Params0R3C1=   0
      Data_VT0R3C1    =   4
      Data_Variant0R3C1=   36
      Data_Params1R3C2=   0
      Data_VT1R3C2    =   4
      Data_Variant1R3C2=   82
      Data_Params2R4C0=   0
      Data_VT2R4C0    =   4
      Data_Variant2R4C0=   40
      Data_Params0R4C1=   0
      Data_VT0R4C1    =   4
      Data_Variant0R4C1=   37
      Data_Params1R4C2=   0
      Data_VT1R4C2    =   4
      Data_Variant1R4C2=   51
      Data_Params2R5C0=   0
      Data_VT2R5C0    =   4
      Data_Variant2R5C0=   50
      Data_Params0R5C1=   0
      Data_VT0R5C1    =   4
      Data_Variant0R5C1=   53
      Data_Params1R5C2=   0
      Data_VT1R5C2    =   4
      Data_Variant1R5C2=   78
      Data_Params2R6C0=   0
      Data_VT2R6C0    =   4
      Data_Variant2R6C0=   60
      Data_Params0R6C1=   0
      Data_VT0R6C1    =   4
      Data_Variant0R6C1=   98
      Data_Params1R6C2=   0
      Data_VT1R6C2    =   4
      Data_Variant1R6C2=   58
      Data_Params2R7C0=   0
      Data_VT2R7C0    =   4
      Data_Variant2R7C0=   70
      Data_Params0R7C1=   0
      Data_VT0R7C1    =   4
      Data_Variant0R7C1=   68
      Data_Params1R7C2=   0
      Data_VT1R7C2    =   4
      Data_Variant1R7C2=   50
      Data_Params2R8C0=   0
      Data_VT2R8C0    =   4
      Data_Variant2R8C0=   80
      Data_Params0R8C1=   0
      Data_VT0R8C1    =   4
      Data_Variant0R8C1=   70
      Data_Params1R8C2=   0
      Data_VT1R8C2    =   4
      Data_Variant1R8C2=   88
      Data_Params2R9C0=   0
      Data_VT2R9C0    =   4
      Data_Variant2R9C0=   90
      Data_Params0R9C1=   0
      Data_VT0R9C1    =   4
      Data_Variant0R9C1=   84
      Data_Params1R9C2=   0
      Data_VT1R9C2    =   4
      Data_Variant1R9C2=   21
      Data_Params2R10C0=   0
      Data_VT2R10C0   =   4
      Data_Variant2R10C0=   100
      Data_Params0R10C1=   0
      Data_VT0R10C1   =   4
      Data_Variant0R10C1=   86
      Data_Params1R10C2=   0
      Data_VT1R10C2   =   4
      Data_Variant1R10C2=   40
      ObjectShape1    =   17
      ObjectNumber1   =   4
      LegendStateXPos1=   98
      LegendStateYPos1=   2
      LegendStatePosHorzRef1=   1
      LegendStatePosVertRef1=   1
      LegendStateJust1=   2
      LegendStateAlign1=   0
   End
End
Attribute VB_Name = "graficarFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_gridFrm As gridFRM

Public Property Set dsiGridFrm(frm As gridFRM)
  Set m_gridFrm = frm
End Property
    
Public Property Get dsiGridFrm() As gridFRM
  Set dsiGridFrm = m_gridFrm
End Property

Private Sub Form_Load()
    
  'referencio grilla
  Set Me.dsiGridFrm = MainMDI.ActiveForm
    
  Me.Graf.GraphEdit = flpDataMenus
    
  'set tipo
  Me.tipo = 0
  Me.Graf.GraphType = Me.tipo
     
  Me.Graf.DataInit = 3
     
  Me.Graf.DataValue(1, 0) = "01/01/2005"
  Me.Graf.DataValue(2, 0) = "01/02/2005"
  Me.Graf.DataValue(3, 0) = "01/03/2005"
  Me.Graf.DataValue(4, 0) = "01/04/2005"
  Me.Graf.DataValue(5, 0) = "01/05/2005"
     
     
  Me.Graf.DataValue(1, 1) = 1000
  Me.Graf.DataValue(2, 1) = 1055
  Me.Graf.DataValue(3, 1) = 2050
  Me.Graf.DataValue(4, 1) = 3000
  Me.Graf.DataValue(5, 1) = 5000
     
  Me.Graf.DataValue(1, 2) = 200
  Me.Graf.DataValue(2, 2) = 300
  Me.Graf.DataValue(3, 2) = 500
  Me.Graf.DataValue(4, 2) = 600
  Me.Graf.DataValue(5, 2) = 900
     
End Sub

Private Sub tipo_click()

  Me.Graf.GraphType = Val(Me.tipo.Text)

End Sub

Private Sub titulo_Click()
  
  If titulo.Value = 1 Then
    Me.Graf.GraphTitle = "Titulo del Grafico"
  Else
    Me.Graf.GraphTitle = ""
  End If
    
End Sub
