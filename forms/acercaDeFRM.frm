VERSION 5.00
Begin VB.Form acercaDeFRM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About DSInfo Data..."
   ClientHeight    =   4875
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5460
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1100
      Left            =   90
      TabIndex        =   6
      Top             =   3165
      Width           =   5300
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "www.dsinfo.com.ar"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   90
         TabIndex        =   8
         Top             =   700
         Width           =   5200
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ventas@dsinfo.com.ar - soporte-aplicaciones@dsinfo.com.ar  soporte-tecnico@dsinfo.com.ar"
         ForeColor       =   &H00000000&
         Height          =   500
         Left            =   80
         TabIndex        =   7
         Top             =   200
         Width           =   5200
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1875
      Left            =   90
      TabIndex        =   3
      Top             =   45
      Width           =   5300
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Expediente: 926839 - 12/05/2011 - Registro de la Propiedad Intelectual"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   555
         TabIndex        =   10
         Top             =   1545
         Width           =   4500
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image2 
         Height          =   1500
         Left            =   70
         Picture         =   "acercaDeFRM.frx":0000
         Stretch         =   -1  'True
         Top             =   50
         Width           =   1500
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Todos los derechos reservados."
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   1695
         TabIndex        =   9
         Top             =   1005
         Width           =   3525
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (c) 2006 por Eduardo Mazzucchelli."
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1700
         TabIndex        =   5
         Top             =   750
         Width           =   3250
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DSInfo Data v.2011.0.1"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1700
         TabIndex        =   4
         Top             =   400
         Width           =   1725
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1250
      Left            =   90
      TabIndex        =   1
      Top             =   1905
      Width           =   5300
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"acercaDeFRM.frx":479E
         ForeColor       =   &H00000000&
         Height          =   1005
         Left            =   75
         TabIndex        =   2
         Top             =   150
         Width           =   5205
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4100
      TabIndex        =   0
      Top             =   4380
      Width           =   1260
   End
End
Attribute VB_Name = "acercaDeFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  
  Unload Me

End Sub

