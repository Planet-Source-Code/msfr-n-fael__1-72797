VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmEmpresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "..."
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   8805
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   8040
      OleObjectBlob   =   "frmEmpresa.frx":0000
      Top             =   120
   End
End
Attribute VB_Name = "frmEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Aplicarskin Me
    Me.Left = PosX(Me)
    Me.Top = PosY(Me)
End Sub
