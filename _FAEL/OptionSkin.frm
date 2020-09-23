VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmTema 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aplicar tema"
   ClientHeight    =   2145
   ClientLeft      =   4950
   ClientTop       =   3570
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4080
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3360
      OleObjectBlob   =   "OptionSkin.frx":0000
      Top             =   1080
   End
   Begin ACTIVESKINLibCtl.SkinLabel sbarInfo 
      Height          =   255
      Left            =   0
      OleObjectBlob   =   "OptionSkin.frx":0234
      TabIndex        =   4
      Top             =   1800
      Width           =   4095
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   720
      OleObjectBlob   =   "OptionSkin.frx":02AE
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.ComboBox cboTema 
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton cmdAplicatema 
      Caption         =   "&Aplicar tema"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "frmTema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private ultElem As Integer
    Private x As Integer
    Private sTmp As String

Private Sub cmdAplicatema_Click()
    strSkinName = cboTema.Text 'Evitar cbo vacío
    Aplicar_skin frmCalc
    Unload Me
End Sub

Private Sub Form_Load()
    File1.Path = App.Path & "\Skins\" 'carga la ruta en el filectrl
    ultElem = 0
    Aplicar_skin Me
    sbarInfo.Caption = "Skin: " & " " & strSkinName
    'Añadir temas al cbo
    x = File1.ListCount
    Do Until x = 0
        x = x - 1
        ultElem = ultElem + 1
        'cboTema.AddItem mnutema(ultElem)
        'Load mnutema(ultElem)
        cboTema.AddItem File1.List(x)
        'mnutema(ultElem).Caption = File1.List(x) 'File1.List(ultElem)   Dir$(App.Path & "\Skins\", vbNormal) & Str(ultElem)
    Loop
    On Local Error Resume Next
    sTmp = Dir$(App.Path & "\Skins\", vbNormal)
    cboTema.Text = strSkinName
End Sub
