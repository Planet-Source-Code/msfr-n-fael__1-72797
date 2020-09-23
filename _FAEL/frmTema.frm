VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmTema 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "..."
   ClientHeight    =   1890
   ClientLeft      =   4950
   ClientTop       =   3570
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame fraTema 
      Caption         =   "Elija tema..."
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdAplicatema 
         Caption         =   "&Aplicar tema"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cboTema 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.SkinLabel sbarInfo 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmTema.frx":0000
         TabIndex        =   1
         Top             =   1320
         Width           =   3015
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "frmTema.frx":007A
      Top             =   240
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
    strSkn = cboTema.Text 'Evitar cbo vacío
    'Graba el actual skin en el registro
    SaveSetting "FAEL", "Skin", "ActualSkn", cboTema.Text
    'Actualiza el nombre del skin actual en la barra de estado
    mdiFAEL.sbrFAEL.Panels(2).Text = Left(strSkn, Len(strSkn) - 4)
    'Aplicarskin mdi
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = App.Title & " .:. Aplicar tema"
    File1.Path = App.Path & "\mmedia\skins\" 'carga la ruta en el filectrl
    ultElem = 0
    Aplicarskin Me
    sbarInfo.Caption = "Tema actual: " & " " & Left(strSkn, Len(strSkn) - 4)
    'Añadir temas al cbo
    x = File1.ListCount
    Do Until x = 0
        x = x - 1
        ultElem = ultElem + 1
        cboTema.AddItem File1.List(x)
    Loop
    On Local Error Resume Next
    sTmp = Dir$(App.Path & "\Skins\", vbNormal)
    cboTema.Text = strSkn
End Sub
