VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmPermisos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "..."
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAplicar 
      Cancel          =   -1  'True
      Caption         =   "&Aplicar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4740
      TabIndex        =   8
      Top             =   1020
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4740
      TabIndex        =   2
      Top             =   180
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4740
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4740
      OleObjectBlob   =   "frmPermisos.frx":0000
      Top             =   1440
   End
   Begin VB.Frame fraSeguridad 
      Caption         =   "Seguridad"
      Height          =   5475
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4515
      Begin VB.OptionButton optUsuarios 
         Caption         =   "&Usuarios restringidos"
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   7
         Top             =   300
         Width           =   1995
      End
      Begin VB.OptionButton optUsuarios 
         Caption         =   "A&dministradores"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   300
         Width           =   1995
      End
      Begin VB.ListBox lstUsuarios 
         Height          =   1230
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   4095
      End
      Begin VB.ListBox lstPermisos 
         Height          =   2985
         Left            =   240
         TabIndex        =   3
         Top             =   2340
         Width           =   4095
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblPermisos 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmPermisos.frx":0234
         TabIndex        =   4
         Top             =   2100
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmPermisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Aplicarskin Me
    Me.Caption = App.Title & " .:. Permisos de usuario"
    Me.Left = PosX(Me)
    Me.Top = PosY(Me)
    'Selecciona administradores al cargar el form
    optUsuarios(0).Value = True
End Sub


Private Sub optUsuarios_Click(Index As Integer)
    If Index = 0 Then ''Llena la lista de Usuarios tipo ADMIN
        lstUsuarios.Clear
        Call Consulta("SELECT * FROM pusr WHERE usr_activo = 1 AND usr_tipo = 1 ORDER BY usr_nombre", rs)
        rs.MoveFirst
        Do
           lstUsuarios.AddItem rs!usr_nombre
           rs.MoveNext
        Loop Until rs.EOF = True
        rs.Close
    Else    'Llena la lista de Usuarios tipo USER
        lstUsuarios.Clear
        Call Consulta("SELECT * FROM pusr WHERE usr_activo = 1 AND usr_tipo = 2 ORDER BY usr_nombre", rs)
        rs.MoveFirst
        Do
           lstUsuarios.AddItem rs!usr_nombre
           rs.MoveNext
        Loop Until rs.EOF = True
        rs.Close
    End If
End Sub
