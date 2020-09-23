VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "..."
   ClientHeight    =   2085
   ClientLeft      =   6750
   ClientTop       =   3660
   ClientWidth     =   5625
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   5415
      Begin VB.PictureBox pctMay 
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   3435
         TabIndex        =   8
         Top             =   1680
         Visible         =   0   'False
         Width           =   3495
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Mayúsculas encendidas..."
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   2220
         End
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar (Esc)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtContraseña 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox cboUsuario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.PictureBox pctSeguridad 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         Left            =   3720
         Picture         =   "frmLogin.frx":0000
         ScaleHeight     =   107.527
         ScaleMode       =   0  'User
         ScaleWidth      =   91.744
         TabIndex        =   7
         Top             =   240
         Width           =   1560
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   120
         OleObjectBlob   =   "frmLogin.frx":0A97
         Top             =   1080
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmLogin.frx":0CCB
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmLogin.frx":0D37
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAceptar_Click()
    
    Call Consulta("pusr", rs)
    rs.MoveFirst
    'Verifica que se haya seleccionado usr & pass
    If cboUsuario.Text = "" Or Trim(txtContraseña.Text) = "" Then
       Call Mensaje(Me, "Debe ingresar usuario y contraseña", "Falta usuario y/o contraseña")
       rs.Close
       Exit Sub
    End If
    'Itera en busca de la combinación usuario contraseña
    Do
        If rs!usr_nombre = Trim(cboUsuario.Text) And rs.Fields("usr_pass") = OCT(Trim(txtContraseña), True) Then
           mdiFAEL.Enabled = True
           mdiFAEL.sbrFAEL.Panels(1).Text = cboUsuario.Text
           Unload Me
           
           'Log de actividades
           '==================
           Exit Do
       End If
       rs.MoveNext
       If rs.EOF = True Then
           Call Mensaje(Me, "La contraseña no coincide", "Acceso denegado")
           txtContraseña.SetFocus
           SendKeys "{Home}+{End}"
       End If
    Loop Until rs.EOF = True
    'Cierra el recordset al terminar cualquier comprobación
    rs.Close
    
End Sub

Private Sub cmdCancelar_Click()
    End
End Sub

Private Sub Form_Load()
    'Centra el form
    Aplicarskin Me
    Me.Caption = App.Title & " .:. Control de acceso"
    Call Consulta("SELECT * FROM pusr WHERE usr_activo = 1 ORDER BY usr_nombre", rs)
    rs.MoveFirst
    Do
       cboUsuario.AddItem rs!usr_nombre
       rs.MoveNext
    Loop Until rs.EOF = True
    rs.Close
    'Me.Left = PosX(Me)
    'Me.Top = PosY(Me)
    
End Sub


