VERSION 5.00
Begin VB.Form frmConexion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Origen de datos / ODBC x.x (mySql)"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame fraConexion 
      Caption         =   "Database conection for MySQL x.x ODBC Driver"
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3855
      Begin VB.TextBox txtPuerto 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Text            =   "3306"
         Top             =   1160
         Width           =   735
      End
      Begin VB.TextBox txtBasededatos 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Text            =   "bdFAEL"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtServidor 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Text            =   "localhost"
         Top             =   330
         Width           =   2055
      End
      Begin VB.TextBox txtUsuario 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Text            =   "root"
         Top             =   1960
         Width           =   2055
      End
      Begin VB.TextBox txtContrasena 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   2360
         Width           =   2055
      End
      Begin VB.ComboBox cboProveedor 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   730
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Puerto"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1160
         Width           =   465
      End
      Begin VB.Label lblProveedor 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   760
         Width           =   735
      End
      Begin VB.Label lblContrasena 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2360
         Width           =   810
      End
      Begin VB.Label lblUsuario 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1960
         Width           =   540
      End
      Begin VB.Label lblServidor 
         AutoSize        =   -1  'True
         Caption         =   "Servidor"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Base de datos"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmConexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    Call Conexion
End Sub

Private Sub cmdCancelar_Click()
    End
End Sub

Private Sub Form_Load()
    cboProveedor.AddItem "MySQL ODBC 5.1 Driver"
    cboProveedor.AddItem "MySQL ODBC 3.51 Driver"
End Sub
