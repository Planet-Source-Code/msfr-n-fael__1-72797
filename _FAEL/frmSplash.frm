VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7560
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   495
         Left            =   2280
         OleObjectBlob   =   "frmSplash.frx":0000
         TabIndex        =   10
         Top             =   2040
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblVersion 
         Height          =   255
         Left            =   5280
         OleObjectBlob   =   "frmSplash.frx":009E
         TabIndex        =   7
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Timer tmrReloj 
         Interval        =   25
         Left            =   2760
         Top             =   2880
      End
      Begin VB.PictureBox Picture1 
         Height          =   2595
         Left            =   120
         ScaleHeight     =   2535
         ScaleWidth      =   1980
         TabIndex        =   4
         Top             =   720
         Width           =   2040
         Begin VB.Image Image1 
            Height          =   2535
            Left            =   0
            Picture         =   "frmSplash.frx":0118
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1980
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "frmSplash.frx":2D23
         TabIndex        =   5
         Top             =   120
         Width           =   6975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   3720
         OleObjectBlob   =   "frmSplash.frx":2DC1
         TabIndex        =   6
         Top             =   3240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   3720
         OleObjectBlob   =   "frmSplash.frx":2E35
         TabIndex        =   0
         Top             =   3000
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   6960
         OleObjectBlob   =   "frmSplash.frx":2EB3
         Top             =   120
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmSplash.frx":30E7
         TabIndex        =   1
         Top             =   3720
         Width           =   3615
      End
      Begin MSComctlLib.ProgressBar pgbCargar 
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Top             =   3600
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblPrograma 
         Height          =   855
         Left            =   2280
         OleObjectBlob   =   "frmSplash.frx":3199
         TabIndex        =   8
         Top             =   1200
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblUsuario 
         Height          =   375
         Left            =   2280
         OleObjectBlob   =   "frmSplash.frx":31FF
         TabIndex        =   9
         Top             =   720
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblPrograma = App.Title
    Aplicarskin Me
    lblVersion = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub tmrReloj_Timer()
    If pgbCargar.Value < 100 Then
        pgbCargar.Value = pgbCargar.Value + 1
    Else
        frmLogin.Show vbModal, mdiFAEL
        Unload Me
    End If
End Sub
