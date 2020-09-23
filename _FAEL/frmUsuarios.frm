VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "..."
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwUsuarios 
      Height          =   2745
      Left            =   120
      TabIndex        =   16
      Top             =   3180
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4842
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3480
      OleObjectBlob   =   "frmUsuarios.frx":0000
      Top             =   2520
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
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
      Left            =   4740
      TabIndex        =   15
      Top             =   300
      Width           =   1455
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "&Actualizar"
      Enabled         =   0   'False
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
      Left            =   4740
      TabIndex        =   14
      Top             =   1020
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
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
      Left            =   4740
      TabIndex        =   13
      Top             =   1740
      Width           =   1455
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
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
      Left            =   4740
      TabIndex        =   6
      Top             =   2460
      Width           =   1455
   End
   Begin VB.Frame fraAdminusuario 
      Caption         =   "Administrar usuarios"
      Height          =   2775
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   4455
      Begin VB.PictureBox pctUsrs 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H8000000D&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   180
         Picture         =   "frmUsuarios.frx":0234
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   17
         ToolTipText     =   "Listar usuarios del sistema"
         Top             =   2160
         Width           =   480
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblUsuario 
         Height          =   255
         Index           =   4
         Left            =   180
         OleObjectBlob   =   "frmUsuarios.frx":07ED
         TabIndex        =   12
         Top             =   1800
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblUsuario 
         Height          =   255
         Index           =   3
         Left            =   180
         OleObjectBlob   =   "frmUsuarios.frx":0869
         TabIndex        =   11
         Top             =   1440
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblUsuario 
         Height          =   255
         Index           =   2
         Left            =   180
         OleObjectBlob   =   "frmUsuarios.frx":08EF
         TabIndex        =   10
         Top             =   1080
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblUsuario 
         Height          =   255
         Index           =   1
         Left            =   180
         OleObjectBlob   =   "frmUsuarios.frx":0961
         TabIndex        =   9
         Top             =   720
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblUsuario 
         Height          =   255
         Index           =   0
         Left            =   180
         OleObjectBlob   =   "frmUsuarios.frx":09CB
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtUsuario 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   2100
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtUsuario 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2100
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtUsuario 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2100
         TabIndex        =   1
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox chkActivo 
         Caption         =   "Activo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2100
         TabIndex        =   5
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txtUsuario 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2100
         MaxLength       =   11
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cboTipo 
         Enabled         =   0   'False
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
         ItemData        =   "frmUsuarios.frx":0A47
         Left            =   2100
         List            =   "frmUsuarios.frx":0A51
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1800
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bLoop As Byte

Private Sub cmdActualizar_Click()
    On Error GoTo errCtrl
    'Verifica que los campos estén diligenciados
    For bLoop = 0 To 3
        If txtUsuario(bLoop) = "" Then
            Call Mensaje(Me, App.Title & " V. " & App.Major - App.Minor & "." & App.Revision & " - Complete la información requerida", "El campo: " & "'" & lblUsuario(bLoop).Caption & "' no puede estar vacío")
            txtUsuario(bLoop).SetFocus
            Exit Sub
        End If
    Next bLoop
    If cboTipo.Text = "" Then
        Call Mensaje(Me, App.Title & " V. " & App.Major - App.Minor & "." & App.Revision & " - Complete la información requerida", "El campo: " & "'" & lblUsuario(bLoop).Caption & "' no puede estar vacío")
        Exit Sub
    End If
    
    'Call Consulta("pusr", rs) No consulta el rs está abierto
    'rs.AddNew
    rs!usr_id = Trim(txtUsuario(0))
    rs!usr_nombre = Trim(txtUsuario(1))
    If Trim(txtUsuario(2)) = Trim(txtUsuario(3)) Then
        rs!usr_pass = OCT(Trim(txtUsuario(2)), True)
    Else
        rs.CancelUpdate
        rs.Close
        Call Mensaje(Me, App.Title & " V. " & App.Major - App.Minor & "." & App.Revision, "La contraseña y su confirmacion no coinciden")
        'Si se equivoca limpia los campos de contraseña
        txtUsuario(2) = ""
        txtUsuario(3) = ""
        Exit Sub
    End If
    If cboTipo.Text = "ADMINISTRADOR" Then
        rs!usr_tipo = 1 'Cod Admin = 1
    Else
        rs!usr_tipo = 2 'Cod Vend = 2
    End If
    rs!usr_activo = chkActivo.Value
    rs.Update
    rs.Close
    'Habilito las cajas de texto and etc..
    For bLoop = 0 To 3
        txtUsuario(bLoop) = ""
        txtUsuario(bLoop).Enabled = False
    Next bLoop
    'cboTipo.Text = ""
    cboTipo.Enabled = False
    chkActivo.Value = 0
    chkActivo.Enabled = False
    cmdNuevo.Caption = "&Nuevo"
    cmdCerrar.Enabled = True
    cmdCancelar.Enabled = False
    cmdActualizar.Enabled = False
    cmdNuevo.Enabled = True
    'txtUsuario(0).SetFocus
    'Control de errores
    Exit Sub
errCtrl:
    Call Mensaje(Me, App.Title & " V. " & App.Major - App.Minor & "." & App.Revision, "Error: " & Err.Number & ": " & Err.Description)
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo errCtrl
    'Si está abierto el rs, cancela la operación actual y lo cierra
    If rs.State = adStateOpen Then
        rs.CancelUpdate
        rs.Close
    End If
    'Habilito las cajas de texto and etc..
    For bLoop = 0 To 3
        txtUsuario(bLoop) = ""
        txtUsuario(bLoop).Enabled = False
    Next bLoop
    'cboTipo.Text = ""
    cboTipo.Enabled = False
    chkActivo.Value = 0
    chkActivo.Enabled = False
    cmdNuevo.Caption = "&Nuevo"
    cmdCerrar.Enabled = True
    cmdCancelar.Enabled = False
    cmdActualizar.Enabled = False
    cmdNuevo.Enabled = True
    'Control de errores
    Exit Sub
errCtrl:
    Call Mensaje(Me, App.Title & " V. " & App.Major - App.Minor & "." & App.Revision, "Error: " & Err.Number & ": " & Err.Description)
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdListar_Click()

End Sub

Private Sub cmdNuevo_Click()
On Error GoTo errCtrl
    If cmdNuevo.Caption = "&Nuevo" Then
        'Habilito las cajas de texto and etc..
        For bLoop = 0 To 3
            txtUsuario(bLoop).Enabled = True
        Next bLoop
        cboTipo.Enabled = True
        chkActivo.Enabled = True
        cmdNuevo.Caption = "&Guardar"
        cmdCerrar.Enabled = False
        cmdCancelar.Enabled = True
        txtUsuario(0).SetFocus
    Else '.Caption = &Guardar
        'Verifica que los campos estén diligenciados
        For bLoop = 0 To 3
            If txtUsuario(bLoop) = "" Then
                Call Mensaje(Me, App.Title & " V. " & App.Major - App.Minor & "." & App.Revision & " - Complete la información requerida", "El campo: " & "'" & lblUsuario(bLoop).Caption & "' no puede estar vacío")
                txtUsuario(bLoop).SetFocus
                Exit Sub
            End If
        Next bLoop
        If cboTipo.Text = "" Then
            Call Mensaje(Me, App.Title & " V. " & App.Major - App.Minor & "." & App.Revision & " - Complete la información requerida", "El campo: " & "'" & lblUsuario(bLoop).Caption & "' no puede estar vacío")
            Exit Sub
        End If
        
        Call Consulta("pusr", rs)
        rs.AddNew
        rs!usr_id = Trim(txtUsuario(0))
        rs!usr_nombre = Trim(txtUsuario(1))
        If Trim(txtUsuario(2)) = Trim(txtUsuario(3)) Then
            rs!usr_pass = OCT(Trim(txtUsuario(2)), True)
        Else
            rs.CancelUpdate
            rs.Close
            Call Mensaje(Me, App.Title & " V. " & App.Major - App.Minor & "." & App.Revision, "La contraseña y su confirmacion no coinciden")
            'Si se equivoca limpia los campos de contraseña
            txtUsuario(2) = ""
            txtUsuario(3) = ""
            Exit Sub
        End If
        If cboTipo.Text = "ADMINISTRADOR" Then
            rs!usr_tipo = 1 'Cod Admin = 1
        Else
            rs!usr_tipo = 2 'Cod Vend = 2
        End If
        rs!usr_activo = chkActivo.Value
        'Deshabilito las cajas de texto and etc..
        For bLoop = 0 To 3
            txtUsuario(bLoop) = ""
            txtUsuario(bLoop).Enabled = False
        Next bLoop
        'cboTipo.Text = ""
        cboTipo.Enabled = False
        chkActivo.Value = 0
        chkActivo.Enabled = False
        cmdNuevo.Caption = "&Nuevo"
        cmdCerrar.Enabled = True
        cmdCancelar.Enabled = False
        cmdActualizar.Enabled = False
        cmdNuevo.Enabled = True
        rs.Update
        rs.Close
    End If
    'Control de errores
    Exit Sub
errCtrl:
    Call Mensaje(Me, App.Title & " V. " & App.Major - App.Minor & "." & App.Revision, "Error: " & Err.Number & ": " & Err.Description)
End Sub

Private Sub Form_Load()
    Aplicarskin Me
    'Centrar el form
    Me.Caption = App.Title & " .:. Usuarios"
    Me.Left = PosX(Me)
    Me.Top = PosY(Me)
    Me.Height = 3580
End Sub

Private Sub pctUsrs_Click()
    Dim usr_tipo As String
    Dim usr_activo As String
    Dim itmX As ListItem
    Dim colX As ColumnHeader
    
    If pctUsrs.Tag = "True" Then
        pctUsrs.Tag = ""
        Me.Height = 3570
    Else
        Me.Height = 6495
        pctUsrs.Tag = "True"
        If lvwUsuarios.ColumnHeaders.Count = 0 Then
            Call Consulta("pusr", rs)
            rs.MoveFirst
            'Agrega los encabezados
            lvwUsuarios.ColumnHeaders.Add , , "Documento", lvwUsuarios.Width / 4
            lvwUsuarios.ColumnHeaders.Add , , "Nombre", lvwUsuarios.Width / 4
            lvwUsuarios.ColumnHeaders.Add , , "Tipo de usuario", lvwUsuarios.Width / 4
            lvwUsuarios.ColumnHeaders.Add , , "Estado", lvwUsuarios.Width / 4
            
            lvwUsuarios.View = lvwReport
            Do While Not rs.EOF
                bLoop = 0
                bLoop = bLoop + 1
                If rs!usr_activo = 1 Then
                    usr_activo = "ACTIVO"
                Else
                    usr_activo = "INACTIVO"
                End If
                If rs!usr_tipo = 1 Then
                    usr_tipo = "ADMINISTRADOR"
                Else
                    usr_tipo = "USUARIO"
                End If
                'Añade un registro al lstv
                Set itmX = lvwUsuarios.ListItems.Add(bLoop, , CStr(rs!usr_id))
                itmX.SubItems(bLoop) = CStr(rs!usr_nombre)
                itmX.SubItems(bLoop + 1) = CStr(usr_tipo)
                itmX.SubItems(bLoop + 2) = CStr(usr_activo)
                rs.MoveNext
            Loop
            rs.Close
        End If
        'pctUsrs.Enabled = False
    End If
End Sub

Private Sub txtUsuario_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0
            If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            If KeyAscii = 13 Then SendKeys "{tab}"
        Case 1
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If KeyAscii = 13 Then SendKeys "{tab}"
        Case 2
            If KeyAscii = 13 Then SendKeys "{tab}"
        Case 3
            If KeyAscii = 13 Then SendKeys "{tab}"
    End Select
End Sub

Private Sub txtUsuario_LostFocus(Index As Integer)
    If Index = 0 Then
        On Error GoTo errCtrl
        Call Consulta("pusr", rs)
        rs.MoveFirst
        Do While Not rs.EOF
            If Val(Trim(txtUsuario(0))) = Val(Trim(rs!usr_id)) Then
                txtUsuario(1) = rs!usr_nombre
                txtUsuario(2) = OCT(rs!usr_pass, False)
                txtUsuario(3) = OCT(rs!usr_pass, False)
                If rs!usr_tipo = 1 Then
                    cboTipo.Text = "ADMINISTRADOR"
                Else
                    cboTipo.Text = "USUARIO"
                End If
                If rs!usr_activo = "1" Then
                    chkActivo.Value = 1
                Else
                    chkActivo.Value = 0
                End If
                'Solo en caso de que encuentre un usuario con el id introducido
                cmdNuevo.Enabled = False
                cmdActualizar.Enabled = True
                Exit Sub 'Sale del ciclo sin cerrar el rs para updt
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If ' Para el caso de la caja 0
    'Control de errores
    Exit Sub
errCtrl:
    Call Mensaje(Me, App.Title & " V. " & App.Major - App.Minor & "." & App.Revision, "Error: " & Err.Number & ": " & Err.Description)
End Sub
