VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm mdiFAEL 
   BackColor       =   &H8000000C&
   Caption         =   "..."
   ClientHeight    =   6300
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9615
   Icon            =   "mdiFAEL.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFAEL.frx":0CCA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbrFAEL 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   5985
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1800
      OleObjectBlob   =   "mdiFAEL.frx":6F64
      Top             =   3720
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "&Admin"
      Begin VB.Menu mnuEmpresa 
         Caption         =   "&Empresa"
      End
      Begin VB.Menu Linea0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUsuarios 
         Caption         =   "&Usuarios"
      End
      Begin VB.Menu mnuPermisos 
         Caption         =   "&Permisos"
      End
      Begin VB.Menu Linea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCerrarsesion 
         Caption         =   "&Cerrar Sesión"
         Shortcut        =   ^L
      End
      Begin VB.Menu Linea2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuInventario 
      Caption         =   "&Inventario"
      Begin VB.Menu mnuCompra 
         Caption         =   "&Compra"
      End
      Begin VB.Menu mnuNotadebito 
         Caption         =   "Nota &Débito"
      End
      Begin VB.Menu mnuRevisión 
         Caption         =   "&Remisión"
      End
      Begin VB.Menu mnuFactura 
         Caption         =   "&Factura"
      End
      Begin VB.Menu mnuNotacredito 
         Caption         =   "Nota &Crédito"
      End
      Begin VB.Menu mnuCotizacion 
         Caption         =   "Coti&zación"
      End
      Begin VB.Menu mnuEntrada 
         Caption         =   "&Entrada"
      End
   End
   Begin VB.Menu mnuCartera 
      Caption         =   "&Cartera"
      Begin VB.Menu mnuClienteCar 
         Caption         =   "&Cliente"
      End
      Begin VB.Menu mnuFacturaCar 
         Caption         =   "&Factura"
      End
      Begin VB.Menu Linea3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInformesCar 
         Caption         =   "&Informes"
      End
      Begin VB.Menu mnuRevisaCar 
         Caption         =   "&Revisa"
      End
   End
   Begin VB.Menu mnuReportes 
      Caption         =   "&Reportes"
      Begin VB.Menu mnuClientesRep 
         Caption         =   "&Clientes"
      End
      Begin VB.Menu mnuProveedoresRep 
         Caption         =   "&Proveedores"
      End
      Begin VB.Menu mnuProductos 
         Caption         =   "Pr&oductos"
      End
   End
   Begin VB.Menu mnuHerramientas 
      Caption         =   "&Herramientas"
      Begin VB.Menu Linea4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpciones 
         Caption         =   "&Opciones..."
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ay&uda"
      Begin VB.Menu mnuContenido 
         Caption         =   "&Contenido"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAcercade 
         Caption         =   "&Acerca de ..."
      End
   End
End
Attribute VB_Name = "mdiFAEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Controla el cierre de la aplicación
Dim blnSesion As Boolean
Dim bLoop As Byte

Private Sub MDIForm_Activate()
    Aplicarskin Me
End Sub

Private Sub MDIForm_Load()
    
    Call Conexion
    Aplicarskin Me
    Me.Caption = App.Title & " - " & App.Major & "." & App.Minor & "." & App.Revision
    blnSesion = False
    'INI - Barra de estado ------------------------------------------------------
    Dim pnlX As Panel
    For bLoop = 1 To 5
        Set pnlX = sbrFAEL.Panels.Add()
        sbrFAEL.Panels(bLoop).AutoSize = sbrContents
    Next bLoop
    'Modifica el primer panel de la sBar - Usuario.
    sbrFAEL.Panels(1).Picture = LoadPicture(App.Path & "\mmedia\usuario.gif")
    sbrFAEL.Panels(1).ToolTipText = "Usuario"
    sbrFAEL.Panels(1).Text = "INICIO DE SESION"
    'Muestra en la sBar la img Tema
    sbrFAEL.Panels(2).Picture = LoadPicture(App.Path & "\mmedia\tema.gif")
    sbrFAEL.Panels(2).ToolTipText = "Tema actual ó skin de la aplicación"
    'Muestra el nombre del skin sin extensión
    sbrFAEL.Panels(2).Text = Left(strSkn, Len(strSkn) - 4)
    sbrFAEL.Panels(3).Picture = LoadPicture(App.Path & "\mmedia\fecha.gif")
    sbrFAEL.Panels(3).ToolTipText = "Fecha de ingreso al sistema"
    'Fecha del sistema para la barra de herramientas
    sbrFAEL.Panels(3).Text = Format$(Now, "dddd,") & Format$(Now, " dd") & _
    " de " & Format$(Now, "mmmm") & " de " & Format$(Now, "yyyy") & _
    " - (" & Format$(Now, "dd/mm/yy") & ")" & " - " & Format$(Now, "hh:mm AMPM")
    'sbrFAEL.Panels(3).Style = sbrDate
    sbrFAEL.Panels(4).Style = sbrCaps
    sbrFAEL.Panels(5).Style = sbrNum 'sbrIns
    sbrFAEL.Panels(6).Style = sbrIns
    'FIN - Barra de estado ------------------------------------------------------
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    cnn.Close
    If blnSesion = False Then End
    blnSesion = True
End Sub

Private Sub mnuAcercade_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuCerrarsesion_Click()
    
    blnSesion = True
    Unload Me
    Load Me
    Me.Show
    Unload frmLogin
    frmLogin.Show
    frmLogin.Refresh
    
End Sub

Private Sub mnuEmpresa_Click()
    frmEmpresa.Show
End Sub

Private Sub mnuOpciones_Click()
    frmTema.Show vbModal, mdiFAEL
End Sub

Private Sub mnuPermisos_Click()
    frmPermisos.Show
End Sub

Private Sub mnuSalir_Click()
    End
End Sub

Private Sub mnuUsuarios_Click()
    frmUsuarios.Show
End Sub
