Attribute VB_Name = "basIni"
Option Explicit
'Conexión general a la bd
Public cnn As ADODB.Connection
'Recordset principal
Public rs As ADODB.Recordset
'Cadena de consulta
Public strSql As String
'Cadena de conexión
Private strCnn As String
'Cadena que contiene el nombre del skin actual (a cargar)
Public strSkn As String

'Inicializa con los efectos de WxP
Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub main()
    'Inicializa los efectos de WxP
    InitCommonControls
    
    'Obtiene del registro el nombre del skin actual
    strSkn = GetSetting("FAEL", "Skin", "ActualSkn")
    'Si no hay datos del SKIN en el registro usa 'este' por defecto y lo graba en el Wreg32
    If strSkn = "" Then
        strSkn = "DiamanteGreen.skn"
        SaveSetting "FAEL", "Skin", "ActualSkn", strSkn
    End If
    mdiFAEL.Show
    frmSplash.Show
    'frmLogin.Show
End Sub

'Implementa un msgbox propio
Public Function Mensaje(ByVal FrmCont As Form, strTitulo As String, strMsg As String)
'On Error Resume Next
    frmMsgBox.Caption = strTitulo
    frmMsgBox.lblMensaje.Caption = strMsg
    frmMsgBox.Show vbModal ', FrmCont
    'DoEvents
End Function

'Consulta la bd
Public Sub Consulta(strSql As String, rsObj As Recordset)
    rs.Open strSql, cnn, adOpenDynamic, adLockOptimistic
End Sub

'Pocisiona el form en la mitad del mdi eje x (left)
Public Function PosX(frm As Form) As Integer
    Dim Ancho As Integer
    Ancho = frm.Width
    PosX = (mdiFAEL.Width - Ancho) / 2
End Function

'Pocisiona el form en la mitad del mdi eje y (top)
Public Function PosY(frm As Form) As Integer
    Dim Alto As Integer
    Alto = frm.Height
    PosY = (mdiFAEL.Height - Alto) / 4
End Function

'Aplica skin a los forms
Public Sub Aplicarskin(ByVal Formulario As Form)
    Formulario.Skin1.LoadSkin App.Path & "\mmedia\skins\" & strSkn 'Winter.skn" - default
    Formulario.Skin1.ApplySkin Formulario.hWnd
End Sub

'Conexión a la bd
Public Sub Conexion()
On Error Resume Next
    Set cnn = New ADODB.Connection
    If GetSetting("FAEL", "DefaultCnn", "blnFlag") = "True" Then
        strCnn = "Driver={" & GetSetting("FAEL", "DefaultCnn", "Driver") & _
        "};Port=" & GetSetting("FAEL", "DefaultCnn", "Port") & _
        ";Server=" & GetSetting("FAEL", "DefaultCnn", "Server") & _
        ";Database=" & GetSetting("FAEL", "DefaultCnn", "Database") & _
        ";Uid=" & GetSetting("FAEL", "DefaultCnn", "Uid") & _
        ";Pwd=" & OCT(GetSetting("FAEL", "DefaultCnn", "Pwd"), False)
    Else
        'frmConexion.Show
        strCnn = "Driver={" & frmConexion.cboProveedor & "};Port=" & _
        frmConexion.txtPuerto & ";Server=" & frmConexion.txtServidor & _
        ";Database=" & frmConexion.txtBasededatos & _
        ";Uid=" & frmConexion.txtUsuario & _
        ";Pwd=" & frmConexion.txtContrasena
    End If
    'El cursor de la conexión de lado del cliente
    cnn.CursorLocation = adUseClient
    cnn.Open strCnn
    If Err.Number = -2147467259 Then
        MsgBox "No se puede conectar con la bd, revise detenidamente el mensaje de error y verifque:" & vbCrLf & _
        "Nombre o IP del servidor, nombre de la bd, puerto, contraseña, driver, etc..." & vbCrLf & vbCrLf & _
        "Error: " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & "Configure correctamente los parámetros", _
        vbCritical, "Error fatal - " & App.Title & " V. " & App.Major - App.Minor & "." & App.Revision
        'End
    End If
    'Agrega datos al registro del sistema, si ya hay datos en el registro. _
    No hace nada
    If Err.Number = 0 Then
        If GetSetting("FAEL", "DefaultCnn", "blnFlag") <> "True" Then
            SaveSetting "FAEL", "DefaultCnn", "Driver", frmConexion.cboProveedor.Text
            SaveSetting "FAEL", "DefaultCnn", "Port", frmConexion.txtPuerto
            SaveSetting "FAEL", "DefaultCnn", "Server", frmConexion.txtServidor
            SaveSetting "FAEL", "DefaultCnn", "Database", frmConexion.txtBasededatos
            SaveSetting "FAEL", "DefaultCnn", "Uid", frmConexion.txtUsuario
            SaveSetting "FAEL", "DefaultCnn", "Pwd", OCT(frmConexion.txtContrasena, True)
            SaveSetting "FAEL", "DefaultCnn", "blnFlag", "True"
            'Descarga el formulario de parámetros de conexión
            Unload frmConexion
        End If
        'Establece una unica vez al cargar el programa el recordset
        Set rs = New Recordset
    Else
        frmConexion.Show vbModal
    End If
End Sub

'Encripta->TRUE y desencripta->FALSE by msfr_n
Public Function OCT(strCadena As String, CryptDecript As Boolean) As String
    Dim Encriptado As String
    Dim Clave As String
    Dim R As String
    Dim x As Integer
    Dim H As String
    If CryptDecript = True Then
        Encriptado = ""
        Clave = strCadena
        For x = 1 To Len(Clave)
            H = Asc(Mid$(Clave, x, 1))
            H = H + 32
            R = Chr$(H - 10)
            Encriptado = Encriptado + R
        Next x
        OCT = Encriptado
    Else
        Encriptado = ""
        Clave = strCadena
        For x = 1 To Len(Clave)
            H = Asc(Mid$(Clave, x, 1))
            H = H + 10
            H = H - 32
            R = Chr$(H)
            Encriptado = Encriptado + R
        Next x
        OCT = Encriptado
    End If
End Function
