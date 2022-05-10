VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm EditorIDE 
   BackColor       =   &H8000000C&
   Caption         =   "PO900 Win Editor"
   ClientHeight    =   9720
   ClientLeft      =   4950
   ClientTop       =   4005
   ClientWidth     =   13770
   Icon            =   "IDE.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer_nav 
      Left            =   5880
      Top             =   720
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   5040
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDE.frx":0ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDE.frx":1464
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDE.frx":19FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDE.frx":2850
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDE.frx":455A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDE.frx":655C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDE.frx":8266
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDE.frx":9F70
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDE.frx":BC7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDE.frx":D984
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDE.frx":F68E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDE.frx":11398
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13770
      _ExtentX        =   24289
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo Proyecto"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Abrir Proyecto"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Guardar proyecto"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Propiedades del proyecto"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ventana de Campos"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ventana de Propiedades"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   7
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Compilar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Globales"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ayuda"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir del Sistema"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog ComunWindows 
      Left            =   4440
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu GrupoArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu PRJ_NUEVO 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu PRJ_ABRIR 
         Caption         =   "Abrir ..."
      End
      Begin VB.Menu PRJ_GUARDAR 
         Caption         =   "Guardar"
      End
      Begin VB.Menu PRJ_GUARDAR_COMO 
         Caption         =   "Guardar como ..."
      End
      Begin VB.Menu PRJ_PROP 
         Caption         =   "Propiedades"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu PRJ_QUIT 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Views 
      Caption         =   "Ver"
      Begin VB.Menu VW_BarraDeCampos 
         Caption         =   "Ventana de Campos"
      End
      Begin VB.Menu VW_VentanaDePropiedades 
         Caption         =   "Ventana de Propiedades"
      End
      Begin VB.Menu visible_toolbar 
         Caption         =   "Barra de Herramientas"
         Checked         =   -1  'True
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu VW_VentanaDeProyecto 
         Caption         =   "Ventana de Proyecto"
      End
   End
   Begin VB.Menu TOOLS 
      Caption         =   "Herramientas"
      Begin VB.Menu TLS_Compilar 
         Caption         =   "Compilar"
      End
      Begin VB.Menu List_lra 
         Caption         =   "Lista LRA Address"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu TLS_Globales 
         Caption         =   "Globales"
      End
   End
   Begin VB.Menu HELP 
      Caption         =   "Ayuda"
      Begin VB.Menu HLP_AcercaDe 
         Caption         =   "Acerca de ..."
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu HLP_Indice 
         Caption         =   "Indice"
      End
      Begin VB.Menu HLP_Temas 
         Caption         =   "Temas"
      End
   End
End
Attribute VB_Name = "EditorIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public path_compilador As String
Dim file_output_bin As String

Private Sub HLP_AcercaDe_Click()
    IdeAcercaDe.Show (1)
End Sub

Private Sub List_lra_Click()
    Dim temp As V_Indice
    Dim idprj As String
    Dim comando As String
    Dim aux As String
    Dim ret
    
    If Proyectos.count = 1 Then
        Set temp = Proyectos(1)
        idprj = temp.archivo
    ElseIf Proyectos.count > 0 Then
        idprj = V_Selec_proyect.Load
        If idprj <> "" Then
            Set temp = Proyectos(idprj)
        End If
    End If
    If idprj <> "" Then
        temp.save_proyecto
        aux = Mid(idprj, Len(temp.path) + 1, Len(idprj))
        If Dir(App.path & "\plugins\Lista_lra.exe") <> "" Then
            comando = App.path & "\plugins\Lista_lra.exe " & idprj & ";" & aux
            ret = Shell(comando, vbNormalFocus)
        End If
    End If
End Sub

Private Sub MDIForm_Load()
    Set Proyectos = New Collection
    Proyectos_Open_New = 0
    PRJ_GUARDAR.Enabled = False
    Toolbar.Buttons(3).Enabled = False
    PRJ_GUARDAR_COMO.Enabled = False
    PRJ_PROP.Enabled = False
    Toolbar.Buttons(4).Enabled = False
    List_lra.Enabled = False
    
    path_compilador = getoffile(App.path + "\IDEPO900.INI", "SISTEMA", "COMPILADOR", "")
    If Dir(path_compilador) <> "" Then
        TLS_Compilar.Enabled = True
        Toolbar.Buttons(9).Enabled = True
    Else
        TLS_Compilar.Enabled = False
        Toolbar.Buttons(9).Enabled = False
        path_compilador = App.path & "\plugins\CpOpm1.exe"
    End If
    
    file_output_bin = getoffile(App.path + "\IDEPO900.INI", "SISTEMA", "OUTPUT_BIN", "\BINARIO\")
    file_output_bin = file_output_bin + getoffile(App.path + "\IDEPO900.INI", "SISTEMA", "BINARIO", "PcodeD.bin")
    
    
'*******************************************************************************************************************
'****************************************** VERSION DEL PROGRAMA ***************************************************
'*******************************************************************************************************************
    Me.Caption = "PO900 Win Editor  Versión " & App.Major & "." & App.Minor & "." & App.Revision
'*******************************************************************************************************************
'*******************************************************************************************************************
'*******************************************************************************************************************
End Sub

Private Sub PRJ_ABRIR_Click()
    Dim file_path As String
    Dim file_name As String
    Dim Test As V_Indice
    
    On Error GoTo finError
    
    ComunWindows.Filter = "Proyecto Panel Operador PO900(*.ppo )|*.ppo"
    ComunWindows.Flags = cdlOFNCreatePrompt
    ComunWindows.CancelError = True
    ComunWindows.ShowOpen
    
    If ComunWindows.filename <> "" Then
        file_path = ComunWindows.filename
        open_project (file_path)
        PRJ_GUARDAR.Enabled = True
        If Dir(App.path & "\plugins\Lista_lra.exe") <> "" Then List_lra.Enabled = True
        Toolbar.Buttons(3).Enabled = True
        PRJ_GUARDAR_COMO.Enabled = True
        PRJ_PROP.Enabled = True
        Toolbar.Buttons(4).Enabled = True
    End If
    
    Exit Sub
finError:
    If Err = &H7FF3 Then
        ComunWindows.filename = ""
    End If
    Resume Next
End Sub

Private Sub PRJ_GUARDAR_Click()
    Dim temp As V_Indice
    Dim idprj As String
    If Proyectos.count = 1 Then
        Set temp = Proyectos(1)
        idprj = temp.archivo
    Else
        idprj = V_Selec_proyect.Load
        If idprj <> "" Then
            Set temp = Proyectos(idprj)
        End If
    End If
    If idprj <> "" Then
        temp.save_proyecto
    End If
End Sub

Private Sub PRJ_GUARDAR_COMO_Click()
    Dim temp As V_Indice
    Dim idprj As String
    If Proyectos.count = 1 Then
        Set temp = Proyectos(1)
        idprj = temp.archivo
    Else
        idprj = V_Selec_proyect.Load
        If idprj <> "" Then
            Set temp = Proyectos(idprj)
        End If
    End If
    If idprj <> "" Then
        temp.save_proyecto (True)
        EditorIDE.PRJ_GUARDAR.Enabled = True
        EditorIDE.Toolbar.Buttons(3).Enabled = True
    End If
End Sub

Private Sub PRJ_NUEVO_Click()
    Proyectos_Open_New = Proyectos_Open_New + 1
    open_project ("Nuevo_" + Hex(Proyectos_Open_New))
    PRJ_GUARDAR_COMO.Enabled = True
    PRJ_PROP.Enabled = True
    Toolbar.Buttons(4).Enabled = True
End Sub

Private Sub PRJ_PROP_Click()
    Dim temp As V_Indice
    Dim idprj As String
    If Proyectos.count = 1 Then
       Set temp = Proyectos(1)
       idprj = temp.archivo
    Else
       idprj = V_Selec_proyect.Load
    End If
    If idprj <> "" Then
       V_Prop.Load (idprj)
    End If
End Sub

Private Sub PRJ_QUIT_Click()
    Unload Me
End Sub

Private Sub Timer_nav_Timer()
    navega.Visible = False
    Timer_nav.Interval = 0
End Sub

Private Sub TLS_Compilar_Click()
    Dim comando As String
    Dim temp As V_Indice
    Dim idprj As String
    Dim path_proyecto As String
    If Proyectos.count > 0 Then
        If Proyectos.count = 1 Then
            Set temp = Proyectos(1)
            idprj = temp.archivo
        Else
            idprj = V_Selec_proyect.Load
        End If
        If idprj <> "" Then
            Set temp = Proyectos.item(idprj)
            path_proyecto = temp.path
            comando = path_compilador + " " + path_proyecto + "SCRNS\" + " " + path_proyecto + file_output_bin
            Shell (comando)
        End If
    End If
    
End Sub

Private Sub TLS_Globales_Click()
    Dim temp As V_Indice
    Dim idprj As String
    If Proyectos.count > 0 Then
        If Proyectos.count = 1 Then
            Set temp = Proyectos(1)
            idprj = temp.archivo
        Else
            idprj = V_Selec_proyect.Load
        End If
        If idprj <> "" Then
            V_Prop_P_G.viewglobals (idprj)
        End If
    End If
End Sub

Private Sub TOOLS_Click()
    If Dir(App.path & "\plugins\Lista_lra.exe") <> "" Then
        List_lra.Enabled = True
    Else
        List_lra.Enabled = False
    End If
    If Dir(path_compilador) <> "" Then
        TLS_Compilar.Enabled = True
    Else
        TLS_Compilar.Enabled = False
    End If
End Sub

Private Sub visible_toolbar_Click()
    If visible_toolbar.Checked = True Then
        visible_toolbar.Checked = False
        Toolbar.Visible = False
    Else
        visible_toolbar.Checked = True
        Toolbar.Visible = True
    End If
End Sub

Private Sub VW_BarraDeCampos_Click()
    V_Tools.Show
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim aux As Long
    Select Case Button.Index
        Case 1
            PRJ_NUEVO_Click
        Case 2
            PRJ_ABRIR_Click
        Case 3
            PRJ_GUARDAR_Click
        Case 4
            PRJ_PROP_Click
        Case 6
            VW_BarraDeCampos_Click
        Case 7
            VW_VentanaDePropiedades_Click
        Case 9
            TLS_Compilar_Click
        Case 10
            TLS_Globales_Click
        Case 14
            aux = MsgBox("¿Esta seguro de querer salir del sistema?", vbYesNo)
            If aux = vbYes Then
                Unload Me
            End If
    End Select

End Sub

Private Sub VW_VentanaDePropiedades_Click()
    properties.Show
End Sub

