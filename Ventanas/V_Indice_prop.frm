VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form V_Prop 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades del Proyecto"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6675
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton B_close 
      Caption         =   "Cerrar"
      Height          =   420
      Left            =   5520
      TabIndex        =   0
      Top             =   3480
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "V_Indice_prop.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lb_compilador"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label_Archivo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text_Coment"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text_Client"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text_Name"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Limites"
      TabPicture(1)   =   "V_Indice_prop.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_lra_lim_inf"
      Tab(1).Control(1)=   "txt_lra_lim_sup"
      Tab(1).Control(2)=   "txt_bit_lim_inf"
      Tab(1).Control(3)=   "txt_bit_lim_sup"
      Tab(1).Control(4)=   "Lb_Word"
      Tab(1).Control(5)=   "Label7"
      Tab(1).Control(6)=   "Label8"
      Tab(1).Control(7)=   "Label9"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Comunicación"
      TabPicture(2)   =   "V_Indice_prop.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label11"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lb_delay"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label10"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txt_delay_resp"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmb_puerto"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmb_velocidad"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Frame1"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      Begin VB.Frame Frame1 
         Caption         =   "RTS Controls"
         Height          =   1455
         Left            =   -74760
         TabIndex        =   26
         Top             =   1800
         Width           =   3255
         Begin VB.CheckBox chk_invertido 
            Alignment       =   1  'Right Justify
            Caption         =   "Invertido"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txt_rst_on 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1845
            MaxLength       =   3
            TabIndex        =   28
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt_rst_off 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1845
            MaxLength       =   3
            TabIndex        =   27
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label14 
            Caption         =   "Cseg"
            Height          =   255
            Left            =   2400
            TabIndex        =   34
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "Cseg"
            Height          =   255
            Left            =   2400
            TabIndex        =   33
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lb_rts_on 
            Caption         =   "Tiempo RTS ON"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label13 
            Caption         =   "Tiempo RTS OFF"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1080
            Width           =   1335
         End
      End
      Begin VB.ComboBox cmb_velocidad 
         Height          =   315
         ItemData        =   "V_Indice_prop.frx":0054
         Left            =   -73080
         List            =   "V_Indice_prop.frx":006D
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox cmb_puerto 
         Height          =   315
         ItemData        =   "V_Indice_prop.frx":009E
         Left            =   -74760
         List            =   "V_Indice_prop.frx":00A8
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txt_delay_resp 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   -73080
         MaxLength       =   3
         TabIndex        =   23
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt_lra_lim_inf 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   -73680
         MaxLength       =   4
         TabIndex        =   15
         Text            =   "0"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txt_lra_lim_sup 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   -71760
         MaxLength       =   4
         TabIndex        =   14
         Text            =   "448"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txt_bit_lim_inf 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   -73680
         MaxLength       =   4
         TabIndex        =   13
         Text            =   "449"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txt_bit_lim_sup 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   -71760
         MaxLength       =   4
         TabIndex        =   12
         Text            =   "512"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox Text_Name 
         Height          =   285
         Left            =   1350
         TabIndex        =   4
         Top             =   1110
         Width           =   5010
      End
      Begin VB.TextBox Text_Client 
         Height          =   285
         Left            =   1350
         TabIndex        =   3
         Top             =   1515
         Width           =   5010
      End
      Begin VB.TextBox Text_Coment 
         Height          =   285
         Left            =   1350
         TabIndex        =   2
         Top             =   1920
         Width           =   5010
      End
      Begin VB.Label Label10 
         Caption         =   " dseg"
         Height          =   255
         Left            =   -72480
         TabIndex        =   32
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lb_delay 
         Caption         =   "Retardo de respuesta"
         Height          =   255
         Left            =   -74760
         TabIndex        =   22
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   " Velocidad"
         Height          =   255
         Left            =   -73080
         TabIndex        =   21
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   " Puerto"
         Height          =   255
         Left            =   -74760
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Lb_Word 
         Caption         =   "Word  Minimo"
         Height          =   255
         Left            =   -74760
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Bitset Minimo"
         Height          =   255
         Left            =   -74760
         TabIndex        =   18
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Word  Maximo"
         Height          =   255
         Left            =   -72840
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Bitset Maximo"
         Height          =   255
         Left            =   -72840
         TabIndex        =   16
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre :"
         Height          =   285
         Left            =   135
         TabIndex        =   11
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente :"
         Height          =   285
         Left            =   135
         TabIndex        =   10
         Top             =   1515
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Comentario :"
         Height          =   285
         Left            =   135
         TabIndex        =   9
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Ruta :"
         Height          =   285
         Left            =   135
         TabIndex        =   8
         Top             =   705
         Width           =   1095
      End
      Begin VB.Label Label_Archivo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1350
         TabIndex        =   7
         Top             =   705
         Width           =   5010
      End
      Begin VB.Label lb_compilador 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1350
         TabIndex        =   6
         Top             =   2370
         Width           =   5010
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Compilador:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2370
         Width           =   1095
      End
   End
End
Attribute VB_Name = "V_Prop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private temp As V_Indice
Public mod_status As Boolean
Dim cambios As Boolean

Private Sub B_close_Click()
    If cambios Then
        temp.lra_limite_inf = Val(txt_lra_lim_inf.Text)
        temp.lra_limite_sup = Val(txt_lra_lim_sup.Text)
        temp.bit_limite_inf = Val(txt_bit_lim_inf.Text)
        temp.bit_limite_sup = Val(txt_bit_lim_sup.Text)
    End If
    Unload Me
End Sub

Private Sub chk_invertido_Click()
    mod_status = True
End Sub

Private Sub cmb_puerto_Change()
    mod_status = True
End Sub

Private Sub cmb_velocidad_Change()
    mod_status = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mod_status Then
        temp.Nombre = Text_Name.Text
        temp.Cliente = Text_Client.Text
        temp.Comentario = Text_Coment.Text
        temp.dataglobal.rts_off = CByte(txt_rst_off.Text)
        temp.dataglobal.rts_on = CByte(txt_rst_on.Text)
        temp.dataglobal.delay_resp = CByte(txt_delay_resp.Text)
        temp.dataglobal.invertido = CBool(chk_invertido.Value)
        If cmb_puerto.ListIndex = 0 Then
            temp.dataglobal.puerto = 1
        Else
            temp.dataglobal.puerto = 2
        End If
        Select Case cmb_velocidad.ListIndex
            Case 0
                temp.dataglobal.velocidad = 1200
            Case 1
                temp.dataglobal.velocidad = 2400
            Case 2
                temp.dataglobal.velocidad = 4800
            Case 3
                temp.dataglobal.velocidad = 9600
            Case 4
                temp.dataglobal.velocidad = 19200
            Case 5
                temp.dataglobal.velocidad = 38400
            Case 6
                temp.dataglobal.velocidad = 57600
        End Select
    End If
End Sub

Private Sub lb_compilador_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim nombre_comp As String
    Dim ret As Variant
    
    On Error GoTo fin
    If Shift = 2 Then
        EditorIDE.ComunWindows.Filter = "Compilador(*.exe )|*.exe"
        EditorIDE.ComunWindows.Flags = cdlOFNCreatePrompt
        EditorIDE.ComunWindows.CancelError = True
        EditorIDE.ComunWindows.ShowOpen
        
        If EditorIDE.ComunWindows.filename <> "" Then
            nombre_comp = Mid(EditorIDE.ComunWindows.filename, Len(EditorIDE.ComunWindows.filename) - 9, 10)
            If nombre_comp = NOMBRE_DEL_COMPILADOR Then
                lb_compilador.Caption = EditorIDE.ComunWindows.filename
                ret = puttofile(App.path + "\IDEPO900.INI", "SISTEMA", "COMPILADOR", lb_compilador.Caption)
                EditorIDE.path_compilador = lb_compilador.Caption
            End If
        End If
    End If
    Exit Sub
fin:
End Sub

Private Sub Text_Client_Change()
    mod_status = True
End Sub

Private Sub Text_Coment_Change()
    mod_status = True
End Sub

Private Sub Text_Name_Change()
    mod_status = True
End Sub

Public Sub Load(idprj As String)
    Set temp = Proyectos(idprj)
    Text_Name.Text = temp.Nombre
    Text_Client.Text = temp.Cliente
    Text_Coment.Text = temp.Comentario
    
    txt_lra_lim_inf.Text = CStr(temp.lra_limite_inf)
    txt_lra_lim_sup.Text = CStr(temp.lra_limite_sup)
    txt_bit_lim_inf.Text = CStr(temp.bit_limite_inf)
    txt_bit_lim_sup.Text = CStr(temp.bit_limite_sup)
    txt_delay_resp.Text = CStr(temp.dataglobal.delay_resp)
    txt_rst_on.Text = CStr(temp.dataglobal.rts_on)
    txt_rst_off.Text = CStr(temp.dataglobal.rts_off)
    
    If temp.dataglobal.invertido = True Then
        chk_invertido.Value = 1
    Else
        chk_invertido.Value = 0
    End If
    If temp.dataglobal.puerto = 1 Then
        cmb_puerto.ListIndex = 0
    Else
        cmb_puerto.ListIndex = 1
    End If
    Select Case temp.dataglobal.velocidad
        Case 1200
            cmb_velocidad.ListIndex = 0
        Case 2400
            cmb_velocidad.ListIndex = 1
        Case 4800
            cmb_velocidad.ListIndex = 2
        Case 9600
            cmb_velocidad.ListIndex = 3
        Case 19200
            cmb_velocidad.ListIndex = 4
        Case 38400
            cmb_velocidad.ListIndex = 5
        Case 57600
            cmb_velocidad.ListIndex = 6
    End Select
    
    texto = temp.archivo
    If InStr(texto, "Nuevo_") Then
        texto = temp.archivo
    Else
        texto = temp.archivo
        aux = InStrRev(texto, "\")
        texto = Mid(texto, aux)
    End If
    Label_Archivo.Caption = texto
    Label_Archivo.ToolTipText = temp.archivo
    mod_status = False
    lb_compilador.Caption = getoffile(App.path + "\IDEPO900.INI", "SISTEMA", "COMPILADOR", "none")
    Me.Show (1)
End Sub

Private Sub txt_bit_lim_inf_Change()
    cambios = True
End Sub

Private Sub txt_bit_lim_sup_Change()
    cambios = True
End Sub

Private Sub txt_delay_resp_Change()
    If Val(txt_delay_resp.Text) > 100 Then txt_delay_resp.Text = "100"
    If Val(txt_delay_resp.Text) < 0 Then txt_delay_resp.Text = "0"
    mod_status = True
End Sub

Private Sub txt_lra_lim_inf_Change()
   cambios = True
End Sub

Private Sub txt_lra_lim_sup_Change()
    cambios = True
End Sub

Private Sub txt_rst_off_Change()
    If Val(txt_rst_off.Text) > 100 Then txt_rst_off.Text = "100"
    If Val(txt_rst_off.Text) < 0 Then txt_rst_off.Text = "0"
    mod_status = True
End Sub

Private Sub txt_rst_on_Change()
    If Val(txt_rst_on.Text) > 100 Then txt_rst_on.Text = "100"
    If Val(txt_rst_on.Text) < 0 Then txt_rst_on.Text = "0"
   mod_status = True
End Sub
