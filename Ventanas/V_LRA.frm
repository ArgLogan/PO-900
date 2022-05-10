VERSION 5.00
Begin VB.Form V_LRA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingrese LRA"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   2430
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FRM_bits 
      Caption         =   "     Bits "
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   2175
      Begin VB.OptionButton Opt_clear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   600
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton opt_set 
         Caption         =   "Set"
         Height          =   195
         Left            =   1080
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txt_ini_bits 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         Height          =   270
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   9
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txt_Largo_bits 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lb_title_bits 
         Caption         =   ". Inicio   /   Largo"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   135
      End
      Begin VB.Label lb_sep_bits 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.CommandButton cmd_aceptar 
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2400
      Width           =   495
   End
   Begin VB.ComboBox cmb_modo 
      Height          =   315
      ItemData        =   "V_LRA.frx":0000
      Left            =   1200
      List            =   "V_LRA.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txt_direccion 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      Height          =   270
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "100"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txt_RTU 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      Height          =   255
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "1"
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Adress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "RTU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Protocolo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "V_LRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SEP_PROTO = "|"
Private Const SEP_RTU = ":"
Private Const SEP_DIREC = "."
Private Const SEP_BITS = "/"
Private Const SEP_BITSET = " "
Private tipo_lra As String
Private str_lra As String
Private limite_max As Integer
Private limite_min As Integer
Dim interno As Boolean
Public Function new_lra(ByVal lra As String, ByVal tipo As String, ByVal max As Integer, ByVal min As Integer) As String
    tipo_lra = tipo
    str_lra = lra
    limite_max = max
    limite_min = min
    Me.Show (1)
    new_lra = str_lra
End Function
Private Sub Check_Click()
        If Check.Value = 0 Then
            txt_ini_bits.Locked = True
            txt_ini_bits.Text = ""
            txt_Largo_bits.Text = ""
            txt_Largo_bits.Locked = True
        Else
            txt_ini_bits.Locked = False
            txt_ini_bits.Text = "0"
            txt_Largo_bits.Text = "1"
            txt_Largo_bits.Locked = False
        End If
End Sub

Private Sub cmd_aceptar_Click()
    Dim bit_set As String
    Dim aux As Byte
    Dim aux2 As Byte
    
    chek_limites
    
    If cmb_modo.ListIndex <> 2 Then
        If tipo_lra = "LRA" Or tipo_lra = "SYNC" Then
            If txt_ini_bits.Text <> "" And txt_Largo_bits.Text = "" Then
                txt_Largo_bits.Text = "1"
            End If
            If txt_ini_bits.Text = "" Then
                str_lra = cmb_modo.Text & SEP_PROTO & Val(txt_RTU.Text) & SEP_RTU & Val(txt_direccion.Text)
            Else
                aux = Val(txt_ini_bits.Text)
                aux2 = 16 - aux
                If aux2 = 0 Then
                    txt_Largo_bits.Text = "1"
                Else
                    If Val(txt_Largo_bits.Text) < 1 Then txt_Largo_bits.Text = "1"
                    If Val(txt_Largo_bits.Text) > aux2 Then txt_Largo_bits.Text = Str(aux2)
                End If
                str_lra = cmb_modo.Text & SEP_PROTO & Val(txt_RTU.Text) & SEP_RTU & Val(txt_direccion.Text) & SEP_DIREC & Val(txt_ini_bits.Text) & SEP_BITS & Val(txt_Largo_bits.Text)
            End If
        Else
            If txt_ini_bits.Text = "" Then txt_ini_bits.Text = "0"
            If opt_set = True Then
                bit_set = "1"
            Else
                bit_set = "0"
            End If
            str_lra = cmb_modo.Text & SEP_PROTO & Val(txt_RTU.Text) & SEP_RTU & Val(txt_direccion.Text) & SEP_DIREC & Val(txt_ini_bits.Text) & SEP_BITSET & bit_set
        End If
    Else
        str_lra = "OFF" 'ahora no deberia entrar nunca aca
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i_aux As Integer
    Dim i_aux2 As Integer
    Dim s_aux As String
    
    txt_direccion.Text = CStr(limite_min)
    If tipo_lra = "LRA" Or tipo_lra = "SYNC" Then
        lb_sep_bits.Visible = True
        lb_title_bits.Caption = ". Inicio   /   Largo"
        txt_Largo_bits.Visible = True
        Opt_clear.Visible = False
        opt_set.Visible = False
    Else
        lb_sep_bits.Visible = False
        lb_title_bits.Caption = ". BIT"
        txt_Largo_bits.Visible = False
        Opt_clear.Visible = True
        opt_set.Visible = True
        interno = True
        Check.Value = 1
        interno = False
        Check.Enabled = False
        cmb_modo.ListIndex = 0
    End If

    If str_lra <> "" Then
        If str_lra = "OFF" Then
            cmb_modo.ListIndex = 0 'estaba en 2
        Else
            If tipo_lra = "LRA" Or tipo_lra = "SYNC" Then
                s_aux = Mid(str_lra, 1, 3)
                If s_aux = "SIM" Then
                    cmb_modo.ListIndex = 0
                Else
                    cmb_modo.ListIndex = 1
                End If
                i_aux = InStr(1, str_lra, SEP_RTU)
                s_aux = Mid(str_lra, 5, (i_aux - 5))
                txt_RTU.Text = s_aux
                i_aux2 = i_aux + 1
                i_aux = InStr(1, str_lra, SEP_DIREC)
                If i_aux <> 0 Then
                    interno = True
                    Check.Value = 1
                    interno = False
                    s_aux = Mid(str_lra, i_aux2, (i_aux - i_aux2))
                    txt_direccion.Text = s_aux
                    i_aux2 = i_aux + 1
                    i_aux = InStr(1, str_lra, SEP_BITS)
                    s_aux = Mid(str_lra, i_aux2, (i_aux - i_aux2))
                    txt_ini_bits.Text = s_aux
                    s_aux = Mid(str_lra, (i_aux + 1), 2)
                    txt_Largo_bits.Text = s_aux
                    txt_Largo_bits.Enabled = True
                Else
                    s_aux = Mid(str_lra, i_aux2, 4)
                    txt_direccion.Text = s_aux
                End If
                If tipo_lra = "SYNC" Then
                    Check.Value = 0
                    Check.Enabled = False
                    txt_ini_bits.Enabled = False
                    txt_Largo_bits.Enabled = False
                End If
            Else
                s_aux = Mid(str_lra, 1, 3)
                If s_aux = "SIM" Then
                    cmb_modo.ListIndex = 0
                Else
                    cmb_modo.ListIndex = 1
                End If
                i_aux = InStr(1, str_lra, SEP_RTU)
                s_aux = Mid(str_lra, 5, (i_aux - 5))
                txt_RTU.Text = s_aux
                i_aux2 = i_aux + 1
                i_aux = InStr(1, str_lra, SEP_DIREC)
                s_aux = Mid(str_lra, i_aux2, (i_aux - i_aux2))
                txt_direccion.Text = s_aux
                i_aux2 = i_aux + 1
                i_aux = InStr(1, str_lra, SEP_BITSET)
                s_aux = Mid(str_lra, i_aux2, (i_aux - i_aux2))
                txt_ini_bits.Text = s_aux
                s_aux = Mid(str_lra, (i_aux + 1), 2)
                If s_aux = "0" Then
                    Opt_clear = True
                Else
                    opt_set = True
                End If
            End If
        End If
    End If
End Sub

Private Function chek_limites()
    If Val(txt_direccion.Text) < limite_min Then
        txt_direccion.Text = CStr(limite_min)
    End If
    If Val(txt_direccion.Text) > limite_max Then
        txt_direccion.Text = CStr(limite_max)
    End If
End Function
Private Sub txt_direccion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmd_aceptar.SetFocus
End Sub

Private Sub txt_direccion_LostFocus()
    If Val(txt_direccion.Text) > 9999 Then txt_direccion.Text = "9999"
    If Val(txt_direccion.Text) < 0 Then txt_direccion.Text = "0"
End Sub

Private Sub txt_ini_bits_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmd_aceptar.SetFocus
End Sub

Private Sub txt_ini_bits_LostFocus()
    Dim aux As Byte
    Dim aux2 As Byte
    
    If Val(txt_ini_bits.Text) > 15 Then txt_ini_bits = "15"
    If Val(txt_ini_bits.Text) < 0 Then txt_ini_bits.Text = "0"
    
    aux = Val(txt_ini_bits.Text)
    aux2 = 16 - aux
    If aux2 = 0 Then
        txt_Largo_bits.Text = "1"
    Else
        If Val(txt_Largo_bits.Text) < 1 Then txt_Largo_bits.Text = "1"
        If Val(txt_Largo_bits.Text) > aux2 Then txt_Largo_bits.Text = CStr(aux2)
    End If
        
End Sub

Private Sub txt_Largo_bits_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmd_aceptar.SetFocus
End Sub

Private Sub txt_Largo_bits_LostFocus()
    Dim aux As Byte
    Dim aux2 As Byte
    
    aux = Val(txt_ini_bits.Text)
    aux2 = 16 - aux
    If aux2 = 0 Then
        txt_Largo_bits.Text = "1"
    Else
        If Val(txt_Largo_bits.Text) < 1 Then txt_Largo_bits.Text = "1"
        If Val(txt_Largo_bits.Text) > aux2 Then txt_Largo_bits.Text = CStr(aux2)
    End If
End Sub

Private Sub txt_RTU_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmd_aceptar.SetFocus
End Sub

Private Sub txt_RTU_LostFocus()
    If Val(txt_RTU.Text) > 128 Then txt_RTU.Text = "128"
    If Val(txt_RTU.Text) < 1 Then txt_RTU.Text = "1"
End Sub
