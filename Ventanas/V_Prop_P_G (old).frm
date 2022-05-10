VERSION 5.00
Begin VB.Form V_Prop_P_G 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd_as_globales 
      Caption         =   "Default"
      Height          =   375
      Left            =   5760
      TabIndex        =   65
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame DataScreen 
      Caption         =   "Pantalla"
      Height          =   1095
      Left            =   120
      TabIndex        =   53
      Top             =   1920
      Width           =   6735
      Begin VB.TextBox NombreScreen 
         Height          =   285
         Left            =   1920
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label9 
         Caption         =   "Nombre de la Pantalla:"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame SYNC_GLOBAL 
      Caption         =   "Globales "
      Height          =   1095
      Left            =   120
      TabIndex        =   26
      Top             =   1920
      Width           =   6735
      Begin VB.TextBox Txt_pant_principal 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   57
         Text            =   "10"
         Top             =   700
         Width           =   375
      End
      Begin VB.TextBox Text_syncTime 
         Height          =   285
         Left            =   5640
         TabIndex        =   29
         Text            =   "50"
         Top             =   340
         Width           =   495
      End
      Begin VB.TextBox Text_syncLRA 
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "SIM|1:100"
         Top             =   340
         Width           =   1215
      End
      Begin VB.CheckBox sync_enable 
         Alignment       =   1  'Right Justify
         Caption         =   "Sync PLC Enable"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lb_pp 
         Caption         =   "Pantalla Principal"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Seg"
         Height          =   255
         Left            =   6120
         TabIndex        =   36
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Sync Time"
         Height          =   255
         Left            =   4800
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Sync LRA"
         Height          =   255
         Left            =   2280
         TabIndex        =   30
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Variables_Pantalla 
      Caption         =   "Variables "
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   5535
      Begin VB.TextBox TextJmpTimeOut 
         Height          =   285
         Left            =   4320
         TabIndex        =   52
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TextJmpEsc 
         Height          =   285
         Left            =   4320
         TabIndex        =   51
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text_TimeAutoCursor 
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
         Left            =   4320
         TabIndex        =   19
         Text            =   "2"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox Text_TimeEdition 
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
         Left            =   1440
         TabIndex        =   18
         Text            =   "30"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox Text_TimeOUT 
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
         Left            =   1440
         TabIndex        =   17
         Text            =   "60"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Text_TimeScan 
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
         Left            =   1440
         TabIndex        =   16
         Text            =   "100"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Seg."
         Height          =   255
         Left            =   4920
         TabIndex        =   35
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Seg."
         Height          =   255
         Left            =   2040
         TabIndex        =   34
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "mSeg."
         Height          =   255
         Left            =   2040
         TabIndex        =   33
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Seg."
         Height          =   255
         Left            =   2040
         TabIndex        =   32
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Tiempo AutoCursor"
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   25
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Salto de TimeOut"
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   24
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Salto de Escape"
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Tiempo Edición"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "TimeOut Pantalla"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Tiempo de Scan"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton Cerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   5760
      TabIndex        =   14
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame TextJmpTextJmp 
      Caption         =   "Definición de Comandos de Teclado "
      Height          =   2850
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   6720
      Begin VB.Frame Frm_teclas 
         Caption         =   "Tecla: "
         Enabled         =   0   'False
         Height          =   2295
         Left            =   4680
         TabIndex        =   59
         Top             =   360
         Width           =   1935
         Begin VB.TextBox txt_salto 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   840
            TabIndex        =   63
            Text            =   "FF"
            Top             =   1770
            Width           =   375
         End
         Begin VB.OptionButton OP_JUMP 
            Caption         =   "JUMP"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   1320
            Width           =   1095
         End
         Begin VB.OptionButton OP_BIT 
            Caption         =   "BIT"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton OP_NONE 
            Caption         =   "NONE"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "Saltar a"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   1800
            Width           =   615
         End
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   12
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   13
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   11
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   10
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   9
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   8
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   7
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   6
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   5
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla F4"
         Height          =   255
         Index           =   13
         Left            =   3480
         TabIndex        =   58
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla F3"
         Height          =   255
         Index           =   12
         Left            =   3480
         TabIndex        =   13
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla F2"
         Height          =   255
         Index           =   11
         Left            =   3480
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla F1"
         Height          =   255
         Index           =   10
         Left            =   3480
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 9"
         Height          =   255
         Index           =   9
         Left            =   2400
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 8"
         Height          =   255
         Index           =   8
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 7"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 6"
         Height          =   255
         Index           =   6
         Left            =   2400
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 5"
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 4"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 3"
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   4
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 2 "
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   3
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 1"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 0 "
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   1
         Top             =   2160
         Width           =   855
      End
   End
End
Attribute VB_Name = "V_Prop_P_G"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ptr_proyect As V_Indice
Dim modificado As Boolean
Dim globaldata As Variant
Dim screendata As Variant
Dim tecla_en_foco As Byte

Public Function viewglobals(idprj As String) As Boolean
    Dim indice As Byte
    Dim j As Byte
    Set ptr_proyect = Proyectos.item(idprj)
    Set globaldata = ptr_proyect.dataglobal
         
    Me.Caption = Me.Caption + " / " + ptr_proyect.Nombre
    SYNC_GLOBAL.Visible = True
    DataScreen.Visible = False
    
    Text_TimeScan.Text = Str(ptr_proyect.dataglobal.tdisplay)
    Text_TimeOUT.Text = Str(globaldata.timeout)
    Text_TimeEdition.Text = Str(globaldata.timeoutedit)
    Text_TimeAutoCursor.Text = Str(globaldata.timeautocursor)
    
    Txt_pant_principal.Text = Hex(globaldata.Pant_principal) 'verifiacar
    TextJmpEsc.Text = Hex(globaldata.escjmp)
    TextJmpTimeOut.Text = Hex(globaldata.timeoutjmp)
    
    If ptr_proyect.dataglobal.syncactivo Then
        sync_enable.Value = 1
    Else
        sync_enable.Value = 0
    End If
    Text_syncLRA.Text = globaldata.synclra
    Text_syncTime.Text = CStr(globaldata.synctime)
    
    For indice = 0 To 13
        If globaldata.keyjmpenable(indice) Then
            TextJmp(indice).Text = "JUMP " & Hex(globaldata.keyjmp(indice))
        Else
            If globaldata.Key_LRA(indice) <> "OFF" Then
                TextJmp(indice).Text = "BITSET"
            Else
                TextJmp(indice).Text = ""
            End If
        End If
    Next indice
    
    modificado = False
    Cmd_as_globales.Visible = True
    Me.Show (1)
    
    viewglobals = True
End Function

Public Function viewlocals(idprj As String, idscreen As Byte) As Boolean
    Dim indice As Byte
    Dim j As Byte
    Set ptr_proyect = Proyectos.item(idprj)
    
    Set globaldata = ptr_proyect.m_Screens.item(genidpan(idscreen))
    

    Me.Caption = Me.Caption + " / " + ptr_proyect.Nombre + " / " + globaldata.name
    
    SYNC_GLOBAL.Visible = False
    DataScreen.Visible = True
    
    NombreScreen.Text = globaldata.name
    Text_TimeScan.Text = Str(ptr_proyect.dataglobal.tdisplay)
    Text_TimeOUT.Text = Str(globaldata.timeout)
    Text_TimeEdition.Text = Str(globaldata.timeoutedit)
    Text_TimeAutoCursor.Text = Str(globaldata.timeautocursor)
    
    TextJmpEsc.Text = Hex(globaldata.escjmp)
    TextJmpTimeOut.Text = Hex(globaldata.timeoutjmp)
    
    For indice = 0 To 13
        If globaldata.keyjmpenable(indice) Then
            TextJmp(indice).Text = "JUMP " & Hex(globaldata.keyjmp(indice))
        End If
        If globaldata.Key_LRA(indice) <> "OFF" Then
            TextJmp(indice).Text = "BITSET"
        Else
            TextJmp(indice).Text = ""
        End If
    Next indice
    
    modificado = False
    
    Me.Show (1)
    
    viewlocals = True

End Function

Private Sub Cerrar_Click()
    Dim indice As Byte
    Dim auxntecla As String
    If modificado Then
        If IsNumeric(Text_TimeScan.Text) Then globaldata.tdisplay = Val(Text_TimeScan.Text)
        If IsNumeric(Text_TimeEdition.Text) Then globaldata.timeoutedit = Val(Text_TimeEdition.Text)
        If IsNumeric(Text_TimeAutoCursor.Text) Then globaldata.timeautocursor = Val(Text_TimeAutoCursor.Text)
    
        If TextJmpEsc.Text <> "" Then
            If Val("&H" + TextJmpEsc.Text) < 255 Then
                globaldata.escjmp = Val("&H" + TextJmpEsc.Text)
            Else
                globaldata.escjmp = &HFF
            End If
        Else
            globaldata.escjmp = &HFF
        End If
        
        
        If TextJmpTimeOut.Text <> "" Then
            If Val("&H" + TextJmpTimeOut.Text) < 254 Then
                globaldata.timeoutjmp = Val("&H" + TextJmpTimeOut.Text)
                If IsNumeric(Text_TimeOUT.Text) Then globaldata.timeout = Val(Text_TimeOUT.Text)
            Else
                globaldata.timeout = 0
                globaldata.timeoutjmp = &HFF
            End If
        Else
            globaldata.timeout = 0
            globaldata.timeoutjmp = &HFF
        End If
        
        If SYNC_GLOBAL.Visible = True Then
            globaldata.syncactivo = sync_enable.Value
            If Text_syncLRA.Text <> "" Then globaldata.synclra = Text_syncLRA.Text
            If IsNumeric(Text_syncTime.Text) Then globaldata.synctime = Val(Text_syncTime.Text)
            If Txt_pant_principal.Text <> "" Then
            If Val("&H" & Txt_pant_principal) < 255 Then
                globaldata.Pant_principal = Val("&H" & Txt_pant_principal)
            Else
                globaldata.Pant_principal = &HFE
            End If
        Else
            globaldata.Pant_principal = &H10
        End If

        Else
            globaldata.name = NombreScreen.Text
            If globaldata.modo <> SC_BORRAR Then
                globaldata.modo = SC_MODIFICADO
            End If
        End If
        
        For indice = 0 To 13
            If TextJmp(indice).Text <> "" And TextJmp(indice).Text <> "BITSET" Then
                auxntecla = Mid(TextJmp(indice).Text, 6, 2)
                If Val("&H" + auxntecla) < 255 Then
                    globaldata.keyjmpenable(indice) = True
                    globaldata.keyjmp(indice) = Val("&H" + auxntecla)
                    globaldata.Key_LRA(indice) = "OFF"
                End If
            Else
                globaldata.keyjmpenable(indice) = False
                If TextJmp(indice).Text = "" Then globaldata.Key_LRA(indice) = "OFF"
            End If
        Next indice
        ptr_proyect.m_prj_mod = True
    End If
    
    Unload Me
End Sub

Private Sub Cmd_as_globales_Click()
Dim indice As Byte

Set screendata = ptr_proyect.m_Screens

End Sub

Private Sub OP_BIT_GotFocus()
    txt_salto.Text = ""
    TextJmp(tecla_en_foco).Text = "BITSET"
    If OP_BIT.Value = False Then
        OP_BIT.Value = True
        TextJmp_DblClick (tecla_en_foco)
    End If
End Sub

Private Sub OP_JUMP_Click()
    TextJmp(tecla_en_foco).Text = "JUMP " & txt_salto.Text
    txt_salto.Enabled = True
    txt_salto.Text = Txt_pant_principal.Text
End Sub

Private Sub OP_NONE_Click()
    txt_salto.Text = ""
    TextJmp(tecla_en_foco).Text = ""
    txt_salto.Enabled = False
End Sub

Private Sub TextJmp_DblClick(Index As Integer)
    Dim ant As String
    Dim j As Byte
    If OP_BIT.Value = True Then
        Select Case Index
            Case 0
               ant = globaldata.Key_LRA(0)
               globaldata.Key_LRA(0) = V_LRA.new_lra(globaldata.Key_LRA(0), "BSFuntion")
               If ant <> globaldata.Key_LRA(0) Then modificado = True
            Case 1
               ant = globaldata.Key_LRA(1)
               globaldata.Key_LRA(1) = V_LRA.new_lra(globaldata.Key_LRA(1), "BSFuntion")
               If ant <> globaldata.Key_LRA(1) Then modificado = True
            Case 2
               ant = globaldata.Key_LRA(2)
               globaldata.Key_LRA(2) = V_LRA.new_lra(globaldata.Key_LRA(2), "BSFuntion")
               If ant <> globaldata.Key_LRA(2) Then modificado = True
            Case 3
               ant = globaldata.Key_LRA(3)
               globaldata.Key_LRA(3) = V_LRA.new_lra(globaldata.Key_LRA(3), "BSFuntion")
               If ant <> globaldata.Key_LRA(3) Then modificado = True
            Case 4
               ant = globaldata.Key_LRA(4)
               globaldata.Key_LRA(4) = V_LRA.new_lra(globaldata.Key_LRA(4), "BSFuntion")
               If ant <> globaldata.Key_LRA(4) Then modificado = True
            Case 5
               ant = globaldata.Key_LRA(5)
               globaldata.Key_LRA(5) = V_LRA.new_lra(globaldata.Key_LRA(5), "BSFuntion")
               If ant <> globaldata.Key_LRA(5) Then modificado = True
            Case 6
               ant = globaldata.Key_LRA(6)
               globaldata.Key_LRA(6) = V_LRA.new_lra(globaldata.Key_LRA(6), "BSFuntion")
               If ant <> globaldata.Key_LRA(6) Then modificado = True
            Case 7
               ant = globaldata.Key_LRA(7)
               globaldata.Key_LRA(7) = V_LRA.new_lra(globaldata.Key_LRA(7), "BSFuntion")
               If ant <> globaldata.Key_LRA(7) Then modificado = True
            Case 8
               ant = globaldata.Key_LRA(8)
               globaldata.Key_LRA(8) = V_LRA.new_lra(globaldata.Key_LRA(8), "BSFuntion")
               If ant <> globaldata.Key_LRA(8) Then modificado = True
            Case 9
               ant = globaldata.Key_LRA(9)
               globaldata.Key_LRA(9) = V_LRA.new_lra(globaldata.Key_LRA(9), "BSFuntion")
               If ant <> globaldata.Key_LRA(9) Then modificado = True
            Case 10
               ant = globaldata.Key_LRA(10)
               globaldata.Key_LRA(10) = V_LRA.new_lra(globaldata.Key_LRA(10), "BSFuntion")
               If ant <> globaldata.Key_LRA(10) Then modificado = True
            Case 11
               ant = globaldata.Key_LRA(11)
               globaldata.Key_LRA(11) = V_LRA.new_lra(globaldata.Key_LRA(11), "BSFuntion")
               If ant <> globaldata.Key_LRA(11) Then modificado = True
            Case 12
               ant = globaldata.Key_LRA(12)
               globaldata.Key_LRA(12) = V_LRA.new_lra(globaldata.Key_LRA(12), "BSFuntion")
               If ant <> globaldata.Key_LRA(12) Then modificado = True
            Case 13
               ant = globaldata.Key_LRA(13)
               globaldata.Key_LRA(13) = V_LRA.new_lra(globaldata.Key_LRA(13), "BSFuntion")
               If ant <> globaldata.Key_LRA(13) Then modificado = True
        End Select
    End If
End Sub

Private Sub NombreScreen_Change()
    modificado = True
End Sub

Private Sub NombreScreen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cerrar.SetFocus
End Sub

Private Sub sync_enable_Click()
    modificado = True
End Sub

Private Sub Text_syncLRA_Change()
    modificado = True
End Sub
Private Sub Text_syncLRA_DblClick()
    Text_syncLRA.Text = V_LRA.new_lra(Text_syncLRA.Text, "LRA")
End Sub

Private Sub Text_syncTime_Change()
    modificado = True
End Sub

Private Sub Text_TimeAutoCursor_Change()
    modificado = True
End Sub

Private Sub Text_TimeEdition_Change()
    modificado = True
End Sub

Private Sub Text_TimeOUT_Change()
    modificado = True
End Sub

Private Sub Text_TimeScan_Change()
    modificado = True
End Sub

Private Sub TextJmp_Change(Index As Integer)
    modificado = True
End Sub

Private Sub TextJmp_GotFocus(Index As Integer)
    tecla_en_foco = Index
    If Index < 10 Then
        Frm_teclas.Caption = "Tecla: " & Hex(Index)
    Else
        Frm_teclas.Caption = "Tecla: F" & Hex(Index - 9)
    End If
    If TextJmp(Index).Text = "" Then
        OP_NONE.Value = True
    End If
    If Mid(TextJmp(Index).Text, 1, 4) = "JUMP" Then
        OP_JUMP.Value = True
    End If
    If TextJmp(Index).Text = "BITSET" Then
        OP_BIT.Value = True
    End If
    Frm_teclas.Enabled = True
End Sub

Private Sub TextJmp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cerrar.SetFocus
End Sub

Private Sub TextJmpEsc_Change()
    modificado = True
End Sub

Private Sub TextJmpTimeOut_Change()
    modificado = True
End Sub

Private Sub Txt_pant_principal_Change()
    modificado = True
End Sub

Private Sub txt_salto_Change()
    TextJmp(tecla_en_foco).Text = "JUMP " & txt_salto.Text
End Sub
