VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form V_Prop_P_G 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5741
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Variables"
      TabPicture(0)   =   "V_Prop_P_G.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chk_default(0)"
      Tab(0).Control(1)=   "TextJmpEsc"
      Tab(0).Control(2)=   "chk_default(3)"
      Tab(0).Control(3)=   "Text_TimeEdition"
      Tab(0).Control(4)=   "Text_TimeAutoCursor"
      Tab(0).Control(5)=   "TextJmpTimeOut"
      Tab(0).Control(6)=   "Text_TimeOUT"
      Tab(0).Control(7)=   "Text_TimeScan"
      Tab(0).Control(8)=   "chk_default(2)"
      Tab(0).Control(9)=   "chk_default(1)"
      Tab(0).Control(10)=   "Label1(2)"
      Tab(0).Control(11)=   "Line3(1)"
      Tab(0).Control(12)=   "Line2(1)"
      Tab(0).Control(13)=   "Line1(1)"
      Tab(0).Control(14)=   "Label7(0)"
      Tab(0).Control(15)=   "Label6(0)"
      Tab(0).Control(16)=   "Label1(5)"
      Tab(0).Control(17)=   "Label1(4)"
      Tab(0).Control(18)=   "Label5(0)"
      Tab(0).Control(19)=   "Label4(0)"
      Tab(0).Control(20)=   "Label1(3)"
      Tab(0).Control(21)=   "Label1(1)"
      Tab(0).Control(22)=   "Label1(0)"
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Teclas"
      TabPicture(1)   =   "V_Prop_P_G.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "LabelTecla(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "LabelTecla(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "LabelTecla(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "LabelTecla(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "LabelTecla(4)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "LabelTecla(5)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "LabelTecla(6)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "LabelTecla(7)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "LabelTecla(8)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "LabelTecla(9)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "LabelTecla(10)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "LabelTecla(11)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "LabelTecla(12)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "LabelTecla(13)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "TextJmp(0)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "TextJmp(1)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "TextJmp(3)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "TextJmp(4)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "TextJmp(5)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "TextJmp(6)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "TextJmp(7)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "TextJmp(8)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "TextJmp(9)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "TextJmp(10)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "TextJmp(11)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "TextJmp(13)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "TextJmp(12)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Frm_teclas"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "TextJmp(2)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).ControlCount=   29
      Begin VB.CheckBox chk_default 
         Caption         =   "(Default)"
         Height          =   195
         Index           =   0
         Left            =   -74880
         TabIndex        =   68
         Top             =   553
         Width           =   975
      End
      Begin VB.TextBox TextJmpEsc 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -72480
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   508
         Width           =   495
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   1920
         Width           =   855
      End
      Begin VB.Frame Frm_teclas 
         Caption         =   "Tecla: "
         Enabled         =   0   'False
         Height          =   2295
         Left            =   4800
         TabIndex        =   44
         Top             =   480
         Width           =   1575
         Begin VB.OptionButton OP_NONE 
            Caption         =   "NONE"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton OP_BIT 
            Caption         =   "BIT"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton OP_JUMP 
            Caption         =   "JUMP"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txt_salto 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   840
            TabIndex        =   45
            Text            =   "FF"
            Top             =   1770
            Width           =   375
         End
         Begin VB.Label Label10 
            Caption         =   "Saltar a"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   49
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
         TabIndex        =   43
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   13
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   11
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   10
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   9
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   8
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   7
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   6
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   5
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox TextJmp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2520
         Width           =   855
      End
      Begin VB.CheckBox chk_default 
         Caption         =   "(Default)"
         Height          =   195
         Index           =   3
         Left            =   -74880
         TabIndex        =   21
         Top             =   2688
         Width           =   975
      End
      Begin VB.TextBox Text_TimeEdition 
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
         Left            =   -72480
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "30"
         Top             =   2643
         Width           =   495
      End
      Begin VB.TextBox Text_TimeAutoCursor 
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
         Left            =   -69600
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "2"
         Top             =   2643
         Width           =   495
      End
      Begin VB.TextBox TextJmpTimeOut 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69600
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1933
         Width           =   495
      End
      Begin VB.TextBox Text_TimeOUT 
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
         Left            =   -72480
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "60"
         Top             =   1933
         Width           =   495
      End
      Begin VB.TextBox Text_TimeScan 
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
         Left            =   -72480
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "100"
         Top             =   1218
         Width           =   495
      End
      Begin VB.CheckBox chk_default 
         Caption         =   "(Default)"
         Height          =   195
         Index           =   2
         Left            =   -74880
         TabIndex        =   15
         Top             =   1978
         Width           =   975
      End
      Begin VB.CheckBox chk_default 
         Caption         =   "(Default)"
         Height          =   195
         Index           =   1
         Left            =   -74880
         TabIndex        =   14
         Top             =   1263
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Salto de Escape"
         Height          =   195
         Index           =   2
         Left            =   -73800
         TabIndex        =   69
         Top             =   553
         Width           =   1335
      End
      Begin VB.Line Line3 
         Index           =   1
         X1              =   -75000
         X2              =   -68160
         Y1              =   1005
         Y2              =   1005
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   -75000
         X2              =   -68160
         Y1              =   1715
         Y2              =   1715
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   -75000
         X2              =   -68160
         Y1              =   2430
         Y2              =   2430
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla F4"
         Height          =   255
         Index           =   13
         Left            =   3480
         TabIndex        =   63
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla F3"
         Height          =   255
         Index           =   12
         Left            =   3480
         TabIndex        =   62
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla F2"
         Height          =   255
         Index           =   11
         Left            =   3480
         TabIndex        =   61
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla F1"
         Height          =   255
         Index           =   10
         Left            =   3480
         TabIndex        =   60
         Top             =   480
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 9"
         Height          =   255
         Index           =   9
         Left            =   2400
         TabIndex        =   59
         Top             =   480
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 8"
         Height          =   255
         Index           =   8
         Left            =   1320
         TabIndex        =   58
         Top             =   480
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 7"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   57
         Top             =   480
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 6"
         Height          =   255
         Index           =   6
         Left            =   2400
         TabIndex        =   56
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 5"
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   55
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 4"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   54
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 3"
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   53
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 2 "
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   52
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 1"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   51
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label LabelTecla 
         Alignment       =   2  'Center
         Caption         =   "Tecla 0 "
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   50
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Seg."
         Height          =   195
         Index           =   0
         Left            =   -69000
         TabIndex        =   30
         Top             =   2688
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Seg."
         Height          =   195
         Index           =   0
         Left            =   -71880
         TabIndex        =   29
         Top             =   2688
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Tiempo AutoCursor"
         Height          =   195
         Index           =   5
         Left            =   -71160
         TabIndex        =   28
         Top             =   2688
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Tiempo Edición"
         Height          =   195
         Index           =   4
         Left            =   -73800
         TabIndex        =   27
         Top             =   2688
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "mSeg."
         Height          =   195
         Index           =   0
         Left            =   -71880
         TabIndex        =   26
         Top             =   1263
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Seg."
         Height          =   195
         Index           =   0
         Left            =   -71880
         TabIndex        =   25
         Top             =   1978
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Salto de TimeOut"
         Height          =   195
         Index           =   3
         Left            =   -71160
         TabIndex        =   24
         Top             =   1978
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "TimeOut Pantalla"
         Height          =   195
         Index           =   1
         Left            =   -73800
         TabIndex        =   23
         Top             =   1978
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Tiempo de Scan"
         Height          =   195
         Index           =   0
         Left            =   -73800
         TabIndex        =   22
         Top             =   1263
         Width           =   1335
      End
   End
   Begin VB.Frame DataScreen 
      Caption         =   "Pantalla"
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5895
      Begin VB.TextBox NombreScreen 
         Height          =   285
         Left            =   1920
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label9 
         Caption         =   "Nombre de la Pantalla:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame SYNC_GLOBAL 
      Caption         =   "Globales "
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5895
      Begin VB.TextBox Txt_pant_inicial 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4800
         TabIndex        =   66
         Text            =   "10"
         Top             =   705
         Width           =   375
      End
      Begin VB.TextBox Txt_pant_principal 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1500
         TabIndex        =   12
         Text            =   "10"
         Top             =   705
         Width           =   375
      End
      Begin VB.TextBox Text_syncTime 
         Height          =   285
         Left            =   4800
         TabIndex        =   4
         Text            =   "50"
         Top             =   345
         Width           =   495
      End
      Begin VB.TextBox Text_syncLRA 
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "SIM|1:100"
         Top             =   345
         Width           =   1215
      End
      Begin VB.CheckBox sync_enable 
         Alignment       =   1  'Right Justify
         Caption         =   "Sync PLC Enable"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Pantalla Inicial"
         Height          =   255
         Left            =   3600
         TabIndex        =   65
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lb_pp 
         Caption         =   "Pantalla Principal"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Seg"
         Height          =   255
         Left            =   5400
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Sync Time"
         Height          =   255
         Left            =   3960
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Sync LRA"
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton Cerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   480
      Width           =   735
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
Dim defaultdatas As Class_GLOBAL
Dim screendata As Variant
Dim tecla_en_foco As Byte

Public Function viewglobals(idprj As String) As Boolean
    Dim indice As Byte
    Dim j As Byte
    Set ptr_proyect = Proyectos.item(idprj)
    Set globaldata = ptr_proyect.dataglobal
    Set defaultdatas = ptr_proyect.dataglobal
         
        
    Me.Caption = "Defaults / " + ptr_proyect.Nombre
    SYNC_GLOBAL.Visible = True
    DataScreen.Visible = False
    
    Text_TimeScan.Text = Str(ptr_proyect.dataglobal.tdisplay)
    Text_TimeOUT.Text = Str(globaldata.timeout)
    Text_TimeEdition.Text = Str(globaldata.timeoutedit)
    Text_TimeAutoCursor.Text = Str(globaldata.timeautocursor)
    
    If globaldata.Pant_principal_enable Then
        Txt_pant_principal.Text = Hex(globaldata.Pant_principal) 'verifiacar
    Else
        Txt_pant_principal.Text = ""
    End If
    If globaldata.Pant_inicial_enable Then
        Txt_pant_inicial.Text = Hex(globaldata.Pant_inicial) 'verifiacar
    Else
        Txt_pant_inicial.Text = ""
    End If
    
    
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
    TextJmpEsc.Locked = False
    Text_TimeScan.Locked = False
    Text_TimeOUT.Locked = False
    TextJmpTimeOut.Locked = False
    Text_TimeEdition.Locked = False
    Text_TimeAutoCursor.Locked = False
    For j = 0 To 3
        chk_default(j).Visible = False
    Next j
    Me.Show (1)
    
    viewglobals = True
End Function

Public Function viewlocals(idprj As String, idscreen As Byte) As Boolean
    Dim indice As Byte
    Dim j As Byte
    Set ptr_proyect = Proyectos.item(idprj)
    Set globaldata = ptr_proyect.m_Screens.item(genidpan(idscreen))
    Set defaultdatas = ptr_proyect.dataglobal

    Me.Caption = "Locales / " + ptr_proyect.Nombre + " / " + globaldata.name
    
    SYNC_GLOBAL.Visible = False
    DataScreen.Visible = True
    
    NombreScreen.Text = globaldata.name
    
    If globaldata.tdisplay_local = True Then
        Text_TimeScan.Text = Str(globaldata.tdisplay)
        chk_default(1).Value = 1
    Else
        Text_TimeScan.Text = Str(defaultdatas.tdisplay)
    End If
    
    If globaldata.timeout_local = True Then
        Text_TimeOUT.Text = Str(globaldata.timeout)
        TextJmpTimeOut.Text = Hex(globaldata.timeoutjmp)
        chk_default(2).Value = 1
    Else
        Text_TimeOUT.Text = Str(defaultdatas.timeout)
        TextJmpTimeOut.Text = Hex(defaultdatas.timeoutjmp)
    End If
    
    If globaldata.timeedit_local = True Then
        Text_TimeEdition.Text = Str(globaldata.timeoutedit)
        Text_TimeAutoCursor.Text = Str(globaldata.timeautocursor)
        chk_default(3).Value = 1
    Else
        Text_TimeEdition.Text = Str(defaultdatas.timeoutedit)
        Text_TimeAutoCursor.Text = Str(defaultdatas.timeautocursor)
    End If
    
    If globaldata.escjmp_local = True Then
        TextJmpEsc.Text = Hex(globaldata.escjmp)
        chk_default(0).Value = 1
    Else
        TextJmpEsc.Text = Hex(defaultdatas.escjmp)
    End If
    
    For indice = 0 To 13
        If globaldata.keyjmpenable(indice) Then
            TextJmp(indice).Text = "JUMP " & Hex(globaldata.keyjmp(indice))
        ElseIf globaldata.Key_LRA(indice) <> "OFF" Then
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
            If Val("&H" + TextJmpTimeOut.Text) < 255 Then
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
        ' Pantalla de Globales
            globaldata.syncactivo = sync_enable.Value
            If Text_syncLRA.Text <> "" Then globaldata.synclra = Text_syncLRA.Text
            If IsNumeric(Text_syncTime.Text) Then globaldata.synctime = Val(Text_syncTime.Text)
            
            If Txt_pant_principal.Text <> "" Then
                globaldata.Pant_principal_enable = True
                If Val("&H" & Txt_pant_principal.Text) < 254 Then
                    globaldata.Pant_principal = Val("&H" & Txt_pant_principal.Text)
                Else
                    globaldata.Pant_principal = 0
                End If
            Else
                globaldata.Pant_principal_enable = False
            End If

            If Txt_pant_inicial.Text <> "" Then
                globaldata.Pant_inicial_enable = True
                If Val("&H" & Txt_pant_inicial.Text) < 254 Then
                    globaldata.Pant_inicial = Val("&H" & Txt_pant_inicial.Text)
                Else
                    globaldata.Pant_inicial = 0
                End If
            Else
                globaldata.Pant_inicial_enable = False
            End If


        Else
            globaldata.name = NombreScreen.Text
            globaldata.escjmp_local = CBool(chk_default(0).Value)
            globaldata.tdisplay_local = CBool(chk_default(1).Value)
            globaldata.timeout_local = CBool(chk_default(2).Value)
            globaldata.timeedit_local = CBool(chk_default(3).Value)

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

Private Sub chk_default_Click(Index As Integer)
    modificado = True
    
    If chk_default(Index).Value = 1 Then
        chk_default(Index).Caption = "Local"
        Select Case Index
            Case 0
                TextJmpEsc.Locked = False
            Case 1
                Text_TimeScan.Locked = False
            Case 2
                Text_TimeOUT.Locked = False
                TextJmpTimeOut.Locked = False
            Case 3
                Text_TimeEdition.Locked = False
                Text_TimeAutoCursor.Locked = False
        End Select
    Else
        chk_default(Index).Caption = "(Default)"
        Select Case Index
            Case 0
                TextJmpEsc.Locked = True
                TextJmpEsc.Text = Hex(defaultdatas.escjmp)
            Case 1
                Text_TimeScan.Locked = True
                Text_TimeScan.Text = Str(defaultdatas.tdisplay)
            Case 2
                Text_TimeOUT.Locked = True
                Text_TimeOUT.Text = Str(defaultdatas.timeout)
                TextJmpTimeOut.Locked = True
                TextJmpTimeOut.Text = Hex(defaultdatas.timeoutjmp)
            Case 3
                Text_TimeEdition.Locked = True
                Text_TimeEdition.Text = Str(defaultdatas.timeoutedit)
                Text_TimeAutoCursor.Locked = True
                Text_TimeAutoCursor.Text = Str(defaultdatas.timeautocursor)
        End Select
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set globaldata = Nothing
    Set defaultdatas = Nothing
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
    If Mid(TextJmp(tecla_en_foco).Text, 1, 4) <> "JUMP" Then
        TextJmp(tecla_en_foco).Text = "JUMP " & txt_salto.Text
        txt_salto.Enabled = True
        txt_salto.Text = Hex(defaultdatas.Pant_principal)
    End If
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
               globaldata.Key_LRA(0) = V_LRA.new_lra(globaldata.Key_LRA(0), "BSFuntion", ptr_proyect.bit_limite_sup, ptr_proyect.bit_limite_inf)
               If ant <> globaldata.Key_LRA(0) Then modificado = True
            Case 1
               ant = globaldata.Key_LRA(1)
               globaldata.Key_LRA(1) = V_LRA.new_lra(globaldata.Key_LRA(1), "BSFuntion", ptr_proyect.bit_limite_sup, ptr_proyect.bit_limite_inf)
               If ant <> globaldata.Key_LRA(1) Then modificado = True
            Case 2
               ant = globaldata.Key_LRA(2)
               globaldata.Key_LRA(2) = V_LRA.new_lra(globaldata.Key_LRA(2), "BSFuntion", ptr_proyect.bit_limite_sup, ptr_proyect.bit_limite_inf)
               If ant <> globaldata.Key_LRA(2) Then modificado = True
            Case 3
               ant = globaldata.Key_LRA(3)
               globaldata.Key_LRA(3) = V_LRA.new_lra(globaldata.Key_LRA(3), "BSFuntion", ptr_proyect.bit_limite_sup, ptr_proyect.bit_limite_inf)
               If ant <> globaldata.Key_LRA(3) Then modificado = True
            Case 4
               ant = globaldata.Key_LRA(4)
               globaldata.Key_LRA(4) = V_LRA.new_lra(globaldata.Key_LRA(4), "BSFuntion", ptr_proyect.bit_limite_sup, ptr_proyect.bit_limite_inf)
               If ant <> globaldata.Key_LRA(4) Then modificado = True
            Case 5
               ant = globaldata.Key_LRA(5)
               globaldata.Key_LRA(5) = V_LRA.new_lra(globaldata.Key_LRA(5), "BSFuntion", ptr_proyect.bit_limite_sup, ptr_proyect.bit_limite_inf)
               If ant <> globaldata.Key_LRA(5) Then modificado = True
            Case 6
               ant = globaldata.Key_LRA(6)
               globaldata.Key_LRA(6) = V_LRA.new_lra(globaldata.Key_LRA(6), "BSFuntion", ptr_proyect.bit_limite_sup, ptr_proyect.bit_limite_inf)
               If ant <> globaldata.Key_LRA(6) Then modificado = True
            Case 7
               ant = globaldata.Key_LRA(7)
               globaldata.Key_LRA(7) = V_LRA.new_lra(globaldata.Key_LRA(7), "BSFuntion", ptr_proyect.bit_limite_sup, ptr_proyect.bit_limite_inf)
               If ant <> globaldata.Key_LRA(7) Then modificado = True
            Case 8
               ant = globaldata.Key_LRA(8)
               globaldata.Key_LRA(8) = V_LRA.new_lra(globaldata.Key_LRA(8), "BSFuntion", ptr_proyect.bit_limite_sup, ptr_proyect.bit_limite_inf)
               If ant <> globaldata.Key_LRA(8) Then modificado = True
            Case 9
               ant = globaldata.Key_LRA(9)
               globaldata.Key_LRA(9) = V_LRA.new_lra(globaldata.Key_LRA(9), "BSFuntion", ptr_proyect.bit_limite_sup, ptr_proyect.bit_limite_inf)
               If ant <> globaldata.Key_LRA(9) Then modificado = True
            Case 10
               ant = globaldata.Key_LRA(10)
               globaldata.Key_LRA(10) = V_LRA.new_lra(globaldata.Key_LRA(10), "BSFuntion", ptr_proyect.bit_limite_sup, ptr_proyect.bit_limite_inf)
               If ant <> globaldata.Key_LRA(10) Then modificado = True
            Case 11
               ant = globaldata.Key_LRA(11)
               globaldata.Key_LRA(11) = V_LRA.new_lra(globaldata.Key_LRA(11), "BSFuntion", ptr_proyect.bit_limite_sup, ptr_proyect.bit_limite_inf)
               If ant <> globaldata.Key_LRA(11) Then modificado = True
            Case 12
               ant = globaldata.Key_LRA(12)
               globaldata.Key_LRA(12) = V_LRA.new_lra(globaldata.Key_LRA(12), "BSFuntion", ptr_proyect.bit_limite_sup, ptr_proyect.bit_limite_inf)
               If ant <> globaldata.Key_LRA(12) Then modificado = True
            Case 13
               ant = globaldata.Key_LRA(13)
               globaldata.Key_LRA(13) = V_LRA.new_lra(globaldata.Key_LRA(13), "BSFuntion", ptr_proyect.bit_limite_sup, ptr_proyect.bit_limite_inf)
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
    Text_syncLRA.Text = V_LRA.new_lra(Text_syncLRA.Text, "SYNC", ptr_proyect.lra_limite_sup, ptr_proyect.lra_limite_inf)
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
        txt_salto.Text = Mid(TextJmp(Index).Text, 6, 2)
        txt_salto.Enabled = True
    End If
    If TextJmp(Index).Text = "BITSET" Then
        OP_BIT.Value = True
    End If
    Frm_teclas.Enabled = True
End Sub

Private Sub TextJmpEsc_Change()
    modificado = True
End Sub

Private Sub TextJmpTimeOut_Change()
    modificado = True
End Sub

Private Sub Txt_pant_inicial_Change()
    modificado = True
    If Txt_pant_inicial.Text <> "" Then
        globaldata.Pant_inicial_enable = True
        If Val("&H" & Txt_pant_inicial.Text) < 254 Then
            globaldata.Pant_inicial = Val("&H" & Txt_pant_inicial.Text)
        Else
            globaldata.Pant_inicial = 0
        End If
    Else
        globaldata.Pant_inicial_enable = False
    End If
End Sub

Private Sub Txt_pant_principal_Change()
    modificado = True
    If Txt_pant_principal.Text <> "" Then
        globaldata.Pant_principal_enable = True
        If Val("&H" & Txt_pant_principal.Text) < 254 Then
            globaldata.Pant_principal = Val("&H" & Txt_pant_principal.Text)
        Else
            globaldata.Pant_principal = 0
        End If
    Else
        globaldata.Pant_principal_enable = False
    End If

End Sub

Private Sub txt_salto_Change()
    If txt_salto.Text <> "" Then
        TextJmp(tecla_en_foco).Text = "JUMP " & txt_salto.Text
    Else
        TextJmp(tecla_en_foco).Text = "JUMP FF"
    End If
End Sub
