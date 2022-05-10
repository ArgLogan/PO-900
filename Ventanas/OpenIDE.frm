VERSION 5.00
Begin VB.Form OpenIDE 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3345
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "OpenIDE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7080
      Begin VB.Timer TimerLED 
         Left            =   120
         Top             =   600
      End
      Begin VB.Timer TimerRunIDE 
         Left            =   120
         Top             =   2400
      End
      Begin VB.Shape EstatusLED 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   80
         Left            =   250
         Shape           =   4  'Rounded Rectangle
         Top             =   1365
         Width           =   120
      End
      Begin VB.Image imgLogo 
         Height          =   1920
         Left            =   195
         Picture         =   "OpenIDE.frx":000C
         Stretch         =   -1  'True
         Top             =   780
         Width           =   1965
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   4
         Top             =   2265
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "Compañía"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   2475
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         Caption         =   "Advertencia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   3060
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Versión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5970
         TabIndex        =   5
         Top             =   1905
         Width           =   885
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Plataforma"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5580
         TabIndex        =   6
         Top             =   1545
         Width           =   1275
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "PatternLCD"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         TabIndex        =   8
         Top             =   975
         Width           =   3360
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "Autorizado a"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   2715
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Producto de la compañía"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   585
         TabIndex        =   7
         Top             =   195
         Width           =   3000
      End
   End
End
Attribute VB_Name = "OpenIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim temp As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
    EditorIDE.Show
    Unload Me
End Sub

Private Sub Form_Load()
    temp = 0
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    lblCompanyProduct.Caption = "Controles Electronicos FERMI S.H."
    lblCompany.Caption = App.CompanyName
    lblCopyright.Caption = App.LegalCopyright
    lblPlatform.Caption = "Windows XP/Me/98se"
    lblLicenseTo.Caption = "Freeware"
    lblWarning.Caption = "La Empresa no se responsabiliza por el uso del producto"
    TimerRunIDE.Interval = 1500
    TimerRunIDE.Enabled = True
    TimerLED.Interval = 500
    TimerLED.Enabled = True
End Sub

Private Sub Frame1_Click()
    EditorIDE.Show
    Unload Me
End Sub

Private Sub imgLogo_Click()
    EditorIDE.Show
    Unload Me
End Sub

Private Sub lblCompany_Click()
    EditorIDE.Show
    Unload Me
End Sub

Private Sub lblCompanyProduct_Click()
    EditorIDE.Show
    Unload Me
End Sub

Private Sub lblCopyright_Click()
    EditorIDE.Show
    Unload Me
End Sub

Private Sub lblLicenseTo_Click()
    EditorIDE.Show
    Unload Me
End Sub

Private Sub lblPlatform_Click()
    EditorIDE.Show
    Unload Me
End Sub

Private Sub lblProductName_Click()
    EditorIDE.Show
    Unload Me
End Sub

Private Sub lblVersion_Click()
    EditorIDE.Show
    Unload Me
End Sub

Private Sub lblWarning_Click()
    EditorIDE.Show
    Unload Me
End Sub

Private Sub TimerLED_Timer()
    If EstatusLED.BackColor = NEGRO Then
        EstatusLED.BackColor = VERDE
    Else
        EstatusLED.BackColor = NEGRO
    End If
End Sub

Private Sub TimerRunIDE_Timer()
    If temp < 1 Then
        temp = temp + 1
        EditorIDE.Show
        Me.SetFocus
    Else
        Unload Me
    End If
End Sub
