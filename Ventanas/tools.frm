VERSION 5.00
Begin VB.Form V_Tools 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tools"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   1335
   Begin VB.Label LB_ALFA 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ALFANUM"
      BeginProperty Font 
         Name            =   "PatternLCD"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Tag             =   "Nuevo"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Lb_MTD 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MTDIGITAL"
      BeginProperty Font 
         Name            =   "PatternLCD"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Tag             =   "Nuevo"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lb_NUM 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NUMERICO"
      BeginProperty Font 
         Name            =   "PatternLCD"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Tag             =   "Nuevo"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lb_Mtext 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MTEXT"
      BeginProperty Font 
         Name            =   "PatternLCD"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Tag             =   "Nuevo"
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lb_CTEXT 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CTEXT"
      BeginProperty Font 
         Name            =   "PatternLCD"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   120
      TabIndex        =   0
      Tag             =   "Nuevo"
      Top             =   120
      Width           =   1100
   End
End
Attribute VB_Name = "V_Tools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LB_ALFA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    y_aux = Int(Y / PASOV)
    x_aux = Int(X / PASOH)
    LB_ALFA.Drag
End Sub

Private Sub lb_CTEXT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    y_aux = Int(Y / PASOV)
    x_aux = Int(X / PASOH)
   lb_CTEXT.Drag
End Sub

Private Sub Lb_MTD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    y_aux = Int(Y / PASOV)
    x_aux = Int(X / PASOH)
    Lb_MTD.Drag
End Sub

Private Sub lb_Mtext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    y_aux = Int(Y / PASOV)
    x_aux = Int(X / PASOH)
    lb_Mtext.Drag
End Sub

Private Sub lb_NUM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    y_aux = Int(Y / PASOV)
    x_aux = Int(X / PASOH)
    lb_NUM.Drag
End Sub
