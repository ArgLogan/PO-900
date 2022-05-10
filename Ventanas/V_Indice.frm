VERSION 5.00
Begin VB.Form V_Indice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nombre Proyeto"
   ClientHeight    =   4260
   ClientLeft      =   6420
   ClientTop       =   2640
   ClientWidth     =   2730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4260
   ScaleWidth      =   2730
   Begin VB.CommandButton B_Borrar 
      Caption         =   "Borrar"
      Enabled         =   0   'False
      Height          =   420
      Left            =   1485
      TabIndex        =   270
      Top             =   3780
      Width           =   1095
   End
   Begin VB.CommandButton B_Pegar 
      Caption         =   "Pegar"
      Enabled         =   0   'False
      Height          =   420
      Left            =   135
      TabIndex        =   269
      Top             =   3780
      Width           =   1095
   End
   Begin VB.CommandButton B_Copiar 
      Caption         =   "Copiar"
      Enabled         =   0   'False
      Height          =   420
      Left            =   1485
      TabIndex        =   268
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton B_Cortar 
      Caption         =   "Cortar"
      Enabled         =   0   'False
      Height          =   420
      Left            =   135
      TabIndex        =   267
      Top             =   3240
      Width           =   1095
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   0
      Left            =   405
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   1
      Left            =   540
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   135
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   2
      Left            =   675
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   135
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   3
      Left            =   810
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   135
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   4
      Left            =   945
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   135
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   5
      Left            =   1080
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   135
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   6
      Left            =   1215
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   135
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   7
      Left            =   1350
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   135
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   8
      Left            =   1485
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   135
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   9
      Left            =   1620
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   135
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   10
      Left            =   1755
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   135
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   11
      Left            =   1890
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   135
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   12
      Left            =   2025
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   135
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   13
      Left            =   2160
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   135
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   14
      Left            =   2295
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   135
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   15
      Left            =   2430
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   135
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   16
      Left            =   405
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   270
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   17
      Left            =   540
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   270
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   18
      Left            =   675
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   270
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   19
      Left            =   810
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   270
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   20
      Left            =   945
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   270
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   21
      Left            =   1080
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   270
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   22
      Left            =   1215
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   270
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   23
      Left            =   1350
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   270
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   24
      Left            =   1485
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   270
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   25
      Left            =   1620
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   270
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   26
      Left            =   1755
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   270
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   27
      Left            =   1890
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   270
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   28
      Left            =   2025
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   270
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   29
      Left            =   2160
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   270
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   30
      Left            =   2295
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   270
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   31
      Left            =   2430
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   270
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   32
      Left            =   405
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   405
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   33
      Left            =   540
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   405
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   34
      Left            =   675
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   405
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   35
      Left            =   810
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   405
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   36
      Left            =   945
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   405
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   37
      Left            =   1080
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   405
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   38
      Left            =   1215
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   405
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   39
      Left            =   1350
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   405
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   40
      Left            =   1485
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   405
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   41
      Left            =   1620
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   405
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   42
      Left            =   1755
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   405
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   43
      Left            =   1890
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   405
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   44
      Left            =   2025
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   405
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   45
      Left            =   2160
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   405
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   46
      Left            =   2295
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   405
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   47
      Left            =   2430
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   405
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   48
      Left            =   405
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   540
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   49
      Left            =   540
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   540
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   50
      Left            =   675
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   540
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   51
      Left            =   810
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   540
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   52
      Left            =   945
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   540
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   53
      Left            =   1080
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   540
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   54
      Left            =   1215
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   540
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   55
      Left            =   1350
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   540
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   56
      Left            =   1485
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   540
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   57
      Left            =   1620
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   540
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   58
      Left            =   1755
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   540
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   59
      Left            =   1890
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   540
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   60
      Left            =   2025
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   540
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   61
      Left            =   2160
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   540
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   62
      Left            =   2295
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   540
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   63
      Left            =   2430
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   540
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   64
      Left            =   405
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   675
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   65
      Left            =   540
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   675
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   66
      Left            =   675
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   675
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   67
      Left            =   810
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   675
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   68
      Left            =   945
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   675
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   69
      Left            =   1080
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   675
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   70
      Left            =   1215
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   675
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   71
      Left            =   1350
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   675
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   72
      Left            =   1485
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   675
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   73
      Left            =   1620
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   675
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   74
      Left            =   1755
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   675
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   75
      Left            =   1890
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   675
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   76
      Left            =   2025
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   675
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   77
      Left            =   2160
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   675
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   78
      Left            =   2295
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   675
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   79
      Left            =   2430
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   675
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   80
      Left            =   405
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   810
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   81
      Left            =   540
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   810
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   82
      Left            =   675
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   810
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   83
      Left            =   810
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   810
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   84
      Left            =   945
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   810
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   85
      Left            =   1080
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   810
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   86
      Left            =   1215
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   810
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   87
      Left            =   1350
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   810
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   88
      Left            =   1485
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   810
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   89
      Left            =   1620
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   810
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   90
      Left            =   1755
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   810
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   91
      Left            =   1890
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   810
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   92
      Left            =   2025
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   810
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   93
      Left            =   2160
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   810
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   94
      Left            =   2295
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   810
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   95
      Left            =   2430
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   810
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   96
      Left            =   405
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   945
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   97
      Left            =   540
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   945
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   98
      Left            =   675
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   945
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   99
      Left            =   810
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   945
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   100
      Left            =   945
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   945
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   101
      Left            =   1080
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   945
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   102
      Left            =   1215
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   945
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   103
      Left            =   1350
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   945
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   104
      Left            =   1485
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   945
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   105
      Left            =   1620
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   945
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   106
      Left            =   1755
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   945
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   107
      Left            =   1890
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   945
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   108
      Left            =   2025
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   945
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   109
      Left            =   2160
      TabIndex        =   109
      TabStop         =   0   'False
      Top             =   945
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   110
      Left            =   2295
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   945
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   111
      Left            =   2430
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   945
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   112
      Left            =   405
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   1080
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   113
      Left            =   540
      TabIndex        =   113
      TabStop         =   0   'False
      Top             =   1080
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   114
      Left            =   675
      TabIndex        =   114
      TabStop         =   0   'False
      Top             =   1080
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   115
      Left            =   810
      TabIndex        =   115
      TabStop         =   0   'False
      Top             =   1080
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   116
      Left            =   945
      TabIndex        =   116
      TabStop         =   0   'False
      Top             =   1080
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   117
      Left            =   1080
      TabIndex        =   117
      TabStop         =   0   'False
      Top             =   1080
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   118
      Left            =   1215
      TabIndex        =   118
      TabStop         =   0   'False
      Top             =   1080
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   119
      Left            =   1350
      TabIndex        =   119
      TabStop         =   0   'False
      Top             =   1080
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   120
      Left            =   1485
      TabIndex        =   120
      TabStop         =   0   'False
      Top             =   1080
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   121
      Left            =   1620
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   1080
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   122
      Left            =   1755
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   1080
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   123
      Left            =   1890
      TabIndex        =   123
      TabStop         =   0   'False
      Top             =   1080
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   124
      Left            =   2025
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   1080
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   125
      Left            =   2160
      TabIndex        =   125
      TabStop         =   0   'False
      Top             =   1080
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   126
      Left            =   2295
      TabIndex        =   126
      TabStop         =   0   'False
      Top             =   1080
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   127
      Left            =   2430
      TabIndex        =   127
      TabStop         =   0   'False
      Top             =   1080
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   128
      Left            =   405
      TabIndex        =   128
      TabStop         =   0   'False
      Top             =   1215
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   129
      Left            =   540
      TabIndex        =   129
      TabStop         =   0   'False
      Top             =   1215
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   130
      Left            =   675
      TabIndex        =   130
      TabStop         =   0   'False
      Top             =   1215
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   131
      Left            =   810
      TabIndex        =   131
      TabStop         =   0   'False
      Top             =   1215
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   132
      Left            =   945
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   1215
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   133
      Left            =   1080
      TabIndex        =   133
      TabStop         =   0   'False
      Top             =   1215
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   134
      Left            =   1215
      TabIndex        =   134
      TabStop         =   0   'False
      Top             =   1215
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   135
      Left            =   1350
      TabIndex        =   135
      TabStop         =   0   'False
      Top             =   1215
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   136
      Left            =   1485
      TabIndex        =   136
      TabStop         =   0   'False
      Top             =   1215
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   137
      Left            =   1620
      TabIndex        =   137
      TabStop         =   0   'False
      Top             =   1215
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   138
      Left            =   1755
      TabIndex        =   138
      TabStop         =   0   'False
      Top             =   1215
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   139
      Left            =   1890
      TabIndex        =   139
      TabStop         =   0   'False
      Top             =   1215
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   140
      Left            =   2025
      TabIndex        =   140
      TabStop         =   0   'False
      Top             =   1215
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   141
      Left            =   2160
      TabIndex        =   141
      TabStop         =   0   'False
      Top             =   1215
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   142
      Left            =   2295
      TabIndex        =   142
      TabStop         =   0   'False
      Top             =   1215
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   143
      Left            =   2430
      TabIndex        =   143
      TabStop         =   0   'False
      Top             =   1215
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   144
      Left            =   405
      TabIndex        =   144
      TabStop         =   0   'False
      Top             =   1350
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   145
      Left            =   540
      TabIndex        =   145
      TabStop         =   0   'False
      Top             =   1350
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   146
      Left            =   675
      TabIndex        =   146
      TabStop         =   0   'False
      Top             =   1350
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   147
      Left            =   810
      TabIndex        =   147
      TabStop         =   0   'False
      Top             =   1350
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   148
      Left            =   945
      TabIndex        =   148
      TabStop         =   0   'False
      Top             =   1350
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   149
      Left            =   1080
      TabIndex        =   149
      TabStop         =   0   'False
      Top             =   1350
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   150
      Left            =   1215
      TabIndex        =   150
      TabStop         =   0   'False
      Top             =   1350
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   151
      Left            =   1350
      TabIndex        =   151
      TabStop         =   0   'False
      Top             =   1350
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   152
      Left            =   1485
      TabIndex        =   152
      TabStop         =   0   'False
      Top             =   1350
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   153
      Left            =   1620
      TabIndex        =   153
      TabStop         =   0   'False
      Top             =   1350
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   154
      Left            =   1755
      TabIndex        =   154
      TabStop         =   0   'False
      Top             =   1350
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   155
      Left            =   1890
      TabIndex        =   155
      TabStop         =   0   'False
      Top             =   1350
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   156
      Left            =   2025
      TabIndex        =   156
      TabStop         =   0   'False
      Top             =   1350
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   157
      Left            =   2160
      TabIndex        =   157
      TabStop         =   0   'False
      Top             =   1350
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   158
      Left            =   2295
      TabIndex        =   158
      TabStop         =   0   'False
      Top             =   1350
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   159
      Left            =   2430
      TabIndex        =   159
      TabStop         =   0   'False
      Top             =   1350
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   160
      Left            =   405
      TabIndex        =   160
      TabStop         =   0   'False
      Top             =   1485
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   161
      Left            =   540
      TabIndex        =   161
      TabStop         =   0   'False
      Top             =   1485
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   162
      Left            =   675
      TabIndex        =   162
      TabStop         =   0   'False
      Top             =   1485
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   163
      Left            =   810
      TabIndex        =   163
      TabStop         =   0   'False
      Top             =   1485
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   164
      Left            =   945
      TabIndex        =   164
      TabStop         =   0   'False
      Top             =   1485
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   165
      Left            =   1080
      TabIndex        =   165
      TabStop         =   0   'False
      Top             =   1485
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   166
      Left            =   1215
      TabIndex        =   166
      TabStop         =   0   'False
      Top             =   1485
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   167
      Left            =   1350
      TabIndex        =   167
      TabStop         =   0   'False
      Top             =   1485
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   168
      Left            =   1485
      TabIndex        =   168
      TabStop         =   0   'False
      Top             =   1485
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   169
      Left            =   1620
      TabIndex        =   169
      TabStop         =   0   'False
      Top             =   1485
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   170
      Left            =   1755
      TabIndex        =   170
      TabStop         =   0   'False
      Top             =   1485
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   171
      Left            =   1890
      TabIndex        =   171
      TabStop         =   0   'False
      Top             =   1485
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   172
      Left            =   2025
      TabIndex        =   172
      TabStop         =   0   'False
      Top             =   1485
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   173
      Left            =   2160
      TabIndex        =   173
      TabStop         =   0   'False
      Top             =   1485
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   174
      Left            =   2295
      TabIndex        =   174
      TabStop         =   0   'False
      Top             =   1485
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   175
      Left            =   2430
      TabIndex        =   175
      TabStop         =   0   'False
      Top             =   1485
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   176
      Left            =   405
      TabIndex        =   176
      TabStop         =   0   'False
      Top             =   1620
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   177
      Left            =   540
      TabIndex        =   177
      TabStop         =   0   'False
      Top             =   1620
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   178
      Left            =   675
      TabIndex        =   178
      TabStop         =   0   'False
      Top             =   1620
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   179
      Left            =   810
      TabIndex        =   179
      TabStop         =   0   'False
      Top             =   1620
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   180
      Left            =   945
      TabIndex        =   180
      TabStop         =   0   'False
      Top             =   1620
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   181
      Left            =   1080
      TabIndex        =   181
      TabStop         =   0   'False
      Top             =   1620
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   182
      Left            =   1215
      TabIndex        =   182
      TabStop         =   0   'False
      Top             =   1620
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   183
      Left            =   1350
      TabIndex        =   183
      TabStop         =   0   'False
      Top             =   1620
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   184
      Left            =   1485
      TabIndex        =   184
      TabStop         =   0   'False
      Top             =   1620
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   185
      Left            =   1620
      TabIndex        =   185
      TabStop         =   0   'False
      Top             =   1620
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   186
      Left            =   1755
      TabIndex        =   186
      TabStop         =   0   'False
      Top             =   1620
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   187
      Left            =   1890
      TabIndex        =   187
      TabStop         =   0   'False
      Top             =   1620
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   188
      Left            =   2025
      TabIndex        =   188
      TabStop         =   0   'False
      Top             =   1620
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   189
      Left            =   2160
      TabIndex        =   189
      TabStop         =   0   'False
      Top             =   1620
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   190
      Left            =   2295
      TabIndex        =   190
      TabStop         =   0   'False
      Top             =   1620
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   191
      Left            =   2430
      TabIndex        =   191
      TabStop         =   0   'False
      Top             =   1620
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   192
      Left            =   405
      TabIndex        =   192
      TabStop         =   0   'False
      Top             =   1755
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   193
      Left            =   540
      TabIndex        =   193
      TabStop         =   0   'False
      Top             =   1755
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   194
      Left            =   675
      TabIndex        =   194
      TabStop         =   0   'False
      Top             =   1755
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   195
      Left            =   810
      TabIndex        =   195
      TabStop         =   0   'False
      Top             =   1755
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   196
      Left            =   945
      TabIndex        =   196
      TabStop         =   0   'False
      Top             =   1755
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   197
      Left            =   1080
      TabIndex        =   197
      TabStop         =   0   'False
      Top             =   1755
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   198
      Left            =   1215
      TabIndex        =   198
      TabStop         =   0   'False
      Top             =   1755
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   199
      Left            =   1350
      TabIndex        =   199
      TabStop         =   0   'False
      Top             =   1755
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   200
      Left            =   1485
      TabIndex        =   200
      TabStop         =   0   'False
      Top             =   1755
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   201
      Left            =   1620
      TabIndex        =   201
      TabStop         =   0   'False
      Top             =   1755
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   202
      Left            =   1755
      TabIndex        =   202
      TabStop         =   0   'False
      Top             =   1755
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   203
      Left            =   1890
      TabIndex        =   203
      TabStop         =   0   'False
      Top             =   1755
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   204
      Left            =   2025
      TabIndex        =   204
      TabStop         =   0   'False
      Top             =   1755
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   205
      Left            =   2160
      TabIndex        =   205
      TabStop         =   0   'False
      Top             =   1755
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   206
      Left            =   2295
      TabIndex        =   206
      TabStop         =   0   'False
      Top             =   1755
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   207
      Left            =   2430
      TabIndex        =   207
      TabStop         =   0   'False
      Top             =   1755
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   208
      Left            =   405
      TabIndex        =   208
      TabStop         =   0   'False
      Top             =   1890
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   209
      Left            =   540
      TabIndex        =   209
      TabStop         =   0   'False
      Top             =   1890
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   210
      Left            =   675
      TabIndex        =   210
      TabStop         =   0   'False
      Top             =   1890
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   211
      Left            =   810
      TabIndex        =   211
      TabStop         =   0   'False
      Top             =   1890
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   212
      Left            =   945
      TabIndex        =   212
      TabStop         =   0   'False
      Top             =   1890
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   213
      Left            =   1080
      TabIndex        =   213
      TabStop         =   0   'False
      Top             =   1890
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   214
      Left            =   1215
      TabIndex        =   214
      TabStop         =   0   'False
      Top             =   1890
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   215
      Left            =   1350
      TabIndex        =   215
      TabStop         =   0   'False
      Top             =   1890
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   216
      Left            =   1485
      TabIndex        =   216
      TabStop         =   0   'False
      Top             =   1890
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   217
      Left            =   1620
      TabIndex        =   217
      TabStop         =   0   'False
      Top             =   1890
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   218
      Left            =   1755
      TabIndex        =   218
      TabStop         =   0   'False
      Top             =   1890
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   219
      Left            =   1890
      TabIndex        =   219
      TabStop         =   0   'False
      Top             =   1890
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   220
      Left            =   2025
      TabIndex        =   220
      TabStop         =   0   'False
      Top             =   1890
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   221
      Left            =   2160
      TabIndex        =   221
      TabStop         =   0   'False
      Top             =   1890
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   222
      Left            =   2295
      TabIndex        =   222
      TabStop         =   0   'False
      Top             =   1890
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   223
      Left            =   2430
      TabIndex        =   223
      TabStop         =   0   'False
      Top             =   1890
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   224
      Left            =   405
      TabIndex        =   224
      TabStop         =   0   'False
      Top             =   2025
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   225
      Left            =   540
      TabIndex        =   225
      TabStop         =   0   'False
      Top             =   2025
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   226
      Left            =   675
      TabIndex        =   226
      TabStop         =   0   'False
      Top             =   2025
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   227
      Left            =   810
      TabIndex        =   227
      TabStop         =   0   'False
      Top             =   2025
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   228
      Left            =   945
      TabIndex        =   228
      TabStop         =   0   'False
      Top             =   2025
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   229
      Left            =   1080
      TabIndex        =   229
      TabStop         =   0   'False
      Top             =   2025
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   230
      Left            =   1215
      TabIndex        =   230
      TabStop         =   0   'False
      Top             =   2025
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   231
      Left            =   1350
      TabIndex        =   231
      TabStop         =   0   'False
      Top             =   2025
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   232
      Left            =   1485
      TabIndex        =   232
      TabStop         =   0   'False
      Top             =   2025
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   233
      Left            =   1620
      TabIndex        =   233
      TabStop         =   0   'False
      Top             =   2025
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   234
      Left            =   1755
      TabIndex        =   234
      TabStop         =   0   'False
      Top             =   2025
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   235
      Left            =   1890
      TabIndex        =   235
      TabStop         =   0   'False
      Top             =   2025
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   236
      Left            =   2025
      TabIndex        =   236
      TabStop         =   0   'False
      Top             =   2025
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   237
      Left            =   2160
      TabIndex        =   237
      TabStop         =   0   'False
      Top             =   2025
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   238
      Left            =   2295
      TabIndex        =   238
      TabStop         =   0   'False
      Top             =   2025
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   239
      Left            =   2430
      TabIndex        =   239
      TabStop         =   0   'False
      Top             =   2025
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   240
      Left            =   405
      TabIndex        =   240
      TabStop         =   0   'False
      Top             =   2160
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   241
      Left            =   540
      TabIndex        =   241
      TabStop         =   0   'False
      Top             =   2160
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   242
      Left            =   675
      TabIndex        =   242
      TabStop         =   0   'False
      Top             =   2160
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   243
      Left            =   810
      TabIndex        =   243
      TabStop         =   0   'False
      Top             =   2160
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   244
      Left            =   945
      TabIndex        =   244
      TabStop         =   0   'False
      Top             =   2160
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   245
      Left            =   1080
      TabIndex        =   245
      TabStop         =   0   'False
      Top             =   2160
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   246
      Left            =   1215
      TabIndex        =   246
      TabStop         =   0   'False
      Top             =   2160
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   247
      Left            =   1350
      TabIndex        =   247
      TabStop         =   0   'False
      Top             =   2160
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   248
      Left            =   1485
      TabIndex        =   248
      TabStop         =   0   'False
      Top             =   2160
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   249
      Left            =   1620
      TabIndex        =   249
      TabStop         =   0   'False
      Top             =   2160
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   250
      Left            =   1755
      TabIndex        =   250
      TabStop         =   0   'False
      Top             =   2160
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   251
      Left            =   1890
      TabIndex        =   251
      TabStop         =   0   'False
      Top             =   2160
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   252
      Left            =   2025
      TabIndex        =   252
      TabStop         =   0   'False
      Top             =   2160
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   253
      Left            =   2160
      TabIndex        =   253
      TabStop         =   0   'False
      Top             =   2160
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   254
      Left            =   2295
      TabIndex        =   254
      TabStop         =   0   'False
      Top             =   2160
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin EDITPO900.ptrpant p_pant 
      Height          =   150
      Index           =   255
      Left            =   2430
      TabIndex        =   255
      TabStop         =   0   'False
      Top             =   2160
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      BorderColor     =   -2147483636
   End
   Begin VB.Line Line2 
      X1              =   2160
      X2              =   2160
      Y1              =   2700
      Y2              =   2970
   End
   Begin VB.Label V_SelecPantName 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   540
      TabIndex        =   272
      Top             =   2700
      Width           =   1500
   End
   Begin VB.Label V_FocoPant 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      Height          =   285
      Left            =   2295
      TabIndex        =   271
      Top             =   2700
      Width           =   285
   End
   Begin VB.Label V_SelecPantNumber 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   285
      Left            =   135
      TabIndex        =   266
      Top             =   2700
      Width           =   285
   End
   Begin VB.Line Line1 
      X1              =   135
      X2              =   2565
      Y1              =   3105
      Y2              =   3105
   End
   Begin VB.Label TXT_Escala_H 
      Caption         =   "F0"
      Height          =   285
      Index           =   4
      Left            =   135
      TabIndex        =   265
      Top             =   2160
      Width           =   285
   End
   Begin VB.Label TXT_Escala_H 
      Caption         =   "C0"
      Height          =   285
      Index           =   3
      Left            =   135
      TabIndex        =   264
      Top             =   1755
      Width           =   285
   End
   Begin VB.Label TXT_Escala_H 
      Caption         =   "80"
      Height          =   285
      Index           =   2
      Left            =   135
      TabIndex        =   263
      Top             =   1215
      Width           =   285
   End
   Begin VB.Label TXT_Escala_H 
      Caption         =   "40"
      Height          =   285
      Index           =   1
      Left            =   135
      TabIndex        =   262
      Top             =   675
      Width           =   285
   End
   Begin VB.Label TXT_Escala_H 
      Caption         =   "00"
      Height          =   285
      Index           =   0
      Left            =   135
      TabIndex        =   261
      Top             =   135
      Width           =   285
   End
   Begin VB.Label TXT_Escala_L 
      Caption         =   "F"
      Height          =   285
      Index           =   4
      Left            =   2430
      TabIndex        =   260
      Top             =   2430
      Width           =   150
   End
   Begin VB.Label TXT_Escala_L 
      Caption         =   "C"
      Height          =   285
      Index           =   3
      Left            =   2025
      TabIndex        =   259
      Top             =   2430
      Width           =   150
   End
   Begin VB.Label TXT_Escala_L 
      Caption         =   "8"
      Height          =   285
      Index           =   2
      Left            =   1485
      TabIndex        =   258
      Top             =   2430
      Width           =   150
   End
   Begin VB.Label TXT_Escala_L 
      Caption         =   "4"
      Height          =   285
      Index           =   1
      Left            =   945
      TabIndex        =   257
      Top             =   2430
      Width           =   150
   End
   Begin VB.Label TXT_Escala_L 
      Caption         =   "0"
      Height          =   285
      Index           =   0
      Left            =   405
      TabIndex        =   256
      Top             =   2430
      Width           =   150
   End
End
Attribute VB_Name = "V_Indice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const GS_MAX = 132    ' largo de buffer para leer y escribir ini files

Private Declare Function WritePrivateProfileString& Lib _
"Kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal Lpsection$, ByVal Lpentry$, ByVal _
    buffer$, ByVal filename$)
Private Declare Function GetPrivateProfileString% Lib _
"Kernel32" Alias "GetPrivateProfileStringA" _
   (ByVal Lpsection$, ByVal Lpentry$, _
    ByVal lpDefault$, ByVal buffer$, ByVal _
    cbBuffer%, ByVal lpFileName$)
    
'variables locales para almacenar los valores de las propiedades
Private m_Archivo As String
Private m_NombreArchivo As String
Private m_RutaArchivo As String

Private m_NombreProyecto As String
Private m_Cliente As String
Private m_Comentario As String

Private m_lra_limite_sup As Integer
Private m_lra_limite_inf As Integer
Private m_bit_limite_sup As Integer
Private m_bit_limite_inf As Integer


Private m_Error As Integer
Public m_prj_mod As Boolean

Private N_FocoPant As Integer

Private m_Data_Global As Class_GLOBAL
Public m_Screens As Collection

Public Property Get lra_limite_sup() As Integer
    lra_limite_sup = m_lra_limite_sup
End Property
Public Property Let lra_limite_sup(new_lra_limite_sup As Integer)
    m_lra_limite_sup = new_lra_limite_sup
End Property

Public Property Get lra_limite_inf() As Integer
    lra_limite_inf = m_lra_limite_inf
End Property
Public Property Let lra_limite_inf(new_lra_limite_inf As Integer)
    m_lra_limite_inf = new_lra_limite_inf
End Property

Public Property Get bit_limite_sup() As Integer
    bit_limite_sup = m_bit_limite_sup
End Property
Public Property Let bit_limite_sup(new_bit_limite_sup As Integer)
    m_bit_limite_sup = new_bit_limite_sup
End Property

Public Property Get bit_limite_inf() As Integer
    bit_limite_inf = m_bit_limite_inf
End Property
Public Property Let bit_limite_inf(new_bit_limite_inf As Integer)
    m_bit_limite_inf = new_bit_limite_inf
End Property

Public Property Get dataglobal() As Class_GLOBAL
    Set dataglobal = m_Data_Global
End Property

Public Property Let archivo(ByVal vdata As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.archivo = 5
    
    If m_Archivo = "" Then
        If InStr(vdata, "Nuevo_") Then
            m_Archivo = vdata
        Else
            If Dir(vdata) <> "" Then
                m_Archivo = vdata
            Else
                ret = MsgBox("Se ingreso al Modulo Proyecto con una ruta invalida", vbOKOnly, "CLASS ProyectoPO900")
                m_Error = -1
            End If
        End If
        
        If m_Error = 0 Then
            Me.Show
            m_prj_mod = False
            If InStr(m_Archivo, "Nuevo_") = 0 Then
                load_proyecto
            Else
                Me.Caption = m_Archivo
                m_NombreProyecto = m_Archivo
                m_Cliente = "?"
                m_Comentario = "Proyecto Nuevo"
                Set m_Data_Global = New Class_GLOBAL
                Set m_Screens = New Collection
            End If
        End If
    Else
        ret = MsgBox("El Modulo de Proyecto esta en Uso", vbOKOnly, "CLASS ProyectoPO900")
    End If
End Property


Public Property Get archivo() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.archivo
    archivo = m_Archivo
End Property

Public Property Get path() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.archivo
    path = m_RutaArchivo
End Property

Public Property Get Nombre() As String
    Nombre = m_NombreProyecto
End Property
Public Property Let Nombre(ByVal vdata As String)
    m_NombreProyecto = vdata
    Me.Caption = m_NombreProyecto
    m_prj_mod = True
End Property

Public Property Get Cliente() As String
    Cliente = m_Cliente
End Property
Public Property Let Cliente(ByVal vdata As String)
    m_Cliente = vdata
    m_prj_mod = True
End Property

Public Property Get Comentario() As String
    Comentario = m_Comentario
End Property
Public Property Let Comentario(ByVal vdata As String)
    m_Comentario = vdata
    m_prj_mod = True
End Property

'**************************************  BOTONES ***********************************
Private Sub B_Borrar_Click()
    m_prj_mod = True
    ret = MsgBox("¿Esta seguro que decea eliminar la plantalla " + V_FocoPant + "?", vbOKCancel)
    If ret = vbOK Then
        If p_pant(N_FocoPant).Set_Estado = SC_NUEVO Then
            p_pant(N_FocoPant).Set_Estado = SC_LIBRE
        Else
            p_pant(N_FocoPant).Set_Estado = SC_BORRAR
        End If
        m_Screens.Remove (genidpan(N_FocoPant))
    End If
   
   Actualiza_Botones
End Sub

Private Sub B_Copiar_Click()
    Dim temp As Class_Pantalla
    
    Set temp = m_Screens(genidpan(N_FocoPant))
    
    Set tempscreen = temp.Clone
    
    Actualiza_Botones
End Sub


Private Sub B_Cortar_Click()
    Dim temp As Class_Pantalla
    
    Set temp = m_Screens(genidpan(N_FocoPant))
    
    Set tempscreen = temp.Clone
 
    m_Screens.Remove (genidpan(N_FocoPant))
    
    Select Case p_pant(N_FocoPant).Set_Estado
        Case SC_USADO:
            p_pant(N_FocoPant).Set_Estado = SC_BORRAR
        Case SC_NUEVO:
            p_pant(N_FocoPant).Set_Estado = SC_LIBRE
        Case SC_MODIFICADO:
            p_pant(N_FocoPant).Set_Estado = SC_BORRAR
    End Select
    
    m_prj_mod = True
    
    Actualiza_Botones
End Sub

Private Sub B_Pegar_Click()
    Dim temp As Class_Pantalla
    Dim ret As Integer
    If (p_pant(N_FocoPant).Set_Estado = SC_LIBRE) Or (p_pant(N_FocoPant).Set_Estado = SC_BORRAR) Then
        Set temp = tempscreen.Clone
        temp.idscreen = genidpan(N_FocoPant)
        temp.Numero = (N_FocoPant)
        m_Screens.Add temp, genidpan(N_FocoPant)
        p_pant(N_FocoPant).Set_Estado = SC_NUEVO
        m_prj_mod = True
    Else
        ret = MsgBox("¿Desea sobre escribir esta la pantalla Nº:" + Hex(N_FocoPant) + "?", vbOKCancel)
        If ret = vbOK Then
            m_Screens.Remove (genidpan(N_FocoPant))
            Set temp = tempscreen.Clone
            temp.idscreen = genidpan(N_FocoPant)
            temp.modo = SC_MODIFICADO
            temp.Numero = N_FocoPant
            m_Screens.Add temp, genidpan(N_FocoPant)
            p_pant(N_FocoPant).Set_Estado = SC_MODIFICADO
            m_prj_mod = True
        End If
    End If
    Actualiza_Botones
End Sub

Private Sub Actualiza_Botones()
    Dim indice As Integer
    
    indice = dec(V_FocoPant.Caption)
    
    If tempscreen Is Nothing Then
        B_Pegar.Enabled = False
    Else
        If p_pant(indice).Set_Open = False Then
            B_Pegar.Enabled = True
        Else
            B_Pegar.Enabled = False
        End If
    End If
    
    If (p_pant(indice).Set_Estado = SC_USADO) Or (p_pant(indice).Set_Estado = SC_NUEVO) Or (p_pant(indice).Set_Estado = SC_MODIFICADO) Then
        If p_pant(indice).Set_Open = True Then
            B_Copiar.Enabled = False
            B_Cortar.Enabled = False
            B_Borrar.Enabled = False
        Else
            B_Copiar.Enabled = True
            B_Cortar.Enabled = True
            B_Borrar.Enabled = True
        End If
    Else
        B_Borrar.Enabled = False
        B_Copiar.Enabled = False
        B_Cortar.Enabled = False
    End If
    
End Sub

Private Sub Form_GotFocus()
    Actualiza_Botones
End Sub

'************************************************************************************************

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    V_SelecPantNumber.Caption = "--"
    V_SelecPantName.Caption = ""
End Sub

Private Sub Form_QueryUnload(cancelar As Integer, modoDescarga As Integer)
    Dim ret As Boolean
    idprj = m_Archivo
    ret = False
    
    close_ViewPant_Open
    If m_prj_mod = True Then
        usersel = MsgBox(m_Archivo + Chr(13) + "¿Desea Gaurdar los cambios?", vbYesNoCancel)
        If usersel = vbYes Then
            ret = save_proyecto
            idprj = m_Archivo
            If ret = False Then
                cancelar = 1
            End If
        ElseIf usersel = vbCancel Then
            cancelar = 1
        Else
            ret = True
        End If
    Else
        ret = True
    End If
    If ret = True Then
        If modoDescarga = 0 Or modoDescarga = 1 Then
           libera_proyecto (idprj)
        End If
    End If
End Sub
Private Sub p_pant_Click(Index As Integer)
    If Index <> 255 Then
        For i = 0 To 255
            p_pant(i).Set_Foco = False
        Next i
        V_FocoPant.Caption = Hex(Index)
        N_FocoPant = Index
        p_pant(Index).Set_Foco = True
        Actualiza_Botones
    End If
End Sub

Private Sub p_pant_DblClick(Index As Integer)
    Dim temppant As Class_Pantalla
    Dim ret
    If Index <> 255 Then
        If p_pant(Index).Set_Open = False Then
            p_pant(Index).Set_Open = True
            If (p_pant(Index).Set_Estado = SC_LIBRE) Or (p_pant(Index).Set_Estado = SC_BORRAR) Then
                p_pant(Index).Set_Estado = SC_NUEVO
                Set temppant = New Class_Pantalla
                temppant.Numero = Index
                temppant.idscreen = genidpan(Index)
                m_Screens.Add temppant, genidpan(Index)
            End If
            Set temppant = m_Screens(genidpan(Index))
            Set temppant.m_pantalla = New pantalla
            ret = temppant.m_pantalla.ViewPantalla(Me, Index)
        End If
        Actualiza_Botones
    End If
End Sub

Private Sub p_pant_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim temppant As Class_Pantalla
    If Index <> 255 Then
        V_SelecPantNumber.Caption = Hex(Index)
        If (p_pant(Index).Set_Estado = SC_USADO) Or (p_pant(Index).Set_Estado = SC_USADO) Then
            Set temppant = m_Screens.item(genidpan(Index))
            V_SelecPantName.Caption = temppant.name
        Else
            V_SelecPantName.Caption = ""
        End If
    End If
End Sub

Private Sub load_proyecto()
    Dim temppant As String
    Dim tempviewpant() As String
                
    Set m_Data_Global = New Class_GLOBAL
    Set m_Screens = New Collection
    
    aux = InStrRev(m_Archivo, "\")
    m_NombreArchivo = Mid(m_Archivo, aux + 1)
    m_RutaArchivo = Mid(m_Archivo, 1, aux)
    
    m_NombreProyecto = getoffile(m_Archivo, "PROYECTO", "NOMBRE", m_NombreArchivo)
    m_Cliente = getoffile(m_Archivo, "PROYECTO", "CLIENTE", "none")
    m_Comentario = getoffile(m_Archivo, "PROYECTO", "COMENTARIO", "none")
    
    m_lra_limite_inf = Val(getoffile(m_Archivo, "LIMITES", "WORD_MIN", "0"))
    m_lra_limite_sup = Val(getoffile(m_Archivo, "LIMITES", "WORD_MAX", "448"))
    m_bit_limite_inf = Val(getoffile(m_Archivo, "LIMITES", "BIT_MIN", "449"))
    m_bit_limite_sup = Val(getoffile(m_Archivo, "LIMITES", "BIT_MAX", "512"))
    
    Me.Caption = m_NombreProyecto
        
    ret = read_gobales(m_Data_Global, m_RutaArchivo + "SCRNS\globales.ini")
    
    For i = 0 To 15
        temppant = getoffile(m_Archivo, "SCREENS", "G_" + Hex(i), "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0")
        tempviewpant = Split(temppant, ",")
        For j = 0 To 15
            If tempviewpant(j) = "1" Then
                ret = read_Pantalla(m_Screens, ((i * 16) + j), m_RutaArchivo + "SCRNS\pant" + Format$((((i * 16) + j)), "00") + ".ini")
                Me.p_pant((i * 16) + j).Set_Estado = SC_USADO
            Else
                Me.p_pant((i * 16) + j).Set_Estado = SC_LIBRE
            End If
        Next j
    Next i

End Sub

Public Function save_proyecto(Optional save_as As Boolean = False) As Boolean
    Dim straux As String
    Dim temp As V_Indice
    Dim arraux(0 To 15) As String
    Dim filename As String
    Dim file As Integer
    Dim col As Integer
    Dim origen As Class_Pantalla
    Dim archivo_temp As String
    Dim ret
    
    save_proyecto = True
    
    For file = 0 To 15
        For col = 0 To 15
            If (Me.p_pant((file * 16) + col).Set_Open = True) Then
                save_proyecto = False
            End If
        Next col
    Next file
    
    If save_proyecto = False Then
        ret = MsgBox("Existen pantalla abiertas antes de guardar se cerraran", vbOKCancel Or vbCritical)
        If ret = vbCancel Then
            Exit Function
        Else
            save_proyecto = True
        End If
    End If
    
    
    If InStr(m_Archivo, "Nuevo_") Or save_as Then
        filename = proyecto_guardar_como(m_Archivo)
        If (Dir(filename) <> "") And (filename <> "") Then
            Set temp = Proyectos(m_Archivo)
            Proyectos.Remove m_Archivo
            m_Archivo = filename
            Proyectos.Add temp, m_Archivo
        Else
            save_proyecto = False
        End If
    End If
    
    If save_proyecto = True Then
        aux = InStrRev(m_Archivo, "\")
        m_NombreArchivo = Mid(m_Archivo, aux + 1)
        m_RutaArchivo = Mid(m_Archivo, 1, aux)
        If Dir(m_Archivo) <> "" Then
            ret = puttofile(m_Archivo, "PROYECTO", "NOMBRE", m_NombreProyecto)
            ret = puttofile(m_Archivo, "PROYECTO", "CLIENTE", m_Cliente)
            ret = puttofile(m_Archivo, "PROYECTO", "COMENTARIO", m_Comentario)
            ret = puttofile(m_Archivo, "LIMITES", "WORD_MIN", CStr(m_lra_limite_inf))
            ret = puttofile(m_Archivo, "LIMITES", "WORD_MAX", CStr(m_lra_limite_sup))
            ret = puttofile(m_Archivo, "LIMITES", "BIT_MIN", CStr(m_bit_limite_inf))
            ret = puttofile(m_Archivo, "LIMITES", "BIT_MAX", CStr(m_bit_limite_sup))
            ret = write_gobales(m_Data_Global, m_RutaArchivo + "SCRNS\globales.ini")

            For file = 0 To 15
                For col = 0 To 15
                    If (Me.p_pant((file * 16) + col).Set_Open = True) Then
                        Set origen = m_Screens.item(genidpan((file * 16) + col))
                        Unload origen.m_pantalla
                    End If
                    
                    If (Me.p_pant((file * 16) + col).Set_Estado <> SC_LIBRE) Then
                        If (Me.p_pant((file * 16) + col).Set_Estado = SC_BORRAR) Then
                            archivo_temp = m_RutaArchivo + "SCRNS\pant" + Format$(((file * 16) + col), "00") + ".ini"
                            If Dir(archivo_temp) <> "" Then
                                Kill (archivo_temp)
                            Else
                                MsgBox ("Error al salvar el proyecto" + Chr$(13) + "Falta:" + archivo_temp)
                            End If
                            Me.p_pant((file * 16) + col).Set_Estado = SC_LIBRE
                            arraux(col) = "0"
                        ElseIf (Me.p_pant((file * 16) + col).Set_Estado = SC_USADO) Then
                            Set origen = m_Screens.item(genidpan((file * 16) + col))
                            origen.modo = SC_USADO
                            Me.p_pant((file * 16) + col).Set_Estado = SC_USADO
                            arraux(col) = "1"
                        Else
                            Set origen = m_Screens.item(genidpan((file * 16) + col))
                            origen.modo = SC_USADO
                            Me.p_pant((file * 16) + col).Set_Estado = SC_USADO
                            ret = write_pantalla(origen, m_RutaArchivo + "SCRNS\pant" + Format$(((file * 16) + col), "00") + ".ini")
                            arraux(col) = "1"
                        End If
                    Else
                        arraux(col) = "0"
                    End If
                Next col
                straux = Join(arraux, ",")
                ret = puttofile(m_Archivo, "SCREENS", "G_" + Hex(file), straux)
            Next file
            m_prj_mod = False
           Else
            MsgBox ("Guardar como error de nombre de archivo")
        End If
    End If
End Function

Private Sub V_FocoPant_Click()
    p_pant(N_FocoPant).Set_Estado = SC_MODIFICADO
    m_prj_mod = True
End Sub

Public Sub close_view_pantalla(id_screen As Byte)
    Dim tempscreen As Class_Pantalla
    Set tempscreen = m_Screens.item(genidpan(id_screen))
    If (tempscreen.modo = SC_MODIFICADO) Or (tempscreen.modo = SC_NUEVO) Then
        m_prj_mod = True
    ElseIf (tempscreen.modo = SC_LIBRE) Or (tempscreen.modo = SC_BORRAR) Then
        m_Screens.Remove (genidpan(id_screen))
    End If
    p_pant(id_screen).Set_Open = False
    p_pant(id_screen).Set_Estado = tempscreen.modo
    Actualiza_Botones
End Sub

Private Sub close_ViewPant_Open()
    Dim tempscreen As Class_Pantalla
    For i = 0 To 254
        If p_pant(i).Set_Open = True Then
            Set tempscreen = m_Screens.item(genidpan(i))
            Unload tempscreen.m_pantalla
        End If
    Next i
    
End Sub
