VERSION 5.00
Begin VB.Form emupantalla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel de Operador"
   ClientHeight    =   6450
   ClientLeft      =   6360
   ClientTop       =   4800
   ClientWidth     =   6450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "main"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "emupantalla.frx":0000
   ScaleHeight     =   6450
   ScaleWidth      =   6450
   Begin VB.Timer Timer_salto 
      Left            =   6000
      Top             =   1200
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000C0C0&
      Height          =   1580
      Index           =   0
      Left            =   1200
      ScaleHeight     =   1515
      ScaleWidth      =   4260
      TabIndex        =   24
      Top             =   960
      Width           =   4320
      Begin EDITPO900.lb_cont lb_contro 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "AB"
         BackStyle       =   0
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   3  'Dot
         Index           =   18
         X1              =   3604
         X2              =   3604
         Y1              =   0
         Y2              =   1520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   3  'Dot
         Index           =   17
         X1              =   3816
         X2              =   3816
         Y1              =   0
         Y2              =   1520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   3  'Dot
         Index           =   16
         X1              =   4028
         X2              =   4028
         Y1              =   0
         Y2              =   1520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   3  'Dot
         Index           =   15
         X1              =   3392
         X2              =   3392
         Y1              =   0
         Y2              =   1520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   3  'Dot
         Index           =   14
         X1              =   3180
         X2              =   3180
         Y1              =   0
         Y2              =   1520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   3  'Dot
         Index           =   13
         X1              =   2968
         X2              =   2968
         Y1              =   0
         Y2              =   1520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   3  'Dot
         Index           =   12
         X1              =   2756
         X2              =   2756
         Y1              =   0
         Y2              =   1520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   3  'Dot
         Index           =   11
         X1              =   2332
         X2              =   2332
         Y1              =   0
         Y2              =   1520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   3  'Dot
         Index           =   10
         X1              =   2120
         X2              =   2120
         Y1              =   0
         Y2              =   1520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   3  'Dot
         Index           =   9
         X1              =   1696
         X2              =   1696
         Y1              =   0
         Y2              =   1520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   3  'Dot
         Index           =   8
         X1              =   1484
         X2              =   1484
         Y1              =   0
         Y2              =   1520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   3  'Dot
         Index           =   7
         X1              =   1272
         X2              =   1272
         Y1              =   0
         Y2              =   1520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   3  'Dot
         Index           =   6
         X1              =   1060
         X2              =   1060
         Y1              =   0
         Y2              =   1520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   3  'Dot
         Index           =   5
         X1              =   848
         X2              =   848
         Y1              =   0
         Y2              =   1520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   3  'Dot
         Index           =   4
         X1              =   636
         X2              =   636
         Y1              =   0
         Y2              =   1520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   3  'Dot
         Index           =   3
         X1              =   1908
         X2              =   1908
         Y1              =   0
         Y2              =   1520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   3  'Dot
         Index           =   2
         X1              =   2544
         X2              =   2544
         Y1              =   0
         Y2              =   1520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   3  'Dot
         Index           =   1
         X1              =   424
         X2              =   424
         Y1              =   0
         Y2              =   1520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   3  'Dot
         Index           =   0
         X1              =   212
         X2              =   212
         Y1              =   0
         Y2              =   1520
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   2
         X1              =   0
         X2              =   4800
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   4800
         Y1              =   760
         Y2              =   760
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   4800
         Y1              =   380
         Y2              =   380
      End
   End
   Begin VB.Timer Timer1 
      Left            =   6000
      Top             =   720
   End
   Begin VB.Label lb_open 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lb_dot 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   4800
      TabIndex        =   23
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label lb_mas_menos 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   3240
      TabIndex        =   22
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label lb_func 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   3
      Left            =   5640
      TabIndex        =   21
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label lb_func 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   2
      Left            =   5640
      TabIndex        =   20
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lb_func 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   1
      Left            =   5640
      TabIndex        =   19
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lb_func 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   0
      Left            =   5640
      TabIndex        =   18
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lb_nums 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   9
      Left            =   4680
      TabIndex        =   17
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lb_nums 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   8
      Left            =   3960
      TabIndex        =   16
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lb_nums 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   7
      Left            =   3240
      TabIndex        =   15
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lb_nums 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   6
      Left            =   4680
      TabIndex        =   14
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lb_nums 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   5
      Left            =   3960
      TabIndex        =   13
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lb_nums 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   4
      Left            =   3240
      TabIndex        =   12
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lb_nums 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   3
      Left            =   4800
      TabIndex        =   11
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lb_nums 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   2
      Left            =   3960
      TabIndex        =   10
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lb_nums 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   1
      Left            =   3240
      TabIndex        =   9
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lb_nums 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   0
      Left            =   3960
      TabIndex        =   8
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label ld_dir 
      BackStyle       =   0  'Transparent
      Height          =   735
      Index           =   3
      Left            =   480
      TabIndex        =   7
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label ld_dir 
      BackStyle       =   0  'Transparent
      Height          =   735
      Index           =   2
      Left            =   1200
      TabIndex        =   6
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label ld_dir 
      BackStyle       =   0  'Transparent
      Height          =   735
      Index           =   1
      Left            =   2040
      TabIndex        =   5
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label ld_dir 
      BackStyle       =   0  'Transparent
      Height          =   735
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lb_esc 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lb_enter 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   2160
      TabIndex        =   2
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label LB_END 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Lb_teclas 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
   Begin VB.Shape sh_status 
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   315
      Top             =   2040
      Width           =   195
   End
   Begin VB.Shape sh_power 
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   315
      Top             =   1240
      Width           =   200
   End
End
Attribute VB_Name = "emupantalla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim coll_campos As Collection
Dim jmp_N(9) As Integer
Dim jmp_F(3) As Integer
Dim tabindex As Byte
Dim indice As Byte
Dim editando As Boolean
Dim primera_vez As Boolean
Dim path As String
Dim jmp_esc As Byte
Dim temp_caption As String
Dim historial_pantalla(5) As Byte
Dim cuenta_pant As Byte
Dim cuenta_esc As Byte
Dim escapando As Boolean
Dim tiempo_salto As Long
Dim jmp_tiempo As Byte

Public Function leer_mia(np_ix As Byte) As Boolean
    Dim j As Byte
    Dim i As Integer
    Dim verifica As Long
    Dim archivo_ini As String
    Dim sep() As String
    Dim aux As String
    Dim seccion As String
    Dim default As String
    Dim key_item As String
    Dim xt As Boolean
    Dim dataitem As String
    Dim item As String
    Dim cuenta_campos  As Integer
    
    Timer_salto.Enabled = False
    archivo_ini = path & "\SCRNS" + "\" + "globales.ini"
    xt = jump_teclas(archivo_ini, True)
    archivo_ini = path & "\SCRNS" + "\" + "pant" + Format$(np_ix, "00") + ".ini"
    If coll_campos Is Nothing Then Set coll_campos = New Collection
    historial_pantalla(cuenta_pant) = np_ix
    If escapando = False Then
        cuenta_pant = cuenta_pant + 1
        If cuenta_pant > 4 Then cuenta_pant = 0
    End If
    
    For j = 1 To 16
        seccion$ = "CAMPO" + Format$(j, "00")
        dataitem = getitem(getoffile(archivo_ini, seccion$, "TIPO", ""))

        If dataitem <> "" Then
            Select Case dataitem
                Case "CTEXT"
                    Set class_temp = New class_CTEXT
                    coll_campos.Add class_temp, "campo CT-" & j
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "TEXTO", "CAMPO" + Format$(j, "00")))
                    class_temp.texto = dataitem
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "POSXY", "1 1"))
                    sep = Split(dataitem, " ", 2)
                    class_temp.x_pos = sep(0)
                    class_temp.y_pos = sep(1)
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "ATRIBUTOS", "00"))
                    class_temp.atributo = dataitem
                
                Case "MTEXT"
                    
                    Set class_temp = New Class_MTEXT
                    coll_campos.Add class_temp, "campo MT-" & j

                    dataitem = getitem(getoffile(archivo_ini, seccion$, "TEXTO", "CAMPO" + Format$(j, "00")))
                    class_temp.texto = dataitem
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "POSXY", "1 1"))
                    sep = Split(dataitem, " ", 2)
                    class_temp.x_pos = sep(0)
                    class_temp.y_pos = sep(1)
                  
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "ATRIBUTOS", "00"))
                    class_temp.atributo = dataitem
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "JMP", "00"))
                    class_temp.jmp = dataitem
                
                Case "ALFANUM"
                    Set class_temp = New Class_ALFA
                    coll_campos.Add class_temp, "campo MT-" & j
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "POSXY", "1 1"))
                    sep = Split(dataitem, " ", 2)
                    class_temp.x_pos = sep(0)
                    class_temp.y_pos = sep(1)
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "ATRIBUTOS", "00"))
                    class_temp.atributo = dataitem
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "LEN", "5"))
                    class_temp.largo = dataitem

                    dataitem = getitem(getoffile(archivo_ini, seccion$, "EDIT", "OFF"))
                    If dataitem = "OFF" Then
                        class_temp.Edit = "False"
                    Else
                        class_temp.Edit = "true"
                    End If
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "LRA", ""))
                    class_temp.LRA = dataitem
                
                Case "MTDIGITAL"
                    Set class_temp = New Class_MTD
                    coll_campos.Add class_temp, "campo TD-" & j
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "POSXY", "1 1"))
                    sep = Split(dataitem, " ", 2)
                    class_temp.x_pos = sep(0)
                    class_temp.y_pos = sep(1)
                    class_temp.name = "0"
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "ATRIBUTOS", "00"))
                    class_temp.atributo = dataitem
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "LRA", ""))
                    class_temp.LRA = dataitem
                
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "LEN", "2"))
                    class_temp.largo = dataitem
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "MT_ITEMS", "2"))
                    class_temp.items = dataitem
                    
                    For i = 0 To (Val(class_temp.items) - 1)
                        default$ = ""
                        item$ = "MTDATA" + Format$(i, "00")
                        dataitem = getitem(getoffile(archivo_ini, seccion$, item$, ""))
                        class_temp.texto(i) = dataitem
                    Next i
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "EDIT", "OFF"))
                    If dataitem = "OFF" Then
                        class_temp.Edit = "False"
                    Else
                        class_temp.Edit = "true"
                    End If
                
                Case "NUMERICO"
                    Set class_temp = New Class_NUM
                    coll_campos.Add class_temp, "campo NU-" & j
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "DEC", "2"))
                    class_temp.dec = dataitem

                    dataitem = getitem(getoffile(archivo_ini, seccion$, "LEN", "2"))
                    class_temp.largo = dataitem
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "EDIT", "OFF"))
                    If dataitem = "OFF" Then
                        class_temp.Edit = "False"
                    Else
                        class_temp.Edit = "true"
                    End If

                    dataitem = getitem(getoffile(archivo_ini, seccion$, "POSXY", "1 1"))
                    sep = Split(dataitem, " ", 2)
                    class_temp.x_pos = sep(0)
                    class_temp.y_pos = sep(1)
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "ATRIBUTOS", "00"))
                    class_temp.atributo = dataitem
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "LRA", ""))
                    class_temp.LRA = dataitem
                
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "NUMERIC_MODE", "0"))
                    class_temp.modo = dataitem
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "GAIN_EXP", "OFF"))
                    If dataitem = "OFF" Then
                        class_temp.recta = "False"
                    Else
                        class_temp.recta = "true"
                    End If
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "GAIN", ""))
                    class_temp.gain = Val(dataitem)
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "OFFSET", "0"))
                    class_temp.offset = dataitem
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "RANGOX", "2048 0"))
                    sep = Split(dataitem, " ")
                    class_temp.rangoX(0) = Val(sep(1))
                    class_temp.rangoX(1) = Val(sep(0))
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "RANGOY", "2048 0"))
                    sep = Split(dataitem, " ")
                    class_temp.rangoY(0) = Val(sep(1))
                    class_temp.rangoY(1) = Val(sep(0))
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "MIN", "0000,00"))
                    class_temp.minimo = dataitem
                    
                    dataitem = getitem(getoffile(archivo_ini, seccion$, "MAX", "0000,00"))
                    class_temp.maximo = Val(dataitem)
                    
            End Select
        End If
    Next j
    tabindex = 0
    For j = 1 To coll_campos.count
        Set class_temp = coll_campos(j)
        Load lb_contro(j)
        cuenta_campos = cuenta_campos + 1
        Set emupantalla.lb_contro(j).cl_campo = coll_campos(j)
        lb_contro(j).Visible = True
        lb_contro(j).Top = ((lb_contro(j).cl_campo.y_pos - 1) * 380) + 15
        lb_contro(j).Left = ((lb_contro(j).cl_campo.x_pos - 1) * 212)
        lb_contro(j).Width = (lb_contro(j).cl_campo.largo * 212) + 20
        lb_contro(j).Caption = lb_contro(j).cl_campo.texto
        lb_contro(j).TabStop = False
        Select Case class_temp.tipo_campo
            Case "CTEXT"
                lb_contro(j).Tag = "campo CT-" & j
            Case "MTEXT"
                lb_contro(j).Tag = "campo MT-" & j
                tabindex = tabindex + 1
                lb_contro(j).TabStop = True
                lb_contro(j).tabindex = tabindex
            Case "ALFANUM"
                emupantalla.lb_contro(j).Tag = "campo AN-" & j
                If lb_contro(j).cl_campo.Edit = True Then
                    tabindex = tabindex + 1
                    lb_contro(j).TabStop = True
                    lb_contro(j).tabindex = tabindex
                End If
            Case "MTDIGITAL"
                lb_contro(j).Tag = "campo TD-" & j
                lb_contro(j).Caption = lb_contro(j).cl_campo.texto(0)
                If lb_contro(j).cl_campo.Edit = True Then
                    tabindex = tabindex + 1
                    lb_contro(j).TabStop = True
                    lb_contro(j).tabindex = tabindex
                End If
            Case "NUMERICO"
                lb_contro(j).Tag = "campo NU-" & j
                lb_contro(j).Caption = cambia(lb_contro(j).Caption)
                If lb_contro(j).cl_campo.Edit = True Then
                    tabindex = tabindex + 1
                    lb_contro(j).TabStop = True
                    lb_contro(j).tabindex = tabindex
                End If
        End Select
        If j = 1 Then
            For i = 0 To 18
                Line4(i).Visible = False
            Next i
            
            For i = 0 To 2
                Line1(i).BorderColor = &HC0FFFF
                Line1(i).BorderStyle = 3
                Line1(i).BorderWidth = 1
            Next i
        End If
    Next j
    xt = jump_teclas(archivo_ini, False)
    If lb_contro.count > 1 Then
        lb_contro(1).SetFocus
        ld_dir_Click (1)
    End If
    If tiempo_salto > 64 Then tiempo_salto = 64
    Timer_salto.Interval = tiempo_salto * 1000
    Timer_salto.Enabled = True
End Function

Private Function jump_teclas(ByVal archivo_ini As String, ByVal globales As Boolean) As Boolean
    Dim key_item As String
    Dim tecla As Byte
    Dim dataitem As String
    Dim variable As String
    Dim separa() As String
    If globales = True Then
        variable = "VARIABLES_DEFAULT"
    Else
        variable = "VARIABLES"
    End If
    
    dataitem = getitem(getoffile(archivo_ini, variable, "TACTIVO", "0 0 255"))
    separa = Split(dataitem, " ")
    tiempo_salto = Val(separa(0))
    jmp_tiempo = Val(separa(1))
    jmp_esc = separa(2)
    
    For tecla = 0 To 9
        key_item = "JMP_TECLA_" + Mid(Str(tecla), 2, 1)
        dataitem = getitem(getoffile(archivo_ini, variable, key_item, "256"))
        If dataitem <> "" Then
            If globales = True Then
                jmp_N(tecla) = Val(dataitem)
            Else
                If Val(dataitem) <> 256 Then jmp_N(tecla) = Val(dataitem)
            End If
        End If
    Next tecla
    
    For tecla = 1 To 4
        key_item = "JMP_TECLA_F" + Mid(Str(tecla), 2, 1)
        dataitem = getitem(getoffile(archivo_ini, variable, key_item, "256"))
        If dataitem <> "" Then
            If globales = True Then
                jmp_F(tecla - 1) = Val(dataitem)
            Else
                If 256 And Val(dataitem) <> 256 Then jmp_F(tecla - 1) = Val(dataitem)
            End If
        End If
    Next tecla
    
    For tecla = 0 To 9
        If jmp_N(tecla) <> 256 Then
            lb_nums(tecla).ToolTipText = "Jump To: " & CStr(Hex(jmp_N(tecla)))
        Else
            lb_nums(tecla).ToolTipText = "No Jump"
        End If
    Next tecla
    
    For tecla = 0 To 3
        If jmp_F(tecla) <> 256 Then
            lb_func(tecla).ToolTipText = "Jump To: " & CStr(Hex(jmp_F(tecla)))
        Else
            lb_func(tecla).ToolTipText = "No Jump"
        End If
    Next tecla
    jump_teclas = True
End Function

Private Sub Form_Load()
    Dim j As Byte
    For j = 0 To 9
        jmp_N(j) = 256
    Next j
    For j = 0 To 3
        jmp_F(j) = 256
    Next j
    Set coll_campos = New Collection
    Timer1.Interval = 500
    sh_power.BackColor = RGB(255, 0, 0)
    sh_status.BackColor = RGB(0, 0, 0)
End Sub

Private Sub lb_contro_GotFocus(Index As Integer)
    Dim aux() As String
    Dim aux2 As String
    indice = Index
    
    If lb_contro(Index).TabStop = False Then
        ld_dir_Click (1)
    Else
        temp_caption = lb_contro(Index).Caption
        lb_contro(Index).Caption = ">" & lb_contro(Index).Caption
        lb_contro(Index).Width = lb_contro(Index).Width + 212
        lb_contro(Index).Left = lb_contro(Index).Left - 212
        lb_contro(Index).ZOrder
    End If
    
    If lb_contro(Index).cl_campo.tipo_campo <> "CTEXT" Then
        If lb_contro(Index).cl_campo.tipo_campo = "MTEXT" Then
            aux2 = lb_contro(Index).cl_campo.jmp
            aux2 = Hex(aux2)
            Lb_teclas.Caption = "Jump To:" & CStr(aux2)
        Else
            aux2 = lb_contro(Index).cl_campo.LRA
            Lb_teclas.Caption = aux2
        End If
    Else
        Lb_teclas.Caption = ""
    End If
    
End Sub

Private Sub lb_contro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim saltar_a As Integer
    Dim j As Integer
    
    If KeyCode = 13 And editando = False Then
        If lb_contro(Index).cl_campo.tipo_campo <> "MTEXT" Then
            ld_dir_Click (1)
        Else
            saltar_a = lb_contro(indice).cl_campo.jmp
            If Not coll_campos Is Nothing Then
                For j = 1 To coll_campos.count
                    Unload lb_contro(j)
                Next j
            Set coll_campos = Nothing
            End If
            indice = 0
            leer_mia (saltar_a)
        End If
    End If
    If editando = True And lb_contro(Index).cl_campo.tipo_campo = "ALFANUM" Then
        If KeyCode = 8 Then
            If Len(lb_contro(Index).Caption) > 0 Then
                lb_contro(Index).Caption = Mid(lb_contro(Index).Caption, 1, Len(lb_contro(Index).Caption) - 1)
            End If
        Else
            lb_contro(Index).Caption = lb_contro(Index).Caption & Chr(KeyCode)
            temp_caption = Mid(lb_contro(Index).Caption, 2, Len(lb_contro(Index).Caption))
        End If
    End If
    If editando = True And lb_contro(Index).cl_campo.tipo_campo = "NUMERICO" Then
        If KeyCode = 8 Then
            If Len(lb_contro(Index).Caption) > 0 Then
                lb_contro(Index).Caption = Mid(lb_contro(Index).Caption, 1, Len(lb_contro(Index).Caption) - 1)
            End If
        Else
            Dim pos_dot As Integer
            If (KeyCode > 48 And KeyCode < 59) Or KeyCode = 190 Then
                lb_contro(Index).Caption = lb_contro(Index).Caption & Chr(KeyCode)
                temp_caption = Mid(lb_contro(Index).Caption, 2, Len(lb_contro(Index).Caption))
            End If
        End If
    End If

End Sub

Private Sub lb_contro_LostFocus(Index As Integer)
    If lb_contro(Index).TabStop = True Then
        lb_contro(Index).Caption = temp_caption
        lb_contro(Index).Width = lb_contro(Index).Width - 212
        lb_contro(Index).Left = lb_contro(Index).Left + 212
    End If
End Sub

Private Sub LB_END_Click()
    Unload Me
End Sub

Private Sub lb_enter_Click()
    Dim saltar_a As Byte
    Dim j As Byte
    If coll_campos.count > 0 Then
        If lb_contro(indice).cl_campo.tipo_campo <> "MTEXT" Then
            If lb_contro(indice).cl_campo.Edit = True Then
                If editando = True Then
                    editando = False
                    lb_contro(indice).BackStyle = 0
                    lb_contro(indice).BackColor = &HC0C0&
                Else
                    editando = True
                    lb_contro(indice).BackStyle = 1
                    lb_contro(indice).BackColor = &H8000&
                End If
            End If
        Else
            saltar_a = lb_contro(indice).cl_campo.jmp
            If Not coll_campos Is Nothing Then
                For j = 1 To coll_campos.count
                    Unload lb_contro(j)
                Next j
            Set coll_campos = Nothing
            End If
            indice = 0
            leer_mia (saltar_a)
        End If
    End If
End Sub

Private Sub lb_esc_Click()
    Dim j As Byte
        If editando = True Then
            editando = False
            lb_contro(indice).BackStyle = 0
            lb_contro(indice).BackColor = &HC0C0&
        Else
            If jmp_esc = 255 Then
                escapando = True
                If coll_campos.count > 0 Then borra
                indice = 0
                If cuenta_pant = 0 Then
                    jmp_esc = historial_pantalla(3)
                    cuenta_pant = 5
                Else
                    If cuenta_pant = 1 Then
                        jmp_esc = historial_pantalla(4)
                        cuenta_pant = cuenta_pant - 1
                    Else
                        jmp_esc = historial_pantalla(cuenta_pant - 2)
                        cuenta_pant = cuenta_pant - 1
                    End If
                End If
            End If
            indice = 0
            borra
            leer_mia (jmp_esc)
            escapando = False
        End If
End Sub

Private Sub lb_func_Click(Index As Integer)
    tiempo_salto = 0
    Select Case Index
        Case 0
            If jmp_F(0) <> 256 Then
                indice = 0
                borra
                leer_mia (jmp_F(0))
            End If
        Case 1
            If jmp_F(1) <> 256 Then
                indice = 0
                borra
                leer_mia (jmp_F(1))
            End If
        Case 2
            If jmp_F(2) <> 256 Then
                indice = 0
                borra
                leer_mia (jmp_F(2))
            End If
        Case 3
            If jmp_F(3) <> 256 Then
                indice = 0
                borra
                leer_mia (jmp_F(3))
            End If
    End Select
End Sub

Private Sub lb_nums_Click(Index As Integer)
    If lb_contro(indice).BorderStyle = 0 Then
        Select Case Index
            Case 0
                If jmp_N(0) <> 256 Then
                    indice = 0
                    borra
                    leer_mia (jmp_N(0))
                End If
            Case 1
                If jmp_N(1) <> 256 Then
                    indice = 0
                    borra
                    leer_mia (jmp_N(1))
                End If
            Case 2
                If jmp_N(2) <> 256 Then
                    indice = 0
                    borra
                    leer_mia (jmp_N(2))
                End If
            Case 3
                If jmp_N(3) <> 256 Then
                    indice = 0
                    borra
                    leer_mia (jmp_N(3))
                End If
            Case 4
                If jmp_N(4) <> 256 Then
                    indice = 0
                    borra
                    leer_mia (jmp_N(4))
                End If
            Case 5
                If jmp_N(5) <> 256 Then
                    indice = 0
                    borra
                    leer_mia (jmp_N(5))
                End If
            Case 6
                If jmp_N(6) <> 256 Then
                    indice = 0
                    borra
                    leer_mia (jmp_N(6))
                End If
            Case 7
                If jmp_N(7) <> 256 Then
                    indice = 0
                    borra
                    leer_mia (jmp_N(7))
                End If
            Case 8
                If jmp_N(8) <> 256 Then
                    indice = 0
                    borra
                    leer_mia (jmp_N(8))
                End If
            Case 9
                If jmp_N(9) <> 256 Then
                    indice = 0
                    borra
                    leer_mia (jmp_N(9))
                End If
        End Select
    Else
        Select Case lb_contro(indice).cl_campo.tipo_campo
            Case "NUMERICO"
        End Select
    
    End If
    
End Sub

Private Sub lb_open_DblClick()
    Dim tempo As Long
    Dim auxiliar As String
    Dim auxiliar2 As String
    On Error Resume Next
    EditorIDE.ComunWindows.ShowOpen
    auxiliar = EditorIDE.ComunWindows.filename
    auxiliar2 = EditorIDE.ComunWindows.FileTitle
    tempo = (InStr(1, auxiliar, auxiliar2)) - 2
    path = Mid(auxiliar, 1, tempo)
    
    If path = "" Then path = App.path
    borra
    leer_mia (0)
End Sub

Private Sub ld_dir_Click(Index As Integer)
    Dim X As Byte
    Dim cuenta As Byte
    Dim cuenta2 As Byte
    Dim count As Byte
    
    If coll_campos.count > 0 Then
        Select Case Index
            Case 1
                If lb_contro(indice).BackStyle = 0 Then
                    Do Until (lb_contro(indice).TabStop = True And X = 1) Or count > 16
                        indice = indice + 1
                        If indice > coll_campos.count Then indice = 1
                        X = 1
                        count = count + 1
                    Loop
                    If lb_contro(indice).TabStop = True Then lb_contro(indice).SetFocus
                End If
            Case 3
                If lb_contro(indice).BackStyle = 0 Then
                    Do Until (lb_contro(indice).TabStop = True And X = 1) Or count > 16
                        indice = indice - 1
                        If indice < 1 Then indice = coll_campos.count
                        X = 1
                        count = count + 1
                    Loop
                    lb_contro(indice).SetFocus
                End If
            Case 0
                If lb_contro(indice).BackStyle = 1 Then
                    If lb_contro(indice).cl_campo.tipo_campo = "MTDIGITAL" Then
                        If lb_contro(indice).cl_campo.name <> "0" Then
                            cuenta = Val(lb_contro(indice).cl_campo.name)
                            cuenta = cuenta - 1
                            lb_contro(indice).cl_campo.name = CStr(cuenta)
                            lb_contro(indice).Caption = lb_contro(indice).cl_campo.texto(cuenta)
                        End If
                    End If
                End If
                
            Case 2
                If lb_contro(indice).BackStyle = 1 Then
                    If lb_contro(indice).cl_campo.tipo_campo = "MTDIGITAL" Then
                        cuenta2 = Val(lb_contro(indice).cl_campo.items) - 1
                        If lb_contro(indice).cl_campo.name <> CStr(cuenta2) Then
                            cuenta = Val(lb_contro(indice).cl_campo.name)
                            cuenta = cuenta + 1
                            lb_contro(indice).cl_campo.name = CStr(cuenta)
                            lb_contro(indice).Caption = lb_contro(indice).cl_campo.texto(cuenta)
                        End If
                    End If
                End If
        End Select
    End If
End Sub
Private Sub borra()
     Dim j As Byte
    If Not coll_campos Is Nothing Then
        For j = 1 To coll_campos.count
            Unload lb_contro(j)
        Next j
        Set coll_campos = Nothing
        For j = 0 To 18
            Line4(j).Visible = True
        Next j
        For j = 0 To 2
            Line1(j).BorderColor = &H80000008
            Line1(j).BorderStyle = 1
            Line1(j).BorderWidth = 3
        Next j
    End If
End Sub
Private Function cambia(ByVal texto As String) As String
    Dim j As Byte
    Dim aux  As String
    
    For j = 1 To Len(texto)
        aux = Mid(texto, j, 1)
        If aux = "#" Then aux = "0"
        Mid(texto, j, 1) = aux
    Next j
    cambia = texto
End Function

Private Sub Timer_salto_Timer()
    indice = 0
    borra
    leer_mia (jmp_tiempo)
End Sub

Private Sub Timer1_Timer()
    If sh_status.BackColor = RGB(0, 0, 0) Then
        sh_status.BackColor = RGB(0, 255, 0)
    Else
        sh_status.BackColor = RGB(0, 0, 0)
    End If
End Sub


