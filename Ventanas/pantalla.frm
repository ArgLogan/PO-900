VERSION 5.00
Begin VB.Form pantalla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel de Operador"
   ClientHeight    =   2940
   ClientLeft      =   6360
   ClientTop       =   4800
   ClientWidth     =   6450
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "pantalla.frx":0000
   LinkTopic       =   "main"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "pantalla.frx":0ECA
   ScaleHeight     =   2940
   ScaleWidth      =   6450
   Begin VB.CommandButton CH_SIZE 
      Caption         =   "Teclado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5400
      TabIndex        =   8
      Top             =   2640
      Width           =   975
   End
   Begin VB.PictureBox Pic_display 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   1200
      ScaleHeight     =   4
      ScaleMode       =   0  'User
      ScaleWidth      =   20
      TabIndex        =   5
      Top             =   960
      Width           =   3975
      Begin EDITPO900.lb_cont lb_control 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "PatternLCD"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line LN_col 
         BorderStyle     =   3  'Dot
         Index           =   0
         X1              =   1.217
         X2              =   1.217
         Y1              =   -0.368
         Y2              =   8.828
      End
      Begin VB.Line LN_fila 
         Index           =   0
         X1              =   5.475
         X2              =   31.027
         Y1              =   2.207
         Y2              =   2.207
      End
   End
   Begin VB.Timer Timer1 
      Left            =   6000
      Top             =   720
   End
   Begin VB.Label PropRun 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label ViewCampos 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   1200
      TabIndex        =   7
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lb_enter 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   2160
      TabIndex        =   4
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Y: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2640
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
      BackStyle       =   0  'Transparent
      Caption         =   "Tecla"
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
      Width           =   5295
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
Attribute VB_Name = "pantalla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cuenta_campos As Byte
Private coll_campos As Collection
Private screen_ref As Class_Pantalla
Private propietario As V_Indice

Public Function ViewPantalla(ByRef prop As V_Indice, ByVal indice_pant As Integer)
    Set propietario = prop
    Set screen_ref = propietario.m_Screens.item(genidpan(indice_pant))
    Set coll_campos = screen_ref.colectcampo
    cuenta_campos = 0
    
    Me.Show
    Lb_teclas.Caption = Hex(screen_ref.Numero) + " : " + screen_ref.name
    While cuenta_campos < coll_campos.count
        cuenta_campos = cuenta_campos + 1
        Load lb_control(cuenta_campos)
        Set lb_control(cuenta_campos).cl_campo = coll_campos.item(cuenta_campos)
        Select Case lb_control(cuenta_campos).cl_campo.tipo_campo
            Case "CTEXT"
                lb_control(cuenta_campos).BackColor = FONDO_COLOR
                lb_control(cuenta_campos).Tag = genidcampo("CTEXT", cuenta_campos)
            Case "MTEXT"
                lb_control(cuenta_campos).BackColor = FONDO_COLOR
                lb_control(cuenta_campos).Tag = genidcampo("MTEXT", cuenta_campos)
            Case "NUMERICO"
                lb_control(cuenta_campos).BackColor = FONDO_COLOR
                lb_control(cuenta_campos).Tag = genidcampo("NUMERICO", cuenta_campos)
            Case "MTDIGITAL"
                lb_control(cuenta_campos).BackColor = FONDO_COLOR
                lb_control(cuenta_campos).Tag = genidcampo("MTDIGITAL", cuenta_campos)
            Case "ALFANUM"
                lb_control(cuenta_campos).BackColor = FONDO_COLOR
                lb_control(cuenta_campos).Tag = genidcampo("ALFANUM", cuenta_campos)
        End Select
        lb_control(cuenta_campos).Caption = lb_control(cuenta_campos).cl_campo.texto
        lb_control(cuenta_campos).Width = lb_control(cuenta_campos).cl_campo.largo
        lb_control(cuenta_campos).Top = (lb_control(cuenta_campos).cl_campo.y_pos - 1)
        lb_control(cuenta_campos).Left = (lb_control(cuenta_campos).cl_campo.x_pos - 1)
        lb_control(cuenta_campos).Visible = True
    Wend
    
End Function


Private Sub CH_SIZE_Click()
    If CH_SIZE.Caption = "Teclado" Then
        Me.Height = 6900
        CH_SIZE.Caption = "Copacto"
    Else
        Me.Height = 3420
        CH_SIZE.Caption = "Teclado"
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
'********************************************************************************************************************
'******************************************* DIBUJA EL DIPLAY *******************************************************
'********************************************************************************************************************
    Pic_display.Width = MAX_COL * PASOH
    Pic_display.Height = MAX_FILA * PASOV
    Pic_display.ScaleMode = 0
    Pic_display.ScaleHeight = MAX_FILA
    Pic_display.ScaleWidth = MAX_COL
    
    Pic_display.BackColor = FONDO_COLOR
    
    For i = 0 To MAX_FILA
        Load LN_fila(i)
        LN_fila(i).BorderColor = GRIS
        LN_fila(i).y1 = i
        LN_fila(i).Y2 = i
        LN_fila(i).x1 = 0
        LN_fila(i).x2 = MAX_COL
        LN_fila(i).Visible = True
    Next i
    
    For i = 0 To MAX_COL
        Load LN_col(i)
        LN_col(i).y1 = 0
        LN_col(i).Y2 = MAX_FILA
        LN_col(i).x1 = i
        LN_col(i).x2 = i
        LN_col(i).Visible = True
        LN_col(i).BorderColor = GRIS
    Next i
    
    sh_power.BackColor = RGB(255, 0, 0)
    sh_status.BackColor = RGB(0, 0, 0)
    Timer1.Interval = 500
    Timer1.Enabled = True
    
    Pic_display.Visible = True
    
    Pic_display.SetFocus

End Sub


'****************************************************************************************************************
'** Funcion para Mover la Ventana sin Borde *********************************************************************
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    X = Int(X)
'    Y = Int(Y)
'    ReleaseCapture
'    SendMessage hwnd, _
'    WM_NCLBUTTONDOWN, _
'    HTCAPTION, 0&
'End Sub
'****************************************************************************************************************


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If properties.v_open = 1 Then properties.clear_propiedades
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then
        Me.Caption = "PANT:" + Hex(screen_ref.Numero)
        properties.clear_propiedades
        propietario.SetFocus
    Else
        Me.Caption = "Panel de Operador PO-900"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim j As Integer
    Set old_foco_campo = Nothing
    If coll_campos.count = 0 Then
        Select Case screen_ref.modo
            Case SC_LIBRE
                screen_ref.modo = SC_LIBRE
            Case SC_NUEVO
                screen_ref.modo = SC_LIBRE
            Case SC_USADO
                screen_ref.modo = SC_BORRAR
            Case SC_MODIFICADO
                screen_ref.modo = SC_BORRAR
            Case SC_BORRAR
                screen_ref.modo = SC_BORRAR
        End Select
    End If
    ordena_campos
    cuenta_campos = 0
    propietario.close_view_pantalla (screen_ref.Numero)
    
End Sub

'****************************************************************************************************************
'**************************** DONDE SE CARGAN LAS PROPIEDADES DEL CAMPO *****************************************
'****************************************************************************************************************
Public Sub lb_control_Click(Index As Integer)
        Dim temp As String
        temp = properties.cargar(Index, Me, propietario)

End Sub

Private Sub lb_control_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    Dim tempX As Integer
    Dim tempY As Integer
    
    tempX = Int(X / PASOV)
    tempY = Int(Y / PASOH)
    
    tempX = tempX + lb_control(Index).cl_campo.x_pos
    tempY = tempY + lb_control(Index).cl_campo.y_pos
    
    Label2.Caption = "X: " & CStr(tempX)
    Label3.Caption = "Y: " & CStr(tempY)

End Sub

Private Sub lb_control_GotFocus(Index As Integer)
        If old_foco_campo Is Nothing Then
        Else
            old_foco_campo.BorderStyle = 0
            old_foco_campo.BackColor = FONDO_COLOR
        End If
        Set old_foco_campo = lb_control(Index)
        lb_control(Index).BorderStyle = 1
        Select Case lb_control(Index).cl_campo.tipo_campo
            Case "CTEXT"
                lb_control(Index).BackColor = CTEXT_COLOR
            Case "MTEXT"
                lb_control(Index).BackColor = MTEXT_COLOR
            Case "NUMERICO"
                lb_control(Index).BackColor = NUMERICO_COLOR
            Case "MTDIGITAL"
                lb_control(Index).BackColor = MTDIGITAL_COLOR
            Case "ALFANUM"
                lb_control(Index).BackColor = ALFANUM_COLOR
        End Select

End Sub

Private Sub lb_control_LostFocus(Index As Integer)
'        lb_control(Index).BorderStyle = 0
'        lb_control(Index).BackColor = FONDO_COLOR
End Sub

Private Sub lb_control_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Dim x1 As Integer
    Dim x2 As Integer
    
    X = Int(X)
    Y = Int(Y)
    
    x2 = Int(X / PASOH)
    
    If lb_control(Index).Tag <> Source.Tag Then
        MsgBox ("no se puede poner un control sobre otro control")
    Else
        If x2 <> x_aux Then
            If x2 > x_aux Then
                x1 = Source.Left + (x2 - x_aux)
            Else
                x1 = Source.Left - (x_aux - x2)
            End If
            Source.cl_campo.x_pos = x1 + 1
            Source.Left = (Source.cl_campo.x_pos - 1)
            lb_control_Click (Index)
            If (screen_ref.modo = SC_USADO) Or (screen_ref.modo = SC_BORRAR) Then
                screen_ref.modo = SC_MODIFICADO
            End If
        End If
    End If
    Source.Drag vbEndDrag
End Sub

'********************************************************************************************************************
'********************************************* BORRA LOS CAMPODS ****************************************************
'********************************************************************************************************************
Private Sub lb_control_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim si As Integer
    Dim i As Byte
    If KeyCode = 46 Then
        si = MsgBox("¿Estas seguro de borrar?", vbOKCancel)
            If si = vbOK Then
                Set old_foco_campo = Nothing
                coll_campos.Remove (lb_control(Index).Tag)
                Unload lb_control(Index)
                
                If coll_campos.count = 0 Then
                    cuenta_campos = 0
                End If
                
                properties.clear_propiedades
               
                If coll_campos.count = 0 Then
                    If screen_ref.modo = SC_NUEVO Then
                        screen_ref.modo = SC_LIBRE
                    Else
                        screen_ref.modo = SC_BORRAR
                    End If
                Else
                    If screen_ref.modo = SC_USADO Then
                        screen_ref.modo = SC_MODIFICADO
                    End If
                End If
            End If
    End If
    If (Shift = 2 And KeyCode = 67) Or (Shift = 2 And KeyCode = 45) Then
        Set temp_cont_cop = lb_control(Index).cl_campo.Clone
    End If
End Sub


'********************************************************************************************************************
'**************************************** FUNCION COPIAR ************************************************************
'********************************************************************************************************************
Private Sub lb_control_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim aux As String
    Dim aux2 As String
    
    X = Int(X)
    Y = Int(Y)
    
    lb_control(Index).BorderStyle = 1
    
    lb_control_Click (Index)
    
    If InStr(lb_control(Index).Tag, "Copiar:") Then
        lb_control(Index).Tag = Mid(lb_control(Index).Tag, Len("Copiar:") + 1)
    End If
    
    If Button = 1 Then
        lb_control(Index).Drag
        
        If Shift = 2 Then
            lb_control(Index).Tag = "Copiar:" + lb_control(Index).Tag
        End If
    End If

End Sub

Private Sub lb_control_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tempX As Integer
    Dim tempY As Integer
    
    y_aux = Int(Y / PASOV)
    x_aux = Int(X / PASOH)
    
    tempX = x_aux + lb_control(Index).cl_campo.x_pos
    tempY = y_aux + lb_control(Index).cl_campo.y_pos
    
    Label2.Caption = "X: " & CStr(tempX)
    Label3.Caption = "Y: " & CStr(tempY)

End Sub

Public Sub LB_END_DblClick()
    Unload Me
End Sub

Private Sub Lb_teclas_DblClick()
    PropRun_DblClick
End Sub

Private Sub Pic_display_Click()
    properties.clear_propiedades
    If old_foco_campo Is Nothing Then
    Else
       old_foco_campo.BorderStyle = 0
       old_foco_campo.BackColor = FONDO_COLOR
    End If
    Set old_foco_campo = Nothing

End Sub

'***************************************************************************************************************
'********************************** FUNCION DROP MUY IMPOTANTE *************************************************
'***************************************************************************************************************
Public Sub Pic_display_DragDrop(Source As Control, X As Single, Y As Single)
    Dim x1 As Integer
    Dim y1 As Integer
    Dim aux As String
    Dim aux2 As String
    Dim aux_campo As Variant
    Dim temp As String
    
    X = Int(X)
    Y = Int(Y)
    

    If Source.Tag = "Nuevo" Or InStr(Source.Tag, "Copiar:") Then 'Define si es un campo Arrastrado o Nuevo
        If coll_campos.count <= 15 Then
            cuenta_campos = cuenta_campos + 1
            Load lb_control(cuenta_campos)
            If Source.Tag = "Nuevo" Then
                Select Case Source.Caption
                    Case "CTEXT"
                        lb_control(cuenta_campos).BackColor = CTEXT_COLOR
                        lb_control(cuenta_campos).Tag = genidcampo("CTEXT", cuenta_campos)
                        Set lb_control(cuenta_campos).cl_campo = New class_CTEXT
                    Case "MTEXT"
                        lb_control(cuenta_campos).BackColor = MTEXT_COLOR
                        lb_control(cuenta_campos).Tag = genidcampo("MTEXT", cuenta_campos)
                        Set lb_control(cuenta_campos).cl_campo = New Class_MTEXT
                    Case "NUMERICO"
                        lb_control(cuenta_campos).BackColor = NUMERICO_COLOR
                        lb_control(cuenta_campos).Tag = genidcampo("NUMERICO", cuenta_campos)
                        Set lb_control(cuenta_campos).cl_campo = New Class_NUM
                    Case "MTDIGITAL"
                        lb_control(cuenta_campos).BackColor = MTDIGITAL_COLOR
                        lb_control(cuenta_campos).Tag = genidcampo("MTDIGITAL", cuenta_campos)
                        Set lb_control(cuenta_campos).cl_campo = New Class_MTD
                    Case "ALFANUM"
                        lb_control(cuenta_campos).BackColor = ALFANUM_COLOR
                        lb_control(cuenta_campos).Tag = genidcampo("ALFANUM", cuenta_campos)
                        Set lb_control(cuenta_campos).cl_campo = New Class_ALFA
                End Select
            Else
                Select Case Source.cl_campo.tipo_campo
                    Case "CTEXT"
                        lb_control(cuenta_campos).BackColor = CTEXT_COLOR
                        lb_control(cuenta_campos).Tag = genidcampo("CTEXT", cuenta_campos)
                    Case "MTEXT"
                        lb_control(cuenta_campos).BackColor = MTEXT_COLOR
                        lb_control(cuenta_campos).Tag = genidcampo("MTEXT", cuenta_campos)
                    Case "NUMERICO"
                        lb_control(cuenta_campos).BackColor = NUMERICO_COLOR
                        lb_control(cuenta_campos).Tag = genidcampo("NUMERICO", cuenta_campos)
                    Case "MTDIGITAL"
                        lb_control(cuenta_campos).BackColor = MTDIGITAL_COLOR
                        lb_control(cuenta_campos).Tag = genidcampo("MTDIGITAL", cuenta_campos)
                    Case "ALFANUM"
                        lb_control(cuenta_campos).BackColor = ALFANUM_COLOR
                        lb_control(cuenta_campos).Tag = genidcampo("ALFANUM", cuenta_campos)
                End Select
                Set lb_control(cuenta_campos).cl_campo = Source.cl_campo.Clone
                Source.Tag = Mid(Source.Tag, Len("Copiar:") + 1)
            End If
            coll_campos.Add lb_control(cuenta_campos).cl_campo, lb_control(cuenta_campos).Tag
            lb_control(cuenta_campos).cl_campo.x_pos = X + 1 - x_aux
            lb_control(cuenta_campos).cl_campo.y_pos = Y + 1 - y_aux
            lb_control(cuenta_campos).Caption = lb_control(cuenta_campos).cl_campo.texto
            lb_control(cuenta_campos).Width = lb_control(cuenta_campos).cl_campo.largo
            lb_control(cuenta_campos).Top = (lb_control(cuenta_campos).cl_campo.y_pos - 1)
            lb_control(cuenta_campos).Left = (lb_control(cuenta_campos).cl_campo.x_pos - 1)
            lb_control(cuenta_campos).Visible = True
            lb_control(cuenta_campos).SetFocus
            If (screen_ref.modo = SC_LIBRE) Then
            screen_ref.modo = SC_NUEVO
            End If
            temp = properties.cargar(cuenta_campos, Me, propietario)
        Else
            MsgBox ("No se puebden poner mas de 16 campos")
        End If
    Else
        x1 = X - x_aux
        y1 = Y - y_aux
        Source.cl_campo.x_pos = x1 + 1
        Source.cl_campo.y_pos = y1 + 1
        Source.Top = Source.cl_campo.y_pos - 1
        Source.Left = Source.cl_campo.x_pos - 1
        Source.Width = Source.cl_campo.largo
    End If
    If (screen_ref.modo = SC_USADO) Or (screen_ref.modo = SC_BORRAR) Then
        screen_ref.modo = SC_MODIFICADO
    End If
End Sub

Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If Source.Tag <> "Nuevo" And InStr(Source.Tag, "Copiar:") = 0 Then
        Source.Drag (vbEndDrag)
    End If
End Sub

Private Sub Pic_display_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Label2.Caption = "X: " & CStr(Int(X) + 1)
    Label3.Caption = "Y: " & CStr(Int(Y) + 1)
End Sub

Private Sub Pic_display_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift = 2 And KeyCode = 86) Or (Shift = 1 And KeyCode = 45) Then
        If coll_campos.count <= 15 Then
            If Not temp_cont_cop Is Nothing Then
                cuenta_campos = cuenta_campos + 1
                Load lb_control(cuenta_campos)
                Select Case temp_cont_cop.tipo_campo
                    Case "CTEXT"
                        lb_control(cuenta_campos).BackColor = CTEXT_COLOR
                        lb_control(cuenta_campos).Tag = genidcampo("CTEXT", cuenta_campos)
                    Case "MTEXT"
                        lb_control(cuenta_campos).BackColor = MTEXT_COLOR
                        lb_control(cuenta_campos).Tag = genidcampo("MTEXT", cuenta_campos)
                    Case "NUMERICO"
                        lb_control(cuenta_campos).BackColor = NUMERICO_COLOR
                        lb_control(cuenta_campos).Tag = genidcampo("NUMERICO", cuenta_campos)
                    Case "MTDIGITAL"
                        lb_control(cuenta_campos).BackColor = MTDIGITAL_COLOR
                        lb_control(cuenta_campos).Tag = genidcampo("MTDIGITAL", cuenta_campos)
                    Case "ALFANUM"
                        lb_control(cuenta_campos).BackColor = ALFANUM_COLOR
                        lb_control(cuenta_campos).Tag = genidcampo("ALFANUM", cuenta_campos)
                End Select
                Set lb_control(cuenta_campos).cl_campo = temp_cont_cop.Clone
                coll_campos.Add lb_control(cuenta_campos).cl_campo, lb_control(cuenta_campos).Tag
                lb_control(cuenta_campos).cl_campo.x_pos = temp_cont_cop.x_pos
                lb_control(cuenta_campos).cl_campo.y_pos = temp_cont_cop.y_pos
                lb_control(cuenta_campos).Caption = lb_control(cuenta_campos).cl_campo.texto
                lb_control(cuenta_campos).Width = lb_control(cuenta_campos).cl_campo.largo
                lb_control(cuenta_campos).Top = (lb_control(cuenta_campos).cl_campo.y_pos - 1)
                lb_control(cuenta_campos).Left = (lb_control(cuenta_campos).cl_campo.x_pos - 1)
                lb_control(cuenta_campos).Visible = True
                lb_control(cuenta_campos).SetFocus
                If (screen_ref.modo = SC_LIBRE) Then
                    screen_ref.modo = SC_NUEVO
                End If
                Dim temp As Variant 'para la funcion
                temp = properties.cargar(cuenta_campos, Me, propietario)
            End If
        Else
            MsgBox ("No se pueden poner mas de 16 campos")
        End If
    End If
End Sub

Private Sub Pic_display_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.Caption = "X: " & CStr(Int(X) + 1)
    Label3.Caption = "Y: " & CStr(Int(Y) + 1)
End Sub

Private Sub PropRun_DblClick()
    Dim ret As Boolean
    ret = V_Prop_P_G.viewlocals(propietario.archivo, screen_ref.Numero)
    If ret = True Then
        Lb_teclas.Caption = Hex(screen_ref.Numero) & " : " & screen_ref.name
    End If
End Sub

Private Sub Timer1_Timer()
    If sh_status.BackColor = RGB(0, 0, 0) Then
        sh_status.BackColor = RGB(0, 255, 0)
    Else
        sh_status.BackColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub ordena_campos()
    Dim temp_coll As Collection
    Dim aux As String
    Dim X As Variant
    Dim j As Integer
    Set temp_coll = New Collection
    For j = 1 To coll_campos.count
        Set X = coll_campos(j)
        aux = genidcampo(X.tipo_campo, j)
        temp_coll.Add X, aux
    Next j
    Set screen_ref.colectcampo = temp_coll
    Set temp_coll = Nothing
End Sub

Private Sub ViewCampos_Click()
    V_Tools.Show
End Sub

Public Sub PropertyMOD()
    If (screen_ref.modo = SC_USADO) Or (screen_ref.modo = SC_BORRAR) Then
        screen_ref.modo = SC_MODIFICADO
    End If
End Sub
