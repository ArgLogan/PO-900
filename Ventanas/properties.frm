VERSION 5.00
Begin VB.Form properties 
   BackColor       =   &H80000009&
   Caption         =   "propiedades"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   2940
   Begin VB.VScrollBar VScroll 
      Height          =   5895
      Left            =   2680
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic_cont 
      BackColor       =   &H80000009&
      Height          =   13695
      Left            =   0
      ScaleHeight     =   13635
      ScaleWidth      =   2940
      TabIndex        =   0
      Top             =   0
      Width           =   3000
      Begin EDITPO900.campo_prop lista_prop 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   13440
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private indice_control As Integer
Dim ptr_control As lb_cont
Dim ptr_proyect As V_Indice
Dim ventana_actual As pantalla
Dim foco As Integer
Dim presskey As Boolean
Public v_open As Byte
Dim mvarContainerTop As Long

Private Sub Form_Resize()
    cambia_tamaño 'Función de cambio de tamaño del formulario
End Sub

Private Sub Form_Unload(Cancel As Integer)
    v_open = 0
End Sub

Private Sub lista_prop_Change(Index As Integer)
    flag_cambio = cambio
End Sub

Private Sub lista_prop_Click(Index As Integer)
    Dim aux As String
    Dim aux2 As String
    lista_prop(Index).SetFocus
'************************************* SOLO PARA CAMPO MULTITEXT DIGITAL *************************************
    aux = Mid(lista_prop(Index).Caption, 1, 5)
    aux2 = ventana_actual.lb_control(indice_control).cl_campo.tipo_campo
    If aux2 = "MTDIGITAL" And aux = "Texto" Then
        aux = Mid(lista_prop(Index).Caption, 6, 2)
        ventana_actual.lb_control(indice_control).Caption = ventana_actual.lb_control(indice_control).cl_campo.texto(Val(aux) - 1)
        ventana_actual.lb_control(indice_control).Width = (ventana_actual.lb_control(indice_control).cl_campo.largo)
    End If
End Sub

Private Sub lista_prop_DblClick(Index As Integer)
    Dim new_lra As String
    flag_cambio = sin_cambio
    If lista_prop(Index).Caption = "LRA" Or lista_prop(Index).Caption = "Trg.LRA" Then
        If lista_prop(Index).Caption = "LRA" Then
            new_lra = V_LRA.new_lra(lista_prop(Index).Text, lista_prop(Index).Caption, ptr_proyect.lra_limite_sup, ptr_proyect.lra_limite_inf)
        Else
            new_lra = V_LRA.new_lra(lista_prop(Index).Text, lista_prop(Index).Caption, ptr_proyect.bit_limite_sup, ptr_proyect.bit_limite_inf)
        End If
        lista_prop(Index).Text = new_lra
    End If
    lista_prop(Index).SetFocus
End Sub

Private Sub lista_prop_DownClick(Index As Integer)
    If lista_prop(Index).Caption = "X" Or lista_prop(Index).Caption = "Y" Or lista_prop(Index).Caption = "Items" Then
        lista_prop(Index).Text = CStr(Val(lista_prop(Index).Text) - 1)
        If Val(lista_prop(Index).Text) < 1 Then lista_prop(Index).Text = "1"
    Else
        Select Case lista_prop(Index).Text
            Case "Falso"
                lista_prop(Index).Text = "True"
            Case "Verdadero"
                lista_prop(Index).Text = "False"
            Case "Blink"
                lista_prop(Index).Text = "Normal"
            Case "Normal"
                lista_prop(Index).Text = "Blink"
            Case "WS1"
                lista_prop(Index).Text = "W1"
            Case "W1"
                lista_prop(Index).Text = "FPS"
            Case "FPS"
                lista_prop(Index).Text = "BIN"
            Case "BIN"
                lista_prop(Index).Text = "HEX"
            Case "HEX"
                lista_prop(Index).Text = "WS1"
            Case "False"
                lista_prop(Index).Text = "True"
            Case "True"
                lista_prop(Index).Text = "False"
        End Select
    End If
End Sub

'*******************************************************************************************************************
'************************************** Herramientas de edicion de campo  ******************************************
'*******************************************************************************************************************

Private Sub lista_prop_GotFocus(Index As Integer)
    lista_prop(Index).BorderStyle = 1
    
    If lista_prop(Index).Caption = "X" Then lista_prop(Index).ud_visible = True
    
    If lista_prop(Index).Caption = "Y" Then lista_prop(Index).ud_visible = True
    
    If lista_prop(Index).Caption = "Gain exp" Then lista_prop(Index).ud_visible = True
    
    If lista_prop(Index).Caption = "Editable" Then
        lista_prop(Index).ud_visible = True
        lista_prop(Index).Locked = True
    End If
    
    If lista_prop(Index).Caption = "Trg.Enable" Then
        lista_prop(Index).ud_visible = True
        lista_prop(Index).Locked = True
    End If
    
    If lista_prop(Index).Caption = "Items" Then
        lista_prop(Index).ud_visible = True
        lista_prop(Index).Locked = True
    End If
    If lista_prop(Index).Caption = "atributos" Then
        lista_prop(Index).ud_visible = True
        lista_prop(Index).Locked = True
    End If
    If lista_prop(Index).Caption = "Modo" Then
        lista_prop(Index).ud_visible = True
        lista_prop(Index).Locked = True
    End If
    If lista_prop(Index).Caption = "LRA" Then
        lista_prop(Index).Locked = True
    End If
    
    If lista_prop(Index).Caption = "Trg.LRA" Then
        lista_prop(Index).Locked = True
    End If
    foco = Index
End Sub

'*******************************************************************************************************************
'*********************************** desplazamiento en la ventana de propiedades  **********************************
'*******************************************************************************************************************
Private Sub lista_prop_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim j As Byte
    If (Index > 0) And (Index < (NPROP + 1)) Then
        If (KeyCode = 13) Then
            If Index <> max_foco Then
                lista_prop((Index + 1)).SetFocus
                foco = Index + 1
             Else
                lista_prop(1).SetFocus
                foco = 1
            End If
            presskey = True
        End If
        
        If KeyCode = 40 Then
            If Index <> max_foco Then
                lista_prop((Index + 1)).SetFocus
                foco = Index + 1
             Else
                lista_prop(1).SetFocus
                foco = 1
             End If
            presskey = True
        End If
        
        If KeyCode = 38 Then
            If Index <> 1 Then
                lista_prop((Index - 1)).SetFocus
                foco = Index - 1
            Else
                lista_prop(max_foco).SetFocus
                foco = max_foco
            End If
            presskey = True
        End If
    End If
End Sub
'**************************************************************************************************************
'********************************** ACTUALIZA LAS PROPIEDADES EN LA COLLECTION ********************************
'**************************************************************************************************************
Private Sub lista_prop_LostFocus(Index As Integer)
    Dim aux As String
    Dim aux2 As String
    lista_prop(Index).BorderStyle = 0
    If lista_prop(Index).Numerico = True Then
        If Len(lista_prop(Index).Text) = 1 And (lista_prop(Index).Text = "-" Or lista_prop(Index).Text = ",") Then
                lista_prop(Index).Text = "0"
        End If
    End If
    
    If flag_cambio = cambio Then
        If lista_prop(Index).Text <> "" Or InStr(1, lista_prop(Index).Caption, "Texto") Or lista_prop(Index).Caption = "Nombre" Then
            Select Case lista_prop(Index).Caption
                Case "X"
                    ventana_actual.lb_control(indice_control).cl_campo.x_pos = CInt(lista_prop(Index).Text)
                    ventana_actual.lb_control(indice_control).Left = ventana_actual.lb_control(indice_control).cl_campo.x_pos - 1
                
                Case "Y"
                    ventana_actual.lb_control(indice_control).cl_campo.y_pos = CInt(lista_prop(Index).Text)
                    ventana_actual.lb_control(indice_control).Top = ventana_actual.lb_control(indice_control).cl_campo.y_pos - 1
                
                Case "atributos"
                        If lista_prop(Index).Text = "Blink" Then
                            aux2 = "01"
                        Else
                            aux2 = "00"
                        End If
                        ventana_actual.lb_control(indice_control).cl_campo.atributo = aux2
                
                Case "Texto"
                    ventana_actual.lb_control(indice_control).cl_campo.texto = lista_prop(Index).Text
                    ventana_actual.lb_control(indice_control).Caption = ventana_actual.lb_control(indice_control).cl_campo.texto
                    ventana_actual.lb_control(indice_control).Width = Len(ventana_actual.lb_control(indice_control).cl_campo.texto)
                
                Case "JMP"
                    ventana_actual.lb_control(indice_control).cl_campo.jmp = Val("&H" + lista_prop(Index).Text)
                
                Case "Decimales"
                    ventana_actual.lb_control(indice_control).cl_campo.dec = lista_prop(Index).Text
                    ventana_actual.lb_control(indice_control).Caption = ventana_actual.lb_control(indice_control).cl_campo.texto
                    ventana_actual.lb_control(indice_control).Width = Len(ventana_actual.lb_control(indice_control).cl_campo.texto)
                    ventana_actual.lb_control(indice_control).Left = ventana_actual.lb_control(indice_control).cl_campo.x_pos - 1
                    
                Case "Enteros"
                    ventana_actual.lb_control(indice_control).cl_campo.ent = CDbl(lista_prop(Index).Text)
                    ventana_actual.lb_control(indice_control).Caption = ventana_actual.lb_control(indice_control).cl_campo.texto
                    ventana_actual.lb_control(indice_control).Width = Len(ventana_actual.lb_control(indice_control).cl_campo.texto)
                    ventana_actual.lb_control(indice_control).Left = ventana_actual.lb_control(indice_control).cl_campo.x_pos - 1
                
                Case "Modo"
                    ventana_actual.lb_control(indice_control).cl_campo.modo = lista_prop(Index).Text
                    ventana_actual.lb_control(indice_control).Caption = ventana_actual.lb_control(indice_control).cl_campo.texto
                Case "Rangox0"
                    ventana_actual.lb_control(indice_control).cl_campo.rangoX(0) = Val(lista_prop(Index).Text)
                Case "Rangox1"
                    ventana_actual.lb_control(indice_control).cl_campo.rangoX(1) = Val(lista_prop(Index).Text)

                Case "Gain exp"
                    ventana_actual.lb_control(indice_control).cl_campo.recta = CBool(lista_prop(Index).Text)
                    
                Case "Gain"
                    ventana_actual.lb_control(indice_control).cl_campo.gain = lista_prop(Index).Text
                    
                Case "Offset"
                    ventana_actual.lb_control(indice_control).cl_campo.offset = lista_prop(Index).Text
                    
                Case "Max"
                    ventana_actual.lb_control(indice_control).cl_campo.maximo = CDbl(lista_prop(Index).Text)
                    
                Case "Min"
                    ventana_actual.lb_control(indice_control).cl_campo.minimo = CDbl(lista_prop(Index).Text)
                
                Case "LRA"
                    ventana_actual.lb_control(indice_control).cl_campo.lra = lista_prop(Index).Text
                
                Case "Editable"
                    ventana_actual.lb_control(indice_control).cl_campo.Edit = lista_prop(Index).Text
                    ventana_actual.lb_control(indice_control).Left = (ventana_actual.lb_control(indice_control).cl_campo.x_pos) - 1
                    ventana_actual.lb_control(indice_control).Width = ventana_actual.lb_control(indice_control).cl_campo.largo
                Case "Rangoy0"
                    ventana_actual.lb_control(indice_control).cl_campo.rangoY(0) = Val(lista_prop(Index).Text)
                Case "Rangoy1"
                    ventana_actual.lb_control(indice_control).cl_campo.rangoY(1) = Val(lista_prop(Index).Text)
                Case "Len"
                    ventana_actual.lb_control(indice_control).cl_campo.largo = Val(lista_prop(Index).Text)
                    If ventana_actual.lb_control(indice_control).cl_campo.tipo_campo <> "MTDIGITAL" Then
                        ventana_actual.lb_control(indice_control).Caption = ventana_actual.lb_control(indice_control).cl_campo.texto
                    Else
                        ventana_actual.lb_control(indice_control).Caption = ventana_actual.lb_control(indice_control).cl_campo.texto(1)
                    End If
                    ventana_actual.lb_control(indice_control).Width = ventana_actual.lb_control(indice_control).cl_campo.largo
                    
                Case "Items"
                    ventana_actual.lb_control(indice_control).cl_campo.items = Val(lista_prop(Index).Text)
                    ventana_actual.lb_control(indice_control).Caption = ventana_actual.lb_control(indice_control).cl_campo.texto
                
                Case "Texto1"
                    ventana_actual.lb_control(indice_control).cl_campo.texto(0) = lista_prop(Index).Text
                Case "Texto2"
                    ventana_actual.lb_control(indice_control).cl_campo.texto(1) = lista_prop(Index).Text
                Case "Texto3"
                    ventana_actual.lb_control(indice_control).cl_campo.texto(2) = lista_prop(Index).Text
                Case "Texto4"
                    ventana_actual.lb_control(indice_control).cl_campo.texto(3) = lista_prop(Index).Text
                Case "Texto5"
                    ventana_actual.lb_control(indice_control).cl_campo.texto(4) = lista_prop(Index).Text
                Case "Texto6"
                    ventana_actual.lb_control(indice_control).cl_campo.texto(5) = lista_prop(Index).Text
                Case "Texto7"
                    ventana_actual.lb_control(indice_control).cl_campo.texto(6) = lista_prop(Index).Text
                Case "Texto8"
                    ventana_actual.lb_control(indice_control).cl_campo.texto(7) = lista_prop(Index).Text
                Case "Texto9"
                    ventana_actual.lb_control(indice_control).cl_campo.texto(8) = lista_prop(Index).Text
                Case "Texto10"
                    ventana_actual.lb_control(indice_control).cl_campo.texto(9) = lista_prop(Index).Text
                Case "Texto11"
                    ventana_actual.lb_control(indice_control).cl_campo.texto(10) = lista_prop(Index).Text
                Case "Texto12"
                    ventana_actual.lb_control(indice_control).cl_campo.texto(11) = lista_prop(Index).Text
                Case "Texto13"
                    ventana_actual.lb_control(indice_control).cl_campo.texto(12) = lista_prop(Index).Text
                Case "Texto14"
                    ventana_actual.lb_control(indice_control).cl_campo.texto(13) = lista_prop(Index).Text
                Case "Texto15"
                    ventana_actual.lb_control(indice_control).cl_campo.texto(14) = lista_prop(Index).Text
                Case "Texto16"
                    ventana_actual.lb_control(indice_control).cl_campo.texto(15) = lista_prop(Index).Text
                Case "Nombre"
                    ventana_actual.lb_control(indice_control).cl_campo.name = lista_prop(Index).Text
                Case "Trg.LRA"
                    ventana_actual.lb_control(indice_control).cl_campo.TRIGGER = lista_prop(Index).Text
                Case "Trg.Enable"
                    ventana_actual.lb_control(indice_control).cl_campo.TRIGGER_ENABLE = CBool(lista_prop(Index).Text)
            End Select
        End If
        Dim temp
        flag_cambio = sin_cambio
        temp = cargar(indice_control, ventana_actual, ptr_proyect)
        If presskey Then
            lista_prop(foco).SetFocus
        End If
        presskey = False
        ventana_actual.PropertyMOD
    End If
    lista_prop(Index).ud_visible = False
End Sub

Private Sub lista_prop_UpClick(Index As Integer)
    If lista_prop(Index).Caption = "X" Or lista_prop(Index).Caption = "Y" Or lista_prop(Index).Caption = "Items" Then
        lista_prop(Index).Text = CStr(Val(lista_prop(Index).Text) + 1)
        If lista_prop(Index).Caption = "X" And Val(lista_prop(Index).Text) > 20 Then lista_prop(Index).Text = "20"
        If lista_prop(Index).Caption = "Y" And Val(lista_prop(Index).Text) > 4 Then lista_prop(Index).Text = "4"
        If lista_prop(Index).Caption = "MT_Items" And Val(lista_prop(Index).Text) > 16 Then lista_prop(Index).Text = "16"
    Else
        Select Case lista_prop(Index).Text
            Case "Falso"
                lista_prop(Index).Text = "True"
            Case "Verdadero"
                lista_prop(Index).Text = "False"
            Case "Blink"
                lista_prop(Index).Text = "Normal"
            Case "Normal"
                lista_prop(Index).Text = "Blink"
            Case "BIN"
                lista_prop(Index).Text = "FPS"
            Case "HEX"
                lista_prop(Index).Text = "BIN"
            Case "WS1"
                lista_prop(Index).Text = "HEX"
            Case "W1"
                lista_prop(Index).Text = "WS1"
            Case "FPS"
                lista_prop(Index).Text = "W1"
            Case "False"
                lista_prop(Index).Text = "True"
            Case "True"
                lista_prop(Index).Text = "False"
        End Select
    End If
End Sub

Private Sub Form_Load()
    cambia_tamaño
    mvarContainerTop = pic_cont.Top
    Me.Left = EditorIDE.ScaleWidth - Me.Width
    Me.Top = 0
    v_open = 1
End Sub
Public Function cargar(ByVal Index As Integer, VENTANA As pantalla, idprj As V_Indice) As Boolean
    Dim i As Byte
    Dim aux As String
    
    Set ptr_proyect = idprj
    
    If flag_cambio = cambio Then
         lista_prop_LostFocus (foco)
    End If
    
     Set ventana_actual = VENTANA
     Set ptr_control = VENTANA.lb_control(Index)
    
     If lista_prop(0).Tag = "" Then
         For i = 1 To NPROP
             Load lista_prop(i)
             lista_prop(i).BorderStyle = 0
             lista_prop(i).Top = ((i - 1) * 240) + (15 * (i - 1))
         Next i
         lista_prop(0).Tag = "m"
     End If
     
     clear_propiedades
     
     flag_cambio = sin_cambio
     properties.Caption = ventana_actual.lb_control(Index).Tag
     
     aux = ptr_control.cl_campo.tipo_campo
     lista_prop(1).Visible = True
     lista_prop(1).Caption = "Nombre"
     lista_prop(1).Text = ptr_control.cl_campo.name
     
     lista_prop(2).Visible = True
     lista_prop(2).Caption = "atributos"
     If ptr_control.cl_campo.atributo = "00" Then
         lista_prop(2).Text = "Normal"
     Else
         lista_prop(2).Text = "Blink"
     End If
     lista_prop(3).Visible = True
     lista_prop(3).Caption = "X"
     lista_prop(3).Text = ptr_control.cl_campo.x_pos
     lista_prop(3).Numerico = True
     
     lista_prop(4).Visible = True
     lista_prop(4).Caption = "Y"
     lista_prop(4).Text = ptr_control.cl_campo.y_pos
     lista_prop(4).Numerico = True
     
     Select Case ptr_control.cl_campo.tipo_campo
         Case "CTEXT"
             lista_prop(5).Visible = True
             lista_prop(5).Caption = "Texto"
             lista_prop(5).Text = ptr_control.cl_campo.texto
             max_foco = 5
         Case "MTEXT"
             lista_prop(5).Visible = True
             lista_prop(5).Caption = "Texto"
             lista_prop(5).Text = ptr_control.cl_campo.texto
             
             lista_prop(6).Visible = True
             lista_prop(6).Caption = "JMP"
             lista_prop(6).Text = Hex(ptr_control.cl_campo.jmp)
             max_foco = 6
         Case "NUMERICO"
             lista_prop(5).Visible = True
             lista_prop(5).Caption = "LRA"
             lista_prop(5).Text = ptr_control.cl_campo.lra
             
             lista_prop(6).Visible = True
             lista_prop(6).Caption = "Modo"
             lista_prop(6).Text = ptr_control.cl_campo.modo
             
             lista_prop(7).Visible = True
             lista_prop(7).Caption = "Enteros"
             lista_prop(7).Text = ptr_control.cl_campo.ent
             lista_prop(7).Numerico = True
             
             lista_prop(8).Visible = True
             lista_prop(8).Caption = "Decimales"
             lista_prop(8).Text = ptr_control.cl_campo.dec
             lista_prop(8).Numerico = True
             lista_prop(8).Limite_inf = 0
             lista_prop(8).Limite_Sup = 4
             
             lista_prop(9).Visible = True
             lista_prop(9).Caption = "Max"
             lista_prop(9).Text = ptr_control.cl_campo.maximo
             lista_prop(9).Numerico = True
             lista_prop(9).Limite_Sup = 999999
             lista_prop(9).Limite_inf = -999999
             
             lista_prop(10).Visible = True
             lista_prop(10).Caption = "Min"
             lista_prop(10).Text = ptr_control.cl_campo.minimo
             lista_prop(10).Numerico = True
             lista_prop(10).Limite_Sup = 999999
             lista_prop(10).Limite_inf = -999999
             
             lista_prop(11).Visible = True
             lista_prop(11).Caption = "Editable"
             lista_prop(11).Text = ptr_control.cl_campo.Edit
             
             lista_prop(12).Visible = True
             lista_prop(12).Caption = "Trg.Enable"
             lista_prop(12).Text = ptr_control.cl_campo.TRIGGER_ENABLE
             
             lista_prop(13).Visible = True
             lista_prop(13).Caption = "Trg.LRA"
             lista_prop(13).Text = ptr_control.cl_campo.TRIGGER
             
             lista_prop(14).Visible = True
             lista_prop(14).Caption = "Gain exp"
             lista_prop(14).Text = ptr_control.cl_campo.recta
             lista_prop(14).Numerico = True
             
             lista_prop(15).Visible = True
             lista_prop(15).Caption = "Gain"
             lista_prop(15).Text = ptr_control.cl_campo.gain
             lista_prop(15).Numerico = True
             
             lista_prop(16).Visible = True
             lista_prop(16).Caption = "Offset"
             lista_prop(16).Text = ptr_control.cl_campo.offset
             lista_prop(16).Numerico = True
             
             lista_prop(17).Visible = True
             lista_prop(17).Caption = "Rangoy0"
             lista_prop(17).Text = ptr_control.cl_campo.rangoY(0)
             lista_prop(17).Numerico = True
             
             lista_prop(18).Visible = True
             lista_prop(18).Caption = "Rangoy1"
             lista_prop(18).Text = ptr_control.cl_campo.rangoY(1)
             lista_prop(18).Numerico = True
                                    
             lista_prop(19).Visible = True
             lista_prop(19).Caption = "Rangox0"
             lista_prop(19).Text = ptr_control.cl_campo.rangoX(0)
             lista_prop(19).Numerico = True
             
             lista_prop(20).Visible = True
             lista_prop(20).Caption = "Rangox1"
             lista_prop(20).Text = ptr_control.cl_campo.rangoX(1)
             lista_prop(20).Numerico = True
             max_foco = 20
         
         Case "MTDIGITAL"
             lista_prop(5).Visible = True
             lista_prop(5).Caption = "LRA"
             lista_prop(5).Text = ptr_control.cl_campo.lra
             
             lista_prop(6).Visible = True
             lista_prop(6).Caption = "Editable"
             lista_prop(6).Text = ptr_control.cl_campo.Edit
             
             lista_prop(7).Visible = True
             lista_prop(7).Caption = "Trg.Enable"
             lista_prop(7).Text = ptr_control.cl_campo.TRIGGER_ENABLE
             
             lista_prop(8).Visible = True
             lista_prop(8).Caption = "Trg.LRA"
             lista_prop(8).Text = ptr_control.cl_campo.TRIGGER
             
             lista_prop(9).Visible = True
             lista_prop(9).Caption = "Len"
             lista_prop(9).Text = CStr(ptr_control.cl_campo.largo)
             lista_prop(9).Numerico = True
             lista_prop(9).Limite_Sup = 20
             lista_prop(9).Limite_inf = 1
             
             lista_prop(10).Visible = True
             lista_prop(10).Caption = "Items"
             lista_prop(10).Text = CStr(ptr_control.cl_campo.items)
             lista_prop(10).Numerico = True
             lista_prop(10).Limite_Sup = 16
             lista_prop(10).Limite_inf = 2
             
             lista_prop(11).Visible = True
             lista_prop(11).Caption = "Texto1"
             lista_prop(11).Text = ptr_control.cl_campo.texto(0)
             
             lista_prop(12).Visible = True
             lista_prop(12).Caption = "Texto2"
             lista_prop(12).Text = ptr_control.cl_campo.texto(1)
             max_foco = 12
             
             If ptr_control.cl_campo.items > 2 Then
                 lista_prop(13).Visible = True
                 max_foco = 13
             End If
             lista_prop(13).Caption = "Texto3"
             lista_prop(13).Text = ptr_control.cl_campo.texto(2)
             
             If ptr_control.cl_campo.items > 3 Then
                 lista_prop(14).Visible = True
                 max_foco = 14
             End If
             lista_prop(14).Caption = "Texto4"
             lista_prop(14).Text = ptr_control.cl_campo.texto(3)
             
             If ptr_control.cl_campo.items > 4 Then
                 lista_prop(15).Visible = True
                 max_foco = 15
             End If
             lista_prop(15).Caption = "Texto5"
             lista_prop(15).Text = ptr_control.cl_campo.texto(4)
             
             If ptr_control.cl_campo.items > 5 Then
                 lista_prop(16).Visible = True
                 max_foco = 16
             End If
             lista_prop(16).Caption = "Texto6"
             lista_prop(16).Text = ptr_control.cl_campo.texto(5)
             
             If ptr_control.cl_campo.items > 6 Then
                 lista_prop(17).Visible = True
                 max_foco = 17
             End If
             lista_prop(17).Caption = "Texto7"
             lista_prop(17).Text = ptr_control.cl_campo.texto(6)
             
             If ptr_control.cl_campo.items > 7 Then
                 lista_prop(18).Visible = True
                 max_foco = 18
             End If
             lista_prop(18).Caption = "Texto8"
             lista_prop(18).Text = ptr_control.cl_campo.texto(7)
             
             If ptr_control.cl_campo.items > 8 Then
                 lista_prop(19).Visible = True
                 max_foco = 19
             End If
             lista_prop(19).Caption = "Texto9"
             lista_prop(19).Text = ptr_control.cl_campo.texto(8)
             
             If ptr_control.cl_campo.items > 9 Then
                 lista_prop(20).Visible = True
                 max_foco = 20
             End If
             lista_prop(20).Caption = "Texto10"
             lista_prop(20).Text = ptr_control.cl_campo.texto(9)
             
             If ptr_control.cl_campo.items > 10 Then
                 lista_prop(21).Visible = True
                 max_foco = 21
             End If
             lista_prop(21).Caption = "Texto11"
             lista_prop(21).Text = ptr_control.cl_campo.texto(10)
             
             If ptr_control.cl_campo.items > 11 Then
                 lista_prop(22).Visible = True
                 max_foco = 22
             End If
             lista_prop(22).Caption = "Texto12"
             lista_prop(22).Text = ptr_control.cl_campo.texto(11)
             
             If ptr_control.cl_campo.items > 12 Then
                 lista_prop(23).Visible = True
                 max_foco = 23
             End If
             lista_prop(23).Caption = "Texto13"
             lista_prop(23).Text = ptr_control.cl_campo.texto(12)
             
             If ptr_control.cl_campo.items > 13 Then
                 lista_prop(24).Visible = True
                 max_foco = 24
             End If
             lista_prop(24).Caption = "Texto14"
             lista_prop(24).Text = ptr_control.cl_campo.texto(13)
             
             If ptr_control.cl_campo.items > 14 Then
                 lista_prop(25).Visible = True
                 max_foco = 25
             End If
             lista_prop(25).Caption = "Texto15"
             lista_prop(25).Text = ptr_control.cl_campo.texto(14)
             
             If ptr_control.cl_campo.items > 15 Then
                 lista_prop(26).Visible = True
                 max_foco = 26
             End If
             lista_prop(26).Caption = "Texto16"
             lista_prop(26).Text = ptr_control.cl_campo.texto(15)
             
         Case "ALFANUM"
             lista_prop(5).Visible = True
             lista_prop(5).Caption = "LRA"
             lista_prop(5).Text = ptr_control.cl_campo.lra
 
             lista_prop(6).Visible = True
             lista_prop(6).Caption = "Len"
             lista_prop(6).Text = CStr(ptr_control.cl_campo.largo)
             
             lista_prop(7).Visible = True
             lista_prop(7).Caption = "Editable"
             lista_prop(7).Text = ptr_control.cl_campo.Edit
             
             lista_prop(8).Visible = True
             lista_prop(8).Caption = "Trg.Enable"
             lista_prop(8).Text = ptr_control.cl_campo.TRIGGER_ENABLE
             
             lista_prop(9).Visible = True
             lista_prop(9).Caption = "Trg.LRA"
             lista_prop(9).Text = ptr_control.cl_campo.TRIGGER

             max_foco = 9
    End Select
    flag_cambio = sin_cambio
    Me.Height = (max_foco + 2) * 255
    indice_control = Index
End Function

Public Function clear_propiedades()
    Dim j As Byte
    Dim a As Byte
    
    If flag_cambio = cambio Then
         lista_prop_LostFocus (foco)
    End If
  
    a = properties.count
    properties.Caption = "Propiedades"
    If a > 5 Then
        For j = 1 To NPROP
            lista_prop(j).BorderStyle = 0
            lista_prop(j).Visible = False
            lista_prop(j).Locked = False
            lista_prop(j).Numerico = False
            lista_prop(j).Limite_inf = -9999
            lista_prop(j).Limite_Sup = 9999
        Next j
    End If
End Function
Private Sub VScroll_Change()
    pic_cont.Top = -VScroll.Value + mvarContainerTop
End Sub
Public Function cambia_tamaño()
    Me.Width = 3030
    If Me.Height > (EditorIDE.ScaleHeight - 200) Then Me.Height = (EditorIDE.ScaleHeight - 200)
    If Me.Height > (max_foco + 2) * 255 Then
        Me.Height = (max_foco + 2) * 255
    End If
    VScroll.Height = Me.ScaleHeight
    If (max_foco + 1) * 255 > Me.Height Then
        VScroll.Visible = True
    Else
        VScroll.Visible = False
    End If
    With VScroll
      .max = (255 * (max_foco + 1)) - (VScroll.Height)
      .SmallChange = 255
      .LargeChange = 765
    End With
End Function

