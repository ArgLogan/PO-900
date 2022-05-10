VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.UserControl campo_prop 
   BackColor       =   &H80000009&
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2685
   ScaleHeight     =   255
   ScaleWidth      =   2685
   Begin ComCtl2.UpDown ud_data 
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   450
      _Version        =   327681
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txt_data 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "PatternLCD"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Width           =   1815
   End
   Begin VB.Line linea 
      X1              =   1200
      X2              =   1200
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Line Line1 
      X1              =   1200
      X2              =   1200
      Y1              =   240
      Y2              =   0
   End
   Begin VB.Label lb_data 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Editable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "campo_prop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_Appearance = 0
Const m_def_BackStyle = 0
Const m_def_ud_visible = 0
'Property Variables:
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_Appearance As Integer
Dim m_BackStyle As Integer
Dim m_ud_visible As Boolean
Dim m_numerico As Boolean
Dim m_limite_inf As Double
Dim m_limite_sup  As Double
Dim m_flag_punto As Byte
Dim m_flag_menos As Byte
Dim m_flag_nume As Byte
'Event Declarations:
Event DblClick() 'MappingInfo=txt_data,txt_data,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txt_data,txt_data,-1,KeyPress
Attribute KeyPress.VB_Description = "Ocurre cuando el usuario presiona y libera una tecla ANSI."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txt_data,txt_data,-1,KeyDown
Attribute KeyDown.VB_Description = "Ocurre cuando el usuario presiona una tecla mientras un objeto tiene el enfoque."
Event Change() 'MappingInfo=txt_data,txt_data,-1,Change
Attribute Change.VB_Description = "Ocurre cuando cambia el contenido de un control."
Event Click() 'MappingInfo=txt_data,txt_data,-1,Click
'Event DblClick()
'Event KeyDown(KeyCode As Integer, Shift As Integer)
'Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Ocurre cuando el usuario libera una tecla mientras un objeto tiene el enfoque."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Ocurre cuando el usuario presiona el botón del mouse mientras un objeto tiene el enfoque."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Ocurre cuando el usuario mueve el mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Ocurre cuando el usuario libera el botón del mouse mientras un objeto tiene el enfoque."
Event UpClick() 'MappingInfo=ud_data,ud_data,-1,UpClick
Attribute UpClick.VB_Description = "Se ha hecho clic en el botón arriba del control UpDown"
Event DownClick() 'MappingInfo=ud_data,ud_data,-1,DownClick
Attribute DownClick.VB_Description = "Se ha hecho clic en el botón abajo del control UpDown"

Private Sub lb_data_Click()
    txt_data.BorderStyle = 1
    txt_data.SetFocus
End Sub

Private Sub txt_data_GotFocus()
    txt_data.SelStart = 0
    txt_data.SelLength = Len(txt_data.Text)
    If InStr(1, txt_data.Text, "-") = 0 Then
        m_flag_menos = 0
    Else
        m_flag_menos = 1
    End If
    If InStr(1, txt_data.Text, ",") = 0 Then
        m_flag_punto = 0
    Else
        m_flag_punto = 1
    End If
    If m_flag_punto = 0 And m_flag_menos = 0 And txt_data.Text = "" Then
        m_flag_nume = 0
    End If
End Sub
Private Sub UserControl_GotFocus()
    m_flag_menos = 0
    m_flag_nume = 0
    m_flag_punto = 0
End Sub

Private Sub UserControl_Initialize()
    linea.ZOrder
End Sub
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar textos y gráficos en un objeto."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Devuelve un objeto Font."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Devuelve o establece si los objetos se dibujan en tiempo de ejecución con efectos 3D."
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indica si un control Label o el color de fondo de un control Shape es transparente u opaco."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txt_data,txt_data,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Devuelve o establece el estilo del borde de un objeto."
    BorderStyle = txt_data.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    txt_data.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Obliga a volver a dibujar un objeto."
     
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=lb_data,lb_data,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Devuelve o establece el texto mostrado en la barra de título de un objeto o bajo el icono de un objeto."
    Caption = lb_data.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lb_data.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

Private Sub ud_data_UpClick()
    RaiseEvent UpClick
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txt_data,txt_data,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Devuelve o establece el texto contenido en el control."
    Text = txt_data.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txt_data.Text() = New_Text
    PropertyChanged "Text"
End Property

Private Sub ud_data_DownClick()
    RaiseEvent DownClick
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,0
Public Property Get ud_visible() As Boolean
    ud_visible = m_ud_visible
End Property

Public Property Let ud_visible(ByVal New_ud_visible As Boolean)
    m_ud_visible = New_ud_visible
    If m_ud_visible = True Then
        ud_data.Visible = True
    Else
        ud_data.Visible = False
    End If
    PropertyChanged "ud_visible"
End Property

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_Appearance = m_def_Appearance
    m_BackStyle = m_def_BackStyle
    m_ud_visible = m_def_ud_visible
    m_numerico = False
    m_limite_sup = 9999
    m_limite_inf = -9999
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    txt_data.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    lb_data.Caption = PropBag.ReadProperty("Caption", "Label1")
    txt_data.Text = PropBag.ReadProperty("Text", "Text1")
    m_ud_visible = PropBag.ReadProperty("ud_visible", m_def_ud_visible)
    txt_data.Locked = PropBag.ReadProperty("Locked", False)
    m_numerico = PropBag.ReadProperty("Numerico", False)
    m_limite_sup = PropBag.ReadProperty("Limite_sup", 9999)
    m_limite_inf = PropBag.ReadProperty("Limite_inf", -9999)

End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", txt_data.BorderStyle, 0)
    Call PropBag.WriteProperty("Caption", lb_data.Caption, "Label1")
    Call PropBag.WriteProperty("Text", txt_data.Text, "Text1")
    Call PropBag.WriteProperty("ud_visible", m_ud_visible, m_def_ud_visible)
    Call PropBag.WriteProperty("Locked", txt_data.Locked, False)
    Call PropBag.WriteProperty("Numerico", m_numerico, False)
    Call PropBag.WriteProperty("Limite_sup", m_limite_sup, 9999)
    Call PropBag.WriteProperty("Limite_inf", m_limite_inf, -9999)

End Sub
Private Sub txt_data_Click()
    RaiseEvent Click
End Sub
Private Sub txt_data_Change()
    If m_numerico = True Then
            If Val(txt_data.Text) > m_limite_sup Then
                txt_data.Text = CStr(m_limite_sup)
            End If
            If Val(txt_data.Text) < m_limite_inf Then
                txt_data.Text = CStr(m_limite_inf)
            End If
    End If
        If InStr(1, txt_data.Text, "-") = 0 Then
        m_flag_menos = 0
    Else
        m_flag_menos = 1
    End If
    If InStr(1, txt_data.Text, ",") = 0 Then
        m_flag_punto = 0
    Else
        m_flag_punto = 1
    End If
    If m_flag_punto = 0 And m_flag_menos = 0 And txt_data.Text = "" Then
        m_flag_nume = 0
    End If
    RaiseEvent Change
End Sub

Private Sub txt_data_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txt_data,txt_data,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determina si se puede modificar un control."
    Locked = txt_data.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txt_data.Locked() = New_Locked
    PropertyChanged "Locked"
End Property
Private Sub txt_data_KeyPress(KeyAscii As Integer)
    Dim aux As Integer
    If m_numerico = True Then
        If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 45 And KeyAscii <> Asc(".") And KeyAscii <> 44 Then
            KeyAscii = 1
        End If
        If IsNumeric(Chr(KeyAscii)) Then m_flag_nume = 1
        
        Select Case KeyAscii
        Case Asc(".")
            If m_flag_punto = 0 Then
                KeyAscii = 44
                m_flag_punto = 1
            Else
                KeyAscii = 1
            End If
        Case 44
            If m_flag_punto = 0 Then
                m_flag_punto = 1
            Else
                KeyAscii = 1
            End If
        Case 45
            If (m_flag_menos = 0 And txt_data.SelStart = 0) Or (txt_data.SelLength = Len(txt_data.Text)) Then
                m_flag_menos = 1
            Else
                KeyAscii = 1
            End If
        End Select
    End If
    RaiseEvent KeyPress(KeyAscii)
End Sub

Public Property Get Numerico() As Boolean
    Numerico = m_numerico
End Property
Public Property Let Numerico(New_numerico As Boolean)
    m_numerico = New_numerico
End Property
Public Property Get Limite_Sup() As Double
    Limite_Sup = m_limite_sup
End Property
Public Property Let Limite_Sup(New_limite_sup As Double)
    m_limite_sup = New_limite_sup
End Property
Public Property Get Limite_inf() As Double
    Limite_inf = m_limite_inf
End Property
Public Property Let Limite_inf(New_limite_inf As Double)
    m_limite_inf = New_limite_inf
End Property
Private Sub txt_data_DblClick()
    RaiseEvent DblClick
End Sub

