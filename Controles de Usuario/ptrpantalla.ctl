VERSION 5.00
Begin VB.UserControl ptrpant 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   ScaleHeight     =   90
   ScaleWidth      =   90
   Begin VB.Shape Push 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "ptrpant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Const BLANCO = &HFFFFFF
 Const GRIS = &H808080
 Const NEGRO = &H0&
 Const AZUL = &HFF0000
 Const ROJO = &HFF&
 Const VERDE = &HFF00&
 Const AMARILLO = &HFFFF&
 Const AZUL_SUAVE = &HFFFF00
 Const ROJO_SUAVE = &H8080FF
 Const VERDE_SUAVE = &H80FF80
 Const VERDE_SUAVE2 = &HC0FFC0
 Const AMARILLO_SUAVE = &H80FFFF
 Const GRIS_SUAVE = &HC0C0C0
 Const GRIS_MUY_SUAVE = &HE0E0E0
 Const MARRON_SUAVE = &H40C0&
 Const NARANJA_SUAVE = &H80C0FF
 Const MARRON = &H4080&
 Const NARANJA = &H80FF&
 
 Const CARADEBOTON = &H8000000F
 Const SEPARADOR = &H8000000C
 
 Private Enum ModoScreenS
    SC_LIBRE = 0
    SC_USADO
    SC_NUEVO
    SC_MODIFICADO
    SC_BORRAR
End Enum

'Default Property Values:
Const m_def_Set_Estado = 0
Const m_def_Set_Foco = 0
Const m_def_Set_Open = 0
'Property Variables:
Dim m_Set_Estado As Integer
Dim m_Set_Foco As Boolean
Dim m_Set_Open As Boolean
'Public m_Ventana As VentanaBASE

'Event Declarations:
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Ocurre cuando el usuario mueve el mouse."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Ocurre cuando el usuario presiona y libera un botón del mouse encima de un objeto."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Ocurre cuando el usuario presiona y libera un botón del mouse y después lo vuelve a presionar y liberar sobre un objeto."




Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Push,Push,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."
    BackColor = Push.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Push.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Push,Push,-1,BorderColor
Public Property Get BorderColor() As Long
Attribute BorderColor.VB_Description = "Devuelve o establece el color del borde de un objeto."
    BorderColor = Push.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As Long)
    Push.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Push,Push,-1,BorderWidth
Public Property Get BorderWidth() As Integer
    BorderWidth = Push.BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
    Push.BorderWidth() = New_BorderWidth
    PropertyChanged "BorderWidth"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get Set_Estado() As Integer
    Set_Estado = m_Set_Estado
End Property

Public Property Let Set_Estado(ByVal New_Set_Estado As Integer)
    m_Set_Estado = New_Set_Estado
    PropertyChanged "Set_Estado"
    Select Case New_Set_Estado
        Case SC_LIBRE
            Push.BackColor = CARADEBOTON
        Case SC_BORRAR
            Push.BackColor = CARADEBOTON
        Case SC_USADO
            Push.BackColor = ROJO
        Case SC_NUEVO
            Push.BackColor = VERDE
        Case SC_MODIFICADO
            Push.BackColor = AMARILLO
    End Select
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,0
Public Property Get Set_Foco() As Boolean
    Set_Foco = m_Set_Foco
End Property

Public Property Let Set_Foco(ByVal New_Set_Foco As Boolean)
    m_Set_Foco = New_Set_Foco
    PropertyChanged "Set_Foco"
    If New_Set_Foco Then
        Push.BorderWidth = 4
        If m_Set_Open = False Then
            Push.BorderColor = NEGRO
        End If
    Else
        If m_Set_Open = False Then
            Push.BorderColor = SEPARADOR
        End If
        Push.BorderWidth = 1
    End If
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,0
Public Property Get Set_Open() As Boolean
    Set_Open = m_Set_Open
End Property

Public Property Let Set_Open(ByVal New_Set_Open As Boolean)
    m_Set_Open = New_Set_Open
    PropertyChanged "Set_Open"
    If New_Set_Open Then
        Push.BorderColor = VERDE
    Else
        Push.BorderColor = SEPARADOR
    End If
End Property

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    m_Set_Estado = m_def_Set_Estado
    m_Set_Foco = m_def_Set_Foco
    m_Set_Open = m_def_Set_Open
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Push.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Push.BorderColor = PropBag.ReadProperty("BorderColor", 0)
    Push.BorderWidth = PropBag.ReadProperty("BorderWidth", 1)
    m_Set_Estado = PropBag.ReadProperty("Set_Estado", m_def_Set_Estado)
    m_Set_Foco = PropBag.ReadProperty("Set_Foco", m_def_Set_Foco)
    m_Set_Open = PropBag.ReadProperty("Set_Open", m_def_Set_Open)
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", Push.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BorderColor", Push.BorderColor, 0)
    Call PropBag.WriteProperty("BorderWidth", Push.BorderWidth, 1)
    Call PropBag.WriteProperty("Set_Estado", m_Set_Estado, m_def_Set_Estado)
    Call PropBag.WriteProperty("Set_Foco", m_Set_Foco, m_def_Set_Foco)
    Call PropBag.WriteProperty("Set_Open", m_Set_Open, m_def_Set_Open)
End Sub

