VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class_MTEXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**************************************
'**************************************
'*** CAMPO MTEXT - Texto Menu *********
'**************************************
'**************************************

Option Explicit
'* Vercion Inicial de Juan al 6/4/2005

Private m_Name As String
Private m_X As Integer
Private m_Y As Integer
Private m_Texto As String
Private m_Largo As Integer
Private m_Atrib As String
Private m_JumpTo As Byte

'*** Variables de la Clase
Private max_x_pos As Integer
'***************************************************************************************************
'*******************************************CONSTANTES**********************************************
'***************************************************************************************************
Private Const m_Tipo = "MTEXT"
'Para activar este evento, use RaiseEvent con la siguiente sintaxis:
'RaiseEvent Change
Public Event Change()

Property Get tipo_campo() As String
    tipo_campo = m_Tipo
End Property

'*** Propiedad Name
Property Get name() As String
    name = m_Name
End Property

Property Let name(ByVal new_Name As String)
    m_Name = new_Name
End Property

'*** Propiedad X
Property Get x_pos() As Integer
    x_pos = m_X
End Property

Property Let x_pos(ByVal new_x As Integer)
    max_x_pos = MAX_LCD_X - m_Largo + 1
    If new_x <= max_x_pos Then
        m_X = new_x
    Else
        m_X = max_x_pos
    End If
    If m_X < 2 Then m_X = 2
    RaiseEvent Change
End Property

'*** Propiedad Y
Property Get y_pos() As Integer
    y_pos = m_Y
End Property

Property Let y_pos(ByVal new_y As Integer)
    If new_y <= MAX_LCD_Y Then
        m_Y = new_y
    Else
        m_Y = MAX_LCD_Y
    End If
    If m_Y < 1 Then m_Y = 1
    RaiseEvent Change
End Property

'*** Propiedad Text
Property Get texto() As String
    texto = m_Texto
End Property

Property Let texto(ByVal New_Text As String)
    max_x_pos = MAX_LCD_X - m_X + 1
    
    If Len(New_Text) < max_x_pos Then
        m_Texto = New_Text
    Else
        m_Texto = Mid(New_Text, 1, max_x_pos)
    End If
    m_Largo = Len(m_Texto)
    RaiseEvent Change
End Property

'*** Propiedad Widty (no necesaria para el INI file)
Property Get largo() As Integer
    largo = m_Largo
End Property

'*** Propiedad Atributo
Property Get atributo() As String
    atributo = m_Atrib
End Property

Property Let atributo(ByVal new_atributo As String)
    m_Atrib = new_atributo
End Property

'*** Propiedad Jump To ...
Property Get jmp() As Byte
    jmp = m_JumpTo
End Property

Property Let jmp(ByVal new_jmp As Byte)
    m_JumpTo = new_jmp
End Property

Private Sub Class_Initialize()
    m_Atrib = "00"
    m_Texto = "MTEXT"
    m_Largo = 5
    m_X = 1
    m_Y = 1
    m_JumpTo = 0
    RaiseEvent Change
End Sub

Public Function Clone() As Class_MTEXT
    Dim nueva As Class_MTEXT
    Set nueva = New Class_MTEXT

    With nueva
        .name = name
        .texto = texto
        .x_pos = x_pos
        .y_pos = y_pos
'       .largo = largo  No tiene let para esta clase
        .atributo = atributo
        .jmp = jmp
    End With
    Set Clone = nueva

End Function


