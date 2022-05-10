VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class_ALFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**************************************
'**************************************
'*** CAMPO ALFANUMERICO ***************
'**************************************
'**************************************

Option Explicit
'* Vercion Inicial de Juan al 6/4/2005

Private m_Name As String
Private m_X As Integer
Private m_Y As Integer
Private m_Largo As Integer
Private m_Atrib As String
Private m_LRA As String
Private m_Editable As Boolean
Private m_trigger As String
Private m_trigger_enable As Boolean

'*** Variables de la Clase
Private max_x_pos As Integer
'********************************************************************************************************************
'**************************************************CONSTANTES********************************************************
'********************************************************************************************************************
Private Const m_Tipo = "ALFANUM"
Private Const max_text = "@@@@@@@@@@@@@@@@@@@@"

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
    If m_Editable Then
        If m_X < 2 Then m_X = 2
    Else
        If m_X < 1 Then m_X = 1
    End If
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

'*** Propiedad Largo
Property Get largo() As Long
    largo = m_Largo
End Property

Property Let largo(ByVal new_largo As Long)
    max_x_pos = MAX_LCD_X - m_X + 1
    If new_largo <= max_x_pos Then
        m_Largo = new_largo
    Else
        m_Largo = max_x_pos
    End If
    RaiseEvent Change
End Property

'*** Propiedad Atributo
Property Get atributo() As String
    atributo = m_Atrib
End Property

Property Let atributo(ByVal new_atributo As String)
    m_Atrib = new_atributo
End Property

'*** Propiedad Text
Property Get texto() As String
    texto = Mid(max_text, 1, m_Largo)
End Property

'*** Propiedad LRA (Link Reference Address)
Property Get lra() As String
    lra = m_LRA
End Property

Property Let lra(ByVal new_lra As String)
    m_LRA = new_lra
End Property

'*** Propiedad Editable
Property Get Edit() As Boolean
    Edit = m_Editable
End Property

Property Let Edit(ByVal new_EDIT As Boolean)
    Dim ret
    m_Editable = new_EDIT
    If m_Editable Then
        If m_X < 2 Then
            If m_Largo = 20 Then
                ret = MsgBox("La propiedad Editable, limita el texto a 19 caractares." + Chr(13) + "Presione SI, si desea truncar los textos a 19 caracteres," + Chr(13) + "Presione NO, si decea anular la edicion.", vbYesNo, "Campo AFLANUM")
                If ret = vbYes Then
                    m_X = 2
                    largo = 19
                Else
                    m_Editable = False
                End If
            Else
                m_X = 2
            End If
        End If
    Else
        If m_X < 1 Then m_X = 1
    End If
    TRIGGER_ENABLE = False
End Property
'*** Propiedad Trigger enable *****
Property Get TRIGGER_ENABLE() As Boolean
    TRIGGER_ENABLE = m_trigger_enable
End Property
Property Let TRIGGER_ENABLE(new_TRIGGER_Enable As Boolean)
    m_trigger_enable = new_TRIGGER_Enable
    If m_Editable = False Then m_trigger_enable = False
End Property

'*** Propiedad Trigger *****
Property Get TRIGGER() As String
    TRIGGER = m_trigger
End Property
Property Let TRIGGER(ByVal new_trigger As String)
    m_trigger = new_trigger
End Property

Private Sub Class_Initialize()
    m_LRA = "SIM|1:100"
    m_Atrib = "00"
    m_Largo = 5
    m_X = 1
    m_Y = 1
    m_Editable = False
    m_trigger = "SIM|1:100.0 0"
    m_trigger_enable = False
    RaiseEvent Change
End Sub

Public Function Clone() As Class_ALFA
    Dim nueva As Class_ALFA
    Set nueva = New Class_ALFA

    With nueva
        .name = name
        .largo = largo
        .x_pos = x_pos
        .y_pos = y_pos
        .atributo = atributo
        .lra = lra
        .Edit = Edit
        .TRIGGER_ENABLE = TRIGGER_ENABLE
        .TRIGGER = TRIGGER
    End With
    Set Clone = nueva

End Function


