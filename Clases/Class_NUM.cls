VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class_NUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**************************************
'**************************************
'*** CAMPO NUMERICO *******************
'**************************************
'**************************************

Option Explicit
'* Vercion Inicial de Juan al 6/4/2005

Private m_Name As String
Private m_X As Integer
Private m_Y As Integer
Private m_Mode As String
Private m_Largo As Integer
Private m_Ent As Integer
Private m_Dec As Integer
Private m_Atrib As String
Private m_LRA As String
Private m_Editable As Boolean
Private m_Maximo As Double
Private m_Minimo As Double
Private m_Escalado As Boolean
Private m_Recta As Boolean
Private m_RangX(0 To 1) As Double
Private m_RangY(0 To 1) As Double
Private m_Gain As Double
Private m_Offset As Double
Private m_trigger As String
Private m_trigger_enable As Boolean

'*** Variables de la Clase
Private max_Largo As Integer
Private limite_Maximo As Double
Private limite_minimo As Double
Private max_x_pos As Integer
'***************************************************************************************************
'*******************************************CONSTANTES**********************************************
'***************************************************************************************************
Private Const m_Tipo = "NUMERICO"
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

'*** Propiedad Atributo
Property Get atributo() As String
    atributo = m_Atrib
End Property

Property Let atributo(ByVal new_atributo As String)
    m_Atrib = new_atributo
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
    m_Editable = new_EDIT
    If m_Editable Then
        If m_X < 2 Then m_X = 2
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

'*** Propiedad Largo
Property Get largo() As Long
    largo = m_Largo
End Property

Property Let largo(ByVal new_largo As Long)
    max_x_pos = MAX_LCD_X - m_X + 1
    
    If m_Dec = 0 Then
        If new_largo > max_Largo Then new_largo = max_Largo
        If new_largo < 1 Then new_largo = 1
        m_Ent = new_largo
    Else
        If new_largo > (max_Largo + 1) Then new_largo = max_Largo + 1
        If new_largo < (m_Dec + 2) Then new_largo = m_Dec + 2
        m_Ent = new_largo - 1 - m_Dec
    End If
    
    m_Largo = new_largo
    
    If m_Largo > max_x_pos Then
        m_X = MAX_LCD_X - m_Largo + 1
    End If
    RaiseEvent Change
End Property

'*** Propiedad Enteros
Property Get ent() As Integer
    ent = m_Ent
End Property

Property Let ent(ByVal new_ent As Integer)
    If m_Dec = 0 Then
        If new_ent > max_Largo Then new_ent = max_Largo
    Else
        If new_ent > (max_Largo + 1 - m_Dec) Then new_ent = (max_Largo + 1 - m_Dec)
    End If
    If new_ent < 1 Then new_ent = 1
    
    m_Ent = new_ent
    
    If m_Dec = 0 Then
        largo = m_Ent
    Else
        largo = m_Ent + 1 + m_Dec
    End If
    maximo = m_Maximo
    minimo = m_Minimo
    RaiseEvent Change
End Property

'*** Propiedad Decimales
Property Get dec() As Integer
    dec = m_Dec
End Property
Property Let dec(ByVal new_dec As Integer)
    If new_dec > (max_Largo + 1 - m_Ent) Then new_dec = m_Largo - m_Ent + 1
    If new_dec < 0 Then new_dec = 0
    Select Case m_Mode
        Case "HEX"
            m_Dec = 0
        Case "BIN"
            m_Dec = 0
        Case Else
            m_Dec = new_dec
    End Select
    If m_Dec = 0 Then
        largo = m_Ent
    Else
        largo = m_Ent + 1 + m_Dec
    End If
    
    maximo = m_Maximo
    minimo = m_Minimo
    RaiseEvent Change
End Property

'*** Propiedad Texto Demo
Property Get texto() As String
    Dim auxtexto As String
    If dec = 0 Then
        auxtexto = Mid("####################", 1, m_Largo)
    Else
        auxtexto = Mid("####################", 1, m_Largo - 1 - m_Dec) + "," + Mid("####################", 1, m_Dec)
    End If
    texto = auxtexto
End Property

'*** Propiedad Modo
Property Get modo() As String
    modo = m_Mode
End Property

Property Let modo(ByVal new_modo As String)
    Select Case new_modo
        Case "WS1"
            m_Mode = new_modo
            max_Largo = 5
        Case "W1"
            m_Mode = new_modo
            max_Largo = 5
        Case "WS2"
            m_Mode = new_modo
            max_Largo = 8
        Case "W2"
            m_Mode = new_modo
            max_Largo = 8
        Case "HEX"
            m_Mode = new_modo
            max_Largo = 4
            m_Dec = 0
        Case "BIN"
            m_Mode = new_modo
            max_Largo = 16
            m_Dec = 0
        Case "FPS"
            m_Mode = new_modo
            max_Largo = 8
        Case "FPD"
            m_Mode = new_modo
            max_Largo = 8
        Case Else
            m_Mode = "WS1"
            max_Largo = 8
    End Select
    RaiseEvent Change
End Property

'*** Propiedad Maximo
Property Get maximo() As Double
    maximo = m_Maximo
End Property

Property Let maximo(ByVal new_maximo As Double)
    Dim aux() As String
    Dim aux2 As String

    Select Case m_Mode
        Case "HEX"
            'm_Maximo = ""
        Case "BIN"
            'm_Maximo = ""
        Case Else
            m_Maximo = max_min_limits(new_maximo)
    End Select
End Property

'*** Propiedad Maximo
Property Get minimo() As Double
    minimo = m_Minimo
End Property

Property Let minimo(ByVal new_minimo As Double)
    Select Case m_Mode
        Case "HEX"
            'm_Minimo = ""
        Case "BIN"
            'm_Minimo = ""
        Case Else
            m_Minimo = max_min_limits(new_minimo)
    End Select
End Property

'*** Propiedad Escalado
Property Get usa_escala() As Boolean
    usa_escala = m_Escalado
End Property

Property Let usa_escala(ByVal new_data As Boolean)
    m_Escalado = new_data
    If m_Escalado = False Then
        m_RangX(1) = 2048
        m_RangX(0) = 0
        m_RangY(1) = 2048
        m_RangY(0) = 0
        m_Gain = 1
        m_Offset = 0
        m_Recta = False
    End If
End Property

'*** Propiedad Recta
Property Get recta() As Boolean
    recta = m_Recta
End Property

Property Let recta(ByVal new_recta As Boolean)
    If m_Escalado Then
        m_Recta = new_recta
    Else
        m_Recta = False
    End If
End Property

'*** Propiedad Rango X
Property Get rangoX(Optional ByVal Index As Byte = 0) As Double
    If (Index <> 0) And (Index <> 1) Then Index = 0
    rangoX = m_RangX(Index)
End Property

Property Let rangoX(Optional ByVal Index As Byte = 0, new_rango As Double)
    If (Index <> 0) And (Index <> 1) Then Index = 0
    m_RangX(Index) = new_rango
End Property

'*** Propiedad Rango Y
Property Get rangoY(Optional ByVal Index As Byte = 0) As Double
    If (Index <> 0) And (Index <> 1) Then Index = 0
    rangoY = m_RangY(Index)
End Property

Property Let rangoY(Optional ByVal Index As Byte = 0, new_rango As Double)
    If (Index <> 0) And (Index <> 1) Then Index = 0
    m_RangY(Index) = new_rango
End Property

'*** Propiedad Ganancia
Property Get gain() As Double
    gain = m_Gain
End Property

Property Let gain(ByVal new_rango As Double)
    m_Gain = new_rango
End Property

'*** Propiedad Ganancia
Property Get offset() As Double
    offset = m_Offset
End Property

Property Let offset(ByVal new_rango As Double)
    m_Offset = new_rango
End Property

'*** Inicia Clase
Private Sub Class_Initialize()
    m_X = 1
    m_Y = 1
    m_Atrib = "00"
    m_LRA = "SIM|1:100"
    m_Largo = 5
    m_Dec = 2
    m_Ent = 2
    m_Mode = "WS1"
    max_Largo = 5
    m_Editable = False
    m_Maximo = 99.99
    m_Minimo = -9.99
    m_Escalado = False
    m_Recta = False
    m_RangX(0) = 0
    m_RangX(1) = 2048
    m_RangY(0) = 0
    m_RangY(1) = 2048
    m_Gain = 1
    m_Offset = 0
    m_trigger = "SIM|1:100.0 0"
    m_trigger_enable = False
    RaiseEvent Change
End Sub

Public Function Clone() As Class_NUM
    Dim nueva As Class_NUM
    Set nueva = New Class_NUM
    Dim i As Integer

    With nueva
        .name = name
        .modo = modo
        .dec = dec
        .largo = largo
        .x_pos = x_pos
        .y_pos = y_pos
        .atributo = atributo
        .lra = lra
        .Edit = Edit
        .maximo = maximo
        .minimo = minimo
        .recta = recta
        .rangoX(0) = rangoX(0)
        .rangoX(1) = rangoX(1)
        .rangoY(0) = rangoY(0)
        .rangoY(1) = rangoY(1)
        .gain = gain
        .offset = offset
        .TRIGGER_ENABLE = TRIGGER_ENABLE
        .TRIGGER = TRIGGER
    End With
    Set Clone = nueva

End Function

Private Function max_min_limits(ByVal valor As Double) As Double
    Dim aux() As String
    Dim aux2 As String
   
    If m_Dec > 0 And InStr(1, CStr(valor), ",") <> 0 Then
        aux = Split(CStr(valor), ",")
    Else
        If InStr(1, CStr(valor), ",") = 0 Then
            ReDim aux(0 To 1)
            aux(0) = valor
            aux(1) = String(m_Dec, "0")
        Else
            aux = Split(CStr(valor), ",")
        End If
    End If
    If Len(aux(0)) > m_Ent Then
        If Mid(aux(0), 1, 1) = "-" Then
            If m_Ent > 1 Then
                aux(0) = "-" & String(m_Ent - 1, "9")
            Else
                aux(0) = "0"
            End If
        Else
            aux(0) = String(m_Ent, "9")
        End If
    End If
    If m_Dec > 0 Then
        If Len(aux(1)) > m_Dec Then
            aux(1) = Mid(aux(1), (Len(aux(1)) - m_Dec), m_Dec)
        End If
        valor = CDbl(aux(0) & "," & aux(1))
    Else
        valor = CDbl(aux(0))
    End If
    max_min_limits = valor
End Function
