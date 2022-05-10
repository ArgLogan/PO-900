VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class_GLOBAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'***************************************************************
'**** Clase de Globales ****************************************
'***************************************************************

Option Explicit

'variables locales para almacenar los valores de las propiedades

Private m_SyncLRA As String
Private m_SyncTime As Integer
Private m_SyncActivo As Boolean

'Variables default de pantalla
Private m_TimeScan As Integer
Private m_TimeOut As Integer
Private m_TimeOutJmp As Byte
Private m_EscapeJmp As Byte
Private m_TimeOutEdit As Integer
Private m_TimeAutoCursor As Integer
Private m_KeyJmp(0 To 13) As Byte
Private m_KeyJmpEnable(0 To 13) As Boolean
Private m_lra_bit_set(13) As String
Private m_Pant_Principal As Byte
Private m_Pant_Principal_Enable As Boolean
Private m_Pant_Inicial As Byte
Private m_Pant_Inicial_Enable As Boolean
Private m_puerto As Byte
Private m_velocidad As Long
Private m_delay_resp As Byte
Private m_rts_on As Byte
Private m_rts_off As Byte
Private m_invertido As Boolean

'*** Variables de la Clase
Private max_x_pos As Integer
'********************************************************************************************************************
'**************************************************CONSTANTES********************************************************
'********************************************************************************************************************
Private Const m_Tipo = "GLOBAL"
'*** Propiedad Invertido
Public Property Get invertido() As Boolean
    invertido = m_invertido
End Property
Public Property Let invertido(new_invertido As Boolean)
    m_invertido = new_invertido
End Property
'*** Propiedad delay respuesta
Public Property Get delay_resp() As Byte
    delay_resp = m_delay_resp
End Property
Public Property Let delay_resp(new_delay_resp As Byte)
    m_delay_resp = new_delay_resp
End Property
'*** Propiedad Puerto Comm
Public Property Get puerto() As Byte
    puerto = m_puerto
End Property
Public Property Let puerto(new_puerto As Byte)
    m_puerto = new_puerto
End Property
'*** Propiedad Velocidad de comunicación
Public Property Get velocidad() As Long
    velocidad = m_velocidad
End Property
Public Property Let velocidad(new_velocidad As Long)
     m_velocidad = new_velocidad
End Property
'*** Propiedad rts on
Public Property Get rts_on() As Byte
    rts_on = m_rts_on
End Property
Public Property Let rts_on(new_rts_on As Byte)
    m_rts_on = new_rts_on
End Property
'*** Propiedad rts off
Public Property Get rts_off() As Byte
    rts_off = m_rts_off
End Property
Public Property Let rts_off(new_rts_off As Byte)
    m_rts_off = new_rts_off
End Property
'*** Propiedad  tipo campo
Public Property Get tipo_campo() As String
    tipo_campo = m_Tipo
End Property
'*** Propiedad Pantalla principal
Public Property Get Pant_principal() As Byte
    Pant_principal = m_Pant_Principal
End Property
Public Property Let Pant_principal(new_pant_principal As Byte)
    If new_pant_principal < 0 Then new_pant_principal = 0
    If new_pant_principal > 254 Then new_pant_principal = 254
    m_Pant_Principal = new_pant_principal
End Property
Public Property Get Pant_principal_enable() As Boolean
    Pant_principal_enable = m_Pant_Principal_Enable
End Property
Public Property Let Pant_principal_enable(ByVal vdata As Boolean)
    m_Pant_Principal_Enable = vdata
End Property
'*** Propiedad Pantalla Inicial
Public Property Get Pant_inicial() As Byte
    Pant_inicial = m_Pant_Inicial
End Property
Public Property Let Pant_inicial(new_pant_inicial As Byte)
    If new_pant_inicial < 0 Then new_pant_inicial = 0
    If new_pant_inicial > 254 Then new_pant_inicial = 254
    m_Pant_Inicial = new_pant_inicial
End Property
'*** Propiedad Pantalla Inicial enable
Public Property Get Pant_inicial_enable() As Boolean
    Pant_inicial_enable = m_Pant_Inicial_Enable
End Property
Public Property Let Pant_inicial_enable(ByVal vdata As Boolean)
    m_Pant_Inicial_Enable = vdata
End Property
'*** Propiedad Syncro LRA
Public Property Get synclra() As String
    synclra = m_SyncLRA
End Property
Public Property Let synclra(ByVal new_lra As String)
    m_SyncLRA = new_lra
End Property

'** Propiedad Syncro Time
Public Property Get synctime() As Integer
    synctime = m_SyncTime
End Property

Public Property Let synctime(ByVal new_time As Integer)
    If new_time > 100 Then new_time = 100
    If new_time < 0 Then new_time = 0
    m_SyncTime = new_time
End Property

'*** Propiedad Salto de Pantalla
Public Property Get syncactivo() As Boolean
    syncactivo = m_SyncActivo
End Property

Public Property Let syncactivo(ByVal new_activo As Boolean)
    m_SyncActivo = new_activo
End Property

'*** Propiedad Time Scan
Public Property Get tdisplay() As Integer
    tdisplay = m_TimeScan
End Property

Public Property Let tdisplay(ByVal new_tdisplay As Integer)
    If new_tdisplay > 3000 Then new_tdisplay = 100
    If new_tdisplay < 0 Then new_tdisplay = 100
    m_TimeScan = new_tdisplay
End Property

'*** Propiedad Tiempo de Salida de pantalla
Public Property Get timeout() As Integer
    timeout = m_TimeOut
End Property

Public Property Let timeout(ByVal new_timeout As Integer)
    If new_timeout > 60 Then new_timeout = 10
    If new_timeout < 0 Then new_timeout = 0
    m_TimeOut = new_timeout
End Property

'*** Propiedad Salto por Tiempo de Salida de Pantalla
Public Property Get timeoutjmp() As Byte
    timeoutjmp = m_TimeOutJmp
End Property

Public Property Let timeoutjmp(ByVal new_salto As Byte)
    m_TimeOutJmp = new_salto
End Property

'*** Propiedad Salto de Escape
Public Property Get escjmp() As Byte
    escjmp = m_EscapeJmp
End Property

Public Property Let escjmp(ByVal new_salto As Byte)
    m_EscapeJmp = new_salto
End Property

'** Propiedad TimeOut Edicion
Public Property Get timeoutedit() As Integer
    timeoutedit = m_TimeOutEdit
End Property

Public Property Let timeoutedit(ByVal new_time As Integer)
    If new_time > 100 Then new_time = 100
    If new_time < 0 Then new_time = 0
    m_TimeOutEdit = new_time
End Property

'*** Propiedad Tiempo de AutoCursor
Public Property Get timeautocursor() As Integer
    timeautocursor = m_TimeAutoCursor
End Property

Public Property Let timeautocursor(ByVal new_time As Integer)
    If new_time > 100 Then new_time = 100
    If new_time < 2 Then new_time = 2
    m_TimeAutoCursor = new_time
End Property

'*** Propiedad Salto de Pantalla
Public Property Get keyjmp(Optional ByVal tecla As Byte = &HA) As Byte
    If tecla <= 13 Then
        keyjmp = m_KeyJmp(tecla)
    End If
End Property

Public Property Let keyjmp(Optional ByVal tecla As Byte = &HA, new_jmp As Byte)
    If tecla <= 13 Then
        m_KeyJmp(tecla) = new_jmp
        m_KeyJmpEnable(tecla) = True
    End If
End Property

'*** Propiedad Salto de Pantalla
Public Property Get keyjmpenable(Optional ByVal tecla As Byte = &HA) As Boolean
    If tecla <= 13 Then
        keyjmpenable = m_KeyJmpEnable(tecla)
    End If
End Property

Public Property Let keyjmpenable(Optional ByVal tecla As Byte = &HA, new_jmp As Boolean)
    If tecla <= 13 Then
        m_KeyJmpEnable(tecla) = new_jmp
    End If
End Property
'**** Prepiedad Bit set LRA******
Public Property Get Key_LRA(ByVal tecla As Byte) As String
    Key_LRA = m_lra_bit_set(tecla)
End Property
Public Property Let Key_LRA(ByVal tecla As Byte, new_key_lra As String)
    m_lra_bit_set(tecla) = new_key_lra
End Property

Private Sub Class_Initialize()
    Dim indice As Integer
    
    m_TimeScan = 100
    m_TimeOut = 0
    m_TimeOutJmp = &H0
    m_EscapeJmp = &HFF
    m_TimeOutEdit = 30
    m_TimeAutoCursor = 2
    m_Pant_Principal = 16
    m_SyncLRA = "SIM|1:100"
    m_SyncTime = 100
    m_SyncActivo = False
    m_invertido = False
    m_delay_resp = 20
    m_rts_on = 10
    m_rts_off = 10
    m_puerto = 1
    m_velocidad = 38400

    For indice = 0 To 13
        m_KeyJmp(indice) = 0
        m_KeyJmpEnable(indice) = False
        m_lra_bit_set(indice) = "OFF"
    Next indice

End Sub

Public Function Clone() As Class_GLOBAL
    Dim i As Byte
    Dim nueva As Class_GLOBAL
    Set nueva = New Class_GLOBAL

    With nueva
        .syncactivo = syncactivo
        .synclra = synclra
        .synctime = synctime
        .Pant_principal = Pant_principal
        .tdisplay = tdisplay
        .timeout = timeout
        .timeoutjmp = timeoutjmp
        .escjmp = escjmp
        .timeoutedit = timeoutedit
        .timeautocursor = timeautocursor
        .delay_resp = delay_resp
        .puerto = puerto
        .velocidad = velocidad
        .rts_on = rts_on
        .rts_off = rts_off
        .invertido = invertido
        
        For i = 0 To 13
            .keyjmp(i) = keyjmp(i)
            .keyjmpenable(i) = keyjmpenable(i)
            .Key_LRA(i) = Key_LRA(i)
        Next i
    End With
    Set Clone = nueva

End Function


 
