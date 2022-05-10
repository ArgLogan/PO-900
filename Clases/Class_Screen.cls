VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class_Pantalla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'***************************************************************
'**** Clase de Pantalla ****************************************
'***************************************************************

Option Explicit

'variables locales para almacenar los valores de las propiedades
Private m_IDscreen As String
Private m_Modo As Byte
Private m_Nombre As String
Private m_Numero As Byte
Private m_TimeScan As Integer
Private m_TimeScan_Local As Boolean
Private m_TimeOut As Integer
Private m_TimeOutJmp As Byte
Private m_TimeOut_Local As Boolean
Private m_EscapeJmp As Byte
Private m_EscapeJmp_Local As Boolean
Private m_TimeOutEdit As Integer
Private m_TimeAutoCursor As Integer
Private m_TimeEdit_Local As Boolean
Private m_KeyJmp(0 To 13) As Byte
Private m_KeyJmpEnable(0 To 13) As Boolean
Private m_ColectCampo As Collection

Public m_pantalla As pantalla
Private m_lra_bit_set(13) As String 'test, ver si queda aca

'*** Variables de la Clase
Private max_x_pos As Integer
'********************************************************************************************************************
'**************************************************CONSTANTES********************************************************
'********************************************************************************************************************
Private Const m_Tipo = "PANTALLA"

Property Get tipo_campo() As String
    tipo_campo = m_Tipo
End Property

'**** Propiedad IDscreen (key de colection screens)
Public Property Get idscreen() As String
    idscreen = m_IDscreen
End Property

Public Property Let idscreen(ByVal vdata As String)
    If (m_IDscreen = "VACIO") Or (m_IDscreen = "COPY") Then
        m_IDscreen = vdata
    End If
End Property

'**** Propiedad Modo (Estado de uso)
Public Property Get modo() As Byte
    If m_IDscreen = "VACIO" Then
        modo = SC_LIBRE
    Else
        modo = m_Modo
    End If
End Property

Public Property Let modo(ByVal vdata As Byte)
    m_Modo = vdata
End Property

'*** Propiedad Numero
Public Property Get Numero() As Byte
    Numero = m_Numero
End Property

Public Property Let Numero(ByVal vdata As Byte)
    m_Numero = vdata
End Property

'*** Propiedad Nombre
Public Property Get name() As String
    name = m_Nombre
End Property

Public Property Let name(ByVal vdata As String)
    m_Nombre = vdata
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

Public Property Get tdisplay_local() As Boolean
    tdisplay_local = m_TimeScan_Local
End Property
Public Property Let tdisplay_local(ByVal nval As Boolean)
    m_TimeScan_Local = nval
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

Public Property Get timeout_local() As Boolean
    timeout_local = m_TimeOut_Local
End Property
Public Property Let timeout_local(ByVal nval As Boolean)
    m_TimeOut_Local = nval
End Property

'*** Propiedad Salto de Escape
Public Property Get escjmp() As Byte
    escjmp = m_EscapeJmp
End Property

Public Property Let escjmp(ByVal new_salto As Byte)
    m_EscapeJmp = new_salto
End Property

Public Property Get escjmp_local() As Boolean
    escjmp_local = m_EscapeJmp_Local
End Property
Public Property Let escjmp_local(ByVal nval As Boolean)
    m_EscapeJmp_Local = nval
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

Public Property Get timeedit_local() As Boolean
    timeedit_local = m_TimeEdit_Local
End Property
Public Property Let timeedit_local(ByVal nval As Boolean)
    m_TimeEdit_Local = nval
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

'*** Propiedad Coleccion de Campos
Public Property Get colectcampo() As Collection
    Set colectcampo = m_ColectCampo
End Property

Public Property Set colectcampo(ByVal vdata As Collection)
    Set m_ColectCampo = vdata
End Property

Private Sub Class_Initialize()
    Dim indice As Integer
    
    m_IDscreen = "VACIO"
    m_Modo = SC_LIBRE
    m_Nombre = "Pantalla"
    m_Numero = 0
    m_TimeScan = 100
    m_TimeScan_Local = False
    m_TimeOut = 60
    m_TimeOutJmp = &H0
    m_TimeOut_Local = False
    m_EscapeJmp = &HFF
    m_EscapeJmp_Local = False
    m_TimeOutEdit = 30
    m_TimeAutoCursor = 2
    m_TimeEdit_Local = False
    
    For indice = 0 To 13
        m_KeyJmp(indice) = 0
        m_KeyJmpEnable(indice) = False
        m_lra_bit_set(indice) = "OFF"
    Next indice

    Set m_ColectCampo = New Collection

End Sub


Public Function Clone() As Class_Pantalla
    Dim i As Byte
    Dim nueva As Class_Pantalla
    Dim tempcampo As Variant
    Dim newcampo As Variant
    Set nueva = New Class_Pantalla

   With nueva
        .idscreen = "COPY"
        .modo = SC_NUEVO
        .name = name
        .Numero = Numero
                
        .tdisplay = tdisplay
        .tdisplay_local = tdisplay_local
        .timeout = timeout
        .timeoutjmp = timeoutjmp
        .timeout_local = timeout_local
        .escjmp = escjmp
        .escjmp_local = escjmp_local
        .timeoutedit = timeoutedit
        .timeautocursor = timeautocursor
        .timeedit_local = timeedit_local
        
        For i = 0 To 13
            .keyjmp(i) = keyjmp(i)
            .keyjmpenable(i) = keyjmpenable(i)
            .Key_LRA(i) = Key_LRA(i)
        Next i
        
        For i = 1 To colectcampo.count
            Set tempcampo = colectcampo.item(i)
            Select Case tempcampo.tipo_campo
                Case "CTEXT"
                    Set newcampo = New Class_MTEXT
                    Set newcampo = tempcampo.Clone
                    .colectcampo.Add newcampo, genidcampo(tempcampo.tipo_campo, i)
                Case "MTEXT"
                    Set newcampo = New Class_MTEXT
                    Set newcampo = tempcampo.Clone
                    .colectcampo.Add newcampo, genidcampo(tempcampo.tipo_campo, i)
                Case "MTDIGITAL"
                    Set newcampo = New Class_MTD
                    Set newcampo = tempcampo.Clone
                    .colectcampo.Add newcampo, genidcampo(tempcampo.tipo_campo, i)
                Case "ALFANUM"
                    Set newcampo = New Class_ALFA
                    Set newcampo = tempcampo.Clone
                    .colectcampo.Add newcampo, genidcampo(tempcampo.tipo_campo, i)
                Case "NUMERICO"
                    Set newcampo = New Class_NUM
                    Set newcampo = tempcampo.Clone
                    .colectcampo.Add newcampo, genidcampo(tempcampo.tipo_campo, i)
            End Select
        
        Next i
        
    End With
    Set Clone = nueva

End Function

 

