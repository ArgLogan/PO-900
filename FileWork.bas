Attribute VB_Name = "FileWork"
Option Explicit

'*****************************************************************************
' Recupera del file de proyecto el
' [grupo]
' campo=default
Public Function getoffile(archivo As String, Grupo As String, campo As String, default As String) As String
        Dim tb$
        Dim rc%
        
        tb$ = String$(GS_MAX, " ")
        rc% = GetPrivateProfileString(Grupo, campo, default, tb$, GS_MAX - 1, archivo)
        tb$ = token(tb$, Chr$(0))
        getoffile = tb$
End Function
'*****************************************************************************
' Escribe al file de proyecto el
' [grupo]
' campo=default
Public Function puttofile(archivo As String, Grupo As String, campo As String, data As String) As String
        Dim rc%
        rc% = WritePrivateProfileString(Grupo, campo, data, archivo)
End Function

Public Function getitem(item As String, Optional coment As Boolean = False) As String
    Dim separa() As String
    If InStr(1, item, CTE_SEP_COMENT) Then
        separa = Split(item, CTE_SEP_COMENT, 2)
        If coment = False Then
            getitem = separa(0)
        Else
            getitem = separa(1)
        End If
    Else
        getitem = item
    End If
End Function


'***************************************************** READ ********************************************************

Function read_gobales(destino As Class_GLOBAL, archivo As String) As Boolean
    Dim dataitem As String
    Dim key_item As String
    Dim seccion As String
    Dim tecla As Byte
    Dim texttemp() As String
    If Dir(archivo) = "" Then
        read_gobales = False
        Exit Function
    End If
       
    
    '[COMM]
    'PORT=1
    destino.puerto = CByte(getitem(getoffile(archivo, "COMM", "PORT", "1")))
    'BAUD=38400
    destino.velocidad = CLng(getitem(getoffile(archivo, "COMM", "BAUD", "38400")))
    'RTSMODE= 0
    destino.invertido = CBool(getitem(getoffile(archivo, "COMM", "RTSMODE", "0")))
    'RTSON= 10
    destino.rts_on = CByte(getitem(getoffile(archivo, "COMM", "RTSON", "10")))
    'RTSOFF=10
    destino.rts_off = CByte(getitem(getoffile(archivo, "COMM", "RTSOFF", "10")))
    'REPLYDELAY = 20
    destino.delay_resp = CByte(getitem(getoffile(archivo, "COMM", "REPLYDELAY", "20")))
    
    ' [SYNC_PLC]
    ' SYNC_LRA=SIM|1:100
    destino.synclra = getitem(getoffile(archivo, "SYNC_PLC", "SYNC_LRA", "SIM|1:100"))
    ' SYNC_TIME=0
    destino.synctime = Val(getitem(getoffile(archivo, "SYNC_PLC", "SYNC_TIME", "0")))
    ' SYNC_PLC_ACTIVO=OFF
    If getitem(getoffile(archivo, "SYNC_PLC", "SYNC_PLC_ACTIVO", "OFF")) = "ON" Then
        destino.syncactivo = True
    Else
        destino.syncactivo = False
    End If
    
    
    ' [VARIABLES_DEFAULT]
    ' PANTALLA_PRINCIPAL = 16
    dataitem = getitem(getoffile(archivo, "VARIABLES_DEFAULT", "PANTALLA_PRINCIPAL", "none"))
    If dataitem <> "none" Then
        destino.Pant_principal_enable = True
        destino.Pant_principal = Val(dataitem)
    End If
    
    ' PANTALLA_INICIAL = 0
    dataitem = getitem(getoffile(archivo, "VARIABLES_DEFAULT", "PANTALLA_INICIAL", "none"))
    If dataitem <> "none" Then
        destino.Pant_inicial_enable = True
        destino.Pant_inicial = Val(dataitem)
    End If
    
    ' TDISPLAY = 10
    destino.tdisplay = Val(getitem(getoffile(archivo, "VARIABLES_DEFAULT", "TDISPLAY", "100")))
    
    ' TACTIVO=30 0
    texttemp = Split(getitem(getoffile(archivo, "VARIABLES_DEFAULT", "TACTIVO", "0 0")), " ")
    destino.timeout = Val(texttemp(0))
    destino.timeoutjmp = Val(texttemp(1))
    
    ' ESCJMP=-1
    dataitem = getitem(getoffile(archivo, "VARIABLES_DEFAULT", "ESCJMP", "-1"))
    If dataitem = "-1" Then
        destino.escjmp = 255
    Else
        destino.escjmp = (Val(dataitem))
    End If
        
    ' TEDIT=30 2
    texttemp = Split(getitem(getoffile(archivo, "VARIABLES_DEFAULT", "TEDIT", "30 2")), " ")
    destino.timeoutedit = Val(texttemp(0))
    destino.timeautocursor = Val(texttemp(1))
    
    
    ' JMP_TECLA_x=x   y  BITSET_TECLA_x=SIM|1:100.1 0
    For tecla = 0 To 9
        'JUMP
        key_item = "JMP_TECLA_" + Mid(Str(tecla), 2, 1)
        dataitem = getitem(getoffile(archivo, "VARIABLES_DEFAULT", key_item, ""))
        If dataitem <> "" Then
            destino.keyjmpenable(tecla) = True
            destino.keyjmp(tecla) = Val(dataitem)
        End If
        
        'BIT SET/CLEAR
        key_item = "BITSET_TECLA_" + Mid(Str(tecla), 2, 1)
        dataitem = getitem(getoffile(archivo, "VARIABLES_DEFAULT", key_item, ""))
        If dataitem <> "" Then
            destino.Key_LRA(tecla) = dataitem
        End If
    Next tecla
    
    For tecla = 1 To 4
        'JUMP
        key_item = "JMP_TECLA_F" + Mid(Str(tecla), 2, 1)
        dataitem = getitem(getoffile(archivo, "VARIABLES_DEFAULT", key_item, ""))
        If dataitem <> "" Then
            destino.keyjmpenable(tecla + 9) = True
            destino.keyjmp(tecla + 9) = Val(dataitem)
        End If
        
        'BIT SET/CLEAR
        key_item = "BITSET_TECLA_F" + Mid(Str(tecla), 2, 1)
        dataitem = getitem(getoffile(archivo, "VARIABLES_DEFAULT", key_item, ""))
        If dataitem <> "" Then
            destino.Key_LRA(tecla + 9) = dataitem
        End If
    Next tecla
    

End Function

Function read_Pantalla(Collection_destino As Collection, idnum As Byte, archivo As String) As Boolean
    Dim dataitem As String
    Dim key_item As String
    Dim seccion As String
    Dim tecla As Byte
    Dim texttemp() As String
    Dim destino As Class_Pantalla
    Dim ret As Boolean
    
    If Dir(archivo) = "" Then
        read_Pantalla = False
        Exit Function
    End If
   
    Set destino = New Class_Pantalla
    destino.modo = SC_USADO
    destino.Numero = idnum
    destino.idscreen = genidpan(idnum)
    
    Collection_destino.Add destino, destino.idscreen
    
    
    ' [VARIABLES]
    ' Number=0 (No se lee del archivo)
    ' Namre = !!!
    destino.name = getitem(getoffile(archivo, "VARIABLES", "NAME", "Pantalla " + Format$(Hex(idnum), "00")))
   
    ' TDISPLAY = 10
    dataitem = getitem(getoffile(archivo, "VARIABLES", "TDISPLAY", "none"))
    If dataitem <> "none" Then
        destino.tdisplay_local = True
        destino.tdisplay = Val(dataitem)
    End If
    
    
    ' TACTIVO=30 0
    dataitem = getitem(getoffile(archivo, "VARIABLES", "TACTIVO", "none"))
    If dataitem <> "none" Then
        destino.timeout_local = True
        texttemp = Split(dataitem, " ")
        destino.timeout = Val(texttemp(0))
        destino.timeoutjmp = Val(texttemp(1))
    End If
    
    ' ESCJMP=-1
    dataitem = getitem(getoffile(archivo, "VARIABLES", "ESCJMP", "none"))
    If dataitem <> "none" Then
        destino.escjmp_local = True
        If dataitem = "-1" Then
            destino.escjmp = 255
        Else
            destino.escjmp = (Val(dataitem))
        End If
    End If
    
    ' TEDIT=30 2
    dataitem = getitem(getoffile(archivo, "VARIABLES", "TEDIT", "none"))
    If dataitem <> "none" Then
        destino.timeedit_local = True
        texttemp = Split(dataitem, " ")
        destino.timeoutedit = Val(texttemp(0))
        destino.timeautocursor = Val(texttemp(1))
    End If
    
    ' JMP_TECLA_x=x   Y BITSET_TECLA_x=SIM|1:100.1 0
    For tecla = 0 To 9
        'JUMP
        key_item = "JMP_TECLA_" + Mid(Str(tecla), 2, 1)
        dataitem = getitem(getoffile(archivo, "VARIABLES", key_item, ""))
        If dataitem <> "" Then
            destino.keyjmpenable(tecla) = True
            destino.keyjmp(tecla) = Val(dataitem)
        End If
    
        'BIT SET/CLEAR
        key_item = "BITSET_TECLA_" + Mid(Str(tecla), 2, 1)
        dataitem = getitem(getoffile(archivo, "VARIABLES", key_item, ""))
        If dataitem <> "" Then
            destino.Key_LRA(tecla) = dataitem
        End If
    Next tecla
    
    For tecla = 1 To 4
        'JUMP
        key_item = "JMP_TECLA_F" + Mid(Str(tecla), 2, 1)
        dataitem = getitem(getoffile(archivo, "VARIABLES", key_item, ""))
        If dataitem <> "" Then
            destino.keyjmpenable(tecla + 9) = True
            destino.keyjmp(tecla + 9) = Val(dataitem)
        End If
        
        'BIT SET/CLEAR
        key_item = "BITSET_TECLA_F" + Mid(Str(tecla), 2, 1)
        dataitem = getitem(getoffile(archivo, "VARIABLES", key_item, ""))
        If dataitem <> "" Then
            destino.Key_LRA(tecla + 9) = dataitem
        End If
    Next tecla
    
    read_Pantalla = read_Campos(destino.colectcampo, archivo)

End Function

Function read_Campos(coll_campos As Collection, archivo As String) As Boolean
    Dim dataitem As String
    Dim datadefault As String
    Dim key_item As String
    Dim seccion As String
    Dim Campo_tipo As String
    Dim texttemp() As String
    Dim NroCampo As Integer
    Dim class_temp As Variant
    Dim i As Integer
    NroCampo = 1
    
    Do
        seccion = "CAMPO" + Format$(NroCampo, "00")
        key_item = "TIPO"
        datadefault = "NO_TYPE"
        dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
        Campo_tipo = dataitem
        If Campo_tipo <> "NO_TYPE" Then
            Select Case dataitem
                Case "CTEXT"
                    Set class_temp = New class_CTEXT
                    coll_campos.Add class_temp, "campo CT-" + Format$(NroCampo, "00")
                    
                    ' [CAMPO01]
                    ' TIPO = CTEXT ; Nombre
                    'class_temp.name = getitem(getoffile(archivo, seccion, key_item, ""), True)
                    key_item = "REM"
                    dataitem = getoffile(archivo, seccion, key_item, "")
                    class_temp.name = dataitem
                    
                    ' TEXTO=MENU PRINCIPAL 16/3
                    key_item = "TEXTO"
                    datadefault = "CTEXT" + Format$(NroCampo, "00")
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.texto = dataitem
                        
                    ' POSXY=02 01
                    key_item = "POSXY"
                    datadefault = "1 1"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    texttemp = Split(dataitem, " ", 2)
                    class_temp.x_pos = Val(texttemp(0))
                    class_temp.y_pos = Val(texttemp(1))
                        
                    ' ATRIBUTOS=07
                    key_item = "ATRIBUTOS"
                    datadefault = "00"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.atributo = dataitem
                    
                Case "MTEXT"
                    Set class_temp = New Class_MTEXT
                    coll_campos.Add class_temp, "campo MT-" + Format$(NroCampo, "00")
                    
                    ' [CAMPO01]
                    ' TIPO = MTEXT ; Nombre
                    'class_temp.name = getitem(getoffile(archivo, seccion, key_item, ""), True)
                    key_item = "REM"
                    dataitem = getoffile(archivo, seccion, key_item, "")
                    class_temp.name = dataitem
                    
                    ' ATRIBUTOS=07
                    key_item = "ATRIBUTOS"
                    datadefault = "00"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.atributo = dataitem
                    
                    ' TEXTO=MENU PRINCIPAL 16/3
                    key_item = "TEXTO"
                    datadefault = "MTEXT" + Format$(NroCampo, "00")
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.texto = dataitem
                    
                    ' POSXY=02 01
                    key_item = "POSXY"
                    datadefault = "1 1"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    texttemp = Split(dataitem, " ", 2)
                    class_temp.x_pos = Val(texttemp(0))
                    class_temp.y_pos = Val(texttemp(1))
                    
                    ' JMP=01
                    key_item = "JMP"
                    datadefault = "0"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    If dataitem = "-1" Then
                        class_temp.jmp = 255
                    Else
                        class_temp.jmp = (Val(dataitem))
                    End If
                
                Case "ALFANUM"
                    Set class_temp = New Class_ALFA
                    coll_campos.Add class_temp, "campo AN-" + Format$(NroCampo, "00")
                    
                    ' [CAMPO01]
                    ' TIPO = ALFANUM ; Nombre
                    'class_temp.name = getitem(getoffile(archivo, seccion, key_item, ""), True)
                    key_item = "REM"
                    dataitem = getoffile(archivo, seccion, key_item, "")
                    class_temp.name = dataitem
                    
                    ' ATRIBUTOS=07
                    key_item = "ATRIBUTOS"
                    datadefault = "00"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.atributo = dataitem
                    
                    ' LRA=SIM|1:100
                    key_item = "LRA"
                    datadefault = "SIM|1:100"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.lra = dataitem
                    
                    ' LEN= 8
                    key_item = "LEN"
                    datadefault = "8"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.largo = Val(dataitem)
                    
                    ' POSXY=02 01
                    key_item = "POSXY"
                    datadefault = "1 1"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    texttemp = Split(dataitem, " ", 2)
                    class_temp.x_pos = Val(texttemp(0))
                    class_temp.y_pos = Val(texttemp(1))
                    
                    ' EDIT=ON
                    key_item = "EDIT"
                    datadefault = "OFF"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    If dataitem = "ON" Then
                        class_temp.Edit = True
                    Else
                        class_temp.Edit = False
                    End If
                    
                    'TRIGGER_ENABLE = ON/OFF
                    key_item = "TRIGGER_ENABLE"
                    datadefault = "OFF"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    If dataitem = "ON" Then
                        class_temp.TRIGGER_ENABLE = True
                    Else
                        class_temp.TRIGGER_ENABLE = False
                    End If
                    
                    'TRIGGER=SIM|1:100.0 0
                    key_item = "TRIGGER_LRA"
                    datadefault = "SIM|1:100.0 0"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.TRIGGER = dataitem
                
                Case "MTDIGITAL"
                    Set class_temp = New Class_MTD
                    coll_campos.Add class_temp, "campo TD-" + Format$(NroCampo, "00")
                    
                    
                    ' [CAMPO01]
                    ' TIPO = MTDIGITAL ; Nombre
                    'class_temp.name = getitem(getoffile(archivo, seccion, key_item, ""), True)
                    key_item = "REM"
                    dataitem = getoffile(archivo, seccion, key_item, "")
                    class_temp.name = dataitem
                    
                    ' ATRIBUTOS=07
                    key_item = "ATRIBUTOS"
                    datadefault = "00"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.atributo = dataitem
                    
                    ' LRA = SIM|1:100
                    key_item = "LRA"
                    datadefault = "SIM|1:100"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.lra = dataitem
                    
                    ' LEN = 8
                    key_item = "LEN"
                    datadefault = "8"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.largo = Val(dataitem)
                    
                    ' POSXY=02 01
                    key_item = "POSXY"
                    datadefault = "1 1"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    texttemp = Split(dataitem, " ", 2)
                    class_temp.x_pos = Val(texttemp(0))
                    class_temp.y_pos = Val(texttemp(1))
                    
                    ' EDIT = ON
                    key_item = "EDIT"
                    datadefault = "OFF"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    If dataitem = "ON" Then
                        class_temp.Edit = True
                    Else
                        class_temp.Edit = False
                    End If
                    
                    'TRIGGER_ENABLE = ON/OFF
                    key_item = "TRIGGER_ENABLE"
                    datadefault = "OFF"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    If dataitem = "ON" Then
                        class_temp.TRIGGER_ENABLE = True
                    Else
                        class_temp.TRIGGER_ENABLE = False
                    End If
                    
                    'TRIGGER=SIM|1:100.0 0
                    key_item = "TRIGGER_LRA"
                    datadefault = "SIM|1:100.0 0"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.TRIGGER = dataitem
                    
                    ' MT_ITEMS = 2
                    key_item = "MT_ITEMS"
                    datadefault = "2"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.items = Val(dataitem)
                    
                    ' MTDATA00 = Texto00
                    For i = 0 To (Val(class_temp.items) - 1)
                        key_item = "MTDATA" + Format$(i, "00")
                        datadefault = "MTD " + Format$(i, "00")
                        dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                        class_temp.texto(i) = dataitem
                    Next i
                    
                
                Case "NUMERICO"
                    Set class_temp = New Class_NUM
                    coll_campos.Add class_temp, "campo NU-" + Format$(NroCampo, "00")
                    
                    ' [CAMPO01]
                    ' TIPO = NUMERICO ; Nombre
                    'class_temp.name = getitem(getoffile(archivo, seccion, key_item, ""), True)
                    key_item = "REM"
                    dataitem = getoffile(archivo, seccion, key_item, "")
                    class_temp.name = dataitem
                    
                    ' ATRIBUTOS=07
                    key_item = "ATRIBUTOS"
                    datadefault = "00"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.atributo = dataitem
                    
                    ' NUMERIC_MODE = WS1
                    key_item = "NUMERIC_MODE"
                    datadefault = "WS1"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.modo = dataitem
                    
                    ' LRA = SIM|1:100
                    key_item = "LRA"
                    datadefault = "SIM|1:100"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.lra = dataitem
                    
                    ' DEC = 2
                    key_item = "DEC"
                    datadefault = "2"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.dec = Val(dataitem)
                    
                    ' LEN = 5
                    key_item = "LEN"
                    datadefault = "5"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.largo = Val(dataitem)
                    
                    ' EDIT = ON
                    key_item = "EDIT"
                    datadefault = "OFF"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    If dataitem = "ON" Then
                        class_temp.Edit = True
                    Else
                        class_temp.Edit = False
                    End If
                    
                    'TRIGGER_ENABLE = ON/OFF
                    key_item = "TRIGGER_ENABLE"
                    datadefault = "OFF"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    If dataitem = "ON" Then
                        class_temp.TRIGGER_ENABLE = True
                    Else
                        class_temp.TRIGGER_ENABLE = False
                    End If
                    
                    'TRIGGER=SIM|1:100.0 0
                    key_item = "TRIGGER_LRA"
                    datadefault = "SIM|1:100.0 0"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.TRIGGER = dataitem
                                       
                    
                    ' POSXY=02 01
                    key_item = "POSXY"
                    datadefault = "1 1"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    texttemp = Split(dataitem, " ", 2)
                    class_temp.x_pos = Val(texttemp(0))
                    class_temp.y_pos = Val(texttemp(1))
                                        
                    ' ESCALA=ON (Aun no implementado)
                    class_temp.usa_escala = True

                    ' GAIN_EXP=OFF
                    key_item = "GAIN_EXP"
                    datadefault = "OFF"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    ' Ojo que funciona al reves si GAIN_EXP = ON  recta = OFF y biceversa
                    If dataitem = "ON" Then
                        class_temp.recta = True
                    Else
                        class_temp.recta = False
                    End If
                    
                    ' GAIN=1
                    key_item = "GAIN"
                    datadefault = "1"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.gain = Val(dataitem)
                    
                    ' OFFSET=0
                    key_item = "OFFSET"
                    datadefault = "0"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.offset = Val(dataitem)
                    
                    ' RANGOX=2048 0
                    key_item = "RANGOX"
                    datadefault = "2048 0"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    texttemp = Split(dataitem, " ", 2)
                    class_temp.rangoX(1) = Val(texttemp(0))
                    class_temp.rangoX(0) = Val(texttemp(1))
                    
                    ' RANGOY=2048 0
                    key_item = "RANGOY"
                    datadefault = "2048 0"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    texttemp = Split(dataitem, " ", 2)
                    class_temp.rangoY(1) = Val(texttemp(0))
                    class_temp.rangoY(0) = Val(texttemp(1))
                    
                    ' MIN = 0
                    key_item = "MIN"
                    datadefault = "0"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.minimo = Val(dataitem)
                               
                    ' MAX = 99.99
                    key_item = "MAX"
                    datadefault = "99.99"
                    dataitem = getitem(getoffile(archivo, seccion, key_item, datadefault))
                    class_temp.maximo = Val(dataitem)
                    
                Case Else
                    Campo_tipo = "NO_TYPE"
                    
            End Select
        End If
        NroCampo = NroCampo + 1
    Loop Until Campo_tipo = "NO_TYPE"
End Function


'***************************************************** WRITE ********************************************************

Function write_gobales(origen As Class_GLOBAL, archivo As String) As Boolean
    Dim dataitem As String
    Dim key_item As String
    Dim seccion As String
    Dim tecla As Byte
    Dim texttemp() As String
    Dim ret
    Dim aux As Integer
    
    If Dir(archivo) <> "" Then
        tecla = InStrRev(archivo, ".ini")
        dataitem = Mid(archivo, 1, tecla) + "bak"
        If Dir(dataitem) <> "" Then Kill (dataitem)
        Name archivo As dataitem
    End If
    '[COMPILADOR]
    'POSX-1=ON
     ret = puttofile(archivo, "COMPILADOR", "POSX-1", "ON")
    
    '[COMM]
    'PORT=1
    ret = puttofile(archivo, "COMM", "PORT", CStr(origen.puerto))
    'BAUD=38400
    ret = puttofile(archivo, "COMM", "BAUD", CStr(origen.velocidad))
    'RTSMODE= 0
    If origen.invertido = True Then
        aux = 1
    Else
        aux = 0
    End If
    ret = puttofile(archivo, "COMM", "RTSMODE", CStr(aux))
    'RTSON= 10
    ret = puttofile(archivo, "COMM", "RTSON", CStr(origen.rts_on))
    'RTSOFF=10
    ret = puttofile(archivo, "COMM", "RTSOFF", CStr(origen.rts_off))
    'REPLYDELAY = 20
    ret = puttofile(archivo, "COMM", "REPLYDELAY", CStr(origen.delay_resp))

     
    ' [SYNC_PLC]
    ' SYNC_LRA=SIM|1:100
    ret = puttofile(archivo, "SYNC_PLC", "SYNC_LRA", origen.synclra)
    
    ' SYNC_TIME=0
    ret = puttofile(archivo, "SYNC_PLC", "SYNC_TIME", CStr(origen.synctime))
    
    ' SYNC_PLC_ACTIVO=OFF
    If origen.syncactivo Then
        ret = puttofile(archivo, "SYNC_PLC", "SYNC_PLC_ACTIVO", "ON")
    Else
        ret = puttofile(archivo, "SYNC_PLC", "SYNC_PLC_ACTIVO", "OFF")
    End If
    
    ' [VARIABLES_DEFAULT]
    'PANTALLA_PRINCIPAL = 16
    If origen.Pant_principal_enable Then
        ret = puttofile(archivo, "VARIABLES_DEFAULT", "PANTALLA_PRINCIPAL", origen.Pant_principal)
    End If
    
    'PANTALLA_INICIAL = 0
    If origen.Pant_inicial_enable Then
        ret = puttofile(archivo, "VARIABLES_DEFAULT", "PANTALLA_INICIAL", origen.Pant_inicial)
    End If
    
    ' tdisplay = 100
    ret = puttofile(archivo, "VARIABLES_DEFAULT", "TDISPLAY", origen.tdisplay)
    
    ' TACTIVO=30 0
    ReDim texttemp(2) As String
    texttemp(0) = CStr(origen.timeout)
    texttemp(1) = CStr(origen.timeoutjmp)
    dataitem = Join(texttemp, " ")
    ret = puttofile(archivo, "VARIABLES_DEFAULT", "TACTIVO", dataitem)
    
    ' ESCJMP = -1
    ret = puttofile(archivo, "VARIABLES_DEFAULT", "ESCJMP", origen.escjmp)
    
    ' TEDIT=30 2
    ReDim texttemp(1) As String
    texttemp(0) = CStr(origen.timeoutedit)
    texttemp(1) = CStr(origen.timeautocursor)
    dataitem = Join(texttemp, " ")
    ret = puttofile(archivo, "VARIABLES_DEFAULT", "TEDIT", dataitem)
    
    ' JMP_TECLA_x=x  Y BITSET_TECLA_X=SIM|1:100.1 0
    For tecla = 0 To 9
        'JUMP
        If origen.keyjmpenable(tecla) Then
            key_item = "JMP_TECLA_" + Mid(Str(tecla), 2, 1)
            dataitem = CStr(origen.keyjmp(tecla))
            ret = puttofile(archivo, "VARIABLES_DEFAULT", key_item, dataitem)
        End If
    
        'BIT SET/CLEAR
        If origen.Key_LRA(tecla) <> "OFF" Then
            key_item = "BITSET_TECLA_" + Mid(Str(tecla), 2, 1)
            dataitem = origen.Key_LRA(tecla)
            ret = puttofile(archivo, "VARIABLES_DEFAULT", key_item, dataitem)
        End If
    Next tecla
    
    For tecla = 1 To 4
        'JUMP
        If origen.keyjmpenable(tecla + 9) Then
            key_item = "JMP_TECLA_F" + Mid(Str(tecla), 2, 1)
            dataitem = CStr(origen.keyjmp(tecla + 9))
            ret = puttofile(archivo, "VARIABLES_DEFAULT", key_item, dataitem)
        End If
    
        'BIT SET/CLEAR
        If origen.Key_LRA(tecla + 9) <> "OFF" Then
            key_item = "BITSET_TECLA_F" + Mid(Str(tecla), 2, 1)
            dataitem = origen.Key_LRA(tecla + 9)
            ret = puttofile(archivo, "VARIABLES_DEFAULT", key_item, dataitem)
        End If
    Next tecla
    
    separa_seccion (archivo)

End Function

Function write_pantalla(origen As Class_Pantalla, archivo As String) As Boolean
    Dim dataitem As String
    Dim key_item As String
    Dim seccion As String
    Dim tecla As Byte
    Dim texttemp() As String
    Dim ret
    
    If Dir(archivo) <> "" Then
        tecla = InStrRev(archivo, ".ini")
        dataitem = Mid(archivo, 1, tecla) + "bak"
        If Dir(dataitem) <> "" Then Kill (dataitem)
        Name archivo As dataitem
    End If
    
    ' [VARIABLES]
    ' NUMBER=0
    ret = puttofile(archivo, "VARIABLES", "NUMBER", CStr(origen.Numero))
    
    ' NAME=Nombre
    ret = puttofile(archivo, "VARIABLES", "NAME", origen.name)
    
    ' tdisplay = 10
    If origen.tdisplay_local Then
        ret = puttofile(archivo, "VARIABLES", "TDISPLAY", origen.tdisplay)
    End If
    
        
    ' TACTIVO=30 0
    If origen.timeout_local Then
        ReDim texttemp(2) As String
        texttemp(0) = CStr(origen.timeout)
        texttemp(1) = CStr(origen.timeoutjmp)
        dataitem = Join(texttemp, " ")
        ret = puttofile(archivo, "VARIABLES", "TACTIVO", dataitem)
    End If
    
    ' ESCJMP
    If origen.escjmp_local Then
        ret = puttofile(archivo, "VARIABLES", "ESCJMP", origen.escjmp)
    End If
    
    
    ' TEDIT=30 2
    If origen.timeedit_local Then
        ReDim texttemp(1) As String
        texttemp(0) = CStr(origen.timeoutedit)
        texttemp(1) = CStr(origen.timeautocursor)
        dataitem = Join(texttemp, " ")
        ret = puttofile(archivo, "VARIABLES", "TEDIT", dataitem)
    End If
    
    ' JMP_TECLA_x=x Y BITSET_TECLA_X=SIM|1:100.1 0
    For tecla = 0 To 9
        'JUMP
        If origen.keyjmpenable(tecla) Then
            key_item = "JMP_TECLA_" + Mid(Str(tecla), 2, 1)
            dataitem = CStr(origen.keyjmp(tecla))
            ret = puttofile(archivo, "VARIABLES", key_item, dataitem)
        End If
        
        'BIT SET/CLEAR
        If origen.Key_LRA(tecla) <> "OFF" Then
            key_item = "BITSET_TECLA_" + Mid(Str(tecla), 2, 1)
            dataitem = origen.Key_LRA(tecla)
            ret = puttofile(archivo, "VARIABLES", key_item, dataitem)
        End If
    Next tecla
    For tecla = 1 To 4
        'JUMP
        If origen.keyjmpenable(tecla + 9) Then
            key_item = "JMP_TECLA_F" + Mid(Str(tecla), 2, 1)
            dataitem = CStr(origen.keyjmp(tecla + 9))
            ret = puttofile(archivo, "VARIABLES", key_item, dataitem)
        End If
    
        'BIT SET/CLEAR
        If origen.Key_LRA(tecla + 9) <> "OFF" Then
            key_item = "BITSET_TECLA_F" + Mid(Str(tecla), 2, 1)
            dataitem = origen.Key_LRA(tecla + 9)
            ret = puttofile(archivo, "VARIABLES", key_item, dataitem)
        End If

    Next tecla
    
    write_pantalla = write_Campos(origen.colectcampo, archivo)
    
    separa_seccion (archivo)

End Function


Function write_Campos(coll_campos As Collection, archivo As String) As Boolean
    Dim dataitem As String
    Dim datadefault As String
    Dim key_item As String
    Dim seccion As String
    Dim Campo_tipo As String
    Dim texttemp() As String
    Dim NroCampo As Integer
    Dim class_temp As Variant
    Dim ret
    Dim i As Integer
    
    
    For NroCampo = 1 To coll_campos.count
        
        Set class_temp = coll_campos.item(NroCampo)
        seccion = "CAMPO" + Format$(NroCampo, "00")
        
            Select Case class_temp.tipo_campo
                Case "CTEXT"
                    ' [CAMPO01]
                    ' TIPO = CTEXT
                    dataitem = class_temp.tipo_campo
                    key_item = "TIPO"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    'REM= comenterio o nombre del campo
                    If class_temp.name <> "" Then
                        dataitem = class_temp.name
                        key_item = "REM"
                        ret = puttofile(archivo, seccion, key_item, dataitem)
                    End If
                    
                    ' POSXY=02 01
                    ReDim texttemp(1) As String
                    texttemp(0) = CStr(class_temp.x_pos)
                    texttemp(1) = CStr(class_temp.y_pos)
                    dataitem = Join(texttemp, " ")
                    key_item = "POSXY"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' TEXTO=MENU PRINCIPAL 16/3
                    dataitem = class_temp.texto
                    key_item = "TEXTO"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                       
                    ' ATRIBUTOS=07
                    dataitem = class_temp.atributo
                    key_item = "ATRIBUTOS"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                Case "MTEXT"
                    
                    ' [CAMPO01]
                    ' TIPO = MTEXT
                    dataitem = class_temp.tipo_campo
                    key_item = "TIPO"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    'REM= comenterio o nombre del campo
                    If class_temp.name <> "" Then
                        dataitem = class_temp.name
                        key_item = "REM"
                        ret = puttofile(archivo, seccion, key_item, dataitem)
                    End If
                    
                    ' POSXY=02 01
                    ReDim texttemp(1) As String
                    texttemp(0) = CStr(class_temp.x_pos)
                    texttemp(1) = CStr(class_temp.y_pos)
                    dataitem = Join(texttemp, " ")
                    key_item = "POSXY"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' ATRIBUTOS=07
                    dataitem = class_temp.atributo
                    key_item = "ATRIBUTOS"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' TEXTO=MENU PRINCIPAL 16/3
                    dataitem = class_temp.texto
                    key_item = "TEXTO"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' JMP=01
                    dataitem = CStr(class_temp.jmp)
                    key_item = "JMP"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                
                Case "ALFANUM"
                    
                    ' [CAMPO01]
                    ' TIPO = ALFANUM
                    dataitem = class_temp.tipo_campo
                    key_item = "TIPO"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    'REM= comenterio o nombre del campo
                    If class_temp.name <> "" Then
                        dataitem = class_temp.name
                        key_item = "REM"
                        ret = puttofile(archivo, seccion, key_item, dataitem)
                    End If
                    
                    ' POSXY=02 01
                    ReDim texttemp(1) As String
                    texttemp(0) = CStr(class_temp.x_pos)
                    texttemp(1) = CStr(class_temp.y_pos)
                    dataitem = Join(texttemp, " ")
                    key_item = "POSXY"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' ATRIBUTOS=07
                    dataitem = class_temp.atributo
                    key_item = "ATRIBUTOS"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' LRA=SIM|1:100
                    dataitem = class_temp.lra
                    key_item = "LRA"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                     ' LEN= 8
                    dataitem = CStr(class_temp.largo)
                    key_item = "LEN"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' EDIT=ON
                    If class_temp.Edit Then
                        dataitem = "ON"
                    Else
                        dataitem = "OFF"
                    End If
                    key_item = "EDIT"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    'TRIGGER_ENABLE
                    If class_temp.TRIGGER_ENABLE Then
                        dataitem = "ON"
                    Else
                        dataitem = "OFF"
                    End If
                    key_item = "TRIGGER_ENABLE"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    'TRIGGER_LRA
                        dataitem = CStr(class_temp.TRIGGER)
                        key_item = "TRIGGER_LRA"
                        ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                Case "MTDIGITAL"
                    ' [CAMPO01]
                    ' TIPO = MTDIGITAL
                    dataitem = class_temp.tipo_campo
                    key_item = "TIPO"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    'REM= comenterio o nombre del campo
                    If class_temp.name <> "" Then
                        dataitem = class_temp.name
                        key_item = "REM"
                        ret = puttofile(archivo, seccion, key_item, dataitem)
                    End If
                    
                    ' POSXY=02 01
                    ReDim texttemp(1) As String
                    texttemp(0) = CStr(class_temp.x_pos)
                    texttemp(1) = CStr(class_temp.y_pos)
                    dataitem = Join(texttemp, " ")
                    key_item = "POSXY"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' ATRIBUTOS=07
                    dataitem = class_temp.atributo
                    key_item = "ATRIBUTOS"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' LRA=SIM|1:100
                    dataitem = class_temp.lra
                    key_item = "LRA"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                     ' LEN= 8
                    dataitem = CStr(class_temp.largo)
                    key_item = "LEN"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' EDIT=ON
                    If class_temp.Edit Then
                        dataitem = "ON"
                    Else
                        dataitem = "OFF"
                    End If
                    key_item = "EDIT"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    'TRIGGER_ENABLE
                    If class_temp.TRIGGER_ENABLE Then
                        dataitem = "ON"
                    Else
                        dataitem = "OFF"
                    End If
                    key_item = "TRIGGER_ENABLE"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    'TRIGGER_LRA
                        dataitem = CStr(class_temp.TRIGGER)
                        key_item = "TRIGGER_LRA"
                        ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' MT_ITEMS = 2
                    dataitem = CStr(class_temp.items)
                    key_item = "MT_ITEMS"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    
                    ' MTDATA00 = Texto00
                    For i = 0 To (Val(class_temp.items) - 1)
                        dataitem = class_temp.texto(i)
                        key_item = "MTDATA" + Format$(i, "00")
                        ret = puttofile(archivo, seccion, key_item, dataitem)
                    Next i
                    
                
                Case "NUMERICO"
                    ' [CAMPO01]
                    ' TIPO = NUMERICO
                    dataitem = class_temp.tipo_campo
                    key_item = "TIPO"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    'REM= comenterio o nombre del campo
                    If class_temp.name <> "" Then
                        dataitem = class_temp.name
                        key_item = "REM"
                        ret = puttofile(archivo, seccion, key_item, dataitem)
                    End If
                    
                    ' POSXY=02 01
                    ReDim texttemp(1) As String
                    texttemp(0) = CStr(class_temp.x_pos)
                    texttemp(1) = CStr(class_temp.y_pos)
                    dataitem = Join(texttemp, " ")
                    key_item = "POSXY"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' ATRIBUTOS=07
                    dataitem = class_temp.atributo
                    key_item = "ATRIBUTOS"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' LRA=SIM|1:100
                    dataitem = class_temp.lra
                    key_item = "LRA"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                     ' LEN= 8
                    dataitem = CStr(class_temp.largo)
                    key_item = "LEN"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' EDIT=ON
                    If class_temp.Edit Then
                        dataitem = "ON"
                    Else
                        dataitem = "OFF"
                    End If
                    key_item = "EDIT"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    'TRIGGER_ENABLE
                    If class_temp.TRIGGER_ENABLE Then
                        dataitem = "ON"
                    Else
                        dataitem = "OFF"
                    End If
                    key_item = "TRIGGER_ENABLE"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    'TRIGGER_LRA
                        dataitem = CStr(class_temp.TRIGGER)
                        key_item = "TRIGGER_LRA"
                        ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' DEC = 2
                    dataitem = CStr(class_temp.dec)
                    key_item = "DEC"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' NUMERIC_MODE = WS1
                    dataitem = class_temp.modo
                    key_item = "NUMERIC_MODE"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                                        
                    ' ESCALA=ON (Aun no implementado)
                    class_temp.usa_escala = True

                    ' GAIN_EXP=OFF
                    ' Ojo que funciona al reves si GAIN_EXP = ON  recta = OFF y biceversa
                    If class_temp.recta Then
                        dataitem = "ON"
                    Else
                        dataitem = "OFF"
                    End If
                    key_item = "GAIN_EXP"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' GAIN=1
                    dataitem = CStr(class_temp.gain)
                    key_item = "GAIN"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' OFFSET=0
                    dataitem = CStr(class_temp.offset)
                    key_item = "OFFSET"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' RANGOX=2048 0
                    ReDim texttemp(1) As String
                    texttemp(0) = CStr(class_temp.rangoX(1))
                    texttemp(1) = CStr(class_temp.rangoX(0))
                    dataitem = Join(texttemp, " ")
                    key_item = "RANGOX"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' RANGOY=2048 0
                    ReDim texttemp(1) As String
                    texttemp(0) = CStr(class_temp.rangoY(1))
                    texttemp(1) = CStr(class_temp.rangoY(0))
                    dataitem = Join(texttemp, " ")
                    key_item = "RANGOY"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' MIN = 0
                    dataitem = CStr(class_temp.minimo)
                    key_item = "MIN"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
                    ' MAX = 99.99
                    dataitem = CStr(class_temp.maximo)
                    key_item = "MAX"
                    ret = puttofile(archivo, seccion, key_item, dataitem)
                    
            End Select
    Next NroCampo
End Function


Private Sub separa_seccion(ByVal archivo_ini As String)
    Dim aux_ As String * 1
    Dim aux_cr As String
    Dim count As Byte
    
    If archivo_ini <> "" Then
        Open archivo_ini For Binary Access Read Write As #1
        Open archivo_ini & "2" For Binary Access Read Write As #2
        
        Get #1, , aux_
        count = 1
        Do While Not EOF(1)
                If aux_ = "[" Then
                    If count > 1 Then
                        aux_cr = Chr(13) & Chr(10)
                        Put #2, , aux_cr
                        Put #2, , aux_
                    Else
                        Put #2, , aux_
                    End If
                    count = 2
                Else
                    Put #2, , aux_
                End If
            Get #1, , aux_
        Loop
    End If
    Close #1
    Close #2
    Kill (archivo_ini)
    Name (archivo_ini & "2") As archivo_ini
End Sub

