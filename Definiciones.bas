Attribute VB_Name = "Definiciones"

'****************************************************************************************************************
'****************************************************************************************************************
'****************************************************************************************************************
Declare Function ReleaseCapture _
  Lib "user32" () As Long

Declare Function SendMessage Lib _
  "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  lParam As Any) As Long

Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
'****************************************************************************************************************
'****************************************************************************************************************
'****************************************************************************************************************
Global Const BLANCO = &HFFFFFF
Global Const GRIS = &H808080
Global Const NEGRO = &H0&
Global Const AZUL = &HFF0000
Global Const ROJO = &HFF&
Global Const VERDE = &HFF00&
Global Const AMARILLO = &HFFFF&
Global Const AZUL_SUAVE = &HFFFF00
Global Const ROJO_SUAVE = &H8080FF
Global Const VERDE_SUAVE = &H80FF80
Global Const VERDE_SUAVE2 = &HC0FFC0
Global Const AMARILLO_SUAVE = &H80FFFF
Global Const GRIS_SUAVE = &HC0C0C0
Global Const GRIS_MUY_SUAVE = &HE0E0E0
Global Const MARRON_SUAVE = &H40C0&
Global Const NARANJA_SUAVE = &H80C0FF
Global Const MARRON = &H4080&
Global Const NARANJA = &H80FF&
Global Const CARADEBOTON = &H8000000F
Global Const SEPARADOR = &H8000000C
'****************************************************************************************************************
'****************************************************************************************************************
'****************************************************************************************************************
 
Global Const CTE_SEP_COMENT = " ;"
 
Global Const MAX_COL = 20
Global Const MAX_FILA = 4
'****************************************************************************************************************
'****************************************************************************************************************
'****************************************************************************************************************

'******************************************************************************************************************
'variables de juan
'Public drag_campo As Byte
Public flag_cambio As Byte
Public max_foco As Byte
Public x_aux As Integer
Public y_aux As Integer
Public coll_pantalla As Collection
Public test_caption As String
Public pant As pantalla
Public cuanta_pant As Byte
Public Pantalla_activa As String
Public Const PASOV = 380
Public Const PASOH = 212
Public Const NPROP = 26
Public temp_cont_cop As Object
Public old_foco_campo As Object

'*******CLASES*********
Public class_temp  As Variant
Public cl_temp_pant As Variant

'constantes
Public Const NOMBRE_DEL_COMPILADOR = "CpOpm1.exe"
Public Const C_TEXT = 0
Public Const M_TEXT = 1
Public Const NUMERIC = 2
Public Const ALFANUM = 3
Public Const MTDIGITAL = 4
Public Const moviendo = 1
Public Const creando = 0
Public Const cambio = 1
Public Const sin_cambio = 0

'******************************************************************************************************************

Public Const FONDO_COLOR = &HC0FFC0
Public Const CTEXT_COLOR = &HC0C0C0
Public Const MTEXT_COLOR = &HC0E0FF
Public Const NUMERICO_COLOR = &HFFC0C0
Public Const MTDIGITAL_COLOR = &HFFC0FF
Public Const ALFANUM_COLOR = &H80FFFF

'----------------------------------------------------------------------------
' Windows 32 API Routines
'----------------------------------------------------------------------------
' declaraciones API windows para in/out config.ini
'----------------------------------------------------------------------------
Global Const GS_MAX = 132    ' largo de buffer para leer y escribir ini files

Declare Function WritePrivateProfileString& Lib _
"Kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal Lpsection$, ByVal Lpentry$, ByVal _
    buffer$, ByVal filename$)
Declare Function GetPrivateProfileString% Lib _
"Kernel32" Alias "GetPrivateProfileStringA" _
   (ByVal Lpsection$, ByVal Lpentry$, _
    ByVal lpDefault$, ByVal buffer$, ByVal _
    cbBuffer%, ByVal lpFileName$)

' Declaracion del Formato de un proyecto
Public Type Stts_Screens
    Abierta As Boolean
    Estado As Integer
End Type

Enum ModoScreen
    SC_LIBRE = 0
    SC_USADO
    SC_NUEVO
    SC_MODIFICADO
    SC_BORRAR
End Enum

Global Const MAX_LCD_X = 20
Global Const MAX_LCD_Y = 4



Public Proyectos As Collection
Public Proyectos_Open_New As Integer
Public tempscreen As Class_Pantalla


Function token$(tmp$, search$)
Dim X%
    X = InStr(1, tmp$, search$)
    If X Then
       token$ = Mid$(tmp$, 1, X - 1)
       tmp$ = Mid$(tmp$, X + 1)
    Else
       token$ = tmp$
       tmp$ = ""
    End If
End Function

'*****************************************************
Public Sub open_project(fileprj As String)
    Dim temp As V_Indice
       
    If CheckPrjOpen(fileprj) Then
        Set temp = Proyectos.item(fileprj)
        temp.SetFocus
    Else
        Set temp = New V_Indice
        temp.archivo = fileprj
        Proyectos.Add temp, fileprj
    End If
    
End Sub

Public Sub libera_proyecto(idprj As String)
    Proyectos.Remove idprj
    If Proyectos.count = 0 Then
        EditorIDE.PRJ_GUARDAR_COMO.Enabled = False
        EditorIDE.Toolbar.Buttons(3).Enabled = False
        EditorIDE.PRJ_GUARDAR.Enabled = False
        EditorIDE.List_lra.Enabled = False
        EditorIDE.PRJ_PROP.Enabled = False
        EditorIDE.Toolbar.Buttons(4).Enabled = False
    End If
End Sub

Public Function proyecto_guardar_como(idprj As String) As String
    
    Dim ruta As String
    On Error GoTo finError
  
    EditorIDE.ComunWindows.Filter = "Proyecto Panel Operador PO900(*.ppo )|*.ppo"
    EditorIDE.ComunWindows.Flags = cdlOFNOverwritePrompt
    EditorIDE.ComunWindows.CancelError = True
    EditorIDE.ComunWindows.ShowSave
    
    If EditorIDE.ComunWindows.filename <> "" Then
        file_path = EditorIDE.ComunWindows.filename
        file_name = EditorIDE.ComunWindows.FileTitle
        If file_path <> idprj Then
            aux = InStrRev(file_path, "\")
            file_name = Mid(file_path, aux + 1)
            file_path = Mid(file_path, 1, aux)
            aux = InStr(file_name, ".ppo")
            file_name = Mid(file_name, 1, aux - 1)
            
            If Dir(file_path + file_name + ".ppo") <> "" Then
                If CheckPrjOpen(file_path + file_name + ".ppo") Then
                    ruta = ""
                    MsgBox ("Se intento guardar un proyecto que sobre otro que se encuentra abierto")
                Else
                    ruta = file_path + file_name + ".ppo"
                End If
            Else
                ruta = file_path + file_name + "\"
                aux = 0
                While Dir(ruta) <> ""
                    ruta = file_path + file_name + " " + Str(aux) + "\"
                    aux = aux + 1
                Wend
            
                MkDir (ruta)
                MkDir (ruta + "SCRNS")
                ruta = ruta + file_name + ".ppo"
                rc% = WritePrivateProfileString("PROYECTO", "NOMBRE", "Vacio", ruta)
            End If
        Else
            ruta = idprj
        End If
        
    End If
    proyecto_guardar_como = ruta
    Exit Function
finError:
    If Err = &H7FF3 Then
        EditorIDE.ComunWindows.filename = ""
    End If
    Resume Next
    

End Function

Public Function CheckPrjOpen(idprj As String) As Boolean
    Dim temp As V_Indice
    CheckPrjOpen = False
    For i = 1 To Proyectos.count
        Set temp = Proyectos.item(i)
        If temp.archivo = idprj Then
            CheckPrjOpen = True
        End If
    Next i
End Function

Public Function dec(hexa As String) As Integer
    dec = CInt("&H" & hexa)
End Function

Public Function genidpan(ByVal indice As Integer) As String
    genidpan = ("PANTALLA_" + Format$(indice, "00"))
End Function

Public Function genidcampo(ByVal tipo As String, ByVal indice As Integer) As String
    Select Case tipo
        Case "CTEXT"
            genidcampo = "campo CT-" + Format$(indice, "00")
        Case "MTEXT"
            genidcampo = "campo MT-" + Format$(indice, "00")
        Case "MTDIGITAL"
            genidcampo = "campo TD-" + Format$(indice, "00")
        Case "ALFANUM"
            genidcampo = "campo AN-" + Format$(indice, "00")
        Case "NUMERICO"
            genidcampo = "campo NU-" + Format$(indice, "00")
    End Select
End Function
