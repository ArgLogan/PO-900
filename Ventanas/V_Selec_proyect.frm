VERSION 5.00
Begin VB.Form V_Selec_proyect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecione Proyecto"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   8130
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List_proyectos 
      Height          =   1230
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   7845
   End
End
Attribute VB_Name = "V_Selec_proyect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IDstr As String

Public Function Load() As String
    Dim indice As Integer
    Dim temp As V_Indice
    For indice = 1 To Proyectos.count
        Set temp = Proyectos.item(indice)
        List_proyectos.AddItem (temp.Nombre + " (" + temp.archivo + ")")
    Next indice
    Me.Show (1)
    Load = IDstr
End Function

Private Sub List_proyectos_DblClick()
    Dim temp As V_Indice
    Set temp = Proyectos.item(List_proyectos.ListIndex + 1)
    IDstr = temp.archivo
    Unload Me
End Sub
