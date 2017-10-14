VERSION 5.00
Begin VB.Form Form0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Por favor elija el tipo de inicio"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "USER.frx":0000
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Iniciar"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipos de Sesion"
      Height          =   1695
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3495
      Begin VB.OptionButton Option3 
         Caption         =   "Administrador"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2775
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Super Usuario"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   3015
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Usuario"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4080
      Top             =   120
   End
End
Attribute VB_Name = "Form0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Option1.Value Then
    Form0.Hide
    Form1.Show
ElseIf Option2.Value Then
    If Setup!SU = Text1.Text Then
        Form0.Hide
        Form1.Show
    Else
        MsgBox "La contraseña del Super Usuario no es correcta", vbExclamation
    End If
ElseIf Option3.Value Then
    If Setup!AD = Text1.Text Then
        Form0.Hide
        Form1.Show
    Else
        MsgBox "La contraseña del Administrador no es correcta", vbExclamation
    End If
End If
Text1.Text = ""
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Activate()
On Error GoTo Alternativa
Set Inventa = OpenDatabase("J:\EACPsys.mdb")
'Set Inventa = OpenDatabase("J:\Pruebas.mdb")
Form1.Image5.Visible = False
Alternativa:
If Error <> "" Then
    Set Inventa = OpenDatabase("C:\Users\Erick M\Desktop\Programa claudio\Codigo Fuente\EACPsys.mdb")
    Form1.Image5.Visible = True
End If
Set Setup = Inventa.OpenRecordset("SETUP")
Setup.MoveFirst
End Sub

Private Sub Option1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command1_Click
End If
If KeyAscii = 18 Then
    Form0.Hide
    Form1.Show
    Form0.Option1.Value = False
    Form0.Option2.Value = False
    Form0.Option3.Value = True
End If
End Sub

Private Sub Option2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1.SetFocus
End If
End Sub

Private Sub Option3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1.SetFocus
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub

Private Sub Timer1_Timer()
If Option1.Value Then
    If Text1.Visible Then
        Text1.Visible = False
    End If
Else
    If Not Text1.Visible Then
        Text1.Visible = True
    End If
End If
End Sub
