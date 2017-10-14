VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form16 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aumento"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Aumento.frx":0000
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   0
      TabIndex        =   4
      Top             =   1680
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aplicar"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Digite el aumento a aplicar en el inventario"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set Tabla = Inventa.OpenRecordset("Inventa")
Tabla.Index = "PrimaryKey"
Tabla.MoveFirst
Command1.Enabled = False
Command2.Enabled = False
While Not Tabla.EOF
    ProgressBar1.Value = Tabla.PercentPosition
    Tabla.Edit
    Tabla!P_unit = (Tabla!P_unit + (Tabla!P_unit * (Text1 / 100)))
    Tabla!P_Venta = (Tabla!P_unit + ((Tabla!P_unit * Tabla!Por_Venta) / 100))
    Tabla.Update
    Tabla.MoveNext
Wend
ProgressBar1.Value = 100
MsgBox "Aumento del " & Text1 & "% aplicado", vbInformation
Call Command2_Click
Command1.Enabled = True
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
Form16.Hide
Form1.Show
End Sub

Private Sub Form_Activate()
Text1 = ""
ProgressBar1.Value = 0
Command1.Enabled = False
Command2.Enabled = True
End Sub

Private Sub Text1_GotFocus()
If Command1.Enabled Then
    Command1.Enabled = False
    Text1.SelStart = 0
    Text1.SelLength = 3
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1 > 100 Then
        Text1 = 100
    End If
    MsgBox "El aumento a aplicar va a ser de un %" & Text1, vbInformation
    Command1.Enabled = True
    Command1.SetFocus
End If
Dim KEY As String
KEY = Chr(KeyAscii)
If (KEY < "0" Or KEY > "9") Then
    If (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End If
End Sub

