VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de proveedores"
   ClientHeight    =   3990
   ClientLeft      =   4080
   ClientTop       =   1305
   ClientWidth     =   7110
   Icon            =   "ProvedorManager.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1800
      TabIndex        =   19
      Top             =   2280
      Width           =   4215
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   17
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   6855
      Begin VB.CommandButton Command4 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   4800
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Actualizar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   4680
      MaxLength       =   2
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3240
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label7 
      Caption         =   "e-mail"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Codigo Nuevo"
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1800
      TabIndex        =   13
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Incoporado el"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Tel1 / Tel2 / Fax"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Dirección"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Limpiar()
Provedores.Index = "PrimaryKey"
Provedores.MoveLast
Text8.Text = Provedores!Codigo + 1
Combo1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Label13.Caption = ""
Command1.Enabled = False
Command2.Enabled = False
Text1.SetFocus
End Sub

Private Sub Nuevo()
Command1.Enabled = True
Command2.Enabled = False
End Sub

Private Sub Buscar()
Provedores.Index = "Provedor"
Provedores.Seek "=", Combo1.Text
If Not Provedores.NoMatch Then
    TMP = Provedores!Codigo
    Text1.Text = TMP
    If Not IsNull(Provedores!Direccion) Then
        Text2.Text = Provedores!Direccion
    Else
        Text2.Text = ""
    End If
    If Not IsNull(Provedores!Tel1) Then
        Text3.Text = Provedores!Tel1
    Else
        Text3.Text = ""
    End If
    If Not IsNull(Provedores!Tel2) Then
        Text4.Text = Provedores!Tel2
    Else
        Text4.Text = ""
    End If
    If Not IsNull(Provedores!Fax) Then
        Text5.Text = Provedores!Fax
    Else
        Text5.Text = ""
    End If
    If Not IsNull(Provedores!email) Then
        Text6.Text = Provedores!email
    Else
        Text6.Text = ""
    End If
    If Not IsNull(Provedores!FEC_ING) Then
        Label13.Caption = Provedores!FEC_ING
    Else
        Label13.Caption = ""
    End If
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub Command1_Click()
If Len(Text1.Text) = 0 Then
    MsgBox "Digite el ID del Vendedor"
    Text1.SelStart = 0
    Text1.SelLength = Len(Combo1.Text)
    Text1.SetFocus
ElseIf Len(Combo1.Text) = 0 Then
    MsgBox "Digite el nombre del Vendedor"
    Combo1.SelStart = 0
    Combo1.SelLength = Len(Combo1.Text)
    Combo1.SetFocus
ElseIf Len(Text2.Text) = 0 Then
    MsgBox "Digite la direccion del Vendedor"
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    Text2.SetFocus
ElseIf Len(Text3.Text) = 0 And Len(Text3.Text) = 0 And Len(Text3.Text) = 0 Then
    MsgBox "Digite el numero de telefono del Vendedor"
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.SetFocus
Else
    Provedores.AddNew
    Provedores!Codigo = Text1.Text
    Provedores!Nombre = Combo1.Text
    Provedores!Direccion = Text2.Text
    Provedores!Tel1 = Text3.Text
    Provedores!Tel2 = Text4.Text
    Provedores!Fax = Text5.Text
    Provedores!email = Text6.Text
    Provedores!FEC_ING = Date
    Provedores.Update
    Text1.SetFocus
    Text1.Text = ""
    Call Limpiar
End If
End Sub

Private Sub Command2_Click()
If Len(Text1.Text) = 0 Then
    MsgBox "Digite el ID del Vendedor"
    Text1.SelStart = 0
    Text1.SelLength = Len(Combo1.Text)
    Text1.SetFocus
ElseIf Len(Combo1.Text) = 0 Then
    MsgBox "Digite el nombre del Vendedor"
    Combo1.SelStart = 0
    Combo1.SelLength = Len(Combo1.Text)
    Combo1.SetFocus
ElseIf Len(Text2.Text) = 0 Then
    MsgBox "Digite la direccion del Vendedor"
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    Text2.SetFocus
ElseIf Len(Text3.Text) = 0 And Len(Text3.Text) = 0 And Len(Text3.Text) = 0 Then
    MsgBox "Digite el numero de telefono del Vendedor"
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.SetFocus
Else
    Provedores.Edit
    Provedores!Codigo = Text1.Text
    Provedores!Nombre = Combo1.Text
    Provedores!Direccion = Text2.Text
    Provedores!Tel1 = Text3.Text
    Provedores!Tel2 = Text4.Text
    Provedores!Fax = Text5.Text
    Provedores!email = Text6.Text
    Provedores.Update
    Text1.SetFocus
    Text1.Text = ""
    Call Limpiar
End If
End Sub

Private Sub Command4_Click()
Form6.Hide
Form1.Show
End Sub

Private Sub Form_Activate()
Text1.Text = ""
Command1.Enabled = False
Command2.Enabled = False
Set Provedores = Inventa.OpenRecordset("Provedores")
Call Limpiar
Provedores.Index = "Provedor"
Provedores.MoveFirst
While Provedores.EOF = False
    Combo1.AddItem Provedores!Nombre
    Provedores.MoveNext
Wend
End Sub

Private Sub Form_Deactivate()
Combo1.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Command4_Click
End Sub

Private Sub Combo1_Change()
Call Buscar
End Sub

Private Sub Combo1_Click()
Call Buscar
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Text1.Text) <> 0 And IsNumeric(Text1.Text) Then
        Provedores.Index = "PrimaryKey"
        Provedores.Seek "=", Text1.Text
        If Provedores.NoMatch Then
            Text1.Text = Text8.Text
        End If
        If Text1.Text <> Text8.Text Then
            If Not Provedores.NoMatch Then
                If Form0.Option3.Value = True Then
                    Command2.Enabled = True
                    Command1.Enabled = False
                End If
                Combo1.Text = Provedores!Nombre
            End If
        Else
            Call Limpiar
            Call Nuevo
        End If
    End If
    Combo1.SetFocus
End If
Dim KEY As String
KEY = Chr(KeyAscii)
If (KEY < "0" Or KEY > "9") Then
    If (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    Text4.SetFocus
End If
Dim KEY As String
KEY = Chr(KeyAscii)
If (KEY < "0" Or KEY > "9") Then
    If (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    Text5.SetFocus
End If
Dim KEY As String
KEY = Chr(KeyAscii)
If (KEY < "0" Or KEY > "9") Then
    If (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    Text6.SetFocus
End If
Dim KEY As String
KEY = Chr(KeyAscii)
If (KEY < "0" Or KEY > "9") Then
    If (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If Command1.Enabled Then
        Command1.SetFocus
    Else
        Command2.SetFocus
    End If
End If
Dim KEY As String
KEY = Chr(KeyAscii)
If (KEY < "0" Or KEY > "9") Then
    If (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
End Sub
