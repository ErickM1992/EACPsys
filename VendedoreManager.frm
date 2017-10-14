VERSION 5.00
Begin VB.Form Form15 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de vendedores"
   ClientHeight    =   3615
   ClientLeft      =   4080
   ClientTop       =   1305
   ClientWidth     =   7110
   Icon            =   "VendedoreManager.frx":0000
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
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
      Top             =   2760
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
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Incoporado el"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
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
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Limpiar()
Vendedores.Index = "PrimaryKey"
Vendedores.MoveLast
Text8.Text = Vendedores!Codigo + 1
Combo1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Label13.Caption = ""
Command1.Enabled = False
Command2.Enabled = False
End Sub

Private Sub Nuevo()
Command1.Enabled = True
Command2.Enabled = False
End Sub

Private Sub Buscar()
Vendedores.Index = "Vendedor"
Vendedores.Seek "=", Combo1.Text
If Not Vendedores.NoMatch Then
    TMP = Vendedores!Codigo
    Text1.Text = TMP
    If Not IsNull(Vendedores!Direccion) Then
        Text2.Text = Vendedores!Direccion
    Else
        Text2.Text = ""
    End If
    If Not IsNull(Vendedores!Tel1) Then
        Text3.Text = Vendedores!Tel1
    Else
        Text3.Text = ""
    End If
    If Not IsNull(Vendedores!Tel2) Then
        Text4.Text = Vendedores!Tel2
    Else
        Text4.Text = ""
    End If
    If Not IsNull(Vendedores!Fax) Then
        Text5.Text = Vendedores!Fax
    Else
        Text5.Text = ""
    End If
    If Not IsNull(Vendedores!FEC_ING) Then
        Label13.Caption = Vendedores!FEC_ING
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
    Vendedores.AddNew
    Vendedores!Codigo = Text1.Text
    Vendedores!Nombre = Combo1.Text
    Vendedores!Direccion = Text2.Text
    Vendedores!Tel1 = Text3.Text
    Vendedores!Tel2 = Text4.Text
    Vendedores!Fax = Text5.Text
    Vendedores!FEC_ING = Date
    Vendedores.Update
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
    Vendedores.Edit
    Vendedores!Codigo = Text1.Text
    Vendedores!Nombre = Combo1.Text
    Vendedores!Direccion = Text2.Text
    Vendedores!Tel1 = Text3.Text
    Vendedores!Tel2 = Text4.Text
    Vendedores!Fax = Text5.Text
    Vendedores.Update
    Text1.SetFocus
    Text1.Text = ""
    Call Limpiar
End If
End Sub

Private Sub Command4_Click()
Form15.Hide
Form1.Show
End Sub

Private Sub Form_Activate()
Text1.Text = ""
Command1.Enabled = False
Command2.Enabled = False
Set Vendedores = Inventa.OpenRecordset("VENDEDORES")
Call Limpiar
Vendedores.Index = "vendedor"
Vendedores.MoveFirst
While Vendedores.EOF = False
    Combo1.AddItem Vendedores!Nombre
    Vendedores.MoveNext
Wend
Text1.SetFocus
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
        Vendedores.Index = "PrimaryKey"
        Vendedores.Seek "=", Text1.Text
        If Vendedores.NoMatch Then
            Text1.Text = Text8.Text
        End If
        If Text1.Text <> Text8.Text Then
            If Not Vendedores.NoMatch Then
                Command2.Enabled = True
                Command1.Enabled = False
                Combo1.Text = Vendedores!Nombre
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

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    Text7.SetFocus
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
