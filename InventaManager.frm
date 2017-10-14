VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Inventario"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "InventaManager.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      MaxLength       =   12
      TabIndex        =   10
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   9
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   26
      Top             =   4200
      Width           =   5655
      Begin VB.CommandButton Command3 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   2880
         TabIndex        =   12
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Equivalente"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   25
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "Precio:"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2760
      TabIndex        =   23
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label12 
      Caption         =   "Ubicación 2"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Ubicación 1"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Stock:"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Por venta:"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Existente:"
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4080
      TabIndex        =   17
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Precio al costo:"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Proveedor:"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Descripción:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Código:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Calcular()
If IsNumeric(Text4.Text) And IsNumeric(Text5.Text) Then
    If Text4.Text > 0 And Text5.Text > 0 Then
        Label15.Caption = FormatCurrency(Text4.Text + ((Text4.Text * Text5.Text) / 100), 2)
    End If
End If
End Sub
Private Sub Limpiar()
Combo1.Text = ""
Label6.Caption = ""
Label13.Caption = ""
Label15.Caption = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
'Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
End Sub
Private Sub Nuevo()
Call Limpiar
End Sub
Private Sub Editar()
If Not IsNull(Tabla!Prove) Then
    Combo1.Text = Tabla!Prove
Else
    Combo1.Text = 0
End If
Text1.Text = Tabla!Codigo
Text2.Text = Tabla!descrip
Text4.Text = Tabla!P_unit
Text5.Text = Tabla!Por_Venta
Text3.Text = Tabla!Equivalente
Label6.Caption = Tabla!Cantidad
Text7.Text = Tabla!Stock
If IsNull(Tabla!gabeta) Then
    Text9.Text = ""
Else
    Text9.Text = Tabla!gabeta
End If
If IsNull(Tabla!ubicacion) Then
    Text10.Text = ""
Else
    Text10.Text = Tabla!ubicacion
End If
Call Combo1_Click
End Sub

Private Sub Combo1_Click()
Provedores.Seek "=", Combo1.Text
If Not Provedores.NoMatch Then
    Label13.Caption = Provedores!Nombre
Else
    Label13.Caption = ""
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Combo1.Text = "" Then
        Combo1.SetFocus
    Else
        Combo1_Click
        Text4.SetFocus
    End If
ElseIf KeyAscii = 27 Then
    Text1.SetFocus
End If
Dim KEY As String
KEY = Chr(KeyAscii)
If (KEY < "0" Or KEY > "9") Then
    If (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Command1_Click()
Set ECGList = Inventa.OpenRecordset("ECGLIST")
ECGList.Index = "PrimaryKey"
ECGList.Seek "=", Text3
If Len(Text4.Text) = 0 Then
    Text4.SetFocus
ElseIf Len(Text5.Text) = 0 Then
    Text5.SetFocus
ElseIf Len(Text7.Text) = 0 Then
    Text7.Text = 0
'ElseIf Len(Text8.Text) = 0 Then
'    Text8.SetFocus
'ElseIf Text4.Text < 0 Then
'    Text4.SetFocus
ElseIf Text5.Text < 0 Then
    Text5.SetFocus
'ElseIf Text6.Text < 0 Then
'    Text6.SetFocus
'ElseIf Text7.Text < 0 Then
'    Text7.SetFocus
ElseIf ECGList.NoMatch Then
    MsgBox "La equivalencia no existe en la lista del sistema", vbCritical
    Text3.SetFocus
ElseIf Len(Label13.Caption) = 0 Then
    MsgBox "Por Favor Digite un Proveedor Valido", vbInformation
    Combo1.SetFocus
ElseIf Combo1.Text < 0 Then
    Combo1.SetFocus
Else
    Tabla.Seek "=", Text1.Text
    If Tabla.NoMatch Then
        MsgBox "Codigo a editar no existe", vbCritical
    Else
        Tabla.Edit
    End If
    Tabla!Prove = Combo1.Text
    Tabla!descrip = Text2.Text
    Tabla!P_unit = Text4.Text
    Tabla!Por_Venta = Text5.Text
    Tabla!P_Venta = Label15.Caption
    If Len(Text6.Text) <> 0 Then
        Tabla!Cantidad = Text6.Text
    End If
    Tabla!Stock = Text7.Text
    Tabla!gabeta = Text9.Text
    Tabla!ubicacion = Text10.Text
    Tabla!Equivalente = UCase(Text3)
    Tabla.Update
    Call Limpiar
    Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
Tabla.Index = "Codigo"
Tabla.Seek "=", Text1.Text
If Tabla.NoMatch Then
    MsgBox "Codigo a borrar no existe", vbCritical
Else
    Tabla.Delete
    Text1.Text = ""
    Call Limpiar
End If
End Sub

Private Sub Command3_Click()
Text1.SetFocus
Form5.Hide
Form1.Show
End Sub

Private Sub Form_Activate()
Set Tabla = Inventa.OpenRecordset("INVENTA")
Set Provedores = Inventa.OpenRecordset("PROVEDORES")
Provedores.Index = "PrimaryKey"
Provedores.MoveFirst
Call Limpiar
While Not Provedores.EOF
    Combo1.AddItem (Provedores!Codigo)
    Provedores.MoveNext
Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Command3_Click
End Sub

Private Sub Text1_GotFocus()
Call Limpiar
Text1.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Tabla.Index = "Codigo"
    Tabla.Seek "=", Mid("0000000000", 1, 10 - Len(Text1.Text)) & Text1.Text
    If Tabla.NoMatch Then
        MsgBox "Codigo a editar no existe", vbCritical
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
        Text1.SetFocus
    Else
        Call Editar
        Text2.SetFocus
    End If
ElseIf KeyAscii = 27 Then
    Call Command3_Click
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsEmpty(Text2.Text) Then
        Text2.SetFocus
    Else
        Combo1.SetFocus
    End If
ElseIf KeyAscii = 27 Then
    Text1.SetFocus
End If
End Sub

Private Sub Text4_Change()
Call Calcular
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsNumeric(Text4.Text) Then
        Text5.SetFocus
    Else
        Text4.SetFocus
    End If
ElseIf KeyAscii = 27 Then
    Text1.SetFocus
End If
Dim KEY As String
KEY = Chr(KeyAscii)
If (KEY < "0" Or KEY > "9") And KEY <> "," Then
    If (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text5_Change()
Call Calcular
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsNumeric(Text5.Text) Then
        Text6.SetFocus
        Call Calcular
    Else
        Text5.SetFocus
    End If
ElseIf KeyAscii = 27 Then
    Text1.SetFocus
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
    If IsNumeric(Text6.Text) Or Len(Text6.Text) = 0 Then
        Text7.SetFocus
    End If
ElseIf KeyAscii = 27 Then
    Text1.SetFocus
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
If KeyAscii = 13 Then
    If IsNumeric(Text7.Text) Then
        Text9.SetFocus
    Else
        Text7.SetFocus
    End If
ElseIf KeyAscii = 27 Then
    Text1.SetFocus
End If
Dim KEY As String
KEY = Chr(KeyAscii)
If (KEY < "0" Or KEY > "9") Then
    If (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsNumeric(Text8.Text) Then
        Text9.SetFocus
    Else
        Text8.SetFocus
    End If
ElseIf KeyAscii = 27 Then
    Text1.SetFocus
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text10.SetFocus
ElseIf KeyAscii = 27 Then
    Text1.SetFocus
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1.SetFocus
ElseIf KeyAscii = 27 Then
    Text1.SetFocus
End If
End Sub
