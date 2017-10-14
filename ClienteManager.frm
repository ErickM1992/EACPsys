VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de clientes"
   ClientHeight    =   4935
   ClientLeft      =   4080
   ClientTop       =   1305
   ClientWidth     =   7110
   Icon            =   "ClienteManager.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   25
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Text            =   "0"
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """¢""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   2
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Text            =   "0"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   6855
      Begin VB.CommandButton Command4 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   4800
         TabIndex        =   23
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Actualizar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   120
         TabIndex        =   9
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
      TabIndex        =   24
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1800
      TabIndex        =   21
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label12 
      Caption         =   "Descuento Permitido"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1800
      TabIndex        =   19
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Facturas pendientes"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1800
      TabIndex        =   17
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Monto adeudado"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Limite de Credito"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Incoporado el"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Tel1 / Tel2 / Fax"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Dirección"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   120
      TabIndex        =   10
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
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Limpiar()
Clientes.Index = "PrimaryKey"
Clientes.MoveLast
Text8.Text = Clientes!Codigo + 1
Combo1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Label9.Caption = FormatCurrency(0, 2)
Label11.Caption = 0
Label13.Caption = ""
Command1.Enabled = False
Command2.Enabled = False
End Sub

Private Sub Nuevo()
Command1.Enabled = True
Command2.Enabled = False
End Sub

Private Sub Buscar()
If Combo1.Text = "- CLIENTE -" Then
    Combo1.Text = ""
End If
Clientes.Index = "Cliente"
Clientes.Seek "=", Combo1.Text
If Not Clientes.NoMatch Then
    TMP = Clientes!Codigo
    Text1.Text = TMP
    Facturas.Index = "Cliente"
    Facturas.Seek "=", Clientes!Codigo
    
    PENDIENTES = 0
    ADEUDA = 0
    If Not IsNull(Clientes!Direccion) Then
        Text2.Text = Clientes!Direccion
    End If
    If Not IsNull(Clientes!Tel1) Then
        Text3.Text = Clientes!Tel1
    End If
    If Not IsNull(Clientes!Tel2) Then
        Text4.Text = Clientes!Tel2
    End If
    If Not IsNull(Clientes!Fax) Then
        Text5.Text = Clientes!Fax
    End If
    If Not IsNull(Clientes!FEC_ING) Then
        Label13.Caption = Clientes!FEC_ING
    End If
    If Not IsNull(Clientes!LIMITE_CRE) Then
        Text6.Text = Clientes!LIMITE_CRE
    End If
    If Not IsNull(Clientes!Descuento) Then
        Text7.Text = Clientes!Descuento
    End If
        Label9 = FormatCurrency(0, 2)
        Label11 = 0
    If Not Facturas.NoMatch Then
        Rango = True
        While Rango And Not Facturas.EOF
            If Facturas!Cli_cod = TMP Then
                Sigla = Mid(Facturas!Num_Fac, 1, 1)
                If Not Facturas!Cancelado Then
                    If (Sigla = "R" Or Sigla = "A" Or Sigla = "T") Then
                        PENDIENTES = PENDIENTES + 1
                        Label11.Caption = PENDIENTES
                        ADEUDA = ADEUDA + Facturas!Monto_S
                        Label9.Caption = FormatCurrency(ADEUDA, 2)
                    End If
                End If
                Facturas.MoveNext
            Else
                Rango = False
            End If
        Wend
    Else
        Label9.Caption = FormatCurrency(0, 2)
        Label11.Caption = 0
    End If
Else
    'Call Limpiar
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub Command1_Click()
If Len(Text1.Text) = 0 Then
    MsgBox "Digite el ID del Cliente"
    Text1.SelStart = 0
    Text1.SelLength = Len(Combo1.Text)
    Text1.SetFocus
ElseIf Len(Combo1.Text) = 0 Then
    MsgBox "Digite el nombre del Cliente"
    Combo1.SelStart = 0
    Combo1.SelLength = Len(Combo1.Text)
    Combo1.SetFocus
ElseIf Len(Text2.Text) = 0 Then
    MsgBox "Digite la direccion del cliente"
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    Text2.SetFocus
ElseIf Len(Text3.Text) = 0 And Len(Text3.Text) = 0 And Len(Text3.Text) = 0 Then
    MsgBox "Digite el numero de telefono del Cliente"
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.SetFocus
Else
    If Len(Text6.Text) = 0 Or Not IsNumeric(Text6.Text) Then
        Text6.Text = 0
    End If
    If Len(Text7.Text) = 0 Or Not IsNumeric(Text7.Text) Then
        Text7.Text = 0
    End If
    Clientes.AddNew
    Clientes!Codigo = Text1.Text
    Clientes!Cliente = Combo1.Text
    Clientes!Direccion = Text2.Text
    Clientes!Tel1 = Text3.Text
    Clientes!Tel2 = Text4.Text
    Clientes!Fax = Text5.Text
    Clientes!FEC_ING = Date
    Clientes!LIMITE_CRE = Text6.Text
    Clientes!Descuento = Text7.Text
    Clientes.Update
    Text1.SetFocus
    Text1.Text = ""
    Call Limpiar
End If
End Sub

Private Sub Command2_Click()
If Len(Text1.Text) = 0 Then
    MsgBox "Digite el ID del Cliente"
    Text1.SelStart = 0
    Text1.SelLength = Len(Combo1.Text)
    Text1.SetFocus
ElseIf Len(Combo1.Text) = 0 Then
    MsgBox "Digite el nombre del Cliente"
    Combo1.SelStart = 0
    Combo1.SelLength = Len(Combo1.Text)
    Combo1.SetFocus
ElseIf Len(Text2.Text) = 0 Then
    MsgBox "Digite la direccion del cliente"
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    Text2.SetFocus
ElseIf Len(Text3.Text) = 0 And Len(Text3.Text) = 0 And Len(Text3.Text) = 0 Then
    MsgBox "Digite el numero de telefono del Cliente"
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.SetFocus
Else
    If Len(Text6.Text) = 0 Or Not IsNumeric(Text6.Text) Then
        Text6.Text = 0
    End If
    If Len(Text7.Text) = 0 Or Not IsNumeric(Text7.Text) Then
        Text7.Text = 0
    End If
    Clientes.Edit
    Clientes!Codigo = Text1.Text
    Clientes!Cliente = Combo1.Text
    Clientes!Direccion = Text2.Text
    Clientes!Tel1 = Text3.Text
    Clientes!Tel2 = Text4.Text
    Clientes!Fax = Text5.Text
    Clientes!LIMITE_CRE = Text6.Text
    Clientes!Descuento = Text7.Text
    Clientes.Update
    Text1.SetFocus
    Text1.Text = ""
    Call Limpiar
End If
End Sub

Private Sub Command3_Click()
Clientes.Delete
Call Limpiar
Text1.SetFocus
Text1.Text = ""
End Sub

Private Sub Command4_Click()
Text1.Text = ""
Form11.Hide
Form1.Show
End Sub

Private Sub Form_Activate()
Text1.Text = ""
Command1.Enabled = False
Command2.Enabled = False
Set Facturas = Inventa.OpenRecordset("FACTURAS")
Set Clientes = Inventa.OpenRecordset("CLIENTES")
Call Limpiar
Clientes.Index = "Cliente"
Clientes.MoveFirst
While Clientes.EOF = False
    If Clientes!Codigo <> 1 Then
        Combo1.AddItem Clientes!Cliente
    End If
    Clientes.MoveNext
Wend
If Form0.Option3.Value Then
    Text1.SetFocus
Else
    Combo1.SetFocus
End If
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
        Clientes.Index = "PrimaryKey"
        Clientes.Seek "=", Text1.Text
        If Clientes.NoMatch Or FormatNumber(Text1.Text) = 1 Then
            Text1.Text = Text8.Text
        End If
        If Text1.Text <> Text8.Text Then
            If Not Clientes.NoMatch Then
                Command2.Enabled = True
                Command1.Enabled = False
                Combo1.Text = Clientes!Cliente
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
If KeyAscii = 13 Then
    KeyAscii = 0
    If Command1.Enabled Then
        Command1.SetFocus
    Else
        If Form0.Option3.Value Then
            Command2.SetFocus
        Else
            Command4.SetFocus
        End If
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
