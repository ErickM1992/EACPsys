VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta por Código o Descripción"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "Busqueda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1080
      MaxLength       =   12
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resultados de Busqueda"
      Height          =   2175
      Left            =   120
      TabIndex        =   31
      Top             =   960
      Width           =   4575
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4335
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   32
         Top             =   1920
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Busqueda por equivalencias"
      Height          =   1215
      Left            =   120
      TabIndex        =   29
      Top             =   3240
      Width           =   4575
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4335
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   135
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle"
      Height          =   3975
      Left            =   4800
      TabIndex        =   12
      Top             =   960
      Width           =   3255
      Begin VB.Label Label1 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Precio"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Precio I.V.I."
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Ubicación 1"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Ubicación 2"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Código"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """¢""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   0
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Última compra"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Equivalencia"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   3600
         Width           =   1455
      End
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6480
      MaxLength       =   10
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3720
      MaxLength       =   40
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label20 
      Caption         =   "Equivalente"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label22 
      Caption         =   "Ubicacion 1"
      Height          =   255
      Left            =   2880
      TabIndex        =   33
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label19 
      Caption         =   "Código"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label18 
      Caption         =   "Ubicación 2"
      Height          =   255
      Left            =   5520
      TabIndex        =   10
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Descripción"
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Equivalencias(ECGNTE As String)
Set ECGList = Inventa.OpenRecordset("ECGLIST")
List2.Clear

ECGList.Index = "PrimaryKey"
ECGList.Seek "=", ECGNTE
Tabla.Index = "Eqvui"
If ECGNTE <> "" Then
    Tabla.Seek "=", ECGNTE
    If Not Tabla.NoMatch Then
        Igual = True
        While Igual
            List2.AddItem Tabla!Codigo & " - " & Tabla!Cantidad
            Tabla.MoveNext
            If Tabla.EOF Then
                Igual = False
            Else
                If ECGNTE <> Tabla!Equivalente Then
                    Igual = False
                End If
            End If
        Wend
    End If
End If
End Sub
Private Sub Busqueda_codigo()
Tabla.Index = "CODIGO"
Tabla.Seek "=", Text1.Text
If Tabla.NoMatch = False Then
    List1.AddItem Mid(Tabla!descrip & "                                           ", 1, 43) & Tabla!Codigo
    List1.SetFocus
    List1.ListIndex = 0
Else
    MsgBox "Codigo no existe"
    Text1.SelStart = 0
    Text1.SelLength = 10
    Text1.SetFocus
End If
End Sub

Private Sub Busqueda_descripcion()
cadena = UCase(Text2.Text)
L = Len(cadena)
    Tabla.Index = "DATOS"
    Tabla.MoveFirst
    While Tabla.EOF = False
        If Mid(Tabla!descrip, 1, L) = cadena Then
            ProgressBar1.Value = Tabla.PercentPosition
            List1.AddItem Mid(Tabla!descrip & "                                           ", 1, 43) & Tabla!Codigo
            Tabla.MoveNext
        Else
            Tabla.MoveNext
        End If
    Wend
    If List1.ListCount > 0 Then
        List1.SetFocus
        List1.ListIndex = 0
    Else
        Text2.SelStart = 0
        Text2.SelLength = 40
        Text2.SetFocus
    End If
ProgressBar1.Value = 100
End Sub

Private Sub Busqueda_Ubicacion1()
    cadena = UCase(Text3.Text)
    L = Len(cadena)
    Tabla.Index = "U1"
    Tabla.MoveFirst
    While Tabla.EOF = False
        If Mid(Tabla!gabeta, 1, L) = cadena Then
            
            List1.AddItem Mid(Tabla!descrip & "                                           ", 1, 43) & Tabla!Codigo
            Tabla.MoveNext
        Else
            Tabla.MoveNext
        End If
    Wend
    If List1.ListCount > 0 Then
        List1.SetFocus
        List1.ListIndex = 0
    Else
        Text2.SelStart = 0
        Text2.SelLength = 40
        Text2.SetFocus
    End If
End Sub

Private Sub Busqueda_Ubicacion2()
    cadena = UCase(Text4.Text)
    L = Len(cadena)
    Tabla.Index = "U2"
    Tabla.MoveFirst
    While Tabla.EOF = False
        If Mid(Tabla!ubicacion, 1, L) = cadena Then
            
            List1.AddItem Mid(Tabla!descrip & "                                           ", 1, 43) & Tabla!Codigo
            Tabla.MoveNext
        Else
            Tabla.MoveNext
        End If
    Wend
    If List1.ListCount > 0 Then
        List1.SetFocus
        List1.ListIndex = 0
    Else
        Text2.SelStart = 0
        Text2.SelLength = 40
        Text2.SetFocus
    End If
End Sub

Private Sub Command1_Click()
List1.Clear
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""
Label12.Caption = ""
Label16.Caption = ""
Label17.Caption = ""
Text5.Text = ""
If Len(Text1.Text) <> 0 Then
    Text1.Text = Mid("000000000", 1, 10 - Len(Text1.Text)) & Text1.Text
End If
Text1.SetFocus
If Len(Text1) Then
    Call Busqueda_codigo
ElseIf Len(Text2) Then
    Call Busqueda_descripcion
ElseIf Len(Text3) Then
    Call Busqueda_Ubicacion1
ElseIf Len(Text4) Then
    Call Busqueda_Ubicacion2
End If
End Sub

Private Sub Command2_Click()
Form2.Hide
List1.Clear
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""
Label12.Caption = ""
Label16.Caption = ""
Form1.Show
End Sub

Private Sub Form_Activate()
ProgressBar1.Value = 100
Set Tabla = Inventa.OpenRecordset("inventa")
Set Setup = Inventa.OpenRecordset("Setup")
Text1.SetFocus
End Sub

Private Sub Form_Deactivate()
Call Command2_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Command2_Click
End Sub

Private Sub List1_Click()
TMP = Mid(List1.Text, 44, 10)
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Tabla.Index = "CODIGO"
Tabla.Seek "=", TMP
If Tabla.NoMatch = False Then
    Label12.Caption = Tabla!Codigo
    If Not IsNull(Tabla!Stock) Then
        If Tabla!Cantidad >= Tabla!Stock Then
            Label6.ForeColor = H80000012
        Else
            Label6.ForeColor = &HFF&
        End If
    End If
    Label6.Caption = Tabla!Cantidad
    Label7.Caption = FormatCurrency((Tabla!P_Venta), 2)
    Label8.Caption = FormatCurrency((Tabla!P_Venta + Tabla!P_Venta * (Setup!IMP / 100)), 2)
    If IsNull(Tabla!gabeta) Then
        Label9.Caption = ""
    Else
        Label9.Caption = Tabla!gabeta
    End If
    If IsNull(Tabla!ubicacion) Then
        Label10.Caption = ""
    Else
        Label10.Caption = Tabla!ubicacion
    End If
    If IsNull(Tabla!Fe_Ult_Com) Then
        Label16.Caption = ""
    Else
        FUV = Tabla!Fe_Ult_Com
        Label16.Caption = Tabla!Fe_Ult_Com
    End If
    If Tabla!Equivalente = "*" Or IsNull(Tabla!Equivalente) Then
        Label17.Caption = ""
    Else
        Label17.Caption = Tabla!Equivalente
    End If
End If
Call Equivalencias(Label17)
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1.SelStart = 0
    Text1.SelLength = 10
    Text1.SetFocus
End If
End Sub

Private Sub List2_DblClick()
Text1 = List2.Text
Call Text1_KeyPress(13)
End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call List2_DblClick
End If
End Sub

Private Sub Text1_GotFocus()
Text2.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1.Text = "" Then
        Text2.SetFocus
    Else
        Call Command1_Click
    End If
End If
If KeyAscii = 27 Then
    Form2.Hide
    Form1.Show
End If
End Sub

Private Sub Text2_GotFocus()
Text1.Text = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text2.Text = "" Then
        Text1.SetFocus
    Else
        Call Command1_Click
    End If
End If
If KeyAscii = 27 Then
    Form2.Hide
    Form1.Show
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub

Private Sub Text5_Change()
    Call Equivalencias(Text5)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    List2.SetFocus
End If
End Sub
