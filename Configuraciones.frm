VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametros"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   Icon            =   "Configuraciones.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Contraseñas"
      Height          =   1815
      Left            =   3840
      TabIndex        =   37
      Top             =   5520
      Width           =   3615
      Begin VB.TextBox Text15 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   43
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text14 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   42
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text13 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   41
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text12 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   40
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label18 
         Caption         =   "Administrador:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Super Usuario:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pie de Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   32
      Top             =   1080
      Width           =   7335
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   36
         Top             =   600
         Width           =   6135
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   35
         Top             =   240
         Width           =   6135
      End
      Begin VB.Label Label16 
         Caption         =   "Linea 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Linea 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Encabezados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   27
      Top             =   0
      Width           =   7335
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         MaxLength       =   40
         TabIndex        =   29
         Top             =   600
         Width           =   6135
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         MaxLength       =   40
         TabIndex        =   28
         Top             =   240
         Width           =   6135
      End
      Begin VB.Label Label14 
         Caption         =   "Linea 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Linea 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   26
      Top             =   7440
      Width           =   3135
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   25
      Top             =   7440
      Width           =   3135
   End
   Begin VB.Frame Frame4 
      Caption         =   "Concecutivos e Impresoras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   7335
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   240
         Top             =   2760
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1920
         Width           =   2415
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1440
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   960
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   14
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   13
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "Imprimir en"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   45
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "¡No alterar estos valores si se esta facturando!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   600
         TabIndex        =   24
         Top             =   2760
         Width           =   6135
      End
      Begin VB.Label Label11 
         Caption         =   "Imprimir en"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   23
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Imprimir en"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Imprimir en"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   21
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Imprimir en"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   20
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Cuentas por Pagar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Proformas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Timbradas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Impuesto / Descuento / Aumento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   3615
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Aumento Genral"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Descuento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Impuesto de Ventas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PRN As Printer
Private Sub Command1_Click()
If Text1.Text = "" Then
    Text1.Text = 0
End If
If Text2.Text = "" Then
    Text2.Text = 0
End If
If Text4.Text = "" Then
    Text4.Text = 0
End If
If Text5.Text = "" Then
    Text5.Text = 0
End If
If Text6.Text = "" Then
    Text6.Text = 0
End If
If Text7.Text = "" Then
    Text7.Text = 0
End If
Setup.Edit
If Text12.Text <> Text13.Text Then
    MsgBox "La contraseña de Super Usuario no coincide"
ElseIf Text14.Text <> Text15.Text Then
    MsgBox "La contraseña del Administrador no coincide"
ElseIf IsNumeric(Text1.Text) = False Or Text1.Text < 0 Then
    MsgBox "Impuesto invalido"
    Text1.SetFocus
    Text1.SelStart = 0
    Text1.SelLength = 10
ElseIf IsNumeric(Text2.Text) = False Or Text2.Text < 0 Then
    MsgBox "Descuento invalido"
    Text2.SetFocus
    Text2.SelStart = 0
    Text2.SelLength = 10
ElseIf IsNumeric(Text4.Text) = False Or Text4.Text < 0 Then
    MsgBox "Consecutivo de Timbradas invalido"
    Text4.SetFocus
    Text4.SelStart = 0
    Text4.SelLength = 10
ElseIf IsNumeric(Text5.Text) = False Or Text5.Text < 0 Then
    MsgBox "Consecutivo de Recibos invalido"
    Text5.SetFocus
    Text5.SelStart = 0
    Text5.SelLength = 10
ElseIf IsNumeric(Text6.Text) = False Or Text6.Text < 0 Then
    MsgBox "Consecutivo de Proformas invalido"
    Text6.SetFocus
    Text6.SelStart = 0
    Text6.SelLength = 10
ElseIf IsNumeric(Text7.Text) = False Or Text7.Text < 0 Then
    MsgBox "Consecutivo de Cuentas por cobrar invalido"
    Text7.SetFocus
    Text7.SelStart = 0
    Text7.SelLength = 10
Else
    Setup!IMP = Text1.Text
    Setup!DES_MAX = Text2.Text
    Setup!TIM_CON = Text4.Text
    Setup!REC_CON = Text5.Text
    Setup!PRO_CON = Text6.Text
    Setup!Cue_Con = Text7.Text
    Setup!Linea1 = Text8.Text
    Setup!linea2 = Text9.Text
    Setup!pie1 = Text10.Text
    Setup!pie2 = Text11.Text
    Setup!SU = Text12.Text
    Setup!AD = Text14.Text
End If

If IsNull(Combo1.Text) Then
    Setup!tim_prn = "- Ninguna -"
Else
    Setup!tim_prn = Combo1.Text
End If
If IsNull(Combo2.Text) Then
    Setup!rec_prn = "- Ninguna -"
Else
    Setup!rec_prn = Combo2.Text
End If
If IsNull(Combo3.Text) Then
    Setup!pro_prn = "- Ninguna -"
Else
    Setup!pro_prn = Combo3.Text
End If
If IsNull(Combo4.Text) Then
    Setup!CUE_PRN = "- Ninguna -"
Else
    Setup!CUE_PRN = Combo4.Text
End If
Setup.Update
Call Command7_Click
End Sub

Private Sub Command7_Click()
Form7.Hide
Combo1.Clear
Combo2.Clear
Combo3.Clear
Combo4.Clear
Form1.Show
End Sub

Private Sub Form_Activate()
Set Setup = Inventa.OpenRecordset("SETUP")
Setup.MoveFirst
Text1.Text = Setup!IMP
Text2.Text = Setup!DES_MAX
Text4.Text = Setup!TIM_CON
Text5.Text = Setup!REC_CON
Text6.Text = Setup!PRO_CON
Text7.Text = Setup!Cue_Con
Text8.Text = Setup!Linea1
Text9.Text = Setup!linea2
Text10.Text = Setup!pie1
Text11.Text = Setup!pie2
Text12.Text = Setup!SU
Text13.Text = Setup!SU
Text14.Text = Setup!AD
Text15.Text = Setup!AD
Combo1.AddItem "- Ninguna -"
For Each PRN In Printers
    Combo1.AddItem PRN.DeviceName
Next
Combo2.AddItem "- Ninguna -"
For Each PRN In Printers
    Combo2.AddItem PRN.DeviceName
Next
Combo3.AddItem "- Ninguna -"
For Each PRN In Printers
    Combo3.AddItem PRN.DeviceName
Next
Combo4.AddItem "- Ninguna -"
For Each PRN In Printers
    Combo4.AddItem PRN.DeviceName
Next
Combo5.AddItem "- Ninguna -"
For Each PRN In Printers
    Combo5.AddItem PRN.DeviceName
Next
I = 0
E = 0
While I < Combo1.ListCount And E = 0
    Combo1.ListIndex = I
    If Combo1.Text <> Setup!tim_prn Then
        I = I + 1
    Else
        E = 1
    End If
    If Combo1.Text <> Setup!tim_prn And I = Combo1.ListCount Then
        Combo1.ListIndex = 0
    End If
Wend
I = 0
E = 0
While I < Combo2.ListCount And E = 0
    Combo2.ListIndex = I
    If Combo2.Text <> Setup!rec_prn Then
        I = I + 1
    Else
        E = 1
    End If
    If Combo2.Text <> Setup!rec_prn And I = Combo2.ListCount Then
        Combo2.ListIndex = 0
    End If
Wend
I = 0
E = 0
While I < Combo3.ListCount And E = 0
    Combo3.ListIndex = I
    If Combo3.Text <> Setup!pro_prn Then
        I = I + 1
    Else
        E = 1
    End If
    If Combo3.Text <> Setup!pro_prn And I = Combo3.ListCount Then
        Combo3.ListIndex = 0
    End If
Wend
I = 0
E = 0
While I < Combo4.ListCount And E = 0
    Combo4.ListIndex = I
    If Combo4.Text <> Setup!CUE_PRN Then
        I = I + 1
    Else
        E = 1
    End If
    If Combo4.Text <> Setup!CUE_PRN And I = Combo4.ListCount Then
        Combo4.ListIndex = 0
    End If
Wend
I = 0
E = 0
While I < Combo5.ListCount And E = 0
    Combo5.ListIndex = I
    If Combo5.Text <> Setup!FAX_PRN Then
        I = I + 1
    Else
        E = 1
    End If
    If Combo5.Text <> Setup!FAX_PRN And I = Combo5.ListCount Then
        Combo5.ListIndex = 0
    End If
Wend
Text1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Command7_Click
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
Text2.SetFocus
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
If (KeyAscii = 13) Then
Text3.SetFocus
End If
Dim KEY As String
KEY = Chr(KeyAscii)
If (KEY < "0" Or KEY > "9") Then
If (KeyAscii <> 8) Then
KeyAscii = 0
End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
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
If (KeyAscii = 13) Then
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
If (KeyAscii = 13) Then
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
If (KeyAscii = 13) Then
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
If (KeyAscii = 13) Then
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

Private Sub Timer1_Timer()
If Label12.Visible = True Then
    Label12.Visible = False
Else
    Label12.Visible = True
End If
End Sub
