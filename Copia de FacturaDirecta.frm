VERSION 5.00
Begin VB.Form Form13 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FACTURACION"
   ClientHeight    =   6105
   ClientLeft      =   2655
   ClientTop       =   2565
   ClientWidth     =   9375
   Icon            =   "FacturaDirecta.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   407
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Tipo de Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   36
      Top             =   120
      Width           =   4455
      Begin VB.OptionButton Option3 
         Caption         =   "Proforma"
         Height          =   255
         Left            =   3000
         TabIndex        =   39
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Factura 2"
         Height          =   255
         Left            =   1560
         TabIndex        =   38
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Factura 1"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Height          =   50
      Left            =   120
      TabIndex        =   35
      Top             =   5520
      Width           =   9135
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6720
      TabIndex        =   32
      Top             =   960
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   31
      Top             =   960
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Exento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      ItemData        =   "FacturaDirecta.frx":000C
      Left            =   120
      List            =   "FacturaDirecta.frx":0013
      Style           =   1  'Checkbox
      TabIndex        =   13
      Top             =   2880
      Width           =   9135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5640
      Width           =   2000
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   5520
      TabIndex        =   0
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Borrar artículos marcados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5640
      Width           =   2550
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar e Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   2000
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   8
      Text            =   "1"
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "0"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   12
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   5520
      MaxLength       =   40
      TabIndex        =   4
      Top             =   1560
      Width           =   3735
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      ItemData        =   "FacturaDirecta.frx":001E
      Left            =   3840
      List            =   "FacturaDirecta.frx":0025
      TabIndex        =   14
      Top             =   1920
      Width           =   5415
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H8000000F&
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
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000F&
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
      ItemData        =   "FacturaDirecta.frx":0030
      Left            =   120
      List            =   "FacturaDirecta.frx":0032
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Modo de pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4800
      TabIndex        =   16
      Top             =   120
      Width           =   4455
      Begin VB.OptionButton Option5 
         Caption         =   "Credito"
         Height          =   375
         Left            =   2280
         TabIndex        =   41
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Contado"
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Line Line4 
      X1              =   608
      X2              =   616
      Y1              =   344
      Y2              =   352
   End
   Begin VB.Line Line3 
      X1              =   616
      X2              =   608
      Y1              =   272
      Y2              =   280
   End
   Begin VB.Line Line2 
      X1              =   288
      X2              =   280
      Y1              =   344
      Y2              =   352
   End
   Begin VB.Line Line1 
      X1              =   280
      X2              =   288
      Y1              =   272
      Y2              =   280
   End
   Begin VB.Shape Shape4 
      Height          =   975
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Width           =   4815
   End
   Begin VB.Shape Shape3 
      Height          =   1215
      Left            =   4200
      Top             =   4080
      Width           =   5055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "¢0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6360
      TabIndex        =   34
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   33
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
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
      Left            =   7440
      TabIndex        =   30
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unitario"
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
      Left            =   5880
      TabIndex        =   29
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cant."
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
      Left            =   5160
      TabIndex        =   28
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripción"
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
      Left            =   360
      TabIndex        =   27
      Top             =   2640
      Width           =   4815
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   26
      Top             =   1920
      Width           =   765
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   25
      Top             =   1560
      Width           =   600
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Búsqueda por parte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3720
      TabIndex        =   24
      Top             =   1560
      Width           =   1680
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2400
      TabIndex        =   23
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2760
      TabIndex        =   15
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Copias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   22
      Top             =   4080
      Width           =   585
   End
   Begin VB.Label Label5 
      Caption         =   "¢0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   21
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Descuento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   285
      TabIndex        =   20
      Top             =   4920
      Width           =   930
   End
   Begin VB.Label Label3 
      Caption         =   "¢0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Subtotal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   495
      TabIndex        =   18
      Top             =   4560
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Descuento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   17
      Top             =   2280
      Width           =   930
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Imprimir()
J = 0
While J < Text5.Text
FDF = False

Dim x As Printer
Printer.PaperSize = 5
For Each x In Printers
    If Option1.Value = True Then
        If x.DeviceName = Setup!tim_prn Then
            Set Printer = x
            Printer.Font = "Courier New"
            Printer.FontSize = 10
        End If
    ElseIf Option2.Value = True Then
        If x.DeviceName = Setup!rec_prn Then
            Set Printer = x
            Printer.Font = "Courier New"
            Printer.FontSize = 10
        End If
    ElseIf Option3.Value = True Then
        If x.DeviceName = Setup!pro_prn Then
            Set Printer = x
            Printer.Font = "Courier New"
            Printer.FontSize = 10
        End If
    End If
Next


If Option1.Value = True Then 'Timbrada
    Vendedores.Index = "PrimaryKey"
    Vendedores.Seek "=", Facturas!VEN_COD
    Detalle.Index = "Detalle"
    Detalle.Seek "=", Facturas!Num_Fac
    Tabla.Index = "CODIGO"
    If Facturas!Cancelado = True Then
        MODO_PAGO = "CONTADO"
    Else
        MODO_PAGO = "CREDITO"
    End If
    Printer.PaperSize = 256
    Printer.Width = 12240
    Printer.Height = 8155
    Printer.Font = "Courier New"
    Printer.FontSize = 10
    While FDF = False And Detalle!Num_Fac = Facturas!Num_Fac
        ACUMULADO = 0
        Printer.Print "                                                                " & MODO_PAGO
        Printer.Print " "
        Printer.Print " "
        Printer.Print "                                                " & Day(Date) & "           " & Month(Date) & "1          " & Year(Date)
        Printer.Print " "
        Printer.Print "         " & Mid(Facturas!CLI_COD & "     ", 1, 5) & "                                                    " & Mid("*           ", 1, 12 - (Len(Facturas!Num_Fac) - 1)) & Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1)
        Printer.Print "         " & Facturas!Cli_Nom
        Printer.Print " "
        Printer.Print " "
        Printer.Print " "
        Printer.Print " "
        i = 1
        While i < 14
            If Not Detalle.EOF Then
                If Detalle!Num_Fac = Facturas!Num_Fac Then
                    Tabla.Seek "=", Detalle!Codigo
                    PRECIO_UNITARIO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!PRE_UNI, 2)))) & FormatNumber(Detalle!PRE_UNI, 2)
                    MONTO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!PRE_UNI * Detalle!CANTIDAD, 2)))) & FormatNumber((Detalle!PRE_UNI * Detalle!CANTIDAD), 2)
                    CANT = Mid("    ", 1, (4 - Len(Detalle!CANTIDAD))) & Detalle!CANTIDAD
                    ACUMULADO = ACUMULADO + (Detalle!CANTIDAD * Detalle!PRE_UNI)
                    Printer.Print Detalle!Codigo & " " & CANT & " " & Mid(Tabla!descrip & "                                        ", 1, 32) & " " & PRECIO_UNITARIO & "     " & MONTO
                    Detalle.MoveNext
                Else
                    FDF = True
                    Printer.Print " "
                End If
            Else
                Printer.Print " "
            End If
            i = i + 1
        Wend
        SUBTOTAL = Mid("            ", 1, 12 - Len(FormatCurrency(ACUMULADO, 2))) & FormatCurrency(ACUMULADO, 2)
        DES_TMP = ACUMULADO * (Facturas!Des / 100)
        Descuento = Mid("            ", 1, 12 - Len(FormatCurrency(DES_TMP, 2))) & FormatCurrency(DES_TMP, 2)
        IMP_TMP = (ACUMULADO - DES_TMP) * (Facturas!IMP / 100)
        IMPUESTO = Mid("            ", 1, 12 - Len(FormatCurrency(IMP_TMP, 2))) & FormatCurrency(IMP_TMP, 2)
        TOT_TMP = (ACUMULADO - DES_TMP) + IMP_TMP
        TOTAL = Mid("            ", 1, 12 - Len(FormatCurrency(TOT_TMP, 2))) & FormatCurrency(TOT_TMP, 2)
        Printer.Print " "
        Printer.Print "                                                                  " & SUBTOTAL
        Printer.Print "                                                                  " & Descuento '32
        Printer.Print "                                                                  " & IMPUESTO '33"
        Printer.Print "                                                                  " & TOTAL '34
        Printer.EndDoc
    Wend
ElseIf Option2.Value = True Then 'Recibo
    Vendedores.Index = "PrimaryKey"
    Vendedores.Seek "=", Facturas!VEN_COD
    Detalle.Index = "Detalle"
    Detalle.Seek "=", Facturas!Num_Fac
    Tabla.Index = "CODIGO"
    If Facturas!Cancelado = True Then
        MODO_PAGO = "CONTADO"
    Else
        MODO_PAGO = "CREDITO"
    End If
    Printer.PaperSize = 256
    Printer.Width = 12240
    Printer.Height = 8155
    Printer.Font = "Courier New"
    Printer.FontSize = 10
    While FDF = False And Facturas!Num_Fac = Detalle!Num_Fac
        Printer.Print Mid(Setup!Linea1 & "                                        ", 1, 40) & "                                   Fecha: " & FormatDateTime(Date, vbShortDate)
        Printer.Print Mid(Setup!linea2 & "                                        ", 1, 40) & "                                   Factura: " & Mid("            ", 1, 12 - Len(Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1))) & Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1)
        Printer.Print "                                                                           Modo de Pago: " & MODO_PAGO
        Printer.Print "Cliente :" & Facturas!CLI_COD
        Printer.Print "Nombre  :" & Facturas!Cli_Nom
        Printer.Print "Vendedor:" & Vendedores!Nombre
        Printer.Print " "
        Printer.Print "   CODIGO      CANT.    ARTICULO                                        VALOR           MONTO   " '07
        i = 1
        While i < 19
            If Not Detalle.EOF Then
                If Detalle!Num_Fac = Facturas!Num_Fac Then
                    Tabla.Seek "=", Detalle!Codigo
                    PRECIO_UNITARIO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!PRE_UNI, 2)))) & FormatNumber(Detalle!PRE_UNI, 2)
                    MONTO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!PRE_UNI * Detalle!CANTIDAD, 2)))) & FormatNumber((Detalle!PRE_UNI * Detalle!CANTIDAD), 2)
                    CANT = Mid("    ", 1, (4 - Len(Detalle!CANTIDAD))) & Detalle!CANTIDAD
                    ACUMULADO = ACUMULADO + (Detalle!CANTIDAD * Detalle!PRE_UNI)
                    Printer.Print " " & Detalle!Codigo & "    " & CANT & "     " & Mid(Tabla!descrip & "                                        ", 1, 40) & "    " & PRECIO_UNITARIO & "    " & MONTO
                    Detalle.MoveNext
                Else
                    FDF = True
                    Printer.Print " "
                End If
            Else
                Printer.Print " "
            End If
            i = i + 1
        Wend
        SUBTOTAL = Mid("            ", 1, 12 - Len(FormatCurrency(ACUMULADO, 2))) & FormatCurrency(ACUMULADO, 2)
        DES_TMP = ACUMULADO * (Facturas!Des / 100)
        Descuento = Mid("            ", 1, 12 - Len(FormatCurrency(DES_TMP, 2))) & FormatCurrency(DES_TMP, 2)
        IMP_TMP = (ACUMULADO - DES_TMP) * (Facturas!IMP / 100)
        IMPUESTO = Mid("            ", 1, 12 - Len(FormatCurrency(IMP_TMP, 2))) & FormatCurrency(IMP_TMP, 2)
        TOT_TMP = (ACUMULADO - DES_TMP) + IMP_TMP
        TOTAL = Mid("            ", 1, 12 - Len(FormatCurrency(TOT_TMP, 2))) & FormatCurrency(TOT_TMP, 2)
        Printer.Print " "
        Printer.Print "                                                                         SUBTOTAL:  " & SUBTOTAL
        Printer.Print "                                                                         DESCUENTO: " & Descuento '32
        Printer.Print "      *** NUESTRAS VENTAS SON EN FIRME ***    _____________________      IMPUESTO:  " & IMPUESTO '33"
        Printer.Print "       *** NO SE ACEPTAN DEVOLUCIONES ***       Recibido Conforme        TOTAL:     " & TOTAL '34
        Printer.EndDoc
    Wend
ElseIf Option3.Value = True Then
    Vendedores.Index = "PrimaryKey"
    Vendedores.Seek "=", Facturas!VEN_COD
    Detalle.Index = "Detalle"
    Detalle.Seek "=", Facturas!Num_Fac
    Tabla.Index = "CODIGO"
    If Facturas!Cancelado = True Then
        MODO_PAGO = "CONTADO"
    Else
        MODO_PAGO = "CREDITO"
    End If
    Printer.PaperSize = 256
    Printer.Width = 12240
    Printer.Height = 8155
    Printer.Font = "Courier New"
    Printer.FontSize = 10
    While FDF = False And Detalle!Num_Fac = Facturas!Num_Fac
        Printer.Print Mid(Setup!Linea1 & "                                        ", 1, 40) & "                                  Fecha: " & FormatDateTime(Date, vbShortDate)
        Printer.Print Mid(Setup!linea2 & "                                        ", 1, 40) & "                                  Proforma: " & Mid("            ", 1, 12 - Len(Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1))) & Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1)
        Printer.Print "                                                                          Modo de Pago: " & MODO_PAGO
        Printer.Print "Cliente :" & Facturas!CLI_COD
        Printer.Print "Nombre  :" & Facturas!Cli_Nom
        Printer.Print "Vendedor:" & Vendedores!Nombre
        Printer.Print " "
        Printer.Print "   CODIGO      CANT.    ARTICULO                                        VALOR           MONTO   " '07
        i = 1
        While i < 19
            If Not Detalle.EOF Then
                If Detalle!Num_Fac = Facturas!Num_Fac Then
                    Tabla.Seek "=", Detalle!Codigo
                    PRECIO_UNITARIO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!PRE_UNI, 2)))) & FormatNumber(Detalle!PRE_UNI, 2)
                    MONTO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!PRE_UNI * Detalle!CANTIDAD, 2)))) & FormatNumber((Detalle!PRE_UNI * Detalle!CANTIDAD), 2)
                    CANT = Mid("    ", 1, (4 - Len(Detalle!CANTIDAD))) & Detalle!CANTIDAD
                    ACUMULADO = ACUMULADO + (Detalle!CANTIDAD * Detalle!PRE_UNI)
                    Printer.Print " " & Detalle!Codigo & "    " & CANT & "     " & Mid(Tabla!descrip & "                                        ", 1, 40) & "    " & PRECIO_UNITARIO & "    " & MONTO
                    Detalle.MoveNext
                Else
                    FDF = True
                    Printer.Print " "
                End If
            Else
                Printer.Print " "
            End If
            i = i + 1
        Wend
        SUBTOTAL = Mid("            ", 1, 13 - Len(FormatCurrency(ACUMULADO, 2))) & FormatCurrency(ACUMULADO, 2)
        DES_TMP = ACUMULADO * (Facturas!Des / 100)
        Descuento = Mid("            ", 1, 13 - Len(FormatCurrency(DES_TMP, 2))) & FormatCurrency(DES_TMP, 2)
        IMP_TMP = (ACUMULADO - DES_TMP) * (Facturas!IMP / 100)
        IMPUESTO = Mid("            ", 1, 13 - Len(FormatCurrency(IMP_TMP, 2))) & FormatCurrency(IMP_TMP, 2)
        TOT_TMP = (ACUMULADO - DES_TMP) + IMP_TMP
        TOTAL = Mid("            ", 1, 13 - Len(FormatCurrency(TOT_TMP, 2))) & FormatCurrency(TOT_TMP, 2)
        Printer.Print " "
        Printer.Print "                                                                         SUBTOTAL:  " & SUBTOTAL
        Printer.Print "                                                                         DESCUENTO: " & Descuento
        Printer.Print Mid(Setup!pie1 & "                                        ", 1, 40) & " _____________________      IMPUESTO:  " & IMPUESTO
        Printer.Print Mid(Setup!pie1 & "                                        ", 1, 40) & "   Recibido Conforme        TOTAL:     " & TOTAL
        Printer.EndDoc
    Wend
End If
J = J + 1
Wend
End Sub

Private Sub Limpiar()
Option1.Value = True
Option2.Value = False
Option3.Value = False
Option4.Value = True
Option5.Value = False
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = "0"
Text5.Text = "1"
Text6.Text = ""
Text7.Text = ""
'Combo1.Text = ""
'Combo2.Text = ""
List1.Clear
List2.Clear
Label3.Caption = "¢0,00"
Label5.Caption = "¢0,00"
Label7.Caption = "¢0,00"
Label10.Caption = ""
End Sub

Private Sub Calcular()
Set Setup = Inventa.OpenRecordset("SETUP")
If Tabla!Por_Venta > Setup!Por_Ven Then
    Text4.Text = 20
    
    UNITARIO = Mid("            ", 1, 12 - Len(FormatNumber((Tabla!P_Unit * 2) - Tabla!P_Unit * (Text4.Text / 100)))) & FormatNumber((Tabla!P_Unit * 2) - Tabla!P_Unit * (Text4.Text / 100))
    MONTO = Mid("            ", 1, 12 - Len(FormatNumber(CDbl(UNITARIO) * Val(Text3.Text), 2))) & FormatNumber(CDbl(UNITARIO) * Val(Text3.Text), 2)
Else
    Text4.Text = Clientes!Descuento
    UNITARIO = Mid("            ", 1, 12 - Len(FormatNumber(Tabla!P_Venta - (Tabla!P_Venta * Clientes!Descuento / 100), 2))) & FormatNumber(Tabla!P_Venta - (Tabla!P_Venta * Clientes!Descuento / 100), 2)
    MONTO = Mid("            ", 1, 12 - Len(FormatNumber(CDbl(UNITARIO) * Val(Text3.Text), 2))) & FormatNumber(CDbl(UNITARIO) * Val(Text3.Text), 2)
End If
Label3.Caption = FormatCurrency((CDbl(MONTO) + (CDbl(Label3.Caption))), 2)
Text3.Text = Mid("    ", 1, 4 - Len(Text3.Text)) & Text3.Text
List2.AddItem Mid(Tabla!descrip & "                                            ", 1, 40) & " " & Text3.Text & " " & UNITARIO & " " & MONTO & "   " & Tabla!Codigo
Text1.SetFocus
End Sub


Private Sub Combo1_Click()
Clientes.Seek "=", Combo1.Text
If Combo1.Text = 1 Then
    Option4.Value = True
    Option5.Enabled = False
    Text6.Text = ""
    Text6.Enabled = True
End If
If (Combo1.Text <> 1 And Clientes.NoMatch = False) Then
    Option5.Enabled = True
    Text6.Text = Clientes!Cliente
    Text6.Enabled = False
End If
If Clientes.NoMatch = True Then
    MsgBox "EL CLIENTE NO EXISTE", vbCritical
    Combo1.SetFocus
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (Combo1.Text = "" Or Combo1.Text < "0" Or Combo1.Text > "9") Then
        Combo1.Text = 1
        Text6.SetFocus
     Else
        Call Combo1_LostFocus
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

Private Sub Combo1_LostFocus()
Clientes.Seek "=", Combo1.Text
If Combo1.Text = "" Then
    Combo1.Text = 1
End If
If Combo1.Text = 1 Then
    Option4.Value = True
    Option5.Enabled = False
    Text6.Text = ""
    Text6.Enabled = True
    Text6.SetFocus
End If
If (Combo1.Text <> 1 And Clientes.NoMatch = False) Then
    Option5.Enabled = True
    Text6.Text = Clientes!Cliente
    Text6.Enabled = False
    Option4.SetFocus
End If
If Clientes.NoMatch = True Then
    MsgBox "EL CLIENTE NO EXISTE", vbCritical
    Combo1.SetFocus
End If
End Sub

Private Sub Combo2_Click()
Vendedores.Seek "=", Combo2.Text
If (Combo2.Text = "" Or Combo2.Text < "0" Or Combo2.Text > "9") Then
    Text7.Text = ""
    Combo2.Text = ""
Else
    If Vendedores.NoMatch = False Then
        Text7.Text = Vendedores!Nombre
    Else
        MsgBox "EL VENDEDOR NO EXISTE", vbCritical
        Combo2.SetFocus
    End If
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (Combo2.Text = "" Or Combo2.Text < "0" Or Combo2.Text > "9") Then
        MsgBox "Digite su código de Vendedor"
        Combo2.SetFocus
        Combo2.Text = ""
     Else
        Text1.SetFocus
    End If
    Call Combo2_Click
End If
Dim KEY As String
KEY = Chr(KeyAscii)
If (KEY < "0" Or KEY > "9") Then
    If (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Combo2_LostFocus()
Vendedores.Seek "=", Combo2.Text
If Combo2.Text = "" Or Combo2.Text < "0" Or Combo2.Text > "9" Then
    Text7.Text = ""
    Combo2.Text = ""
Else
    If Vendedores.NoMatch = False Then
        Text7.Text = Vendedores!Nombre
    Else
        MsgBox "Vendedor no existe"
        Combo2.SetFocus
    End If
End If
End Sub

Private Sub Command1_Click()
Facturas.AddNew
Detalle.AddNew
Setup.MoveFirst
Setup.Edit
If Option1.Value = True Then
    Facturas!Num_Fac = "T" & Setup!TIM_CON
    Setup!TIM_CON = Setup!TIM_CON + 1
ElseIf Option2.Value = True Then
    Facturas!Num_Fac = "R" & Setup!REC_CON
    Setup!REC_CON = Setup!REC_CON + 1
ElseIf Option3.Value = True Then
    Facturas!Num_Fac = "P" & Setup!PRO_CON
    Setup!PRO_CON = Setup!PRO_CON + 1
End If
Setup.Update
If Option4.Value = True Then
    Facturas!Cancelado = True
    Facturas!Monto_F = Label7.Caption
    Facturas!Monto_S = 0
ElseIf Option5.Value = True Then
    Facturas!Cancelado = False
    Facturas!Monto_F = Label7.Caption
    Facturas!Monto_S = Label7.Caption
End If
If Clientes.NoMatch Then
    Facturas!CLI_COD = 1
    Facturas!Cli_Nom = Combo1.Text

Else
    Facturas!CLI_COD = Clientes!Codigo
    Facturas!Cli_Nom = Clientes!Cliente
End If
If Check1.Value = 1 Then
    Facturas!IMP = 0
Else
    Facturas!IMP = Setup!IMP
End If
Facturas!VEN_COD = Combo2.Text
Facturas!fecha = Date
Facturas!Des = Text4.Text
Facturas!Anulada = False
FAC_NUM = Facturas!Num_Fac
Facturas.Update
Tabla.Index = "CODIGO"
i = 0
While i < List2.ListCount
    List2.ListIndex = i
    Detalle!Num_Fac = FAC_NUM
    Detalle!CANTIDAD = Mid(List2.Text, 42, 4)
    Detalle!Codigo = Mid(List2.Text, 75, 10)
    Detalle!PRE_UNI = Mid(List2.Text, 47, 12)
    If Not Option3.Value Then
        Tabla.Seek "=", Mid(List2.Text, 75, 10)
        Tabla.Edit
        Tabla!CANTIDAD = Tabla!CANTIDAD - Detalle!CANTIDAD
        Tabla.Update
    End If
    Detalle.Update
    Detalle.AddNew
    i = i + 1
Wend
Facturas.Index = "PrimaryKey"
Facturas.Seek "=", FAC_NUM
Call Imprimir
Call Limpiar
Call Command3_Click
End Sub

Private Sub Command1_GotFocus()
Command1.BackColor = RGB(100, 200, 255)
End Sub

Private Sub Command1_LostFocus()
Command1.BackColor = &H8000000F
End Sub

Private Sub Command2_Click()
NUM = 0
i = 0
While i < List2.ListCount
    List2.ListIndex = i
    If List2.Selected(i) = True Then
        NUM = NUM + Mid(List2.Text, 60, 12)
        List2.RemoveItem (i)
    Else
        i = i + 1
    End If
Wend
Label3.Caption = FormatCurrency(FormatNumber(Label3.Caption, 2) - NUM)
End Sub

Private Sub Command2_GotFocus()
Command2.BackColor = RGB(100, 200, 255)
End Sub

Private Sub Command2_LostFocus()
Command2.BackColor = &H8000000F
End Sub

Private Sub Command3_Click()
Call Limpiar
Form13.Hide
Form1.Show
End Sub

Private Sub Command3_GotFocus()
Command3.BackColor = RGB(100, 200, 255)
End Sub

Private Sub Command3_LostFocus()
Command3.BackColor = &H8000000F
End Sub

Private Sub Form_Activate()
Set Setup = Inventa.OpenRecordset("SETUP")
Set Tabla = Inventa.OpenRecordset("INVENTA")
Set Clientes = Inventa.OpenRecordset("CLIENTES")
Set Vendedores = Inventa.OpenRecordset("VENDEDORES")
Set Facturas = Inventa.OpenRecordset("FACTURAS")
Set Detalle = Inventa.OpenRecordset("DETALLE")
Combo1.Clear
Combo2.Clear
Text4.Enabled = False
Text5.Enabled = False
'Llenar el combo con vendedores
Vendedores.Index = "PrimaryKey"
Vendedores.MoveFirst
While Vendedores.EOF = False
    Combo2.AddItem Vendedores!Codigo
    Vendedores.MoveNext
Wend
'Llenar el combo con clientes
Clientes.Index = "PrimaryKey"
Clientes.MoveFirst
While Clientes.EOF = False
    Combo1.AddItem Clientes!Codigo
    Clientes.MoveNext
Wend
Combo1.Text = 1
Option1.Value = True
Option1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Limpiar
Form3.Hide
Form1.Show
End Sub

Private Sub List1_DblClick()
Text1.SetFocus
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1.SetFocus
End If
End Sub

Private Sub Option1_Click()
Text5.Text = 1
Text5.Enabled = False
End Sub

Private Sub Option1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo1.SetFocus
End If
End Sub

Private Sub Option2_Click()
Text5.Enabled = True
End Sub

Private Sub Option2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo1.SetFocus
End If
End Sub

Private Sub Option3_Click()
Text5.Enabled = True
End Sub

Private Sub Option3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo1.SetFocus
End If
End Sub

Private Sub Option4_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    Combo1.SetFocus
ElseIf KeyAscii = 13 Then
    Text6.Enabled = True
    Combo2.SetFocus
End If
End Sub

Private Sub Option5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        Combo1.SetFocus
    End If
    If KeyAscii = 13 Then
        Text6.Enabled = False
        Combo2.SetFocus
    End If
End Sub

Private Sub List1_Click()
    TMP = Mid(List1.Text, 44, 10)
    Text1.Text = ""
    Text2.Text = ""
    Tabla.Index = "CODIGO"
    Tabla.Seek "=", TMP
    If Tabla.NoMatch = False Then
        Text1.Text = Tabla!Codigo
        Label10.Caption = Tabla!CANTIDAD
    End If
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = 10
Text2.Text = ""
Text3.Text = ""
Text3.Enabled = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1.Text <> "" Then
        List1.Clear
        If Len(Text1.Text) > 0 Then
            Text1.Text = Mid("0000000000", 1, 10 - Len(Text1.Text)) & Text1.Text
        End If
        Tabla.Index = "CODIGO"
        Tabla.Seek "=", Text1.Text
        If Tabla.NoMatch = False Then
            Label10.Caption = Tabla!CANTIDAD
            List1.AddItem Mid(Tabla!descrip & "                                             ", 1, 45) & Tabla!Codigo
            List1.SetFocus
            List1.ListIndex = List1.ListIndex
            Text3.Enabled = True
            Text3.SetFocus
        Else
            MsgBox "EL CODIGO NO EXISTE", vbCritical
            Text1.SelStart = 0
            Text1.SelLength = 10
            Text1.SetFocus
        End If
    Else
        Text2.Text = ""
        Text2.SetFocus
    End If
End If
If KeyAscii = 27 Then
    'Call Text4_LostFocus
    Text4.Enabled = True
    Text4.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    List1.Clear
    If Text2.Text = "" Then
        MsgBox "POR FAVOR DIGITE ALGUN CODIGO O ARTICULO", vbExclamation
        Text1.Text = ""
        Text1.SetFocus
    End If
    If Text2.Text <> "" Then
        cadena = UCase(Text2.Text)
        l = Len(cadena)
        Tabla.Index = "DATOS"
        Tabla.MoveFirst
        While Tabla.EOF = False
            If Mid(Tabla!descrip, 1, l) = cadena Then
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
            MsgBox "EL CODIGO NO EXISTE", vbCritical
            Label10.Caption = ""
            Text2.SelStart = 0
            Text2.SelLength = 40
            Text2.SetFocus
        End If
    End If
End If
End Sub

Private Sub Text3_GotFocus()
If Label10.Caption = 0 Then
    MsgBox "NO HAY ARTICULOS", vbCritical
    Text1.SetFocus
    Text1.SelStart = 0
    Text1.SelLength = 10
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text3.Text = "" Or Val(Text3.Text) = 0 Then
        MsgBox "CANTIDAD INVALIDA", vbCritical
        Text3.Text = ""
        Text3.SetFocus
    End If
    If Text3.Text <> "" Then
        If (Val(Text3.Text) > CDbl(Label10.Caption)) Then
            MsgBox "CANTIDAD INVALIDA", vbCritical
            Text3.SelStart = 0
            Text3.SelLength = 5
            Text3.SetFocus
        Else
            Call Calcular
        End If
    End If
End If
If KeyAscii = 27 Then
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
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Option1.Enabled = False Then
        Text5.SetFocus
    Else
        Command1.SetFocus
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

Private Sub Text4_LostFocus()
If Combo1.Text = "" Then
    Combo1.Text = 1
End If
If Combo1.Text = 1 Then
    Max = Setup!DES_MAX
Else
    Max = Clientes!Descuento
    If Max = 0 Then
        Max = Setup!DES_MAX
    End If
End If
TMP_DEC = 0
If (Text4.Text = "" Or Text4.Text < "0" Or Text4.Text > "9") Then
    Text4.Text = 0
    Text4.SelStart = 0
    Text4.SelLength = 3
    Label5.Caption = FormatCurrency((Label3.Caption * (Text4.Text / 100)), 2)
    Label7.Caption = FormatCurrency((CDbl(Label3.Caption) - CDbl(Label5.Caption)) * ((Setup!IMP / 100) + 1), 2)

Else
    If CInt(Text4.Text) > Max Then
        MsgBox "El descuento máximo permitido es de " & Max & " %", vbInformation
        Text4.Text = 0
        Text4.SelStart = 0
        Text4.SelLength = 3
        Text4.SetFocus
    Else
        Label5.Caption = FormatCurrency((Label3.Caption * (Text4.Text / 100)), 2)
        Label7.Caption = FormatCurrency((CDbl(Label3.Caption) - CDbl(Label5.Caption)) * ((Setup!IMP / 100) + 1), 2)
    End If
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text5.Text = 0 Then
        Text5.Text = 1
    Else
        Command1.SetFocus
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
    Combo2.SetFocus
End If
End Sub

