VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   Icon            =   "Facturacion.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   3840
      TabIndex        =   44
      Top             =   3000
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
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
      Height          =   315
      Left            =   1200
      TabIndex        =   42
      Text            =   "1"
      Top             =   4800
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   8880
      Top             =   840
   End
   Begin VB.ComboBox Text6 
      Height          =   315
      Left            =   1200
      TabIndex        =   40
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Frame Frame3 
      Height          =   495
      Left            =   2160
      TabIndex        =   37
      Top             =   2640
      Width           =   1335
      Begin VB.CheckBox Check1 
         Caption         =   "Exento"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   32
      Top             =   3600
      Width           =   9135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7320
      TabIndex        =   31
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox Text7 
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
      Height          =   315
      Left            =   5880
      TabIndex        =   23
      Top             =   1320
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Borrar articulos marcados"
      Height          =   375
      Left            =   4200
      TabIndex        =   22
      Top             =   4800
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar e Imprimir"
      Height          =   375
      Left            =   1800
      TabIndex        =   21
      Top             =   4800
      Width           =   2295
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
      Height          =   315
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   14
      Text            =   "0"
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
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
      Height          =   315
      Left            =   1320
      TabIndex        =   13
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text1 
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
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   11
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text2 
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
      TabIndex        =   10
      Top             =   1800
      Width           =   3735
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   9
      Top             =   2280
      Width           =   5415
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
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
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
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Modo de pago"
      Height          =   735
      Left            =   5400
      TabIndex        =   1
      Top             =   120
      Width           =   3855
      Begin VB.OptionButton Option5 
         Caption         =   "Credito"
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Contado"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Documento"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.OptionButton Option3 
         Caption         =   "Proforma"
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Factura"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Factura 1"
         Height          =   255
         Left            =   4560
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label21 
      Caption         =   "Vendedor"
      Height          =   255
      Left            =   4680
      TabIndex        =   43
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   "Copias"
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3120
      TabIndex        =   39
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   315
      Left            =   7440
      TabIndex        =   36
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unitario"
      Height          =   315
      Left            =   5880
      TabIndex        =   35
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cant."
      Height          =   315
      Left            =   5160
      TabIndex        =   34
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripción"
      Height          =   315
      Left            =   360
      TabIndex        =   33
      Top             =   3360
      Width           =   4815
   End
   Begin VB.Label Label15 
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
      Height          =   255
      Left            =   8160
      TabIndex        =   30
      Top             =   -1200
      Width           =   3615
   End
   Begin VB.Label Label14 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label13 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Código:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "Busqueda por parte:"
      Height          =   255
      Left            =   3840
      TabIndex        =   26
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "de"
      Height          =   255
      Left            =   2160
      TabIndex        =   25
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
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
      Height          =   315
      Left            =   2640
      TabIndex        =   24
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "0,00"
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
      Left            =   6720
      TabIndex        =   20
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   19
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "0,00"
      Height          =   255
      Left            =   2040
      TabIndex        =   18
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Descuento:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "0,00"
      Height          =   255
      Left            =   2040
      TabIndex        =   16
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Subtotal:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Descuento"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Calcular()
UNITARIO = Mid("            ", 1, 12 - Len(CStr(FormatNumber(Tabla!P_Venta, 2)))) & CStr(FormatNumber(Tabla!P_Venta, 2))
MONTO = Mid("            ", 1, 12 - Len(FormatNumber(UNITARIO * Text3.Text, 2))) & FormatNumber(UNITARIO * Text3.Text, 2)
Label3.Caption = FormatCurrency((FormatNumber(MONTO, 2) + CDbl(Label3.Caption)))
If Len(Text3.Text) = 1 Then
    Text3.Text = "   " & Text3.Text
ElseIf Len(Text3.Text) = 2 Then
    Text3.Text = "  " & Text3.Text
ElseIf Len(Text3.Text) = 3 Then
    Text3.Text = " " & Text3.Text
End If
List2.AddItem Mid(Tabla!descrip & "                                            ", 1, 40) & " " & Text3.Text & " " & UNITARIO & " " & MONTO & "   " & Tabla!Codigo
Text1.SetFocus
End Sub

Private Sub Facturar()

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

If Combo1.Text = "000001" Then
    Facturas!Cli_cod = Combo1.Text
    Facturas!Cli_Nom = Text6.Text
Else
    Facturas!Cli_cod = Clientes!Codigo
    Facturas!Cli_Nom = Clientes!Cliente
End If

If Option4.Value = True Then
    Facturas!Cancelado = True
    Facturas!Monto_F = Label7.Caption
    Facturas!Monto_S = 0
ElseIf Option5.Value = True Then
    Facturas!Cancelado = False
    Facturas!Monto_F = Label7.Caption
    Facturas!Monto_S = Label7.Caption
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
fac_num = Facturas!Num_Fac
Facturas.Update
Tabla.Index = "CODIGO"
I = 0



While I < List2.ListCount
    List2.ListIndex = I
    Detalle!Num_Fac = fac_num
    Detalle!Cantidad = Mid(List2.Text, 42, 4)
    Detalle!Codigo = Mid(List2.Text, 75, 10)
    Detalle!Pre_Uni = Mid(List2.Text, 47, 12)
    If Not Option3.Value Then
        Tabla.Seek "=", Mid(List2.Text, 75, 10)
        Tabla.Edit
        Tabla!Cantidad = Tabla!Cantidad - Detalle!Cantidad
        Tabla.Update
    End If
    Detalle.Update
    Detalle.AddNew
    I = I + 1
Wend
Facturas.Index = "PrimaryKey"
Facturas.Seek "=", fac_num
Call Imprimir
Call Limpiar
Form3.Hide
Form1.Show
End Sub
Private Sub Limpiar()
Check1.Value = 0
Option1.Value = False
Option2.Value = True
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
Combo1.Text = ""
Combo2.Text = ""
List1.Clear
List2.Clear
Label3.Caption = "0,00"
Label5.Caption = "0,00"
Label7.Caption = "0,00"
Label10.Caption = ""
Label20.Caption = ""
End Sub

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
            Printer.PaperSize = 1
            Ancho = Printer.Width
            Alto = Printer.Height / 2
            Printer.PaperSize = 256
            Printer.Width = Ancho
            Printer.Height = Alto
            Printer.Font = "Courier New"
            Printer.FontSize = 10
        End If
    ElseIf Option2.Value = True Then
        If x.DeviceName = Setup!rec_prn Then
            Set Printer = x
            Printer.PaperSize = 1
            Ancho = 12240
            Alto = 8155
            Printer.PaperSize = 256
            Printer.Width = Ancho
            Printer.Height = Alto
            Printer.Font = "Courier New"
            Printer.FontSize = 10
        End If
    ElseIf Option3.Value = True Then
        If x.DeviceName = Setup!pro_prn Then
            Set Printer = x
            Printer.PaperSize = 1
            Ancho = 12240
            Alto = 8155
            Printer.PaperSize = 256
            Printer.Width = Ancho
            Printer.Height = Alto
            Printer.Font = "Courier New"
            Printer.FontSize = 10
        End If
    End If
Next
SUBTOTAL = FormatCurrency(0, 2)
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
    While FDF = False 'And Detalle!Num_Fac = Facturas!Num_Fac
        ACUMULADO = 0
        Printer.Print "                                                                " & MODO_PAGO
        Printer.Print " "
        Printer.Print " "
        Printer.Print "                                                " & Day(Date) & "           " & Month(Date) & "           " & Year(Date)
        Printer.Print " "
        Printer.Print "         " & Mid(Facturas!Cli_cod & "     ", 1, 5) & "                                                    " & Mid("*           ", 1, 12 - (Len(Facturas!Num_Fac) - 1)) & Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1)
        Printer.Print "         " & Facturas!Cli_Nom
        Printer.Print " "
        Printer.Print " "
        Printer.Print " "
        Printer.Print " "
        I = 1
        While I < 14
            If Not Detalle.EOF Then
                If Detalle!Num_Fac = Facturas!Num_Fac Then
                    Tabla.Seek "=", Detalle!Codigo
                    PRECIO_UNITARIO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!Pre_Uni, 2)))) & FormatNumber(Detalle!Pre_Uni, 2)
                    MONTO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!Pre_Uni * Detalle!Cantidad, 2)))) & FormatNumber((Detalle!Pre_Uni * Detalle!Cantidad), 2)
                    CANT = Mid("    ", 1, (4 - Len(Detalle!Cantidad))) & Detalle!Cantidad
                    ACUMULADO = ACUMULADO + (Detalle!Cantidad * Detalle!Pre_Uni)
                    Printer.Print Detalle!Codigo & " " & CANT & " " & Mid(Tabla!descrip & "                                        ", 1, 32) & " " & PRECIO_UNITARIO & "     " & MONTO
                    Detalle.MoveNext
                Else
                    FDF = True
                    Printer.Print " "
                End If
            Else
                FDF = True
                Printer.Print " "
            End If
            I = I + 1
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
    While FDF = False And Facturas!Num_Fac = Detalle!Num_Fac
        Printer.Print Mid(Setup!Linea1 & "                                        ", 1, 40) & "                                   Fecha: " & FormatDateTime(Date, vbShortDate)
        Printer.Print Mid(Setup!linea2 & "                                        ", 1, 40) & "                                   Proforma:" & Mid("            ", 1, 12 - Len(Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1))) & Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1)
        Printer.Print "                                                                           Modo de Pago: " & MODO_PAGO
        Printer.Print "Cliente :" & Facturas!Cli_cod
        Printer.Print "Nombre  :" & Facturas!Cli_Nom
        Printer.Print "Vendedor:" & Vendedores!Nombre
        Printer.Print " "
        Printer.Print "   CODIGO      CANT.    ARTICULO                                        VALOR           MONTO   " '07
        I = 1
        While I < 18
            If Not Detalle.EOF Then
                If Detalle!Num_Fac = Facturas!Num_Fac Then
                    Tabla.Seek "=", Detalle!Codigo
                    PRECIO_UNITARIO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!Pre_Uni, 2)))) & FormatNumber(Detalle!Pre_Uni, 2)
                    MONTO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!Pre_Uni * Detalle!Cantidad, 2)))) & FormatNumber((Detalle!Pre_Uni * Detalle!Cantidad), 2)
                    CANT = Mid("    ", 1, (4 - Len(Detalle!Cantidad))) & Detalle!Cantidad
                    ACUMULADO = ACUMULADO + (Detalle!Cantidad * Detalle!Pre_Uni)
                    Printer.Print " " & Detalle!Codigo & "    " & CANT & "     " & Mid(Tabla!descrip & "                                        ", 1, 40) & "    " & PRECIO_UNITARIO & "    " & MONTO
                    Detalle.MoveNext
                Else
                    FDF = True
                    Printer.Print " "
                End If
            Else
                Printer.Print " "
            End If
            I = I + 1
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
    While FDF = False And Detalle!Num_Fac = Facturas!Num_Fac
        Printer.Print Mid(Setup!Linea1 & "                                        ", 1, 40) & "                                  Fecha: " & FormatDateTime(Date, vbShortDate)
        Printer.Print Mid(Setup!linea2 & "                                        ", 1, 40) & "                                  PROFORMA: " & Mid("            ", 1, 12 - Len(Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1))) & Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1)
        Printer.Print "                                                                          Modo de Pago: " & MODO_PAGO
        Printer.Print "Cliente :" & Facturas!Cli_cod
        Printer.Print "Nombre  :" & Facturas!Cli_Nom
        Printer.Print "Vendedor:" & Vendedores!Nombre
        Printer.Print " "
        Printer.Print "   CODIGO      CANT.    ARTICULO                                        VALOR           MONTO   " '07
        I = 1
        While I < 18
            If Not Detalle.EOF Then
                If Detalle!Num_Fac = Facturas!Num_Fac Then
                    Tabla.Seek "=", Detalle!Codigo
                    PRECIO_UNITARIO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!Pre_Uni, 2)))) & FormatNumber(Detalle!Pre_Uni, 2)
                    MONTO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!Pre_Uni * Detalle!Cantidad, 2)))) & FormatNumber((Detalle!Pre_Uni * Detalle!Cantidad), 2)
                    CANT = Mid("    ", 1, (4 - Len(Detalle!Cantidad))) & Detalle!Cantidad
                    ACUMULADO = ACUMULADO + (Detalle!Cantidad * Detalle!Pre_Uni)
                    Printer.Print " " & Detalle!Codigo & "    " & CANT & "     " & Mid(Tabla!descrip & "                                        ", 1, 40) & "    " & PRECIO_UNITARIO & "    " & MONTO
                    Detalle.MoveNext
                Else
                    FDF = True
                    Printer.Print " "
                End If
            Else
                Printer.Print " "
            End If
            I = I + 1
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
        Printer.Print "                                                                         DESCUENTO: " & Descuento
        Printer.Print Mid(Setup!pie1 & "                                        ", 1, 40) & " _____________________           IMPUESTO:  " & IMPUESTO
        Printer.Print Mid(Setup!pie2 & "                                        ", 1, 40) & "   Recibido Conforme             TOTAL:     " & TOTAL
        Printer.EndDoc
    Wend
End If
J = J + 1
ACUMULADO = 0
Wend
End Sub
Private Sub Buscar()
    List1.Clear
    If Len(Text1.Text) = 1 Then
        Text1.Text = "000000000" & Text1.Text
    ElseIf Len(Text1.Text) = 2 Then
        Text1.Text = "00000000" & Text1.Text
    ElseIf Len(Text1.Text) = 3 Then
        Text1.Text = "0000000" & Text1.Text
    ElseIf Len(Text1.Text) = 4 Then
        Text1.Text = "000000" & Text1.Text
    ElseIf Len(Text1.Text) = 5 Then
        Text1.Text = "00000" & Text1.Text
    ElseIf Len(Text1.Text) = 6 Then
        Text1.Text = "0000" & Text1.Text
    ElseIf Len(Text1.Text) = 7 Then
        Text1.Text = "000" & Text1.Text
    ElseIf Len(Text1.Text) = 8 Then
        Text1.Text = "00" & Text1.Text
    ElseIf Len(Text1.Text) = 9 Then
        Text1.Text = "0" & Text1.Text
    End If

If Text1.Text = "" And Text2.Text = "" Then
    MsgBox "Por favor digite algún dato"
    Text1.SetFocus
ElseIf Text1.Text = "" And Text2.Text <> "" Then
    cadena = UCase(Text2.Text)
    L = Len(cadena)
    Tabla.Index = "DATOS"
    Tabla.MoveFirst
    While Tabla.EOF = False
        If Mid(Tabla!descrip, 1, L) = cadena Then
            List1.AddItem Mid(Tabla!descrip & "                                           ", 1, 43) & Tabla!Codigo
            ProgressBar1.Value = Tabla.PercentPosition
            Tabla.MoveNext
        Else
            Tabla.MoveNext
        End If
    Wend
    ProgressBar1.Value = 100
    If List1.ListCount > 0 Then
        List1.SetFocus
        List1.ListIndex = 0
    Else
        Text2.SelStart = 0
        Text2.SelLength = 40
        Text2.SetFocus
    End If
Else
    Tabla.Index = "CODIGO"
    Tabla.Seek "=", Text1.Text
    If Tabla.NoMatch = False Then
        Label10.Caption = Tabla!Cantidad
        List1.AddItem Mid(Tabla!descrip & "                                             ", 1, 45) & Tabla!Codigo
        List1.SetFocus
        List1.ListIndex = List1.ListIndex
        Text3.Enabled = True
        Text3.SetFocus
    Else
        MsgBox "Codigo no existe"
        Text1.SelStart = 0
        Text1.SelLength = 10
        Text1.SetFocus
    End If
End If
End Sub

Private Sub Check1_Click()
Label5.Caption = FormatCurrency((Label3.Caption * (Text4.Text / 100)), 2)
If Check1.Value = 0 Then
     Label7.Caption = FormatCurrency((CDbl(Label3.Caption) - CDbl(Label5.Caption)) * ((Setup!IMP / 100) + 1), 2)
Else
     Label7.Caption = FormatCurrency((CDbl(Label3.Caption) - CDbl(Label5.Caption)), 2)
End If
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Or KeyAscii = 13 Then
    If Not Option1.Value Then
        Text5.SelStart = 0
        Text5.SelLength = 40
        Text5.SetFocus
    Else
        Command1.SetFocus
    End If
End If
End Sub

Private Sub Combo1_Click()
If Len(Combo1.Text) = 0 Then
ElseIf IsNumeric(Combo1.Text) Then
    Clientes.Index = "PrimaryKey"
    Clientes.Seek "=", FormatNumber(Combo1.Text)
    If Not Clientes.NoMatch Then
        Text6.Text = Clientes!Cliente
        Text4.Text = Clientes!Descuento
    End If
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo1.Text = Mid("000000", 1, 6 - Len(Combo1.Text)) & Combo1.Text
    Call Combo1_Click
    If Combo1.Text = "000001" Then
        Text6.SelStart = 0
        Text6.SelLength = 50
        Text6.Enabled = True
        Text6.SetFocus
    Else
        Option4.SetFocus
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

Private Sub Combo2_Click()
If Len(Combo2.Text) = 0 Then
ElseIf IsNumeric(Combo2.Text) Then
    Vendedores.Index = "PrimaryKey"
    Vendedores.Seek "=", Combo2.Text
    If Vendedores.NoMatch = False Then
        Text7.Text = Vendedores!Nombre
    End If
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Combo2_Click
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
Clientes.Index = "PrimaryKey"
Clientes.Seek "=", Combo1.Text
Vendedores.Index = "PrimaryKey"
Vendedores.Seek "=", Combo2.Text
Call Check1_Click
If Len(Combo1.Text) < 6 Then
    MsgBox "Por favor especifique el cliente", vbCritical
ElseIf Not IsNumeric(Combo1.Text) Then
    MsgBox "Por favor especifique un cliente valido", vbCritical
ElseIf Clientes.NoMatch Or (Len(Text6) = 0) Then
    MsgBox "Por favor especifique un cliente valido", vbCritical
ElseIf Len(Combo2.Text) = 0 Then
    MsgBox "Por favor especifiquese como vendedor", vbCritical
ElseIf Not IsNumeric(Combo2.Text) Then
    MsgBox "Por favor especifique un vendedor valido", vbCritical
ElseIf Vendedores.NoMatch Or (Len(Text7) = 0) Then
    MsgBox "Por favor especifique un vendedor valido", vbCritical
ElseIf List2.ListCount = 0 Then
    MsgBox "No se ha facturado ningun articulo", vbCritical
ElseIf Len(Text5.Text) = 0 Then
    MsgBox "La cantidad de copias a imprimir es invalida", vbCritical
ElseIf Not IsNumeric(Text5.Text) Then
    MsgBox "La cantidad de copias a imprimir es invalida", vbCritical
ElseIf Not List2.ListCount < 11 Then
    MsgBox "Error: El maximo de artículos aceptado es de 11"
Else
    Call Facturar
End If
End Sub

Private Sub Command2_Click()
NUM = 0
I = 0
While I < List2.ListCount
List2.ListIndex = I
If List2.Selected(I) = True Then
    NUM = NUM + Mid(List2.Text, 60, 12)
    List2.RemoveItem (I)
Else
    I = I + 1
End If
Wend
Label3.Caption = FormatCurrency(FormatNumber(Label3.Caption, 2) - NUM)
End Sub

Private Sub Command3_Click()
Call Limpiar
Form3.Hide
Form1.Show
End Sub

Private Sub Form_Activate()
Set Setup = Inventa.OpenRecordset("SETUP")
Set Tabla = Inventa.OpenRecordset("INVENTA")
Set Vendedores = Inventa.OpenRecordset("VENDEDORES")
Set Facturas = Inventa.OpenRecordset("FACTURAS")
Set Detalle = Inventa.OpenRecordset("DETALLE")
Combo1.Clear
Combo2.Clear
Text6.Clear
Vendedores.Index = "PrimaryKey"
Vendedores.MoveFirst
While Vendedores.EOF = False
    Combo2.AddItem Vendedores!Codigo
    Vendedores.MoveNext
Wend
Set Clientes = Inventa.OpenRecordset("CLIENTES")
Clientes.Index = "Cliente"
Clientes.MoveFirst
'Combo1.AddItem ("000001")
While Clientes.EOF = False
    Combo1.AddItem Mid("000000", 1, 6 - Len(Clientes!Codigo)) & Clientes!Codigo
    Text6.AddItem Clientes!Cliente
    Clientes.MoveNext
Wend
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
ProgressBar1.Value = 100
Option2.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Command3_Click
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1.SetFocus
End If
End Sub

Private Sub Option1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo1.SetFocus
End If
End Sub
Private Sub Option2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo1.SetFocus
End If
End Sub
Private Sub Option3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo1.SetFocus
End If
End Sub

Private Sub Option4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo2.SetFocus
End If
End Sub

Private Sub Option5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
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
    Label10.Caption = Tabla!Cantidad
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
I = 0
B = False
If KeyAscii = 13 Then
    If Text1.Text <> "" Then
        If List2.ListCount > 0 Then
            List2.ListIndex = I
            While List2.ListCount <> I And Not B
                List2.ListIndex = I
                If Mid(List2.Text, 75, 10) = (Mid("0000000000", 1, 10 - Len(Text1)) & Text1) Then
                    B = True
                    List2.ListIndex = -1
                    Text1.SelStart = 0
                    Text1.SelLength = 10
                    MsgBox "El acticulo ya está facturado", vbExclamation
                End If
                I = I + 1
            Wend
        End If
        If B = False Then
            Call Buscar
            List2.ListIndex = -1
        End If
    Else
        Text2.SetFocus
    End If
ElseIf KeyAscii = 27 Then
    Call Text4_LostFocus
    Text4.SelStart = 0
    Text4.SelLength = 3
    Text4.SetFocus
End If
End Sub

Private Sub Text2_GotFocus()
Text1.Text = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Buscar
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text3.Text <> "" Then
    If IsNumeric(Text3.Text) = True Then
        If Option3 Or (0 < CDbl(Text3.Text) And CDbl(Text3.Text) <= CDbl(Label10.Caption)) Then
            Call Calcular
        Else
            MsgBox "Cantidad invalida"
            Text3.SelStart = 0
            Text3.SelLength = 40
            Text3.SetFocus
            If Label10.Caption = 0 Then
                Text1.SetFocus
                Text1.SelStart = 0
                Text1.SelLength = 10
            End If
        End If
    Else
        MsgBox "Por favor digite una cantidad"
        Text3.Text = ""
        Text3.SetFocus
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

Private Sub Text4_Change()
Call Text4_LostFocus
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Check1.SetFocus
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
If Len(Combo1) <> 0 Then
    Clientes.Seek "=", Combo1
    If Not Clientes.NoMatch Then
        Max = Clientes!Descuento
        If Max = 0 Then
            Max = Setup!DES_MAX
        End If
    Else
        Max = Setup!DES_MAX
    End If
    TMP_DEC = 0
    If (Text4.Text = "" Or Text4.Text < "0" Or Text4.Text > "9") Then
        Text4.Text = 0
        Text4.SelStart = 0
        Text4.SelLength = 3
        Label5.Caption = FormatCurrency((Label3.Caption * (Text4.Text / 100)), 2)
        If Not Check1.Value Then
            Label7.Caption = FormatCurrency((CDbl(Label3.Caption) - CDbl(Label5.Caption)) * ((Setup!IMP / 100) + 1), 2)
        Else
            Label7.Caption = FormatCurrency((CDbl(Label3.Caption) - CDbl(Label5.Caption)), 2)
        End If
    Else
        If CInt(Text4.Text) > Max Then
            MsgBox "El descuento máximo permitido es de " & Max & " %", vbInformation
            Text4.Text = 0
            Text4.SelStart = 0
            Text4.SelLength = 3
            Text4.SetFocus
        Else
            Label5.Caption = FormatCurrency((Label3.Caption * (Text4.Text / 100)), 2)
            If Not Check1.Value Then
                Label7.Caption = FormatCurrency((CDbl(Label3.Caption) - CDbl(Label5.Caption)) * ((Setup!IMP / 100) + 1), 2)
            Else
                Label7.Caption = FormatCurrency((CDbl(Label3.Caption) - CDbl(Label5.Caption)), 2)
            End If
        End If
    End If
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text5.Text < 1 Then
        MsgBox "Cantidad de páginas invalida"
        Text5.SelStart = 0
        Text5.SelLength = 4
        Text5.SetFocus
    Else
        If Command1.Enabled Then
            Command1.SetFocus
        End If
    End If
End If
Dim KEY As String
KEY = Chr(KeyAscii)
If (KEY < "1" Or KEY > "9") Then
    If (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text6_Click()
Clientes.Index = "Cliente"
Clientes.Seek "=", Text6.Text
If Not Clientes.NoMatch Then
    Combo1.Text = Mid("000000", 1, 6 - Len(Clientes!Codigo)) & Clientes!Codigo
    Combo1.SetFocus
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Text6.Text) = 0 Then
        'Clientes.Seek "=", "1"
        Text6.Text = Clientes!Cliente
    End If
    Option4.SetFocus
End If
End Sub

Private Sub Timer2_Timer()
If Combo1.Text = "000001" Then 'And Len(Combo1.Text) > 0 Then
    If Option5.Enabled Then
        Option5.Enabled = False
        Option4.Value = True
        Text6.Enabled = True
    End If
    If Not Text6.Enabled Then
        Text6.Enabled = True
    End If
Else
    If Not Option5.Enabled Then
        Option5.Enabled = True
    End If
    If Text6.Enabled Then
        Text6.Enabled = False
    End If
End If
If Option1.Value Then
    If Text5.Enabled Then
        Text5.Enabled = False
        Text5.Text = 1
    End If
Else
    If Not Text5.Enabled Then
        Text5.Enabled = True
    End If
End If
If List2.ListCount = 0 Then
    If Command1.Enabled Then
        Command1.Enabled = False
    End If
Else
    If Not Command1.Enabled Then
        Command1.Enabled = True
    End If
End If
If List2.ListCount > 0 Then
    If Option1.Enabled And Option3 Then
        Option1.Enabled = False
        Option2.Enabled = False
    End If
Else
    If Not Option1.Enabled And Option3 Then
        Option1.Enabled = True
        Option2.Enabled = True
    End If
End If
End Sub
