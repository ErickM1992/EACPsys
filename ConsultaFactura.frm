VERSION 5.00
Begin VB.Form Form10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar Factura"
   ClientHeight    =   8055
   ClientLeft      =   1830
   ClientTop       =   1635
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ConsultaFactura.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   7200
      Width           =   8295
      Begin VB.CommandButton Command4 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Limpiar Pantalla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Imprimir Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Consultar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5520
      TabIndex        =   5
      Top             =   360
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3855
      Begin VB.OptionButton Option3 
         Caption         =   "Proforma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Factura 1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.ListBox List1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   10455
   End
   Begin VB.Label Label1 
      Caption         =   "Número de Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Mostrar()
FDF = False
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
    While FDF = False
        ACUMULADO = 0
        List1.AddItem ("                                                                " & MODO_PAGO)
        List1.AddItem (" ")
        List1.AddItem (" ")
        List1.AddItem ("                                                " & Day(Facturas!fecha) & "           " & Month(Facturas!fecha) & "           " & Year(Facturas!fecha))
        List1.AddItem (" ")
        List1.AddItem ("         " & Mid(Facturas!Cli_cod & "     ", 1, 5) & "                                                    " & Mid("*           ", 1, 12 - (Len(Facturas!Num_Fac) - 1)) & Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1))
        List1.AddItem ("         " & Facturas!Cli_Nom)
        List1.AddItem (" ")
        List1.AddItem (" ")
        List1.AddItem (" ")
        List1.AddItem (" ")
        I = 1
        While I < 14
            If Not Detalle.EOF Then
                If Detalle!Num_Fac = Facturas!Num_Fac Then
                    Tabla.Seek "=", Detalle!Codigo
                    PRECIO_UNITARIO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!Pre_Uni, 2)))) & FormatNumber(Detalle!Pre_Uni, 2)
                    MONTO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!Pre_Uni * Detalle!Cantidad, 2)))) & FormatNumber((Detalle!Pre_Uni * Detalle!Cantidad), 2)
                    CANT = Mid("    ", 1, (4 - Len(Detalle!Cantidad))) & Detalle!Cantidad
                    ACUMULADO = ACUMULADO + (Detalle!Cantidad * Detalle!Pre_Uni)
                    List1.AddItem (Detalle!Codigo & " " & CANT & " " & Mid(Tabla!descrip & "                                        ", 1, 32) & " " & PRECIO_UNITARIO & "     " & MONTO)
'                                           1234567891123456
                    Detalle.MoveNext
                Else
                    List1.AddItem (" ")
                End If
            Else
                List1.AddItem (" ")
            End If
            I = I + 1
            If Not Detalle.EOF Then
                If Detalle!Num_Fac <> Facturas!Num_Fac Then
                    FDF = True
                End If
            Else
                FDF = True
            End If
        Wend
        SUBTOTAL = Mid("            ", 1, 12 - Len(FormatCurrency(ACUMULADO, 2))) & FormatCurrency(ACUMULADO, 2)
        DES_TMP = ACUMULADO * (Facturas!Des / 100)
        Descuento = Mid("            ", 1, 12 - Len(FormatCurrency(DES_TMP, 2))) & FormatCurrency(DES_TMP, 2)
        IMP_TMP = (ACUMULADO - DES_TMP) * (Facturas!IMP / 100)
        IMPUESTO = Mid("            ", 1, 12 - Len(FormatCurrency(IMP_TMP, 2))) & FormatCurrency(IMP_TMP, 2)
        TOT_TMP = (ACUMULADO - DES_TMP) + IMP_TMP
        TOTAL = Mid("            ", 1, 12 - Len(FormatCurrency(TOT_TMP, 2))) & FormatCurrency(TOT_TMP, 2)
        List1.AddItem (" ")
        List1.AddItem ("                                                                  " & SUBTOTAL)
        List1.AddItem ("                                                                  " & Descuento)
        List1.AddItem ("                                                                  " & IMPUESTO)
        List1.AddItem ("                                                                  " & TOTAL)
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
    While FDF = False
        List1.AddItem (Mid(Setup!Linea1 & "                                        ", 1, 40) & "                                   Fecha: " & FormatDateTime(Facturas!fecha, vbShortDate))
        List1.AddItem (Mid(Setup!linea2 & "                                        ", 1, 40) & "                                   Factura Nº:" & Mid(" ", 1, 12 - Len(Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1))) & Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1))
        List1.AddItem ("                                                                           Modo de Pago: " & MODO_PAGO)
        List1.AddItem ("Cliente :" & Facturas!Cli_cod)
        List1.AddItem ("Nombre  :" & Facturas!Cli_Nom)
        List1.AddItem ("Vendedor:" & Vendedores!Nombre)
        List1.AddItem (" ")
        List1.AddItem ("   CODIGO      CANT.    ARTICULO                                        VALOR           MONTO   ")
        I = 1
        While I < 19
            If Not Detalle.EOF Then
                If Detalle!Num_Fac = Facturas!Num_Fac Then
                    Tabla.Seek "=", Detalle!Codigo
                    PRECIO_UNITARIO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!Pre_Uni, 2)))) & FormatNumber(Detalle!Pre_Uni, 2)
                    MONTO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!Pre_Uni * Detalle!Cantidad, 2)))) & FormatNumber((Detalle!Pre_Uni * Detalle!Cantidad), 2)
                    CANT = Mid("    ", 1, (4 - Len(Detalle!Cantidad))) & Detalle!Cantidad
                    ACUMULADO = ACUMULADO + (Detalle!Cantidad * Detalle!Pre_Uni)
                    List1.AddItem (" " & Detalle!Codigo & "    " & CANT & "     " & Mid(Tabla!descrip & "                                        ", 1, 40) & "    " & PRECIO_UNITARIO & "    " & MONTO)
                    Detalle.MoveNext
                Else
                    List1.AddItem (" ")
                End If
            Else
                List1.AddItem (" ")
            End If
            I = I + 1
            If Not Detalle.EOF Then
                If Detalle!Num_Fac <> Facturas!Num_Fac Then
                    FDF = True
                End If
            Else
                FDF = True
            End If
        Wend
        SUBTOTAL = Mid("            ", 1, 12 - Len(FormatCurrency(ACUMULADO, 2))) & FormatCurrency(ACUMULADO, 2)
        DES_TMP = ACUMULADO * (Facturas!Des / 100)
        Descuento = Mid("            ", 1, 12 - Len(FormatCurrency(DES_TMP, 2))) & FormatCurrency(DES_TMP, 2)
        IMP_TMP = (ACUMULADO - DES_TMP) * (Facturas!IMP / 100)
        IMPUESTO = Mid("            ", 1, 12 - Len(FormatCurrency(IMP_TMP, 2))) & FormatCurrency(IMP_TMP, 2)
        TOT_TMP = (ACUMULADO - DES_TMP) + IMP_TMP
        TOTAL = Mid("            ", 1, 12 - Len(FormatCurrency(TOT_TMP, 2))) & FormatCurrency(TOT_TMP, 2)
        List1.AddItem (" ")
        List1.AddItem ("                                                                         SUBTOTAL: " & SUBTOTAL)
        List1.AddItem ("                                                                         DESCUENTO:" & Descuento)
        List1.AddItem ("      *** NUESTRAS VENTAS SON EN FIRME ***    _____________________      IMPUESTO: " & IMPUESTO)
        List1.AddItem ("       *** NO SE ACEPTAN DEVOLUCIONES ***       Recibido Conforme        TOTAL:    " & TOTAL)
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
    While FDF = False
        List1.AddItem (Mid(Setup!Linea1 & "                                        ", 1, 40) & "                                  Fecha: " & FormatDateTime(Facturas!fecha, vbShortDate))
        List1.AddItem (Mid(Setup!linea2 & "                                        ", 1, 40) & "                                  PROFORMA Nº: " & Mid("    ", 1, 12 - Len(Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1))) & Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1))
        List1.AddItem ("                                                                          Modo de Pago: " & MODO_PAGO)
        List1.AddItem ("Cliente :" & Facturas!Cli_cod)
        List1.AddItem ("Nombre  :" & Facturas!Cli_Nom)
        List1.AddItem ("Vendedor:" & Vendedores!Nombre)
        List1.AddItem (" ")
        List1.AddItem ("   CODIGO      CANT.    ARTICULO                                        VALOR           MONTO   ")
        I = 1
        While I < 19
            If Not Detalle.EOF Then
                If Detalle!Num_Fac = Facturas!Num_Fac Then
                    Tabla.Seek "=", Detalle!Codigo
                    PRECIO_UNITARIO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!Pre_Uni, 2)))) & FormatNumber(Detalle!Pre_Uni, 2)
                    MONTO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!Pre_Uni * Detalle!Cantidad, 2)))) & FormatNumber((Detalle!Pre_Uni * Detalle!Cantidad), 2)
                    CANT = Mid("    ", 1, (4 - Len(Detalle!Cantidad))) & Detalle!Cantidad
                    ACUMULADO = ACUMULADO + (Detalle!Cantidad * Detalle!Pre_Uni)
                    List1.AddItem (" " & Detalle!Codigo & "    " & CANT & "     " & Mid(Tabla!descrip & "                                        ", 1, 40) & "    " & PRECIO_UNITARIO & "    " & MONTO)
                    Detalle.MoveNext
                Else
                    FDF = True
                    List1.AddItem (" ")
                End If
            Else
                List1.AddItem (" ")
            End If
            I = I + 1
            If Detalle.EOF Then
                If Detalle!Num_Fac <> Facturas!Num_Fac Then
                    FDF = True
                End If
            Else
                FDF = True
            End If
        Wend
        SUBTOTAL = Mid("            ", 1, 13 - Len(FormatCurrency(ACUMULADO, 2))) & FormatCurrency(ACUMULADO, 2)
        DES_TMP = ACUMULADO * (Facturas!Des / 100)
        Descuento = Mid("            ", 1, 13 - Len(FormatCurrency(DES_TMP, 2))) & FormatCurrency(DES_TMP, 2)
        IMP_TMP = (ACUMULADO - DES_TMP) * (Facturas!IMP / 100)
        IMPUESTO = Mid("            ", 1, 13 - Len(FormatCurrency(IMP_TMP, 2))) & FormatCurrency(IMP_TMP, 2)
        TOT_TMP = (ACUMULADO - DES_TMP) + IMP_TMP
        TOTAL = Mid("            ", 1, 13 - Len(FormatCurrency(TOT_TMP, 2))) & FormatCurrency(TOT_TMP, 2)
        List1.AddItem (" ")
        List1.AddItem ("                                                                         SUBTOTAL: " & SUBTOTAL)
        List1.AddItem ("                                                                         DESCUENTO:" & Descuento)
        List1.AddItem ("      *** NUESTRAS VENTAS SON EN FIRME ***    _____________________      IMPUESTO: " & IMPUESTO)
        List1.AddItem ("       *** NO SE ACEPTAN DEVOLUCIONES ***       Recibido Conforme        TOTAL:    " & TOTAL)
    Wend
End If
End Sub
Private Sub Imprimir()
FDF = False
Dim X As Printer
Printer.PaperSize = 5
For Each X In Printers
    If Option1.Value = True Then
        If X.DeviceName = Setup!tim_prn Then
            Set Printer = X
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
        If X.DeviceName = Setup!rec_prn Then
            Set Printer = X
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
        If X.DeviceName = Setup!pro_prn Then
            Set Printer = X
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
        Printer.Print "                                                " & Day(Facturas!fecha) & "           " & Month(Facturas!fecha) & "           " & Year(Facturas!fecha)
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
        Printer.Print Mid(Setup!Linea1 & "                                        ", 1, 40) & "                                   Fecha: " & Facturas!fecha
        Printer.Print Mid(Setup!linea2 & "                                        ", 1, 40) & "                                   Factura Nº: " & Mid("            ", 1, 12 - Len(Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1))) & Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1)
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
        Printer.Print Mid(Setup!Linea1 & "                                        ", 1, 40) & "                                  Fecha: " & Facturas!fecha
        Printer.Print Mid(Setup!linea2 & "                                        ", 1, 40) & "                                  Proforma Nº: " & Mid("            ", 1, 12 - Len(Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1))) & Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1)
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
        Printer.Print Mid(Setup!pie1 & "          ", 1, 40) & " _____________________           IMPUESTO:  " & IMPUESTO
        Printer.Print Mid(Setup!pie2 & "          ", 1, 40) & "   Recibido Conforme             TOTAL:     " & TOTAL
        Printer.EndDoc
    Wend
End If
End Sub
Private Sub Command1_Click()
List1.Clear
If Option1.Value Then
    Documento = "T"
ElseIf Option2.Value Then
    Documento = "R"
ElseIf Option3.Value Then
    Documento = "P"
End If
Factura = Documento & Text1.Text ' R o P 23, 24, 15 etc, concatena una letra con un valor de busqueda.
Facturas.Index = "PrimaryKey"
Facturas.Seek "=", Factura ' Busqueme en la posicion que optiene la variable factura,
If Facturas.NoMatch Then
    MsgBox "Factura no existe", vbCritical
Else
    Call Mostrar
End If
Option2.SetFocus
End Sub

Private Sub Command2_Click()
If Option1.Value Then
    Documento = "T"
ElseIf Option2.Value Then
    Documento = "R"
ElseIf Option3.Value Then
    Documento = "P"
End If
Factura = Documento & Text1.Text
Facturas.Index = "PrimaryKey"
Facturas.Seek "=", Factura
If Facturas.NoMatch Then
    MsgBox "Factura no existe", vbCritical
Else
    Call Imprimir
End If
Option2.SetFocus
End Sub

Private Sub Command3_Click()
List1.Clear
Option2.SetFocus
End Sub

Private Sub Command4_Click()
Option1.Value = True
Text1.Text = ""
Form10.Hide
Form1.Show
End Sub

Private Sub Form_Activate()
Option1.Value = True
Option2.Value = False
Option3.Value = False
Text1.Text = ""
List1.Clear
Option2.SetFocus
Set Setup = Inventa.OpenRecordset("SETUP")
Set Vendedores = Inventa.OpenRecordset("VENDEDORES")
Set Facturas = Inventa.OpenRecordset("FACTURAS")
Set Detalle = Inventa.OpenRecordset("DETALLE")
Set Tabla = Inventa.OpenRecordset("INVENTA")
Setup.MoveFirst
Facturas.MoveFirst
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Command4_Click
End Sub

Private Sub Option1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1.SetFocus
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
    Command1.SetFocus
End If
End Sub
