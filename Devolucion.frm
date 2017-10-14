VERSION 5.00
Begin VB.Form Form12 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Devolucion"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   Icon            =   "Devolucion.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "->"
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2760
      MaxLength       =   4
      TabIndex        =   25
      Top             =   1320
      Width           =   975
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
      Height          =   2610
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   2160
      Width           =   9135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   19
      Top             =   120
      Width           =   2775
      Begin VB.OptionButton Option1 
         Caption         =   "Factura 1"
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
         Left            =   3000
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Factura"
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
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar y cerrar"
      Height          =   495
      Left            =   6720
      TabIndex        =   7
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aplicar devolución"
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   5160
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   26
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   23
      Top             =   1920
      Width           =   4815
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cant."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   22
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   21
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7440
      TabIndex        =   20
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   17
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label10 
      Height          =   255
      Left            =   5640
      TabIndex        =   16
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label9 
      Height          =   255
      Left            =   5640
      TabIndex        =   15
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   5640
      TabIndex        =   14
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label7 
      Caption         =   "Total:"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Subtotal:"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Listado de articulos:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Vendedor:"
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre del Cliente"
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Codigo de Cliente:"
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Factura:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Limpiar()
List2.Clear
Label11 = 0
Label12 = 0
End Sub

Private Sub Comprobante(info)
Dim x As Printer
Printer.PaperSize = 5
For Each x In Printers
    If x.DeviceName = Setup!rec_prn Then
            Set Printer = x
            Printer.Font = "Courier New"
            Printer.FontSize = 10
    End If
Next
Detalle.Index = "Detalle"
Detalle.Seek "=", info
Printer.PaperSize = 256
Printer.Width = 12240
Printer.Height = 8155
Printer.Font = "Courier New"
Printer.FontSize = 10
Tabla.Index = "PrimaryKey"
L = 0
While L < List2.ListCount
    List2.ListIndex = L
    If List2.Selected(L) = True Then
        Dev = Dev + 1
    End If
    L = L + 1
Wend
While FDF = False And Facturas!Num_Fac = Detalle!Num_Fac And Dev <> 0
    Printer.Print Mid(Setup!Linea1 & "                                        ", 1, 40) & "                                   Fecha: " & FormatDateTime(Date, vbShortDate)
    Printer.Print Mid(Setup!linea2 & "                                        ", 1, 40) & "                                   Factura: " & Mid("            ", 1, 12 - Len(Mid(Detalle!Num_Fac, 2, Len(Detalle!Num_Fac) - 1))) & Mid(Detalle!Num_Fac, 2, Len(Detalle!Num_Fac) - 1)
    Printer.Print "Cliente :" & Facturas!Cli_cod
    Printer.Print "Nombre  :" & Facturas!Cli_Nom
    Printer.Print "Vendedor:" & Vendedores!Nombre
    Printer.Print " "
    Printer.Print "Por Concepto de los Siguientes articulos devueltos:"
    Printer.Print "   CODIGO      CANT.    ARTICULO                                        VALOR           MONTO   " '07
    I = 1
    J = 1
    L = 0
    While I < 19
        If Not Detalle.EOF Then
            If (Detalle!Num_Fac = Facturas!Num_Fac) Then 'And Detalle!Devuelto = True Then 'And l <> List2.ListCount Then
            
                While Dev <> 0 'L < List2.ListCount
                    List2.ListIndex = L
                    If List2.Selected(L) = True Then
                
                        Tabla.Seek "=", Mid(List2.Text, 1, 10)
                        While (Detalle!Num_Fac <> Facturas!Num_Fac Or Detalle!Codigo <> Tabla!Codigo)
                            Detalle.MoveNext
                        Wend
                        PRECIO_UNITARIO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!Pre_Uni, 2)))) & FormatNumber(Detalle!Pre_Uni, 2)
                        MONTO = Mid("            ", 1, (12 - Len(FormatNumber(Detalle!Pre_Uni * Detalle!Dev_Can, 2)))) & FormatNumber((Detalle!Pre_Uni * Detalle!Dev_Can), 2)
                        CANT = Mid("    ", 1, (4 - Len(Detalle!Dev_Can))) & Detalle!Dev_Can
                        ACUMULADO = ACUMULADO + (Detalle!Dev_Can * Detalle!Pre_Uni)
                        Printer.Print " " & Detalle!Codigo & "    " & CANT & "     " & Mid(Tabla!descrip & "                                        ", 1, 40) & "    " & PRECIO_UNITARIO & "    " & MONTO
                        'Detalle.MoveNext
                        I = I + 1
                        Dev = Dev - 1
                
                    End If
                    L = L + 1
                Wend
                
            End If
            If Dev = 0 Then
                Printer.Print " "
                I = I + 1
                If Not FDF Then
                    FDF = True
                End If
            End If
        End If
    Wend
    SUBTOTAL = Mid("            ", 1, 12 - Len(FormatCurrency(ACUMULADO, 2))) & FormatCurrency(ACUMULADO, 2)
    DES_TMP = ACUMULADO * (Facturas!Des / 100)
    Descuento = Mid("            ", 1, 12 - Len(FormatCurrency(DES_TMP, 2))) & FormatCurrency(DES_TMP, 2)
    IMP_TMP = (ACUMULADO - DES_TMP) * (Facturas!IMP / 100)
    IMPUESTO = Mid("            ", 1, 12 - Len(FormatCurrency(IMP_TMP, 2))) & FormatCurrency(IMP_TMP, 2)
    TOT_TMP = (ACUMULADO - DES_TMP) + IMP_TMP
    TOTAL = Mid("            ", 1, 12 - Len(FormatCurrency(TOT_TMP, 2))) & FormatCurrency(TOT_TMP, 2)
    Printer.Print "Se le acredita el Monto de "
    Printer.Print "                                                                         SUBTOTAL:  " & SUBTOTAL
    Printer.Print "                                                                         DESCUENTO: " & Descuento '32
    Printer.Print "                                            _____________________        IMPUESTO:  " & IMPUESTO '33"
    Printer.Print "                                            Firma de Autorización        TOTAL:     " & TOTAL '34
    Printer.EndDoc
Wend
End Sub
Private Sub Command1_Click()
NUM = 0
I = 0
While I < List2.ListCount
    List2.ListIndex = I
    If List2.Selected(I) = True Then
        Detalle.Index = "DetalleCodigo"
        Detalle.Seek "=", Mid(List2.Text, 1, 10)
        While Facturas!Num_Fac <> Detalle!Num_Fac
            Detalle.MoveNext
        Wend
        Detalle.Edit
        Detalle!Dev_Can = Detalle!Dev_Can + (Detalle!Cantidad - CInt(Mid(List2, 56, 4)))
        If Detalle!Cantidad = Detalle!Dev_Can Then
            Detalle!Devuelto = True
        Else
            Detalle!Dev_Can = Detalle!Cantidad
            Detalle!Devuelto = True
        End If
        Detalle.Update
    End If
    I = I + 1
Wend
Call Comprobante(Facturas!Num_Fac)
Call Command2_Click
End Sub


Private Sub Command2_Click()
Form12.Hide
Form1.Show
End Sub

Private Sub Command3_Click()
If Text2 > Label14 Then
    Text2 = 0
    Text2.SelStart = 0
    Text2.SelLength = 4
Else
    If List2.Selected(List2.ListIndex) Then
        A = Mid(List2, 1, 54)
        T = Mid(List2, 56, 4) - Text2
        B = Mid("   ", 1, 4 - Len(T)) & T
        C = Mid(List2, 59, 200)
        POSICION = List2.ListIndex
        List2.RemoveItem (POSICION)
        List2.AddItem A & B & C, POSICION
        List2.ListIndex = POSICION
        List2.Selected(POSICION) = True
    End If
End If
End Sub

Private Sub Form_Activate()
Option1.Value = True
Option2.Value = False
Text1 = ""
Text2 = 0
Label8 = ""
Label9 = ""
Label10 = ""
Label11 = 0
Label12 = 0
List2.Clear
Option2.SetFocus
End Sub

Private Sub List2_Click()
C = Mid(List2, 56, 4)
Label14 = (CInt(C))
Text2 = 0
End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
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

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Set Facturas = Inventa.OpenRecordset("FACTURAS")
    Facturas.Index = "PrimaryKey"
    If Option1.Value = True Then
        Tipo = "T"
    ElseIf Option2.Value = True Then
        Tipo = "R"
    ElseIf Option3.Value = True Then
        Tipo = "P"
    End If
    If IsNull(Tipo) Then
        MsgBox "Por favor eliga el tipo de documento", vbCritical
    Else
        Facturas.Seek "=", Tipo & Text1.Text
        If Facturas.NoMatch Then
            MsgBox "Número de factura invalido", vbCritical
        Else
            Label8.Caption = Facturas!Cli_cod
            Label9.Caption = Facturas!Cli_Nom
            Set Vendedores = Inventa.OpenRecordset("Vendedores")
            Vendedores.Index = "PrimaryKey"
            Vendedores.Seek "=", Facturas!VEN_COD
            Label10.Caption = Vendedores!Nombre
            Call Limpiar
            Call Cargar(Tipo)
            List2.SetFocus
            Call List2_Click
        End If
    End If
End If
End Sub

Private Sub Cargar(Tipo)
Set Detalle = Inventa.OpenRecordset("Detalle")
Set Tabla = Inventa.OpenRecordset("INVENTA")
Set Setup = Inventa.OpenRecordset("Setup")
List2.Clear
Detalle.Index = "Detalle"
Detalle.Seek "=", Tipo & Text1.Text
Tabla.Index = "CODIGO"
While Facturas!Num_Fac = Detalle!Num_Fac
    If Detalle!Devuelto = False Then
        Tabla.Seek "=", Detalle!Codigo
        List2.AddItem Detalle!Codigo & "   " & Mid(Tabla!descrip & "                                            ", 1, 40) & " " & Mid("    ", 1, 4 - Len(Detalle!Cantidad - Detalle!Dev_Can)) & Detalle!Cantidad - Detalle!Dev_Can & " " & Mid("            ", 1, 12 - Len(FormatCurrency(Detalle!Cantidad * Detalle!Pre_Uni, 2))) & FormatCurrency(Detalle!Cantidad * Detalle!Pre_Uni, 2) & "     " & Detalle!Dev_Can
        Label11.Caption = FormatCurrency(Label11.Caption + (Detalle!Cantidad * Detalle!Pre_Uni), 2)
        Label12.Caption = FormatCurrency(Label11.Caption + Label11.Caption * Facturas!IMP / 100, 2)
    End If
    Detalle.MoveNext
Wend
End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = 4
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command3_Click
    List2.SetFocus
End If
Dim KEY As String
KEY = Chr(KeyAscii)
If (KEY < "0" Or KEY > "9") Then
    If (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End If
End Sub
