VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   2295
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Printer.PaperSize = 256
    Printer.FontName = "Draft 17cpi"
    Printer.FontSize = 10
    Printer.Width = 12240
    Printer.Height = 8155
'   Printer.Print "1234567891123456789212345678931234567894123456789512345678961234567897123456789812345678991234567890123456" '01
    Printer.Print "   " '01 ENCABEZADO
    Printer.Print " " '02 ENCABEZADO
    Printer.Print " " '03 ENCABEZADO
    Printer.Print " " '04 ENCABEZADO
    Printer.Print " " '05 ENCABEZADO
    Printer.Print " " '06 ENCABEZADO
    Printer.Print "     CODIGO        CANT.       ARTICULO                                        VALOR           MONTO   " '07
    Printer.Print "   1234567890      1234        1234567890123456789012345678901234567890    123456789012    123456789012" '08
    Printer.Print "12345678911234567892123456789312345678941234567895123456789612345678971234567898123456789912345678901234567891123456789212345678931234567894" '09
    Printer.Print "   " '10
    Printer.Print "   " '11
    Printer.Print "   " '12
    Printer.Print "   " '13
    Printer.Print "   " '14
    Printer.Print "   " '15
    Printer.Print "   " '16
    Printer.Print "   " '17
    Printer.Print "   " '18
    Printer.Print "   " '19
    Printer.Print "   " '20
    Printer.Print "   " '21
    Printer.Print "   " '22
    Printer.Print "   " '23
    Printer.Print "   " '24
    Printer.Print "   " '25
    Printer.Print "   " '26
    Printer.Print "   " '27
    Printer.Print "   " '28
    Printer.Print "   " '29
    Printer.Print "   " '30
    Printer.Print "   " '31
    Printer.Print "                                                                              SUBTOTAL:  123456789012" '32
    Printer.Print "      *** NUESTRAS VENTAS SON EN FIRME ***                                    DESCUENTO: 123456789012" '33
    Printer.Print "       *** NO SE ACEPTAN DEVOLUCIONES ***                                     TOTAL:     123456789012" '34
    Printer.EndDoc
End Sub
