VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abono a Facturas"
   ClientHeight    =   4095
   ClientLeft      =   3060
   ClientTop       =   1800
   ClientWidth     =   9495
   Icon            =   "FacturasAbono.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Cerrar"
      Height          =   255
      Left            =   6000
      TabIndex        =   24
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Modificar"
      Height          =   975
      Left            =   4200
      TabIndex        =   18
      Top             =   2640
      Width           =   5175
      Begin VB.CommandButton Command3 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Restaurar"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   3480
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2280
      MaxLength       =   7
      TabIndex        =   17
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecciones las Facturas a Abonar"
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   9255
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
         Height          =   1410
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   9015
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Abono"
         Height          =   255
         Left            =   7200
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto cancelado"
         Height          =   255
         Left            =   5520
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto facturado"
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Factura"
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   23
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Saldo"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Monto a abonar"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Tipo de Factura"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Numero de Factura"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Limpiar()
Text1.Text = ""
Text3.Text = 0
Combo1.Text = ""
List1.Clear
Label5.Caption = FormatCurrency(0, 2)
Label8.Caption = FormatCurrency(0, 2)
Label10.Caption = FormatCurrency(0, 2)
Text1.SetFocus
Command1.Enabled = False
End Sub
Private Sub Buscar()
List1.Clear
Clientes.Index = "Cliente"
Clientes.Seek "=", Combo1.Text
If Not Clientes.NoMatch Then
    TMP = Clientes!Codigo
    Facturas.Index = "FechaBarrido"
    Facturas.MoveFirst
    'Facturas.Seek "=", Clientes!Codigo, Date
    While Not Facturas.EOF And Facturas.NoMatch = False
        fecha = Facturas!fecha
        If Clientes!Codigo = Facturas!Cli_cod Then
            If Facturas!Cancelado = False And Facturas!Anulada = False Then
                If Mid(Facturas!Num_Fac, 1, 1) = "T" Then
                    Documento = "Timbrada"
                ElseIf Mid(Facturas!Num_Fac, 1, 1) = "R" Then
                    Documento = "Factura "
                ElseIf Mid(Facturas!Num_Fac, 1, 1) = "A" Then
                    Documento = "Añadida "
                End If
                If Mid(Facturas!Num_Fac, 1, 1) <> "P" Then
                    List1.AddItem (Facturas!fecha & " " & Documento & " " & Mid("            ", 1, 12 - (Len(Facturas!Num_Fac) - 1)) & Mid(Facturas!Num_Fac, 2, Len(Facturas!Num_Fac) - 1) & " " & Mid("             ", 1, 13 - Len(FormatCurrency(Facturas!Monto_F, 2))) & FormatCurrency(Facturas!Monto_F, 2) & " " & Mid("             ", 1, 13 - Len(FormatCurrency(Facturas!Monto_F - Facturas!Monto_S, 2))) & FormatCurrency(Facturas!Monto_F - Facturas!Monto_S, 2)) & " " & Mid("             ", 1, 13 - Len(FormatCurrency(0, 2))) & FormatCurrency(0, 2)
                End If
            End If
        End If
        Facturas.MoveNext
    Wend
    If List1.ListCount > 0 Then
        List1.SetFocus
    End If
End If
End Sub
Private Sub Combo1_Change()
Call Buscar
End Sub

Private Sub Combo1_Click()
Call Buscar
End Sub

Private Sub Command1_Click()
I = 0
Setup.MoveFirst
Recibos.AddNew
Recibos!Recibo = Setup!Cue_Con
Recibos!Cliente = Clientes!Codigo
Recibos_ACTUAL = Recibos!Recibo
Recibos!fecha = Date
Recibos.Update
Setup.Edit
Setup!Cue_Con = Setup!Cue_Con + 1
Setup.Update
While I < List1.ListCount
    List1.ListIndex = I
    If Mid(List1.Text, 76, 1) = "*" Then
        Rec_Detalle.AddNew
        Rec_Detalle!Recibo = Recibos_ACTUAL
        LISTA_TIP = Mid(List1.Text, 10, 8)
        LISTA_FAC = Mid(List1.Text, 19, 12)
        LISTA_MDF = Mid(List1.Text, 32, 13)
        LISTA_CAN = Mid(List1.Text, 46, 13)
        LISTA_SAL = Label10.Caption
        LISTA_ABO = Mid(List1.Text, 60, 13)
        
        If LISTA_TIP = "Timbrada" Then
            Tipo = "T"
        ElseIf LISTA_TIP = "Factura " Then
            Tipo = "R"
        ElseIf LISTA_TIP = "Añadida " Then
            Tipo = "A"
        End If
        Rec_Detalle!Factura = Tipo & CDbl(LISTA_FAC)
        Facturas.Index = "PrimaryKey"
        Facturas.Seek "=", Rec_Detalle!Factura
        Facturas.Edit
        Facturas!Monto_S = Facturas!Monto_F - LISTA_CAN
        Facturas!FECHA_ABONO = Date
        If LISTA_SAL = "¢0,00" Then
            Facturas!Cancelado = True
        End If
        Facturas.Update
        Rec_Detalle!Facturado = LISTA_MDF
        Rec_Detalle!Saldo = LISTA_SAL
        Rec_Detalle!abonado = LISTA_ABO
        Rec_Detalle.Update
    End If
    I = I + 1
Wend
Call Limpiar
Call Buscar
End Sub

Private Sub Command2_Click()
If Mid(List1.Text, 76, 1) = "*" Then
    POSICION = List1.ListIndex
    info = List1.Text
    Cancelado = FormatNumber(Mid(List1.Text, 46, 13), 2)
    abono = FormatNumber(Mid("             ", 1, 13 - Len(FormatCurrency(Mid(List1.Text, 60, 13), 2))) & FormatCurrency(Mid(List1.Text, 60, 13), 2), 2)
    Cancelado = CDbl(Cancelado) - CDbl(abono)
    abono = FormatNumber(Mid("             ", 1, 13 - Len(FormatCurrency(0, 2))) & FormatCurrency(0, 2), 2)
    List1.RemoveItem (POSICION)
    List1.AddItem (Mid(info, 1, 45) & Mid("             ", 1, 13 - Len(FormatCurrency(Cancelado))) & FormatCurrency(Cancelado) & " " & Mid("             ", 1, 13 - Len(FormatCurrency(abono))) & FormatCurrency(abono) & "   ")
    List1.ListIndex = POSICION
End If
End Sub

Private Sub Command3_Click()
Dim x As Printer
For Each x In Printers
    If x.DeviceName = Setup!pro_prn Then
        Set Printer = x
        Printer.Font = "Courier New"
        Printer.FontSize = 10
        Printer.PaperSize = 1
    End If
Next
FDF = False
List1.ListIndex = 0
FAC_ACU = 0
CAN_FAC = 0
SAL_ACU = 0
While FDF = False And List1.ListIndex <> List1.ListCount
    If Not Command1.Enabled Then
        Command1.Enabled = True
    End If
    Printer.Print Mid(Setup!Linea1 & "                                        ", 1, 40) & "                                   Fecha: " & FormatDateTime(Date, vbShortDate)
    Printer.Print Mid(Setup!linea2 & "                                        ", 1, 40) & "                                   Recibo: " & Mid("            ", 1, 12 - Len(Setup!Cue_Con - 1)) & Setup!Cue_Con
    Printer.Print "Cliente :" & Text1.Text
    Printer.Print "Nombre  :" & Combo1.Text
    Printer.Print " "
    Printer.Print "Recibo de abono a facturas:"
                  '12-12-12 factura 1             1234567890123 1234567890123 1234567890123 1234567890123
    Printer.Print "FECHA    TIPO      FACTURA     FACTURADO     CANCELADO     ABONADO       SALDO" '07
    I = 1
    While I < 55
        If FDF = False Then
            Facturado = FormatNumber(Mid(List1.Text, 33, 13), 2)
            Cancelado = FormatNumber(Mid(List1.Text, 46, 13), 2)
            Saldo = FormatCurrency(Facturado - Cancelado, 2)
            Printer.Print Mid(List1.Text, 1, 73) & " " & Mid("               ", 1, 13 - Len(Saldo)) & Saldo
            FAC_ACU = FAC_ACU + FormatNumber(Facturado, 2)
            CAN_ACU = CAN_ACU + FormatNumber(Cancelado, 2)
            SAL_ACU = SAL_ACU + FormatNumber(Saldo, 2)
            If List1.ListIndex <> (List1.ListCount - 1) Then
                List1.ListIndex = List1.ListIndex + 1
            Else
                FDF = True
            End If
        Else
            Printer.Print " "
        End If
        I = I + 1
    Wend
    Printer.Print "                                                                         Facturado: " & FormatCurrency(FormatNumber(FAC_ACU, 2), 2)
    Printer.Print "                                                                         Cancelado: " '& FormatCurrency(FormatNumber(CAN_ACU, 2), 2)
    Printer.Print "   _____________________                      ____________________       Saldo:     " & FormatCurrency(FormatNumber(SAL_ACU, 2), 2)
    Printer.Print "     Firma de Recibido                        Vendedor Responsable"
    Printer.EndDoc
Wend
End Sub

Private Sub Command4_Click()
Form8.Hide
Form1.Show
End Sub

Private Sub Form_Activate()
Call Limpiar
Set Setup = Inventa.OpenRecordset("SETUP")
Set Facturas = Inventa.OpenRecordset("FACTURAS")
Set Clientes = Inventa.OpenRecordset("CLIENTES")
Set Recibos = Inventa.OpenRecordset("RECIBOS")
Set Rec_Detalle = Inventa.OpenRecordset("RECIBOS DETALLE")
Clientes.Index = "Cliente"
Clientes.MoveFirst
While Clientes.EOF = False
    If Clientes!Codigo <> 1 Then
        Combo1.AddItem Clientes!Cliente
    End If
    Clientes.MoveNext
Wend
End Sub

Private Sub Form_Deactivate()
Call Command4_Click
End Sub

Private Sub List1_Click()
If List1.ListCount > 0 Then
    Label5.Caption = Mid(List1.Text, 10, 8)
    Label8.Caption = Mid(List1.Text, 19, 12)
    Cancelado = FormatNumber(Mid(List1.Text, 46, 13), 2)
    Facturado = FormatNumber(Mid(List1.Text, 33, 13), 0)
    Label10.Caption = FormatCurrency(Facturado - Cancelado, 2)
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
End If
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3.SetFocus
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Len(Text1.Text) > 1 Then
    If Text1.Text <> 1 Then
        Clientes.Index = "PrimaryKey"
        Clientes.Seek "=", Text1.Text
        If Not Clientes.NoMatch Then
            Combo1.Text = Clientes!Cliente
        End If
    End If
End If
End Sub

Private Sub Timer1_Timer()
If List1.ListCount > 0 Then
    If Not Line1.Visible Then
        Line1.Visible = True
        Line2.Visible = True
        Line3.Visible = True
    Else
        Line1.Visible = False
        Line2.Visible = False
        Line3.Visible = False
    End If
Else
    Line1.Visible = False
    Line2.Visible = False
    Line3.Visible = False
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TMP = FormatNumber(Mid(List1.Text, 60, 13), 0)
    If Len(Text3.Text) > 0 And Mid(List1.Text, 76, 1) <> "*" Then
        If IsNumeric(Text3.Text) Then
            If Text3.Text <= 0 Then
                MsgBox "El monto no puede ser 0"
            Else '76
                POSICION = List1.ListIndex
                info = List1.Text
                Cancelado = FormatNumber(Mid(List1.Text, 46, 13), 2)
                abono = FormatNumber(Mid("             ", 1, 13 - Len(FormatCurrency(Text3.Text, 2))) & FormatCurrency(Text3.Text, 2), 2)
                Facturado = FormatNumber(Mid(List1.Text, 33, 13), 0)
                If Not (CDbl(Cancelado) + CDbl(abono)) > Facturado Then
                    Cancelado = CDbl(Cancelado) + CDbl(abono)
                    List1.RemoveItem (POSICION)
                    List1.AddItem (Mid(info, 1, 45) & Mid("             ", 1, 13 - Len(FormatCurrency(Cancelado))) & FormatCurrency(Cancelado) & " " & Mid("             ", 1, 13 - Len(FormatCurrency(abono))) & FormatCurrency(abono) & "   " & "*")
                    List1.ListIndex = POSICION
                    List1.SetFocus
                Else
                    MsgBox "Monto invalido"
                End If
            End If
        End If
    Else
        MsgBox "Este elemento ya fue modificado"
    End If
End If
End Sub
