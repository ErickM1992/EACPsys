VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form14 
   Caption         =   "Reportes..."
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   Icon            =   "ReporteTemporal.frx":0000
   LinkTopic       =   "Form14"
   ScaleHeight     =   7575
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   22
      Top             =   7320
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   2775
      Left            =   5640
      TabIndex        =   21
      Top             =   360
      Width           =   3135
      Begin VB.OptionButton Option7 
         Caption         =   "Reporte de Clientes"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Reporte de Ventas"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2895
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Ventas por Vendedor"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   2895
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Control de Existencias"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   2895
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Historico de Articulos"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   2895
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Listado de Articulos"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   2895
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Rotación de Articulos"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   2895
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9120
      Top             =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cerrar"
      Height          =   735
      Left            =   9000
      TabIndex        =   13
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir Pantalla Actual"
      Height          =   735
      Left            =   9000
      TabIndex        =   12
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar"
      Height          =   735
      Left            =   9000
      TabIndex        =   11
      Top             =   480
      Width           =   2295
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
      Height          =   3660
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   11175
   End
   Begin MSComCtl2.MonthView MonthView2 
      Height          =   2370
      Left            =   2880
      TabIndex        =   8
      Top             =   480
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   59768833
      CurrentDate     =   38234
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   59768833
      CurrentDate     =   38234
   End
   Begin VB.Label T4 
      Height          =   135
      Left            =   2160
      TabIndex        =   20
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label T3 
      Height          =   255
      Left            =   1560
      TabIndex        =   19
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label T2 
      Height          =   255
      Left            =   840
      TabIndex        =   18
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label T1 
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Final"
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha de Inicio"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
T1.Caption = ""
T2.Caption = ""
T3.Caption = ""
T4.Caption = ""
If Option1.Value Then
    Call ReporteVentas
ElseIf Option2.Value Then
    Call VentasVendedor
ElseIf Option3.Value Then
    Call Existencias
ElseIf Option4.Value Then
    Call Historico
ElseIf Option5.Value Then
    Call Listado
ElseIf Option6.Value Then
    Call Rotacion
ElseIf Option7.Value Then
    Call Clientela
End If
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
End Sub

Private Sub Clientela()
List1.Clear
Facturas_Tot = 0
Monto_Tot = 0
Set Clientes = Inventa.OpenRecordset("CLIENTES")
Set Facturas = Inventa.OpenRecordset("FACTURAS")
Clientes.Index = "Cliente"
List1.AddItem "Cod Nombre del Cliente                       Fac Deuda total  Limite cred. Des", 0
              '123 1234567890123456789012345678901234567890 123 123456789012 123456789012 123
While Not Clientes.EOF
    TMP = Clientes!Codigo
    
    Facturas.Index = "Cliente"
    Facturas.Seek "=", Clientes!Codigo
    Facturas_Cli = 0
    Monto_Cli = 0
    If Not IsNull(Clientes!LIMITE_CRE) Then
        Limite_Cli = FormatCurrency(Clientes!LIMITE_CRE, 2)
    Else
        Limite_Cli = FormatCurrency(0, 2)
    End If
    If Not IsNull(Clientes!Descuento) Then
        Descuento_Cli = Clientes!Descuento
    Else
        Descuento_Cli = 0
    End If
    If Not Facturas.NoMatch Then
        rango = True
        While rango And Not Facturas.EOF
            If Facturas!Cli_cod = TMP And TMP <> 1 Then
                Sigla = Mid(Facturas!Num_Fac, 1, 1)
                If Not Facturas!Cancelado Then
                    If (Sigla = "R" Or Sigla = "A" Or Sigla = "T") Then
                        Facturas_Cli = Facturas_Cli + 1
                        Monto_Cli = Monto_Cli + Facturas!Monto_S
                    End If
                End If
                Facturas.MoveNext
            Else
                rango = False
            End If
        Wend
    Else
    End If
    If TMP <> 1 Then
        Codigo_Cli = Mid("   ", 1, 3 - Len(Clientes!Codigo)) & Clientes!Codigo
        Nombre_Cli = Clientes!Cliente & Mid("                                        ", 1, 40 - Len(Clientes!Cliente))
        Facturas_Cli = Mid("   ", 1, 3 - Len(Facturas_Cli)) & Facturas_Cli
        Monto_Tot = Monto_Tot + Monto_Cli
        Facturas_Tot = Facturas_Tot + Facturas_Cli
        Monto_Cli = Mid("             ", 1, 13 - Len(FormatCurrency(Monto_Cli, 2))) & FormatCurrency(Monto_Cli, 2)
        Limite_Cli = Mid("             ", 1, 13 - Len(Limite_Cli)) & Limite_Cli
        Descuento_Cli = Mid("   ", 1, 3 - Len(Descuento_Cli)) & Descuento_Cli
        List1.AddItem (Codigo_Cli & " " & Nombre_Cli & " " & Facturas_Cli & " " & Monto_Cli & " " & Limite_Cli & " " & Descuento_Cli)
    End If
    Clientes.MoveNext
Wend
List1.AddItem ""
List1.AddItem "Facturas pendientes: " & Facturas_Tot
List1.AddItem "Monto total: " & FormatCurrency(Monto_Tot, 2)
Clientes.Close
Facturas.Close
End Sub
Private Sub ReporteVentas()
List1.Clear
B = False
Set Facturas = Inventa.OpenRecordset("Facturas")
Facturas.Index = "Fecha"
Facturas.MoveFirst
While B = False And Facturas.EOF = False
    ProgressBar1.Value = Facturas.PercentPosition
    If Facturas!fecha < MonthView1.Value Then
         Facturas.MoveNext
    Else
        B = True
    End If
    
Wend
rango = False
Dim Contado As Double
Dim Credito As Double
Dim Cancelado As Double
Dim Des_Mon As Double
Dim Sub_Mon As Double
Contado = 0
Con_Con = 0
Credito = 0
Cre_Con = 0
Cancelado = 0
T = 0
R = 0
p = 0
'If B Then
    List1.AddItem ("  Fecha   Tipo       Factura    Cliente                               Des.   %         Monto ")
    T1 = "  Fecha   Tipo       Factura    Cliente                               Des.   %         Monto "
'End If

    While rango = False And Facturas.EOF = False
        ProgressBar1.Value = Facturas.PercentPosition
        If MonthView2.Value >= Facturas!fecha Then
            Tipo = Mid(Facturas!Num_Fac, 1, 1)
            Numero = Mid(Facturas!Num_Fac, 2, 50)
            If Tipo = "T" Then
                Tipo = "Factura 1"
                T = T + 1
            End If
            If Tipo = "R" Then
                Tipo = "Factura 2"
                R = R + 1
            End If
            If Tipo = "P" Then
                Tipo = "Proforma "
                p = p + 1
            End If
            If Facturas!Cancelado And Tipo <> "Proforma " And Facturas!Anulada = False Then
                ok = " "
            Else
                If Tipo = "Proforma " Then
                    ok = " "
                Else
                    ok = "¤"
                End If
            End If
            Cliente = Mid(Facturas!Cli_Nom & Mid("                                      ", 1, 38 - Len(Mid(Facturas!Cli_Nom, 1, 38))), 1, 28)
            Descuento = FormatCurrency(Facturas!Monto_F * Facturas!Des / (100 - Facturas!Des))
            Descuento = Mid("             ", 1, 13 - Len(Descuento)) & Descuento
            List1.AddItem (ok & " " & Facturas!fecha & " " & Tipo & " " & Mid("          ", 1, 10 - Len(Numero)) & Numero & " " & Cliente & " " & Descuento & " " & Mid("   ", 1, 3 - Len(Facturas!Des)) & Facturas!Des & " " & Mid("             ", 1, 13 - Len(FormatCurrency(Facturas!Monto_F))) & FormatCurrency(Facturas!Monto_F))
            If Tipo <> "Proforma " And Facturas!Anulada = False Then
                If Facturas!Cancelado Then
                    Contado = FormatNumber(Facturas!Monto_F + Contado, 2)
                    Con_Con = Con_Con + 1
                    Des_Mon = Des_Mon + FormatNumber(Descuento, 0)
                    Sub_Mon = (Facturas!Monto_F / ((100 - Facturas!Des) / 100)) + Sub_Mon
                Else
                    Credito = FormatNumber(Facturas!Monto_F + Credito, 2)
                    Cre_Con = Cre_Con + 1
                    Cancelado = FormatNumber(Cancelado + (Facturas!Monto_F - Facturas!Monto_S), 2)
                End If
            End If
            Facturas.MoveNext
        Else
            rango = True
        End If
    Wend
List1.AddItem ("")
List1.AddItem ("Desde el " & MonthView1.Value & " hasta el " & MonthView2.Value & " se han emitido " & List1.ListCount - 2 & " facturas")
List1.AddItem ("de la cuales " & T & " son Factura 1, " & R & " son Factura 2 y " & p & " son Proformas")
List1.AddItem ("Monto Total en " & Con_Con & " Facturas de Contado: " & FormatCurrency(Contado, 2))
List1.AddItem ("Monto Total en " & Cre_Con & " Facturas de Credito: " & FormatCurrency(Credito, 2) & ", donde se han cancelado " & FormatCurrency(Cancelado))
List1.AddItem ("Monto Total en facturas de contado: " & FormatCurrency(Sub_Mon, 2))
List1.AddItem ("Monto Total en descuentos para estas facturas: " & FormatCurrency(Des_Mon, 2))
List1.AddItem ("Monto Total Facturado: " & FormatCurrency(Credito + Contado, 2))
List1.AddItem ("Total de Entradas: " & FormatCurrency(FormatNumber(Cancelado, 0) + FormatNumber(Contado, 0), 2))
End Sub

Private Sub VentasVendedor()
Set Facturas = Inventa.OpenRecordset("Facturas")
Set Vendedores = Inventa.OpenRecordset("Vendedores")
Set VXV = Inventa.OpenRecordset("vxv")
VXV.Index = "PrimaryKey"
Facturas.Index = "Fecha"
Facturas.MoveFirst
While Not Facturas.EOF
    Progreso = Facturas.PercentPosition / 2
    ProgressBar1.Value = Progreso
    If MonthView1.Value <= Facturas!fecha And MonthView2.Value >= Facturas!fecha Then
        VXV.Seek "=", Facturas!VEN_COD
        If VXV.NoMatch Then
            VXV.AddNew
            VXV!Codigo = Facturas!VEN_COD
            VXV!Vendido = VXV!Vendido + (Facturas!Monto_F - Facturas!Monto_S)
            VXV.Update
        Else
            VXV.Edit
            VXV!Vendido = VXV!Vendido + (Facturas!Monto_F - Facturas!Monto_S)
            VXV.Update
        End If
    End If
    Facturas.MoveNext
    If Facturas.EOF Then
    x = Date
    End If
Wend
If Not VXV.EOF Then
    VXV.MoveFirst
    Vendedores.Index = "PrimaryKey"
    List1.Clear
    List1.AddItem ("Cod Nombre                              Total Vendido")
    While Not VXV.EOF
        ProgressBar1.Value = Progreso + (VXV.PercentPosition / 2)
        Vendedores.Seek "=", VXV!Codigo
        List1.AddItem (Mid("   ", 1, 3 - Len(VXV!Codigo)) & VXV!Codigo & " " & Vendedores!Nombre & Mid("                                   ", 1, 35 - Len(Vendedores!Nombre)) & " " & Mid("                    ", 1, 20 - Len(FormatCurrency(VXV!Vendido, 2))) & FormatCurrency(VXV!Vendido, 2))
        VXV.MoveNext
    Wend
    VXV.MoveFirst
    While Not VXV.EOF
        VXV.Delete
        VXV.MoveFirst
    Wend
    ProgressBar1.Value = 100
End If
End Sub

Private Sub Existencias()
List1.Clear
Set Inventa = OpenDatabase("C:\Users\Erick M\Desktop\Nueva carpeta (3)\Codigo Fuente\EACPsys.mdb")
Set Tabla = Inventa.OpenRecordset("Inventa")
Tabla.MoveFirst
List1.AddItem (" Control de Existencias")
List1.AddItem ("Codigo     Descripcion                              Cantidad Stock")
While Not Tabla.EOF
    ProgressBar1.Value = Tabla.PercentPosition
    If Tabla!Cantidad <= Tabla!Stock Then
        T1.Caption = " Control de Existencias"
        T2.Caption = "Codigo     Descripcion                              Cantidad Stock"
        List1.AddItem (Tabla!Codigo & " " & Tabla!descrip & Mid("                                             ", 1, 40 - Len(Tabla!descrip)) & " " & Mid("        ", 1, 8 - Len(Tabla!Cantidad)) & Tabla!Cantidad & " " & Mid("     ", 1, 5 - Len(Tabla!Stock)) & Tabla!Stock)
    End If
    Tabla.MoveNext
Wend
End Sub

Private Sub Historico()
List1.Clear
B = False
TXT_COD = Mid("0000000000", 1, 10 - Len(Text1.Text)) & Text1.Text
TXT_COD = UCase(TXT_COD)
Set Tabla = Inventa.OpenRecordset("Inventa")
Set Facturas = Inventa.OpenRecordset("Facturas")
Set Detalle = Inventa.OpenRecordset("Detalle")
Set FacturaIN = Inventa.OpenRecordset("Factura de Ingreso")
Set DetalleIN = Inventa.OpenRecordset("DETALLE DE INGRESO")
'****************************************************************************************************

'FacturaIN.Index = "Codigo"
'DetalleIN.Index = "Detallecodigo"
'DetalleIN.Seek "=", Text1.Text
List1.AddItem (" Entradas ")
List1.AddItem ("Fecha    Prov.       Factura Descripcion                              Cantidad")
T1.Caption = " Entradas "
T2.Caption = "Fecha    Prov.       Factura Descripcion                              Cantidad"
DetalleIN.Index = "Fecha"
DetalleIN.MoveFirst
While Not DetalleIN.EOF
    Progreso = DetalleIN.PercentPosition / 2
    ProgressBar1.Value = Progreso
    If (UCase(DetalleIN!Codigo) = TXT_COD) And (MonthView1.Value <= DetalleIN!fecha) And (MonthView2.Value >= DetalleIN!fecha) Then
        Prov = Mid(DetalleIN!Num_Fac, 1, 3)
        Numero = Mid(DetalleIN!Num_Fac, 4, 50)
        Tabla.Index = "Codigo"
        Tabla.Seek "=", DetalleIN!Codigo
        List1.AddItem (DetalleIN!fecha & " " & Prov & " " & Mid("               ", 1, 15 - Len(Numero)) & Numero & " " & Tabla!descrip & Mid("                                        ", 1, 40 - Len(Tabla!descrip)) & " " & Mid("        ", 1, 8 - Len(DetalleIN!Cantidad)) & DetalleIN!Cantidad)
    End If
    DetalleIN.MoveNext
Wend

'****************************************************************************************************
Detalle.Index = "DetalleCodigo"
Detalle.Seek "=", TXT_COD
Facturas.Index = "PrimaryKey"
rango = False
List1.AddItem (" ")
List1.AddItem (" Salidas")
List1.AddItem ("Fecha    Tipo       Factura        Descripcion                              Cantidad")
T3.Caption = " Salidas"
T4.Caption = "Fecha    Tipo       Factura        Descripcion                              Cantidad"
If Not Detalle.NoMatch Then
    While rango = False And Detalle.EOF = False
        ProgressBar1.Value = Progreso + (Detalle.PercentPosition / 2)
        Facturas.Seek "=", Detalle!Num_Fac
        If Detalle!Codigo = TXT_COD And MonthView1.Value <= Facturas!fecha And MonthView2.Value >= Facturas!fecha Then
            Tipo = Mid(Detalle!Num_Fac, 1, 1)
            Numero = Mid(Detalle!Num_Fac, 2, 50)
            If Tipo = "T" Then
                Tipo = "Factura 1"
                T = T + 1
            End If
            If Tipo = "R" Then
                Tipo = "Factura 2"
                R = R + 1
            End If
            If Tipo = "P" Then
                Tipo = "Proforma "
                p = p + 1
            End If
            If Tipo <> "Proforma " Then
                Tabla.Index = "Codigo"
                Tabla.Seek "=", Detalle!Codigo
                If Facturas!Anulada Then
                    List1.AddItem (Facturas!fecha & " " & Tipo & "*" & Mid("               ", 1, 15 - Len(Numero)) & Numero & " " & Tabla!descrip & Mid("                                        ", 1, 40 - Len(Tabla!descrip)) & " " & Mid("        ", 1, 8 - Len(Detalle!Cantidad)) & Detalle!Cantidad)
                Else
                    List1.AddItem (Facturas!fecha & " " & Tipo & " " & Mid("               ", 1, 15 - Len(Numero)) & Numero & " " & Tabla!descrip & Mid("                                        ", 1, 40 - Len(Tabla!descrip)) & " " & Mid("        ", 1, 8 - Len(Detalle!Cantidad)) & Detalle!Cantidad)
                End If
            End If
            Detalle.MoveNext
        Else
            Detalle.MoveNext
        End If
    Wend
End If
List1.AddItem ("  ")
End Sub

Private Sub Listado()
List1.Clear
Set Tabla = Inventa.OpenRecordset("Inventa")
Set Setup = Inventa.OpenRecordset("Setup")
Tabla.Index = "Datos"
Tabla.MoveFirst
List1.AddItem (" Listado de Articulos")                                     '1234567   1234567
               '1234567890 1234567890123456789012345678901234 12345678 1234567890123 123456789 1234567890123
List1.AddItem (" Codigo     Descripcion                       Cantidad      P. Costo %Utilidad   P. Unitario Ubicación")
T1 = " Listado de Articulos"
T2 = " Codigo     Descripcion                       Cantidad      P. Costo %Utilidad   P. Unitario Ubicación"
A_U = 0
A_V = 0
While Not Tabla.EOF
    ProgressBar1.Value = Tabla.PercentPosition
    If Not IsNull(Tabla!gabeta) Then
        U1 = Mid("       ", 1, 7 - Len(Tabla!gabeta)) & Tabla!gabeta
    Else
        U1 = "       "
    End If
    If Not IsNull(Tabla!ubicacion) Then
        U2 = Mid("       ", 1, 7 - Len(Tabla!ubicacion)) & Tabla!ubicacion
    Else
        U2 = "       "
    End If
    D = Mid(Tabla!descrip & Mid("                                             ", 1, 40 - Len(Tabla!descrip)), 1, 34)
    P_U = Mid("             ", 1, 13 - Len(FormatCurrency(Tabla!P_unit, 2))) & FormatCurrency(Tabla!P_unit, 2)
    P_V = Mid("             ", 1, 13 - Len(FormatCurrency(Tabla!P_Venta, 2))) & FormatCurrency(Tabla!P_Venta, 2)
    U = Mid("       ", 1, 9 - Len(FormatNumber(Tabla!Por_Venta, 2))) & FormatNumber(Tabla!Por_Venta, 2)
    If Tabla!Codigo <> "EACP" Then
        List1.AddItem (Tabla!Codigo & " " & D & " " & Mid("        ", 1, 8 - Len(Tabla!Cantidad)) & Tabla!Cantidad & " " & P_U & " " & U & " " & P_V & " " & U1 & "¦" & U2)
        A_U = (P_U * Tabla!Cantidad) + A_U
        A_V = (P_V * Tabla!Cantidad) + A_V
    End If
    Tabla.MoveNext
Wend
List1.AddItem ("Total de precios al costo: " & FormatCurrency(A_U, 2))
List1.AddItem ("Total de precios de venta: " & FormatCurrency(A_V, 2))
End Sub

Private Sub Rotacion()
List1.Clear
Set Tabla = Inventa.OpenRecordset("Inventa")
Set Detalle = Inventa.OpenRecordset("Detalle")
Detalle.Index = "DetalleCodigo"
Set Facturas = Inventa.OpenRecordset("Facturas")
Facturas.Index = "PrimaryKey"
Set Setup = Inventa.OpenRecordset("Setup")
Tabla.Index = "Datos"
Tabla.MoveFirst
List1.AddItem (" Listado de Articulos")                                     '1234567   1234567
               '1234567890 1234567890123456789012345678901234 12345678 1234567890123 123456789012345 1234567890123
List1.AddItem (" Codigo     Descripcion                       Cantidad      P. Costo Ubicación       P. Unitario")
T1 = " Listado de Articulos"
T2 = " Codigo     Descripcion                       Cantidad      P. Costo %Utilidad   P. Unitario Ubicación"
A_U = 0
A_V = 0
While Not Tabla.EOF
    ProgressBar1.Value = Tabla.PercentPosition
'    If Tabla!Codigo = "000STK4152" Then
'        p = 0
'    End If
    Detalle.Seek "=", Tabla!Codigo
    If Detalle.NoMatch Then
        If Not IsNull(Tabla!gabeta) Then
            U1 = Mid("       ", 1, 7 - Len(Tabla!gabeta)) & Tabla!gabeta
        Else
            U1 = "       "
        End If
        If Not IsNull(Tabla!ubicacion) Then
            U2 = Mid("       ", 1, 7 - Len(Tabla!ubicacion)) & Tabla!ubicacion
        Else
            U2 = "       "
        End If
        D = Mid(Tabla!descrip & Mid("                                             ", 1, 40 - Len(Tabla!descrip)), 1, 34)
        P_U = Mid("             ", 1, 13 - Len(FormatCurrency(Tabla!P_unit, 2))) & FormatCurrency(Tabla!P_unit, 2)
        P_V = Mid("             ", 1, 13 - Len(FormatCurrency(Tabla!P_Venta, 2))) & FormatCurrency(Tabla!P_Venta, 2)
        U = Mid("        ", 1, 9 - Len(FormatNumber(Tabla!Por_Venta, 2))) & FormatNumber(Tabla!Por_Venta, 2)
        If Tabla!Codigo <> "EACP" Then
            List1.AddItem (Tabla!Codigo & " " & D & " " & Mid("        ", 1, 8 - Len(Tabla!Cantidad)) & Tabla!Cantidad & " " & P_U & " " & U1 & "¦" & U2 & " " & P_V)
            A_U = (P_U * Tabla!Cantidad) + A_U
            A_V = (P_V * Tabla!Cantidad) + A_V
        End If
        Tabla.MoveNext
    Else
        Facturas.Seek "=", Detalle!Num_Fac
        Articulo = Detalle!Codigo
        If Not Detalle.EOF Then
            conteo = 0
            While Detalle!Codigo = Articulo
                If Not Facturas.NoMatch Then
                    If Not (MonthView1.Value <= Facturas!fecha And MonthView2.Value >= Facturas!fecha) Then
                        conteo = conteo + 1
                    End If
                End If
                Detalle.MoveNext
                If Detalle.EOF Then
                    Detalle.MoveFirst
                End If
                ok = 1
            Wend
            If 0 < conteo Then
                If Not IsNull(Tabla!gabeta) Then
                    U1 = Mid("       ", 1, 7 - Len(Tabla!gabeta)) & Tabla!gabeta
                Else
                    U1 = "       "
                End If
                If Not IsNull(Tabla!ubicacion) Then
                    U2 = Mid("       ", 1, 7 - Len(Tabla!ubicacion)) & Tabla!ubicacion
                Else
                    U2 = "       "
                End If
                D = Mid(Tabla!descrip & Mid("                                             ", 1, 40 - Len(Tabla!descrip)), 1, 34)
                P_U = Mid("             ", 1, 13 - Len(FormatCurrency(Tabla!P_unit, 2))) & FormatCurrency(Tabla!P_unit, 2)
                P_V = Mid("             ", 1, 13 - Len(FormatCurrency(Tabla!P_Venta, 2))) & FormatCurrency(Tabla!P_Venta, 2)
                U = Mid("       ", 1, 8 - Len(FormatNumber(Tabla!Por_Venta, 2))) & FormatNumber(Tabla!Por_Venta, 2)
                If Tabla!Codigo <> "EACP" Then
                    List1.AddItem (Tabla!Codigo & " " & D & " " & Mid("        ", 1, 8 - Len(Tabla!Cantidad)) & Tabla!Cantidad & " " & P_U & " " & U1 & "¦" & U2 & " " & P_V)
                    A_U = (P_U * Tabla!Cantidad) + A_U
                    A_V = (P_V * Tabla!Cantidad) + A_V
                End If
            End If
            Tabla.MoveNext
        End If
    End If
Wend
List1.AddItem ("Total de precios al costo: " & FormatCurrency(A_U, 2))
List1.AddItem ("Total de precios de venta: " & FormatCurrency(A_V, 2))
End Sub

Private Sub Command2_Click()
Set Setup = Inventa.OpenRecordset("Setup")
Dim x As Printer
For Each x In Printers
    If x.DeviceName = Setup!pro_prn Then
        Set Printer = x
        Printer.Font = "Courier New"
        Printer.FontSize = 10
        Printer.PaperSize = 1
        If Option5.Value Then
            Printer.Orientation = 2
        End If
        I = 0
        C = 0
        If Option1.Value Or Option2.Value Then
            Printer.Print T1
            I = 1
        ElseIf Option3.Value Or Option5.Value Then
            Printer.Print T1
            Printer.Print T2
            I = 2
            C = C + 1
        End If
        List1.ListIndex = I
        While I < List1.ListCount
            ProgressBar1.Value = (I / List1.ListCount) * 100
            List1.ListIndex = I
            If C = 0 And I > 2 Then
                If Option1.Value Or Option2.Value Then
                    Printer.Print T1
                ElseIf Option3.Value Or Option5.Value Then
                    Printer.Print T1
                    Printer.Print T2
                    C = C + 1
                End If
            End If
            Printer.Print List1.Text
            If C = 64 And Printer.Orientation = 1 Then
                C = 0
            ElseIf C = 39 And Printer.Orientation = 2 Then
                C = 0
            Else
                C = C + 1
            End If
            I = I + 1
        Wend
        Printer.EndDoc
    End If
    ProgressBar1.Value = 100
Next
End Sub

Private Sub Command3_Click()
Form14.Hide
Form1.Show
End Sub

Private Sub Form_Activate()
MonthView1.Value = Date
MonthView2.Value = Date
List1.Clear
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
If Form0.Option3 Then
    Option1.Value = True
Else
    Option3.Value = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
Form1.Show
End Sub

Private Sub Option1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1.SetFocus
End If
End Sub

Private Sub Option2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1.SetFocus
End If
End Sub

Private Sub Option3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1.SetFocus
End If
End Sub

Private Sub Option4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1.SetFocus
End If
End Sub

Private Sub Option5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1.SetFocus
End If
End Sub

Private Sub Option6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1.SetFocus
End If
End Sub

Private Sub Option7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1.SetFocus
End If
End Sub

Private Sub Option8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1.SetFocus
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
If Form0.Option3.Value Then
    If Not Option1.Enabled And Not Option2.Enabled Then
        Option1.Enabled = True
        Option2.Enabled = True
        Option7.Enabled = True
        Option1.Value = True
    End If
Else
    If Option1.Enabled And Option2.Enabled Then
        Option1.Enabled = False
        Option2.Enabled = False
        Option7.Enabled = False
        Option3.Value = True
    End If
End If
If Option4.Value Then
    If Not Text1.Visible Then
        Text1.Visible = True
    End If
Else
    If Text1.Visible Then
        Text1.Visible = False
    End If
End If
If Option5.Value Or Option3.Value Or Option7.Value Then
    If MonthView1.Value Then
        MonthView1.Enabled = False
        MonthView2.Enabled = False
    End If
Else
    If Not MonthView1.Enabled Then
        MonthView1.Enabled = True
        MonthView2.Enabled = True
    End If
End If
If Option1.Value Or Option2.Value Or Option4.Value Or Option6.Value Then
    If Not MonthView1.Enabled Then
        MonthView1.Enabled = True
        MonthView2.Enabled = True
    End If
Else
    If MonthView1.Value Then
        MonthView1.Enabled = False
        MonthView2.Enabled = False
    End If
End If
End Sub
