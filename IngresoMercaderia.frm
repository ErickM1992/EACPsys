VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Mercadería"
   ClientHeight    =   5670
   ClientLeft      =   -135
   ClientTop       =   -45
   ClientWidth     =   7830
   Icon            =   "IngresoMercaderia.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   378
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   522
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   4800
      TabIndex        =   1
      Text            =   "1"
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   3000
      TabIndex        =   0
      Text            =   "1"
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame Frame5 
      Height          =   50
      Left            =   120
      TabIndex        =   35
      Top             =   5040
      Width           =   7455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5160
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5160
      Width           =   2000
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
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
      Left            =   6840
      MaxLength       =   7
      TabIndex        =   10
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
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
      Left            =   4680
      MaxLength       =   7
      TabIndex        =   9
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      TabIndex        =   8
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      TabIndex        =   7
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   6240
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   3480
      TabIndex        =   26
      Top             =   13080
      Width           =   5895
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   7575
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   1440
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
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
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   720
      Width           =   3855
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Valor agregado"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   19
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Tipo de Cambio"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   18
      Top             =   120
      Width           =   1575
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Stock"
      Height          =   195
      Left            =   840
      TabIndex        =   37
      Top             =   2160
      Width           =   495
   End
   Begin VB.Line Line4 
      X1              =   504
      X2              =   512
      Y1              =   176
      Y2              =   184
   End
   Begin VB.Line Line3 
      X1              =   512
      X2              =   504
      Y1              =   120
      Y2              =   128
   End
   Begin VB.Line Line2 
      X1              =   256
      X2              =   248
      Y1              =   176
      Y2              =   184
   End
   Begin VB.Line Line1 
      X1              =   248
      X2              =   256
      Y1              =   120
      Y2              =   128
   End
   Begin VB.Shape Shape3 
      Height          =   975
      Left            =   3720
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3840
      TabIndex        =   36
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Precio al costo"
      Height          =   195
      Left            =   285
      TabIndex        =   34
      Top             =   2520
      Width           =   1050
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Por venta"
      Height          =   195
      Left            =   630
      TabIndex        =   33
      Top             =   2880
      Width           =   690
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Ubicación 1"
      Height          =   195
      Left            =   3720
      TabIndex        =   32
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Ubicación 2"
      Height          =   195
      Left            =   5760
      TabIndex        =   31
      Top             =   2880
      Width           =   855
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
      TabIndex        =   30
      Top             =   3480
      Width           =   4815
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
      TabIndex        =   29
      Top             =   3480
      Width           =   735
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
      TabIndex        =   28
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lista de Artículos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   3240
      TabIndex        =   27
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7800
      TabIndex        =   23
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Ultima Factura"
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
      Left            =   6480
      TabIndex        =   22
      Top             =   360
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Factura"
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
      Left            =   6240
      TabIndex        =   21
      Top             =   840
      Width           =   660
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3720
      TabIndex        =   20
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Descripción"
      Height          =   195
      Left            =   495
      TabIndex        =   17
      Top             =   1440
      Width           =   840
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Left            =   840
      TabIndex        =   16
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cantidad"
      Height          =   195
      Left            =   705
      TabIndex        =   15
      Top             =   1800
      Width           =   630
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Proveedor"
      Height          =   195
      Left            =   600
      TabIndex        =   14
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Guardar()
Codigo = Mid("***", 1, 3 - Len(Combo1.Text)) & Combo1.Text & Text5.Text
FacturaIN.Index = "PrimaryKey"
FacturaIN.Seek "=", Codigo
If FacturaIN.NoMatch Then
    FacturaIN.AddNew
    FacturaIN!fac_num = Codigo
    FacturaIN!Pro_Cod = Combo1.Text
    FacturaIN.Update
    DetalleIN.AddNew
    DetalleIN!Num_Fac = Codigo
    DetalleIN!fecha = Date
    DetalleIN!Codigo = Text2.Text
    DetalleIN!Cantidad = Text4.Text
    DetalleIN!Pre_Uni = Label10.Caption
    DetalleIN!fecha = Date
Else
    DetalleIN.AddNew
    DetalleIN!Num_Fac = Codigo
    DetalleIN!Codigo = Text2.Text
    DetalleIN!Cantidad = Text4.Text
    DetalleIN!Pre_Uni = Label10.Caption
    DetalleIN!fecha = Date
End If
Tabla.Seek "=", Mid("0000000000", 1, 10 - Len(Text2.Text)) & Text2.Text
If Tabla.NoMatch Then
    Tabla.AddNew
    Tabla!Codigo = Text2.Text
    Tabla!descrip = Text3.Text
    Tabla!Prove = Combo1
    Tabla!Cantidad = Text4.Text
    Tabla!Stock = Text12.Text
    Tabla!P_Venta = FormatNumber(Label10.Caption)
    Tabla!Por_Venta = Text9.Text
    Tabla!P_unit = Text8.Text * Text10.Text * Text11.Text
    Tabla!gabeta = Text6.Text
    Tabla!ubicacion = Text7.Text
    Tabla!Fe_Ult_Com = Date
    Tabla!EQUIVALENTE = "*"
Else
    Tabla.Edit
    Tabla!Cantidad = Tabla!Cantidad + Text4.Text
    Tabla!P_unit = Text8.Text * Text10.Text * Text11.Text
    Tabla!P_Venta = FormatNumber(Label10.Caption)
    Tabla!Por_Venta = Text9.Text
    Tabla!Fe_Ult_Com = Date
    Tabla!Stock = Text12.Text
    Tabla!Prove = Combo1
End If
Tabla.Update
DetalleIN.Update
Call Limpiar
Text11.SelStart = 0
Text11.SelLength = 40
Text11.SetFocus
End Sub

Private Sub Listado()
Codigo = Mid("***", 1, 3 - Len(Combo1.Text)) & Combo1.Text & Text5.Text
FacturaIN.Index = "PrimaryKey"
DetalleIN.Index = "Detalle"
Tabla.Index = "PrimaryKey"
FacturaIN.Seek "=", Codigo
If Not FacturaIN.NoMatch Then
    DetalleIN.Seek "=", FacturaIN!fac_num
    ok = True
    List1.Clear
    If Not DetalleIN.EOF Then
        While ok
            If FacturaIN!fac_num = DetalleIN!Num_Fac Then
                Tabla.Seek "=", DetalleIN!Codigo
                List1.AddItem ("  " & Tabla!descrip & Mid("                                       ", 1, 39 - Len(Mid(Tabla!descrip, 1, 39))) & " " & Mid("     ", 1, 5 - Len(DetalleIN!Cantidad)) & DetalleIN!Cantidad & " " & Mid("            ", 1, 12 - Len(DetalleIN!Pre_Uni)) & DetalleIN!Pre_Uni)
                DetalleIN.MoveNext
                If DetalleIN.EOF Then
                    ok = False
                End If
            Else
                ok = False
            End If
        Wend
    End If
Else
    List1.Clear
End If
End Sub

Private Sub Limpiar()
Combo1.ListIndex = -1
Call Combo1_Click
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text9.Text = ""
Text10.Text = ""
Text12.Text = ""
Label2.Caption = ""
Label10.Caption = ""
List1.Clear
End Sub

Private Sub Precio()
If IsNumeric(Text10.Text) And IsNumeric(Text9.Text) Then
    If Text10.Text >= 0 And Text9.Text >= 0 Then
        COLONES = FormatNumber(Text10.Text, 2) * FormatNumber(Text8.Text, 2) * FormatNumber(Text11.Text, 2)
        MONTO = FormatCurrency(COLONES + ((FormatNumber(Text9.Text, 2) * COLONES) / 100), 2)
        Label10.Caption = MONTO
    '    Label15.Caption = FormatCurrency(Text4.Text + ((Text4.Text * Text5.Text) / 100), 2)
    End If
End If
End Sub

Private Sub Combo1_Click()
If Combo1 = "" Then
    Text1.Text = "- Seleccione un proveedor -"
Else
    Provedores.Seek "=", Combo1.Text
    Text1.Text = Provedores!Nombre
    FacturaIN.Index = "PrimaryKey"
    FacturaIN.Seek "=", Combo1.Text
    If Not FacturaIN.NoMatch Then
        While Not FacturaIN.EOF
            If Combo1.Text = FacturaIN!Pro_Cod Then
                FacturaIN.MoveNext
            Else
                FacturaIN.MovePrevious
            End If
        Wend
        If FacturaIN.EOF Then
            FacturaIN.MoveLast
        End If
        Label5.Caption = FacturaIN!fac_num
    Else
        Label5.Caption = "No existe"
    End If
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If 0 < Len(Combo1.Text) Then
        If Combo1.Text = 1 Then
            Text1.SetFocus
        Else
            Text5.SetFocus
        End If
    End If
End If
End Sub

Private Sub Command1_Click()
Form9.Hide
Call Limpiar
Text8.Text = 1
Form1.Show
End Sub

Private Sub Command1_GotFocus()
Command1.BackColor = RGB(100, 200, 255)
End Sub

Private Sub Command1_LostFocus()
Command1.BackColor = &H8000000F
End Sub

Private Sub Command2_Click()
If Len(Label2.Caption) = 0 Then
    MsgBox "No se ha introducido ningun codigo", vbExclamation
ElseIf Label2.Caption = "Articulo Nuevo" And Text3.Text = "" Then
    MsgBox "No ha digitado la descripción del codigo", vbExclamation
ElseIf Len(Combo1) = 0 Then
    MsgBox "Por favor seleccione un proveedor", vbExclamation
ElseIf Len(Text5.Text) = 0 Then
    MsgBox "Por favor digite la factura del producto", vbExclamation
ElseIf Len(Text4.Text) = 0 Then
    MsgBox "Por favor digite la cantidad", vbExclamation
ElseIf Len(Text12.Text) = 0 Then
    MsgBox "Por favor digite el limite de existencias", vbExclamation
ElseIf Len(Text10.Text) = 0 Then
    MsgBox "Digite el valor del articulo", vbExclamation
ElseIf Not IsNumeric(Text10.Text) Then
    MsgBox "Digite un valor valido", vbExclamation
ElseIf Len(Text9.Text) = 0 Then
    MsgBox "Digite % de Utilidad"
Else
    Call Guardar
End If
End Sub

Private Sub Command2_GotFocus()
Command2.BackColor = RGB(100, 200, 255)
End Sub

Private Sub Command2_LostFocus()
Command2.BackColor = &H8000000F
End Sub

Private Sub Form_Activate()
Set Tabla = Inventa.OpenRecordset("INVENTA")
Set Provedores = Inventa.OpenRecordset("PROVEDORES")
Set FacturaIN = Inventa.OpenRecordset("FACTURA DE INGRESO")
Set DetalleIN = Inventa.OpenRecordset("DETALLE DE INGRESO")
Provedores.Index = "PrimaryKey"
Provedores.MoveFirst
While Not Provedores.EOF
    Combo1.AddItem (Provedores!Codigo)
    Provedores.MoveNext
Wend
Call Limpiar
Text8.SelStart = 0
Text8.SelLength = Len(Text8.Text)
Text8.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Command1_Click
End Sub

Private Sub Label14_Click()
Form9.Hide
Form1.Show
End Sub

Private Sub Text10_Change()
Call Precio
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text9.SetFocus
End If
Dim KEY As String
KEY = Chr(KeyAscii)
If (KEY < "0" Or KEY > "9") And KEY <> "," Then
    If (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo1.SetFocus
End If
Dim KEY As String
KEY = Chr(KeyAscii)
If (KEY < "0" Or KEY > "9") And KEY <> "," Then
    If (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text10.SetFocus
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
    Text2.Text = Mid("0000000000", 1, 10 - Len(Text2.Text)) & Text2.Text
    Tabla.Index = "Codigo"
    Tabla.Seek "=", Text2.Text
    If Tabla.NoMatch Then
        Label2.Caption = "Articulo Nuevo"
        Text3.Locked = False
        Text3.Enabled = True
        Text6.Enabled = True
        Text7.Enabled = True
        Text3.SetFocus
    Else
        Label2.Caption = "Articulo Existente"
        Text3.Locked = True
        Text3.Text = Tabla!descrip
        If Not IsNull(Tabla!gabeta) Then
            Text6.Text = Tabla!gabeta
        Else
            Text6.Text = ""
        End If
        If Not IsNull(Tabla!ubicacion) Then
            Text7.Text = Tabla!ubicacion
        Else
            Text7.Text = ""
        End If
        Text9.Text = Tabla!Por_Venta
        Text4.SetFocus
    End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text12.SetFocus
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
    Call Listado
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

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text7.SetFocus
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command2.SetFocus
End If
End Sub

Private Sub Text8_GotFocus()
Text8.BackColor = RGB(255, 255, 255)
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text11.SelStart = 0
    Text11.SelLength = 40
    Text11.SetFocus
End If
Dim KEY As String
KEY = Chr(KeyAscii)
If (KEY < "0" Or KEY > "9") And KEY <> "," Then
    If (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text8_LostFocus()
Text8.BackColor = &H8000000F
End Sub

Private Sub Text9_Change()
Call Precio
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text6.Enabled And Text7.Enabled Then
        Text6.SetFocus
    Else
        Command2.SetFocus
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
