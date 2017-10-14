VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EACPsys"
   ClientHeight    =   4875
   ClientLeft      =   540
   ClientTop       =   1110
   ClientWidth     =   6180
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   412
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1560
      Top             =   3360
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   3360
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4500
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   2
            Enabled         =   0   'False
            Object.Width           =   5671
            MinWidth        =   1411
            Text            =   "E. A. C. P. Systems"
            TextSave        =   "E. A. C. P. Systems"
            Object.ToolTipText     =   "E. A. C. P. Systems"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2593
            TextSave        =   "02/10/2017"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "04:34:PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   1560
      Picture         =   "Menu.frx":25CA
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   840
      Picture         =   "Menu.frx":2A0C
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   120
      Picture         =   "Menu.frx":314E
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "Menu.frx":3590
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "Menu.frx":39D2
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   1440
      Left            =   1560
      Picture         =   "Menu.frx":3E14
      Top             =   1440
      Width           =   2970
   End
   Begin VB.Menu Inventario 
      Caption         =   "&Inventario"
      Begin VB.Menu Consultas 
         Caption         =   "Consultas"
         Shortcut        =   {F1}
      End
      Begin VB.Menu Mercaderia 
         Caption         =   "Ingreso de Mercaderia"
         Shortcut        =   ^M
      End
      Begin VB.Menu Inven 
         Caption         =   "Mantenimienton de Inventario"
         Shortcut        =   ^I
      End
      Begin VB.Menu Aumento 
         Caption         =   "Aplicar aumento..."
         Shortcut        =   ^J
      End
   End
   Begin VB.Menu Facturas 
      Caption         =   "Facturación"
      Begin VB.Menu Facturación 
         Caption         =   "Facturar..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu Anular_facturas 
         Caption         =   "Anular facturas..."
         Shortcut        =   ^A
      End
      Begin VB.Menu Devolucion 
         Caption         =   "Devolución..."
         Shortcut        =   ^D
      End
      Begin VB.Menu Consultas_de_factura 
         Caption         =   "Consultar factura..."
         Shortcut        =   ^O
      End
      Begin VB.Menu FacturaPersonalizada 
         Caption         =   "Factura Personalizada"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu Cuentas_por_cobrar 
      Caption         =   "Cuentas &por Cobrar"
      Begin VB.Menu AbonoFacturas 
         Caption         =   "Abono de Facturas"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu Reportes 
      Caption         =   "&Reportes"
   End
   Begin VB.Menu Configuracion 
      Caption         =   "&Configuración"
      Begin VB.Menu Parametros 
         Caption         =   "Parámetros"
         Shortcut        =   ^E
      End
      Begin VB.Menu Vendedores 
         Caption         =   "Vendedores"
         Shortcut        =   ^V
      End
      Begin VB.Menu Clientes 
         Caption         =   "Clientes"
         Shortcut        =   ^C
      End
      Begin VB.Menu Proveedores 
         Caption         =   "Proveedores"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu Cerrar 
      Caption         =   "&Cerrar"
      Begin VB.Menu Sesion 
         Caption         =   "Cambiar inicio de sesión"
         Shortcut        =   ^U
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Atto As Date
Dim Ima As Date
Dim Soshite As Integer

Private Sub AbonoFacturas_Click()
If Not Form0.Option1.Value Then
    Form1.Hide
    Form8.Show
Else
    MsgBox "Acceso Restringido", vbCritical
End If
End Sub

Private Sub Anular_facturas_Click()
If Form0.Option1.Value Then
    MsgBox "Acceso Restringido", vbCritical
Else
    Form1.Hide
    Form4.Show
End If
End Sub

Private Sub Aumento_Click()
If Form0.Option3.Value Then
    Form1.Hide
    Form16.Show
Else
    MsgBox "Acceso Restringido", vbCritical
End If
End Sub

Private Sub Clientes_Click()
If Form0.Option3.Value Then
    Form1.Hide
    Form11.Show
    Form11.Text1.Enabled = True
    Form11.Command1.Enabled = True
Else
    Form1.Hide
    Form11.Show
    Form11.Text1.Enabled = False
    Form11.Command1.Enabled = False
End If
End Sub

Private Sub Consultas_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub Consultas_de_factura_Click()
Form1.Hide
Form10.Show
End Sub

Private Sub Devolucion_Click()
If Not Form0.Option1.Value Then
    Form1.Hide
    Form12.Show
Else
    MsgBox "Acceso Restringido", vbCritical
End If
End Sub

Private Sub Facturación_Click()
Form1.Hide
Form3.Show
End Sub

Private Sub FacturaPersonalizada_Click()
If Not Form0.Option1.Value Then
    Form1.Hide
    Form13.Show
Else
    MsgBox "Acceso Restringido", vbCritical
End If
End Sub

Private Sub Form_Activate()
If Not Form0.Option1 Then
    Atto = Time
    Timer2.Enabled = True
End If
If "J:\Pruebas.mdb" = Inventa.Name Then
    Image6.Visible = True
Else
    Image6.Visible = False
End If
End Sub

Private Sub Form_Deactivate()
If Not Form0.Option1 Then
    Timer2.Enabled = False
End If
End Sub

Private Sub Form_Load()
Form1.BackColor = &H0&
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Salir_Click
End Sub

Private Sub Inven_Click()
If Not Form0.Option1.Value Then
    Form1.Hide
    Form5.Show
Else
    MsgBox "Acceso Restringido", vbCritical
End If
End Sub

Private Sub Mercaderia_Click()
If Not Form0.Option1.Value Then
    Form1.Hide
    Form9.Show
Else
    MsgBox "Acceso Restringido", vbCritical
End If
End Sub

Private Sub Parametros_Click()
If Form0.Option3.Value Then
    Form1.Hide
    Form7.Show
Else
    MsgBox "Acceso Restringido", vbCritical
End If
End Sub

Private Sub Proveedores_Click()
If Form0.Option3.Value Then
    Form1.Hide
    Form6.Show
    Form6.Command1.Enabled = True
Else
    Form1.Hide
    Form6.Show
    Form6.Command1.Enabled = False
End If
End Sub

Private Sub Reportes_Click()
Form14.Show
Me.Hide
End Sub

Private Sub Salir_Click()
End
End Sub

Private Sub Sesion_Click()
Form1.Hide
Form0.Show
End Sub

Private Sub Timer1_Timer()
If Form0.Option1.Value Then
    If Not Image2.Visible Then
        Image2.Visible = True
        Image3.Visible = False
        Image4.Visible = False
    End If
End If
If Form0.Option2.Value Then
    If Not Image3.Visible Then
        Image2.Visible = False
        Image3.Visible = True
        Image4.Visible = False
    End If
End If
If Form0.Option3.Value Then
    If Not Image4.Visible Then
        Image2.Visible = False
        Image3.Visible = False
        Image4.Visible = True
    End If
End If
End Sub

Private Sub Timer2_Timer()
If Form0.Option2 Or Form0.Option3 Then
    Ima = Time
    Soshite = Minute(Ima - Atto)
    If Soshite = 5 Then
        Form0.Option1 = True
        Form0.Option2 = False
        Form0.Option3 = False
        Timer2.Enabled = False
    End If
End If
End Sub

Private Sub Vendedores_Click()
If Form0.Option3 Then
    Form1.Hide
    Form15.Show
End If
End Sub
