VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anular Factura"
   ClientHeight    =   1935
   ClientLeft      =   8745
   ClientTop       =   1440
   ClientWidth     =   5415
   Icon            =   "AnularFacturas.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Anular"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Documento"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.OptionButton Option3 
         Caption         =   "Proforma"
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Factura"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Factura 1"
         Height          =   255
         Left            =   4800
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Factura a anular"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2295
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Limpiar()
Option1.Value = True
Option2.Value = False
Option3.Value = False
Text1.Text = ""
End Sub
Private Sub Command1_Click()
Set Facturas = Inventa.OpenRecordset("FACTURAS")
Set Detalle = Inventa.OpenRecordset("Detalle")
Set Tabla = Inventa.OpenRecordset("Inventa")
Facturas.Index = "PrimaryKey"
Tabla.Index = "PrimaryKey"
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
    If Facturas.NoMatch = False Then
        If Facturas!Anulada = False Then
            Factura = Tipo & Text1.Text
            Facturas.Edit
            Facturas!Anulada = True
            Facturas.Update
            MsgBox "Factura anulada", vbInformation
            ok = True
            Detalle.Index = "Detalle"
            Detalle.Seek "=", Factura
            While ok = True
                Tabla.Seek "=", Detalle!Codigo
                Tabla.Edit
                Tabla!Cantidad = Tabla!Cantidad + Detalle!Cantidad
                Tabla.Update
                Detalle.MoveNext
                If Detalle.EOF Then
                    ok = False
                Else
                    If Detalle!Num_Fac <> Facturas!Num_Fac Then
                        ok = False
                    End If
                End If
            Wend
            Call Command2_Click
        Else
            MsgBox "Esa factura anulada ya está anulada", vbInformation
        End If
    Else
        MsgBox "Número de factura invalido", vbCritical
    End If
End If
End Sub

Private Sub Command2_Click()
Option1.Value = False
Option2.Value = False
Option3.Value = False
Text1.Text = ""
Form4.Hide
Form1.Show
End Sub

Private Sub Form_Activate()
Call Limpiar
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Command2_Click
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
    Call Command1_Click
End If
End Sub
