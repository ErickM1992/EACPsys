VERSION 5.00
Begin VB.Form ProvedorManager 
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Text            =   "Text5"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   765
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "ProvedorManager.frx":0000
      Top             =   1080
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   600
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   120
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   4815
      Begin VB.CommandButton Command2 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Telefono 2:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Telefono 1:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Dirección:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Provedor:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "ProvedorManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
