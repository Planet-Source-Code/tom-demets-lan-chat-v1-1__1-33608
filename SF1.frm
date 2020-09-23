VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   120
      Top             =   2520
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "Version: 1.1"
      BeginProperty Font 
         Name            =   "Magneto"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "LAN Server"
      BeginProperty Font 
         Name            =   "Enviro"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   4
      Height          =   855
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1095
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Hemoglobin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If App.PrevInstance = True Then
MsgBox "LAN chat is allready activated", vbInformation, "LAN chat"
End
Else
MsgBox "PLEASE VOTE FOR ME", vbInformation, "LAN chat"
End If
End Sub

Private Sub Timer1_Timer()
Unload Me
Form2.Show
End Sub
