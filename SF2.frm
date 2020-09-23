VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAN chat - SERVER"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   ClipControls    =   0   'False
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   6720
      Top             =   6720
   End
   Begin VB.Frame FraCommercial 
      BackColor       =   &H80000009&
      Caption         =   "Commercial"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2160
      TabIndex        =   63
      Top             =   1200
      Width           =   6375
      Begin VB.TextBox TxtCommercial 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   67
         Top             =   2880
         Width           =   2895
      End
      Begin VB.ListBox lstCommercial 
         Height          =   3420
         Left            =   3360
         TabIndex        =   65
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label lblDelete 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Delete"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   71
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label lbladdmsg 
         BackColor       =   &H80000009&
         Caption         =   "If you want to add a message, type it and click 'Add'. If you want to delete a message, select it and click 'Delete'"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   70
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label lblAdd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   69
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label lblinsmsg 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Insert message here:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label lblcommsg 
         BackColor       =   &H80000009&
         Caption         =   "Commercial Messages:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   66
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label lblcominfo 
         BackColor       =   &H80000009&
         Caption         =   "With this, you can make the server display several messages on the commercial space provided at the top of the program."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   64
         Top             =   480
         Width           =   6135
      End
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Index           =   0
      Left            =   7680
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7200
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   8160
      Top             =   6720
   End
   Begin VB.Frame FraOptions 
      BackColor       =   &H80000014&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2160
      TabIndex        =   46
      Top             =   1200
      Width           =   6375
      Begin VB.TextBox TxtNick 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1200
         TabIndex        =   54
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox chkGaming5min 
         BackColor       =   &H80000014&
         Caption         =   "Set Status to Gaming when i'm not active for "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   2520
         Value           =   1  'Checked
         Width           =   4095
      End
      Begin VB.TextBox TxtTimeOut 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4320
         TabIndex        =   52
         Text            =   "5"
         Top             =   2520
         Width           =   375
      End
      Begin VB.CheckBox chkPlaySoundUserOL 
         BackColor       =   &H80000014&
         Caption         =   "Play sound when user gets On-Line"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   3120
         Value           =   1  'Checked
         Width           =   6135
      End
      Begin VB.CheckBox chkPlaySoundUserMSG 
         BackColor       =   &H80000014&
         Caption         =   "Play sound when user sends a message"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   3720
         Value           =   1  'Checked
         Width           =   6135
      End
      Begin VB.CheckBox chkSeePrograms 
         BackColor       =   &H80000014&
         Caption         =   "Allow other people to see which programs i'm running"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   4320
         Value           =   1  'Checked
         Width           =   6135
      End
      Begin VB.OptionButton OpMeOnLine 
         BackColor       =   &H80000014&
         Caption         =   "On-Line"
         Height          =   255
         Left            =   1200
         TabIndex        =   48
         Top             =   1560
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.OptionButton OpMeGaming 
         BackColor       =   &H80000014&
         Caption         =   "Gaming"
         Height          =   255
         Left            =   1200
         TabIndex        =   47
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label lblSpecifyuserandprivacy 
         BackColor       =   &H80000014&
         Caption         =   "Specify user and privacy options:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblNickName 
         BackStyle       =   0  'Transparent
         Caption         =   "NickName:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblMinutes 
         BackColor       =   &H80000014&
         Caption         =   "minutes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   58
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H80000014&
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblOK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "OK"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         TabIndex        =   56
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label lblCancel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5160
         TabIndex        =   55
         Top             =   5040
         Width           =   1095
      End
   End
   Begin VB.Frame FraChat 
      BackColor       =   &H80000014&
      Caption         =   "Chat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2160
      TabIndex        =   42
      Top             =   1200
      Width           =   6375
      Begin VB.TextBox TxtChat 
         Height          =   3975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         Top             =   480
         Width           =   6135
      End
      Begin VB.TextBox TxtMsg 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Top             =   4560
         Width           =   4935
      End
      Begin VB.Label lblSEND 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SEND"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5160
         TabIndex        =   45
         Top             =   4920
         Width           =   1095
      End
   End
   Begin VB.Frame FraUsers 
      BackColor       =   &H80000014&
      Caption         =   "Users"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2160
      TabIndex        =   37
      Top             =   1200
      Width           =   6375
      Begin VB.ListBox lstUserOnLine 
         Height          =   4260
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   39
         Top             =   960
         Width           =   2895
      End
      Begin VB.ListBox LstUserPrograms 
         Height          =   4260
         Left            =   3360
         TabIndex        =   38
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label lblUsersConnected 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Users connected:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label lblProgramsRunning 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Programs running:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3360
         TabIndex        =   40
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Frame FraAbout 
      BackColor       =   &H80000009&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2160
      TabIndex        =   29
      Top             =   1200
      Width           =   6375
      Begin VB.Label lblCredits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Credits:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lblTomDemets 
         BackColor       =   &H80000009&
         Caption         =   "Tom Demets"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label lblemail 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "tom_demets@hotmail.com"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   4680
         Width           =   2295
      End
      Begin VB.Label lblAboutCompany 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Hemoglobin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   2655
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   1455
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1440
         Width           =   6135
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   4
         Height          =   1215
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   1560
         Width           =   5895
      End
      Begin VB.Label lblLanChat 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "LAN Server"
         BeginProperty Font 
            Name            =   "Enviro"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   975
         Left            =   360
         TabIndex        =   32
         Top             =   1680
         Width           =   5655
      End
      Begin VB.Label LblVersion 
         BackColor       =   &H8000000E&
         Caption         =   "Version: 1.0"
         BeginProperty Font 
            Name            =   "Magneto"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   31
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label lblHemoYear 
         BackColor       =   &H8000000E&
         Caption         =   "Hemoglobin           2001-2002 "
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   5040
         Width           =   2415
      End
   End
   Begin VB.Frame FraHelp 
      BackColor       =   &H80000009&
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2160
      TabIndex        =   18
      Top             =   1200
      Width           =   6375
      Begin VB.Label lblControuble 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Connection trouble:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   6135
      End
      Begin VB.Label lblIfyoucant 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "If you can't connect to the server, there are three options:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   6135
      End
      Begin VB.Label lblOne 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "One:  the server is down"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   1200
         Width           =   5775
      End
      Begin VB.Label lblTwo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Two:  you mistyped the server IP adress"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   25
         Top             =   1560
         Width           =   4095
      End
      Begin VB.Label lblThree 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Three:  the connection port 7000 and following are busy"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   1920
         Width           =   4695
      End
      Begin VB.Label lblSollution 
         BackColor       =   &H80000009&
         Caption         =   "Sollution:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2640
         Width           =   6135
      End
      Begin VB.Label lblAskifServer 
         BackColor       =   &H80000009&
         Caption         =   "Ask if the server is down"
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   3000
         Width           =   5775
      End
      Begin VB.Label lblRetypeIP 
         BackColor       =   &H80000009&
         Caption         =   "If not: retype the IP adress"
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   3360
         Width           =   5415
      End
      Begin VB.Label lblDisableFW 
         BackColor       =   &H80000009&
         Height          =   735
         Left            =   1080
         TabIndex        =   20
         Top             =   3720
         Width           =   5175
      End
      Begin VB.Label lblPORTSCAN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click here to execute a portscan "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         Top             =   4560
         Width           =   2775
      End
   End
   Begin VB.Frame FraConnection 
      BackColor       =   &H80000009&
      Caption         =   "Connection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2160
      TabIndex        =   1
      Top             =   1200
      Width           =   6375
      Begin MSWinsockLib.Winsock Client 
         Left            =   5760
         Top             =   1200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.TextBox txtServerIP 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   13
         Text            =   "Insert Server IP"
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox txtLocalIP 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   11
         Text            =   "Insert local IP"
         Top             =   2040
         Width           =   2055
      End
      Begin VB.ComboBox cmbService 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "SF2.frx":0000
         Left            =   2760
         List            =   "SF2.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblConProgress 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   480
         TabIndex        =   16
         Top             =   3960
         Width           =   4335
      End
      Begin VB.Label lblCancelConnect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Disconnect"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label lblConnect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label lblServerIP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         Caption         =   "Server IP:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lblLocalip 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         Caption         =   "Local IP adress:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label lblService 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         Caption         =   "Service:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lbldescCon 
         BackStyle       =   0  'Transparent
         Caption         =   "Select the service you wish to connect to. Insert local IP and server's IP adress:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   6135
      End
   End
   Begin VB.Label lblmnuCommercial 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Commercial"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   62
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label lblCommercial 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "LAN chat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   61
      Top             =   240
      Width           =   6375
   End
   Begin VB.Label lblmnuUsers 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Users"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   2160
      X2              =   8520
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblmnuAbout 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label lblmnuHelp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label lblmnuChat 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CHAT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblmnuOptions 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   0
      X1              =   2160
      X2              =   8520
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblmnuConnection 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Connection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblCompany 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Hemoglobin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   2040
      X2              =   2040
      Y1              =   0
      Y2              =   6720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   0
      X1              =   2040
      X2              =   2040
      Y1              =   0
      Y2              =   6720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   6735
      Left            =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------
'This is LANchat, it was made by me: Tom Demets
' if you wish to contact me for any reason at all,
' send an e-mail to tom_demets@hotmail.com
'You can use the code if you want to, but i spent allot of work
'and time into this app, so i would appreciate it if you should
'mention my name in the credits.
'This said, have fun and GAME ON
'------------------------------------------------

'THIS IS THE SERVER
Dim PortScan As Boolean
Dim I As Integer
Dim Lannick As String
Dim OldNick As String
Dim StatusGaming As Boolean
Dim StatusOnLine As Boolean
Dim Gaming5min As Boolean
Dim PlaysoundUserMSG As Boolean
Dim PlaysoundUserOL As Boolean
Dim NumOfProcess As Long
Dim SeePrograms As Boolean
Dim Portscanning As Boolean
Dim FirstCon As Boolean
Dim Connected As Boolean
Dim InternetConnected As Boolean
Dim IData As Integer
Dim U As Integer
Dim ConPort As Integer
Dim RemItem As Integer
Dim C As Integer
Dim A As Integer
Dim Msg As String
Dim U0, U1, U2, U3, U4, U5, U6, U7, U8, U9, U10 As Integer 'max 10 user who could have connected and left
Dim AbandonedPort As Integer
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Private Const ERROR_SUCCESS = 0&
Private Const APINULL = 0&
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const RAS_MAXENTRYNAME As Integer = 256
Private Const RAS_MAXDEVICETYPE As Integer = 16
Private Const RAS_MAXDEVICENAME As Integer = 128
Private Const RAS_RASCONNSIZE As Integer = 412

Private Sub chkGaming5min_Click()
'if user has selected to set to gaming after 5min, set the bool val as true
If chkGaming5min.Value = 0 Then
Gaming5min = False
Else
Gaming5min = True
End If
Timer1.Enabled = True
'set a timer for some specified minutes, like msn, set away after i wasnt active for....minutes
End Sub

Private Sub chkPlaySoundUserMSG_Click()
'if user has selected to be warned when people send a msg, set the bool val as true
If chkPlaySoundUserMSG = 0 Then
PlaysoundUserMSG = False
Else
PlaysoundUserMSG = True
End If
End Sub

Private Sub chkPlaySoundUserOL_Click()
'if user has selected to be warned when people get online, set the bool val as true
If chkPlaySoundUserOL = 0 Then
PlaysoundUserOL = False
Else
PlaysoundUserOL = True
End If
End Sub

Private Sub chkSeePrograms_Click()
'if user has selected that other people can see the progs he/she is running, set the bool val as true
If chkSeePrograms = 0 Then
SeePrograms = False
Else
SeePrograms = True
End If
End Sub

Private Sub Client_Close()
Connected = False

Client.Close
DoEvents 'so the winsock can 'concentrate'  on just doing that
Me.Caption = "LAN chat - Not Connected"

lblmnuConnection.Enabled = True
lblmnuOptions.Enabled = False
lblmnuChat.Enabled = False
lblmnuHelp.Enabled = True
lblmnuAbout.Enabled = True
lblmnuUsers.Enabled = False
lblmnuCommercial.Enabled = False

FraAbout.Visible = False
FraChat.Visible = False
FraConnection.Visible = True
FraHelp.Visible = False
FraOptions.Visible = False
FraUsers.Visible = False
FraCommercial.Visible = False
End Sub

Private Sub Client_Connect()
If FirstCon = False Then
lblmnuConnection.Enabled = False
lblmnuOptions.Enabled = True
lblmnuChat.Enabled = False
lblmnuHelp.Enabled = True
lblmnuAbout.Enabled = True
lblmnuUsers.Enabled = False
lblmnuCommercial.Enabled = False
lblConProgress = lblConProgress & vbCrLf & "Connected at port " & IData
Me.Caption = "LAN chat - Connected"
Connected = True
Wait 2
Else
lblConProgress = lblConProgress & vbCrLf & "Connected at port 7000"
End If
End Sub

Private Sub Client_DataArrival(ByVal bytesTotal As Long)
If FirstCon = True Then 'ok, this is the tricky part, you want one winsock control
                                     'but it has to accept multiple cons.
Client.GetData IData 'this is an integer
DoEvents
Client.Close 'the principle is this, you make the client connect to a port that is always availlable on the server
DoEvents
Client.Connect txtServerIP, IData 'e.g.:port 7000 then, when the client connects, the server accepts it, creates a
DoEvents
FirstCon = False 'new socket listening on another port e.g: 7001 and sends data back that he is listening on that port
Exit Sub 'and you make the client REconnect to port 7001
End If

If FirstCon = False Then
Dim Data As String 'if you send the port data via String, it gives an error, so i splitted it in 2
Client.GetData Data 'gets the data sent by the the server, this is now string data
DoEvents
Dim ID As String 'every data sent by the server gots an ID so we can identify if the data wants us to do anything special

ID = Split(Data, "+")(0) 'this is the ID
 
If ID = "||CHAT||" Then
'this is a regular chat message
If PlaysoundUserMSG = True Then
'play a sound and display the msg
Module2.UserMSG
Msg = Split(Data, "+")(1)  'the data excists out of multiple parts, the second part is the message
TxtChat.Text = TxtChat.Text & vbCrLf & vbCrLf & Msg
Else
'do not play a sound, but display the msg
Msg = Split(Data, "+")(1)
TxtChat.Text = TxtChat.Text & vbCrLf & vbCrLf & Msg
End If
End If

If ID = "||USERCON||" Then
If PlaysoundUserOL = True Then
'play a sound, update the list and set the user in the chatbox
'play sound
Module2.UserOnline
User = Split(Data, "+")(1)
lstUserOnLine.AddItem User, lstUserOnLine.ListCount
TxtChat.Text = TxtChat.Text & vbCrLf & vbCrLf & User & " has connected to the LAN chat Server"
Else
'do not play a sound but update the list and set the user in the chatbox
User = Split(Data, "+")(1)
lstUserOnLine.AddItem User, lstUserOnLine.ListCount
TxtChat.Text = TxtChat.Text & vbCrLf & vbCrLf & User & " has connected to the LAN chat Server"
End If
End If

If ID = "||USERLEFT||" Then   'If someone has left the chat, he must be deleted out of the
Del = Split(Data, "+")(1) 'user presence list. DEL= the index of the user in the UserOnlineList
On Error Resume Next
Port = Split(Data, "+")(2) 'this is the abandoned port
If Port <> "" Or Port <> "0" Then
AbandonedPort = Port
PortShift
End If
On Error Resume Next
Dim N
For N = 0 To lstUserOnLine.ListCount
Nick = lstUserOnLine.List(N)
If Nick = Del Then 'found nick
lstUserOnLine.RemoveItem N
TxtChat.Text = TxtChat.Text & vbCrLf & vbCrLf & Nick & " has left LAN chat"
End If
Next N
I = I - 1 'user has left
End If

If ID = "||PROGRAM||" Then
PROGRAM = Split(Data, "+")(1) 'This is the program that is running on the other user

lstUserOnLine.ListIndex = U 'setting the correct user
LstUserPrograms.AddItem PROGRAM, LstUserPrograms.ListCount
End If

If ID = "||REQPROG||" Then
'update the programs that we are running on this terminal
If SeePrograms = True Then
NumOfProcess = Module1.GetActiveProcess
Dim T
For T = 1 To NumOfProcess
                                'PROGRAM LST      PROGRAM
    Client.SendData "||PROGRAM||+" & Module1.exePath(T)
    DoEvents
      Next T
Else
        Client.SendData "||PROGRAM||+" & "Unauthorized"
        DoEvents
End If
End If


If ID = "||COMMERCIAL||" Then 'this is the banner on top of the app, it says for now: LAN chat
COMMERCIAL = Split(Data, "+")(1) 'but, it can say all kinds of things, its especially usefull if you want
lblCommercial = COMMERCIAL 'people to join in a game
End If

End If
End Sub

Private Sub Client_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
lblConProgress = lblConProgress & vbCrLf & "Error while trying to connect"
Wait 2
lblConProgress = ""
Client.Close
DoEvents

lblmnuConnection.Enabled = True
lblmnuOptions.Enabled = False
lblmnuChat.Enabled = False
lblmnuHelp.Enabled = True
lblmnuAbout.Enabled = True
lblmnuUsers.Enabled = False
lblmnuCommercial.Enabled = False

FraAbout.Visible = False
FraChat.Visible = False
FraConnection.Visible = True
FraHelp.Visible = False
FraOptions.Visible = False
FraUsers.Visible = False
FraCommercial.Visible = False
End Sub

Private Sub Form_Activate()
With Tray
.cbSize = Len(Me.Caption)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon
.szTip = Me.Caption & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, Tray
End Sub

Private Sub Form_Load()
lblmnuConnection.ForeColor = &H0&
lblmnuOptions.ForeColor = &H0&
lblmnuChat.ForeColor = &H0&
lblmnuHelp.ForeColor = &H0&
lblmnuAbout.ForeColor = &H0&
lblmnuUsers.ForeColor = &H0&
lblmnuCommercial.ForeColor = &H0&

lblmnuConnection.Enabled = True
lblmnuOptions.Enabled = False
lblmnuChat.Enabled = False
lblmnuHelp.Enabled = True
lblmnuAbout.Enabled = True
lblmnuUsers.Enabled = False
lblmnuCommercial.Enabled = False

FraAbout.Visible = False
FraChat.Visible = False
FraConnection.Visible = True
FraHelp.Visible = False
FraOptions.Visible = False
FraUsers.Visible = False
FraCommercial.Visible = False

cmbService.AddItem "Internet Server", 0
cmbService.AddItem "LAN chat Server", 1
cmbService.ListIndex = 1

lstCommercial.AddItem "Welcome", 0
lstCommercial.AddItem "To", 1
lstCommercial.AddItem "LAN chat", 2

TxtChat.Text = TxtChat.Text & "Welcome to LAN chat"

I = 0 'number of connected winsocks=0
C = 0 'its the current commercial message being displayed right now, set to zero

ConPort = "7000"
Winsock1.Close
DoEvents
Winsock1.LocalPort = 7000
DoEvents
Winsock1.Listen
DoEvents

FirstCon = True 'set because the first time we receive port info, the second time we don't
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Msg As Long
Msg = X / Screen.TwipsPerPixelX
Select Case Msg
Case WM_LBUTTONDBLCLK:
Me.WindowState = vbNormal
Me.Show
Case WM_LBUTTONDOWN:

Case WM_LBUTTONUP:
                       
Case WM_RBUTTONDBLCLK:
            
Case WM_RBUTTONDOWN:

Case WM_RBUTTONUP:
                   
End Select
lblmnuConnection.ForeColor = &H0&
lblmnuOptions.ForeColor = &H0&
lblmnuChat.ForeColor = &H0&
lblmnuHelp.ForeColor = &H0&
lblmnuAbout.ForeColor = &H0&
lblmnuUsers.ForeColor = &H0&
lblmnuCommercial.ForeColor = &H0&
lblSEND.ForeColor = &H0&
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
Me.Visible = False
End If
End Sub

Private Sub Form_Terminate()
Unload Form1
Form_Unload 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Connected = True Then
GoBack:
If Lannick = "" Then
Lannick = "Server"
Client.SendData "||USERCON||" & "+" & Lannick
DoEvents
'now a lannick has been set, the routine will also do the else statement
GoTo GoBack
Else
Client.SendData "||USERLEFT||" & "+" & Lannick & "+" & IData
DoEvents
End If
End If
Client.Close
Shell_NotifyIcon NIM_DELETE, Tray
MsgBox "Don't forget to vote!!", vbInformation, "LAN chat"
End
End Sub

Private Sub FraChat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSEND.ForeColor = &H0&
End Sub

Private Sub FraCommercial_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmnuConnection.ForeColor = &H0&
lblmnuOptions.ForeColor = &H0&
lblmnuChat.ForeColor = &H0&
lblmnuHelp.ForeColor = &H0&
lblmnuAbout.ForeColor = &H0&
lblmnuUsers.ForeColor = &H0&
lblAdd.ForeColor = &H0&
lblDelete.ForeColor = &H0&
lblmnuCommercial.ForeColor = &H0&
End Sub

Private Sub FraConnection_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmnuConnection.ForeColor = &H0&
lblmnuOptions.ForeColor = &H0&
lblmnuChat.ForeColor = &H0&
lblmnuHelp.ForeColor = &H0&
lblmnuAbout.ForeColor = &H0&
lblmnuUsers.ForeColor = &H0&
lblmnuCommercial.ForeColor = &H0&
lblConnect.ForeColor = &H0&
lblCancelConnect.ForeColor = &H0&
End Sub

Private Sub FraHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPORTSCAN.ForeColor = &H0&
End Sub

Private Sub FraOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOK.ForeColor = &H0&
lblCancel.ForeColor = &H0&
End Sub

Private Sub lblAdd_Click()
lstCommercial.AddItem TxtCommercial, lstCommercial.ListCount 'adding a commercial message to the list
TxtCommercial.Text = ""
End Sub

Private Sub lblAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAdd.ForeColor = &HFF&
lblDelete.ForeColor = &H0&
End Sub

Private Sub lblCancel_Click()
If Lannick = "" Then 'when user presses cancel, and the nick is not set, we have a nick of "" so, we don't allow that
MsgBox "Enter your Nickname", vbInformation, "LAN chat"
End If
End Sub

Private Sub lblCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOK.ForeColor = &H0&
lblCancel.ForeColor = &HFF&
End Sub

Private Sub lblCancelConnect_Click()
'speaks for itself
lblConProgress = lblConProgress & vbCrLf & "Cancelling connection request."
Client.Close
DoEvents
lblConProgress = lblConProgress & vbCrLf & "Disconnected"
Wait 2
lblConProgress = ""
End Sub

Private Sub lblCancelConnect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCancelConnect.ForeColor = &HFF&
lblConnect.ForeColor = &H0&
End Sub

Private Sub lblConnect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblConnect.ForeColor = &HFF&
lblCancelConnect.ForeColor = &H0&
End Sub

Private Sub lblDelete_Click()
On Error Resume Next
lstCommercial.RemoveItem RemItem 'when you click on the commercial list, it sets an index of what you clicked so the app knows what to delete
If lstCommercial.ListCount = 0 Then
lstCommercial.AddItem "LAN chat", ListCount
C = 0 'makes the commercial msg's begin from the top
End If
End Sub

Private Sub lblDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAdd.ForeColor = &H0&
lblDelete.ForeColor = &HFF&
End Sub

Private Sub lblmnuAbout_Click()
FraAbout.Visible = True
FraChat.Visible = False
FraConnection.Visible = False
FraHelp.Visible = False
FraOptions.Visible = False
FraUsers.Visible = False
FraCommercial.Visible = False
End Sub

Private Sub lblmnuChat_Click()
FraAbout.Visible = False
FraChat.Visible = True
FraConnection.Visible = False
FraHelp.Visible = False
FraOptions.Visible = False
FraUsers.Visible = False
FraCommercial.Visible = False
End Sub

Private Sub lblmnuCommercial_Click()
FraAbout.Visible = False
FraChat.Visible = False
FraConnection.Visible = False
FraHelp.Visible = False
FraOptions.Visible = False
FraUsers.Visible = False
FraCommercial.Visible = True
End Sub

Private Sub lblmnuCommercial_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmnuConnection.ForeColor = &H0&
lblmnuOptions.ForeColor = &H0&
lblmnuChat.ForeColor = &H0&
lblmnuHelp.ForeColor = &H0&
lblmnuAbout.ForeColor = &H0&
lblmnuUsers.ForeColor = &H0&
lblmnuCommercial.ForeColor = &HFF&
End Sub

Private Sub lblmnuConnection_Click()
FraAbout.Visible = False
FraChat.Visible = False
FraConnection.Visible = True
FraHelp.Visible = False
FraOptions.Visible = False
FraUsers.Visible = False
FraCommercial.Visible = False
End Sub

Private Sub lblmnuconnection_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmnuConnection.ForeColor = &HFF&
lblmnuOptions.ForeColor = &H0&
lblmnuChat.ForeColor = &H0&
lblmnuHelp.ForeColor = &H0&
lblmnuAbout.ForeColor = &H0&
lblmnuUsers.ForeColor = &H0&
lblmnuCommercial.ForeColor = &H0&
End Sub

Private Sub lblmnuHelp_Click()
FraAbout.Visible = False
FraChat.Visible = False
FraConnection.Visible = False
FraHelp.Visible = True
FraOptions.Visible = False
FraUsers.Visible = False
FraCommercial.Visible = False
End Sub

Private Sub lblmnuOptions_Click()
FraAbout.Visible = False
FraChat.Visible = False
FraConnection.Visible = False
FraHelp.Visible = False
FraOptions.Visible = True
FraUsers.Visible = False
FraCommercial.Visible = False
End Sub

Private Sub lblmnuoptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmnuOptions.ForeColor = &HFF&
lblmnuConnection.ForeColor = &H0&
lblmnuChat.ForeColor = &H0&
lblmnuHelp.ForeColor = &H0&
lblmnuAbout.ForeColor = &H0&
lblmnuUsers.ForeColor = &H0&
lblmnuCommercial.ForeColor = &H0&
End Sub
Private Sub lblmnuchat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmnuChat.ForeColor = &HFF&
lblmnuConnection.ForeColor = &H0&
lblmnuOptions.ForeColor = &H0&
lblmnuHelp.ForeColor = &H0&
lblmnuAbout.ForeColor = &H0&
lblmnuUsers.ForeColor = &H0&
lblmnuCommercial.ForeColor = &H0&
End Sub

Private Sub lblmnuhelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmnuHelp.ForeColor = &HFF&
lblmnuConnection.ForeColor = &H0&
lblmnuOptions.ForeColor = &H0&
lblmnuChat.ForeColor = &H0&
lblmnuAbout.ForeColor = &H0&
lblmnuUsers.ForeColor = &H0&
lblmnuCommercial.ForeColor = &H0&
End Sub

Private Sub lblmnuAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmnuAbout.ForeColor = &HFF&
lblmnuConnection.ForeColor = &H0&
lblmnuOptions.ForeColor = &H0&
lblmnuChat.ForeColor = &H0&
lblmnuHelp.ForeColor = &H0&
lblmnuUsers.ForeColor = &H0&
lblmnuCommercial.ForeColor = &H0&
End Sub

Private Sub lblConnect_Click()
On Error Resume Next

If cmbService.ListIndex = 0 Then 'when the connect via Internet was selected
InternetConnected = InternetGetConnectedState(0&, 0&) 'check if we are allready connected
If InternetConnected = True Then 'yes, we are connected
Client.Close
DoEvents
Client.Connect txtServerIP, 7000
lblConProgress = lblConProgress & vbCrLf & "Trying to connect to " & txtServerIP & " at port 7000"
Else 'no, we are not connected
con = MsgBox("You are not connected to the internet! Would you like to make a connection?", vbInformation + vbYesNo, "LAN chat")
If con = vbYes Then 'if user wants to make a con
InternetAutodial 0, 0 'default screen to dial in on your ISP (don't know if this works for ADSL or Cable, since you are always on-line (i guess))
End If
If con = vbNo Then 'user doesnt want to make a con to the internet
'so do nothing
End If
End If
End If

If cmbService.ListIndex = 1 Then 'if user has selected LAN
Client.Close
DoEvents
Client.Connect txtServerIP, 7000 'connect to server via alwayz available port: 7000
lblConProgress = lblConProgress & vbCrLf & "Trying to connect to " & txtServerIP & " at port 7000"
End If
End Sub

Public Sub Wait(Pause As Integer)
'not mine from MS Excel => timer but modified
Dim PauseTime, Start, Finish, TotalTime
    PauseTime = Pause
    Start = Timer
    Do While Timer < Start + PauseTime
        DoEvents
    Loop
    Finish = Timer
End Sub

Private Sub lblmnuUsers_Click()
FraAbout.Visible = False
FraChat.Visible = False
FraConnection.Visible = False
FraHelp.Visible = False
FraOptions.Visible = False
FraUsers.Visible = True
FraCommercial.Visible = False
End Sub

Private Sub lblmnuUsers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmnuAbout.ForeColor = &H0&
lblmnuConnection.ForeColor = &H0&
lblmnuOptions.ForeColor = &H0&
lblmnuChat.ForeColor = &H0&
lblmnuHelp.ForeColor = &H0&
lblmnuUsers.ForeColor = &HFF&
lblmnuCommercial.ForeColor = &H0&
End Sub

Private Sub lblOK_Click()
'first set all the options, or else we won't know if we need to play a sound or not :)
If chkPlaySoundUserMSG.Value = 1 Then 'option: playsound when user sends msg: yes or no
PlaysoundUserMSG = True
Else
PlaysoundUserMSG = False
End If

If chkPlaySoundUserOL.Value = 1 Then 'option: playsound when user connets
PlaysoundUserOL = True
Else
PlaysoundUserOL = False
End If

If chkSeePrograms.Value = 1 Then 'option: if others can see the programs that you are running
SeePrograms = True
Else
SeePrograms = False
End If

'nick
If TxtNick.Text = "" Then
MsgBox "Enter your Nickname", vbInformation, "LAN chat" 'if user did not fill in something
Else
Lannick = TxtNick 'if user has filled in something

If OldNick = "" Then              'this is also kind of tricky:
Client.SendData "||USERCON||" & "+" & Lannick
DoEvents
Else                                        'first, if its the first time a user sets his nick, then don't keep track of it
If OldNick <> Lannick Then      'if the previous nick is not the same as the nick we just filled in then get rid of it
                            
                            'ID                              'Oldnick name
Client.SendData "||USERLEFT||" & "+" & OldNick
DoEvents
'we are just changing nick so we don't need to send abandoned port data
Client.SendData "||USERCON||" & "+" & Lannick
DoEvents
End If
End If

If OpMeGaming = True Then
StatusGaming = True
StatusOnLine = False
Client.SendData "||CHAT||" & "+" & Lannick & " is now gaming" 'sets our status to gaming and notifies other users that you are away
DoEvents 'so winsock only concentrates on this and doesn't mix this with other data (if you have some trouble with sending data, this could be your sollution)
End If
If OpMeOnLine = True Then
StatusGaming = False
StatusOnLine = True
Client.SendData "||CHAT||" & "+" & Lannick & " has joined" 'sets our stat to online and notifies other users that you are online
DoEvents
End If

'just buttons
lblmnuConnection.Enabled = False
lblmnuOptions.Enabled = True
lblmnuChat.Enabled = True
lblmnuHelp.Enabled = True
lblmnuAbout.Enabled = True
lblmnuUsers.Enabled = True
lblmnuCommercial.Enabled = True

'this is dependant from the options so, we save this for the last
If Gaming5min = True Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If

'update the oldnick so that we can compare the next time we press ok
OldNick = TxtNick.Text
End If
End Sub

Private Sub lblOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOK.ForeColor = &HFF&
lblCancel.ForeColor = &H0&
End Sub

Private Sub lblPORTSCAN_Click()
If Connected = True Then
MsgBox "You are allready connected to the server, can't run portscan", vbInformation, "LAN chat"
Else
Dim A, Z
Z = 0
For A = 7000 To 9000 - 1 'specify the ports (i dont think there will be more than 2000 people at the LAN , if there are, just change the last port) :)
Portscanning = True
'this peace is also not mine, but don't know who he/she is, if you see this, then contact me and i will put your name here
On Error GoTo error 'if we have an error, it means we found an open port
Client.Close
Client.LocalPort = A
Client.Listen
DoEvents
Next A
error:
If Err.Number = 10048 Then
MsgBox "Open port found: " & A, vbCritical, "LAN chat - Ports left: " & 7200 - A
Z = Z + 1
End If
Resume Next
MsgBox "Portscan complete. Open ports found: " & Z, vbInformation, "LAN chat - Scanned ports 7000 to 7200"
Portscanning = False
End If
End Sub

Private Sub lblPORTSCAN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPORTSCAN.ForeColor = &HFF&
End Sub

Private Sub lblSEND_Click()
CurTime = 0
If Gaming5min = True Then
Timer1.Enabled = True
End If
If TxtMsg <> "" Then 'this kinda explains itself
                            'ID                 'Nick                   'MSG
Client.SendData "||CHAT||" & "+" & Lannick & ": " & TxtMsg
TxtMsg.Text = ""
End If
End Sub

Private Sub lblSEND_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSEND.ForeColor = &HFF&
End Sub

Private Sub lstCommercial_Click()
RemItem = lstCommercial.ListIndex 'index the item so we can delete it later
End Sub

Private Sub lstUserOnLine_Click()
U = lstUserOnLine.ListIndex
Client.SendData "||REQPROG||"  'this means we want to request what others are running
DoEvents                                        'especially usefull to know what others are doing like playing counterstrike, or MOHAA
LstUserPrograms.Clear
End Sub

Private Sub Timer1_Timer()
CurTime = CurTime + 1

If CurTime = TxtTimeOut.Text Then 'if time has expired then
'set to gaming and send msg to server that we are gaming

Gaming5min = True
OpMeGaming.Value = True
OpMeOnLine.Value = False
Timer1.Enabled = False
Client.SendData "||CHAT||" & "+" & Lannick & " is now gaming"
DoEvents
Else
Gaming5min = False
End If

End Sub

Private Sub Timer2_Timer()
If Connected = True Then
'this is the commercial timer, if it triggers, we send a commercial message to the server
'from there, its sends it to every user connected
If lstCommercial.ListCount <= C Then C = 0
Client.SendData "||COMMERCIAL||" & "+" & lstCommercial.List(C) 'index of the com list message
DoEvents
C = C + 1

End If
End Sub

Private Sub TxtChat_Change()
'not mine, i found it on PSC but, i forgot who wrote it... if you see this and it's yours,
'please give me some proof you wrote it (i don't care what). I will update
'my submission and you will be standing here instead of these stupid lines of txt
On Error Resume Next
    TxtChat.SelLength = 0
    If Len(TxtChat.Text) > 0 Then
        If Right$(TxtChat.Text, 1) = vbCrLf Then
            TxtChat.SelStart = Len(TxtChat.Text) - 1
            Exit Sub
        End If
        TxtChat.SelStart = Len(TxtChat.Text)
    End If
End Sub

Private Sub txtLocalIP_Click()
If txtLocalIP = "Insert local IP" Then
txtLocalIP = ""
Else
End If
End Sub

Private Sub TxtMsg_KeyPress(KeyAscii As Integer)
CurTime = 0
If Gaming5min = True Then
Timer1.Enabled = True
End If
If TxtMsg <> "" Then
If KeyAscii = 13 Then
lblSEND_Click
KeyAscii = 0
TxtMsg.Text = ""
End If
Else
If KeyAscii = 13 Then
KeyAscii = 0
TxtMsg.Text = ""
End If
End If
End Sub

Private Sub txtServerIP_Click()
If txtServerIP = "Insert Server IP" Then
txtServerIP = ""
Else
End If
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If U9 = "" Then                                     'noboby has left the chat

I = I + 1                                               'ok, here's the magic, i admitt, there are
ConPort = ConPort + 1                           'probably better ways, but this is my version
Winsock1.Close                                      'first make an index of all connection requests
DoEvents                                            'then specify the conport eg: 7000+1=7001 (first user who has connected)
Winsock1.Accept (requestID)             'then the second user will connect at 7000+2 because I will be 2 by then
DoEvents                                            'accept the id
Winsock1.SendData (ConPort)             'send the port the client has to connect on
DoEvents

Load Winsock2(I)                            'this is the code form Dustin Davis but, it wasn't processed in a chat program so i made my version (I REPEAT NOT THE SAME CODE)

DoEvents                                            'this is my code, its more difficult but it works fine
Winsock2(I).Close                               'closing the newly created winsock to avoid errors
DoEvents                                            'still closing
Winsock2(I).LocalPort = ConPort
DoEvents
Winsock2(I).Listen                              'listening on the conport eg: 7001
DoEvents

Else                                    'Somebody has allready left the chat and we need to re-use the port
I = I + 1                                               'someone has connected
Winsock1.Close                                      'Close it to avoid error
DoEvents
Winsock1.Accept (requestID)                 'accept the con
DoEvents
Winsock1.SendData (U9)                      'send the buffered and abandoned port, so we can 'recycle' it
DoEvents

DoEvents                                            'this is my code, its more difficult but it works fine
Winsock2(I).Close                               'closing the newly created winsock to avoid errors
DoEvents                                            'still closing
Winsock2(I).LocalPort = U9
DoEvents
Winsock2(I).Listen                              'listening on the conport eg: 7001
DoEvents
U9 = ""                                                'set the buffer back to "" so a new buffer can be filled
'AbandonedPort = ""
End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'this is when we are in big sh*t, the server has crashed and can't accept new connections
Me.Caption = "Winsock listening on port 7000 is down!"

Winsock1.Close          'reboot the listening sock
DoEvents
Winsock1.LocalPort = 7000
DoEvents
Winsock1.Listen
DoEvents
Wait 2
Me.Caption = "Server rebooted"
Wait 2
Me.Caption = "LAN chat - SERVER"
End Sub

Private Sub Winsock2_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Winsock2(I).Close                               'this is the sequel for the client, this winsock is listening
DoEvents                                            'the conport and it needs to initialize
Winsock2(I).Accept requestID            'listen for connection from client
DoEvents                                            'when we get this, it means that the client is connecting to us, and that we can
Winsock1.Close                                      'reset the available listening winsock at port 7000
DoEvents
Winsock1.LocalPort = 7000
DoEvents
Winsock1.Listen
DoEvents
End Sub

Private Sub Winsock2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next 'this is to prevent that the server is going down because a user has disconnected
Dim Data As String
Call Winsock2(Index).GetData(Data, , bytesTotal) 'this is the actual server, it gets the data from who sends it
DoEvents
Dim E
For E = 1 To I
Winsock2(E).SendData Data           'and sends it to everyone connected... it doesn't care about the types of data
                                                        'like ||chat|| or something like that, it lets the client worry about that
DoEvents 'so that the winsock sends its data and does nothing else!
Next E
End Sub

Private Sub Winsock2_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'this means, that someone has left LANchat permanently (disconnected) or that we are in BIG ....
Me.Caption = "Winsock down at port " & 7000 + Index
'notifies the server
Beep
I = I - 1 'a user has left
AbandonedPort = 7000 + Index
'-------------------------------------------
PortShift
Wait 2
Me.Caption = "LAN chat - SERVER"
End Sub

Private Sub PortShift()
'we should index the winsocks closed, or else if someone closes and goes back online we have a lot of open ports who can't be used again
If U0 = "" Then                 '\\\\\\\\\\\\\\\\\\
U0 = AbandonedPort       'every U is a memory location for a port.
End If                                'As you can see, if we call portshift, the data goes
If U1 = "" Then                 'straight to U9. If there are eg: 3 people disconnecting
U1 = AbandonedPort      'but no one connects, we lose a lot of ports. cause the port
U0 = ""                              'index counts up.
End If                               'For the people who still don't understand it :)
If U2 = "" Then                '5 People connect
U2 = AbandonedPort     'Ports open from 7001 to 7006
U1 = ""                             '3 people disconnect, these ports would normally stay open
End If                               'The ports they were connected on,
If U3 = "" Then                 'are sent to the server, who saves them in this buffer
U3 = AbandonedPort      'If this buffer WOULDN'T exist, and someone wants to reconnect
U2 = ""                             'he/she would reconnect @ port 7007 and the ports 7001 to 7007 would
End If                              'stay open. This is not good especially for the server... if he is gaming
If U4 = "" Then                 'and because of all the open ports, he/she can't game anymore, it would be a realy
U4 = AbandonedPort      'stupid program
U3 = ""                             'WITH the buffer: the abandoned ports can be re-used so the server
End If                              'has no trouble handling all the connections because there will never be a port that is
If U5 = "" Then                 'wasted
U5 = AbandonedPort
U4 = ""
End If
If U6 = "" Then
U6 = AbandonedPort
U5 = ""
End If
If U7 = "" Then
U7 = AbandonedPort
U6 = ""
End If
If U8 = "" Then
U8 = AbandonedPort
End If
If U9 = "" Then
U9 = AbandonedPort
'-------------------------------------------
End If
End Sub
