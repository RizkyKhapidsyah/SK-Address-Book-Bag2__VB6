VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form DIALER1 
   Caption         =   "MSComm Phone Dialer"
   ClientHeight    =   1545
   ClientLeft      =   4005
   ClientTop       =   7350
   ClientWidth     =   4275
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1545
   ScaleWidth      =   4275
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   2295
      Begin VB.OptionButton OptionCom4 
         Caption         =   "COM4"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptionCom1 
         Caption         =   "COM1"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptionCom2 
         Caption         =   "COM2"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
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
      Height          =   348
      Left            =   1680
      TabIndex        =   1
      Top             =   7000
      Width           =   852
   End
   Begin VB.CommandButton DialButton 
      Caption         =   "Dial"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   720
      TabIndex        =   0
      Top             =   7000
      Visible         =   0   'False
      Width           =   852
   End
End
Attribute VB_Name = "DIALER1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------
'   DIALER.FRM
'   Copyright (c) 1994 Crescent Software, Inc.
'   by Carl Franklin
'
'   Updated by Anton de Jong
'
'   Demonstrates how to dial phone numbers with a modem.
'
'   For this program to work, your telephone and
'   modem must be connected to the same phone line.
'--------------------------------------------------------
Option Explicit

DefInt A-Z
Dim CancelFlag

Private Sub CancelButton_Click()
    ' CancelFlag tells the Dial procedure to exit.
   ' CancelFlag = True
    'CancelButton.Enabled = False
End Sub



Private Sub DialButton_Click()
    Dim Number$, Temp$
    
    DialButton.Enabled = False
    
    CancelButton.Enabled = True
    
    ' Get the number to dial.
    Number$ = Form10b11a.txtPhone.Text 'InputBox$("Enter phone number:", Number$)
        If Number$ = "" Then Exit Sub
    'Temp$ = Status
    'Status = "Dialing - " + Number$
    
    ' Dial the selected phone number.
   Dial Number$

    DialButton.Enabled = True
   
    CancelButton.Enabled = False

   
End Sub

Private Sub Form_Load()
    ' Setting InputLen to 0 tells MSComm to read the entire
    ' contents of the input buffer when the Input property
    ' is used.
    MSComm1.InputLen = 0
    DialButton_Click
End Sub

Private Sub OptionCom1_Click()
OptionCom1.Value = True
End Sub

Private Sub OptionCom2_Click()
OptionCom2.Value = True
End Sub

Private Sub OptionCom4_Click()
OptionCom4.Value = True
End Sub

