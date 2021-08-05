VERSION 5.00
Begin VB.Form frmAddressBook1 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   5655
   ClientLeft      =   645
   ClientTop       =   645
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   377
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   Begin VB.PictureBox Picture6 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   22
      Top             =   1.05e5
      Visible         =   0   'False
      Width           =   975
      Begin VB.CommandButton cmdReplace 
         Caption         =   "Replace"
         Height          =   375
         Left            =   0
         TabIndex        =   28
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtNull 
         Height          =   285
         Left            =   0
         TabIndex        =   27
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtExt 
         Height          =   285
         Left            =   0
         TabIndex        =   26
         Text            =   ".ini"
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdCopyList 
         Caption         =   "Copy All"
         Height          =   255
         Left            =   0
         TabIndex        =   25
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtRemExt 
         Height          =   285
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   1560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.FileListBox File1 
         Height          =   300
         Left            =   0
         TabIndex        =   23
         Top             =   1800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Replace with"
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Find What"
         Height          =   495
         Left            =   0
         TabIndex        =   29
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7080
      Top             =   840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   1.05e5
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00800080&
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1275
      ScaleWidth      =   9075
      TabIndex        =   12
      Top             =   0
      Width           =   9135
      Begin VB.CommandButton cmdChange 
         Caption         =   "View and Chage ini file"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   44
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         TabIndex        =   43
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   5400
         TabIndex        =   41
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox cboINI 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         ItemData        =   "frmAddressBook1.frx":0000
         Left            =   1560
         List            =   "frmAddressBook1.frx":0002
         TabIndex        =   39
         Top             =   0
         Width           =   6015
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   21
         Top             =   600
         Width           =   2535
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add Entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H000000FF&
         Caption         =   "Delete Entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800080&
         Caption         =   "Press F2 to add  F3 to Delete an Entry"
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
         Height          =   735
         Left            =   7680
         TabIndex        =   42
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800080&
         Caption         =   "Type your entry:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   4215
      Left            =   5760
      ScaleHeight     =   4155
      ScaleWidth      =   3435
      TabIndex        =   7
      Top             =   1320
      Width           =   3495
      Begin VB.HScrollBar HsbView 
         Height          =   285
         Left            =   0
         TabIndex        =   11
         Top             =   3840
         Width           =   3225
      End
      Begin VB.VScrollBar VsbView 
         Height          =   3855
         Left            =   3240
         TabIndex        =   10
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picViewZ 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   3855
         Left            =   0
         ScaleHeight     =   253
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   213
         TabIndex        =   8
         Top             =   0
         Width           =   3255
         Begin VB.PictureBox picViewX 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1995
            Left            =   -480
            ScaleHeight     =   133
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   259
            TabIndex        =   9
            Top             =   -240
            Width           =   3885
         End
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808000&
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   5475
      TabIndex        =   5
      Top             =   1320
      Width           =   5535
      Begin VB.PictureBox Picture7 
         BackColor       =   &H0000FFFF&
         Height          =   855
         Left            =   120
         ScaleHeight     =   795
         ScaleWidth      =   3555
         TabIndex        =   32
         Top             =   120
         Width           =   3615
         Begin VB.CommandButton Command1 
            BackColor       =   &H0000FF00&
            Caption         =   "Add New Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1440
            TabIndex        =   33
            Text            =   "Type new name"
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label lbllpAppName 
            BackColor       =   &H0000FFFF&
            Caption         =   "Add new name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   1215
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00800000&
         Height          =   3975
         Left            =   3840
         ScaleHeight     =   3915
         ScaleWidth      =   1515
         TabIndex        =   15
         Top             =   120
         Width           =   1575
         Begin VB.OptionButton OptionCom4 
            BackColor       =   &H00800000&
            Caption         =   "COM4"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   360
            TabIndex        =   38
            Top             =   3480
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptionCom1 
            BackColor       =   &H00800000&
            Caption         =   "COM1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   360
            TabIndex        =   37
            Top             =   3000
            Width           =   855
         End
         Begin VB.OptionButton OptionCom2 
            BackColor       =   &H00800000&
            Caption         =   "COM2"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   360
            TabIndex        =   36
            Top             =   3240
            Width           =   855
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00FFFF00&
            Caption         =   "View Home Page"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            Picture         =   "frmAddressBook1.frx":0004
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00FFFF00&
            Caption         =   "Send Email"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            Picture         =   "frmAddressBook1.frx":030E
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00FFFF00&
            Caption         =   "Dial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            Picture         =   "frmAddressBook1.frx":0750
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   2040
            Width           =   1335
         End
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2940
         Left            =   1440
         TabIndex        =   6
         Top             =   1080
         Width           =   2295
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   5280
      ScaleHeight     =   375
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   7000
      Visible         =   0   'False
      Width           =   15
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7800
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.TextBox txtPhone 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmAddressBook1.frx":0B92
      Top             =   1.05e5
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtEdit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4215
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   72
      X2              =   32
      Y1              =   64
      Y2              =   88
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Click one page in this box"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   19
      Top             =   720
      Width           =   2040
   End
End
Attribute VB_Name = "frmAddressBook1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim inCustKey As String
Dim inCustVal As String * 255
Dim File As String
Dim MyiniSetting As Variant, iniSetting As Integer
Dim iniSet As Integer
Dim lRet As Long
Dim sBuff As String * 255
Dim sUser As String
Dim i As Integer
Dim FileNum
Dim FileNames As String


Private Sub cboINI_Change()
On Error Resume Next
Dim Index As Integer
Dim Answer
Index = FindFirstMatch(cboINI, cboINI.Text, -1, True)
cboINI.ListIndex = Index
If Index <> -1 Then Answer = MsgBox(cboINI.Text & " is the " & Index & "th entry in your INI file.", vbYesNo)
If Answer = vbYes Then
cmdAdd.Enabled = True
End If
End Sub

Private Sub cboINI_Click()
txtPhone.Text = cboINI.Text
End Sub

Private Sub cboINI_DropDown()
'cboINI.Clear
While GetPrivateProfileString(Text1.Text, Text1.Text & CStr(i + 1), "", sBuff, 255, App.Path & "\" & List2.Text & ".ini")
 sUser = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
 If sUser <> "Deleted" Then cboINI.AddItem sUser
 i = i + 1
 Wend
 On Error Resume Next
 cboINI.ListIndex = 1

End Sub

Private Sub cboINI_KeyDown(KeyCode As Integer, Shift As Integer)
'ini Add and Delete

Select Case KeyCode

Case vbKeyF2:
 cmdAdd_Click
Case vbKeyF3: 'Deletes from ini
cmdDel_Click
    End Select

End Sub

Private Sub cmdAdd_Click()

If cboINI.Text = ".." Then Exit Sub
While GetPrivateProfileString(Text1.Text, Text1.Text & CStr(i + 1), "", sBuff, 255, App.Path & "\" & List2.Text & ".ini")
 sUser = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
 i = i + 1
 Wend
    inCustKey = RTrim(Text1.Text)  'Trim spaces out of ini file
    inCustVal = cboINI.Text
     lpAppName = Text1.Text
    lpFileName = File
lonStatus = WritePrivateProfileString(lpAppName, inCustKey & CStr(i + 1), inCustVal, lpFileName)

cboINI.Clear 'Clear then fill combo with new ini entries

i = 0 'Set I to 0 to start the read of the ini file at the begining

'Gets all entries
While GetPrivateProfileString(Text1.Text, Text1.Text & CStr(i + 1), "", sBuff, 255, App.Path & "\" & List2.Text & ".ini")
    sUser = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
    If sUser <> "Deleted" Then cboINI.AddItem sUser
    i = i + 1
Wend
Command2_Click
On Error Resume Next
    cboINI.ListIndex = 1 'Sets combotext to 1st entry
End Sub

Private Sub cmdDel_Click()
On Error GoTo errhandler
If cboINI.Text = ".." Then Exit Sub
    i = 0 'Set I to zero so the loop starts at the begining
    
    'Gets all entries
    While GetPrivateProfileString(Text1.Text, Text1.Text & CStr(i + 1), "", sBuff, 255, App.Path & "\" & List2.Text & ".ini")
        sUser = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
        If sUser <> "Deleted" Then cboINI.AddItem sUser
        i = i + 1
        If sUser = cboINI.Text Then GoTo Delini
     Wend
     
 
Delini:
If cboINI.Text = ".." Then Exit Sub
    File = App.Path & "\" & List2.Text & ".ini"
    lpAppName = Text1.Text
    inCustKey = RTrim(Text1.Text)
    inCustVal = cboINI.Text
                               
    lpFileName = File
          On Error Resume Next
       cboINI.RemoveItem (cboINI.ListIndex)
            inCustVal = "Deleted"
                lonStatus = WritePrivateProfileString(lpAppName, inCustKey & CStr(i + 0), inCustVal, lpFileName)
    cboINI.Clear
    i = 0
    
    While GetPrivateProfileString(Text1.Text, Text1.Text & CStr(i + 1), "", sBuff, 255, App.Path & "\" & List2.Text & ".ini")
        sUser = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
        If sUser <> "Deleted" Then cboINI.AddItem sUser
        i = i + 1
    Wend
    cboINI.ListIndex = 1
Exit Sub
errhandler:
End Sub

Private Sub cmdChange_Click()
On Error Resume Next
cmdOK.Visible = True
cmdChange.Visible = False
txtEdit.Visible = True
Picture3.Visible = False
FileNum = FreeFile
Open App.Path & "\" & List2.Text & ".ini" For Input As FileNum
txtEdit.Text = Input(LOF(FileNum), FileNum)
Close FileNum
End Sub

Private Sub cmdOK_Click()
FileNum = FreeFile
txtEdit.Visible = False
Picture3.Visible = True
On Error Resume Next
Open App.Path & "\" & List2.Text & ".ini" For Output As FileNum
Print #FileNum, txtEdit.Text
Close #FileNum
cmdOK.Visible = False
cboINI.Clear
i = 0
While GetPrivateProfileString(Text1.Text, Text1.Text & CStr(i + 1), "", sBuff, 255, App.Path & "\" & List2.Text & ".ini")
    sUser = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
    If sUser <> "Deleted" Then cboINI.AddItem sUser
    i = i + 1
Wend

On Error Resume Next
    cboINI.ListIndex = 1
End Sub

Private Sub Command1_Click()
If Text1.Text = "" Then Exit Sub
cmdOK_Click
Command3_Click
cmdOK.Visible = False
cmdChange.Visible = True
Command2.Enabled = True
cmdAdd.Enabled = True
File = App.Path & "\" & List2.Text & ".ini"
lpAppName = Text1.Text
inCustKey = RTrim(Text1.Text)
inCustVal = cboINI.Text
                           
lpDefault = 0
lpFileName = File

KeyPreview = True

While GetPrivateProfileString(Text1.Text, Text1.Text & CStr(i + 1), "", sBuff, 255, App.Path & "\" & List2.Text & ".ini")
    sUser = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
    If sUser <> "Deleted" Then cboINI.AddItem sUser
    i = i + 1
Wend

If cboINI.ListCount = 0 Then
    inCustVal = ".."
    lonStatus = WritePrivateProfileString(lpAppName, inCustKey & CStr(i + 1), inCustVal, lpFileName)
While GetPrivateProfileString(Text1.Text, Text1.Text & CStr(i + 1), "", sBuff, 255, App.Path & "\" & List2.Text & ".ini")
    sUser = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
    If sUser <> "Deleted" Then cboINI.AddItem sUser
    i = i + 1
Wend
 End If

On Error GoTo errhandler
cboINI.ListIndex = 1
Exit Sub
errhandler:
End Sub

Private Sub Command2_Click()
txtEdit.Text = ""
On Error GoTo errhandler
cmdOK.Visible = True
FileNum = FreeFile

Open App.Path & "\" & List2.Text & ".ini" For Input As FileNum
txtEdit.Text = Input(LOF(FileNum), FileNum)
Close FileNum
Call GetNames
List1.Text = Text1.Text
Exit Sub
errhandler:
Exit Sub
End Sub

Private Sub Command3_Click()
cmdOK.Visible = False

File = App.Path & "\" & List2.Text & ".ini"
lpAppName = Text1.Text
inCustKey = RTrim(Text1.Text)
inCustVal = cboINI.Text
                           
lpDefault = 0
lpFileName = File

KeyPreview = True

While GetPrivateProfileString(Text1.Text, Text1.Text & CStr(i + 1), "", sBuff, 255, App.Path & "\" & List2.Text & ".ini")
    sUser = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
    If sUser <> "Deleted" Then cboINI.AddItem sUser
    i = i + 1
Wend

If cboINI.ListCount = 0 Then
    inCustVal = ".."
    lonStatus = WritePrivateProfileString(lpAppName, inCustKey & CStr(i + 1), inCustVal, lpFileName)
While GetPrivateProfileString(Text1.Text, Text1.Text & CStr(i + 1), "", sBuff, 255, App.Path & "\" & List2.Text & ".ini")
    sUser = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
    If sUser <> "Deleted" Then cboINI.AddItem sUser
    i = i + 1
cboINI.ListIndex = 0
Wend
On Error GoTo errhandler
    cboINI.ListIndex = 1
End If
On Error Resume Next
cboINI.ListIndex = 1
Exit Sub
errhandler:

End Sub

Private Sub Command4_Click()
Call GetNames
End Sub

Private Sub Command5_Click()
 
    On Error Resume Next
    txtEdit.SelStart = 0
    txtEdit.SelLength = Len(txtEdit.Text)
    txtEdit.SelText = ""
    cmdAdd.Enabled = False
    cmdDel.Enabled = False
    picViewX.Picture = LoadPicture()
    cmdOK_Click
End Sub

Private Sub Command6_Click()
DIALER1.Show
End Sub

Private Sub Command7_Click()
Shell ("Start mailto:" & txtPhone.Text), vbHide
End Sub

Private Sub Command8_Click()
Shell ("Start " & txtPhone.Text), vbHide
End Sub

Private Sub File1_Click()
picViewX.Picture = LoadPicture()
Picture4.Visible = True
Picture5.Visible = True
cmdChange.Visible = True
On Error Resume Next
FileNum = FreeFile
Open App.Path & "\" & List2.Text & ".ini" For Input As FileNum
txtEdit.Text = Input(LOF(FileNum), FileNum)
Close FileNum
Command5.Caption = "Clear all entries from " & File1.filename
cboINI.ListIndex = 0
Call GetNames
End Sub

Private Sub Form_Activate()
 cmdCopyList_Click
End Sub

Private Sub Form_Load()
cmdOK.Visible = False
cmdChange.Visible = False
txtEdit.Visible = False
cmdChange.Visible = False
Command2.Enabled = False
cmdAdd.Enabled = False
cmdDel.Enabled = False
Picture4.Visible = False
Picture5.Visible = False
File1.Pattern = "*.ini"
    On Error Resume Next
    FileNames = App.Path & "\Pages.txt"
    Dim mHandle
    mHandle = FreeFile
    Open FileNames For Binary As #mHandle
    Close #mHandle

    maxImageWidth = Screen.Width / Screen.TwipsPerPixelX
    maxImageHeight = Screen.Height / Screen.TwipsPerPixelY
    Dim i, j
    i = maxImageWidth / picViewZ.Width
    j = maxImageHeight / picViewZ.Height
    HsbView.Max = (i - 1) * picViewZ.Width
    VsbView.Max = (j - 1) * picViewZ.Height
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
End Sub


Private Sub List1_Click()
'Call GetKeys

cmdOK.Visible = True
cmdChange.Visible = True
Command2.Enabled = True
cmdAdd.Enabled = True
cmdDel.Enabled = True
Text1.Text = List1.Text
On Error Resume Next

    HsbView.Value = 0
    VsbView.Value = 0
   
    On Error Resume Next
   picViewX.Picture = LoadPicture()
   picViewX.Picture = LoadPicture(App.Path & "\" & List1.Text & ".gif")
    
    Screen.MousePointer = vbNormal
Text1.Text = List1.Text
Command1_Click
End Sub
Private Sub HsbView_Change()
    picViewX.Left = -HsbView.Value
End Sub

Private Sub List2_Click()
picViewX.Picture = LoadPicture()
Picture4.Visible = True
Picture5.Visible = True
cmdChange.Visible = True
On Error Resume Next
FileNum = FreeFile
Open App.Path & "\" & List2.Text & ".ini" For Input As FileNum
txtEdit.Text = Input(LOF(FileNum), FileNum)
Close FileNum
Command5.Caption = "Clear all entries from " & List2.Text
cboINI.ListIndex = 0
Call GetNames
End Sub

Private Sub Timer1_Timer()
Dim MySize
On Error GoTo errhandler
MySize = FileLen(List2.Text & ".ini")
Text2.Text = MySize & " bytes"
If Val(MySize) = 2 Then List1.Clear
Exit Sub
errhandler:
End Sub

Private Sub VsbView_Change()
    picViewX.Top = -VsbView.Value
End Sub

Private Sub GetNames()
List1.Clear
iniFile = App.Path & "\" & List2.Text & ".ini"
File = FreeFile

Open iniFile For Input As File
  Do Until EOF(File)
    Line Input #File, buffer
      If Left(buffer, 1) = "[" Then
        iniName = Mid(buffer, 2, Len(buffer) - 2)
        List1.AddItem (iniName)
      End If
  Loop
Close File
End Sub



Private Sub OptionCom1_Click()
DIALER1.OptionCom1.Value = True
End Sub

Private Sub OptionCom2_Click()
DIALER1.OptionCom2.Value = True
End Sub

Private Sub OptionCom4_Click()
DIALER1.OptionCom4.Value = True
End Sub
Private Sub FillList2()
    Dim mHandle
    Dim tmp As String
    List2.Clear
    mHandle = FreeFile
    
    Open FileNames For Input As #mHandle
    Do While Not EOF(mHandle)
        Line Input #mHandle, tmp
        If Len(Trim(tmp)) > 0 Then
            List2.AddItem tmp
        End If
    Loop
    Close #mHandle
    If List2.ListCount > 0 Then
        List2.ListIndex = 0
    End If
End Sub

Private Function LineThere(inText) As Boolean
    Dim strLine As String
    Dim h, i
    h = List2.ListIndex
    LineThere = False
    For i = 0 To List2.ListCount - 1
        strLine = LTrim(Trim(List2.List(i)))
        If strLine = inText Then
             LineThere = True
             Exit For
        End If
    Next i
    List2.ListIndex = h
End Function

Private Sub SortFileLines()
    Dim strLine As String
    Dim intLineNums As Integer
    Dim arrLines() As String
    Dim mHandle As Variant
    mHandle = FreeFile
      ' We know file is in existence
    Open FileNames For Input As #mHandle
    Do While Not EOF(mHandle)
        Line Input #mHandle, strLine
        strLine = LTrim(Trim(strLine))
        intLineNums = intLineNums + 1
        ReDim Preserve arrLines(1 To intLineNums)
        arrLines(intLineNums) = strLine
    Loop
    Close #mHandle
    SelectionSort arrLines, 1, intLineNums
    mHandle = FreeFile
    Open FileNames For Output As #mHandle
    Dim i
    For i = 1 To intLineNums
         If Len(Trim(arrLines(i))) > 0 Then
             Print #mHandle, arrLines(i)
         End If
    Next i
    Close #mHandle
End Sub

Private Sub SelectionSort(inSortList() As String, ByVal inStart As Integer, ByVal inEnd As Integer)
    Dim i, j, intSelect
    Dim strSelect As String, strTemp As String
    For i = inStart To (inEnd - 1)
        intSelect = i
        strSelect = inSortList(i)
        For j = i + 1 To inEnd
            If StrComp(inSortList(j), strSelect, vbTextCompare) < 0 Then
                strSelect = inSortList(j)
                intSelect = j
            End If
        Next j
        inSortList(intSelect) = inSortList(i)
        inSortList(i) = strSelect
    Next i
End Sub

Private Sub cmdCopyList_Click()
    txtRemExt.Text = ""
    Dim i As Integer
    For i = 0 To File1.ListCount - 1
        txtRemExt.Text = txtRemExt.Text & File1.List(i) & vbCrLf
    Next i
    Call CopyText
    cmdReplace_Click
    Dim FileNum
    FileNum = FreeFile
    Open "Pages.txt" For Output As FileNum
    Print #FileNum, txtRemExt.Text
    Close FileNum
    FillList2
End Sub

Public Sub CopyText()
    txtRemExt.SelLength = Len(txtRemExt.Text)
    Clipboard.Clear
    Clipboard.SetText txtRemExt.SelText
End Sub
Public Sub RemoveString(Entire As String, Word As String, Replace As String)
    Dim i As Integer
    i = 1
    Dim LeftPart
    Do While True

        i = InStr(1, UCase$(Entire), UCase$(Word)) 'This is case sensitive

        If i = 0 Then
            Exit Do
        Else
            LeftPart = Left(Entire, i - 1)
            Entire = LeftPart & Replace & Right(Entire, Len(Entire) - Len(Word) - Len(LeftPart))
        End If
    Loop
txtRemExt.Text = ""
    txtRemExt.Text = Entire
End Sub

Private Sub cmdReplace_Click()
If txtExt.Text = "" Then
Exit Sub
End If
    RemoveString txtRemExt.Text, txtExt.Text, txtNull.Text
On Error Resume Next
End Sub
