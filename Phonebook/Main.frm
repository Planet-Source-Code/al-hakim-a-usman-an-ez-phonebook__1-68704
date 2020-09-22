VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form PhoneBook 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Phone book A1.0"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6800
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Entries"
      TabPicture(0)   =   "Main.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Search"
      TabPicture(1)   =   "Main.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "SearchGrid"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Status / About"
      TabPicture(2)   =   "Main.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame4 
         Caption         =   "About"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   19
         Top             =   1800
         Width           =   5775
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Programming Level: Beginner"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   25
            Top             =   1080
            Width           =   3075
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Phone book  A1.0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   24
            Top             =   360
            Width           =   1830
         End
         Begin VB.Label lblMyLink 
            AutoSize        =   -1  'True
            Caption         =   "www40.brinkster.com/hackimusman"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   1440
            TabIndex        =   23
            Top             =   1440
            Width           =   3645
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Visit me "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   22
            Top             =   1440
            Width           =   900
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Through Visual Basic 6.0 Compiler"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   21
            Top             =   840
            Width           =   3570
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Programmed by Al-Hakim A. Usman"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   20
            Top             =   600
            Width           =   3720
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Entry Status"
         Height          =   975
         Left            =   -74880
         TabIndex        =   16
         Top             =   600
         Width           =   5775
         Begin VB.Label lblStatus 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2880
            TabIndex        =   18
            Top             =   360
            Width           =   2340
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Number of Entries :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   17
            Top             =   360
            Width           =   2340
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Find by Name"
         Height          =   735
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   5775
         Begin VB.TextBox txtSearch 
            BackColor       =   &H80000007&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   5535
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5775
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Height          =   495
            Left            =   4440
            TabIndex        =   12
            Top             =   2160
            Width           =   1215
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid EntryGrid 
            Height          =   2055
            Left            =   120
            TabIndex        =   11
            Top             =   1200
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   3625
            _Version        =   393216
            BackColor       =   0
            ForeColor       =   65280
            BackColorFixed  =   -2147483638
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12632256
            ForeColorSel    =   -2147483641
            BackColorBkg    =   -2147483641
            GridColor       =   -2147483641
            GridColorFixed  =   0
            GridColorUnpopulated=   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   495
            Left            =   4440
            TabIndex        =   10
            Top             =   2760
            Width           =   1215
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete"
            Height          =   495
            Left            =   4440
            TabIndex        =   9
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Update"
            Height          =   495
            Left            =   4440
            TabIndex        =   8
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Height          =   495
            Left            =   4440
            TabIndex        =   7
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddNew 
            Caption         =   "Add New"
            Height          =   495
            Left            =   4440
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtName 
            BackColor       =   &H80000007&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   1320
            TabIndex        =   2
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox txtNumber 
            BackColor       =   &H80000007&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   1320
            TabIndex        =   4
            Top             =   720
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "Mobile Number"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Name"
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1215
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid SearchGrid 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   13
         Top             =   1200
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   4471
         _Version        =   393216
         BackColor       =   0
         ForeColor       =   65280
         BackColorFixed  =   -2147483638
         ForeColorFixed  =   -2147483641
         BackColorSel    =   12632256
         ForeColorSel    =   -2147483641
         BackColorBkg    =   -2147483641
         GridColor       =   -2147483641
         GridColorFixed  =   0
         GridColorUnpopulated=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "PhoneBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Public Enum StartWindowState
START_HIDDEN = 0
START_NORMAL = 4
START_MINIMIZED = 2
START_MAXIMIZED = 3
End Enum

Public Function ShellDocument(sDocName As String, _
Optional ByVal Action As String = "Open", _
Optional ByVal Parameters As String = vbNullString, _
Optional ByVal Directory As String = vbNullString, _
Optional ByVal WindowState As StartWindowState) As Boolean
Dim Response
Response = ShellExecute(&O0, Action, sDocName, Parameters, Directory, WindowState)
Select Case Response
    Case Is < 33
        ShellDocument = False
    Case Else
        ShellDocument = True
End Select
End Function

Private Sub cmdAddNew_Click()
Call TxtEnabler(True, True)
Call TxtCleaner("", "")
Call CmdEnabler(True, True, False, False, True)
txtName.SetFocus
End Sub

Private Sub cmdCancel_Click()
Call TxtEnabler(False, False)
Call TxtCleaner("", "")
Call CmdEnabler(True, False, False, False, False)
update_Buffer = ""
End Sub

Private Sub cmdDelete_Click()
Dim delCnf As Integer
delCnf = MsgBox("You are about to delete this entry are you sure?", vbOKCancel, "Request confirm")
If delCnf = 1 Then
    sqlQuery = "Delete*from entries where MobileNo='" & update_Buffer & "'"
    myConn.Execute (sqlQuery)
    myRS.Open "Select*from entries ORDER BY Name ASC", myConn, 3, 3
    Set EntryGrid.DataSource = myRS
    Set SearchGrid.DataSource = myRS
    lblStatus.Caption = myRS.RecordCount
    Set myRS = Nothing
    nam = Trim(txtName.Text)
    
    MsgBox "Phone number of " & nam & " has been deleted.", vbInformation
    
    update_Buffer = ""
    nam = ""
    
    Call TxtEnabler(False, False)
    Call TxtCleaner("", "")
    Call CmdEnabler(True, False, False, False, False)
End If
End Sub

Private Sub cmdEdit_Click()
nam = Trim(txtName.Text)
num = Trim(txtNumber.Text)

sqlQuery = "Update entries set Name='" & nam & "', MobileNo='" & num & "' where MobileNo='" & update_Buffer & "'"
myConn.Execute (sqlQuery)
myRS.Open "Select*from entries ORDER BY Name ASC", myConn, 3, 3
Set EntryGrid.DataSource = myRS
Set SearchGrid.DataSource = myRS
Set myRS = Nothing

MsgBox "Entry has been updated successfuly.", vbInformation, "Progress report"

Call TxtEnabler(False, False)
Call TxtCleaner("", "")
Call CmdEnabler(True, False, False, False, False)

nam = ""
num = ""
update_Buffer = ""
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdSave_Click()
nam = Trim(txtName.Text)
num = Trim(txtNumber.Text)

sqlQuery = "Insert into entries values('" & nam & "','" & num & "')"
myConn.Execute (sqlQuery)
myRS.Open "Select*from entries ORDER BY Name ASC", myConn, 3, 3
Set EntryGrid.DataSource = myRS
Set SearchGrid.DataSource = myRS
lblStatus.Caption = myRS.RecordCount
Set myRS = Nothing

MsgBox "New number has been saved.", vbInformation, "Progress report"

Call TxtEnabler(False, False)
Call TxtCleaner("", "")
Call CmdEnabler(True, False, False, False, False)

nam = ""
num = ""
End Sub

Private Sub EntryGrid_DblClick()
Call TxtEnabler(True, True)
Call TxtCleaner("", "")
Call CmdEnabler(False, False, True, True, True)
txtName.SetFocus


Dim X As Integer
X = EntryGrid.Row
With EntryGrid
    txtName.Text = .TextMatrix(X, 1)
    txtNumber.Text = .TextMatrix(X, 2)
    update_Buffer = .TextMatrix(X, 2)
End With


End Sub

Private Sub Form_Load()
dBPath = App.Path & "\PhoneNumber.mdb"
myConn.Open "Driver={Microsoft Access Driver (*.mdb)}; dbq=" & dBPath

With EntryGrid '4095 - 340
    .ColWidth(0) = 195
    .ColWidth(1) = 2060
    .ColWidth(2) = 1500
End With

With SearchGrid '5775 -340
    .ColWidth(0) = 200
    .ColWidth(1) = 3000
    .ColWidth(2) = 2235
End With


myRS.Open "Select*from entries ORDER BY Name ASC", myConn, 3, 3
Set EntryGrid.DataSource = myRS
Set SearchGrid.DataSource = myRS
lblStatus.Caption = myRS.RecordCount
Set myRS = Nothing

Call TxtEnabler(False, False)
Call CmdEnabler(True, False, False, False, False)
End Sub
Private Function CmdEnabler(cmdVal1, cmdVal2, cmdVal3, cmdVal4, cmdVal5)
    cmdAddNew.Enabled = cmdVal1
    cmdSave.Enabled = cmdVal2
    cmdEdit.Enabled = cmdVal3
    cmdDelete.Enabled = cmdVal4
    cmdCancel.Enabled = cmdVal5
End Function
Private Function TxtEnabler(txtBool1, txtBool2)
    txtName.Enabled = txtBool1
    txtNumber.Enabled = txtBool2
End Function
Private Function TxtCleaner(txtVal1, txtVal2)
    txtName.Text = txtVal1
    txtNumber.Text = txtVal2
End Function

Private Sub Form_Resize()
If Me.WindowState <> 1 Then
    SSTab1.Left = (Me.Width / 2) - (SSTab1.Width / 2) - 50
    SSTab1.Top = (Me.Height / 2) - (SSTab1.Height / 2) - 250
End If
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMyLink.ForeColor = &H800000
End Sub

Private Sub lblMyLink_Click()
ShellDocument "http://www40.brinkster.com/hackimusman"
End Sub

Private Sub lblMyLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMyLink.ForeColor = &HFF0000
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
txtSearch.SetFocus
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
Dim Nam_Buffer As String
Nam_Buffer = "%" + Trim(txtSearch.Text) + "%"
'If Len(Nam_Buffer) = 1 Or Len(Nam_Buffer) > 1 Then
    myRS.Open "Select*from entries where Name LIKE  '" & Nam_Buffer & "'", myConn, adOpenKeyset, adLockPessimistic
    Set SearchGrid.DataSource = myRS
    Set myRS = Nothing
'End If
Nam_Buffer = ""
End Sub
