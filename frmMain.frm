VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Core Address Book"
   ClientHeight    =   4215
   ClientLeft      =   1590
   ClientTop       =   1740
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4620
   Begin VB.Frame Frame2 
      Caption         =   "Address"
      Height          =   1365
      Left            =   75
      TabIndex        =   15
      Top             =   900
      Width           =   4470
      Begin VB.TextBox txtCity 
         Appearance      =   0  'Flat
         BackColor       =   &H0074A5A2&
         DataField       =   "City"
         DataSource      =   "datContacts"
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   975
         Width           =   2250
      End
      Begin VB.TextBox txtZip 
         Appearance      =   0  'Flat
         BackColor       =   &H0074A5A2&
         DataField       =   "Zip"
         DataSource      =   "datContacts"
         Height          =   270
         Left            =   3735
         TabIndex        =   7
         Top             =   975
         Width           =   600
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H0074A5A2&
         DataField       =   "Address"
         DataSource      =   "datContacts"
         Height          =   270
         Left            =   105
         TabIndex        =   4
         Top             =   435
         Width           =   4245
      End
      Begin VB.TextBox txtState 
         Appearance      =   0  'Flat
         BackColor       =   &H0074A5A2&
         DataField       =   "State"
         DataSource      =   "datContacts"
         Height          =   270
         Left            =   2505
         TabIndex        =   6
         Top             =   990
         Width           =   1110
      End
      Begin VB.Label Label9 
         Caption         =   "City:"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   795
         Width           =   450
      End
      Begin VB.Label Label7 
         Caption         =   "Address:"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "State:"
         Height          =   180
         Left            =   2490
         TabIndex        =   17
         Top             =   795
         Width           =   450
      End
      Begin VB.Label Label5 
         Caption         =   "Zip:"
         Height          =   180
         Left            =   3720
         TabIndex        =   16
         Top             =   795
         Width           =   285
      End
   End
   Begin VB.TextBox txtLastName 
      Appearance      =   0  'Flat
      BackColor       =   &H0074A5A2&
      DataField       =   "LastName"
      DataSource      =   "datContacts"
      Height          =   270
      Left            =   2115
      TabIndex        =   2
      Top             =   495
      Width           =   1800
   End
   Begin VB.Frame Frame1 
      Caption         =   "Name"
      Height          =   795
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   4470
      Begin VB.TextBox txtMI 
         Appearance      =   0  'Flat
         BackColor       =   &H0074A5A2&
         DataField       =   "MI"
         DataSource      =   "datContacts"
         Height          =   285
         Left            =   3960
         TabIndex        =   3
         Top             =   405
         Width           =   390
      End
      Begin VB.TextBox txtFirstName 
         Appearance      =   0  'Flat
         BackColor       =   &H0074A5A2&
         DataField       =   "FirstName"
         DataSource      =   "datContacts"
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   405
         Width           =   1800
      End
      Begin VB.Label Label4 
         Caption         =   "MI:"
         Height          =   180
         Left            =   3960
         TabIndex        =   14
         Top             =   210
         Width           =   285
      End
      Begin VB.Label Label3 
         Caption         =   "Last:"
         Height          =   180
         Left            =   2040
         TabIndex        =   13
         Top             =   210
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "First:"
         Height          =   180
         Left            =   135
         TabIndex        =   11
         Top             =   210
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Contact"
      Height          =   1350
      Left            =   75
      TabIndex        =   19
      Top             =   2295
      Width           =   4470
      Begin VB.TextBox txtEmailAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H0074A5A2&
         DataField       =   "Email Address"
         DataSource      =   "datContacts"
         Height          =   285
         Left            =   90
         TabIndex        =   10
         Top             =   960
         Width           =   4260
      End
      Begin VB.TextBox txtWorkPhone 
         Appearance      =   0  'Flat
         BackColor       =   &H0074A5A2&
         DataField       =   "Work Phone"
         DataSource      =   "datContacts"
         Height          =   270
         Left            =   2295
         TabIndex        =   9
         Top             =   435
         Width           =   2055
      End
      Begin VB.TextBox txtHomePhone 
         Appearance      =   0  'Flat
         BackColor       =   &H0074A5A2&
         DataField       =   "Home Phone"
         DataSource      =   "datContacts"
         Height          =   270
         Left            =   105
         TabIndex        =   8
         Top             =   435
         Width           =   2070
      End
      Begin VB.Label Label11 
         Caption         =   "Email Address:"
         Height          =   195
         Left            =   105
         TabIndex        =   22
         Top             =   750
         Width           =   1425
      End
      Begin VB.Label Label10 
         Caption         =   "Work Phone:"
         Height          =   180
         Left            =   2295
         TabIndex        =   21
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Home Phone:"
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Height          =   510
      Left            =   60
      TabIndex        =   23
      Top             =   3645
      Width           =   4485
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   270
         Left            =   3480
         TabIndex        =   25
         Top             =   165
         Width           =   930
      End
      Begin VB.Data datContacts 
         Appearance      =   0  'Flat
         BackColor       =   &H00517771&
         Connect         =   "Access"
         DatabaseName    =   "C:\Documents and Settings\Administrator\Desktop\Address Book\Book.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   315
         Left            =   30
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Contacts"
         Top             =   495
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.CommandButton cmdBacku 
         Caption         =   "Back"
         Height          =   270
         Left            =   2565
         TabIndex        =   26
         Top             =   165
         Width           =   930
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   270
         Left            =   1920
         TabIndex        =   27
         Top             =   165
         Width           =   660
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "View All"
         Height          =   270
         Left            =   1185
         TabIndex        =   30
         Top             =   165
         Width           =   750
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   270
         Left            =   585
         TabIndex        =   28
         Top             =   165
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   270
         Left            =   90
         TabIndex        =   29
         Top             =   165
         Width           =   510
      End
   End
   Begin VB.Label Label2 
      Caption         =   "First:"
      Height          =   180
      Left            =   2055
      TabIndex        =   12
      Top             =   270
      Width           =   1695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMinimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuAddContact 
         Caption         =   "Add New Contact"
      End
      Begin VB.Menu mnuRemContact 
         Caption         =   "Remove Contact"
      End
      Begin VB.Menu mnuViewAll 
         Caption         =   "View All Contacts"
      End
      Begin VB.Menu mnuSearchContact 
         Caption         =   "Search For Contact"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    datContacts.Recordset.AddNew
End Sub

Private Sub cmdAll_Click()
    On Error Resume Next
    datContacts.RecordSource = "SELECT * FROM Contacts;"
    datContacts.Refresh
End Sub

Private Sub cmdBacku_Click()
    On Error Resume Next
    If datContacts.Recordset.AbsolutePosition = 0 Then Exit Sub
    datContacts.Recordset.MovePrevious
End Sub

Private Sub cmdDelete_Click()
    Dim Ays As Integer ' Are you Sure?
    Ays = MsgBox("Are you Sure you want to Delete this Person?", vbQuestion + vbYesNo, "Address Book")
    Select Case Ays
        Case 6
            datContacts.Recordset.Delete
            datContacts.Refresh
        Case 7
            Exit Sub
    End Select
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    Dim txtObj As TextBox
    datContacts.Recordset.MoveNext
    If txtFirstName = ("") Then
        datContacts.Recordset.MovePrevious
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdSearch_Click()
    Dim NewSearch As String
    NewSearch = InputBox("Last Name of Contact:", "Search")
    If NewSearch = ("") Then Exit Sub
    datContacts.RecordSource = ("SELECT * From Contacts Where LastName='" & NewSearch & "' ")
    datContacts.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Ays As Integer
    Ays = MsgBox("Are You Sure you Want to Exit?", vbQuestion + vbYesNo, "Address Book")
    Select Case Ays
        Case 6
            End
        Case 7
            Cancel = 1
    End Select
End Sub

Private Sub Label7_Click()
    MsgBox datContacts.Recordset.AbsolutePosition
End Sub

Private Sub mnuAbout_Click()
  Call MsgBox("This program was written as a learning project, and was never meant for anything other than that " & _
                "... just don't copy it and call it your own.. And btw.. I Love you Jennifer!", vbInformation + vbOKOnly, "About Address Book")
End Sub

Private Sub mnuAddContact_Click()
    datContacts.Recordset.AddNew
End Sub

Private Sub mnuClose_Click()
    Dim Ays As Integer
    Ays = MsgBox("Are You Sure you Want to Exit?", vbQuestion + vbYesNo, "Address Book")
    Select Case Ays
        Case 6
            End
        Case 7
            Exit Sub
    End Select
End Sub

Private Sub mnuMinimize_Click()
    Me.WindowState = 1
End Sub

Private Sub mnuRemContact_Click()
    Dim Ays As Integer ' Are you Sure?
    Ays = MsgBox("Are you Sure you want to Delete this Person?", vbQuestion + vbYesNo, "Address Book")
    Select Case Ays
        Case 6
            datContacts.Recordset.Delete
            datContacts.Refresh
        Case 7
            Exit Sub
    End Select
End Sub

Private Sub mnuSearchContact_Click()
    Dim NewSearch As String
    NewSearch = InputBox("Last Name of Contact:", "Search")
    If NewSearch = ("") Then Exit Sub
    datContacts.RecordSource = ("SELECT * From Contacts Where LastName='" & NewSearch & "' ")
    datContacts.Refresh
End Sub

Private Sub mnuViewAll_Click()
    On Error Resume Next
    datContacts.RecordSource = "SELECT * FROM Contacts;"
    datContacts.Refresh
End Sub
