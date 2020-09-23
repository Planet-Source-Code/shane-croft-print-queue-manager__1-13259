VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Print Queue Manager"
   ClientHeight    =   6165
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "0"
      Top             =   3480
      Width           =   495
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   25
      Top             =   5910
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   13732
            Text            =   "Status:"
            TextSave        =   "Status:"
            Object.ToolTipText     =   "Current Status"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4920
      Top             =   3360
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   23
      Text            =   "10"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Automaticly Refresh Every"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get Print Queues"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2880
      TabIndex        =   19
      Top             =   360
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   2760
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh Selected Print Queue"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Print Queue Jobs"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   7815
      Begin MSComctlLib.ListView ListView1 
         Height          =   1695
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   2990
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Print Queue Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   2880
      TabIndex        =   4
      Top             =   960
      Width           =   5055
      Begin VB.TextBox txtPrintStatus 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox txtModel 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox txtPrintPath 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox txtLocation 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label6 
         Caption         =   "Print Status"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Printer Path"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblText 
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Sec."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   24
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Total Print Jobs:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   21
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "You may also Right Click on a selected print queue job to use the tools on that print job."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   20
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Total Print Queues:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Print Queues"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Computer Name or IP Address"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Menu menufile 
      Caption         =   "File"
      Begin VB.Menu menuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu menuTools 
      Caption         =   "Tools"
      Begin VB.Menu MenuPause 
         Caption         =   "Pause Print Queue"
         Shortcut        =   {F1}
      End
      Begin VB.Menu MenuResume 
         Caption         =   "Resume Print Queue"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MenuPriority 
         Caption         =   "Change Print Job Priority"
         Shortcut        =   {F3}
      End
      Begin VB.Menu MenuRemove 
         Caption         =   "Remove Print Job"
         Shortcut        =   {F4}
      End
      Begin VB.Menu MenuPurge 
         Caption         =   "Purge All Print Jobs"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ContainerPrint As IADsContainer
Dim pq As IADsPrintQueue
Dim pQOps As IADsPrintQueueOperations
Dim pj As IADsPrintJob
Dim pjOps As IADsPrintJobOperations
Public Sub List_Add(List As ComboBox, txt As String)
On Error Resume Next
    Combo1.AddItem txt
End Sub
Public Sub List_Load(thelist As ComboBox, FileName As String)
    'Loads a file to a ComboBox
    On Error Resume Next
    Dim TheContents As String
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Input As fFile
    Do
        Line Input #fFile, TheContents$
        Call List_Add(Combo1, TheContents$)
    Loop Until EOF(fFile)
    Close fFile
End Sub

Public Sub List_Save(thelist As ComboBox, FileName As String)
    'Save a ComboBox as FileName
    On Error Resume Next
    Dim Save As Long
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Append As fFile
        Print #fFile, Combo1.Text
    Close fFile
End Sub
Private Function GetCurrentPrintQueue() As IADsPrintQueue
 If (List1.Text = "") Then
    Set GetCurrentPrintQueue = Nothing
    Exit Function
 End If
    
 Set GetCurrentPrintQueue = ContainerPrint.GetObject("printQueue", List1.Text)

End Function

Private Sub Check1_Click()
If Check1.Value = 1 Then
Timer2.Enabled = True
Else
Timer2.Enabled = False
End If
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command2_Click
 DoEvents
 End If
End Sub

Private Sub Command1_Click()
Call List1_Click
End Sub

Private Sub Command2_Click()
On Error Resume Next
List1.Clear
ListView1.ListItems.Clear

Call List_Save(Combo1, App.Path & "\History.ini")
DoEvents
StatusBar1.Panels(1).Text = "Status: Getting Print Queues...This may take a few min."
DoEvents
Dim ComputerName As String
Dim PrintQueue As IADsPrintQueue

DoEvents
ComputerName = Combo1.Text
Set ContainerPrint = GetObject("WinNT://" & ComputerName)
DoEvents
ContainerPrint.Filter = Array("PrintQueue")
For Each PrintQueue In ContainerPrint
     List1.AddItem PrintQueue.Name
Next
DoEvents
StatusBar1.Panels(1).Text = "Status:"
DoEvents
End Sub

Private Sub Form_Load()
On Error Resume Next
Call List_Load(Combo1, App.Path & "\History.ini")
DoEvents
'---- Print Jobs
ListView1.ColumnHeaders.Add , , "Priority"
ListView1.ColumnHeaders.Add , , "Description"
ListView1.ColumnHeaders.Add , , "User"
ListView1.ColumnHeaders.Add , , "Pages"
ListView1.ColumnHeaders.Add , , "Printed"
ListView1.ColumnHeaders.Add , , "Status"
ListView1.ColumnHeaders.Add , , "Size"
ListView1.ColumnHeaders.Add , , "Submitted"
ListView1.ColumnHeaders.Add , , "Name/ID"
    
    
'--- Set to Report View ---------------
ListView1.View = 3
End Sub

Private Sub List1_Click()

Set pq = GetCurrentPrintQueue()
StatusBar1.Panels(1).Text = "Status: Getting Print Queue Information..."
DoEvents
'-------Print Queue --------------------------
On Error Resume Next
txtDescription = pq.Description
txtPrintPath = pq.PrinterPath
txtLocation = pq.Location
txtModel = pq.Model




'----- Print Queue Operations ---------
Set pQOps = pq
txtPrintStatus = GetPrintStatus(pQOps.status)

'---- Print Jobs and Print Job Operations ---------------------
ListView1.ListItems.Clear ' Clear the user interface
For Each pj In pQOps.PrintJobs
   Set pjOps = pj
   Set newLine = ListView1.ListItems.Add(, , pj.Priority)
   newLine.SubItems(1) = pj.Description
   newLine.SubItems(2) = pj.User
   newLine.SubItems(3) = pj.TotalPages
   newLine.SubItems(4) = CStr(pjOps.PagesPrinted)
   newLine.SubItems(5) = GetJobStatus(pjOps.status)
   newLine.SubItems(6) = pj.Size \ 1024 & "KB"
   newLine.SubItems(7) = pj.TimeSubmitted
   newLine.SubItems(8) = pj.Name
Next
StatusBar1.Panels(1).Text = "Status:"
DoEvents
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Make sure it's the right button.
    If Button And vbRightButton _
        Then PopupMenu menuTools
End Sub

Private Sub menuExit_Click()
End
End Sub

Private Sub MenuPause_Click()
On Error Resume Next
StatusBar1.Panels(1).Text = "Status: Puaseing Print Queue..."
DoEvents
Set pq = GetCurrentPrintQueue()
DoEvents
Set pQOps = pq
DoEvents

pQOps.Pause
DoEvents
Call List1_Click
DoEvents
StatusBar1.Panels(1).Text = "Status:"
DoEvents
End Sub

Private Sub MenuPriority_Click()

Dim sNewName As String
sNewName = Trim(InputBox("Please enter a new Priority for this Print Job:"))
If sNewName <> "" Then
Exit Sub
Else
StatusBar1.Panels(1).Text = "Status: Changeing Priority..."
DoEvents
Set pq = GetCurrentPrintQueue()
DoEvents
Set pQOps = pq
DoEvents
For Each PrintJob In PrintQueue.PrintJobs
          If PrintJob.Name = ListView1.SelectedItem.SubItems(8) Then
          pj.Priority = sNewName
          pj.SetInfo
            End If
            Next
End If
StatusBar1.Panels(1).Text = "Status:"
DoEvents
End Sub

Private Sub MenuPurge_Click()
On Error Resume Next
StatusBar1.Panels(1).Text = "Status: Purgeing All Print Jobs..."
DoEvents
Set pq = GetCurrentPrintQueue()
DoEvents
Set pQOps = pq
DoEvents

pQOps.Purge
DoEvents
Call List1_Click
DoEvents
StatusBar1.Panels(1).Text = "Status:"
DoEvents
End Sub

Private Sub MenuRemove_Click()
On Error Resume Next
StatusBar1.Panels(1).Text = "Status: Removeing Selected Print Job..."
DoEvents
Set pq = GetCurrentPrintQueue()
DoEvents
Set pQOps = pq
DoEvents

pQOps.PrintJobs.Remove ListView1.SelectedItem.SubItems(8)
DoEvents
Call List1_Click
DoEvents
StatusBar1.Panels(1).Text = "Status:"
DoEvents
End Sub

Private Sub MenuResume_Click()
On Error Resume Next
StatusBar1.Panels(1).Text = "Status: Resumeing Print Queue..."
DoEvents
Set pq = GetCurrentPrintQueue()
DoEvents
Set pQOps = pq
DoEvents

pQOps.Resume
DoEvents
Call List1_Click
DoEvents
StatusBar1.Panels(1).Text = "Status:"
DoEvents
End Sub

Private Sub Timer1_Timer()
Label7.Caption = "Total Print Queues: " & List1.ListCount
Label9.Caption = "Total Print Jobs: " & ListView1.ListItems.Count

If Combo1.Text = "" Then
Command2.Enabled = False
Else
Command2.Enabled = True
End If

If List1.Text = "" Then
Check1.Enabled = False
Text1.Enabled = False
Else
Check1.Enabled = True
Text1.Enabled = True
End If

If List1.ListCount = 0 Then
Command1.Enabled = False

Else
Command1.Enabled = True

End If
End Sub

Private Sub Timer2_Timer()
Dim x As Long

If Check1.Enabled = False Then
Timer2.Enabled = False
Exit Sub
End If

x = Text2.Text
x = x + 1
Text2.Text = x

If x > Text1.Text Then
Call List1_Click
Text2.Text = "0"
End If
End Sub
