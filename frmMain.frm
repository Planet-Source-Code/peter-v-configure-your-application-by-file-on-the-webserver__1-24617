VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   Caption         =   "Application Control  by internet"
   ClientHeight    =   4740
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   1605
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Text            =   "frmMain.frx":0000
      Top             =   3000
      Width           =   4455
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3720
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   0
      TabIndex        =   9
      Top             =   240
      Width           =   855
   End
   Begin VB.Frame FraTXT 
      Caption         =   "TextBox's"
      Height          =   1215
      Left            =   960
      TabIndex        =   6
      ToolTipText     =   "FraTXT"
      Top             =   120
      Width           =   1215
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "Text2"
         ToolTipText     =   "Text2"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "Text1"
         ToolTipText     =   "Text1"
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame FraCheck 
      Caption         =   "Checkbox's"
      Height          =   1215
      Left            =   2280
      TabIndex        =   3
      ToolTipText     =   "FraCheck"
      Top             =   120
      Width           =   1095
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame FraButton 
      Caption         =   "Buttons"
      Height          =   1215
      Left            =   3480
      TabIndex        =   0
      ToolTipText     =   "FraButton"
      Top             =   120
      Width           =   1215
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Data from the local or remote file :"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Message from Local file or internet file on your site"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label lblMESSAGE 
      Caption         =   "Message is:"
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "lblMESSAGE"
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuSettings 
         Caption         =   "Global Settings"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Start()
Dim count, Item As Integer
On Error GoTo DataERROR

'Listbox
List1.Enabled = modRAFile.GetValue(modSettings.FileData, "LIST1", ".ENABLED")
List1.Visible = modRAFile.GetValue(modSettings.FileData, "LIST1", ".VISIBLE")
'-------- check item count for the listbox..
count = modRAFile.GetValue(modSettings.FileData, "LIST1", "ITEMS")
For n = 1 To count
    List1.AddItem modRAFile.GetValue(modSettings.FileData, "LIST1", "ITEM" & n)
Next n

'Fra Text Settings...
FraTXT.Enabled = modRAFile.GetValue(modSettings.FileData, "FraTXT", ".ENABLED")
FraTXT.Visible = modRAFile.GetValue(modSettings.FileData, "FraTXT", ".VISIBLE")
Text1.Text = modRAFile.GetValue(modSettings.FileData, "FraTXT", "TEXT1")
Text2.Text = modRAFile.GetValue(modSettings.FileData, "FraTXT", "TEXT2")
Text1.Enabled = modRAFile.GetValue(modSettings.FileData, "FraTXT", "TEXT1.ENABLED")
Text2.Enabled = modRAFile.GetValue(modSettings.FileData, "FraTXT", "TEXT2.ENABLED")
'Fra Button Settings..
FraButton.Enabled = modRAFile.GetValue(modSettings.FileData, "FraButton", ".ENABLED")
FraButton.Visible = modRAFile.GetValue(modSettings.FileData, "FraButton", ".VISIBLE")
Command1.Visible = modRAFile.GetValue(modSettings.FileData, "FraButton", "COMMAND1.VISIBLE")
Command2.Visible = modRAFile.GetValue(modSettings.FileData, "FraButton", "COMMAND2.VISIBLE")
'lblMESSAGE
lblMESSAGE.Caption = modRAFile.GetValue(modSettings.FileData, "lblMESSAGE", ".CAPTION")
'FraCheck
FraCheck.Visible = modRAFile.GetValue(modSettings.FileData, "FraCheck", ".VISIBLE")
FraCheck.Enabled = modRAFile.GetValue(modSettings.FileData, "FraCheck", ".ENABLED")
Check1.Value = modRAFile.GetValue(modSettings.FileData, "FraCheck", "CHECK1.VALUE")
Check1.Enabled = modRAFile.GetValue(modSettings.FileData, "FraCheck", "CHECK1.ENABLED")
Check2.Value = modRAFile.GetValue(modSettings.FileData, "FraCheck", "CHECK2.VALUE")
Check2.Enabled = modRAFile.GetValue(modSettings.FileData, "FraCheck", "CHECK2.ENABLED")
Exit Sub
DataERROR:
    MsgBox "A Subject or Item  is not Found  !!", vbCritical
End Sub

Private Sub Form_Load()
'Text3 = " blallalalalall " & vbCrLf & "[BOX]" & vbCrLf & "data1 =  XtestingX;" & vbCrLf & "data12 = OkayEnd;" & vbCrLf
Text3 = modSettings.FileData

End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
'On Error GoTo Error1:
Dim FileEnd As Boolean
Dim DataX

Select Case State
   Case icHostResolvingHost
   Case icConnecting
   Debug.Print "Connecting to Host"
   Case icConnected
   Debug.Print "Connected To Host"
   Case icRequesting
   Debug.Print "Requesting data"
   Case icRequestSent
   Debug.Print "Request Send"
   Case icReceivingResponse
   Debug.Print "Control receives response from host"
   Case icResponseReceived
   
   Debug.Print "Successfully receive response from host "
     Do
     DoEvents
     Loop While Inet1.StillExecuting = True
     
     Do
     DataX = Inet1.GetChunk(20000)
     modSettings.FileData = modSettings.FileData & DataX
     DoEvents
     Loop While Len(DataX) > 0
     Text3 = modSettings.FileData
     Dialog.Visible = False
    frmMain.Show
    frmMain.Start
     'MsgBox modSettings.FileData
     If InStr(1, Text3, "404 Not Found") > 0 Then MsgBox "File not Found on this Server!"
   Case icResponseCompleted
        Debug.Print "Data All Received"
        
   Case icError
     MsgBox "An Error has detected ! Host not found"
   End Select
   Exit Sub
End Sub

Private Sub mnuExit_Click()
Unload Me
Unload Dialog
End Sub

Private Sub mnuSettings_Click()
Dialog.Show


End Sub
