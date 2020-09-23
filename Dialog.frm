VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Internet Settings .. "
   ClientHeight    =   3975
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option2 
      Caption         =   "Program Controlled by Local File"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Value           =   -1  'True
      Width           =   3255
   End
   Begin VB.Frame FraLocal 
      Caption         =   "Controlled by file Local "
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   4575
      Begin VB.TextBox txtLocFile 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Text            =   "c:\login.txt"
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label3 
         Caption         =   "Give the Path + Filename and extension:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Program Controlled by Internet"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.Frame FraInternet 
      Caption         =   "Controlled by internet "
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4575
      Begin VB.TextBox txtURL 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "http://users.skynet.be/verburgh.peter/TESTING"
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Text            =   "Test.txt"
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Give the  URL where the file exist on the www"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Give the Filename  and extension:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Help"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   1815
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
MsgBox "Read the ReadMe.txt included in the zip file", vbInformation
End Sub

Private Sub OKButton_Click()
modSettings.FileData = ""
frmMain.Text3 = ""
On Error GoTo FileError:
Dim fso1, f

modSettings.Filename1 = txtFile
modSettings.URL = txtURL
'Check what is used :  LOCAL file  or REMOTE file ..

If Option1.Value = True Then   'if REMOTE file wanted..
modSettings.Filename1 = txtFile
modSettings.URL = txtURL
frmMain.Inet1.Execute modSettings.URL & "/" & modSettings.Filename1, "GET" 'Form resizen
'MsgBox modSettings.URL & "/" & modSettings.Filename1
frmMain.Show
Exit Sub
Else
    modSettings.Filename1 = txtLocFile
    Set fso1 = CreateObject("Scripting.FileSystemObject")
        If (fso1.FileExists(modSettings.Filename1)) Then
        'Okay .... file Exist...
            Set f = fso1.OpenTextFile(modSettings.Filename1, 1)
            modSettings.FileData = f.ReadAll
            f.Close
            Unload Me
            frmMain.Show
            frmMain.Text3 = modSettings.FileData
            frmMain.Start
            Exit Sub
        Else
            MsgBox "Source file " & modSettings.Filename1 & "  NOT found !!  Check & change the Settings !! ", vbCritical
            Exit Sub
        End If
End If
FileError:
MsgBox "Remote File not Found !! ", vbCritical
End Sub

Private Sub Option1_Click()
FraInternet.Enabled = True
FraLocal.Enabled = False
End Sub

Private Sub Option2_Click()
FraInternet.Enabled = False
FraLocal.Enabled = True
End Sub

