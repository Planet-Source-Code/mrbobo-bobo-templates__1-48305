VERSION 5.00
Begin VB.Form frmMRU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MRUTemplate"
   ClientHeight    =   3255
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   2190
      Left            =   1140
      Picture         =   "frmMRU.frx":0000
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMRU.frx":2C10
      Height          =   675
      Left            =   660
      TabIndex        =   0
      Top             =   2340
      Width           =   3675
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuFileSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMRU 
         Caption         =   ""
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMRU 
         Caption         =   ""
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMRUSP1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMRUClear 
         Caption         =   "Clear recent files list"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMRUSP2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmMRU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2003********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive


'Please visit our web site at www.psst.com.au



'Templates
'If you dont know what a Template is then...
'A template is a form, module, user control etc. that
'contains commonly used snippets of code. Templates are
'stored in the VB6 Template directory - usually
'"C:\Program Files\Microsoft Visual Studio\VB98\Template"
'When you add a form, module, user control etc. to your project
'VB presents you with a dialog to select the template
'you wish to use. It fills this dialog with items found
'in the appropriate directory in the Template folder.
'You are free to place items in these folders which you
'wish to use. These items are called "Templates" and
'can dramatically reduce development time. If you
'wish, you could copy these two modules to "Template\Modules"
'and they will then also become available to any future projects.
'If the folder "Modules" does not exist, just create it
'and VB will recognise it.

'This project contains two such Templates...
'Commondialog Open/Save - "ModCmnDlg.bas"
'Recently used files management - "ModMRU.bas"

'Using "ModCmnDlg.bas" means that you do not have to use
'MS Common Dialog Control (comdlg32.ocx) for the simple
'task of opening and saving files.
'
'Using "ModMRU.bas" enables the easy addition of recently
'used files to menus in this form. This can make your
'application easier for the user to use and gives a
'more professional feel to the application

'This form demonstrates the usage of these templates.


Option Explicit
'Variables used for the convenience of the user
'- the Common dialog initial directories
Private OpenDir As String
Private SaveDir As String

Private Sub Form_Load()
    Dim z As Long
    'Allocate menus - see ModMRU declarations section
    ReDim MRU(1 To 5)
    For z = 1 To 5
        Set MRU(z) = mnuMRU(z)
    Next
    Set mruClear = mnuMRUClear
    Set mruSP1 = mnuMRUSP1
    Set mruSP2 = mnuMRUSP2
    'Last used directories used for initial directory of the commondialog
    OpenDir = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "OpenDir")
    SaveDir = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "SaveDir")
    'retrieve recent file list from registry and place in menus
    GetMRUs
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Save last used directories used for initial directory of the commondialog
    SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "OpenDir", OpenDir
    SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "SaveDir", SaveDir
    'save recent file list to registry
    SaveMRUs
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOpen_Click()
    With cmndlg
        .Filter = "All files (*.*)|*.*"
        .Flags = 5
        .Initdir = OpenDir
        .Ownerform = hwnd
        ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
        OpenDir = PathOnly(.FileName)
        Tag = .FileName
        Caption = App.Title & " - " & .Filetitle
        AddMRU .FileName
        'Load file code
    End With
End Sub

Private Sub mnuFileSave_Click()
'save file code
End Sub

Private Sub mnuFileSaveAs_Click()
    With cmndlg
        .Filter = "All files (*.*)|*.*"
        .Flags = 5
        .Initdir = SaveDir
        .Ownerform = hwnd
        'prefill the dialog box with the files' name
        .FileName = IIf(Len(Tag) = 0, "Untitled.txt", FileOnly(Tag))
        .OverwritePrompt = True
        ShowSave
        If Len(.FileName) = 0 Then Exit Sub
        SaveDir = PathOnly(.FileName)
        Caption = App.Title & " - " & .Filetitle
        AddMRU .FileName
        'save file code
    End With
End Sub

Private Sub mnuMRU_Click(Index As Integer)
    If FileExists(MRUColl(Index)) Then
        Tag = MRUColl(Index)
        Caption = App.Title & " - " & FileOnly(Tag)
        'Load file code
    Else
        MsgBox "File not found", vbCritical, "PSST Software"
        RemoveMRU Index
    End If
End Sub

Private Sub mnuMRUClear_Click()
    If MsgBox("Are you sure ypu wish to clear your recent files list ?", vbYesNo, "PSST Software") = vbNo Then Exit Sub
    Set MRUColl = New Collection
    MRUsToMenu
End Sub



