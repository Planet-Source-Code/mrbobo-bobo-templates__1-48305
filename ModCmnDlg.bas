Attribute VB_Name = "ModCmnDlg"
'******************************************************************
'***************Copyright PSST 2003********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive

'Please visit our web site at www.psst.com.au

Option Explicit
'Commondialog API - more efficient than using MS Common Dialog Control (comdlg32.ocx)
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXPLORER = &H80000
'UDT that makes calling the commondialog easier
Public Type CMDialog
    Ownerform As Long
    Filter As String
    Filetitle As String
    FilterIndex As Long
    FileName As String
    DefaultExtension As String
    OverwritePrompt As Boolean
    AllowMultiSelect As Boolean
    Initdir As String
    Dialogtitle As String
    Flags As Long
End Type
Public cmndlg As CMDialog
'****************COMMONDIALOG CODE*********************
Public Sub ShowOpen()
    Dim OFName As OPENFILENAME
    Dim temp As String
    With cmndlg
        OFName.lStructSize = Len(OFName)
        OFName.hwndOwner = .Ownerform
        OFName.hInstance = App.hInstance
        OFName.lpstrFilter = Replace(.Filter, "|", Chr(0))
        OFName.lpstrFile = Space$(254)
        OFName.nMaxFile = 255
        OFName.lpstrFileTitle = Space$(254)
        OFName.nMaxFileTitle = 255
        OFName.lpstrInitialDir = .Initdir
        OFName.lpstrTitle = .Dialogtitle
        OFName.nFilterIndex = .FilterIndex
        OFName.Flags = .Flags Or OFN_EXPLORER Or IIf(.AllowMultiSelect, OFN_ALLOWMULTISELECT, 0)
        If GetOpenFileName(OFName) Then
            .FilterIndex = OFName.nFilterIndex
            If .AllowMultiSelect Then
                temp = Replace(Trim$(OFName.lpstrFile), Chr(0), ";")
                If Right(temp, 2) = ";;" Then temp = Left(temp, Len(temp) - 2)
                .FileName = temp
            Else
                .FileName = StripTerminator(Trim$(OFName.lpstrFile))
                .Filetitle = StripTerminator(Trim$(OFName.lpstrFileTitle))
            End If
        Else
            .FileName = ""
        End If
    End With
End Sub
Public Sub ShowSave()
    Dim OFName As OPENFILENAME
    With cmndlg
        OFName.lStructSize = Len(OFName)
        OFName.hwndOwner = .Ownerform
        OFName.hInstance = App.hInstance
        OFName.lpstrFilter = Replace(.Filter, "|", Chr(0))
        OFName.nMaxFile = 255
        OFName.lpstrFileTitle = Space$(254)
        OFName.nMaxFileTitle = 255
        OFName.lpstrInitialDir = .Initdir
        OFName.lpstrTitle = .Dialogtitle
        OFName.nFilterIndex = .FilterIndex
        OFName.lpstrDefExt = .DefaultExtension
        OFName.lpstrFile = .FileName & Space$(254 - Len(.FileName))
        OFName.Flags = .Flags Or IIf(.OverwritePrompt, OFN_OVERWRITEPROMPT, 0)
        If GetSaveFileName(OFName) Then
            .FileName = StripTerminator(Trim$(OFName.lpstrFile))
            .Filetitle = StripTerminator(Trim$(OFName.lpstrFileTitle))
            .FilterIndex = OFName.nFilterIndex
        Else
            .FileName = ""
        End If
    End With
End Sub

'****************STRING FUNCTIONS*********************
Public Function StripTerminator(ByVal strString As String) As String
    'Removes chr(0)'s from the end of a string
    'API tends to do this
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function


