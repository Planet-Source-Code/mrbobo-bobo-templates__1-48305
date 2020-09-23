Attribute VB_Name = "ModMRU"
'******************************************************************
'***************Copyright PSST 2003********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive

'Please visit our web site at www.psst.com.au

Option Explicit
'Menus for use by this module in the calling form
Public mruClear As Menu
Public mruSP1 As Menu
Public mruSP2 As Menu
Public MRU() As Menu
'Collection to hold recent files
Public MRUColl As New Collection
'****************RECENT FILE CODE*********************

Public Sub GetMRUs()
    Dim MySettings() As String
    Dim z As Long
    Dim cnt As Long
    'Placing a dummy entry avoids any errors should there be no recent files
    SaveSetting "PSST SOFTWARE\" + App.Title, "Recent", "A", "Empty"
    'Retrieve all the recent files
    MySettings = GetAllSettings("PSST SOFTWARE\" + App.Title, "Recent")
    For z = LBound(MySettings, 1) To UBound(MySettings, 1)
        If MySettings(z, 1) <> "Empty" Then
            If FileExists(MySettings(z, 1)) Then
                'If the file still exists add it to our collection
                MRUColl.add MySettings(z, 1)
                cnt = cnt + 1
                If cnt > UBound(MRU) Then Exit For
            End If
        End If
    Next
    'Load up the menus
    MRUsToMenu
End Sub
Public Sub SaveMRUs()
    Dim z As Long
    'Placing a dummy entry avoids any errors when we delete an entire key
    SaveSetting "PSST SOFTWARE\" + App.Title, "Recent", "A", "Empty"
    'Delete all old entries
    DeleteSetting "PSST SOFTWARE\" + App.Title, "Recent"
    'Save all the current recent files
    If MRUColl.count > 0 Then
        For z = 1 To MRUColl.count
            SaveSetting "PSST SOFTWARE\" + App.Title, "Recent", CStr(z), MRUColl(z)
        Next
    End If
End Sub


Public Sub MRUsToMenu()
    Dim z As Long
    'Hide all the MRU menus
    For z = 1 To UBound(MRU)
        MRU(z).Visible = False
        MRU(z).Caption = ""
    Next
    If MRUColl.count > 0 Then
        For z = 1 To MRUColl.count
            If z > UBound(MRU) Then Exit For
            'Show the menu with the filename
            MRU(z).Visible = True
            MRU(z).Caption = FileOnly(MRUColl(z))
        Next
    End If
    'Show auxilary menus if required
    If Not mruClear Is Nothing Then mruClear.Visible = CBool(MRUColl.count)
    If Not mruSP1 Is Nothing Then mruSP1.Visible = CBool(MRUColl.count)
    If Not mruSP2 Is Nothing Then mruSP2.Visible = CBool(MRUColl.count)
End Sub

Public Sub AddMRU(mPath As String)
    Dim z As Long
    If MRUColl.count > 0 Then
        'If this file is already in the list then
        'remove it because we want it to appear
        'as the first menu
        For z = 1 To MRUColl.count
            If MRUColl(z) = mPath Then
                MRUColl.Remove z
                Exit For
            End If
        Next
        If MRUColl.count > 0 Then
            'Add it as the first item
            MRUColl.add mPath, , 1
        Else
            'just add it
            MRUColl.add mPath
        End If
    Else
        'add the file to the collection
        MRUColl.add mPath
    End If
    'If somehow we got more than the menus remove excess items
    'this should not happen
    If MRUColl.count > UBound(MRU) Then
        For z = MRUColl.count To UBound(MRU) + 1 Step -1
            MRUColl.Remove z
        Next
    End If
    'Update the menus
    MRUsToMenu
End Sub

Public Sub RemoveMRU(Index As Integer)
    If Index < MRUColl.count Then
        'Removing the file from the collection
        MRUColl.Remove Index
        'Update the menus
        MRUsToMenu
    End If
End Sub
'****************FILEEXISTS*********************
'Normally I would not use the FileSystemObject as it can be very
'slow for large tasks. But for this task, which is very small,
'it has the overwhelming advantage of coping with network paths
'without any extra code
Public Function FileExists(mPath) As Boolean
    Dim fs As Object
    Dim mAttr As VbFileAttribute
    On Error GoTo woops
    Set fs = CreateObject("Scripting.FileSystemObject")
    mAttr = GetAttr(mPath)
    If mAttr And vbVolume Then
        FileExists = fs.DriveExists(mPath)
    ElseIf mAttr And vbDirectory Then
        FileExists = fs.FolderExists(mPath)
    Else
        FileExists = fs.FileExists(mPath)
    End If
    Set fs = Nothing
    Exit Function
woops:
    Set fs = Nothing
    FileExists = False
End Function

'****************STRING FUNCTIONS*********************
Public Function FileOnly(ByVal filepath As String) As String
    FileOnly = Mid$(filepath, InStrRev(filepath, "\") + 1)
End Function
Public Function PathOnly(ByVal filepath As String) As String
    Dim temp As String
    temp = Mid$(filepath, 1, InStrRev(filepath, "\"))
    If Right(temp, 1) = "\" And Len(temp) > 3 Then temp = Left(temp, Len(temp) - 1)
    PathOnly = temp
End Function


