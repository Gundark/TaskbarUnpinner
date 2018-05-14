' VBScript to unpin unwanted or default shortcuts from the Windows taskbar.

' Based on Stefan Hofbauer's code from https://pinto10blog.wordpress.com/2016/09/10/pinto10/


Set objShell = CreateObject("WScript.Shell")
userProfilePath = objShell.ExpandEnvironmentStrings("%UserProfile%")
programData = objShell.ExpandEnvironmentStrings("%ProgramData%")


Set sho = CreateObject("Shell.Application")

Set folder = sho.Namespace(userProfilePath + "\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar")
processNamespace(folder)

' For some unknown reason, the namespace above doesn't work for File Explorer, but this namespace does.

Set folder = sho.Namespace(programData + "\Microsoft\Windows\Start Menu Places")
processNamespace(folder)


' Functions ----------------------------------------------------------------------------------

Function processNamespace(folder)
    For Each shellFolderItem In folder.Items
        If unwanted(shellFolderItem.Name) Then
            ' For debugging, uncomment the call to 'showVerbs' below:
            ' call showVerbs(shellFolderItem)
            call unpin(shellFolderItem)
        End If
    Next
End Function


Function unwanted(itemName)
    unwanted = false

    If stringContains(itemName, "Word 2013") OR _
       stringContains(itemName, "Excel 2013") OR _
       stringContains(itemName, "Outlook 2013") OR _
       stringContains(itemName, "Internet Explorer") OR _
       stringContains(itemName, "File Explorer") Then

        unwanted = true
    End If
End Function


Function stringContains(sourceStr, checkStr)
    stringContains = InStr(1, sourceStr, checkStr, vbTextCompare) > 0
End Function


Function unpin(shellFolderItem)
    unpinPresent = canApplicationBeUnpinned(shellFolderItem)

    If (unpinPresent) Then
        shellFolderItem.InvokeVerb("taskbarunpin")
    End If
End Function


Function canApplicationBeUnpinned(shellFolderItem)
    Set verbList = shellFolderItem.Verbs()

    canApplicationBeUnpinned = false
    For Each verb In verbList
        If verb.Name = "Unpin from tas&kbar" Then
            canApplicationBeUnpinned = true
            Exit For
        End If
    Next
End Function


Function showVerbs(shellFolderItem)
    doubleNewline = vbCrLf & vbCrLf
    Set verbList = shellFolderItem.Verbs()

    displayString = shellFolderItem.Name & vbCrLf _
                    & vbCrLf & "Available Verbs" & vbCrLf & "------------------------" & vbCrLf & shellFolderItem.Path

    For Each verb In verbList
        displayString = displayString & vbCrLf & "  - " & verb.Name
    Next

    unpinPresent = canApplicationBeUnpinned(shellFolderItem)
    displayString = displayString & doubleNewline & "Unpin present = " & unpinPresent

    MsgBox(displayString)
End Function
