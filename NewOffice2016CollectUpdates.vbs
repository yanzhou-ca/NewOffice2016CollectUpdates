' Office 2016 Patch Extractor
' Extracts and renames all installed Microsoft Office 2016 patches based on modified date

Option Explicit

' Declare variables
Dim oMsi, oFso, oWShell
Dim Patches, SumInfo
Dim patch, msp, qView, record
Dim sTargetFolder, sMessage
Dim patchCount, successCount, failedCount, totalPatches
Dim errorLog, logFile
Dim fileCounter
Dim patchDates, localPackage

' Constants
Const OFFICEID = "000-0000000FF1CE}"
Const PRODUCTCODE_EMPTY = ""
Const MACHINESID = ""
Const MSIINSTALLCONTEXT_MACHINE = 4
Const MSIPATCHSTATE_APPLIED = 1
Const MSIOPENDATABASEMODE_PATCHFILE = 32
Const PID_TEMPLATES = 7 'PatchTargets

' Create objects
Set oMsi = CreateObject("WindowsInstaller.Installer")
Set oFso = CreateObject("Scripting.FileSystemObject")
Set oWShell = CreateObject("WScript.Shell")

' Create the target folder (use %TEMP% as fallback if C:\Office2016Updates fails)
sTargetFolder = oWShell.ExpandEnvironmentStrings("%SystemDrive%") & "\Office2016Updates\"
If Not oFso.FolderExists(sTargetFolder) Then
    On Error Resume Next
    oFso.CreateFolder sTargetFolder
    If Err.Number <> 0 Then
        sTargetFolder = oWShell.ExpandEnvironmentStrings("%TEMP%") & "\Office2016Updates\"
        If Not oFso.FolderExists(sTargetFolder) Then oFso.CreateFolder sTargetFolder
        Err.Clear
    End If
    On Error Goto 0
End If

' Error log file path (only created if needed)
errorLog = sTargetFolder & "error_log.txt"
Set logFile = Nothing

' Show initial message
sMessage = "Patches are being copied to the " & sTargetFolder & " folder." & vbCrLf & _
           "A Windows Explorer window will open after the script has completed." & vbCrLf & _
           "Run as Administrator if copying fails."
oWShell.Popup sMessage, 20, "Office 2016 Updates Extractor"

' Initialize counters
totalPatches = 0
patchCount = 0
successCount = 0
failedCount = 0
fileCounter = -1 ' Start at -1 so first increment is 0

' Dictionary to store patches and their modified dates
Set patchDates = CreateObject("Scripting.Dictionary")

' Get all applied patches
On Error Resume Next
Set Patches = oMsi.PatchesEx(PRODUCTCODE_EMPTY, MACHINESID, MSIINSTALLCONTEXT_MACHINE, MSIPATCHSTATE_APPLIED)
If Err.Number <> 0 Then
    Set logFile = oFso.CreateTextFile(errorLog, True)
    logFile.WriteLine "Office 2016 Patch Extractor Log - " & Now()
    logFile.WriteLine "------------------------------------"
    logFile.WriteLine "Error getting patches: " & Err.Description
    logFile.Close
    MsgBox "Error getting patches: " & Err.Description, 16, "Error"
    WScript.Quit
End If
On Error Goto 0

' Enumerate all patches and store modified dates
For Each patch In Patches
    On Error Resume Next
    localPackage = patch.PatchProperty("LocalPackage")
    totalPatches = totalPatches + 1
    If Err.Number <> 0 Then
        If logFile Is Nothing Then
            Set logFile = oFso.CreateTextFile(errorLog, True)
            logFile.WriteLine "Office 2016 Patch Extractor Log - " & Now()
            logFile.WriteLine "------------------------------------"
        End If
        logFile.WriteLine "Patch #" & totalPatches & " - Error getting LocalPackage: " & Err.Description
        Err.Clear
    ElseIf localPackage <> "" Then
        If oFso.FileExists(localPackage) Then
            patchDates(localPackage) = oFso.GetFile(localPackage).DateLastModified
            If logFile Is Nothing Then
                Set logFile = oFso.CreateTextFile(errorLog, True)
                logFile.WriteLine "Office 2016 Patch Extractor Log - " & Now()
                logFile.WriteLine "------------------------------------"
            End If
            logFile.WriteLine "Patch #" & totalPatches & " - Detected: " & localPackage
        Else
            If logFile Is Nothing Then
                Set logFile = oFso.CreateTextFile(errorLog, True)
                logFile.WriteLine "Office 2016 Patch Extractor Log - " & Now()
                logFile.WriteLine "------------------------------------"
            End If
            logFile.WriteLine "Patch #" & totalPatches & " - File not found: " & localPackage
        End If
    End If
    On Error Goto 0
Next

' Sort patches by modified date
Dim arrPatches: arrPatches = patchDates.Keys()
Call QuickSort(arrPatches, LBound(arrPatches), UBound(arrPatches))

' Process sorted patches
For Each localPackage In arrPatches
    On Error Resume Next
    Set msp = oMsi.OpenDatabase(localPackage, MSIOPENDATABASEMODE_PATCHFILE)
    If Err.Number = 0 Then
        Set SumInfo = msp.SummaryInformation
        If InStr(SumInfo.Property(PID_TEMPLATES), OFFICEID) > 0 Then
            Set qView = msp.OpenView("SELECT `Property`,`Value` FROM MsiPatchMetadata WHERE `Property`='StdPackageName'")
            qView.Execute
            Set record = qView.Fetch()
            If Not record Is Nothing Then
                patchCount = patchCount + 1
                fileCounter = fileCounter + 1
                
                ' Create filename: num-OriginalFileName.msp
                Dim newFileName: newFileName = fileCounter & "-" & record.StringData(2)
                Dim destFile: destFile = sTargetFolder & newFileName
                
                ' Copy file with retry logic
                Dim retryCount: retryCount = 0
                Dim copySuccess: copySuccess = False
                
                Do While retryCount < 3 And Not copySuccess
                    retryCount = retryCount + 1
                    oFso.CopyFile localPackage, destFile, True
                    If Err.Number <> 0 Then
                        If logFile Is Nothing Then
                            Set logFile = oFso.CreateTextFile(errorLog, True)
                            logFile.WriteLine "Office 2016 Patch Extractor Log - " & Now()
                            logFile.WriteLine "------------------------------------"
                        End If
                        logFile.WriteLine "Patch #" & (fileCounter + 1) & " - Attempt " & retryCount & " - Failed to copy: " & localPackage & " to " & destFile & " - " & Err.Description
                        Err.Clear
                        WScript.Sleep 1000
                    Else
                        copySuccess = True
                        successCount = successCount + 1
                        logFile.WriteLine "Patch #" & (fileCounter + 1) & " - Successfully copied: " & localPackage & " to " & destFile
                    End If
                Loop
                
                If Not copySuccess Then
                    failedCount = failedCount + 1
                    logFile.WriteLine "Patch #" & (fileCounter + 1) & " - Failed to copy after retries: " & localPackage & " to " & destFile
                End If
            Else
                logFile.WriteLine "Patch #" & (fileCounter + 2) & " - No StdPackageName found for: " & localPackage
            End If
        End If
    Else
        If logFile Is Nothing Then
            Set logFile = oFso.CreateTextFile(errorLog, True)
            logFile.WriteLine "Office 2016 Patch Extractor Log - " & Now()
            logFile.WriteLine "------------------------------------"
        End If
        logFile.WriteLine "Patch #" & (fileCounter + 2) & " - Failed to open database for " & localPackage & ": " & Err.Description
        Err.Clear
    End If
    On Error Goto 0
Next

' Show results
sMessage = "Office 2016 Patch extraction complete." & vbCrLf & _
           "Total patches detected: " & totalPatches & vbCrLf & _
           "Office patches found: " & patchCount & vbCrLf & _
           "Successfully copied: " & successCount

If failedCount > 0 Then
    sMessage = sMessage & vbCrLf & "Failed to copy: " & failedCount & vbCrLf & _
               "See error_log.txt for details."
End If

If Not logFile Is Nothing Then
    logFile.WriteLine "------------------------------------"
    logFile.WriteLine "Summary:"
    logFile.WriteLine "Total patches detected: " & totalPatches
    logFile.WriteLine "Office patches found: " & patchCount
    logFile.WriteLine "Successfully copied: " & successCount
    logFile.WriteLine "Failed to copy: " & failedCount
    logFile.Close
End If

oWShell.Popup sMessage, 15, "Office 2016 Updates Extractor"

' Open the folder
oWShell.Run "explorer /e," & Chr(34) & sTargetFolder & Chr(34)

' Clean up
Set oMsi = Nothing
Set oFso = Nothing
Set oWShell = Nothing
Set Patches = Nothing
Set patch = Nothing
Set msp = Nothing
Set qView = Nothing
Set record = Nothing
Set SumInfo = Nothing
Set logFile = Nothing
Set patchDates = Nothing

' QuickSort function to sort by Date Modified (ascending)
Sub QuickSort(arr, low, high)
    Dim i, j, pivot, temp
    If low < high Then
        pivot = arr(low)
        i = low
        j = high
        Do While i < j
            Do While patchDates(arr(j)) >= patchDates(pivot) And i < j
                j = j - 1
            Loop
            Do While patchDates(arr(i)) <= patchDates(pivot) And i < j
                i = i + 1
            Loop
            If i < j Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Loop
        arr(low) = arr(j)
        arr(j) = pivot
        Call QuickSort(arr, low, j - 1)
        Call QuickSort(arr, j + 1, high)
    End If
End Sub