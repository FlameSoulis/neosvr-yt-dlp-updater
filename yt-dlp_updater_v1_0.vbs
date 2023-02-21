FILE_TO_UPDATE = "yt-dlp.exe"
APP_TITLE = "NEOS yt-dlp Updater v1.0"

Set WshShell = CreateObject("WScript.Shell")
Set fileSys = CreateObject("Scripting.FileSystemObject")

introMsg = "Welcome to the Unofficial NeosVR yt-dlp Updater!" & vbCrLf
introMsg = introMsg & "Original by Knackrack615; Modified by Flame Soulis" & vbCrLf & vbCrLf
introMsg = introMsg & "Are you updating the Steam Build (Yes) or Website build (No)?"
response = MsgBox(introMsg, vbQuestion + vbYesNo, APP_TITLE)

If response = vbNo Then
    ' Let's check the obvious: did the user just put it in 'C:/Neos'?
    inputMsg = "Please specify where NeosVR was installed to." & vbCrLf & vbCrLf
    inputMsg = inputMsg & "If unsure, the default installation has been filled in."
    installPath = InputBox(inputMsg, APP_TITLE, "C:/Neos")

    If fileSys.FolderExists(installPath) = False Then
        ' Welp, we tried!
        errorMsg = "The specified location does not exist!" & vbCrLf & vbCrLf
        errorMsg = errorMsg & "Please try run again."
        MsgBox errorMsg, vbCritical, "Failed - " & APP_TITLE
        Wscript.Quit
    End If

    ' This part is easy!
    installPath = installPath & "/app/RuntimeData"
    ' Does THIS place exist?
    If fileSys.FolderExists(installPath) = False Then
        ' Uh...
        errorMsg = "The specified location exists, but the app's folder appears to be empty!" & vbCrLf & vbCrLf
        errorMsg = errorMsg & "Has NeosVR been downloaded?"
        MsgBox errorMsg, vbCritical, "Failed - " & APP_TITLE
        Wscript.Quit
    End If

    'Home stretch!
    executablePath = installPath & "/" & FILE_TO_UPDATE
Else
    ' Knackrack's Method
    ' Prepare for unforseen consequences
    Err.Clear
    On Error Resume Next
    keyPath = "HKEY_CLASSES_ROOT\neos\Registered Path"
    registeredPath = WshShell.RegRead(keyPath)

    ' Okay, do we even have a result from the registry?
    If Err.Number <> 0 Then
        ' Bail!
        bailMsg = "We could not find NeosVR's installation location in the registry!" & vbCrLf & vbCrLf
        bailMsg = bailMsg & "Is it actually installed via Steam?"
        MsgBox bailMsg, vbCritical, "Failed - " & APP_TITLE
        Wscript.Quit
    End If

    ' Continue
    executablePath = registeredPath & "\" & FILE_TO_UPDATE
End If

' If we are here, then we have out information in 'executablePath'
if fileSys.FileExists(executablePath) = False Then
    ' WHY?!?!!?!
    bailMsg = "We could not find yt-dlp.exe in the following path:" & vbCrLf & vbCrLf
    bailMsg = bailMsg & executablePath & vbCrLf & vbCrLf
    bailMsg = bailMsg & "Please navigate to your NEOS installation directory and update yt-dlp manually."
    MsgBox bailMsg, vbCritical, "Failed - " & APP_TITLE
    Wscript.Quit
End If

' Prompt user to confirm the path is correct
confirmMsg = "yt-dlp has been detected in the following location:" & vbCrLf & vbCrLf
confirmMsg = confirmMsg & executablePath & vbCrLf & vbCrLf & "Is this correct?"
response = MsgBox(confirmMsg, vbQuestion + vbYesNo, "Confirm Path - " & APP_TITLE)

If response = vbYes Then
    ' Hide window and run executable with "-U" argument
    WshShell.Run """" & executablePath & """ -U", 0, True

    ' Show message box when execution is complete
    resultMsg = "yt-dlp has been updated to the latest version!" & vbCrLf & vbCrLf
    resultMsg = resultMsg & "You should now be able to watch YouTube videos normally."
    MsgBox resultMsg, vbInformation, "Update Completed - " & APP_TITLE
Else
    ' Show message box if user clicked "No"
      cancelMsg = "Execution cancelled." & vbCrLf & vbCrLf
      cancelMsg = cancelMsg & "Please navigate to your NEOS installation directory and update yt-dlp manually."
      MsgBox cancelMsg, vbExclamation, "Cancelled - " & APP_TITLE
End If