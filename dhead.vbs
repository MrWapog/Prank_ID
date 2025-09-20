' prank_with_congrats.vbs
Option Explicit

' CONFIG - tweak these as you like
Const MAX_SPAWNS = 25    ' maximum total extra windows allowed across all instances
Const YES_SPAWNS = 10    ' how many "Good luck..." boxes to open when user clicks Yes

Dim arg0, spawnCount, resp, i

' If this instance was launched just to show the congrats message, show it and exit
If WScript.Arguments.Count > 0 Then
    arg0 = WScript.Arguments(0)
    If LCase(arg0) = "congrats" Then
        MsgBox "Good luck on your journey, smart one", vbInformation, "Congratulations"
        WScript.Quit
    End If
End If

' Otherwise, try to read spawnCount from argument (numeric)
spawnCount = 0
If WScript.Arguments.Count > 0 Then
    On Error Resume Next
    spawnCount = CInt(WScript.Arguments(0))
    On Error GoTo 0
End If

Do
    resp = MsgBox("Are you an idiot?", vbYesNo + vbExclamation, "Hi dhead")

    If resp = vbNo Then
        If spawnCount < MAX_SPAWNS Then
            spawnCount = spawnCount + 1
            Dim shell
            Set shell = CreateObject("WScript.Shell")
            shell.Run "wscript """ & WScript.ScriptFullName & """ " & spawnCount, 1, False
            Set shell = Nothing
        End If
        ' If spawn limit reached, No will not spawn more windows.
    ElseIf resp = vbYes Then
        ' On Yes: spawn up to YES_SPAWNS "good luck" boxes but do not exceed MAX_SPAWNS
        For i = 1 To YES_SPAWNS
            If spawnCount < MAX_SPAWNS Then
                spawnCount = spawnCount + 1
                Dim shell2
                Set shell2 = CreateObject("WScript.Shell")
                shell2.Run "wscript """ & WScript.ScriptFullName & """ congrats", 1, False
                Set shell2 = Nothing
            Else
                Exit For
            End If
        Next
        Exit Do
    End If

Loop

' Final message in the original instance (optional)
MsgBox "HAHAHA, Knew it!", vbInformation, ":)"
