''''
' SYNOPSIS
'   Utility class to perform common functions.
''''''''
Class UtilityBelt
    ''''
    ' SYNOPSIS
    '   Pauses script execution until the Enter key is pressed.
    ''''''''
    Sub pause(ByVal message)
        WScript.Echo(message)
        Dim key: key = WScript.StdIn.Read(1)
    End Sub

    ''''
    ' SYNOPSIS
    '   Re-starts the script using CScript if it was launched with
    '   WScript.
    ''''''''
    Sub forceConsoleMode(ByVal permanent)
        'http://ask.metafilter.com/79481/vbscript-printing-to-command-line
        Dim strStartExe: strStartExe = UCase(Mid(WScript.FullName, instrRev(WScript.FullName, "\") + 1))

        If Not strStartExe = "CSCRIPT.EXE" Then
            Dim arguments, argument
            For Each argument In WScript.Arguments
                arguments = arguments & " " & argument
            Next

            Dim wshell: Set wshell = CreateObject("WScript.Shell")
            If permanent Then
                wshell.Run("cmd /k cscript.exe """ & WScript.ScriptFullName & """ " & arguments)
            Else
                wshell.Run("cscript.exe """ & WScript.ScriptFullName & """ " & arguments)
            End If
            Set wshell = Nothing
            WScript.Quit
        End If
    End Sub

    Sub clearScreen
        WScript.StdOut.WriteBlankLines(50)
    End Sub
End Class
