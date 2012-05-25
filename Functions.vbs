'Copyright (c) 2012 David Heinemann
'
'Permission is hereby granted, free of charge, to any person obtaining a
'copy of this software and associated documentation files (the
'"Software"), to deal in the Software without restriction, including
'without limitation the rights to use, copy, modify, merge, publish,
'distribute, sublicense, and/or sell copies of the Software, and to
'permit persons to whom the Software is furnished to do so, subject to
'the following conditions:
'
'The above copyright notice and this permission notice shall be included
'in all copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
'OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
'MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
'IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
'CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
'TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
'SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

OPTION EXPLICIT
' Confirm that the backup directory exists
Function checkIfDirExists(ByVal givenDir)
    Dim objFso: Set objFso = CreateObject("Scripting.FileSystemObject")
    If Not objFso.FolderExists(givenDir) Then
        checkIfDirExists = False
    Else
        checkIfDirExists = True
    End If
End Function

' Fetch a list of DBs on the instance
Function getDatabaseList
    WScript.Echo "Generating a list of available databases..."
    Dim dbListCmd
    dbListCmd = """" + mysqlDir + "mysql.exe"""
    dbListCmd = dbListCmd + " --user=" + username
    dbListCmd = dbListCmd + " --password=" + password
    dbListCmd = dbListCmd + " --port=" + port
    dbListCmd = dbListCmd + " --execute ""show databases"" --skip-column-names"

    Dim objShell: Set objShell = WScript.CreateObject("WScript.Shell")
    Dim objExec: Set objExec = objShell.Exec(dbListCmd)

    Do While objExec.Status = 0
        WScript.Sleep 1000
    Loop

    Dim dbList(): ReDim dbList(0)
    Do While objExec.StdOut.AtEndOfStream <> True
        Dim db: db = objExec.StdOut.ReadLine

        If dbList(UBound(dbList)) <> "" Then
            ReDim Preserve dbList(UBound(dbList) + 1)
        End If

        dbList(UBound(dbList)) = db
    Loop

    getDatabaseList = dbList
End Function

' Backup a given database
Sub backupDatabase(ByVal db)
    WScript.Echo ("Backing up " + db + "...")
    Dim backupTime: backupTime = getBackupTime
    Dim backupDate: backupDate = getBackupDate
    Dim dbBackupFile: dbBackupFile = db + "-" + backupDate + "-" + backupTime + ".sql"
    Dim dbBackupArchive: dbBackupArchive = dbBackupFile + "." + archiveExt
    Dim finalBackupDir: finalBackupDir = rootBackupDir + "\" + db + "\"

    Dim dbBackupCmd
    dbBackupCmd = "cmd.exe /C """ + mysqlDir + "mysqldump.exe"""
    dbBackupCmd = dbBackupCmd + " " + db
    dbBackupCmd = dbBackupCmd + " > " + dbBackupFile

    Dim archiveCmd
    archiveCmd = """" + szExe + """" + " a"
    archiveCmd = archiveCmd + " -t" + archiveType
    archiveCmd = archiveCmd + " -mx" + compLevel
    archiveCmd = archiveCmd + " " + dbBackupArchive
    archiveCmd = archiveCmd + " " + dbBackupFile

    Dim objShell: Set objShell = WScript.CreateObject("WScript.Shell")
    objShell.Run dbBackupCmd, 0, True
    objShell.Run archiveCmd, 0, True

    Dim objFso: Set objFso = CreateObject("Scripting.FileSystemObject")
    ' Delete uncompressed backup
    If objFso.FileExists(dbBackupFile) Then
        Dim objFile: Set objFile = objFso.GetFile(dbBackupFile)
        objFile.Delete
    End If

    If Not objFso.FolderExists(finalBackupDir) Then
        objFso.CreateFolder(finalBackupDir)
    End If

    ' Move the compressed backup
    If Not objFso.FileExists(finalBackupDir + dbBackupArchive) Then
        If objFso.FileExists(dbBackupArchive) Then
            objFso.MoveFile dbBackupArchive, finalBackupDir
        End If
    End If

    'Finally, remove any outdated backups.
    pruneBackups db, pruneAge
End Sub

' Convert the current time into 24 hour HHMM format
Function getBackupTime
    Dim timeArray: timeArray = Split(Time(), ":")

    ' Convert time to 24 hours
    If Mid(timeArray(2), 4, 2) = "PM" Then
        timeArray(0) = timeArray(0) + 12
    ElseIf Len(timeArray(0)) = 1 Then
        timeArray(0) = "0" + timeArray(0)
    End If

    getBackupTime = CStr(timeArray(0)) + CStr(timeArray(1))
End Function

' Convert the current date into YYYY-MM-DD format
Function getBackupDate
    Dim dateArray: dateArray = Split(Date(), "/")
    getBackupDate = dateArray(2) + "-" + dateArray(1) + "-" + dateArray(0)
End Function

' Ensure each directory variable has a trailing backslash
Function correctDirSlashes(ByVal directory)
    Dim lastChar: lastChar = Mid(directory, Len(directory), 1)
    If lastChar = "\" Then
        correctDirSlashes = directory
    Else
        correctDirSlashes = directory + "\"
    End If
End Function

' Delete backups older than a given age (in days)
Sub pruneBackups(ByVal db, ByVal maxAgeInDays)
    WScript.Echo "Pruning " + db + " backups..."
    Dim objFso: Set objFso = CreateObject("Scripting.FileSystemObject")
    Dim dbDir: Set dbDir = objFso.GetFolder(rootBackupDir + db)

    Dim dbBackup
    For Each dbBackup in dbDir.Files
        Dim backupAge: backupAge = Datediff("d", dbBackup.DateCreated, date())
        If backupAge > maxAgeInDays Then
            dbBackup.Delete
            WScript.Echo dbBackup + " deleted (Age is " + CStr(backupAge) + " days)."
        End If
    Next
End Sub
