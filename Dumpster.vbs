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
'''''''''''''''''''''''''''''''
' Settings
'''''''''''''''''''''''''''''''
' PATH TO MYSQL BIN DIRECTORY
Dim mysqlDir: mysqlDir = "C:\Program Files\MySQL\MySQL Server 5.5\bin"

' PATH TO 7ZA.EXE
' 7za.exe is the command-line version of 7-zip; download it from
' http://www.7-zip.org/download.html
Dim szExe: szExe = "7za" 

' PATH TO BACKUP DIRECTORY
' Must have trailing slash
Dim backupDir: backupDir = "C:\foo\bar\"

' MYSQL USERNAME
Dim username: username = "root"

' MYSQL PASSWORD
Dim password: password = "xyzzy"

' MYSQL PORT
Dim port: port = "3306"

' DATABASES TO IGNORE
Dim ignoreList: ignoreList = Array("information_schema", "performance_schema", "mysql")

'''''''''''''''''''''''''''''''
' Advanced Settings
'''''''''''''''''''''''''''''''
' ARCHIVE FORMAT
' e.g. zip, 7z, gzip, tar, etc. See 7-Zip help for more information.
Dim archiveType: archiveType = "zip"

' ARCHIVE EXTENSION
' Should match the format selected above.
Dim archiveExt: archiveExt = "zip"

' ARCHIVE COMPRESSION LEVEL
' Must be between 0 & 9. 0 = no compression; 9 = Ultra compression.
Dim compLevel: compLevel = "9"

'''''''''''''''''''''''''''''''
' The Serious Business
'''''''''''''''''''''''''''''''
' Confirm that the backup directory exists
Dim objFso:  Set objFso  = CreateObject("Scripting.FileSystemObject")
If Not objFso.FolderExists(backupDir) Then
    WScript.Echo "Error: Target backup directory (" + backupDir + ") does not exist. Please create it."
    WScript.Quit
End If

Dim dbList, db
dbList = getDatabaseList
For Each db In dbList
    Dim ignoredDb
    Dim isIgnored: isIgnored = False

    ' Ignore blacklisted DBs
    For Each ignoredDb In ignoreList
        If db = ignoredDb Then
            isIgnored = True
        End If
    Next

    If isIgnored = False Then
        backupDatabase(db)
    End If
Next

' Fetch a list of DBs on the instance
Function getDatabaseList
    Dim dbListCmd
    dbListCmd = """" + mysqlDir + "\mysql.exe"""
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
    Dim backupTime: backupTime = getBackupTime
    Dim backupDate: backupDate = getBackupDate
    Dim dbBackupFile: dbBackupFile = db + "-" + backupDate + "-" + backupTime + ".sql"
    Dim dbBackupArchive: dbBackupArchive = dbBackupFile + "." + archiveExt

    Dim dbBackupCmd
    dbBackupCmd = "cmd.exe /C """ + mysqlDir + "\mysqldump.exe"""
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

    Dim objFso:  Set objFso  = CreateObject("Scripting.FileSystemObject")
    ' Delete uncompressed backup
    If objFso.FileExists(dbBackupFile) Then
        Dim objFile: Set objFile = objFso.GetFile(dbBackupFile)
        objFile.Delete
    End If

    ' Move the compressed backup
    If Not objFso.FileExists(backupDir + dbBackupArchive) Then
        If objFso.FileExists(dbBackupArchive) Then
            objFso.MoveFile dbBackupArchive, backupDir
        End If
    End If
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
