Imports System.Collections.Generic
Imports System.IO

Module Module1

    Sub Main()
        Try

            Dim clp = New CommandLineParser(Command, " ")
            Dim filename = getFileName(clp)
            Dim mode = getMode(clp)
            Dim revisionNumber As Integer = getRevision(clp)

            Dim lines = getLines(filename)

            Dim modeUpdateNeeded As Boolean
            Dim versionUpdateNeeded As Boolean
            checkIfChangesNeeded(lines, mode, revisionNumber, modeUpdateNeeded, versionUpdateNeeded)

            If modeUpdateNeeded Or versionUpdateNeeded Then
                If modeUpdateNeeded Then adjustMode(lines, mode)
                If versionUpdateNeeded Then adjustVersion(lines, revisionNumber)
                writeNewFile(filename, lines)
            End If
        Catch e As Exception
            Console.Error.WriteLine("Error : " & e.Message)
        End Try
    End Sub

    Private Sub adjustMode( _
                    ByVal pLines As List(Of String), _
                    ByVal pMode As String)
        Dim i = 0
        Do While i < pLines.Count
            Dim s = pLines(i)
            If s = "VersionCompatible32=""1""" Then
                ' this line will be rewritten only if needed, so delete it
                pLines.Remove(i)
                i = i - 1
            ElseIf startsWith(s, "CompatibleMode=") Then
                If pMode = "P" Then
                    pLines.Remove(i)
                    pLines.Insert(i, "CompatibleMode=""1""")
                    Console.WriteLine("File adjusted to Project Compatibility")
                ElseIf pMode = "B" Then
                    pLines.Remove(i)
                    pLines.Insert(i, "CompatibleMode=""2""")
                    pLines.Insert(i + 1, "VersionCompatible32=""1""")
                    i = i + 1
                    Console.WriteLine("File adjusted to Binary Compatibility")
                ElseIf pMode = "N" Then
                    pLines.Remove(i)
                    pLines.Insert(i, "CompatibleMode=""0""")
                    Console.WriteLine("File adjusted to No Compatibility")
                End If
            End If
            i = i + 1
        Loop
    End Sub

    Private Sub adjustVersion( _
                    ByVal pLines As List(Of String), _
                    ByVal pRevisionNumber As Integer)
        Dim i = 0
        Do While i < pLines.Count
            Dim s = pLines(i)
            If s.StartsWith("RevisionVer=") Then
                pLines.Remove(i)
                pLines.Insert(i, "RevisionVer=" & CStr(pRevisionNumber))
                Console.WriteLine("Revision version set to " & CStr(pRevisionNumber))
            End If
            i = i + 1
        Loop
    End Sub

    Private Sub checkIfChangesNeeded( _
                    ByVal pLines As List(Of String), _
                    ByVal pMode As String, _
                    ByVal pRevisionNumber As Integer, _
                    ByRef pModeUpdateNeeded As Boolean, _
                    ByRef pVersionUpdateNeeded As Boolean)
        For Each s In pLines
            If s.StartsWith("CompatibleMode=") Then
                If pMode = "P" And s = "CompatibleMode=""1""" Then
                    Console.WriteLine("Already in Project Compatibility mode")
                ElseIf pMode = "B" And s = "CompatibleMode=""2""" Then
                    Console.WriteLine("Already in Binary Compatibility mode")
                Else
                    pModeUpdateNeeded = True
                End If
            ElseIf s.StartsWith("RevisionVer=") Then
                If s <> ("RevisionVer=" & CLng(pRevisionNumber)) Then pVersionUpdateNeeded = True
            End If
        Next
    End Sub

    Private Function getFileName(ByVal pClp As CommandLineParser) As String
        getFileName = pClp.Arg(0)
        If getFileName = "" Then Throw New ArgumentException("Project filename must be supplied as first argument")
    End Function

    Private Function getLines(ByVal pFilename As String) As List(Of String)
        Dim reader = File.OpenText(pFilename)

        Dim lines = New List(Of String)

        Do While Not reader.EndOfStream
            lines.Add(reader.ReadLine)
        Loop

        reader.Close()

        Return lines
    End Function

    Private Function getMode(ByVal pClp As CommandLineParser) As String
        If Not pClp.IsSwitchSet("mode") Then Throw New ArgumentException("Must supply /mode switch")

        getMode = UCase$(pClp.SwitchValue("mode"))
        If getMode <> "P" And getMode <> "B" And getMode <> "N" Then Throw New ArgumentException("Mode must be 'P' or 'B' or 'N'")
    End Function

    Private Function getRevision(ByVal pClp) As Integer
        Dim revString As String
        revString = pClp.Arg(1)
        If revString = "" Then Throw New ArgumentException("Product revision number must be supplied as second argument")
        Dim revision As Integer
        If Not Integer.TryParse(revString, revision) Or revision < 0 Or revision > 9999 Then Throw New ArgumentException("Product revision number must be an integer 0-9999")
        Return revision
    End Function

    Private Function startsWith(ByVal s As String, ByVal pSubStr As String) As Boolean
        startsWith = (Left$(UCase$(s), Len(pSubStr)) = UCase$(pSubStr))
    End Function

    Private Sub writeNewFile(ByVal pFilename As String, ByRef pLines As List(Of String))
        Dim writer = File.CreateText(pFilename)

        For Each s In pLines
            writer.WriteLine(s)
        Next

        writer.Close()
    End Sub

End Module
