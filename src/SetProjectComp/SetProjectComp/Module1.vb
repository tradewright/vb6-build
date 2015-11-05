Imports System.Collections.Generic
Imports System.IO

Module Module1

    Sub Main()
        Try

            Dim clp = New CommandLineParser(Command, " ")
            Dim filename = getFileName(clp)
            Dim mode = getMode(clp)
            Dim majorVersionNumber = getVersionNumberPart(clp, 1, "Major")
            Dim minorVersionNumber = getVersionNumberPart(clp, 2, "Minor")
            Dim revisionNumber = getVersionNumberPart(clp, 3, "Revision")

            Dim lines = getLines(filename)

            If adjust(lines, mode, majorVersionNumber, minorVersionNumber, revisionNumber) Then
                writeNewFile(filename, lines)
            End If
        Catch e As Exception
            Console.Error.WriteLine("Error : " & e.Message)
        End Try
    End Sub

    Private Function adjust( _
                    ByVal pLines As List(Of String), _
                    ByVal pMode As String, _
                    ByVal pMajorVersionNumber As String, _
                    ByVal pMinorVersionNumber As String, _
                    ByVal pRevisionNumber As String) As Boolean
        Dim adjusted = False
        Dim versionAdjusted = False

        Dim i = 0
        Do While i < pLines.Count
            Dim s = pLines(i)
            If s = "VersionCompatible32=""1""" Then
                ' this line will be rewritten only if needed, so delete it
                pLines.RemoveAt(i)
                i -= 1
            ElseIf s.StartsWith("CompatibleMode=") Then
                adjusted = adjustCompatibleMode(pLines, i, pMode)
            ElseIf s.StartsWith("MajorVer=") Then
                versionAdjusted = adjustVersion(pLines, i, "MajorVer=", pMajorVersionNumber)
            ElseIf s.StartsWith("MinorVer=") Then
                versionAdjusted = adjustVersion(pLines, i, "MinorVer=", pMinorVersionNumber)
            ElseIf s.StartsWith("RevisionVer=") Then
                versionAdjusted = adjustVersion(pLines, i, "RevisionVer=", pRevisionNumber)
            End If
            i += 1
        Loop

        If versionAdjusted Then
            Console.WriteLine(String.Format("Version set to {0}.{1}.{2}", pMajorVersionNumber, pMinorVersionNumber, pRevisionNumber))
        End If
        Return adjusted Or versionAdjusted
    End Function

    Private Function adjustCompatibleMode(ByVal pLines As List(Of String), ByRef pIndex As Integer, ByVal pMode As String) As Boolean
        Dim adjustedLine = "CompatibleMode=" & """" & getModeNumber(pMode) & """"

        If pLines(pIndex) = adjustedLine Then Return False

        pLines(pIndex) = adjustedLine
        If pMode = "B" Then
            pLines.Insert(pIndex + 1, "VersionCompatible32=""1""")
            pIndex += 1
        End If
        Console.WriteLine("Project compatibility adjusted to " & pMode)
        Return True
    End Function

    Private Function adjustVersion(ByVal pLines As List(Of String), ByRef pIndex As Integer, ByVal pLinePrefix As String, ByVal pVersionNumber As String) As Boolean
        Dim adjustedLine = pLinePrefix & pVersionNumber

        If pLines(pIndex) = adjustedLine Then Return False

        pLines(pIndex) = adjustedLine
        Return True
    End Function

    Private Function getModeNumber(ByVal pMode As String) As String
        Dim result = String.Empty
        If pMode = "P" Then
            result = "1"
        ElseIf pMode = "B" Then
            result = "2"
        ElseIf pMode = "N" Then
            result = "0"
        End If
        Return result
    End Function

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

    Private Function getVersionNumberPart(ByVal pClp As CommandLineParser, ByVal argumentNumber As Integer, ByVal versionPart As String) As String
        Dim versionString As String
        versionString = pClp.Arg(argumentNumber)
        If versionString = "" Then Throw New ArgumentException(String.Format("Product {0} version must be supplied as argument {1}", versionPart, argumentNumber))
        Dim version As Integer
        If Not Integer.TryParse(versionString, version) Or version < 0 Or version > 9999 Then Throw New ArgumentException(String.Format("Product {0} version must be an integer 0-9999", versionPart))
        Return versionString
    End Function

    Private Sub writeNewFile(ByVal pFilename As String, ByRef pLines As List(Of String))
        Dim writer = File.CreateText(pFilename)

        For Each s In pLines
            writer.WriteLine(s)
        Next

        writer.Close()
    End Sub

End Module
