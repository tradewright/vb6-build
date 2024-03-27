Imports System.Collections.Generic
Imports System.IO

Module Module1

    Sub Main()
        Try

            Dim clp = New CommandLineParser(Command, " ")
            Dim projectFilename = getProjectFileName(clp)
            Dim mode = getMode(clp)
            Dim majorVersionNumber = getVersionNumberPart(clp, 1, "Major")
            Dim minorVersionNumber = getVersionNumberPart(clp, 2, "Minor")
            Dim revisionNumber = getVersionNumberPart(clp, 3, "Revision")
            Dim normaliseRefsAndObjects = clp.IsSwitchSet("n")
            Dim objectFilename = getObjectFilename(clp)

            Dim lines = getLines(projectFilename)

            If adjust(lines, mode, majorVersionNumber, minorVersionNumber, revisionNumber, objectFilename, normaliseRefsAndObjects) Then
                writeNewFile(projectFilename, lines)
            End If
        Catch e As Exception
            Console.Error.WriteLine("Error : " & e.Message)
        End Try
    End Sub

    Private Function adjust(
                    pLines As List(Of String),
                    pMode As String,
                    pMajorVersionNumber As String,
                    pMinorVersionNumber As String,
                    pRevisionNumber As String,
                    pObjectFilename As String,
                    pNormaliseRefsAndObjects As Boolean) As Boolean
        Dim adjusted = False
        Dim versionAdjusted = False

        Dim i = 0
        Do While i < pLines.Count
            Dim s = pLines(i)
            If s = "VersionCompatible32=""1""" Then
                ' this line will be rewritten only if needed, so delete it
                pLines.RemoveAt(i)
                adjusted = True
                i -= 1
            ElseIf s.StartsWith("CompatibleMode=") Then
                adjusted = adjusted Or adjustCompatibleMode(pLines, i, pMode)
            ElseIf s.StartsWith("MajorVer=") Then
                versionAdjusted = versionAdjusted Or adjustVersion(pLines, i, "MajorVer=", pMajorVersionNumber)
            ElseIf s.StartsWith("MinorVer=") Then
                versionAdjusted = versionAdjusted Or adjustVersion(pLines, i, "MinorVer=", pMinorVersionNumber)
            ElseIf s.StartsWith("RevisionVer=") Then
                versionAdjusted = versionAdjusted Or adjustVersion(pLines, i, "RevisionVer=", pRevisionNumber)
            ElseIf s.StartsWith("ExeName32=") Then
                adjusted = adjusted Or adjustExeName(pLines, i, pObjectFilename)
            ElseIf s.StartsWith("Reference=") And pNormaliseRefsAndObjects Then
                adjusted = adjusted Or normaliseReference(pLines, i)
            ElseIf s.StartsWith("Object=") And pNormaliseRefsAndObjects Then
                adjusted = adjusted Or normaliseObject(pLines, i)
            End If
            i += 1
        Loop

        If versionAdjusted Then
            Console.WriteLine(String.Format("Version set to {0}.{1}.{2}", pMajorVersionNumber, pMinorVersionNumber, pRevisionNumber))
        End If
        Return adjusted Or versionAdjusted
    End Function

    Private Function adjustCompatibleMode(pLines As List(Of String), ByRef pIndex As Integer, pMode As String) As Boolean
        Dim adjustedLine = $"CompatibleMode=""{getModeNumber(pMode)}"""

        If pLines(pIndex) = adjustedLine Then Return False

        pLines(pIndex) = adjustedLine
        If pMode = "B" Then
            pLines.Insert(pIndex + 1, "VersionCompatible32=""1""")
            pIndex += 1
        End If
        Console.WriteLine("Project compatibility adjusted to " & pMode)
        Return True
    End Function

    Private Function adjustExeName(pLines As List(Of String), ByRef pIndex As Integer, pObjectFilename As String) As Boolean
        If pObjectFilename = "" Then Return False
        Dim adjustedLine = $"ExeName32=""{pObjectFilename}"""
        If adjustedLine.ToUpperInvariant() = pLines(pIndex).ToUpperInvariant() Then Return False
        pLines(pIndex) = adjustedLine
        Return True
    End Function

    Private Function adjustVersion(pLines As List(Of String), ByRef pIndex As Integer, pLinePrefix As String, pVersionNumber As String) As Boolean
        Dim adjustedLine = pLinePrefix & pVersionNumber

        If pLines(pIndex) = adjustedLine Then Return False

        pLines(pIndex) = adjustedLine
        Return True
    End Function

    Private Function getModeNumber(pMode As String) As String
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

    Private Function getObjectFileName(pClp As CommandLineParser) As String
        getObjectFileName = pClp.SwitchValue("ObjectFileName")
    End Function

    Private Function getProjectFileName(pClp As CommandLineParser) As String
        getProjectFileName = pClp.Arg(0)
        If getProjectFileName = "" Then Throw New ArgumentException("Project filename must be supplied as first argument")
    End Function

    Private Function getLines(pFilename As String) As List(Of String)
        Dim reader = File.OpenText(pFilename)

        Dim lines = New List(Of String)

        Do While Not reader.EndOfStream
            lines.Add(reader.ReadLine)
        Loop

        reader.Close()

        Return lines
    End Function

    Private Function getMode(pClp As CommandLineParser) As String
        If Not pClp.IsSwitchSet("mode") Then Throw New ArgumentException("Must supply /mode switch")

        getMode = UCase$(pClp.SwitchValue("mode"))
        If getMode <> "P" And getMode <> "B" And getMode <> "N" Then Throw New ArgumentException("Mode must be 'P' or 'B' or 'N'")
    End Function

    Private Function getVersionNumberPart(pClp As CommandLineParser, argumentNumber As Integer, versionPart As String) As String
        Dim versionString As String
        versionString = pClp.Arg(argumentNumber)
        If versionString = "" Then Throw New ArgumentException($"Product {versionPart} version must be supplied as argument {argumentNumber}")
        Dim version As Integer
        If Not Integer.TryParse(versionString, version) Or version < 0 Or version > 9999 Then Throw New ArgumentException($"Product {versionPart} version must be an integer 0-9999")
        Return versionString
    End Function

    Private Function normaliseObject(pLines As List(Of String), pIndex As Integer) As Boolean
        Dim adjusted As Boolean

        Dim elements = pLines(pIndex).Split({"#"c})
        If elements(1) <> "0.0" Then
            elements(1) = "0.0"
            adjusted = True
        End If

        Dim s = elements(2).ToUpperInvariant()
        If s <> elements(2) Then
            elements(2) = s
            adjusted = True
        End If

        If adjusted Then pLines(pIndex) = String.Join("#", elements)
        Return adjusted
    End Function

    Private Function normaliseReference(pLines As List(Of String), pIndex As Integer) As Boolean
        Dim adjusted As Boolean
        Dim s As String

        Dim elements = pLines(pIndex).Split({"#"c})

        If elements(1) <> "0.0" Then
            elements(1) = "0.0"
            adjusted = True
        End If

        Dim pathEls = elements(3).ToUpperInvariant.Split({"\"c})
        ' note that this element in some References ends with something like this: ...\vbscript.dll\3
        For i = 0 To pathEls.Length - 1
            Dim pathEl = pathEls(i)
            If pathEl.IndexOf(".DLL") <> -1 Or pathEl.IndexOf(".TLB") <> -1 Then
                s = pathEl
                If i <> pathEls.Length - 1 Then s = s & "\" & pathEls(i + 1)
                Exit For
            End If
        Next
        If elements(3) <> s Then
            elements(3) = s
            adjusted = True
        End If

        s = elements(4).ToUpperInvariant()
        If elements(4) <> s Then
            elements(4) = s
            adjusted = True
        End If

        If adjusted Then pLines(pIndex) = String.Join("#", elements)
        Return adjusted
    End Function

    Private Sub writeNewFile(pFilename As String, ByRef pLines As List(Of String))
        Dim writer = File.CreateText(pFilename)

        For Each s In pLines
            writer.WriteLine(s)
        Next

        writer.Close()
    End Sub

End Module
