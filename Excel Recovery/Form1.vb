Option Explicit On
Imports System.IO
Imports System.Text.RegularExpressions
Public Class Form1

#Region "Functions"

    Public Function DelFromLeft(ByVal sChars As String, _
            ByVal sLine As String) As String

        ' Removes unwanted characters from left of given string
        '  EXAMPLE
        '      MsgBox DelFromLeft("THIS", "THIS IS A TEST")
        '        displays  "IS A TEST"


        Dim iCount As Integer
        Dim sChar As String

        DelFromLeft = ""
        ' Remove unwanted characters to left of folder name
        If InStr(sLine, sChars) > 0 Then
            For iCount = 1 To Len(sChars)
                ' Retrieve character from start string to 
                'look for in folder string (sLine)
                sChar = Mid$(sChars, iCount, 1)
                ' Remove all characters to left of found string
                sLine = Mid$(sLine, InStr(sLine, sChar) + 1)

            Next iCount
        End If
        DelFromLeft = sLine
        Exit Function

    End Function

#End Region

#Region "DIMs"

    Dim shadowLinkFolderName As New List(Of String)
    Dim nonErrorShadowPathList As New List(Of String)
    Dim filename As String
    Dim sFileShadowPath As String
    Dim sFileShadowName As String
    Dim sFileShadowPathDate As String
    Dim pathToComboBoxSelectedFile As String
    Dim selectedsFileShadowPathDate As String
    Dim selectedPreviousVersion As String
    Dim saveShadowPath As String
    Dim matchCount As Integer
    Dim counterVariable As Integer
    Dim previousVersionCounterVariable As Integer
    Dim comboBoxChoiceIndex As Integer = 0
    Dim comboBoxIndex As Integer = 0

#End Region

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Dim OFD As New OpenFileDialog
        Dim matches As MatchCollection
        Dim objProcessReader As StreamReader = Nothing
        Dim objProcessErrorReader As StreamReader = Nothing
        Dim myStreamWriter As StreamWriter = Nothing
        Dim fileCompareWriter As StreamWriter = Nothing
        Dim fileCompareReader As StreamReader = Nothing
        Dim fileCompareBoolean As Boolean
        Dim fileCompBooleanError As Boolean
        Dim objProcessOut As String
        Dim objProcessError As String
        Dim driveLetter As String
        Dim sFile As String
        Dim fileCompareOut As String
        Dim previoussFileShadowPath As String
        Dim prevsFileShadowPathDate As String
        Dim sFileShadowSize As Long
        Dim prevsFileShadowPathSize As Long

        Try
            ComboBox1.Items.Clear()

            With OFD

                .ShowDialog()
                filename = .FileName
                PathTb.Text = .FileName

            End With

            ComboBox1.Items.Clear()
            sFile = PathTb.Text
            shadowLinkFolderName.Clear()
            nonErrorShadowPathList.Clear()

            MsgBox("Please wait while Previous Version Explorer searches for previous " _
                   & "versions of your file.", MsgBoxStyle.Information)

            'Find out the number of vss shadow snapshots (restore 
            'points). All shadows apparently have a linkable path 
            '\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy#,
            'where # is a simple one or two or three digit integer.

            Using objProcess As Process = New Process

                objProcess.StartInfo.UseShellExecute = False
                objProcess.StartInfo.CreateNoWindow = True
                objProcess.StartInfo.RedirectStandardOutput = True
                objProcess.StartInfo.RedirectStandardError = True
                objProcess.StartInfo.FileName() = "vssadmin"
                objProcess.StartInfo.Arguments() = "List Shadows"
                objProcess.Start()

                objProcessReader = objProcess.StandardOutput
                objProcessOut = objProcessReader.ReadToEnd
                objProcessErrorReader = objProcess.StandardError
                objProcessError = objProcessErrorReader.ReadToEnd
                objProcess.WaitForExit()
                objProcess.Close()

            End Using

            'MsgBox("objProcessOut is: " & objProcessOut & " and " & _
            '"objProcessError is: " & objProcessError)

            ' Call Regex.Matches method.

            driveLetter = sFile.Substring(0, 2)

            matches = Regex.Matches(objProcessOut, _
            "\\\\\?\\GLOBALROOT\\Device\\HarddiskVolumeShadowCopy[0-9]+")
            counterVariable = 0
            matchCount = matches.Count

            If matchCount = 0 Then

                MsgBox("There are no saved Restore Points for your machine, so no previous versions " _
                       & "of your file are avilable. To enable previous versions of your file, you " _
                       & "must turn on System Protection for the drive your file is stored on. You " _
                       & "turn on System Protection in the System App of the Control Panel. This will " _
                       & "allow you to recover future previous versions not current ones.")

                Exit Sub

            Else

                ' Loop over matches.

                For Each m As Match In matches

                    'MsgBox(m.ToString)
                    shadowLinkFolderName.Add(driveLetter & "\" & DelFromLeft( _
                        "\\?\GLOBALROOT\Device\HarddiskVolume", (m.ToString())))
                    sFileShadowPath = (shadowLinkFolderName(counterVariable) & DelFromLeft( _
                        driveLetter, sFile))
                    'MsgBox(sFileShadowPath)

                    'Here I create temporary folders off the C: 
                    'drive which are mapped to each snapshot.

                    Using myProcess As Process = New Process

                        myProcess.StartInfo.FileName = "cmd.exe"
                        myProcess.StartInfo.UseShellExecute = False
                        myProcess.StartInfo.RedirectStandardInput = True
                        myProcess.StartInfo.RedirectStandardOutput = True
                        myProcess.StartInfo.CreateNoWindow = True
                        myProcess.Start()
                        myStreamWriter = myProcess.StandardInput
                        myStreamWriter.WriteLine("mklink /d " & _
                        (shadowLinkFolderName(counterVariable).ToString) _
                        & " " & (m.ToString()) & "\")
                        'MsgBox("Check if the link was created.")
                        myStreamWriter.Flush()
                        myStreamWriter.Close()
                        myProcess.WaitForExit()
                        myProcess.Close()

                    End Using

                    Dim sFileShadowPathInfo As New FileInfo(sFileShadowPath)

                    'MsgBox(sFileShadowPath)
                    'MsgBox(sFileShadowPathInfo.Exists.ToString)

                    'Here I check if the file exists in the filing system of 
                    'the shadow image to which I've just created a shortcut. 
                    'If it does not, I delete the just created shortcut.

                    If sFileShadowPathInfo.Exists = False Then

                        Using myProcess2 As New Process

                            myProcess2.StartInfo.FileName = "cmd.exe"
                            myProcess2.StartInfo.UseShellExecute = False
                            myProcess2.StartInfo.RedirectStandardInput = True
                            myProcess2.StartInfo.RedirectStandardOutput = True
                            myProcess2.StartInfo.CreateNoWindow = True
                            myProcess2.Start()
                            myStreamWriter = myProcess2.StandardInput
                            myStreamWriter.WriteLine("rmdir " & shadowLinkFolderName(counterVariable))
                            myStreamWriter.Flush()
                            myStreamWriter.Close()
                            myStreamWriter = Nothing
                            myProcess2.WaitForExit()
                            myProcess2.Close()

                        End Using

                        counterVariable = counterVariable + 1

                        Continue For

                    Else

                        'Here I compare our recovery target file against the shadow 
                        'copies. One shadow file copy is compared for each iteration 
                        'of the loop. If the string "no difference encountered is found" 
                        'then I know this shadow copy of the file is not worth looking 
                        'at, as it is the same as the recovery target. Addditonally if 
                        'the file compare error returns "FC: cannot open", then I end the
                        'match iteration of the loop to and go to the next one.

                        Using fileCompare As Process = New Process

                            fileCompare.StartInfo.FileName = "cmd.exe"
                            fileCompare.StartInfo.UseShellExecute = False
                            fileCompare.StartInfo.RedirectStandardInput = True
                            fileCompare.StartInfo.RedirectStandardOutput = True
                            fileCompare.StartInfo.CreateNoWindow = True
                            fileCompare.Start()
                            fileCompareWriter = fileCompare.StandardInput
                            fileCompareWriter.WriteLine("fc """ & sFile & """ """ _
                                            & sFileShadowPath & """")
                            fileCompareWriter.Flush()
                            fileCompareWriter.Close()
                            fileCompareReader = fileCompare.StandardOutput
                            fileCompareOut = fileCompareReader.ReadToEnd
                            fileCompareReader.Close()
                            fileCompare.WaitForExit()
                            fileCompare.Close()

                        End Using

                        fileCompareBoolean = fileCompareOut.Contains("no differences encountered").ToString
                        fileCompBooleanError = fileCompareOut.Contains("FC: cannot open").ToString

                        'MsgBox(fileCompareBoolean)
                        'MsgBox(fileCompBooleanError)

                        If fileCompBooleanError = "True" Then

                            counterVariable = counterVariable + 1

                            Continue For

                        End If

                        If fileCompareBoolean = "True" Then

                            counterVariable = counterVariable + 1

                            Continue For

                        End If

                        'Here I take a positive result of a file difference between
                        'the target and the shadow copy, and I write it out to a combo 
                        'box on the form, so it can be chosen. I also only keep the 
                        'first instance of a different shadow file as the others are 
                        'identical. I distinguish if they are the same by date and size.

                        sFileShadowPathDate = sFileShadowPathInfo.LastWriteTime
                        sFileShadowSize = sFileShadowPathInfo.Length
                        sFileShadowName = sFileShadowPathInfo.Name

                        If ComboBox1.Items.Count = 0 Then

                            ComboBox1.ResetText()

                            ComboBox1.Items.Add("Please choose from the drop down list of recovered " _
                                & "previous versions of your file.")
                            ComboBox1.Items.Add("File Name: " _
                                & sFileShadowName & " Last Modified: " & sFileShadowPathDate _
                                & " Size in Bytes: " & sFileShadowSize)
                            nonErrorShadowPathList.Add(sFileShadowPath)
                            counterVariable = counterVariable + 1

                            Dim firstComboBoxEntryIndexNumber As Integer = 0

                            Continue For

                        End If

                        previousVersionCounterVariable = ComboBox1.Items.Count - 2
                        previoussFileShadowPath = nonErrorShadowPathList(previousVersionCounterVariable)

                        Dim prevsFileShadowPathInfo As New FileInfo(previoussFileShadowPath)

                        prevsFileShadowPathSize = prevsFileShadowPathInfo.Length
                        prevsFileShadowPathDate = prevsFileShadowPathInfo.LastWriteTime

                        If String.Equals(sFileShadowPathDate, prevsFileShadowPathDate) _
                            And Long.Equals(sFileShadowSize, prevsFileShadowPathSize) Then

                            counterVariable = counterVariable + 1

                            Continue For

                        Else

                            ComboBox1.Items.Add("File Name: " _
                            & sFileShadowName & " Last Modified: " & sFileShadowPathDate _
                            & " Size in Bytes: " & sFileShadowSize)
                            nonErrorShadowPathList.Add(sFileShadowPath)
                            counterVariable = counterVariable + 1

                            Continue For

                        End If

                    End If

                Next m

                MsgBox("Processing has finished and should have returned previous " _
                       & "versions, if they exist.", MsgBoxStyle.Information)
                Try

                    If ComboBox1.Items.Count = 0 Then

                        ComboBox1.ResetText()
                        ComboBox1.Items.Add("There are no saved previous version copies " _
                                                    & "that are different from your target corrupt file.")
                        ComboBox1.SelectedItem = "There are no saved previous version copies " _
                                                    & "that are different from your target corrupt file."
                        MsgBox("There are no saved previous version copies that are different from " _
                            & "your target corrupt file. This could mean System Protection has not been " _
                            & "turned on for the drive were your file resides. Turning on System Protection " _
                            & "for your drive can be done in the System App of the Control Panel and will " _
                            & "help save future backup copies but will not allow the recovery of any present ones.")


                        Exit Sub

                    Else

                        ComboBox1.SelectedItem = "Please choose from the drop down list of recovered " _
                                & "previous versions of your file."    ' The first item has index 0 '

                    End If
                Catch ex As Exception

                End Try
            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub
    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        Try
            Dim x As Integer

            For x = 0 To shadowLinkFolderName.Count - 1

                Dim myProcess As New Process()
                myProcess.StartInfo.FileName = "cmd.exe"
                myProcess.StartInfo.UseShellExecute = False
                myProcess.StartInfo.RedirectStandardInput = True
                myProcess.StartInfo.RedirectStandardOutput = True
                myProcess.StartInfo.CreateNoWindow = True
                myProcess.Start()
                Dim myStreamWriter As StreamWriter = myProcess.StandardInput

                myStreamWriter.WriteLine("rmdir " & shadowLinkFolderName(x))
                myStreamWriter.Close()
                myProcess.WaitForExit()
                myProcess.Close()

            Next

            Me.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub Button1_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try

            Dim saveFileDialog1 As New SaveFileDialog()
            Dim pathToComboBoxSelectedFileInfo As New FileInfo(pathToComboBoxSelectedFile)
            Dim pathToComboBoxSelectedFileName As String = pathToComboBoxSelectedFileInfo.Name
            Dim pathToComboBoxSelectedFileDate As String = pathToComboBoxSelectedFileInfo.LastWriteTime
            Dim pathToComboBoxSelectedFileExtension As String = pathToComboBoxSelectedFileInfo.Extension

            saveFileDialog1.Filter = pathToComboBoxSelectedFileExtension & " Files (*" & _
                pathToComboBoxSelectedFileExtension & ")|*" & pathToComboBoxSelectedFileExtension & _
                "|All files (*.*)|*.*"
            saveFileDialog1.DefaultExt = pathToComboBoxSelectedFileExtension
            saveFileDialog1.RestoreDirectory = True

            If saveFileDialog1.ShowDialog() = DialogResult.OK Then

                saveShadowPath = saveFileDialog1.FileName

                If System.IO.File.Exists(pathToComboBoxSelectedFile) = True Then

                    System.IO.File.Copy(pathToComboBoxSelectedFile, saveShadowPath, True)
                    MsgBox(("The Previous version of " & pathToComboBoxSelectedFileName _
                                  & " last modified on " & pathToComboBoxSelectedFileDate _
                                  & " was saved to a new location: " & saveShadowPath) & ".")

                Else

                    MsgBox("Can't connect to previous version file.")

                End If

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

        Try

            'Here I assign pathToComboBoxSelectedFile to the choice made in the combo box.

            If ComboBox1.SelectedItem = "There are no saved previous version copies " _
                            & "that are different from your target corrupt file." Then

                Exit Sub

            ElseIf ComboBox1.SelectedItem = "Please choose from the drop down list of recovered " _
                            & "previous versions of your file." Then

                Exit Sub

            Else

                comboBoxChoiceIndex = ComboBox1.SelectedIndex
                pathToComboBoxSelectedFile = nonErrorShadowPathList(comboBoxChoiceIndex - 1)

            End If

            'MsgBox("ComboBox1.selectedindex: " & comboBoxChoiceIndex)
            'MsgBox(pathToComboBoxSelectedFile)

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub
End Class
