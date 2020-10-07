Imports System.Data.SqlClient
Imports System.IO

Module Module1
    Private logFile As String
    Sub Main()
        Dim cRunModel As clsRunModels
        Dim startTime As DateTime
        Dim endTime As DateTime
        Dim elapsedTime As TimeSpan

        Dim reportDir As String = My.Settings.ReportDirectory & Today.ToString("yyyyMMdd") & "\"
        Dim Subject As String = String.Empty, Body As String = String.Empty, modelName As String = String.Empty
        logFile = My.Settings.ReportDirectory & "reportLog_" & Today.ToString("yyyyMMdd") & ".txt"
        Try

            Dim files() As String = IO.Directory.GetFiles(My.Settings.ModelDirectory, "*.sql")

            For Each mfile As String In files
                cRunModel = New clsRunModels
                cRunModel.reportDir = reportDir
                cRunModel.scriptName = mfile
                writeFile("Start " & mfile)
                startTime = Now
                cRunModel.runModel()
                endTime = Now
                writeFile("End " & mfile)
                elapsedTime = endTime.Subtract(startTime)
                writeFile(String.Format("{3} {0}hr : {1}min : {2}sec", elapsedTime.Hours, elapsedTime.Minutes, elapsedTime.Seconds, mfile))
                modelName = System.Text.RegularExpressions.Regex.Match(mfile, "(?s)(?<=Models).+").ToString.Replace(".sql", "")
                If System.IO.File.Exists(reportDir & modelName & ".rpt") Then
                    Body = My.Computer.FileSystem.ReadAllText(reportDir & modelName & ".rpt")
                    If Body.Contains("Msg") Then
                        Subject = "MB Model Failed "
                    Else
                        Subject = "MB Model Completed "

                    End If
                    emailReportAttach(Subject & modelName, Body, My.Settings.recipients, My.Settings.ReportDirectory & modelName & ".rpt")
                Else
                    Subject = "Report File Not Found " & modelName
                    emailReport(Subject, "Report File Not Found", My.Settings.recipients)
                End If


                cRunModel = Nothing
            Next

            emailReport("MB Models Completed", "Models finished", My.Settings.recipients)
        Catch ex As Exception

            emailReport("Error Running MB Models", ex.ToString, My.Settings.recipients)
        End Try
    End Sub
    Private Sub writeFile(message As String)
        Dim strMessage As String = Now.ToString("yyyyMMdd hh:mm:ss") & " " & message
        Using outStream As StreamWriter = File.AppendText(logFile)
            outStream.WriteLine(strMessage)
        End Using
    End Sub
    Private Sub writeFile(message As String, fileName As String)
        Dim strMessage As String = Now.ToString("yyyyMMdd hh:mm:ss") & " " & message
        Using outStream As StreamWriter = File.AppendText(fileName)
            outStream.WriteLine(strMessage)
        End Using
    End Sub
    Public Sub emailReport(subject As String, message As String, recipients As String)
        Dim connectionString As String = "Server=tigerwood;Initial Catalog=Master;User Id=JBARASH;Password=Northern#Lights!"
        Using conn = New SqlClient.SqlConnection(connectionString)
            conn.Open()

            Using cmd As SqlCommand = New SqlCommand("msdb.dbo.sp_send_dbmail", conn)
                cmd.CommandType = CommandType.StoredProcedure

                cmd.Parameters.AddWithValue("@recipients", recipients)
                cmd.Parameters.AddWithValue("Subject", subject)
                cmd.Parameters.AddWithValue("@body", message)
                cmd.Parameters.AddWithValue("@body_format", "HTML")
                cmd.Parameters.AddWithValue("@profile_name", "Auto")
                'cmd.Parameters.AddWithValue("@file_attachments", fileName)
                Try
                    cmd.ExecuteNonQuery()
                Catch exs As SqlException
                    writeFile(exs.ToString, My.Settings.ReportDirectory & "ErrorLog_" & Today.ToString("yyyyMMdd") & ".txt")
                Catch ex As Exception
                    writeFile(ex.ToString, My.Settings.ReportDirectory & "ErrorLog_" & Today.ToString("yyyyMMdd") & ".txt")
                End Try
            End Using
        End Using

    End Sub
    Public Sub emailReportAttach(subject As String, message As String, recipients As String, fileName As String)
        Dim connectionString As String = "Server=tigerwood;Initial Catalog=Master;User Id=JBARASH;Password=Northern#Lights!"
        Using conn = New SqlClient.SqlConnection(connectionString)
            conn.Open()

            Using cmd As SqlCommand = New SqlCommand("msdb.dbo.sp_send_dbmail", conn)
                cmd.CommandType = CommandType.StoredProcedure

                cmd.Parameters.AddWithValue("@recipients", recipients)
                cmd.Parameters.AddWithValue("Subject", subject)
                cmd.Parameters.AddWithValue("@body", message)
                cmd.Parameters.AddWithValue("@body_format", "TEXT")
                cmd.Parameters.AddWithValue("@profile_name", "Auto")
                cmd.Parameters.AddWithValue("@file_attachments", fileName)
                Try
                    cmd.ExecuteNonQuery()
                Catch exs As SqlException
                    writeFile(exs.ToString, My.Settings.ReportDirectory & "ErrorLog_{DateTime.Today:dd-MMM-yyyy}.txt")
                Catch ex As Exception
                    writeFile(ex.ToString, My.Settings.ReportDirectory & "ErrorLog_{DateTime.Today:dd-MMM-yyyy}.txt")
                End Try
            End Using
        End Using

    End Sub

End Module
