Imports System.Data.SqlClient
Imports System.IO

Public Class clsRunModels
    Public recipients As String
    Public cipher As String
    Public scriptName As String
    Public reportDir As String
    Public Sub runModel()
        Dim modelName As String = String.Empty

        Dim ServerName As String = "tigerwood"
        Dim DatabaseName As String = "MeritBaseDatamart"
        Dim DoubleQuote As String = Chr(34)
        Dim UserName As String = "JBARASH"
        Dim Password As String = "Northern#Lights!"
        Dim reportFile As String = String.Empty



        modelName = System.Text.RegularExpressions.Regex.Match(scriptName, "(?s)(?<=Models).+").ToString.Replace(".sql", "")
        reportFile = DoubleQuote & reportDir & modelName & ".rpt" & DoubleQuote
        modelName = scriptName.Replace("*.Models", "").Replace(".sql", "")
        scriptName = DoubleQuote & My.Settings.ModelDirectory & scriptName & DoubleQuote
        Dim Process = New Process()
        Process.StartInfo.UseShellExecute = False
        Process.StartInfo.RedirectStandardOutput = True
        Process.StartInfo.RedirectStandardError = True
        Process.StartInfo.CreateNoWindow = True
        Process.StartInfo.FileName = My.Settings.SQLCmd
        Process.StartInfo.Arguments =
            $"-S {ServerName} -d {DatabaseName} -i {scriptName} -U {UserName} -P {Password} -o {reportFile}"
        Process.StartInfo.WorkingDirectory = My.Settings.ModelDirectory
        Process.Start()
        'Dim output As String = Process.StandardOutput.ReadToEnd()
        'MessageBox.Show(output)
        Dim outputError As String = Process.StandardError.ReadToEnd()
        If outputError.Length > 0 Then
            Dim strFile As String = My.Settings.ReportDirectory & "ErrorLog_" & Today.ToString("yyyyMMdd") & ".txt"
            File.AppendAllText(strFile, outputError & Environment.NewLine)
        End If
        Process.WaitForExit()






    End Sub






End Class
