Imports System.IO
Imports System.Threading
Imports System.ComponentModel
Imports System.Text
Imports System.Net.Mail



Public Class Main_Screen

    Private busyworking As Boolean = False

    Private lastinputline As String = ""
    Private inputlines As Long = 0
    Private highestPercentageReached As Integer = 0
    Private inputlinesprecount As Long = 0
    Private pretestdone As Boolean = False
    Private primary_PercentComplete As Integer = 0
    Private percentComplete As Integer

    Private inputfile As String = ""
    Private outputfile As String = ""
    Private emailaddress As String = ""
    Private mailserver1 As String = ""
    Private mailserver1port As String = ""
    Private mailserver2 As String = ""
    Private mailserver2port As String = ""
    Private webmasteraddress As String = ""
    Private webmasterdisplay As String = ""
    Private webroot As String = ""
    Private webroottranslate As String = ""

    ' Constants
    Const WdDoNotSaveChanges = 0

    ' see WdSaveFormat enumeration constants: 
    ' http://msdn2.microsoft.com/en-us/library/bb238158.aspx
    Const wdFormatPDF = 17  ' PDF format. 
    Const wdFormatXPS = 18  ' XPS format. 


    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "", Optional ByVal ExitApplication As Boolean = True)
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message()
                If ExitApplication = True Then
                    Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ": " & ex.Message.ToString & vbCrLf & vbCrLf & "Please note due to the nature of this application, this application will now attempt to close itself."
                Else
                    Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ": " & ex.Message.ToString & vbCrLf & vbCrLf & ""
                End If


                Display_Message1.Timer1.Interval = 1000
                Display_Message1.ShowDialog()
                Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                dir = Nothing
                Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ": " & ex.ToString)
                filewriter.WriteLine("")
                filewriter.Flush()
                filewriter.Close()
                filewriter = Nothing
            End If
            ex = Nothing
            identifier_msg = Nothing
            If ExitApplication = True Then
                Me.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in the application's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub



    Private Sub RunWorker()
        Try
            If busyworking = False Then
                If My.Computer.FileSystem.FileExists(TextBox1.Text) Then
                    busyworking = True
                    inputlines = 0
                    lastinputline = ""
                    highestPercentageReached = 0
                    inputlinesprecount = 0
                    pretestdone = False
                    TextBox1.Enabled = False
                    TextBox2.Enabled = False
                    TextBox3.Enabled = False
                    BackgroundWorker1.RunWorkerAsync(TextBox1.Text)
                Else
                    Dim Display_Message1 As New Display_Message()
                    Display_Message1.Message_Textbox.Text = "Specified input file cannot be located. This application will now shutdown."
                    Display_Message1.Timer1.Interval = 1000
                    Display_Message1.ShowDialog()
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex, "StartWorker")
        End Try
    End Sub

    ' This event handler is where the actual work is done.
    Private Sub backgroundWorker1_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        ' Get the BackgroundWorker object that raised this event.
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)

        ' Assign the result of the computation
        ' to the Result property of the DoWorkEventArgs
        ' object. This is will be available to the 
        ' RunWorkerCompleted eventhandler.
        e.Result = MainWorkerFunction(worker, e)
        sender = Nothing
        e = Nothing
        worker.Dispose()
        worker = Nothing
    End Sub 'backgroundWorker1_DoWork

    ' This event handler deals with the results of the
    ' background operation.
    Private Sub backgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            busyworking = False
            ' First, handle the case where an exception was thrown.
            If Not (e.Error Is Nothing) Then
                Error_Handler(e.Error, "backgroundWorker1_RunWorkerCompleted")
            ElseIf e.Cancelled Then
                ' Next, handle the case where the user canceled the 
                ' operation.
                ' Note that due to a race condition in 
                ' the DoWork event handler, the Cancelled
                ' flag may not have been set, even though
                ' CancelAsync was called.
                'Me.ToolStripStatusLabel1.Text = "Operation Cancelled" & "   (" & inputlines & " of " & inputlinesprecount & ")"
                Me.ToolStripStatusLabel1.Text = "Operation Cancelled"
                Me.ProgressBar1.Value = 0

            Else
                ' Finally, handle the case where the operation succeeded.
                'Me.ToolStripStatusLabel1.Text = "Operation Completed" & "   (" & inputlines & " of " & inputlinesprecount & ")"
                Me.ToolStripStatusLabel1.Text = "Operation Completed"
                Me.ProgressBar1.Value = 100
            End If

            TextBox1.Enabled = True
            TextBox2.Enabled = True
            TextBox3.Enabled = True

            sender = Nothing
            e = Nothing
            Me.Close()
        Catch ex As Exception
            Error_Handler(ex, "Operation Completed")
        End Try
    End Sub 'backgroundWorker1_RunWorkerCompleted

    Private Sub backgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Try
            Me.ProgressBar1.Value = e.ProgressPercentage
            'If lastinputline.StartsWith("Operation Completed") Then
            'Me.ToolStripStatusLabel1.Text = lastinputline
            'Else
            'Me.ToolStripStatusLabel1.Text = lastinputline & "   (" & inputlines & " of " & inputlinesprecount & ")"
            Me.ToolStripStatusLabel1.Text = lastinputline
            'End If
            sender = Nothing
            e = Nothing
        Catch ex As Exception
            Error_Handler(ex, "Progess Changed")
        End Try
    End Sub

    Function MainWorkerFunction(ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs) As Boolean
        Dim result As Boolean = False
        Try

            lastinputline = "Initialising Operation"
            primary_PercentComplete = 0
            worker.ReportProgress(0)
            inputlinesprecount = 2
            Me.pretestdone = True

            If worker.CancellationPending Then
                e.Cancel = True
                Return False
            End If

            primary_PercentComplete = 0
            worker.ReportProgress(0)

            inputlines = 0
            lastinputline = ""
            Dim filename As String = ""
            If My.Computer.FileSystem.FileExists(TextBox1.Text) = True Then
                Dim finfo11 As FileInfo = New FileInfo(TextBox1.Text)
                filename = finfo11.Name
                finfo11 = Nothing
                If My.Computer.FileSystem.FileExists(TextBox2.Text) = False Then
                    lastinputline = "Running Converter"
                    worker.ReportProgress(33)
                    Try


                        Dim wdocs As Microsoft.Office.Interop.Word.Documents
                        Dim wdoc As Microsoft.Office.Interop.Word.Document
                        Dim wdo As Microsoft.Office.Interop.Word.Application = New Microsoft.Office.Interop.Word.Application
                        wdocs = wdo.Documents
                        wdoc = wdocs.Open(TextBox1.Text)
                        wdoc.SaveAs(TextBox2.Text, wdFormatPDF)

                        wdoc.Close(WdDoNotSaveChanges)
                        wdo.Quit(WdDoNotSaveChanges)
                        wdoc = Nothing
                        wdoc = Nothing
                        wdo = Nothing
                    Catch ex As Exception
                        Try
                            Dim wdo As Microsoft.Office.Interop.Word.Application = New Microsoft.Office.Interop.Word.Application
                            Dim wdoc As Microsoft.Office.Interop.Word.Document
                            Dim wdocs As Microsoft.Office.Interop.Word.Documents
                            Dim sPrevPrinter As String
                            Dim oDistiller As ACRODISTXLib.PdfDistiller = New ACRODISTXLib.PdfDistiller()
                            wdocs = wdo.Documents
                            sPrevPrinter = wdo.ActivePrinter
                            wdo.ActivePrinter = "Acrobat Distiller"
                            wdoc = wdocs.Open(TextBox1.Text)
                            wdo.ActiveDocument.PrintOut(False, , , TextBox2.Text & ".ps")
                            wdoc.Close(WdDoNotSaveChanges)
                            wdo.ActivePrinter = sPrevPrinter
                            wdo.Quit()
                            wdo = Nothing
                            oDistiller.FileToPDF(TextBox2.Text & ".ps", TextBox2.Text, "Print")
                            oDistiller = Nothing
                            My.Computer.FileSystem.DeleteFile(TextBox2.Text & ".ps")
                            wdoc = Nothing
                            wdocs = Nothing
                        Catch ex1 As Exception
                            Error_Handler(ex1, "Running Converter", False)
                        End Try
                    End Try
                    lastinputline = "Sending Notification Mail Out"
                    worker.ReportProgress(66)

                    If My.Computer.FileSystem.FileExists(TextBox2.Text) = True Then
                        '*********************
                        'Send Mail Out - file was converted
                        Try
                            Dim obj As SmtpClient
                            If mailserver1port.Length > 0 Then
                                obj = New SmtpClient(mailserver1, mailserver1port)
                            Else
                                obj = New SmtpClient(mailserver1)
                            End If

                            Dim msg As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage

                            msg.Subject = "PDF Generator Notification"
                            Dim fromaddress As MailAddress = New MailAddress(webmasteraddress, webmasterdisplay)
                            msg.From = fromaddress
                            msg.ReplyTo = fromaddress
                            msg.To.Add(TextBox3.Text)

                            msg.IsBodyHtml = False

                            obj.DeliveryMethod = SmtpDeliveryMethod.Network
                            obj.EnableSsl = False
                            obj.UseDefaultCredentials = True

                            msg.Body = "This is a notification message from PDF Generator running on " & webroottranslate & " to inform you that your document has been successfully converted to PDF format."
                            msg.Body = msg.Body & vbCrLf & vbCrLf & "You can go ahead and download the file from: " & vbCrLf & vbTab & Uri.EscapeUriString(TextBox2.Text.Replace(webroot, webroottranslate).Replace("\", "/"))
                            msg.Body = msg.Body & vbCrLf & vbCrLf & "Please remember that this file will be available for the next 48 hours only and will be removed from the server at " & Format(Now.AddHours(48), "HH:mm:ss dd/MM/yyyy")
                            obj.Send(msg)
                            obj = Nothing
                        Catch ex As Exception
                            Try
                                Dim obj As SmtpClient
                                If mailserver2port.Length > 0 Then
                                    obj = New SmtpClient(mailserver2, mailserver2port)
                                Else
                                    obj = New SmtpClient(mailserver2)
                                End If
                                Dim msg As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage

                                msg.Subject = "PDF Generator Notification"
                                Dim fromaddress As MailAddress = New MailAddress(webmasteraddress, webmasterdisplay)
                                msg.From = fromaddress
                                msg.ReplyTo = fromaddress
                                msg.To.Add(TextBox3.Text)

                                msg.IsBodyHtml = False

                                obj.DeliveryMethod = SmtpDeliveryMethod.Network
                                obj.EnableSsl = False
                                obj.UseDefaultCredentials = True
                                msg.Body = "This is a notification message from PDF Generator running on " & webroottranslate & " to inform you that your document has been successfully converted to PDF format."
                                msg.Body = msg.Body & vbCrLf & vbCrLf & "You can go ahead and download the file from: " & vbCrLf & vbTab & Uri.EscapeUriString(TextBox2.Text.Replace(webroot, webroottranslate).Replace("\", "/"))
                                msg.Body = msg.Body & vbCrLf & vbCrLf & "Please remember that this file will be available for the next 48 hours only and will be removed from the server at " & Format(Now.AddHours(48), "HH:mm:ss dd/MM/yyyy")
                                obj.Send(msg)
                                obj = Nothing
                            Catch ex1 As Exception
                                Error_Handler(ex, "Send Mail")
                            End Try
                        End Try
                    Else
                        '*********************
                        'Send Mail Out - file wasn't converted
                        Try
                            Dim obj As SmtpClient
                            If mailserver1port.Length > 0 Then
                                obj = New SmtpClient(mailserver1, mailserver1port)
                            Else
                                obj = New SmtpClient(mailserver1)
                            End If

                            Dim msg As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage

                            msg.Subject = "PDF Generator Notification"
                            Dim fromaddress As MailAddress = New MailAddress(webmasteraddress, webmasterdisplay)
                            msg.From = fromaddress
                            msg.ReplyTo = fromaddress
                            msg.To.Add(TextBox3.Text)

                            msg.IsBodyHtml = False

                            obj.DeliveryMethod = SmtpDeliveryMethod.Network
                            obj.EnableSsl = False
                            obj.UseDefaultCredentials = True
                            msg.Body = "This is a notification message from PDF Generator running on " & webroottranslate & " to inform you that there was a problem in converting your submitted file (" & filename & "). It is suggested that you try submitting a simple text file, just to ensure that this service is in fact operating correctly."
                            obj.Send(msg)
                            obj = Nothing
                        Catch ex As Exception
                            Try
                                Dim obj As SmtpClient
                                If mailserver2port.Length > 0 Then
                                    obj = New SmtpClient(mailserver2, mailserver2port)
                                Else
                                    obj = New SmtpClient(mailserver2)
                                End If
                                Dim msg As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage

                                msg.Subject = "PDF Generator Notification"
                                Dim fromaddress As MailAddress = New MailAddress(webmasteraddress, webmasterdisplay)
                                msg.From = fromaddress
                                msg.ReplyTo = fromaddress
                                msg.To.Add(TextBox3.Text)

                                msg.IsBodyHtml = False

                                obj.DeliveryMethod = SmtpDeliveryMethod.Network
                                obj.EnableSsl = False
                                obj.UseDefaultCredentials = True
                                msg.Body = "This is a notification message from PDF Generator running on " & webroottranslate & " to inform you that there was a problem in converting your submitted file (" & filename & "). It is suggested that you try submitting a simple text file, just to ensure that this service is in fact operating correctly."

                                obj.Send(msg)
                                obj = Nothing
                            Catch ex1 As Exception
                                Error_Handler(ex, "Send Mail")
                            End Try
                        End Try
                    End If

                Else
                    e.Cancel = True
                    lastinputline = "File to create already exists on the system"
                    Dim Display_Message1 As New Display_Message()
                    Display_Message1.Message_Textbox.Text = "File to create already exists on the system. This application will now shut itself down."
                    Display_Message1.Timer1.Interval = 1000
                    Display_Message1.ShowDialog()
                    Me.Close()
                End If
            Else
                e.Cancel = True
                lastinputline = "File to convert doesn't seem to exist on the system"
                Dim Display_Message1 As New Display_Message()
                Display_Message1.Message_Textbox.Text = "File to convert doesn't seem to exist on the system. This application will now shut itself down."
                Display_Message1.Timer1.Interval = 1000
                Display_Message1.ShowDialog()
                Me.Close()
            End If

            result = True

        Catch ex As Exception
            Error_Handler(ex, "Main Worker Function")
        End Try
        worker.Dispose()
        worker = Nothing
        e = Nothing
        Return result

    End Function


    Private Sub LoadSettings()
        Try
            Dim configfile As String = (Application.StartupPath & "\config.sav").Replace("\\", "\")
            If My.Computer.FileSystem.FileExists(configfile) Then
                Dim reader As StreamReader = New StreamReader(configfile)
                Dim lineread As String
                Dim variablevalue As String
                While reader.Peek <> -1
                    lineread = reader.ReadLine
                    If lineread.IndexOf("=") <> -1 Then

                        variablevalue = lineread.Remove(0, lineread.IndexOf("=") + 1)

                        If lineread.StartsWith("mailserver1=") Then
                            mailserver1 = variablevalue
                            If mailserver1.Length < 1 Then
                                mailserver1 = "mail.uct.ac.za"
                            End If
                        End If
                        If lineread.StartsWith("mailserver1port=") Then
                            mailserver1port = variablevalue
                            If mailserver1port.Length < 1 Then
                                mailserver1port = "25"
                            End If
                        End If
                        If lineread.StartsWith("mailserver2=") Then
                            mailserver2 = variablevalue
                            If mailserver2.Length < 1 Then
                                mailserver2 = "obe1.com.uct.ac.za"
                            End If
                        End If
                        If lineread.StartsWith("mailserver2port=") Then
                            mailserver2port = variablevalue
                            If mailserver2port.Length < 1 Then
                                mailserver2port = "25"
                            End If
                        End If
                        If lineread.StartsWith("webmasteraddress=") Then
                            webmasteraddress = variablevalue
                            If webmasteraddress.Length < 1 Then
                                webmasteraddress = "com-webmaster@uct.ac.za"
                            End If
                        End If
                        If lineread.StartsWith("webmasterdisplay=") Then
                            webmasterdisplay = variablevalue
                            If webmasterdisplay.Length < 1 Then
                                webmasterdisplay = "Commerce Webmaster"
                            End If
                        End If
                        If lineread.StartsWith("webroot=") Then
                            webroot = variablevalue
                            If webroot.Length < 1 Then
                                webroot = "C:\Inetpub\wwwroot"
                            End If
                        End If
                        If lineread.StartsWith("webroottranslate=") Then
                            webroottranslate = variablevalue
                            If webroottranslate.Length < 1 Then
                                webroottranslate = "http://www.commerce.uct.ac.za"
                            End If
                        End If
                    End If
                End While
                reader.Close()
                reader = Nothing
            End If
        Catch ex As Exception
            Error_Handler(ex, "Load Settings")
        End Try
    End Sub

    Private Sub SaveSettings()
        Try
            Dim configfile As String = (Application.StartupPath & "\config.sav").Replace("\\", "\")
            Dim writer As StreamWriter = New StreamWriter(configfile, False)
            writer.WriteLine("mailserver1=" & mailserver1)
            writer.WriteLine("mailserver1port=" & mailserver1port)
            writer.WriteLine("mailserver2=" & mailserver2)
            writer.WriteLine("mailserver2port=" & mailserver2port)
            writer.WriteLine("webmasteraddress=" & webmasteraddress)
            writer.WriteLine("webmasterdisplay=" & webmasterdisplay)
            writer.WriteLine("webroot=" & webroot)
            writer.WriteLine("webroottranslate=" & webroottranslate)

            writer.Flush()
            writer.Close()
            writer = Nothing
        Catch ex As Exception
            Error_Handler(ex, "Save Settings")
        End Try
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            Me.Text = My.Application.Info.ProductName & " " & Format(My.Application.Info.Version.Major, "0000") & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & Format(My.Application.Info.Version.Revision, "00") & ""
            LoadSettings()
            Me.ToolStripStatusLabel1.Text = "Application Loaded"
            If My.Application.CommandLineArgs.Count < 3 Then
                Dim Display_Message1 As New Display_Message()
                Display_Message1.Message_Textbox.Text = "Incorrect number of command line arguments detected. The usage pattern for this application is:" & vbCrLf & """PDF Generator DOC2PDF.exe"" #inputfile.doc# #outputfile.pdf# #myemail@address.com#." & vbCrLf & vbCrLf & "This application will now shut itself down."
                Display_Message1.Timer1.Interval = 1000
                Display_Message1.ShowDialog()
                Me.Close()
            Else
                TextBox1.Text = My.Application.CommandLineArgs(0)
                TextBox2.Text = My.Application.CommandLineArgs(1)
                TextBox3.Text = My.Application.CommandLineArgs(2)
                RunWorker()
            End If

        Catch ex As Exception
            Error_Handler(ex, "Application Load")
        End Try
    End Sub

    Private Sub Form1_Close(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            Me.ToolStripStatusLabel1.Text = "Application Closing"
            SaveSettings()
        Catch ex As Exception
            Error_Handler(ex, "Application Close")
        End Try
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Try
            Me.ToolStripStatusLabel1.Text = "About displayed"
            AboutBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display About Screen")
        End Try
    End Sub

    Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem.Click
        Try
            Me.ToolStripStatusLabel1.Text = "Help displayed"
            HelpBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display Help Screen")
        End Try
    End Sub


End Class
