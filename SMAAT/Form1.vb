Imports System.Net
Imports System.Net.Mail

Public Class SMAAT
    Public element1 As String = "<!DOCTYPE html>
<html lang='en' xmlns='http://www.w3.org/1999/xhtml'>
<head>
<meta charset='utf-8' />
<title>MAIL</title>
<script type='text/java script'></script>
</head>
<body>
<h1 id='Mailh'></h1>
<h1 id='Fromh'></h1>
<p id='Fromp'></p>
<h1 id='Subjecth'></h1>
<p id='Subjectp'></p>
<h1 id='Bodyh'></h1>
<p id='Bodyp'></p>
<script>
var chara =0;
//MAILh
function timedText0() 
{

var Mailh = 'MAIL';
var at0 = 0;
var interval = setInterval(animateMailh,50);
function animateMailh()
{
	if (at0 < Mailh.length)
	{
		var para = document.getElementById('Mailh');
		para.innerHTML = para.innerHTML + Mailh[at0];
		at0++;
		}
		else
		{
			clearInterval(interval);
			timedText2();
			}
			}
}
timedText0();

//Fromh
function timedText2() 
{

var Fromh = 'From:';
var at0 = 0;
var interval = setInterval(animateFromh,50);
function animateFromh()
{
	if (at0 < Fromh.length)
	{
		var para = document.getElementById('Fromh');
		para.innerHTML = para.innerHTML + Fromh[at0];
		at0++;
		}
		else
		{
			clearInterval(interval);
			timedText3();
			}
			}
}

//Fromp

function timedText3() 
{

var Fromp = '"


    Public element2 As String = "';
var at0 = 0;
var interval = setInterval(animateFromp,50);
function animateFromp()
{
	if (at0 < Fromp.length)
	{
		var para = document.getElementById('Fromp');
		para.innerHTML = para.innerHTML + Fromp[at0];
		at0++;
		}
		else
		{
			clearInterval(interval);
			timedText4();
			}
			}
}

//Subjecth

function timedText4() 
{

var Subjecth = 'Subject:';
var at0 = 0;
var interval = setInterval(animateSubjecth,50);
function animateSubjecth()
{
	if (at0 < Subjecth.length)
	{
		var para = document.getElementById('Subjecth');
		para.innerHTML = para.innerHTML + Subjecth[at0];
		at0++;
		}
		else
		{
			clearInterval(interval);
			timedText5();
			}
			}
}

//Subjectp

function timedText5() 
{

var Subjectp = '"


    Public element3 As String = "';
var at0 = 0;
var interval = setInterval(animateSubjectp,50);
function animateSubjectp()
{
	if (at0 < Subjectp.length)
	{
		var para = document.getElementById('Subjectp');
		para.innerHTML = para.innerHTML + Subjectp[at0];
		at0++;
		}
		else
		{
			clearInterval(interval);
			timedText6();
			}
			}
}

//Bodyh

function timedText6() 
{
var Bodyh = 'Body:'
var at = 0;
var interval = setInterval(animateBodyh,50);
function animateBodyh()
{
	if (at < Bodyh.length)
	{
		var para = document.getElementById('Bodyh');
		para.innerHTML = para.innerHTML + Bodyh[at];
		at++;
		}
		else
		{
			clearInterval(interval);
			timedText1();
			}
			}
			}



//Bodyp
function timedText1() 
{
var Bodyp ="

    Public element4 As String = "
var at = 0;
var interval = setInterval(animatepara,50);
function animatepara()
{
	if (at < Bodyp.length)
	{
		var para = document.getElementById('Bodyp');
		para.innerHTML = para.innerHTML + Bodyp[at];
		at++;
		}
		else
		{
			clearInterval(interval);
			}
			}
			}
			</script>
			</body>
			</html>"


    Private Sub SMAAT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Me.Icon = My.Resources._20281
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub



    Private Sub CloseMe_Click(sender As Object, e As EventArgs) Handles CloseMe.Click
        Try
            Dim p As Process = Process.GetCurrentProcess
            p.Kill()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub CloseMe1_Click(sender As Object, e As EventArgs) Handles CloseMe1.Click
        Try
            Dim p As Process = Process.GetCurrentProcess
            p.Kill()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Minimize_Click(sender As Object, e As EventArgs) Handles Minimize.Click
        Try
            Me.WindowState = FormWindowState.Minimized
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Minimize1_Click(sender As Object, e As EventArgs) Handles Minimize1.Click
        Try
            Me.WindowState = FormWindowState.Minimized
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Function MakeHTML() As String
        Dim returntext As String = ""
        If MailB.Text.Length = 0 Then
            returntext = element1 + FromEmail.Text + element2 + Subject.Text + element3 + " ' " + "Empty ';" + element4
        Else
            returntext = element1 + FromEmail.Text + element2 + Subject.Text + element3 + Dividethebody() + " ;" + element4
        End If
        Return returntext
    End Function
    Public Function Dividethebody() As String
        Dim returntext As String = ""
        Dim line As String = ""
        For i = 0 To (MailB.Text.Length - 1)
            If line.Length = 50 Then
                returntext = returntext + " '" + line + "' + "
                line = ""
            ElseIf i = (MailB.Text.Length - 1) Then
                line = line + MailB.Text(i)
                returntext = returntext + " '" + line + "' + "
                line = ""
            End If
            line = line + MailB.Text(i)
        Next
        returntext = returntext + "' '"
        Return returntext
    End Function
    Public Sub SendMail()
        Try
            Dim Mail As MailMessage = New MailMessage()
            Dim SmtpServer As SmtpClient = New SmtpClient("smtp.gmail.com")
            Mail.From = New MailAddress(FromEmail.Text)
            Mail.To.Add(Toaddress.Text)
            Mail.Subject = Subject.Text
            Mail.Body = ""
            SmtpServer.Port = 587
            SmtpServer.Credentials = New System.Net.NetworkCredential(FromEmail.Text, Password.Text)
            SmtpServer.EnableSsl = True
            Dim body As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(MakeHTML)
            Dim stream As New IO.MemoryStream(body)
            Dim myattachment As Attachment = New Attachment(stream, "Mail.html", Nothing)
            Mail.Attachments.Add(myattachment)
1:          Dim result = MsgBox("Do you want to attach something else  with this mail?", MsgBoxStyle.OkCancel)
            If result = MsgBoxResult.Ok Then
                Dim result1 = OpenFileDialog1.ShowDialog()
                If result1 = DialogResult.OK Then
                    For Each file In OpenFileDialog1.FileNames
                        Dim myattachment1 As Attachment = New Attachment(file)
                        Mail.Attachments.Add(myattachment1)
                    Next
                    GoTo 1
                End If
            End If
            SmtpServer.Send(Mail)
            MsgBox("Mail sent")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Send_Click(sender As Object, e As EventArgs) Handles Send.Click
        SendMail()
    End Sub

    Private Sub Send1_Click(sender As Object, e As EventArgs) Handles Send1.Click
        SendMail()
    End Sub
End Class
