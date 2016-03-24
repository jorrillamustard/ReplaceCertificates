Imports System.IO
Imports System.ServiceProcess
Imports System.Windows.Forms
Imports System.IO.Compression

Module pubss


    Dim SSPath As String = "C:\ProgramData\Resolution1\SiteServer\siteserver.config"

    Sub Main()
        Try

            Dim externalIP As String = getMYIP()
            Dim ParentIP As String = getParentIP()
            Dim privateCert As String = getPrivateCert()
            If File.Exists(SSPath + ".original") Then

            Else
            makeBackup(SSPath)
            End If
            Console.WriteLine("Configuring Public SS...")
            SSconfigExternalIP(externalIP)
            SSconfigParentIP(ParentIP)
            SSconfigPrivate(privateCert)
            Console.WriteLine("Restarting Site Server...")
            restartService()
            Console.WriteLine("Cleaning Public Site Server...")
            cleanup()
        Catch ex As Exception
            Console.WriteLine(ex)
        End Try
    End Sub

    Function makeBackup(path As String)
        File.Copy(path, path + ".original")
        Return True
    End Function

    Function getParentIP()
        Dim ParentIP As String = My.Application.CommandLineArgs(1)
        Return ParentIP
    End Function

    Function getMYIP()
        Dim myIP As String = My.Application.CommandLineArgs(0)
        Return myIP
    End Function

    Function getPrivateCert()
        Dim cert As String = My.Application.CommandLineArgs(2)
        Return cert
    End Function

    Private Sub SSconfigPrivate(cert As String)
        Dim text As String = IO.File.ReadAllText(SSPath)
        Dim find As String = "PrivateCert="""
        Dim xst As Integer = text.IndexOf(find) + find.Length
        Dim xend As Integer = text.IndexOf("Results") - 3
        Dim xsub As String = text.Substring(xst, (xend - xst) + 1)
        text = text.Replace(xsub, cert)
        Dim Writer As System.IO.StreamWriter
        Writer = New System.IO.StreamWriter(SSPath) '<-- Where to write to
        Writer.Write(text)
        Writer.Close()
    End Sub


    Private Sub SSconfigExternalIP(pubIP As String)
        Dim text As String = IO.File.ReadAllText(SSPath)
        Dim find As String = "ExternalAddress="""
        Dim xst As Integer = text.IndexOf(find) + find.Length
        Dim xend As Integer = text.IndexOf("InternalAddress") - 3
        Dim xsub As String = text.Substring(xst, (xend - xst) + 1)
        text = text.Replace(xsub, pubIP)
        Dim Writer As System.IO.StreamWriter
        Writer = New System.IO.StreamWriter(SSPath) '<-- Where to write to
        Writer.Write(text)
        Writer.Close()
    End Sub

    Private Sub SSconfigParentIP(pubIP As String)
        Dim text As String = IO.File.ReadAllText(SSPath)
        Dim find As String = "Parent="""
        ' Console.WriteLine(find)
        Dim xst As Integer = text.IndexOf(find) + find.Length
        ' Console.WriteLine(xst)
        Dim xend As Integer = text.IndexOf("ExternalAddress") - 3
        ' Console.WriteLine(xend)
        Dim xsub As String = text.Substring(xst, (xend - xst) + 1)
        text = text.Replace(xsub, pubIP + ":54545")
        Dim Writer As System.IO.StreamWriter
        Writer = New System.IO.StreamWriter(SSPath) '<-- Where to write to
        Writer.Write(text)
        Writer.Close()
    End Sub

    Function restartService()
        Dim sc2 As New ServiceController()
        sc2.ServiceName = "SiteServer"

        Try

            sc2.Stop()
            Console.WriteLine("Stopping Site Server Service...")
            sc2.WaitForStatus(ServiceControllerStatus.Stopped)


            sc2.Start()
            Console.WriteLine("Starting Site Server Service...")
            sc2.WaitForStatus(ServiceControllerStatus.Running)

            Return True
        Catch ex As Exception
            Console.WriteLine(ex)
            Return False
        End Try
    End Function

    Function cleanup()
        File.Delete("C:\unzip.ps1")
        File.Delete("C:\pubss.zip")
    End Function


End Module
