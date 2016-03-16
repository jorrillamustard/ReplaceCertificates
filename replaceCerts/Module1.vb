Imports System.IO
Imports System.ServiceProcess
Imports System.Configuration
Imports System.Windows.Forms
Imports System.IO.Compression

Module module1
    Dim publicCert As String = ""
    Dim privateCert As String = ""
    Dim oldPub As String = ""
    Dim oldPrivate As String = ""
    Dim help As String = "Format: replacecerts.exe <-n,-d> <New Public Cert Name> <New Private Cert Name>"

    Dim cAppConfig As Configuration = ConfigurationManager.OpenExeConfiguration(My.Application.Info.DirectoryPath + "\replaceCerts.exe")
    Dim asSettings As AppSettingsSection = cAppConfig.AppSettings

    Dim cAppConfig1 As Configuration = ConfigurationManager.OpenExeConfiguration(asSettings.Settings.Item("DriveLetter").Value + ":\Program Files\Resolution1\Work Manager\Infrastructure.WorkExecutionServices.Host.exe")
    Dim asSettings1 As AppSettingsSection = cAppConfig1.AppSettings


    Dim WMPath As String = asSettings.Settings.Item("DriveLetter").Value + ":\Program Files\Resolution1\Work Manager\Infrastructure.WorkExecutionServices.Host.exe.config"
    Dim SSPath As String = "C:\ProgramData\Resolution1\SiteServer\siteserver.config"

    'main function for application
    Sub Main()


        If checkOSSLinstalled() = True Then
            GoTo start
        Else
            OSSLDownload()
        End If
start:

        Dim flag As Boolean = parse()

        If flag = False Then
            GoTo closeme
        Else

            If File.Exists(WMPath + ".original") And File.Exists(SSPath + ".original") Then
                GoTo Replace
            Else
                Try
                    File.Copy(WMPath, WMPath + ".original")
                    File.Copy(SSPath, SSPath + ".original")
                Catch ex As Exception
                    Console.WriteLine(ex)
                End Try
            End If

Replace:
            FindReplaceString(WMPath, oldPub, publicCert)
            FindReplaceString(WMPath, oldPrivate, privateCert + ".adp12")
            FindReplaceString(SSPath, oldPub, publicCert)
            FindReplaceString(SSPath, oldPrivate, privateCert)

            restartService()
            SSClick()
            cleanup()
        End If
closeme:
    End Sub


    Function checkOSSLinstalled()
        If Directory.Exists(asSettings.Settings.Item("OSSLDirectory").Value + "\bin") Then
            If File.Exists(asSettings.Settings.Item("OSSLDirectory").Value + "\bin\openssl.cnf") Then
                Return True
            Else
                File.Copy(asSettings.Settings.Item("OSSLDirectory").Value + "\share\openssl.cnf", asSettings.Settings.Item("OSSLDirectory").Value + "\bin\openssl.cnf")
                Return True
            End If
        Else
                Return False
        End If
    End Function


    Function OSSLDownload()
        Dim nl As String = Environment.NewLine
        Dim input As String
        If checkOSSLinstalled() = True Then
            Console.WriteLine("OpenSSL appears to be installed. Would you like to install anyway?" + nl)
            input = Console.ReadLine()
            If input.ToLower.Equals("y") Then
                Try
                    Console.WriteLine("Downloading OpenSSL..." + nl)
                    DownloadFile("https://docs.google.com/uc?authuser=0&id=0B2whQAhi_lmzdU1qLUlzTjhHSDA&export=download", asSettings.Settings.Item("OSSLDirectory").Value + "\OSSL.zip")
                    Console.WriteLine("Extracting OpenSSL..." + nl)
                    ZipFile.ExtractToDirectory(asSettings.Settings.Item("OSSLDirectory").Value + "\OSSL.zip", asSettings.Settings.Item("OSSLDirectory").Value)
                    File.Copy(asSettings.Settings.Item("OSSLDirectory").Value + "\share\openssl.cnf", asSettings.Settings.Item("OSSLDirectory").Value + "\bin\openssl.cnf")
                    Console.WriteLine("OpenSSL is now installed to " + asSettings.Settings.Item("OSSLDirectory").Value + nl)
                    Main()
                Catch ex As Exception
                    Console.WriteLine(ex)
                    Console.WriteLine("Error in downloading OpenSSL..Exiting...")
                    Exit Function
                End Try
            Else
                Exit Function
            End If
        Else
            Console.WriteLine("OpenSSL is not in the path defined, would you like to install it?" + nl)
            input = Console.ReadLine()
            If input.ToLower.Equals("y") Then
                Try
                    Console.WriteLine("Downloading OpenSSL..." + nl)
                    DownloadFile("https://docs.google.com/uc?authuser=0&id=0B2whQAhi_lmzdU1qLUlzTjhHSDA&export=download", asSettings.Settings.Item("OSSLDirectory").Value + "\OSSL.zip")
                    Console.WriteLine("Extracting OpenSSL..." + nl)
                    ZipFile.ExtractToDirectory(asSettings.Settings.Item("OSSLDirectory").Value + "\OSSL.zip", asSettings.Settings.Item("OSSLDirectory").Value)
                    File.Copy(asSettings.Settings.Item("OSSLDirectory").Value + "\share\openssl.cnf", asSettings.Settings.Item("OSSLDirectory").Value + "\bin\openssl.cnf")
                    Console.WriteLine("OpenSSL is now installed to " + asSettings.Settings.Item("OSSLDirectory").Value + nl)
                    Main()
                Catch ex As Exception
                    Console.WriteLine(ex)
                    Console.WriteLine("Error in downloading OpenSSL..Exiting...")
                    Exit Function
                End Try
            Else
                Exit Function
            End If
        End If
    End Function

    Public Sub DownloadFile(ByVal _URL As String, ByVal _SaveAs As String)
        Try
            Dim _WebClient As New System.Net.WebClient()
            ' Downloads the resource with the specified URI to a local file.
            _WebClient.DownloadFile(_URL, _SaveAs)
        Catch _Exception As Exception
            ' Error
            Console.WriteLine("Exception caught in process: {0}", _Exception.ToString())
        End Try
    End Sub

    'find and replace cert strings
    Private Sub FindReplaceString(ByVal fileName As String, ByVal oldValue As String, ByVal newValue As String)

        Using fs As New FileStream(fileName, FileMode.Open, FileAccess.ReadWrite)

            Dim sr As New StreamReader(fs)

            Dim fileContent As String = sr.ReadToEnd()
            fs.SetLength(0)
            sr.Close() ' Close the Reader
            fileContent = fileContent.Replace(oldValue, newValue)

            Dim sw As New StreamWriter(fileName)
            sw.Write(fileContent)
            sw.Flush()
            sw.Close()
        End Using
    End Sub

    'set cert info for changing certs
    Function parse()
        Dim l As Integer = asSettings.Settings.Item("CertFolder").Value.Length + 1

        oldPub = asSettings1.Settings.Item("SSAgentCertFile").Value.Remove(0, l)
        oldPrivate = asSettings1.Settings.Item("SSCommunicationCertPath").Value.Remove(0, l)

        If My.Application.CommandLineArgs.Count > 0 And My.Application.CommandLineArgs(0).ToLower.Equals("-z") Then
            manager()
            Return False
        ElseIf My.Application.CommandLineArgs.Count > 1 And My.Application.CommandLineArgs(0).ToLower.Equals("-n") Then

            If Directory.Exists(asSettings.Settings.Item("OSSLDirectory").Value) And File.Exists(asSettings.Settings.Item("OSSLDirectory").Value + "\bin\openssl.cnf") Then


                Dim pubTemp = My.Application.CommandLineArgs(1).ToLower()
                Dim privTemp = My.Application.CommandLineArgs(2).ToLower()
                Dim contFlag As Boolean = createCerts(pubTemp, privTemp)
                If contFlag = True Then
                    publicCert = pubTemp
                    privateCert = privTemp
                    Return True
                Else
                    Console.WriteLine("An error ocurred during Certificate Creation, Please try again
or try create certs manually and use -d")
                End If
            Else
                Console.WriteLine("It appears that OpenSSL is not installed/ or installed in the correct directory. Please install to """ + asSettings.Settings.Item("OSSLDirectory").Value + """ and make sure the
""openssl.cnf"" file is in the \bin folder. This information is also in the ReadMe.txt")
                Return False
            End If

        ElseIf My.Application.CommandLineArgs.Count > 1 And My.Application.CommandLineArgs(0).ToLower.Equals("-d") Then

            publicCert = My.Application.CommandLineArgs(1).ToLower
            privateCert = My.Application.CommandLineArgs(2).ToLower
            Return True
            ElseIf My.Application.CommandLineArgs.Count > 0 And My.Application.CommandLineArgs(0).Equals("?") Then
                Console.WriteLine(help)
                Return False

            Else
                Console.Write("Incorrect Format Please use ? to see help")
            Return False
        End If
    End Function

    'Management Function for additional info
    Function manager()
        Dim intInput As Integer = 0
        Console.WriteLine("")
        Console.WriteLine("Main Menu")
        Console.WriteLine("==========================")
        Console.WriteLine("1. view paths")
        Console.WriteLine("2. Create Certs")
        Console.WriteLine("3. apply SSCONFIG")
        Console.WriteLine("4. Get Current Cert Names")
        Console.WriteLine("5. Check OpenSSL")
        Console.WriteLine("6. Exit" & vbNewLine)

        Console.WriteLine("Enter your choice: ")
        intInput = Val(Console.ReadLine())

        Select Case intInput
            Case 1
                Console.WriteLine(asSettings.Settings.Item("OSSLDirectory").Value + "\bin\openssl.exe")
                Console.WriteLine(asSettings.Settings.Item("privateCert").Value)
                Console.WriteLine(asSettings.Settings.Item("publicCert").Value)
                Exit Select
            Case 2

                Dim nl As String = Environment.NewLine
                    Dim input As String
                Dim x = Console.ReadLine()

                Console.WriteLine("please enter the public cert name inluding .crt" + nl)
                input = Console.ReadLine()
                Dim pub = input
                Console.WriteLine("please enter the private cert name including .pem" + nl)
                input = Console.ReadLine()
                Dim priv = input
                createCerts(pub, priv)
                Console.WriteLine("Certificates created!")
                Exit Select
            Case 3
                SSClick()
            Case 4
                GetCertNames()

            Case 5
                OSSLDownload()
            Case 6
                Exit Function
            Case Else
                Console.WriteLine("Please select from 1 to 6")
        End Select
    End Function

    Function GetCertNames()

        Dim l As Integer = asSettings.Settings.Item("CertFolder").Value.Length + 1
        Dim x As String
        Dim privateDisplay As String = asSettings1.Settings.Item("SSAgentCertFile").Value.Remove(0, l)
        Dim publicDisplay As String = asSettings1.Settings.Item("SSCommunicationCertPath").Value.Remove(0, l)

        Console.WriteLine("Private Certificate: " + privateDisplay)
        Console.WriteLine("Public Certificate: " + publicDisplay)

        Console.WriteLine("Would you like to replace these with OpenSSL?")


        Try
            Dim nl As String = Environment.NewLine
            Dim input As String
            x = Console.ReadLine()
            If x.ToLower.Equals("y") Then
                Console.WriteLine("please enter the public cert name inluding .crt" + nl)
                input = Console.ReadLine()
                Dim pub = input
                Console.WriteLine("please enter the private cert name including .pem" + nl)
                input = Console.ReadLine()
                Dim priv = input
                createCerts(pub, priv)
                publicCert = pub
                privateCert = priv
                FindReplaceString(WMPath, oldPub, publicCert)
                FindReplaceString(WMPath, oldPrivate, privateCert + ".adp12")
                FindReplaceString(SSPath, oldPub, publicCert)
                FindReplaceString(SSPath, oldPrivate, privateCert)

                restartService()
                SSClick()
                cleanup()
            Else
                Exit Function
            End If
        Catch ex As Exception
            Console.WriteLine(ex)
        End Try
    End Function

    'Use OpenSSL to create new Cert pair
    Function createCerts(pub As String, prive As String)
        Try
            Console.WriteLine("Generating Key...")
            Shell(asSettings.Settings.Item("OSSLDirectory").Value + "\bin\openssl.exe" + " " + "genrsa -out " + asSettings.Settings.Item("OSSLDirectory").Value + "\OSSLkey.key" + " " + asSettings.Settings.Item("keySize").Value,, True)

            Console.WriteLine("Key Created, Generating Request...")

            Shell(asSettings.Settings.Item("OSSLDirectory").Value + "\bin\openssl.exe" + " " + "req -new -key " + asSettings.Settings.Item("OSSLDirectory").Value + "\OSSLKey.key" + " -out " + asSettings.Settings.Item("OSSLDirectory").Value + "\OSSLRequest.csr -config " +
             asSettings.Settings.Item("OSSLDirectory").Value + "\bin\openssl.cnf -subj /C=" +
             asSettings.Settings.Item("Country").Value + "/ST=" +
             asSettings.Settings.Item("State").Value + "/L=" +
             asSettings.Settings.Item("City").Value + "/O=" +
             asSettings.Settings.Item("Orginization").Value + "/OU=" +
             asSettings.Settings.Item("OrginizationalUnit").Value + "/CN=" +
             asSettings.Settings.Item("CommonName").Value,, True)

            Console.WriteLine("Request Complete, Generating Certificates...")

            Shell(asSettings.Settings.Item("OSSLDirectory").Value + "\bin\openssl.exe" + " " + "x509 -req -days 3650 -in " + asSettings.Settings.Item("OSSLDirectory").Value + "\" + "OSSLRequest.csr -signkey " + asSettings.Settings.Item("OSSLDirectory").Value + "\OSSLKey.key -out " + asSettings.Settings.Item("CertFolder").Value + "\" + pub,, True)

            If createPrivate(pub, prive) = True Then
                Console.WriteLine("New Certificate Creation Complete")
                Return True
            Else
                Console.WriteLine("New Certificate Creation Failed To create Private Cert. Please retry Or generate manually And use (-d)")
                Return False
            End If

        Catch ex As Exception
            Console.WriteLine(ex)
            Return False
        End Try

    End Function

    'Create and run batch to output the PEM Private Cert
    Function createPrivate(pub As String, priv As String)
        Try

            Dim sb As New Text.StringBuilder

            sb.AppendLine("@echo off")
            sb.AppendLine("type " + asSettings.Settings.Item("CertFolder").Value + "\" + pub + " " + asSettings.Settings.Item("OSSLDirectory").Value + "\OSSLKey.key > " + asSettings.Settings.Item("CertFolder").Value + "\" + priv)
            IO.File.WriteAllText(asSettings.Settings.Item("OSSLDirectory").Value + "\ConvertPrivate.bat", sb.ToString())
            Shell(asSettings.Settings.Item("OSSLDirectory").Value + "\ConvertPrivate.bat")
            Return True
        Catch ex As Exception
            Console.WriteLine(ex)
            Return False
        End Try
    End Function

    'Restart the WorkManager and SiteServer Service
    Function restartService()
        Dim sc As New ServiceController()
        sc.ServiceName = "R1WorkManager"

        Dim sc2 As New ServiceController()
        sc2.ServiceName = "SiteServer"

        Try
            sc.Stop()
            Console.WriteLine("Stopping Work Manager Service...")
            sc.WaitForStatus(ServiceControllerStatus.Stopped)

            sc2.Stop()
            Console.WriteLine("Stopping Site Server Service...")
            sc2.WaitForStatus(ServiceControllerStatus.Stopped)

            sc.Start()
            Console.WriteLine("Starting Work Manager Service...")
            sc.WaitForStatus(ServiceControllerStatus.Running)
            sc2.Start()
            Console.WriteLine("Starting Site Server Service...")
            sc.WaitForStatus(ServiceControllerStatus.Running)

        Catch ex As Exception
            Console.WriteLine(ex)
        End Try
    End Function

    Declare Auto Function FindWindow Lib "USER32.DLL" (
    ByVal lpClassName As String,
    ByVal lpWindowName As String) As IntPtr

    ' Activate an application window.
    Declare Auto Function SetForegroundWindow Lib "USER32.DLL" _
    (ByVal hWnd As IntPtr) As Boolean

    ' Send a series of key presses to the SSConfig application.
    Function SSClick()
        Process.Start(asSettings.Settings.Item("DriveLetter").Value + ":\Program Files\Resolution1\SiteServer\SS_Config.exe")
        Threading.Thread.Sleep(6000)
        ' Get a handle to the SSConfig application. The window class
        ' and window name were obtained using the Spy++ tool.
        Dim SSHandle As IntPtr = FindWindow("WindowsForms10.Window.8.app.0.2bf8098_r9_ad1", "Site Server Configuration")

        ' 
        SetForegroundWindow(SSHandle)
            SendKeys.SendWait("{TAB 39}")
        SendKeys.SendWait("{ENTER}")
        Threading.Thread.Sleep(10000)
        SendKeys.SendWait("{ENTER}")
        Threading.Thread.Sleep(6000)
        SendKeys.SendWait("{ENTER}")
        SendKeys.SendWait("{TAB}")
        SendKeys.SendWait("{ENTER}")





    End Function

    'Cleans up the mess made from OpenSSL and the PEM Cert
    Function cleanup()

        File.Delete(asSettings.Settings.Item("OSSLDirectory").Value + "\" + "OSSLRequest.csr")
        File.Delete(asSettings.Settings.Item("OSSLDirectory").Value + "\" + "OSSLkey.key")
        File.Delete(asSettings.Settings.Item("OSSLDirectory").Value + "\" + "ConvertPrivate.bat")
        Threading.Thread.Sleep(6000)
        File.Delete(asSettings.Settings.Item("CertFolder").Value + "\" + privateCert)
        Return True
    End Function



End Module
