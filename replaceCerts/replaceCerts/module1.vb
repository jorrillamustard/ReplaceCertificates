Imports System.IO
Imports System.ServiceProcess
Imports System.Configuration
Imports System.Windows.Forms
Imports System.IO.Compression
Imports System.Management.Automation
Imports System.Management.Automation.Runspaces
Imports System.Collections.ObjectModel
Imports System.Text

Module module1
    Dim publicCert As String = ""
    Dim privateCert As String = ""
    Dim oldPub As String = ""
    Dim oldPrivate As String = ""
    Dim SSPrivate As String = ""
    Dim user As String = ""
    Dim Pass As String = ""
    Dim help As String = "Format: replacecerts.exe <-n,-d> <New Public Cert Name> <New Private Cert Name>"
    Dim nl As String = Environment.NewLine
    Dim cAppConfig As Configuration = ConfigurationManager.OpenExeConfiguration(My.Application.Info.DirectoryPath + "\replaceCerts.exe")
    Dim asSettings As AppSettingsSection = cAppConfig.AppSettings

    Dim cAppConfig1 As Configuration = ConfigurationManager.OpenExeConfiguration(asSettings.Settings.Item("DriveLetter").Value + ":\Program Files\Resolution1\Work Manager\Infrastructure.WorkExecutionServices.Host.exe")
    Dim asSettings1 As AppSettingsSection = cAppConfig1.AppSettings


    Dim WMPath As String = asSettings.Settings.Item("DriveLetter").Value + ":\Program Files\Resolution1\Work Manager\Infrastructure.WorkExecutionServices.Host.exe.config"
    Dim SSPath As String = "C:\ProgramData\Resolution1\SiteServer\siteserver.config"

    'main function for application
    Sub Main()

        'initial check to see if openssl is installed
        If checkOSSLinstalled() = True Then
            GoTo start
        Else
            Directory.CreateDirectory(asSettings.Settings.Item("OSSLDirectory").Value)
            OSSLDownload()
        End If
start:
        'start the process of creating certificates and changing configurations
        Dim flag As Boolean = parse()

        If flag = False Then
            GoTo closeme 'close the app if certs were not created or changed properly
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
            'Replace the cert names in config files
Replace:
            FindReplaceString(WMPath, oldPub, publicCert)
            FindReplaceString(WMPath, oldPrivate, privateCert + ".adp12")
            FindReplaceString(SSPath, oldPub, publicCert)
            FindReplaceString(SSPath, SSPrivate, privateCert)

            'restart services and apply SS and cleanup
            restartService()
            'check if public ss needs configuring
            If My.Application.CommandLineArgs(3).Equals("-p") Then
                If My.Application.CommandLineArgs(5).Equals("-user") Then
                    user = My.Application.CommandLineArgs(6)
                End If
                If My.Application.CommandLineArgs(7).Equals("-pass") Then
                    Pass = My.Application.CommandLineArgs(8)
                End If
                SSconfigChild(My.Application.CommandLineArgs(4).ToString)
                SSClick()
                cleanup()
                Console.WriteLine("Now Configuring Public Site Server...")
                PublicSS()
                ' restartService()
            Else
                SSClick()
                cleanup()
            End If
        End If
closeme:
    End Sub

    'edits SS config to add in Child information
    Private Sub SSconfigChild(pubIP As String)
        Dim text As String = IO.File.ReadAllText(SSPath)
        Dim find As String = "Children="""
        Dim xst As Integer = text.IndexOf(find) + find.Length
        Dim xend As Integer = text.IndexOf(":54545") - 1
        Dim xsub As String = text.Substring(xst, (xend - xst) + 1)
        text = text.Replace(xsub, pubIP)
        Dim Writer As System.IO.StreamWriter
        Writer = New System.IO.StreamWriter(SSPath) '<-- Where to write to
        Writer.Write(text)
        Writer.Close()
    End Sub

    'Checks if OSSL is installed
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


    'Main function for OpenSSL Download
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

    'Downloads OpenSSL
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
        If oldPrivate.Contains(".adp12") Then
            SSPrivate = oldPrivate
        Else
            SSPrivate = oldPrivate + ".adp12"
        End If

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

    'function for configuring public SS
    Function PublicSS()
        Dim publicSSIP As String = My.Application.CommandLineArgs(4).ToString

        Try
            If installPowershellFunction() = True Then
                Dim RootIP As String = pubIP()
                Directory.CreateDirectory(asSettings.Settings.Item("OSSLDirectory").Value + "\pubss")
                Dim path As String = Directory.GetCurrentDirectory()
                'Console.WriteLine(path)
                Console.WriteLine("Creating Public SS Payload...")
                FileCopy(path + "\pubss.exe", asSettings.Settings.Item("OSSLDirectory").Value + "\pubss\pubss.exe")
                'FileCopy(asSettings.Settings.Item("OSSLDirectory").Value + "\RSSPUBIP.txt", asSettings.Settings.Item("OSSLDirectory").Value + "\pubss\RSSPUBIP.txt")
                FileCopy(asSettings1.Settings.Item("SSCommunicationCertPath").Value, asSettings.Settings.Item("OSSLDirectory").Value + "\pubss\" + privateCert + ".adp12")
                Dim rootSSIP As String = pubIP()
                ' Console.WriteLine(publicSSIP + ", " + rootSSIP) 'test write to view IP's (need to comment out)
                'Console.WriteLine("Creating Public SS Payload...")
                ZipFile.CreateFromDirectory(asSettings.Settings.Item("OSSLDirectory").Value + "\pubss", asSettings.Settings.Item("OSSLDirectory").Value + "\pubss.zip")
                Console.WriteLine("Creating Session Script...")
                If asSettings.Settings.Item("URLSession").Value.Equals("y") Then
                    Console.WriteLine("Sending Files to PubSS...")
                    UnzipScript()
                    CreatePUBSSscriptURLFormat(publicSSIP, RootIP)
                    Console.WriteLine("Configuring Public SS...")
                    powershellCMD(asSettings.Settings.Item("OSSLDirectory").Value + "\sendFiles.ps1")

                Else
                    CreatePUBSSscript(publicSSIP)
                    executePubssConfig(publicSSIP, RootIP)
                    Console.WriteLine("Sending Files to PubSS...")
                    powershellCMD(asSettings.Settings.Item("OSSLDirectory").Value + "\sendFiles.ps1")
                    Console.WriteLine("Configuring Public SS...")
                    powershellCMD(asSettings.Settings.Item("OSSLDirectory").Value + "\ConfigPubss.ps1")
                End If
                additionalCleanup()
                Console.WriteLine("Complete!")
                Else
                    Console.WriteLine("Issue configuring Public Site Server...")
            End If
        Catch ex As Exception
            Console.WriteLine(ex)
        End Try

    End Function

    'Creates the SendFiles.ps1 to get the files down to the server
    Function CreatePUBSSscript(Pip As String)
        Using sw As StreamWriter = File.CreateText(asSettings.Settings.Item("OSSLDirectory").Value + "\sendFiles.ps1")
            sw.WriteLine("$password = ConvertTo-SecureString " + """" + Pass + """" + " -AsPlainText -Force")
            sw.WriteLine("$cred= New-Object System.Management.Automation.PSCredential (" + """" + user + """" + ", $password )")
            sw.WriteLine("import-module " + Directory.GetCurrentDirectory + "\Send-File.psm1")
            sw.WriteLine("$session = New-PSSession -ComputerName " + Pip + " -credential $cred")
            ' sw.WriteLine("mkdir C:\pubss")
            sw.WriteLine("Send-File -path " + asSettings.Settings.Item("OSSLDirectory").Value + "\pubss\pubss.exe," + asSettings.Settings.Item("OSSLDirectory").Value + "\pubss\" + privateCert + ".adp12 -destination C:\pubss -Session $session")

        End Using
    End Function

    'function to do connection for URL format (specifically for Azure)
    Function CreatePUBSSscriptURLFormat(Pip As String, Rip As String)
        Using sw As StreamWriter = File.CreateText(asSettings.Settings.Item("OSSLDirectory").Value + "\sendFiles.ps1")
            sw.WriteLine("$password = ConvertTo-SecureString " + """" + Pass + """" + " -AsPlainText -Force")
            sw.WriteLine("$cred= New-Object System.Management.Automation.PSCredential (" + """" + user + """" + ", $password )")
            sw.WriteLine("import-module " + Directory.GetCurrentDirectory + "\Send-File.psm1")
            sw.WriteLine("$SessionOptions = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck")
            sw.WriteLine("$session = New-PSSession -ConnectionUri https://" + Pip + ":5986 -credential $cred -SessionOption $SessionOptions")
            ' sw.WriteLine("Send-File -path " + asSettings.Settings.Item("OSSLDirectory").Value + "\pubss\pubss.exe -destination C:\pubss -Session $session")
            ' sw.WriteLine("Send-File -path " + asSettings.Settings.Item("OSSLDirectory").Value + "\pubss\" + privateCert + " -destination C:\pubss -Session $session")
            sw.WriteLine("Send-File -path " + asSettings.Settings.Item("OSSLDirectory").Value + "\pubss.zip," + asSettings.Settings.Item("OSSLDirectory").Value + "\unzip.ps1 -destination C:\ -Session $session")
            sw.WriteLine("Invoke-Command -Session $session -scriptBlock {C:\unzip.ps1}")

            Dim command As String = "Invoke-Command -Session $session -scriptBlock {C:\pubss\pubss.exe " + Pip + " " + Rip + " c:\pubss\" + privateCert + ".adp12}"
            Dim newcommand As String = command.Replace(vbCr, "").Replace(vbLf, "")
            sw.WriteLine(newcommand)
        End Using
    End Function


    Function UnzipScript()
        Using sw As StreamWriter = File.CreateText(asSettings.Settings.Item("OSSLDirectory").Value + "\unzip.ps1")
            sw.WriteLine("mkdir C:\pubss -Force")
            sw.WriteLine("$shell = new-object -com shell.application")
            sw.WriteLine("$zip = $shell.NameSpace('C:\pubss.zip')")
            sw.WriteLine("foreach($item in $zip.items())")
            sw.WriteLine("{")
            sw.WriteLine("$shell.Namespace('C:\pubss').copyhere($item,0x14)")
            sw.WriteLine("}")
        End Using
    End Function


    'Creates script to execute the files
    Function executePubssConfig(Pip As String, Rip As String)
        Dim command As String = "winrs -r: " + Pip + " -u:" + user + " -p:" + Pass + " C:\pubss\pubss.exe " + Pip + " " + Rip + " c:\pubss\" + privateCert + ".adp12"
        Dim newcommand As String = command.Replace(vbCr, "").Replace(vbLf, "")
        Using sw As StreamWriter = File.CreateText(asSettings.Settings.Item("OSSLDirectory").Value + "\ConfigPubss.ps1")
            sw.WriteLine(newcommand)
        End Using
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
        Console.WriteLine("6. PubIP")
        Console.WriteLine("7. Exit" & vbNewLine)

        Console.WriteLine("Enter your choice:  ")
        intInput = Val(Console.ReadLine())

        Select Case intInput
            Case 1
                Console.WriteLine(asSettings.Settings.Item("OSSLDirectory").Value + "\bin\openssl.exe")
                Dim l As Integer = asSettings.Settings.Item("CertFolder").Value.Length + 1
                Dim x As String
                Dim privateDisplay As String = asSettings1.Settings.Item("SSAgentCertFile").Value.Remove(0, l)
                Dim publicDisplay As String = asSettings1.Settings.Item("SSCommunicationCertPath").Value.Remove(0, l)
                Console.WriteLine(SSPath)

                Console.WriteLine("Private Certificate: " + privateDisplay)
                Console.WriteLine("Public Certificate: " + publicDisplay)
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
                Console.WriteLine(pubIP)

            Case 7
                Exit Function
            Case Else
                Console.WriteLine("Please select from 1 to 6")
        End Select
    End Function

    Function GetCertNames()

        Dim l As Integer = asSettings.Settings.Item("CertFolder").Value.Length + 1
        Dim x As String
        Dim publicdisplay As String = asSettings1.Settings.Item("SSAgentCertFile").Value.Remove(0, l)
        Dim privatedisplay As String = asSettings1.Settings.Item("SSCommunicationCertPath").Value.Remove(0, l)

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
            Return True
        Catch ex As Exception
            Console.WriteLine(ex)
            Return False
        End Try
    End Function

    'Used for SSClick Function -need to be at class level!
    Declare Auto Function FindWindow Lib "USER32.DLL" (
    ByVal lpClassName As String,
    ByVal lpWindowName As String) As IntPtr

    ' Activate an application window. -need to be at class level!
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
        Try
            File.Delete(asSettings.Settings.Item("OSSLDirectory").Value + "\" + "OSSLRequest.csr")
            File.Delete(asSettings.Settings.Item("OSSLDirectory").Value + "\" + "OSSLkey.key")
            File.Delete(asSettings.Settings.Item("OSSLDirectory").Value + "\" + "ConvertPrivate.bat")
            Threading.Thread.Sleep(6000)
            'File.Delete(asSettings.Settings.Item("CertFolder").Value + "\" + privateCert)

            Return True
        Catch ex As Exception
            Console.WriteLine(ex)

        End Try
    End Function

    Function additionalCleanup()
        Try
            Console.WriteLine("Cleaning up...")
            Directory.Delete(asSettings.Settings.Item("OSSLDirectory").Value + "\pubss", True)
            File.Delete(asSettings.Settings.Item("OSSLDirectory").Value + "\sendFiles.ps1")
            File.Delete(asSettings.Settings.Item("OSSLDirectory").Value + "\unzip.ps1")
            File.Delete(asSettings.Settings.Item("OSSLDirectory").Value + "\pubss.zip")
            File.Delete(asSettings.Settings.Item("OSSLDirectory").Value + "\ConfigPubss.ps1")
            powershellCMD("get-pssession | remove-pssession")
        Catch ex As Exception
            Console.WriteLine(ex)
        End Try

    End Function

    'gets public ip address
    Function pubIP()
        Dim IpCommand As String = "$wc = New-object System.Net.WebClient" + nl + " 
$wc.DownloadString(""http://myexternalip.com/raw"")"

        Dim PublicIP As String = powershellCMD(IpCommand)

        ' Using sw As StreamWriter = File.CreateText(asSettings.Settings.Item("OSSLDirectory").Value + "\RSSPUBIP.txt")
        'sw.WriteLine(PublicIP)
        'sw.Close()
        '  End Using


        Return PublicIP
    End Function

    'installs Send-File Function for sending using WinRM
    Function installPowershellFunction()
        If File.Exists(Directory.GetCurrentDirectory + "\Send-File.psm1") Then
            powershellCMD("import-module " + Directory.GetCurrentDirectory + "\Send-File.psm1")
            Console.WriteLine("Installing Send-File Function for Powershell...")
            Return True
        Else
            Console.WriteLine("Could not find the Send-File Function...")
            Return False
        End If


    End Function

    'uses powershell to run commands - opens stream to shell
    Private Function powershellCMD(ByVal script As String) As String
        'Create Powershell Runspace'
        Dim myRunSpace As Runspace = RunspaceFactory.CreateRunspace()
        'open the runspace'
        myRunSpace.Open()

        'create pipeline and feed it script text'
        Dim myPipeline As Pipeline = myRunSpace.CreatePipeline()
        myPipeline.Commands.AddScript(script)
        Dim myCommand As Command = New Command(script)

        ' add an extra command to transform the script output objects into nicely formatted strings 
        ' remove this line to get the actual objects that the script returns. For example, the script 
        ' "Get-Process" returns a collection of System.Diagnostics.Process instances. 
        myPipeline.Commands.Add("Out-String")

        'execute the script
        Dim results As Collection(Of PSObject) = myPipeline.Invoke()

        'close the runspace
        myRunSpace.Close()

        Dim myStringBuilder As New StringBuilder()

        For Each obj As PSObject In results
            myStringBuilder.AppendLine(obj.ToString())
        Next

        'return the script results as string
        Return myStringBuilder.ToString()
    End Function

    Function checkTrustedHost()

    End Function

End Module
