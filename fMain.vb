Imports System.Diagnostics
Imports System.IO
Imports System.Management
Imports System.Media
Imports System.Security.Principal
Imports Microsoft.Win32

' OP2Mapper2Helper
' https://github.com/leviathan400/OP2Mapper2Helper
'
' A helper tool that registers all required VB6 COM libraries for OP2Mapper2,
' shows their current registration locations, and launches the mapper.
'
' Outpost 2: Divided Destiny is a real-time strategy video game released in 1997.

Public Class fMain

    Public ApplicationName As String = "OP2Mapper2Helper"
    Public ApplicatioVersion As String = "0.3.0"
    Public ApplicatioBuild As String = "0015"
    Public ApplicationDescription As String = "Tool that registers required COM libraries For OP2Mapper2 and shows their registration locations."
    Public ApplicationWebsite As String = "https//github.com/leviathan400/OP2Mapper2Helper"

    Public WindowsOSVersion As String = "Unknown"

    Public Mapper2Path As String = IO.Path.Combine(Application.StartupPath, "Mapper2.exe")

    Private MapperFiles As String() = {
        "cPopMenu6.ocx",
        "OP2Editor.dll",
        "SSubTmr6.dll",
        "vbalDTab6.ocx",
        "vbalTBar6.ocx",
        "vbalTreeView6.ocx"
    }

    Private Sub fMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Icon = My.Resources.Mapper2_Icon_256
        Me.Text = "OP2Mapper2 Helper"

        GetRealWindowsVersion() ' Set WindowsOSVersion
        'AppendToConsole("Detected OS " & GetRealWindowsVersion())
        AppendToConsole("Windows OS: " & WindowsOSVersion)

        ShowAdminStatus()

        CheckFileExists(Mapper2Path)

        DisplayFileVersion(Mapper2Path)

        AppendToConsole("")

        Dim player As New SoundPlayer(My.Resources.beep8)
        player.Play()
    End Sub

    Private Sub AppendToConsole(text As String)
        txtConsole.AppendText(text & vbCrLf)
        txtConsole.SelectionStart = txtConsole.Text.Length
        txtConsole.ScrollToCaret()
    End Sub

    Public Sub CheckFileExists(filename As String)
        Dim filePath As String = IO.Path.Combine(Application.StartupPath, filename)

        If IO.File.Exists(filePath) Then
            'AppendToConsole(filename & " present")
        Else
            AppendToConsole(IO.Path.GetFileName(filename) & " Not found In folder!")
        End If
    End Sub

    Private Function GetWindowsVersion() As String '
        ' Example: AppendToConsole("Detected OS " & GetWindowsVersion())
        Try
            Dim searcher As New ManagementObjectSearcher("Select Caption, Version FROM Win32_OperatingSystem")
            For Each os As ManagementObject In searcher.Get()
                Dim caption = os("Caption").ToString().Trim()
                Dim version = os("Version").ToString().Trim()
                Return $"{caption} (Version {version})"
            Next
        Catch ex As Exception
            Return "Unable To detect OS " & ex.Message
        End Try

        Return "Unknown Windows Version"
    End Function

    Private Function GetRealWindowsVersion() As String
        Try
            Dim searcher As New System.Management.ManagementObjectSearcher(
            "Select Caption, Version FROM Win32_OperatingSystem")

            For Each os As System.Management.ManagementObject In searcher.Get()
                Dim caption As String = os("Caption").ToString().Trim()
                Dim version As String = os("Version").ToString().Trim()
                Dim lowerCaption As String = caption.ToLowerInvariant()

                If lowerCaption.Contains("windows 11") Then
                    WindowsOSVersion = "Windows 11"
                ElseIf lowerCaption.Contains("windows 10") Then
                    WindowsOSVersion = "Windows 10"
                ElseIf lowerCaption.Contains("windows 8.1") Then
                    WindowsOSVersion = "Windows 8.1"
                ElseIf lowerCaption.Contains("windows 8") Then
                    WindowsOSVersion = "Windows 8"
                ElseIf lowerCaption.Contains("windows 7") Then
                    WindowsOSVersion = "Windows 7"
                ElseIf lowerCaption.Contains("vista") Then
                    WindowsOSVersion = "Windows Vista"
                ElseIf lowerCaption.Contains("xp") Then
                    WindowsOSVersion = "Windows XP"
                ElseIf lowerCaption.Contains("server") Then
                    WindowsOSVersion = "Windows Server"
                    If lowerCaption.Contains("server 2022") Then
                        WindowsOSVersion = "Windows Server 2022"
                    End If
                Else
                    WindowsOSVersion = "Unknown"
                End If

                Return $"{caption} (Version {version})"
            Next

        Catch ex As Exception
            WindowsOSVersion = "Unknown"
            Return "Unable To detect OS " & ex.Message
        End Try

        WindowsOSVersion = "Unknown"
        Return "Unknown Windows Version"
    End Function

    Private Sub ShowAdminStatus()
        If IsRunningAsAdministrator() Then
            AppendToConsole("Running As Administrator Yes")
        Else
            AppendToConsole("Running As Administrator No")
        End If
    End Sub

    Private Function IsRunningAsAdministrator() As Boolean
        Try
            Dim identity = WindowsIdentity.GetCurrent()
            Dim principal = New WindowsPrincipal(identity)
            Return principal.IsInRole(WindowsBuiltInRole.Administrator)
        Catch
            Return False
        End Try
    End Function

    Public Function FindRegisteredComDll(target As String) As List(Of String)
        Dim results As New List(Of String)()

        ' Search HKCR\CLSID and HKCR\WOW6432Node\CLSID
        Dim roots = {
        "HKEY_CLASSES_ROOT\CLSID",
        "HKEY_CLASSES_ROOT\WOW6432Node\CLSID"
    }

        For Each root In roots
            Using clsidKey As RegistryKey = Registry.ClassesRoot.OpenSubKey(If(root.Contains("WOW"), "WOW6432Node\CLSID", "CLSID"))
                If clsidKey Is Nothing Then Continue For

                For Each subKeyName In clsidKey.GetSubKeyNames()
                    Using subKey = clsidKey.OpenSubKey(subKeyName & "\InprocServer32")
                        If subKey Is Nothing Then Continue For

                        Dim ComPath = TryCast(subKey.GetValue(""), String)
                        If String.IsNullOrEmpty(ComPath) Then Continue For

                        ' Match by filename or path
                        If ComPath.IndexOf(target, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            results.Add(ComPath)
                        End If
                    End Using
                Next
            End Using
        Next

        Return results.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
    End Function

    Public Sub DisplayFileVersion(filePath As String)
        Try
            Dim versionInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(filePath)

            AppendToConsole($"OP2Mapper2 Version {versionInfo.FileMajorPart}.{versionInfo.FileMinorPart}.{versionInfo.FilePrivatePart}")

        Catch ex As Exception
            Console.WriteLine("Error " & ex.Message)
        End Try
    End Sub

    Public Function GetDllVersion(dllPath As String) As String
        Try
            Dim info = FileVersionInfo.GetVersionInfo(dllPath)

            Return $"FileVersion={info.FileVersion}, " &
               $"ProductVersion={info.ProductVersion}, " &
               $"Description={info.FileDescription}"
        Catch ex As Exception
            Return "(Version info unavailable)"
        End Try
    End Function

    Private Sub FindCOMRegistration(ByVal ComName As String)
        ' Show COM Status
        Dim paths = FindRegisteredComDll(ComName)

        If paths.Count = 0 Then
            AppendToConsole(ComName & " Not registered.")
        Else
            AppendToConsole("Found " & ComName & " registration")
            For Each p In paths
                AppendToConsole("" & p)
                'AppendToConsole(GetDllVersion(p))
            Next
        End If
    End Sub

    Private Sub DisableButtons()
        btnRegisterLibraries.Enabled = False
        btnShowLibraries.Enabled = False
        btnRunMapper.Enabled = False
    End Sub

    Private Sub EnableButtons()
        btnRegisterLibraries.Enabled = True
        btnShowLibraries.Enabled = True
        btnRunMapper.Enabled = True
    End Sub

    Private Sub RegisterComLibrary(fileName As String)
        ' Helper to register one COM library with regsvr32
        Dim fullPath As String = Path.Combine(Application.StartupPath, fileName)

        If Not File.Exists(fullPath) Then
            AppendToConsole("   Error " & fileName & " Not found In " & Application.StartupPath)
            Return
        End If

        Try
            Dim p As New Process()
            p.StartInfo.FileName = "regsvr32.exe"
            p.StartInfo.Arguments = "/s """ & fullPath & """"
            p.StartInfo.UseShellExecute = True   ' No redirection, just run it

            ' Force UAC prompt for this app instead of requiring
            ' "Run As administrator", uncomment the next line:
            ' p.StartInfo.Verb = "runas"

            p.Start()
            p.WaitForExit()

            If p.ExitCode = 0 Then
                AppendToConsole("   OK")
            Else
                AppendToConsole("   regsvr32 Exit code " & p.ExitCode)
            End If

        Catch ex As Exception
            AppendToConsole("   Error " & ex.Message)
        End Try
    End Sub
    Private Sub btnRegisterLibraries_Click(sender As Object, e As EventArgs) Handles btnRegisterLibraries.Click
        DisableButtons()

        AppendToConsole("--------------------------------------------")
        AppendToConsole(" OP2Mapper2 Installation")
        AppendToConsole("--------------------------------------------")
        AppendToConsole("Welcome To the installer For OP2Mapper2")
        AppendToConsole("You may need To be logged On As a system administrator")
        AppendToConsole("If you are On Windows NT, 2000, Or XP.")
        AppendToConsole("The installation may still succeed without being an administrator,")
        AppendToConsole("however, If you Get errors On run, you will need To install again")
        AppendToConsole("As an administrator.")
        AppendToConsole("")
        'AppendToConsole("---------------------------")
        'AppendToConsole(" Running install script...")
        'AppendToConsole("---------------------------")

        AppendToConsole("Registering cPopMenu6.ocx...")
        RegisterComLibrary("cPopMenu6.ocx")
        AppendToConsole("Registering OP2Editor.dll...")
        RegisterComLibrary("OP2Editor.dll")
        AppendToConsole("Registering SSubTmr6.dll...")
        RegisterComLibrary("SSubTmr6.dll")
        AppendToConsole("Registering vbalDTab6.ocx...")
        RegisterComLibrary("vbalDTab6.ocx")
        AppendToConsole("Registering vbalTbar6.ocx...")
        RegisterComLibrary("vbalTbar6.ocx")
        AppendToConsole("Registering vbalTreeView6.ocx...")
        RegisterComLibrary("vbalTreeView6.ocx")

        'AppendToConsole("---------------------------")
        'AppendToConsole(" Finishing installation...")
        'AppendToConsole("---------------------------")
        AppendToConsole("")
        AppendToConsole("Click 'Run Mapper' when you’re ready.")
        AppendToConsole("")

        EnableButtons()
    End Sub

    Private Sub btnShowLibraries_Click(sender As Object, e As EventArgs) Handles btnShowLibraries.Click
        DisableButtons()

        'FindCOMRegistration("OP2Editor")

        ShowLibrariesMapper()

        AppendToConsole("")

        ShowLibrariesSystem()

        EnableButtons()
    End Sub

    Private Sub ShowLibrariesMapper()
        Dim LibraryFileCount As Int16 = 0

        AppendToConsole("COM Registration Status:")
        For Each LibraryFile In MapperFiles
            'FindCOMRegistration(Path.GetFileNameWithoutExtension(LibraryFile))
            FindCOMRegistration(LibraryFile)
            LibraryFileCount = LibraryFileCount + 1
        Next
        'AppendToConsole(LibraryFileCount & " COM Library Files Checked")
    End Sub

    Private Sub ShowLibrariesSystem()
        AppendToConsole("VB6 COM Registration Status:")
        FindCOMRegistration("MSCOMCTL.OCX")
        FindCOMRegistration("COMDLG32.OCX")
    End Sub

    Private Sub btnRunMapper_Click(sender As Object, e As EventArgs) Handles btnRunMapper.Click
        Try
            If Not File.Exists(Mapper2Path) Then
                AppendToConsole("")
                AppendToConsole("Mapper2.exe not found in folder!")
                Exit Sub
            End If

            AppendToConsole("")
            AppendToConsole("Launching OP2Mapper2...")
            Process.Start(Mapper2Path)

        Catch ex As Exception
            AppendToConsole("ERROR launching: " & ex.Message)
        End Try
    End Sub

    Private Sub btnAbout_Click(sender As Object, e As EventArgs) Handles btnAbout.Click
        Using AboutForm As New fAbout()
            AboutForm.ApplicationTitle = ApplicationName
            AboutForm.ApplicationVersion = ApplicatioVersion
            AboutForm.ApplicationBuild = ApplicatioBuild
            AboutForm.ApplicationDescription = ApplicationDescription
            AboutForm.ApplicationCopyright = "Outpost Universe © 2025"
            AboutForm.ApplicationWebsite = ApplicationWebsite
            AboutForm.ShowDialog(Me)
        End Using
    End Sub

End Class
