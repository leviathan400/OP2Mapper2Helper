Imports System.Diagnostics
Imports System.IO
Imports System.Management
Imports System.Media
Imports System.Security.Principal
Imports Microsoft.Win32

' OP2Mapper2Helper
'
' A helper tool that registers all required VB6 COM libraries for OP2Mapper2, shows their current registration locations, and launches the mapper.
Public Class fMain

    Public Version As String = "0.2.0"

    Public WindowsOSVersion As String = "Unknown"

    Private MapperFiles As String() = {
        "cPopMenu6.ocx",
        "OP2Editor.dll",
        "SSubTmr6.dll",
        "vbalDTab6.ocx",
        "vbalTBar6.ocx",
        "vbalTreeView6.ocx"
    }

    Private Sub fMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Icon = My.Resources.Mapper2_Icon
        Me.Text = "OP2Mapper2 Helper"

        AppendToConsole("Detected OS: " & GetRealWindowsVersion())
        'AppendToConsole("WindowsOSVersion Short Code: " & WindowsOSVersion)

        ShowAdminStatus()

        AppendToConsole("")

        Dim player As New SoundPlayer(My.Resources.beep8)
        player.Play()
    End Sub

    Private Sub AppendToConsole(text As String)
        txtConsole.AppendText(text & vbCrLf)
        txtConsole.SelectionStart = txtConsole.Text.Length
        txtConsole.ScrollToCaret()
    End Sub

    Private Function GetWindowsVersion() As String '
        ' Example: AppendToConsole("Detected OS: " & GetWindowsVersion())

        Try
            Dim searcher As New ManagementObjectSearcher("SELECT Caption, Version FROM Win32_OperatingSystem")
            For Each os As ManagementObject In searcher.Get()
                Dim caption = os("Caption").ToString().Trim()
                Dim version = os("Version").ToString().Trim()
                Return $"{caption} (Version {version})"
            Next
        Catch ex As Exception
            Return "Unable to detect OS: " & ex.Message
        End Try

        Return "Unknown Windows Version"
    End Function

    Private Function GetRealWindowsVersion() As String
        Try
            Dim searcher As New System.Management.ManagementObjectSearcher(
            "SELECT Caption, Version FROM Win32_OperatingSystem")

            For Each os As System.Management.ManagementObject In searcher.Get()
                Dim caption As String = os("Caption").ToString().Trim()
                Dim version As String = os("Version").ToString().Trim()
                Dim lowerCaption As String = caption.ToLowerInvariant()

                If lowerCaption.Contains("windows 11") Then
                    WindowsOSVersion = "Win11"
                ElseIf lowerCaption.Contains("windows 10") Then
                    WindowsOSVersion = "Win10"
                ElseIf lowerCaption.Contains("windows 8.1") Then
                    WindowsOSVersion = "Win8.1"
                ElseIf lowerCaption.Contains("windows 8") Then
                    WindowsOSVersion = "Win8"
                ElseIf lowerCaption.Contains("windows 7") Then
                    WindowsOSVersion = "Win7"
                ElseIf lowerCaption.Contains("vista") Then
                    WindowsOSVersion = "Vista"
                ElseIf lowerCaption.Contains("xp") Then
                    WindowsOSVersion = "XP"
                ElseIf lowerCaption.Contains("server") Then
                    WindowsOSVersion = "WinServer"
                Else
                    WindowsOSVersion = "Unknown"
                End If

                Return $"{caption} (Version {version})"
            Next

        Catch ex As Exception
            WindowsOSVersion = "Unknown"
            Return "Unable to detect OS: " & ex.Message
        End Try

        WindowsOSVersion = "Unknown"
        Return "Unknown Windows Version"
    End Function

    Private Sub ShowAdminStatus()
        If IsRunningAsAdministrator() Then
            AppendToConsole("Running as Administrator: Yes")
        Else
            AppendToConsole("Running as Administrator: No")
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
            AppendToConsole(ComName & " not registered.")
        Else
            AppendToConsole("Found " & ComName & " registration:")
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
            AppendToConsole("   ERROR: " & fileName & " not found in " & Application.StartupPath)
            Return
        End If

        Try
            Dim p As New Process()
            p.StartInfo.FileName = "regsvr32.exe"
            p.StartInfo.Arguments = "/s """ & fullPath & """"
            p.StartInfo.UseShellExecute = True   ' no redirection, just run it
            ' If you want to force UAC prompt for this app instead of requiring
            ' "Run as administrator", uncomment the next line:
            ' p.StartInfo.Verb = "runas"

            p.Start()
            p.WaitForExit()

            If p.ExitCode = 0 Then
                AppendToConsole("   OK")
            Else
                AppendToConsole("   regsvr32 exit code: " & p.ExitCode)
            End If

        Catch ex As Exception
            AppendToConsole("   ERROR: " & ex.Message)
        End Try
    End Sub
    Private Sub btnRegisterLibraries_Click(sender As Object, e As EventArgs) Handles btnRegisterLibraries.Click
        DisableButtons()

        AppendToConsole("--------------------------------------------")
        AppendToConsole(" OP2Mapper2 Installation")
        AppendToConsole("--------------------------------------------")
        AppendToConsole("Welcome to the installer for OP2Mapper2")
        AppendToConsole("You may need to be logged on as a system administrator")
        AppendToConsole("if you are on Windows NT, 2000, or XP.")
        AppendToConsole("The installation may still succeed without being an administrator,")
        AppendToConsole("however, if you get errors on run, you will need to install again")
        AppendToConsole("as an administrator.")
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

        AppendToConsole("---------------------------")
        AppendToConsole(" Finishing installation...")
        AppendToConsole("---------------------------")
        AppendToConsole("")
        AppendToConsole("Click 'Run Mapper' when you’re ready.")
        AppendToConsole("")

        EnableButtons()
    End Sub


    Private Sub btnShowLibraries_Click(sender As Object, e As EventArgs) Handles btnShowLibraries.Click
        DisableButtons()

        'FindCOMRegistration("OP2Editor")

        ShowLibrariesMapper()

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
        AppendToConsole(LibraryFileCount & " COM Library Files Checked")
    End Sub

    Private Sub ShowLibrariesSystem()
        AppendToConsole("VB6 COM Registration Status:")
        FindCOMRegistration("MSCOMCTL.OCX")
        FindCOMRegistration("COMDLG32.OCX")
    End Sub

    Private Sub btnRunMapper_Click(sender As Object, e As EventArgs) Handles btnRunMapper.Click
        Try
            Dim exe = Path.Combine(Application.StartupPath, "Mapper2.exe")

            If Not File.Exists(exe) Then
                AppendToConsole("Mapper2.exe not found in folder!")
                Exit Sub
            End If

            AppendToConsole("Launching OP2Mapper2...")
            Process.Start(exe)

        Catch ex As Exception
            AppendToConsole("ERROR launching: " & ex.Message)
        End Try
    End Sub

End Class
