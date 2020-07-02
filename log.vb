Imports EHICOM
Imports System.IO
Imports System.Text

Public Class Logging
    Private DoLogging As Boolean = False

    Private _LogPath As String
    Public Property LogPath() As String
        Get
            Return _LogPath
        End Get
        Set(ByVal value As String)
            _LogPath = value
        End Set
    End Property

    Private _LogLevelSetting As Integer
    Public Property LogLevelSetting() As Integer
        Get
            Return _LogLevelSetting
        End Get
        Set(ByVal value As Integer)
            _LogLevelSetting = value
        End Set
    End Property

    Private _LogFileFormat As String
    Public Property LogFileFormat() As String
        Get
            Return _LogFileFormat
        End Get
        Set(ByVal value As String)
            _LogFileFormat = value
        End Set
    End Property


    Public Sub New(ByRef comApp As EHICOM.Application)
        Try

            Dim eimodule As String = Nothing
            Select Case comApp.ModuleType
                Case EIModuleType.eiUnknownModule
                    eimodule = "eiUnknownModule"
                Case EIModuleType.eiScan
                    eimodule = "Scan"
                Case EIModuleType.eiInterpret
                    eimodule = "Interpret"
                Case EIModuleType.eiVerify
                    eimodule = "Verify"
                Case EIModuleType.eiTransfer
                    eimodule = "Transfer"
                Case EIModuleType.eiOptimizer
                    eimodule = "Optimizer"
                Case EIModuleType.eiManager
                    eimodule = "Manager"
            End Select



            Dim cLogPath As String = comApp.Configuration(EIConfigurationScope.eiGlobalSiteScope).Segment("ENCOMA_Logging").KeyValue("LogPath")

            If cLogPath.Length > 0 Then
                LogPath = cLogPath & "\" & eimodule & "\"
            End If

            LogLevelSetting = comApp.Configuration(EIConfigurationScope.eiGlobalSiteScope).Segment("ENCOMA_Logging").KeyValue("LogLevel")
            LogFileFormat = comApp.Configuration(EIConfigurationScope.eiGlobalSiteScope).Segment("ENCOMA_Logging").KeyValue("LogFileFormat")

        Catch ex As Exception
            'MsgBox("Logging Exception: " & ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' Write log entry on file system or in eventlog. Settings can be made in eiglobal -> section ENCOMA_Logging
    ''' </summary>
    ''' <param name="LogText"></param>
    ''' <param name="LogLevel">1=Exceptions // 2=Error // 3=Information // 4=Debug</param>
    ''' <remarks>0 = nothing</remarks>
    Public Sub Write(ByVal LogText As String, Optional ByVal LogLevel As Integer = 4)
        If LogLevel Then

        End If
        Dim CurrentLog As String = String.Empty
        If LogLevelSetting >= LogLevel And LogLevel > 0 Then
            Dim LevelType As String = String.Empty
            Select Case LogLevel
                Case "1"
                    LevelType = "Exc // "
                Case "2"
                    LevelType = "Err // "
                Case "3"
                    LevelType = "Info // "
                Case "4"
                    LevelType = "Debug // "
            End Select

            Try
                CurrentLog = LogPath & DateTime.Now.ToString(LogFileFormat) & ".log"
                If Not Directory.Exists(Path.GetDirectoryName(CurrentLog)) Then
                    Directory.CreateDirectory(Path.GetDirectoryName(CurrentLog))
                End If
            Catch ex As Exception
            End Try


            Try
                Using writer As New IO.StreamWriter(CurrentLog, True, Encoding.Unicode)
                    Dim CurrentTime As String = DateTime.Now.ToString("ddMMyyy HH:mm:ss.ffff")
                    writer.WriteLine(CurrentTime & " // Host:" & Environment.MachineName.ToString() & " // " & LevelType & LogText)
                End Using
            Catch ex As Exception
            End Try

        End If
    End Sub
End Class
