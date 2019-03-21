VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DefaultLogMessageFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Logger.Implementation")
Option Explicit
Implements ILogMessageFormatter

Private Function ILogMessageFormatter_FormatMessage(ByVal level As LogLevel, ByVal loggerName As String, ByVal message As String) As String
    ILogMessageFormatter_FormatMessage = Framework.Strings.Format("{0:s}\t{1}\t[{2}]\t{3}", Now, loggerName, FormatLogLevel(level), message)
End Function

Private Function FormatLogLevel(ByVal level As LogLevel) As String

    Select Case level

        Case LogLevel.DebugLevel
            FormatLogLevel = "DEBUG"

        Case LogLevel.ErrorLevel
            FormatLogLevel = "ERROR"

        Case LogLevel.FatalLevel
            FormatLogLevel = "FATAL"

        Case LogLevel.InfoLevel
            FormatLogLevel = "INFO"

        Case LogLevel.TraceLevel
            FormatLogLevel = "TRACE"

        Case LogLevel.WarnLevel
            FormatLogLevel = "WARNING"

    End Select

End Function
