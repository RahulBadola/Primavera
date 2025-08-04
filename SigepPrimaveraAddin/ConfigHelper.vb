Imports System.Configuration
Imports System.Reflection
Imports System.IO

Public Class ConfigHelper

    Private _FileMap As ExeConfigurationFileMap
    Private _ConfigFile As String
    Private _Config As System.Configuration.Configuration

    Public Sub New()

        _ConfigFile = System.Reflection.Assembly.GetExecutingAssembly().Location & ".config"
        _FileMap = New ExeConfigurationFileMap()
        _FileMap.ExeConfigFilename = _ConfigFile
        _Config = ConfigurationManager.OpenMappedExeConfiguration(_FileMap, ConfigurationUserLevel.None)

    End Sub
    Public Function GetConfigKey(key As String) As String

        Return _Config.AppSettings.Settings(key).Value

    End Function


End Class
