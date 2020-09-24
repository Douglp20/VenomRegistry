Public Class VenomRegistry
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Public Sub New()
    End Sub
    Private Function getFullPath(KeySection As String) As String
        Dim path As String = "Software\\VB and VBA Program Settings\\circuit\" + KeySection
        Return path

    End Function
    Public Sub SaveSetting(keySection As String, keyString As String, keySetting As String)
        On Error GoTo Err

        Dim key As Microsoft.Win32.RegistryKey
        key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(getFullPath(keySection), True)

        If key Is Nothing Then
            key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(getFullPath(keySection))
        End If
        key.SetValue(keyString, keySetting)

        Exit Sub
Err:

        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Function GetSetting(keySection As String, keyString As String) As String

        On Error GoTo Err

        Dim strName As String
        Dim key As Microsoft.Win32.RegistryKey

        key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(getFullPath(keySection), True)

        If key Is Nothing Then
            strName = ""
        Else
            If key.GetValue(keyString, "").ToString = "" Then
                strName = ""
            Else
                strName = key.GetValue(keyString).ToString
            End If
        End If
        Return strName.ToString

        Exit Function
Err:

        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function GetSetting(keySection As String, keyString As String, keyDefault As String) As String

        On Error GoTo Err
        Dim strName As String
        Dim key As Microsoft.Win32.RegistryKey


        key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(getFullPath(keySection), True)

        If key Is Nothing Then
            strName = keyDefault
        Else
            If key.GetValue(keyString, "").ToString = "" Then
                strName = keyDefault
            Else
                Select Case key.GetValue(keyString).GetType.ToString
                    Case "System.String"
                        strName = key.GetValue(keyString).ToString
                    Case Else
                        strName = keyDefault
                End Select

            End If
        End If
        Return strName.ToString
        Exit Function
Err:

        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
End Class
