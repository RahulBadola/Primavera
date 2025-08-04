Imports System.Windows.Forms
Imports PO.Objects
Imports System.Data
Public Class SnapshotBrowser
    Private Shared _versionId As Integer = 0
    Private Shared _POProjId As Integer = 0
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        '' fill drop-down lists
        ddVersions.DataSource = Nothing
        ddVersions.ValueMember = "Key"
        ddVersions.DisplayMember = "Value"
        ddVersions.DataSource = New BindingSource(DataRepository.Current.ProjSnapshotversion(_POProjId), Nothing)
    End Sub
    Public Shared Function selectVersion(ByVal projectID As Integer) As Integer
        _POProjId = projectID
        '_versionId = projectID
        Dim thisForm As New SnapshotBrowser()
        thisForm.Focus()
        thisForm.StartPosition = FormStartPosition.CenterParent
        thisForm.ShowDialog()
        Return _versionId
    End Function

    'Private Sub ddVersions_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles ddVersions.SelectedValueChanged
    '    _versionId = ddVersions.SelectedValue
    'End Sub

    Private Sub btnOK_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click
        _versionId = ddVersions.SelectedValue
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        _versionId = 0
        Me.Close()
    End Sub
End Class