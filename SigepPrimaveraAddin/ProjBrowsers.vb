Imports System.Windows.Forms
Imports PO.Objects
Imports System.Data

Public Class ProjBrowser

    Private Shared _POProjId As Integer = 0
    Private _P6ProjId As Integer = 0
    Private _UserId As String ' TODO: to be used to check access to projects

    Private Sub New(p6ProjId As Integer, ByVal userId As String)


        InitializeComponent()
        Dim cfg As New ConfigHelper()
        Dim companies As Dictionary(Of Integer, String)
        Dim getLoggedUserFromSysOp As Boolean = ((New ConfigHelper()).GetConfigKey("VerifySIGEPUserByProxy") = "1")
        If getLoggedUserFromSysOp Then : _UserId = Environment.UserName
        Else : _UserId = userId
        End If
        _P6ProjId = p6ProjId
        SigepPrimaveraApp.WriteLog("Current user: " & _UserId)
        SigepPrimaveraApp.WriteLog("Current P6 project: " & _P6ProjId)
        Try

            companies = DataRepository.Current.CompanyList(_UserId)
            If companies.Count > 0 Then
                cmbCompany.DataSource = Nothing
                cmbCompany.ValueMember = "Key"
                cmbCompany.DisplayMember = "Value"
                cmbCompany.DataSource = New BindingSource(companies, Nothing)
            End If
            cmbStatus.DataSource = Nothing
            cmbStatus.ValueMember = "Key"
            cmbStatus.DisplayMember = "Value"
            cmbStatus.DataSource = New BindingSource(DataRepository.Current.ProjStatus, Nothing)

            Dim currentVersion As String = cfg.GetConfigKey("version").ToString()
            If Not String.IsNullOrEmpty(currentVersion) Then lbVersion.Text = "Version " & currentVersion
        Catch ex As Exception
            SigepPrimaveraApp.WriteLog(ex.Message)
        End Try

    End Sub
    Public Shared Function SelectProject(p6ProjId As Integer, ByVal userId As String) As Integer
        Dim thisForm As New ProjBrowser(p6ProjId, userId)
        thisForm.ShowDialog()
        thisForm.Focus()

        thisForm.StartPosition = FormStartPosition.CenterParent
        Return _POProjId
    End Function

    Private Sub OnCancelClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        _POProjId = 0
        Me.Close()
    End Sub

    Private Function GetProjectInfo(ws As PM6DBServiceController.PM6DBServiceController, ByVal projId As Integer, ByRef code As String, ByRef name As String) As Boolean

        code = String.Empty
        name = String.Empty

        Dim buf() As Byte = ws.GetSigepProjectInfo(projId)
        Dim data As New ZDataSet(buf)
        If data.Tables(0).Rows.Count > 0 Then
            Dim row As DataRow = data.Tables(0).Rows(0)
            code = row(0)
            name = row(1)
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub OnProjectsCellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grdProjects.CellDoubleClick
        Dim row As Integer = e.RowIndex
        If row < 0 Then '???
            MessageBox.Show(Me, "Please select a PO project to publish to", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        ExecuteProjectAssociation()

        '_POProjId = grdProjects.Rows(row).Cells(0).Value
        '_POProjId = SnapshotBrowser.selectVersion(_POProjId)
        'Me.Close()
    End Sub

    Private Sub OnPublishClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPublish.Click
        ExecuteProjectAssociation()
    End Sub

    Private Sub ExecuteProjectAssociation()
        If grdProjects.SelectedRows.Count = 0 Then
            MessageBox.Show(Me, "Please select a PO project to publish to", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        _POProjId = grdProjects.SelectedRows(0).Cells(0).Value

        ''Add warning message before the process start'
        ''checks if a P6 schedule is already linked to a SIGEP project different from the one the user is publishing to"
        '' Added By Indramani
        Dim cfg As New ConfigHelper()
        Dim ws As New PM6DBServiceController.PM6DBServiceController(cfg.GetConfigKey("WebServiceUrl"))
        Dim refSIGEPProj As Integer = ws.GetPM6AssociateProject(_P6ProjId)
        Dim currSigepProjectCode As String = String.Empty
        Dim currSigepProjectname As String = String.Empty
        Dim newSigepProjectCode As String = String.Empty
        Dim newSigepProjectname As String = String.Empty
        Dim result As MsgBoxResult

        If Not GetProjectInfo(ws, _POProjId, newSigepProjectCode, newSigepProjectname) Then Throw New Exception("Unknown project")
        If refSIGEPProj > 0 AndAlso _POProjId <> refSIGEPProj Then
            ' the P6 project we're coming from is already referenced in SIGEP
            ' the P6 project is linked to a SIGEP project which is not the one that has been selected for publication
            If Not GetProjectInfo(ws, refSIGEPProj, currSigepProjectCode, currSigepProjectname) Then Throw New Exception("Unknown project")
            result = MessageBox.Show(Me, "This schedule is already linked to project " & currSigepProjectCode & " - " & currSigepProjectname & " in SIGEP. Are you sure you wish to remove the existing link and link it to project project " & newSigepProjectCode & " - " & newSigepProjectname & " ?", Me.Text, MsgBoxStyle.OkCancel, MessageBoxIcon.Warning)
            If result <> MsgBoxResult.Ok Then
                _POProjId = 0
                Return ' nothing is done
            End If
        End If

        ' warning
        result = MessageBox.Show(Me, "This action will overwrite any existing data for project " & newSigepProjectCode & " - " & newSigepProjectname & ". Do you wish to continue?", Me.Text, MsgBoxStyle.OkCancel, MessageBoxIcon.Warning)
        If result <> MsgBoxResult.Ok Then
            _POProjId = 0
            Return ' nothing is done
        End If


        _POProjId = SnapshotBrowser.selectVersion(_POProjId)
        Me.Close()
    End Sub

    Private Sub OnCleanFiltersClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCleanFilters.Click

        cmbCustomer.SelectedIndex = -1
        cmbLocation.SelectedIndex = -1
        cmbStatus.SelectedIndex = -1
        txtDescription.Text = String.Empty

    End Sub

    Private Sub OnSearchClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSearch.Click

        grdProjects.DataSource = Nothing

        Dim status As Integer = -1
        Dim location As Integer = -1
        Dim cust As String = String.Empty
        Dim descr As String = txtDescription.Text.Trim()
        Dim companyId As Integer = cmbCompany.SelectedValue '' System.Configuration.ConfigurationManager.AppSettings("companyId")
        Dim cfg As New ConfigHelper()
        Dim ws As New PM6DBServiceController.PM6DBServiceController(cfg.GetConfigKey("WebServiceUrl"))


        If cmbStatus.SelectedIndex > -1 Then status = CInt(cmbStatus.SelectedValue)
        If cmbCustomer.SelectedIndex > -1 Then cust = cmbCustomer.SelectedValue
        If cmbLocation.SelectedIndex > -1 Then location = CInt(cmbLocation.SelectedValue)

        Dim buf() As Byte = ws.Projects(companyId, _UserId, descr, status, location, cust)
        If buf Is Nothing Then Return
        Dim data As New ZDataSet(buf)

        grdProjects.DataSource = data.Tables(0)
        grdProjects.Columns(0).Visible = False
        grdProjects.Refresh()
        grdProjects.ClearSelection()

        For i As Integer = 1 To grdProjects.Columns.Count - 1
            grdProjects.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        Next
    End Sub

    Private Sub OnFormLoad(sender As Object, e As EventArgs) Handles Me.Load
        Me.BringToFront()
        '' Me.Focus()
    End Sub

    Private Sub cmbCompany_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles cmbCompany.SelectedValueChanged
        '' fill drop-down lists
        cmbLocation.DataSource = Nothing
        cmbLocation.ValueMember = "Key"
        cmbLocation.DisplayMember = "Value"
        cmbLocation.DataSource = New BindingSource(DataRepository.Current.Locations, Nothing)


        cmbCustomer.DataSource = Nothing
        cmbCustomer.ValueMember = "Key"
        cmbCustomer.DisplayMember = "Value"
        cmbCustomer.DataSource = New BindingSource(DataRepository.Current.Customers, Nothing)


    End Sub
End Class