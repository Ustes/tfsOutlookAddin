Imports Microsoft.Office.Tools.Ribbon
Imports System
Imports System.Collections.ObjectModel
Imports Microsoft.TeamFoundation.Client
Imports Microsoft.TeamFoundation.Framework.Common
Imports Microsoft.TeamFoundation.Framework.Client
Imports Microsoft.TeamFoundation.Proxy
Imports Microsoft.TeamFoundation.Server
Imports Microsoft.TeamFoundation.WorkItemTracking.Client
Imports Microsoft.TeamFoundation.WorkItemTracking.WpfControls
Imports Microsoft.Office.Interop.Outlook
Imports System.Collections
Imports System.IO

Public Class RibbonTfs
    Private configurationServer As TfsConfigurationServer = Nothing
    Private currentProject As TfsTeamProjectCollection = Nothing
    Private currentWorkItemStore As WorkItemStore = Nothing
    Private currentProjectName As String
    Private currentUser As String

    Private Sub RibbonTfs_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

        Dim projUri As String = GetSetting("tfsOutlookAddin", "TFS", "Server", "")
        Dim svrName As String = GetSetting("tfsOutlookAddin", "TFS", "ServerName", "").Replace("Server : ", "")
        Dim projName As String = GetSetting("tfsOutlookAddin", "TFS", "ProjectName", "").Replace("Project : ", "")
        If Not String.IsNullOrEmpty(projUri) Then
            Dim u As New Uri(projUri)
            currentProject = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(u)
            lblSvr.Label = svrName
            lblTFSInfo.Label = projName

            lblSvr.Label = "Server : " & svrName
            lblTFSInfo.Label = "Project : " & projName


            currentWorkItemStore = currentProject.GetService(Of WorkItemStore)()
            loadOptions(projName)

        End If

        Dim item As inspector = TryCast(Me.Context, Inspector)

        currentUser = item.Application.Session.CurrentUser.Name

    End Sub

    Private Sub btnCreate_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCreate.Click
        Try
            Dim wit As String = cboWorkItemType.Text
            Dim selectedUser As String = cboAssignedTo.Text
            currentProjectName = lblTFSInfo.Label.Replace("Project :", "").Trim

            If String.IsNullOrEmpty(selectedUser) Then
                selectedUser = currentUser
            End If

            If Not String.IsNullOrEmpty(wit) And Not String.IsNullOrEmpty(currentProjectName) Then
                Dim item As inspector = TryCast(Me.Context, Inspector)
                Dim mail As MailItem = TryCast(item.CurrentItem, MailItem)

                Dim workItemTypes As WorkItemTypeCollection = currentWorkItemStore.Projects(currentProjectName).WorkItemTypes

                Dim sw As WorkItemType = workItemTypes(wit)
                Dim wi As New WorkItem(sw)
                wi.Title = mail.Subject
                wi.Description = mail.HTMLBody.ToString
              
                wi.Fields("Assigned to").Value = selectedUser


                Dim mailPath As String = String.Format("{0}{1}.msg", Path.GetTempPath, Path.GetFileNameWithoutExtension(Path.GetRandomFileName()))
                mail.SaveAs(mailPath)


                Dim wiAttach As New Microsoft.TeamFoundation.WorkItemTracking.Client.Attachment(mailPath, mail.Subject)
                wi.Attachments.Add(wiAttach)

                For Each att As Microsoft.Office.Interop.Outlook.Attachment In mail.Attachments
                    Try

                        Dim attPath As String = String.Format("{0}{1}", Path.GetTempPath, att.FileName)
                        att.SaveAsFile(attPath)
                        Dim wiAtt As New Microsoft.TeamFoundation.WorkItemTracking.Client.Attachment(attPath, att.DisplayName)
                        wi.Attachments.Add(wiAtt)
                    Catch ex As System.Exception
                        Continue For
                    End Try

                Next

                Dim errorMessage As String = ""
                For Each field As Field In wi.Fields
                    If field.Status <> FieldStatus.Valid Then
                        errorMessage = errorMessage & String.Format("{0} - {1} = {2}", field.Status.ToString, field.ReferenceName, field.Value) & Environment.NewLine
                    End If
                Next
                If Not String.IsNullOrEmpty(errorMessage) Then
                    MsgBox(errorMessage)
                    Exit Sub
                End If

                Dim errs As ArrayList = wi.Validate
                If errs.Count = 0 Then
                    Try
                        wi.Save()
                        Dim wiUrl As String = String.Format("{0}", wi.Uri.PathAndQuery)
                        'System.Diagnostics.Process.Start(wiUrl)
                        MsgBox(String.Format("Work Item #{0} Successfully Created!!!", wi.Id.ToString))

                    Catch ex As TeamFoundationPropertyValidationException
                        MsgBox(ex.Message.ToString)
                    End Try

                Else
                    Dim errMsg As New StringBuilder
                    For Each fld As Field In errs
                        errMsg.AppendLine(String.Format("{0}", fld.ReferenceName) & Environment.NewLine)
                    Next
                    MsgBox("Validation Errors: " & errMsg.ToString)
                End If

            End If

        Catch ex As System.Exception
            MsgBox(ex.Message.ToString, vbOKOnly, "tfsOutlookAddin")
        End Try
      
    End Sub

    Private Sub btnSelectProject_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSelectProject.Click

        'If IsNothing(currentProject) Then
        Dim dp As New TeamProjectPicker(TeamProjectPickerMode.SingleProject, False)
        dp.ShowDialog()
        For Each p As ProjectInfo In dp.SelectedProjects
            currentProjectName = p.Name
            Exit For
        Next
        currentProject = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(dp.SelectedTeamProjectCollection.Uri)

        lblSvr.Label = "Server : " & currentProject.ToString
        lblTFSInfo.Label = "Project : " & currentProjectName
        SaveSetting("tfsOutlookAddin", "TFS", "ServerName", lblSvr.Label)
        SaveSetting("tfsOutlookAddin", "TFS", "ProjectName", lblTFSInfo.Label)
        SaveSetting("tfsOutlookAddin", "TFS", "Server", currentProject.Uri.ToString)
        'End If

        currentWorkItemStore = currentProject.GetService(Of WorkItemStore)()


        loadOptions(currentProjectName)



    End Sub

    Private Sub loadOptions(projectName As String)
        Dim newrib As RibbonDropDownItem

        cboWorkItemType.Items.Clear()
        cboArea.Items.Clear()
        cboIteration.Items.Clear()
        For Each w As WorkItemType In currentWorkItemStore.Projects(projectName).WorkItemTypes
            newrib = Me.Factory.CreateRibbonDropDownItem()
            newrib.Label = w.Name
            cboWorkItemType.Items.Add(newrib)
        Next


        For Each w As Node In currentWorkItemStore.Projects(projectName).AreaRootNodes
            newrib = Me.Factory.CreateRibbonDropDownItem()
            newrib.Label = w.Path
            cboArea.Items.Add(newrib)
            For Each c As Node In w.ChildNodes
                newrib = Me.Factory.CreateRibbonDropDownItem()
                newrib.Label = c.Path
                cboArea.Items.Add(newrib)
            Next
        Next



        For Each w As Node In currentWorkItemStore.Projects(projectName).IterationRootNodes
            newrib = Me.Factory.CreateRibbonDropDownItem()
            newrib.Label = w.Path
            cboIteration.Items.Add(newrib)
            For Each c As Node In w.ChildNodes
                newrib = Me.Factory.CreateRibbonDropDownItem()
                newrib.Label = c.Path
                cboIteration.Items.Add(newrib)
            Next
        Next




        Dim identityManagementService As Microsoft.TeamFoundation.Framework.Client.IIdentityManagementService = currentProject.GetService(Of IIdentityManagementService)()
        Dim collectionWideValidUsers As TeamFoundationIdentity = identityManagementService.ReadIdentity(IdentitySearchFactor.DisplayName,
                                                                              "Project Collection Valid Users",
                                                                              MembershipQuery.Expanded,
                                                                              ReadIdentityOptions.None)

        Dim validMembers As TeamFoundationIdentity() = identityManagementService.ReadIdentities(collectionWideValidUsers.Members,
                                                                    MembershipQuery.Expanded,
                                                                    ReadIdentityOptions.ExtendedProperties)

        Dim memberNames As List(Of String) = validMembers.Where(Function(t As TeamFoundationIdentity) Not t.IsContainer _
                                                                    And Not t.DisplayName.Contains("Microsoft") _
                                                                    And Not t.DisplayName.Contains("TeamFoundationService") _
                                                                    And t.Descriptor.IdentityType <> "Microsoft.TeamFoundation.UnauthenticatedIdentity" _
                                                                    ).Select(Function(t) t.DisplayName).ToList()

        For Each m As String In memberNames
            newrib = Me.Factory.CreateRibbonDropDownItem()
            newrib.Label = m
            cboAssignedTo.Items.Add(newrib)
        Next



        If cboArea.Items.Count = 0 Then
            newrib = Me.Factory.CreateRibbonDropDownItem()
            newrib.Label = projectName
            cboArea.Items.Add(newrib)
        End If

        If cboIteration.Items.Count = 0 Then
            newrib = Me.Factory.CreateRibbonDropDownItem()
            newrib.Label = projectName
            cboIteration.Items.Add(newrib)
        End If
        If cboIteration.Items.Count = 0 Then
            newrib = Me.Factory.CreateRibbonDropDownItem()
            newrib.Label = currentUser
            cboAssignedTo.Items.Add(newrib)
        End If





    End Sub

End Class