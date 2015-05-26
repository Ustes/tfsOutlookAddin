Partial Class RibbonTfs
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.lblSvr = Me.Factory.CreateRibbonLabel
        Me.lblTFSInfo = Me.Factory.CreateRibbonLabel
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.cboWorkItemType = Me.Factory.CreateRibbonComboBox
        Me.cboArea = Me.Factory.CreateRibbonComboBox
        Me.cboIteration = Me.Factory.CreateRibbonComboBox
        Me.cboAssignedTo = Me.Factory.CreateRibbonComboBox
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.btnSelectProject = Me.Factory.CreateRibbonButton
        Me.btnCreate = Me.Factory.CreateRibbonButton
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.Tab1.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group3.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.btnSelectProject)
        Me.Group4.Label = "TFS Server"
        Me.Group4.Name = "Group4"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.lblSvr)
        Me.Group2.Items.Add(Me.lblTFSInfo)
        Me.Group2.Label = "TFS Info"
        Me.Group2.Name = "Group2"
        '
        'lblSvr
        '
        Me.lblSvr.Label = "TFS Server"
        Me.lblSvr.Name = "lblSvr"
        '
        'lblTFSInfo
        '
        Me.lblTFSInfo.Label = "Project name"
        Me.lblTFSInfo.Name = "lblTFSInfo"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.cboWorkItemType)
        Me.Group1.Items.Add(Me.cboArea)
        Me.Group1.Items.Add(Me.cboIteration)
        Me.Group1.Items.Add(Me.Separator1)
        Me.Group1.Items.Add(Me.cboAssignedTo)
        Me.Group1.Label = "Work Item Info"
        Me.Group1.Name = "Group1"
        '
        'cboWorkItemType
        '
        Me.cboWorkItemType.Label = "Work Item Type"
        Me.cboWorkItemType.Name = "cboWorkItemType"
        Me.cboWorkItemType.Text = Nothing
        '
        'cboArea
        '
        Me.cboArea.Label = "Area"
        Me.cboArea.Name = "cboArea"
        Me.cboArea.Text = Nothing
        '
        'cboIteration
        '
        Me.cboIteration.Label = "Iteration"
        Me.cboIteration.Name = "cboIteration"
        Me.cboIteration.Text = Nothing
        '
        'cboAssignedTo
        '
        Me.cboAssignedTo.Label = "Assigned To"
        Me.cboAssignedTo.Name = "cboAssignedTo"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.btnCreate)
        Me.Group3.Label = "Actions"
        Me.Group3.Name = "Group3"
        '
        'btnSelectProject
        '
        Me.btnSelectProject.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnSelectProject.Image = Global.TfsOutlookAddIn.My.Resources.Resources.magnifier5
        Me.btnSelectProject.Label = "Select Project"
        Me.btnSelectProject.Name = "btnSelectProject"
        Me.btnSelectProject.ShowImage = True
        '
        'btnCreate
        '
        Me.btnCreate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnCreate.Image = Global.TfsOutlookAddIn.My.Resources.Resources.writing9
        Me.btnCreate.Label = "Create Work Item"
        Me.btnCreate.Name = "btnCreate"
        Me.btnCreate.ShowImage = True
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'RibbonTfs
        '
        Me.Name = "RibbonTfs"
        Me.RibbonType = "Microsoft.Outlook.Mail.Read"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnCreate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents cboWorkItemType As Microsoft.Office.Tools.Ribbon.RibbonComboBox
    Friend WithEvents btnSelectProject As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents lblTFSInfo As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents lblSvr As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents cboArea As Microsoft.Office.Tools.Ribbon.RibbonComboBox
    Friend WithEvents cboIteration As Microsoft.Office.Tools.Ribbon.RibbonComboBox
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents cboAssignedTo As Microsoft.Office.Tools.Ribbon.RibbonComboBox
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property RibbonTfs() As RibbonTfs
        Get
            Return Me.GetRibbon(Of RibbonTfs)()
        End Get
    End Property
End Class
