Attribute VB_Name = "MRSAPUtilities"
Option Explicit
' /////////////////////////////////////////////////////////////////////////////////////////////////
' Utilities procedures to use and re-use in other Main procedures.
' RSAP = Robot Structural Analysis Professional
'
' Developer: Nuno Teixeira
' Email: teixeiranh@gmail.com
' /////////////////////////////////////////////////////////////////////////////////////////////////
'
Private robApp As RobotApplication
Private RobotNodes As RobotNodeServer
Private RLS As RobotNodeRigidLinkServer
Private NodeSel As RobotSelection
Private RLdata As RobotNodeRigidLinkData
Private Label As RobotLabel
Private AllNodesCol As RobotNodeCollection
Private I_PT_FRAME As IRobotActiveModelType

' /////////////////////////////////////////////////////////////////////////////////////////////////
' Verify if RSAP is opened.
' /////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Sub VerifyIfRobotIsOpened()
    Set robApp = New RobotApplication
    If Not robApp.Visible Then
        Set robApp = Nothing
        MsgBox "Start Robot and Load Model!", vbOKOnly, "Error"
        End
    Else
        If (robApp.Project.Type <> I_PT_FRAME) And _
        (robApp.Project.Type <> I_PT_SHELL) And _
        (robApp.Project.Type <> I_PT_BUILDING) Then
            MsgBox "Structure type should be Frame3D or Shell or Building!", vbOKOnly, "Error"
            End
        End If
    End If
End Sub

' /////////////////////////////////////////////////////////////////////////////////////////////////
' Verify if we have any node created in project.
' /////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Sub VerifyIfNodesWereCreated()
    Set robApp = New RobotApplication
    Set AllNodesCol = robApp.Project.structure.Nodes.GetAll
    If AllNodesCol.Count = 0 Then
        MsgBox "Please create nodes in Robot!", vbOKOnly, "Error!"
        End
    End If
End Sub

' /////////////////////////////////////////////////////////////////////////////////////////////////
' Verify if there are any nodes currently selected.
' /////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Sub VerifyIfNodesAreSelected()
    Set robApp = New RobotApplication
    Set NodeSel = robApp.Project.structure.Selections.Get(I_OT_NODE)
    If NodeSel.Count = 0 Then
        MsgBox "Please select nodes in Robot!", vbOKOnly, "Error!"
        End
    End If
End Sub

' /////////////////////////////////////////////////////////////////////////////////////////////////
' Clears all the references.
' /////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Sub CleanVariables()
    Set robApp = Nothing
    Set AllNodesCol = Nothing
    Set RLS = Nothing
    Set RLdata = Nothing
    Set RobotNodes = Nothing
    Set NodeSel = Nothing
End Sub

' /////////////////////////////////////////////////////////////////////////////////////////////////
' Set the states for RSAP.
' /////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Sub SetStates()
    Set robApp = New RobotApplication
    robApp.Visible = True
    robApp.Interactive = 1
    robApp.UserControl = True
End Sub

' /////////////////////////////////////////////////////////////////////////////////////////////////
' Rigid link label creator.
' /////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Sub RigidLinkLabelCreate()
    Set Label = robApp.Project.structure.labels.Create(I_LT_NODE_RIGID_LINK, "rl_fixed")
    Set RLdata = Label.Data
    RLdata.UX = True
    RLdata.UY = True
    RLdata.UZ = True
    RLdata.RX = True
    RLdata.RY = True
    RLdata.RZ = True
    robApp.Project.structure.labels.Store Label
End Sub




