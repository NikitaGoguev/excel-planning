VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} QuarterPlanTaskCommentForm
   Caption         =   "Task comment options"
   ClientHeight    =   5885
   ClientLeft      =   90
   ClientTop       =   425
   ClientWidth     =   8220.001
   OleObjectBlob   =   "QuarterPlanTaskCommentForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "QuarterPlanTaskCommentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mAccepted As Boolean
Private mArtifactOptionCount As Long
Private mAdjacentOptionCount As Long

Public Sub InitializeOptions(ByVal artifactCaption As String, ByVal artifactOptions As Collection, ByVal adjacentCaption As String, ByVal adjacentOptions As Collection)
    Me.Caption = TaskEstimateTextCommentOptionsTitle()
    fraArtifacts.Caption = artifactCaption
    fraAdjacentTeams.Caption = adjacentCaption
    cmdOk.Caption = TaskEstimateTextOk()
    cmdCancel.Caption = TaskEstimateTextCancel()
    fraArtifacts.Font.Name = DESIGN_FONT_NAME
    fraArtifacts.Font.Size = DESIGN_FONT_SIZE
    fraAdjacentTeams.Font.Name = DESIGN_FONT_NAME
    fraAdjacentTeams.Font.Size = DESIGN_FONT_SIZE
    cmdOk.Font.Name = DESIGN_FONT_NAME
    cmdOk.Font.Size = DESIGN_FONT_SIZE
    cmdCancel.Font.Name = DESIGN_FONT_NAME
    cmdCancel.Font.Size = DESIGN_FONT_SIZE
    mAccepted = False
    AddOptionsToFrame fraArtifacts, artifactOptions, "a", mArtifactOptionCount
    AddOptionsToFrame fraAdjacentTeams, adjacentOptions, "n", mAdjacentOptionCount
End Sub

Private Sub AddOptionsToFrame(ByVal targetFrame As Object, ByVal options As Collection, ByVal namePrefix As String, ByRef optionCount As Long)
    Dim index As Long
    Dim checkbox As Object
    Dim optionText As String
    Dim nextTop As Double

    optionCount = 0
    nextTop = 12
    For index = 1 To options.Count
        optionText = Trim$(CStr(options.Item(index)))
        If optionText <> "" Then
            optionCount = optionCount + 1
            Set checkbox = targetFrame.Controls.Add("Forms.CheckBox.1", "qpcb_" & namePrefix & "_" & CStr(optionCount), True)
            checkbox.Caption = optionText
            checkbox.Tag = optionText
            checkbox.Left = 8
            checkbox.Top = nextTop
            checkbox.Width = targetFrame.Width - 24
            checkbox.Height = 18
            checkbox.Value = False
            checkbox.Font.Name = DESIGN_FONT_NAME
            checkbox.Font.Size = DESIGN_FONT_SIZE
            nextTop = nextTop + 21
        End If
    Next index
    targetFrame.ScrollHeight = nextTop + 4
    targetFrame.ScrollTop = 0
End Sub

Public Property Get Accepted() As Boolean
    Accepted = mAccepted
End Property

Public Property Get ArtifactOptionCount() As Long
    ArtifactOptionCount = mArtifactOptionCount
End Property

Public Property Get AdjacentOptionCount() As Long
    AdjacentOptionCount = mAdjacentOptionCount
End Property

Public Function SelectedValues() As Collection
    Dim result As New Collection
    AddSelectedValues fraArtifacts, result
    AddSelectedValues fraAdjacentTeams, result
    Set SelectedValues = result
End Function

Private Sub AddSelectedValues(ByVal targetFrame As Object, ByVal result As Collection)
    Dim control As Object

    For Each control In targetFrame.Controls
        If TypeName(control) = "CheckBox" Then
            If CBool(control.Value) Then result.Add CStr(control.Tag)
        End If
    Next control
End Sub

Public Sub SelectOptionForTest(ByVal optionText As String)
    If SelectOptionInFrameForTest(fraArtifacts, optionText) Then Exit Sub
    If SelectOptionInFrameForTest(fraAdjacentTeams, optionText) Then Exit Sub
    Err.Raise 9, "SelectOptionForTest", "Option was not found: " & optionText
End Sub

Private Function SelectOptionInFrameForTest(ByVal targetFrame As Object, ByVal optionText As String) As Boolean
    Dim control As Object

    For Each control In targetFrame.Controls
        If TypeName(control) = "CheckBox" Then
            If StrComp(Trim$(CStr(control.Tag)), Trim$(optionText), vbTextCompare) = 0 Then
                control.Value = True
                SelectOptionInFrameForTest = True
                Exit Function
            End If
        End If
    Next control
End Function

Public Sub AcceptForTest()
    mAccepted = True
End Sub

Public Sub CancelForTest()
    mAccepted = False
End Sub

Private Sub cmdOk_Click()
    mAccepted = True
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    mAccepted = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        mAccepted = False
        Me.Hide
    End If
End Sub
