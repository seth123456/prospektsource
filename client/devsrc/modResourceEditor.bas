Attribute VB_Name = "modResourceEditor"
Option Explicit

Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public TempResource As ResourceRec
Public MapResource() As MapResourceRec

Public Type ResourceRec
    name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    sound As String * NAME_LENGTH
    
    ResourceType As Byte
    ResourceImage As Long
    ExhaustedImage As Long
    ItemReward As Long
    ToolRequired As Long
    health As Long
    RespawnTime As Long
    Animation As Long
    Effect As Long
End Type

Private Type MapResourceRec
    X As Long
    Y As Long
    ResourceState As Byte
End Type

' ////////////////
' // Resource Editor //
' ////////////////
Public Sub ResourceEditorInit()
Dim I As Long
Dim SoundSet As Boolean

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    If frmEditor_Resource.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Resource.lstIndex.ListIndex + 1

    ' add the array to the combo
    frmEditor_Resource.cmbSound.Clear
    frmEditor_Resource.cmbSound.AddItem "None."
    For I = 1 To UBound(soundCache)
        frmEditor_Resource.cmbSound.AddItem soundCache(I)
    Next
    ' finished populating
    
    With frmEditor_Resource
        .scrlExhaustedPic.Max = Count_Resource
        .scrlNormalPic.Max = Count_Resource
        .scrlAnimation.Max = MAX_ANIMATIONS
        .scrlReward.Max = MAX_ITEMS
        .scrlEffect.Max = MAX_EFFECTS
        
        .txtName.Text = Trim$(Resource(EditorIndex).name)
        .txtMessage.Text = Trim$(Resource(EditorIndex).SuccessMessage)
        .txtMessage2.Text = Trim$(Resource(EditorIndex).EmptyMessage)
        .cmbType.ListIndex = Resource(EditorIndex).ResourceType
        .scrlNormalPic.Value = Resource(EditorIndex).ResourceImage
        .scrlExhaustedPic.Value = Resource(EditorIndex).ExhaustedImage
        .scrlReward.Value = Resource(EditorIndex).ItemReward
        .scrlTool.Value = Resource(EditorIndex).ToolRequired
        .txtHealth.Text = Resource(EditorIndex).health
        .scrlRespawn.Value = Resource(EditorIndex).RespawnTime
        .scrlAnimation.Value = Resource(EditorIndex).Animation
        .scrlEffect.Value = Resource(EditorIndex).Effect
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For I = 0 To .cmbSound.ListCount
                If .cmbSound.List(I) = Trim$(Resource(EditorIndex).sound) Then
                    .cmbSound.ListIndex = I
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
    End With
    
    Resource_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub ResourceEditorOk()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    For I = 1 To MAX_RESOURCES
        If Resource_Changed(I) Then
            Call SendSaveResource(I)
        End If
    Next
    
    Unload frmEditor_Resource
    Editor = 0
    ClearChanged_Resource
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub ResourceEditorCancel()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Resource
    ClearChanged_Resource
    ClearResources
    SendRequestResources
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Resource()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    ZeroMemory Resource_Changed(1), MAX_RESOURCES * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Resource", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearResource(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearResource", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearResources()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    For I = 1 To MAX_RESOURCES
        Call ClearResource(I)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
