Attribute VB_Name = "modEventEditor"
Option Explicit

Public Events(1 To MAX_EVENTS) As EventWrapperRec
Public TempEvent As EventWrapperRec
Public Switches(1 To MAX_SWITCHES) As String
Public Variables(1 To MAX_VARIABLES) As String

Public Type SubEventRec
    Type As EventType
    HasText As Boolean
    Text() As String
    HasData As Boolean
    Data() As Long
End Type

Public Type EventWrapperRec
    name As String
    chkSwitch As Byte
    chkVariable As Byte
    chkHasItem As Byte
    
    SwitchIndex As Long
    SwitchCompare As Byte
    VariableIndex As Long
    VariableCompare As Byte
    VariableCondition As Long
    HasItemIndex As Long
    
    HasSubEvents As Boolean
    SubEvents() As SubEventRec
    
    Trigger As Byte
    WalkThrought As Byte
    Animated As Byte
    Graphic(0 To 2) As Long
    Layer As Byte
End Type

' /////////////////
' // Event Editor //
' /////////////////
Public Sub Events_ClearChanged()
    Dim I As Long
    For I = 1 To MAX_EVENTS
        Event_Changed(I) = False
    Next I
End Sub

Public Sub EventEditorInit()
Dim I As Long
    With frmEditor_Events
        If .Visible = False Then Exit Sub
        .scrlOpenShop.Max = MAX_SHOPS
        .scrlGiveItemID.Max = MAX_ITEMS
        .scrlPlayAnimationAnim.Max = MAX_ANIMATIONS
        .scrlWarpMap.Max = MAX_MAPS
        .scrlMessageSprite.Max = Count_Char
        .scrlGraphic.Max = Count_Event
        
        .cmbHasItem.Clear
        .cmbBranchItem.Clear
        For I = 1 To MAX_ITEMS
            .cmbHasItem.AddItem Trim$(Item(I).name)
            .cmbBranchItem.AddItem Trim$(Item(I).name)
        Next
        .cmbHasItem.ListIndex = 0
        .cmbBranchItem.ListIndex = 0
        
        .cmbSwitch.Clear
        .cmbPlayerSwitch.Clear
        .cmbBranchSwitch.Clear
        For I = 1 To MAX_SWITCHES
            .cmbSwitch.AddItem I & ". " & Switches(I)
            .cmbPlayerSwitch.AddItem I & ". " & Switches(I)
            .cmbBranchSwitch.AddItem I & ". " & Switches(I)
        Next
        .cmbSwitch.ListIndex = 0
        .cmbPlayerSwitch.ListIndex = 0
        .cmbBranchSwitch.ListIndex = 0
        
        .cmbVariable.Clear
        .cmbPlayerVar.Clear
        .cmbBranchVar.Clear
        For I = 1 To MAX_VARIABLES
            .cmbVariable.AddItem I & ". " & Variables(I)
            .cmbPlayerVar.AddItem I & ". " & Variables(I)
            .cmbBranchVar.AddItem I & ". " & Variables(I)
        Next
        .cmbVariable.ListIndex = 0
        .cmbPlayerVar.ListIndex = 0
        .cmbBranchVar.ListIndex = 0
        
        .cmbBranchClass.Clear
        .cmbChangeClass.Clear
        For I = 1 To MAX_CLASSES
            .cmbBranchClass.AddItem Trim$(Class(I).name)
            .cmbChangeClass.AddItem Trim$(Class(I).name)
        Next
        .cmbBranchClass.ListIndex = 0
        .cmbChangeClass.ListIndex = 0
        
        .cmbBranchSkill.Clear
        .cmbChangeSkills.Clear
        For I = 1 To MAX_SPELLS
            .cmbBranchSkill.AddItem Trim$(Spell(I).name)
            .cmbChangeSkills.AddItem Trim$(Spell(I).name)
        Next
        .cmbBranchSkill.ListIndex = 0
        .cmbChangeSkills.ListIndex = 0
        
        .cmbChatBubbleTarget.Clear
        For I = 1 To MAX_MAP_NPCS
            .cmbChatBubbleTarget.AddItem CStr(I) & ". "
        Next
        .cmbChatBubbleTarget.ListIndex = 0
        
        .cmbPlaySound.Clear
        For I = 1 To UBound(soundCache)
            .cmbPlaySound.AddItem (soundCache(I))
        Next
        .cmbPlaySound.ListIndex = 0
        
        .cmbPlayBGM.Clear
        For I = 1 To UBound(musicCache)
            .cmbPlayBGM.AddItem (musicCache(I))
        Next
        .cmbPlayBGM.ListIndex = 0
        
        EditorIndex = .lstIndex.ListIndex + 1
        .txtName = Trim$(Events(EditorIndex).name)
        .chkPlayerSwitch.Value = Events(EditorIndex).chkSwitch
        .chkPlayerVar.Value = Events(EditorIndex).chkVariable
        .chkHasItem.Value = Events(EditorIndex).chkHasItem
        .cmbPlayerSwitch.ListIndex = Events(EditorIndex).SwitchIndex
        .cmbPlayerSwitchCompare.ListIndex = Events(EditorIndex).SwitchCompare
        .cmbPlayerVar.ListIndex = Events(EditorIndex).VariableIndex
        .cmbPlayerVarCompare.ListIndex = Events(EditorIndex).VariableCompare
        .txtPlayerVariable.Text = Events(EditorIndex).VariableCondition
        .cmbHasItem.ListIndex = Events(EditorIndex).HasItemIndex - 1
        .cmbTrigger.ListIndex = Events(EditorIndex).Trigger
        .chkWalkthrought.Value = Events(EditorIndex).WalkThrought
        .chkAnimated.Value = Events(EditorIndex).Animated
        .scrlGraphic.Value = Events(EditorIndex).Graphic(0)
        .cmbLayer.ListIndex = Events(EditorIndex).Layer
        Call .PopulateSubEventList
    End With
    Event_Changed(EditorIndex) = True
End Sub

Public Sub EventEditorOk()
Dim I As Long
    For I = 1 To MAX_EVENTS
        If Event_Changed(I) Then
            Call Events_SendSaveEvent(I)
        End If
    Next I
    
    Unload frmEditor_Events
    Events_ClearChanged
    Editor = 0
End Sub

Public Sub EventEditorCancel()
    Editor = 0
    Unload frmEditor_Events
    Events_ClearChanged
    ClearEvents
    Events_SendRequestEventsData
End Sub

' *********************
' ** Event Utilities **
' *********************
Public Function GetSubEventCount(ByVal Index As Long)
    GetSubEventCount = 0
    If Index <= 0 Or Index > MAX_EVENTS Then Exit Function
    If Events(Index).HasSubEvents Then
        GetSubEventCount = UBound(Events(Index).SubEvents)
    End If
End Function

Public Sub ClearEvents()
    Dim I As Long
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    For I = 1 To MAX_EVENTS
        Call ClearEvent(I)
    Next I
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearEvents", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearEvent(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If Index <= 0 Or Index > MAX_EVENTS Then Exit Sub
    
    Call ZeroMemory(ByVal VarPtr(Events(Index)), LenB(Events(Index)))
    Events(Index).name = vbNullString
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearEvent", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

' *****************
' ** Event Logic **
' *****************
Public Sub Events_SetSubEventType(ByVal EIndex As Long, ByVal SIndex As Long, ByVal EType As EventType)
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    If SIndex < LBound(Events(EIndex).SubEvents) Or SIndex > UBound(Events(EIndex).SubEvents) Then Exit Sub
    
    'We are ok, allocate
    With Events(EIndex).SubEvents(SIndex)
        .Type = EType
        Select Case .Type
            Case Evt_Message
                .HasText = True
                .HasData = True
                ReDim Preserve .Text(1 To 1)
                ReDim Preserve .Data(1 To 2)
            Case Evt_Menu
                If Not .HasText Then ReDim .Text(1 To 2)
                If UBound(.Text) < 2 Then ReDim Preserve .Text(1 To 2)
                If Not .HasData Then ReDim .Data(1 To 1)
                .HasText = True
                .HasData = True
            Case Evt_OpenShop
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_GOTO
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_GiveItem
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 3)
            Case Evt_PlayAnimation
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 3)
            Case Evt_Warp
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 4)
            Case Evt_Switch
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 2)
            Case Evt_Variable
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 4)
            Case Evt_AddText
                .HasText = True
                .HasData = True
                ReDim Preserve .Text(1 To 1)
                ReDim Preserve .Data(1 To 2)
            Case Evt_Chatbubble
                .HasText = True
                .HasData = True
                ReDim Preserve .Text(1 To 1)
                ReDim Preserve .Data(1 To 2)
            Case Evt_Branch
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 6)
            Case Evt_ChangeSkill
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 2)
            Case Evt_ChangeLevel
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 2)
            Case Evt_ChangeSprite
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_ChangePK
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_ChangeClass
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_ChangeSex
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_ChangeExp
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 2)
            Case Evt_SetAccess
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_CustomScript
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_OpenEvent
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 4)
            Case Evt_ChangeGraphic
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 4)
            Case Evt_ChangeVitals
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 3)
            Case Evt_PlaySound
                .HasText = True
                .HasData = False
                Erase .Data
                ReDim Preserve .Text(1 To 1)
            Case Evt_PlayBGM
                .HasText = True
                .HasData = False
                Erase .Data
                ReDim Preserve .Text(1 To 1)
            Case Evt_SpecialEffect
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 5)
            Case Else
                .HasText = False
                .HasData = False
                Erase .Text
                Erase .Data
        End Select
    End With
End Sub


Public Function GetComparisonOperatorName(ByVal opr As ComparisonOperator) As String
    Select Case opr
        Case GEQUAL
            GetComparisonOperatorName = ">="
            Exit Function
        Case LEQUAL
            GetComparisonOperatorName = "<="
            Exit Function
        Case GREATER
            GetComparisonOperatorName = ">"
            Exit Function
        Case LESS
            GetComparisonOperatorName = "<"
            Exit Function
        Case EQUAL
            GetComparisonOperatorName = "="
            Exit Function
        Case NOTEQUAL
            GetComparisonOperatorName = "><"
            Exit Function
    End Select
    GetComparisonOperatorName = "Unknown"
End Function

Public Function GetEventTypeName(ByVal EventIndex As Long, SubIndex As Long) As String
Dim evtType As EventType
evtType = Events(EventIndex).SubEvents(SubIndex).Type
    Select Case evtType
        Case Evt_Message
            GetEventTypeName = "@Show Message: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "'"
            Exit Function
        Case Evt_Menu
            GetEventTypeName = "@Show Choices"
            Exit Function
        Case Evt_Quit
            GetEventTypeName = "@Exit Event"
            Exit Function
        Case Evt_OpenShop
            If Events(EventIndex).SubEvents(SubIndex).Data(1) > 0 Then
                GetEventTypeName = "@Open Shop: " & Events(EventIndex).SubEvents(SubIndex).Data(1) & "-" & Trim$(Shop(Events(EventIndex).SubEvents(SubIndex).Data(1)).name)
            Else
                GetEventTypeName = "@Open Shop: " & Events(EventIndex).SubEvents(SubIndex).Data(1) & "- None "
            End If
            Exit Function
        Case Evt_OpenBank
            GetEventTypeName = "@Open Bank"
            Exit Function
        Case Evt_GiveItem
            GetEventTypeName = "@Change Item"
            Exit Function
        Case Evt_ChangeLevel
            GetEventTypeName = "@Change Level"
            Exit Function
        Case Evt_PlayAnimation
            If Events(EventIndex).SubEvents(SubIndex).Data(1) > 0 Then
                GetEventTypeName = "@Play Animation: " & Events(EventIndex).SubEvents(SubIndex).Data(1) & "." & Trim$(Animation(Events(EventIndex).SubEvents(SubIndex).Data(1)).name) & " {" & Events(EventIndex).SubEvents(SubIndex).Data(2) & ", " & Events(EventIndex).SubEvents(SubIndex).Data(3) & "}"
            Else
                GetEventTypeName = "@Play Animation: None {" & Events(EventIndex).SubEvents(SubIndex).Data(2) & ", " & Events(EventIndex).SubEvents(SubIndex).Data(3) & "}"
            End If
            Exit Function
        Case Evt_Warp
            GetEventTypeName = "@Warp to: " & Events(EventIndex).SubEvents(SubIndex).Data(1) & " {" & Events(EventIndex).SubEvents(SubIndex).Data(2) & ", " & Events(EventIndex).SubEvents(SubIndex).Data(3) & "}"
            Exit Function
        Case Evt_GOTO
            GetEventTypeName = "@GoTo: " & Events(EventIndex).SubEvents(SubIndex).Data(1)
            Exit Function
        Case Evt_Switch
            If Events(EventIndex).SubEvents(SubIndex).Data(2) = 1 Then
                GetEventTypeName = "@Change Switch: " & Events(EventIndex).SubEvents(SubIndex).Data(1) + 1 & "." & Switches(Events(EventIndex).SubEvents(SubIndex).Data(1) + 1) & " = True"
            Else
                GetEventTypeName = "@Change Switch: " & Events(EventIndex).SubEvents(SubIndex).Data(1) + 1 & "." & Switches(Events(EventIndex).SubEvents(SubIndex).Data(1) + 1) & " = False"
            End If
            Exit Function
        Case Evt_Variable
            GetEventTypeName = "@Change Variable: "
            Exit Function
        Case Evt_AddText
            Select Case Events(EventIndex).SubEvents(SubIndex).Data(2)
                Case 0: GetEventTypeName = "@Add text: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "' {" & GetColorString(Events(EventIndex).SubEvents(SubIndex).Data(1)) & ", Player}"
                Case 1: GetEventTypeName = "@Add text: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "' {" & GetColorString(Events(EventIndex).SubEvents(SubIndex).Data(1)) & ", Map}"
                Case 2: GetEventTypeName = "@Add text: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "' {" & GetColorString(Events(EventIndex).SubEvents(SubIndex).Data(1)) & ", Global}"
            End Select
            Exit Function
        Case Evt_Chatbubble
            GetEventTypeName = "@Show chatbubble"
            Exit Function
        Case Evt_Branch
            GetEventTypeName = "@Conditional branch"
            Exit Function
        Case Evt_ChangeSkill
            GetEventTypeName = "@Change Spells"
            Exit Function
        Case Evt_ChangeSprite
            GetEventTypeName = "@Change Sprite: " & Events(EventIndex).SubEvents(SubIndex).Data(1)
            Exit Function
        Case Evt_ChangePK
            Select Case Events(EventIndex).SubEvents(SubIndex).Data(1)
                Case 0: GetEventTypeName = "@Change PK: NO"
                Case 1: GetEventTypeName = "@Change PK: YES"
            End Select
            Exit Function
        Case Evt_ChangeClass
            If Events(EventIndex).SubEvents(SubIndex).Data(1) > 0 Then
                GetEventTypeName = "@Change Class: " & Events(EventIndex).SubEvents(SubIndex).Data(1)
            Else
                GetEventTypeName = "@Change Class: None"
            End If
            Exit Function
        Case Evt_ChangeSex
            Select Case Events(EventIndex).SubEvents(SubIndex).Data(1)
                Case 0: GetEventTypeName = "@Change Sex: MALE"
                Case 1: GetEventTypeName = "@Change Sex: FEMALE"
            End Select
            Exit Function
        Case Evt_ChangeExp
            GetEventTypeName = "@Change Exp"
            Exit Function
        Case Evt_SetAccess
            GetEventTypeName = "@Set Access: " & Events(EventIndex).SubEvents(SubIndex).Data(1)
            Exit Function
        Case Evt_CustomScript
            GetEventTypeName = "@Custom Script: " & Events(EventIndex).SubEvents(SubIndex).Data(1)
            Exit Function
        Case Evt_OpenEvent
            Select Case Events(EventIndex).SubEvents(SubIndex).Data(3)
                Case 0: GetEventTypeName = "@Open Event: {" & Events(EventIndex).SubEvents(SubIndex).Data(1) & ", " & Events(EventIndex).SubEvents(SubIndex).Data(2) & "}"
                Case 1: GetEventTypeName = "@Close Event: {" & Events(EventIndex).SubEvents(SubIndex).Data(1) & ", " & Events(EventIndex).SubEvents(SubIndex).Data(2) & "}"
            End Select
            Exit Function
        Case Evt_ChangeGraphic
            GetEventTypeName = "@Change graphic: " & Events(EventIndex).SubEvents(SubIndex).Data(3) & " {" & Events(EventIndex).SubEvents(SubIndex).Data(1) & ", " & Events(EventIndex).SubEvents(SubIndex).Data(2) & "}"
            Exit Function
        Case Evt_ChangeVitals
            GetEventTypeName = "@Change Vitals"
            Exit Function
        Case Evt_PlaySound
            GetEventTypeName = "@Play Sound: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "'"
            Exit Function
        Case Evt_PlayBGM
            GetEventTypeName = "@Play Music: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "'"
            Exit Function
        Case Evt_FadeoutBGM
            GetEventTypeName = "Stop Music"
            Exit Function
        Case Evt_SpecialEffect
            GetEventTypeName = "@Special Effect"
            Exit Function
    End Select
    GetEventTypeName = "Unknown"
End Function

Public Function GetColorString(color As Long)
    Select Case color
        Case 0
            GetColorString = "Black"
        Case 1
            GetColorString = "Blue"
        Case 2
            GetColorString = "Green"
        Case 3
            GetColorString = "Cyan"
        Case 4
            GetColorString = "Red"
        Case 5
            GetColorString = "Magenta"
        Case 6
            GetColorString = "Brown"
        Case 7
            GetColorString = "Grey"
        Case 8
            GetColorString = "Dark Grey"
        Case 9
            GetColorString = "Bright Blue"
        Case 10
            GetColorString = "Bright Green"
        Case 11
            GetColorString = "Bright Cyan"
        Case 12
            GetColorString = "Bright Red"
        Case 13
            GetColorString = "Pink"
        Case 14
            GetColorString = "Yellow"
        Case 15
            GetColorString = "White"

    End Select
End Function
