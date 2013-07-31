Attribute VB_Name = "modHandleData"
Option Explicit

' ******************************************
' ** Parses and handles String packets    **
' ******************************************
Public Function GetAddress(FunAddr As Long) As Long
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    GetAddress = FunAddr
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetAddress", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Sub InitMessages()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SDevLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(SClassesData) = GetAddress(AddressOf HandleClassesData)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SItemEditor) = GetAddress(AddressOf HandleItemEditor)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SNpcEditor) = GetAddress(AddressOf HandleNpcEditor)
    HandleDataSub(SUpdateNpc) = GetAddress(AddressOf HandleUpdateNpc)
    HandleDataSub(SAnimationEditor) = GetAddress(AddressOf HandleAnimationEditor)
    HandleDataSub(SUpdateAnimation) = GetAddress(AddressOf HandleUpdateAnimation)
    HandleDataSub(SShopEditor) = GetAddress(AddressOf HandleShopEditor)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SSpellEditor) = GetAddress(AddressOf HandleSpellEditor)
    HandleDataSub(SUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(SResourceCache) = GetAddress(AddressOf HandleResourceCache)
    HandleDataSub(SResourceEditor) = GetAddress(AddressOf HandleResourceEditor)
    HandleDataSub(SUpdateResource) = GetAddress(AddressOf HandleUpdateResource)
    HandleDataSub(SEventData) = GetAddress(AddressOf Events_HandleEventData)
    HandleDataSub(SEventEditor) = GetAddress(AddressOf Events_HandleEventEditor)
    HandleDataSub(SSwitchesAndVariables) = GetAddress(AddressOf HandleSwitchesAndVariables)
    HandleDataSub(SEffectEditor) = GetAddress(AddressOf HandleEffectEditor)
    HandleDataSub(SUpdateEffect) = GetAddress(AddressOf HandleUpdateEffect)
    HandleDataSub(SMapReport) = GetAddress(AddressOf HandleMapReport)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitMessages", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub HandleData(ByRef Data() As Byte)
Dim buffer As clsBuffer
Dim MsgType As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    MsgType = buffer.ReadLong
    
    If MsgType < 0 Then
        DestroySuite
        Exit Sub
    End If

    If MsgType >= SMSG_COUNT Then
        DestroySuite
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), MyIndex, buffer.ReadBytes(buffer.Length), 0, 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub HandleAlertMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Msg As String
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Msg = buffer.ReadString 'Parse(1)
    
    Set buffer = Nothing
    Call SetStatus(Msg)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAlertMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub HandleLoginOk(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, MyName As String

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' Now we can receive game data
    MyIndex = buffer.ReadLong
    
    ' player high index
    MyName = buffer.ReadString
    
    Set buffer = Nothing
    
    SetStartData
    SendMapReport
    
    frmMain.tlbrMain.Enabled = True
    frmMain.tlbrSec.Enabled = True
    frmMain.Caption = "Developer Suite - " & Trim$(MyName)
    frmMain.picLogin.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleLoginOk", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleItemEditor()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    With frmEditor_Item
        Editor = EDITOR_ITEM
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_ITEMS
            .lstIndex.AddItem I & ": " & Trim$(Item(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        frmMain.Visible = False
        ItemEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleItemEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcEditor()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    With frmEditor_NPC
        Editor = EDITOR_NPC
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_NPCS
            .lstIndex.AddItem I & ": " & Trim$(Npc(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        frmMain.Visible = False
        NpcEditorInit
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    n = buffer.ReadLong
    
    NpcSize = LenB(Npc(n))
    ReDim NpcData(NpcSize - 1)
    NpcData = buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(Npc(n)), ByVal VarPtr(NpcData(0)), NpcSize
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpellEditor()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    With frmEditor_Spell
        Editor = EDITOR_SPELL
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_SPELLS
            .lstIndex.AddItem I & ": " & Trim$(Spell(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        frmMain.Visible = False
        SpellEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpellEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim spellnum As Long
Dim buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    spellnum = buffer.ReadLong
    
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateSpell", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResourceEditor()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    With frmEditor_Resource
        Editor = EDITOR_RESOURCE
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_RESOURCES
            .lstIndex.AddItem I & ": " & Trim$(Resource(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        frmMain.Visible = False
        ResourceEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResourceEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim ResourceNum As Long
Dim buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ResourceNum = buffer.ReadLong
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = buffer.ReadBytes(ResourceSize)
    
    ClearResource ResourceNum
    
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateResource", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleShopEditor()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    With frmEditor_Shop
        Editor = EDITOR_SHOP
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_SHOPS
            .lstIndex.AddItem I & ": " & Trim$(Shop(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        frmMain.Visible = False
        ShopEditorInit
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleShopEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim shopnum As Long
Dim buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    shopnum = buffer.ReadLong
    
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    ShopData = buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopnum)), ByVal VarPtr(ShopData(0)), ShopSize
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAnimationEditor()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    With frmEditor_Animation
        Editor = EDITOR_ANIMATION
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_ANIMATIONS
            .lstIndex.AddItem I & ": " & Trim$(Animation(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        frmMain.Visible = False
        AnimationEditorInit
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAnimationEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Animationnum As Long
Dim buffer As clsBuffer
Dim AnimationSize As Long
Dim AnimationData() As Byte
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Animationnum = buffer.ReadLong
    
    AnimationSize = LenB(Animation(Animationnum))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(Animationnum)), ByVal VarPtr(AnimationData(0)), AnimationSize
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEffectEditor()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    With frmEditor_Effect
        Editor = EDITOR_EFFECT
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_EFFECTS
            .lstIndex.AddItem I & ": " & Trim$(Effect(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        frmMain.Visible = False
        EffectEditorInit
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEffectEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateEffect(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Effectnum As Long
Dim buffer As clsBuffer
Dim EffectSize As Long
Dim EffectData() As Byte
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Effectnum = buffer.ReadLong
    
    EffectSize = LenB(Effect(Effectnum))
    ReDim EffectData(EffectSize - 1)
    EffectData = buffer.ReadBytes(EffectSize)
    CopyMemory ByVal VarPtr(Effect(Effectnum)), ByVal VarPtr(EffectData(0)), EffectSize
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateEffect", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub Events_HandleEventData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim EIndex As Long, S As Long, SCount As Long, d As Long, DCount As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    EIndex = buffer.ReadLong
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    
    Events(EIndex).name = buffer.ReadString
    Events(EIndex).chkSwitch = buffer.ReadByte
    Events(EIndex).chkVariable = buffer.ReadByte
    Events(EIndex).chkHasItem = buffer.ReadByte
    Events(EIndex).SwitchIndex = buffer.ReadLong
    Events(EIndex).SwitchCompare = buffer.ReadByte
    Events(EIndex).VariableIndex = buffer.ReadLong
    Events(EIndex).VariableCompare = buffer.ReadByte
    Events(EIndex).VariableCondition = buffer.ReadLong
    Events(EIndex).HasItemIndex = buffer.ReadLong
    SCount = buffer.ReadLong
    If SCount > 0 Then
        ReDim Events(EIndex).SubEvents(1 To SCount)
        Events(EIndex).HasSubEvents = True
        For S = 1 To SCount
            With Events(EIndex).SubEvents(S)
                .Type = buffer.ReadLong
                'Textz
                DCount = buffer.ReadLong
                If DCount > 0 Then
                    ReDim .Text(1 To DCount)
                    .HasText = True
                    For d = 1 To DCount
                        .Text(d) = buffer.ReadString
                    Next d
                Else
                    Erase .Text
                    .HasText = False
                End If
                'Dataz
                DCount = buffer.ReadLong
                If DCount > 0 Then
                    ReDim .Data(1 To DCount)
                    .HasData = True
                    For d = 1 To DCount
                        .Data(d) = buffer.ReadLong
                    Next d
                Else
                    Erase .Data
                    .HasData = False
                End If
            End With
        Next S
    Else
        Events(EIndex).HasSubEvents = False
        Erase Events(EIndex).SubEvents
    End If
    
    Events(EIndex).Trigger = buffer.ReadByte
    Events(EIndex).WalkThrought = buffer.ReadByte
    Events(EIndex).Animated = buffer.ReadByte
    For S = 0 To 2
        Events(EIndex).Graphic(S) = buffer.ReadLong
    Next
    Events(EIndex).Layer = buffer.ReadByte
    
    Set buffer = Nothing
End Sub

Public Sub Events_HandleEventEditor(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    With frmEditor_Events
        Editor = EDITOR_EVENT
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_EVENTS
            .lstIndex.AddItem I & ": " & Trim$(Events(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        frmMain.Visible = False
        EventEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Events_HandleEventEditor", "modEvents", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
Private Sub HandleSwitchesAndVariables(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String, I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    For I = 1 To MAX_SWITCHES
        Switches(I) = buffer.ReadString
    Next
    
    For I = 1 To MAX_VARIABLES
        Variables(I) = buffer.ReadString
    Next
    
    Set buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSwitchesAndVariables", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapReport(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim MapNum As Integer

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    frmMain.lstMaps.Clear
    
    For MapNum = 1 To MAX_MAPS
        frmMain.lstMaps.AddItem MapNum & ": " & buffer.ReadString
    Next MapNum
    
    Set buffer = Nothing
End Sub

Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim X As Long
Dim Y As Long
Dim I As Long
Dim buffer As clsBuffer
Dim MapNum As Long
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()

    MapNum = buffer.ReadLong
    Map.name = buffer.ReadString
    Map.Music = buffer.ReadString
    Map.Revision = buffer.ReadLong
    Map.Moral = buffer.ReadByte
    Map.Up = buffer.ReadLong
    Map.Down = buffer.ReadLong
    Map.Left = buffer.ReadLong
    Map.Right = buffer.ReadLong
    Map.BootMap = buffer.ReadLong
    Map.BootX = buffer.ReadByte
    Map.BootY = buffer.ReadByte
    Map.MaxX = buffer.ReadByte
    Map.MaxY = buffer.ReadByte
    Map.BossNpc = buffer.ReadLong
    
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For I = 1 To MapLayer.Layer_Count - 1
                Map.Tile(X, Y).Layer(I).X = buffer.ReadLong
                Map.Tile(X, Y).Layer(I).Y = buffer.ReadLong
                Map.Tile(X, Y).Layer(I).Tileset = buffer.ReadLong
                Map.Tile(X, Y).Autotile(I) = buffer.ReadByte
            Next
            Map.Tile(X, Y).Type = buffer.ReadByte
            Map.Tile(X, Y).Data1 = buffer.ReadLong
            Map.Tile(X, Y).Data2 = buffer.ReadLong
            Map.Tile(X, Y).Data3 = buffer.ReadLong
            Map.Tile(X, Y).Data4 = buffer.ReadLong
            Map.Tile(X, Y).DirBlock = buffer.ReadByte
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Map.Npc(X) = buffer.ReadLong
    Next
    Map.Fog = buffer.ReadByte
    Map.FogSpeed = buffer.ReadByte
    Map.FogOpacity = buffer.ReadByte
    
    Map.Red = buffer.ReadByte
    Map.Green = buffer.ReadByte
    Map.Blue = buffer.ReadByte
    Map.Alpha = buffer.ReadByte
    
    Map.Panorama = buffer.ReadByte
    
    Map.Weather = buffer.ReadLong
    Map.WeatherIntensity = buffer.ReadLong
    Map.BGS = buffer.ReadString
    
    Set buffer = Nothing
    initAutotiles
    HasMap = True
    CurrentMap = MapNum
    SetStatus "Received map #" & MapNum
    UpdateCamera

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResourceCache(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ' if in map editor, we cache shit ourselves
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Resource_Index = buffer.ReadLong
    Resources_Init = False

    If Resource_Index > 0 Then
        ReDim Preserve MapResource(0 To Resource_Index)

        For I = 0 To Resource_Index
            MapResource(I).ResourceState = buffer.ReadByte
            MapResource(I).X = buffer.ReadLong
            MapResource(I).Y = buffer.ReadLong
        Next

        Resources_Init = True
    Else
        ReDim MapResource(0 To 1)
    End If

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResourceCache", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub HandleClassesData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim I As Long
Dim z As Long, X As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = 1
    ' Max classes
    MAX_CLASSES = buffer.ReadLong 'CByte(Parse(n))
    ReDim Class(1 To MAX_CLASSES)
    n = n + 1

    For I = 1 To MAX_CLASSES

        With Class(I)
            .name = buffer.ReadString 'Trim$(Parse(n))
            .Vital(Vitals.HP) = buffer.ReadLong 'CLng(Parse(n + 1))
            .Vital(Vitals.MP) = buffer.ReadLong 'CLng(Parse(n + 2))
            
            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)
            ' loop-receive data
            For X = 0 To z
                .MaleSprite(X) = buffer.ReadLong
            Next
            
            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)
            ' loop-receive data
            For X = 0 To z
                .FemaleSprite(X) = buffer.ReadLong
            Next
                            
            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = buffer.ReadLong
            Next
        End With

        n = n + 10
    Next

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleClassesData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
