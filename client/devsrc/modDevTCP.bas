Attribute VB_Name = "modDevTCP"
Option Explicit
' ******************************************
' ** Communcation to server, TCP          **
' ** Winsock Control (mswinsck.ocx)       **
' ** String packets (slow and big)        **
' ******************************************
Private PlayerBuffer As clsBuffer

Sub TcpInit(ByVal IP As String, Port As Long)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set PlayerBuffer = New clsBuffer

    ' connect
    frmMain.Socket.RemoteHost = Trim$(IP)
    frmMain.Socket.RemotePort = Trim$(Port)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TcpInit", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub DestroyTCP()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    frmMain.Socket.Close
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DestroyTCP", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
Dim buffer() As Byte
Dim pLength As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    frmMain.Socket.GetData buffer, vbUnicode, DataLength
    
    PlayerBuffer.WriteBytes buffer()
    
    If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Do While pLength > 0 And pLength <= PlayerBuffer.Length - 4
        If pLength <= PlayerBuffer.Length - 4 Then
            PlayerBuffer.ReadLong
            HandleData PlayerBuffer.ReadBytes(pLength)
        End If

        pLength = 0
        If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Loop
    PlayerBuffer.Trim
    DoEvents
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "IncomingData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Function ConnectToServer(ByVal I As Long) As Boolean
Dim Wait As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = timeGetTime
    frmMain.Socket.Close
    frmMain.Socket.Connect
    
    SetStatus "Connecting to server..."
    
    ' Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsConnected) And (timeGetTime <= Wait + 3000)
        DoEvents
    Loop
    
    ConnectToServer = IsConnected

    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConnectToServer", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Function IsConnected() As Boolean
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If frmMain.Socket.State = sckConnected Then
        IsConnected = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsConnected", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub SendData(ByRef Data() As Byte)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If IsConnected Then
        Set buffer = New clsBuffer
                
        buffer.WriteLong (UBound(Data) - LBound(Data)) + 1
        buffer.WriteBytes Data()
        frmMain.Socket.SendData buffer.ToArray()
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDevLogin(ByVal name As String, ByVal Password As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
   On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CDevLogin
    buffer.WriteString name
    buffer.WriteString Password
    buffer.WriteLong App.Major
    buffer.WriteLong App.Minor
    buffer.WriteLong App.Revision
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDevLogin", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditItem()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditItem
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveItem(ByVal itemNum As Long)
Dim buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    ItemSize = LenB(Item(itemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemNum)), ItemSize
    buffer.WriteLong CSaveItem
    buffer.WriteLong itemNum
    buffer.WriteBytes ItemData
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestItems()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestItems
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestItems", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditSpell()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditSpell
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditSpell", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveSpell(ByVal spellnum As Long)
Dim buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellnum)), SpellSize
    
    buffer.WriteLong CSaveSpell
    buffer.WriteLong spellnum
    buffer.WriteBytes SpellData
    SendData buffer.ToArray()
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveSpell", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestSpells()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestSpells
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestSpells", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditNpc()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditNpc
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditNpc", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveNpc(ByVal npcNum As Long)
Dim buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    NpcSize = LenB(Npc(npcNum))
    ReDim NpcData(NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(Npc(npcNum)), NpcSize
    buffer.WriteLong CSaveNpc
    buffer.WriteLong npcNum
    buffer.WriteBytes NpcData
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveNpc", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestNPCS()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestNPCS
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestNPCS", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditResource()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditResource
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditResource", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveResource(ByVal ResourceNum As Long)
Dim buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    buffer.WriteLong CSaveResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveResource", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestResources()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestResources
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestResources", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditShop()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditShop
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditShop", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveShop(ByVal shopnum As Long)
Dim buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopnum)), ShopSize
    buffer.WriteLong CSaveShop
    buffer.WriteLong shopnum
    buffer.WriteBytes ShopData
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveShop", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestShops()
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestShops
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestShops", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditAnimation()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditAnimation
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditAnimation", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveAnimation(ByVal Animationnum As Long)
Dim buffer As clsBuffer
Dim AnimationSize As Long
Dim AnimationData() As Byte

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    AnimationSize = LenB(Animation(Animationnum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(Animationnum)), AnimationSize
    buffer.WriteLong CSaveAnimation
    buffer.WriteLong Animationnum
    buffer.WriteBytes AnimationData
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveAnimation", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestAnimations()
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestAnimations
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestAnimations", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditEffect()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditEffect
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditEffect", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveEffect(ByVal Effectnum As Long)
Dim buffer As clsBuffer
Dim EffectSize As Long
Dim EffectData() As Byte

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    EffectSize = LenB(Effect(Effectnum))
    ReDim EffectData(EffectSize - 1)
    CopyMemory EffectData(0), ByVal VarPtr(Effect(Effectnum)), EffectSize
    buffer.WriteLong CSaveEffect
    buffer.WriteLong Effectnum
    buffer.WriteBytes EffectData
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveEffect", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestEffects()
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEffects
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEffects", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub RequestSwitchesAndVariables()
    Dim I As Long, buffer As clsBuffer
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestSwitchesAndVariables
    SendData buffer.ToArray
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RequestSwitchesAndVariables", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendSwitchesAndVariables()
    Dim I As Long, buffer As clsBuffer
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSwitchesAndVariables
    For I = 1 To MAX_SWITCHES
        buffer.WriteString Switches(I)
    Next
    For I = 1 To MAX_VARIABLES
        buffer.WriteString Variables(I)
    Next
    SendData buffer.ToArray
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSwitchesAndVariables", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub Events_SendSaveEvent(ByVal EIndex As Long)
    Dim buffer As clsBuffer
    Dim I As Long, d As Long
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    Set buffer = New clsBuffer
    
    buffer.WriteLong CSaveEventData
    buffer.WriteLong EIndex
    buffer.WriteString Events(EIndex).name
    buffer.WriteByte Events(EIndex).chkSwitch
    buffer.WriteByte Events(EIndex).chkVariable
    buffer.WriteByte Events(EIndex).chkHasItem
    buffer.WriteLong Events(EIndex).SwitchIndex
    buffer.WriteByte Events(EIndex).SwitchCompare
    buffer.WriteLong Events(EIndex).VariableIndex
    buffer.WriteByte Events(EIndex).VariableCompare
    buffer.WriteLong Events(EIndex).VariableCondition
    buffer.WriteLong Events(EIndex).HasItemIndex
    If Events(EIndex).HasSubEvents Then
        buffer.WriteLong UBound(Events(EIndex).SubEvents)
        For I = 1 To UBound(Events(EIndex).SubEvents)
            With Events(EIndex).SubEvents(I)
                buffer.WriteLong .Type
                If .HasText Then
                    buffer.WriteLong UBound(.Text)
                    For d = 1 To UBound(.Text)
                        buffer.WriteString .Text(d)
                    Next d
                Else
                    buffer.WriteLong 0
                End If
                If .HasData Then
                    buffer.WriteLong UBound(.Data)
                    For d = 1 To UBound(.Data)
                        buffer.WriteLong .Data(d)
                    Next d
                Else
                    buffer.WriteLong 0
                End If
            End With
        Next I
    Else
        buffer.WriteLong 0
    End If
    
    buffer.WriteByte Events(EIndex).Trigger
    buffer.WriteByte Events(EIndex).WalkThrought
    buffer.WriteByte Events(EIndex).Animated
    For I = 0 To 2
        buffer.WriteLong Events(EIndex).Graphic(I)
    Next
    buffer.WriteByte Events(EIndex).Layer
    
    SendData buffer.ToArray

    Set buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Events_SendSaveEvent", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub Events_SendRequestEditEvents()
    Dim buffer As clsBuffer
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditEvents
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Events_SendRequestEditEvents", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
Public Sub Events_SendRequestEventsData()
    Dim buffer As clsBuffer
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEventsData
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Events_SendRequestEventsData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
Public Sub Events_SendRequestEventData(ByVal I As Long)
    Dim buffer As clsBuffer
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEventData
    buffer.WriteLong I
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Events_SendRequestEventData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMapReport()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CMapReport
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMapReport", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpTo(ByVal MapNum As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If MapNum = 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong CDevMap
    buffer.WriteLong MapNum
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WarpTo", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMap()
Dim packet As String
Dim X As Long
Dim Y As Long
Dim I As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    Set buffer = New clsBuffer

    With Map
        buffer.WriteLong CMapData
        buffer.WriteLong CurrentMap
        buffer.WriteString Trim$(.name)
        buffer.WriteString Trim$(.Music)
        buffer.WriteByte .Moral
        buffer.WriteLong .Up
        buffer.WriteLong .Down
        buffer.WriteLong .Left
        buffer.WriteLong .Right
        buffer.WriteLong .BootMap
        buffer.WriteByte .BootX
        buffer.WriteByte .BootY
        buffer.WriteByte .MaxX
        buffer.WriteByte .MaxY
        buffer.WriteLong .BossNpc
    End With

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            With Map.Tile(X, Y)
                For I = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteLong .Layer(I).X
                    buffer.WriteLong .Layer(I).Y
                    buffer.WriteLong .Layer(I).Tileset
                    buffer.WriteByte .Autotile(I)
                Next
                buffer.WriteByte .Type
                buffer.WriteLong .Data1
                buffer.WriteLong .Data2
                buffer.WriteLong .Data3
                buffer.WriteLong .Data4
                buffer.WriteByte .DirBlock
            End With
        Next
    Next

    With Map
        For X = 1 To MAX_MAP_NPCS
            buffer.WriteLong .Npc(X)
        Next
    End With
    
    buffer.WriteByte Map.Fog
    buffer.WriteByte Map.FogSpeed
    buffer.WriteByte Map.FogOpacity
    
    buffer.WriteByte Map.Red
    buffer.WriteByte Map.Green
    buffer.WriteByte Map.Blue
    buffer.WriteByte Map.Alpha
    
    buffer.WriteByte Map.Panorama
    
    buffer.WriteLong Map.Weather
    buffer.WriteLong Map.WeatherIntensity
    buffer.WriteString Trim$(Map.BGS)

    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
