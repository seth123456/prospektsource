Attribute VB_Name = "modItemEditor"
Option Explicit
Public Item(1 To MAX_ITEMS) As ItemRec
Public TempItem As ItemRec

' Item constants
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_WEAPON As Byte = 1
Public Const ITEM_TYPE_ARMOR As Byte = 2
Public Const ITEM_TYPE_HELMET As Byte = 3
Public Const ITEM_TYPE_SHIELD As Byte = 4
Public Const ITEM_TYPE_CONSUME As Byte = 5
Public Const ITEM_TYPE_SPELL As Byte = 6
Public Const ITEM_TYPE_EVENT As Byte = 7

Public Type ItemRec
    name As String * NAME_LENGTH
    Desc As String * 255
    sound As String * NAME_LENGTH
    
    Pic As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    ClassReq As Long
    AccessReq As Long
    LevelReq As Long
    Price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Byte
    Rarity As Byte
    Speed As Long
    BindType As Byte
    Stat_Req(1 To Stats.Stat_Count - 1) As Byte
    Animation As Long
    Paperdoll As Long
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    Projectile As Long
    Range As Byte
    Rotation As Integer
    Ammo As Long
    isTwoHanded As Byte
    Stackable As Byte
    Effect As Long
End Type

' /////////////////
' // Item Editor //
' /////////////////
Public Sub ItemEditorInit()
Dim I As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    If frmEditor_Item.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1

    ' add the array to the combo
    frmEditor_Item.cmbSound.Clear
    frmEditor_Item.cmbSound.AddItem "None."
    For I = 1 To UBound(soundCache)
        frmEditor_Item.cmbSound.AddItem soundCache(I)
    Next
    ' finished populating
    
    ' set max values
    frmEditor_Item.scrlPic.Max = Count_Item
    frmEditor_Item.scrlAnim.Max = MAX_ANIMATIONS
    frmEditor_Item.scrlPaperdoll.Max = Count_Paperdoll
    frmEditor_Item.scrlProjectilePic.Max = Count_Projectile
    frmEditor_Item.scrlProjectileAmmo.Max = MAX_ITEMS
    frmEditor_Item.scrlEffect.Max = MAX_EFFECTS

    With Item(EditorIndex)
        frmEditor_Item.txtName.Text = Trim$(.name)
        If .Pic > frmEditor_Item.scrlPic.Max Then .Pic = 0
        frmEditor_Item.scrlPic.Value = .Pic
        frmEditor_Item.cmbType.ListIndex = .Type
        frmEditor_Item.scrlAnim.Value = .Animation
        frmEditor_Item.txtDesc.Text = Trim$(.Desc)
        frmEditor_Item.chkStackable.Value = .Stackable
        frmEditor_Item.scrlEffect.Value = .Effect
        
        ' find the sound we have set
        If frmEditor_Item.cmbSound.ListCount >= 0 Then
            For I = 0 To frmEditor_Item.cmbSound.ListCount
                If frmEditor_Item.cmbSound.List(I) = Trim$(.sound) Then
                    frmEditor_Item.cmbSound.ListIndex = I
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Item.cmbSound.ListIndex = -1 Then frmEditor_Item.cmbSound.ListIndex = 0
        End If

        ' Type specific settings
        If (frmEditor_Item.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmEditor_Item.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
            frmEditor_Item.fraEquipment.Visible = True
            frmEditor_Item.scrlDamage.Value = .Data2
            frmEditor_Item.cmbTool.ListIndex = .Data3
            frmEditor_Item.scrlProjectilePic.Value = .Projectile
            frmEditor_Item.scrlProjectileRange.Value = .Range
            frmEditor_Item.scrlProjectileRotation.Value = .Rotation
            frmEditor_Item.scrlProjectileAmmo.Value = .Ammo
            frmEditor_Item.chkTwoHanded.Value = .isTwoHanded

            If .Speed < 100 Then .Speed = 100
            frmEditor_Item.scrlSpeed.Value = .Speed
            
            ' loop for stats
            For I = 1 To Stats.Stat_Count - 1
                frmEditor_Item.scrlStatBonus(I).Value = .Add_Stat(I)
            Next
            
            frmEditor_Item.scrlPaperdoll = .Paperdoll
        Else
            frmEditor_Item.fraEquipment.Visible = False
        End If
        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_WEAPON) Then
            frmEditor_Item.Frame4.Visible = True
        Else
            frmEditor_Item.Frame4.Visible = False
        End If

        If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_CONSUME Then
            frmEditor_Item.fraVitals.Visible = True
            frmEditor_Item.scrlAddHp.Value = .AddHP
            frmEditor_Item.scrlAddMP.Value = .AddMP
            frmEditor_Item.scrlAddExp.Value = .AddEXP
        Else
            frmEditor_Item.fraVitals.Visible = False
        End If

        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            frmEditor_Item.fraSpell.Visible = True
            frmEditor_Item.scrlSpell.Value = .Data1
        Else
            frmEditor_Item.fraSpell.Visible = False
        End If
        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_EVENT) Then
            frmEditor_Item.fraEvent.Visible = True
            frmEditor_Item.scrlEvent.Value = .Data1
        Else
            frmEditor_Item.fraEvent.Visible = False
        End If

        ' Basic requirements
        frmEditor_Item.scrlAccessReq.Value = .AccessReq
        frmEditor_Item.scrlLevelReq.Value = .LevelReq
        
        ' loop for stats
        For I = 1 To Stats.Stat_Count - 1
            frmEditor_Item.scrlStatReq(I).Value = .Stat_Req(I)
        Next
        
        ' Build cmbClassReq
        frmEditor_Item.cmbClassReq.Clear
        frmEditor_Item.cmbClassReq.AddItem "None"

        For I = 1 To MAX_CLASSES
            frmEditor_Item.cmbClassReq.AddItem Class(I).name
        Next

        frmEditor_Item.cmbClassReq.ListIndex = .ClassReq
        ' Info
        frmEditor_Item.txtPrice.Text = .Price
        frmEditor_Item.cmbBind.ListIndex = .BindType
        frmEditor_Item.scrlRarity.Value = .Rarity
         
        EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    End With

    Item_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub ItemEditorOk()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    For I = 1 To MAX_ITEMS
        If Item_Changed(I) Then
            Call SendSaveItem(I)
        End If
    Next
    
    Unload frmEditor_Item
    Editor = 0
    ClearChanged_Item
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub ItemEditorCancel()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Item
    ClearChanged_Item
    ClearItems
    SendRequestItems
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Item()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    ZeroMemory Item_Changed(1), MAX_ITEMS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Item", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearItems()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    For I = 1 To MAX_ITEMS
        Call ClearItem(I)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearTempItem()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(TempItem), LenB(TempItem))
    TempItem.name = vbNullString
    TempItem.Desc = vbNullString
    TempItem.sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearTempItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
