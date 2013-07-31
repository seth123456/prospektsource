Attribute VB_Name = "modEnumerations"
Option Explicit

' The order of the packets must match with the server's packet enumeration

' Packets sent by server to client
Public Enum ServerPackets
    SAlertMsg = 1
    SLoginOk
    SDevLoginOk
    SNewCharClasses
    SClassesData
    SInGame
    SPlayerInv
    SPlayerInvUpdate
    SPlayerWornEq
    SPlayerHp
    SPlayerMp
    SPlayerStats
    SPlayerData
    SPlayerMove
    SNpcMove
    SPlayerDir
    SNpcDir
    SPlayerXY
    SPlayerXYMap
    SAttack
    SNpcAttack
    SCheckForMap
    SMapData
    SMapItemData
    SMapNpcData
    SMapDone
    SGlobalMsg
    SAdminMsg
    SPlayerMsg
    SMapMsg
    SSpawnItem
    SItemEditor
    SUpdateItem
    SREditor
    SSpawnNpc
    SNpcDead
    SNpcEditor
    SUpdateNpc
    SEditMap
    SShopEditor
    SUpdateShop
    SSpellEditor
    SUpdateSpell
    SSpells
    SLeft
    SResourceCache
    SResourceEditor
    SUpdateResource
    SSendPing
    SActionMsg
    SPlayerEXP
    SBlood
    SAnimationEditor
    SUpdateAnimation
    SAnimation
    SMapNpcVitals
    SCooldown
    SClearSpellBuffer
    SSayMsg
    SOpenShop
    SStunned
    SMapWornEq
    SBank
    STrade
    SCloseTrade
    STradeUpdate
    STradeStatus
    STarget
    SHotbar
    SHighIndex
    SSound
    STradeRequest
    SPartyInvite
    SPartyUpdate
    SPartyVitals
    SStartTutorial
    SChatBubble
    SMapReport
    SEventData
    SEventEditor
    SEventUpdate
    SSwitchesAndVariables
    SEventOpen
    SCreateProjectile
    SEventGraphic
    SClientTime
    SPlaySound
    SPlayBGM
    SFadeoutBGM
    SEffectEditor
    SUpdateEffect
    SEffect
    SSpecialEffect
    SThreshold
    SSwearFilter
    ' Make sure SMSG_COUNT is below everything else
    SMSG_COUNT
End Enum

' Packets sent by client to server
Public Enum ClientPackets
    CNewAccount = 1
    CLogin
    CDevLogin
    CAddChar
    CUseChar
    CSayMsg
    CEmoteMsg
    CBroadcastMsg
    CPlayerMsg
    CPlayerMove
    CPlayerDir
    CUseItem
    CAttack
    CUseStatPoint
    CWarpMeTo
    CWarpToMe
    CWarpTo
    CSetSprite
    CRequestNewMap
    CMapData
    CNeedMap
    CMapGetItem
    CMapDropItem
    CMapRespawn
    CMapReport
    CKickPlayer
    CBanList
    CBanDestroy
    CBanPlayer
    CRequestEditMap
    CRequestEditItem
    CSaveItem
    CRequestEditNpc
    CSaveNpc
    CRequestEditShop
    CSaveShop
    CRequestEditSpell
    CSaveSpell
    CSetAccess
    CWhosOnline
    CSetMotd
    CTarget
    CSpells
    CCast
    CQuit
    CSwapInvSlots
    CRequestEditResource
    CSaveResource
    CCheckPing
    CUnequip
    CRequestPlayerData
    CRequestItems
    CRequestNPCS
    CRequestResources
    CSpawnItem
    CRequestEditAnimation
    CSaveAnimation
    CRequestAnimations
    CRequestSpells
    CRequestShops
    CRequestLevelUp
    CForgetSpell
    CCloseShop
    CBuyItem
    CSellItem
    CChangeBankSlots
    CDepositItem
    CWithdrawItem
    CCloseBank
    CAdminWarp
    CTradeRequest
    CAcceptTrade
    CDeclineTrade
    CTradeItem
    CUntradeItem
    CHotbarChange
    CHotbarUse
    CSwapSpellSlots
    CAcceptTradeRequest
    CDeclineTradeRequest
    CPartyRequest
    CAcceptParty
    CDeclineParty
    CPartyLeave
    CFinishTutorial
    CSwitchesAndVariables
    CRequestSwitchesAndVariables
    CSaveEventData
    CRequestEditEvents
    CRequestEventData
    CRequestEventsData
    CChooseEventOption
    CRequestEditEffect
    CSaveEffect
    CRequestEffects
    CDevMap
    ' Make sure CMSG_COUNT is below everything else
    CMSG_COUNT
End Enum


Public HandleDataSub(SMSG_COUNT) As Long

' Stats used by Players, Npcs and Classes
Public Enum Stats
    Strength = 1
    Endurance
    Intelligence
    Agility
    Willpower
    ' Make sure Stat_Count is below everything else
    Stat_Count
End Enum

' Vitals used by Players, Npcs and Classes
Public Enum Vitals
    HP = 1
    MP
    ' Make sure Vital_Count is below everything else
    Vital_Count
End Enum

' Equipment used by Players
Public Enum Equipment
    Weapon = 1
    Armor
    Helmet
    Shield
    ' Make sure Equipment_Count is below everything else
    Equipment_Count
End Enum

' Layers in a map
Public Enum MapLayer
    Ground = 1
    Mask
    Mask2
    Fringe
    Fringe2
    Roof
    ' Make sure Layer_Count is below everything else
    Layer_Count
End Enum

' Sound entities
Public Enum SoundEntity
    seAnimation = 1
    seItem
    seNpc
    seResource
    seSpell
    seEffect
    ' Make sure SoundEntity_Count is below everything else
    SoundEntity_Count
End Enum

Public Enum GUIType
    GUI_CHAT = 1
    GUI_HOTBAR
    GUI_MENU
    GUI_BARS
    GUI_INVENTORY
    GUI_SPELLS
    GUI_CHARACTER
    GUI_OPTIONS
    GUI_PARTY
    GUI_DESCRIPTION
    GUI_MAINMENU
    GUI_SHOP
    GUI_BANK
    GUI_TRADE
    GUI_CURRENCY
    GUI_DIALOGUE
    GUI_EVENTCHAT
    GUI_TUTORIAL
    GUI_Count
End Enum

Public Enum ButtonType
    Button_Inventory = 1
    Button_Spells
    Button_Character
    Button_Options
    Button_Trade
    Button_Party
    Button_Login
    Button_Register
    Button_Credits
    Button_Exit
    Button_LoginAccept
    Button_RegisterAccept
    Button_ClassAccept
    Button_ClassNext
    Button_NewCharAccept
    Button_GenderLeft
    Button_GenderRight
    Button_AddStats1
    Button_AddStats2
    Button_AddStats3
    Button_AddStats4
    Button_AddStats5
    Button_ShopBuy
    Button_ShopSell
    Button_ShopExit
    Button_PartyInvite
    Button_PartyDisband
    Button_MusicOn
    Button_MusicOff
    Button_SoundOn
    Button_SoundOff
    Button_DebugOn
    Button_DebugOff
    Button_AutotileOn
    Button_AutotileOff
    Button_FullscreenOn
    Button_FullscreenOff
    Button_ChatUp
    Button_ChatDown
    Button_TradeAccept
    Button_TradeDecline
    Button_Count
End Enum

Public Enum EventType
    Evt_Message = 0
    Evt_Menu
    Evt_Quit
    Evt_OpenShop
    Evt_OpenBank
    Evt_GiveItem
    Evt_ChangeLevel
    Evt_PlayAnimation
    Evt_Warp
    Evt_GOTO
    Evt_Switch
    Evt_Variable
    Evt_AddText
    Evt_Chatbubble
    Evt_Branch
    Evt_ChangeSkill
    Evt_ChangeSprite
    Evt_ChangePK
    Evt_ChangeClass
    Evt_ChangeSex
    Evt_ChangeExp
    Evt_SetAccess
    Evt_CustomScript
    Evt_OpenEvent
    Evt_ChangeGraphic
    Evt_ChangeVitals
    Evt_PlaySound
    Evt_PlayBGM
    Evt_FadeoutBGM
    Evt_SpecialEffect
    'EventType_Count should be below everything else
    EventType_Count
End Enum

Public Enum ComparisonOperator
    GEQUAL = 0
    LEQUAL
    GREATER
    LESS
    EQUAL
    NOTEQUAL
    'ComparisonOperator_Count should be below everything else
    ComparisonOperator_Count
End Enum

Public Enum SEffectType
    SEFFECT_TYPE_FADEIN = 1
    SEFFECT_TYPE_FADEOUT
    SEFFECT_TYPE_FLASH
    SEFFECT_TYPE_FOG
    SEFFECT_TYPE_WEATHER
    SEFFECT_TYPE_TINT
End Enum
