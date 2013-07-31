Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map(1 To MAX_CACHED_MAPS) As MapRec
Public MapCache(1 To MAX_CACHED_MAPS) As Cache
Public PlayersOnMap(1 To MAX_CACHED_MAPS) As Long
Public ResourceCache(1 To MAX_CACHED_MAPS) As ResourceCacheRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Bank(1 To MAX_PLAYERS) As BankRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_CACHED_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_CACHED_MAPS) As MapDataRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Party(1 To MAX_PARTYS) As PartyRec
Public Effect(1 To MAX_EFFECTS) As EffectRec
Public Events(1 To MAX_EVENTS) As EventWrapperRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Switches(1 To MAX_SWITCHES) As String
Public Variables(1 To MAX_VARIABLES) As String

Public Class() As ClassRec
Public GameTime As TimeRec
Public SwearFilter() As SwearFilterRec

' server-side
Public Options As OptionsRec

Private Type OptionsRec
    Game_Name As String
    MOTD As String
    Port As Long
    Logs As Byte
    HighIndexing As Byte
    StartMap As Long
    StartX As Long
    StartY As Long
End Type

Private Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerInvRec
    Num As Long
    Value As Long
    Bound As Byte
End Type

Private Type Cache
    Data() As Byte
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Private Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Private Type PlayerRec
    ' Account
    Login As String * ACCOUNT_LENGTH
    Password As String * PASS_LENGTH
    
    ' General
    Name As String * ACCOUNT_LENGTH
    Sex As Byte
    Class As Long
    Sprite As Long
    Level As Byte
    exp As Long
    Access As Byte
    PK As Byte
    
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Long
    
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    
    ' Hotbar
    Hotbar(1 To MAX_HOTBAR) As HotbarRec
    
    ' Position
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte
    
    ' Tutorial
    TutorialState As Byte
    
    ' Banned
    isBanned As Byte
    isMuted As Byte
    
    Switches(0 To MAX_SWITCHES) As Byte
    Variables(0 To MAX_VARIABLES) As Long
    
    EventOpen(1 To MAX_EVENTS) As Byte
    EventGraphic(1 To MAX_EVENTS) As Byte
    Threshold As Byte
End Type

Private Type SpellBufferRec
    Spell As Long
    Timer As Long
    target As Long
    tType As Byte
End Type

Private Type DoTRec
    Used As Boolean
    Spell As Long
    Timer As Long
    Caster As Long
    StartTime As Long
End Type

Private Type TempPlayerRec
    ' Non saved local vars
    Buffer As clsBuffer
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    targetType As Byte
    target As Long
    GettingMap As Byte
    SpellCD(1 To MAX_PLAYER_SPELLS) As Long
    InShop As Long
    StunTimer As Long
    StunDuration As Long
    InBank As Boolean
    ' trade
    TradeRequest As Long
    InTrade As Long
    TradeOffer(1 To MAX_INV) As PlayerInvRec
    AcceptTrade As Boolean
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' spell buffer
    spellBuffer As SpellBufferRec
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' party
    inParty As Long
    partyInvite As Long
    ' chat
    inEventWith As Long
    inEventMap As Long
    CurrentEvent As Long
End Type

Private Type TileDataRec
    x As Long
    y As Long
    Tileset As Long
End Type

Private Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To MapLayer.Layer_Count - 1) As Byte
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    DirBlock As Byte
End Type

Private Type MapRec
    Name As String * NAME_LENGTH
    Music As String * MUSIC_LENGTH
    BGS As String * MUSIC_LENGTH
    
    Revision As Long
    Moral As Byte
    
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    MaxX As Byte
    MaxY As Byte
    
    Tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Long
    BossNpc As Long
    Fog As Byte
    FogSpeed As Byte
    FogOpacity As Byte
    
    Red As Byte
    Green As Byte
    Blue As Byte
    Alpha As Byte
    
    Panorama As Byte
    
    Weather As Long
    WeatherIntensity As Long
End Type

Private Type ClassRec
    Name As String * NAME_LENGTH
    Stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    
    startItemCount As Long
    StartItem() As Long
    StartValue() As Long
    
    startSpellCount As Long
    StartSpell() As Long
End Type

Private Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
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

Private Type MapItemRec
    Num As Long
    Value As Long
    x As Byte
    y As Byte
    ' ownership + despawn
    playerName As String
    playerTimer As Long
    canDespawn As Boolean
    despawnTimer As Long
    Bound As Boolean
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    Sound As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    Stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    exp As Long
    Animation As Long
    Damage As Long
    Level As Long
    
    ' Npc drops
    DropChance(1 To MAX_NPC_DROPS) As Double
    DropItem(1 To MAX_NPC_DROPS) As Byte
    DropItemValue(1 To MAX_NPC_DROPS) As Integer
    
    ' Casting
    Spell(1 To MAX_NPC_SPELLS) As Long
    
    Event As Long
    Projectile As Long
    ProjectileRange As Byte
    Rotation As Integer
    Moral As Byte
    Effect As Long
End Type

Private Type MapNpcRec
    Num As Long
    target As Long
    targetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    x As Byte
    y As Byte
    Dir As Byte
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
    StunDuration As Long
    StunTimer As Long
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' spell casting
    spellBuffer As SpellBufferRec
    SpellCD(1 To MAX_NPC_SPELLS) As Long
    ' Event
    e_lastDir As Byte
    inEventWith As Long
    ' pathfinding
    arPath() As tPoint
    hasPath As Boolean
    targetX As Long
    targetY As Long
    pathLoc As Long
End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    costitem As Long
    costvalue As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Type As Byte
    mpCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    x As Long
    y As Long
    Dir As Byte
    Duration As Long
    Interval As Long
    Range As Byte
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
    Vital(1 To Vitals.Vital_Count - 1) As Long
    VitalType(1 To Vitals.Vital_Count - 1) As Byte
    Effect As Long
End Type

Private Type MapDataRec
    Npc() As MapNpcRec
End Type

Private Type MapResourceRec
    ResourceState As Byte
    ResourceTimer As Long
    x As Long
    y As Long
    cur_health As Long
End Type

Private Type ResourceCacheRec
    Resource_Count As Long
    ResourceData() As MapResourceRec
End Type

Private Type ResourceRec
    Name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
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

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    LoopTime(0 To 1) As Long
End Type

Private Type SubEventRec
    Type As EventType
    HasText As Boolean
    Text() As String * 250
    HasData As Boolean
    Data() As Long
End Type

Private Type EventWrapperRec
    Name As String * NAME_LENGTH
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

Private Type TimeRec
     Minute As Byte
     Hour As Byte
     Day As Byte
     Month As Byte
     Year As Long
End Type

Private Type EffectRec
    Name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    isMulti As Byte
    MultiParticle(1 To MAX_MULTIPARTICLE) As Long
    Type As Long
    Sprite As Long
    Particles As Long
    Size As Single
    Alpha As Single
    Decay As Single
    Red As Single
    Green As Single
    Blue As Single
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Duration As Single
    XSpeed As Single
    YSpeed As Single
    XAcc As Single
    YAcc As Single
    Modifier As Byte
End Type

Private Type SwearFilterRec
    BadWord As String
    NewWord As String
End Type
