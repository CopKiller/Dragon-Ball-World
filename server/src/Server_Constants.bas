Attribute VB_Name = "Server_Constants"
Option Explicit

' System of compressing
Public Declare Function Compress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Public Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

' Connection details
Public Const GAME_SERVER_IP As String = "127.0.0.1" ' "46.23.70.66"
Public Const GAME_SERVER_PORT As Long = 7001 ' the port used by the main game server

Public Const GAME_NAME As String = "Crystalshire"
Public Const GAME_WEBSITE As String = "http://www.crystalshire.com"

' API
Public Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

' path constants
Public Const ADMIN_LOG As String = "admin.log"
Public Const PLAYER_LOG As String = "player.log"

' Version constants
Public Const CLIENT_MAJOR As Byte = 1
Public Const CLIENT_MINOR As Byte = 8
Public Const CLIENT_REVISION As Byte = 0
Public Const MAX_LINES As Long = 500 ' Used for frmServer.txtText

' ********************************************************
' * The values below must match with the client's values *
' ********************************************************
' General constants
Public Const MAX_PLAYERS As Long = 200
Public Const MAX_ITEMS As Long = 255
Public Const MAX_NPCS As Long = 255
Public Const MAX_ANIMATIONS As Long = 255
Public Const MAX_INV As Long = 35
Public Const MAX_MAP_ITEMS As Long = 255
Public Const MAX_MAP_NPCS As Long = 30
Public Const MAX_SHOPS As Long = 50
Public Const MAX_PLAYER_SPELLS As Long = 35
Public Const MAX_SPELLS As Long = 255
Public Const MAX_TRADES As Long = 35
Public Const MAX_RESOURCES As Long = 100
Public Const MAX_LEVELS As Long = 20
Public Const MAX_BANK As Long = 99
Public Const MAX_HOTBAR As Long = 12
Public Const MAX_PARTYS As Long = 35
Public Const MAX_PARTY_MEMBERS As Long = 4
Public Const MAX_CONVS As Byte = 255
Public Const MAX_NPC_DROPS As Byte = 30
Public Const MAX_NPC_SPELLS As Byte = 10
Public Const MAX_PROJECTILE_PLAYER As Byte = 25
Public Const MAX_PROJECTILE_MAP As Byte = 125

' server-side stuff
Public Const ITEM_SPAWN_TIME As Long = 30000 ' 30 seconds
Public Const ITEM_DESPAWN_TIME As Long = 600000 ' 10 minutes
Public Const MAX_DOTS As Long = 30

' text color constants
Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15
Public Const DarkBrown As Byte = 16
Public Const Gold As Byte = 17

Public Const SayColor As Byte = White
Public Const GlobalColor As Byte = BrightBlue
Public Const BroadcastColor As Byte = White
Public Const TellColor As Byte = BrightGreen
Public Const EmoteColor As Byte = BrightCyan
Public Const AdminColor As Byte = BrightCyan
Public Const HelpColor As Byte = BrightBlue
Public Const WhoColor As Byte = BrightBlue
Public Const JoinLeftColor As Byte = DarkGrey
Public Const NpcColor As Byte = Brown
Public Const AlertColor As Byte = Red
Public Const NewMapColor As Byte = BrightBlue

' Boolean constants
Public Const NO As Byte = 0
Public Const YES As Byte = 1

' String constants
Public Const NAME_LENGTH As Byte = 20
Public Const ACCOUNT_LENGTH As Byte = 12
Public Const EMAIL_LENGTH As Byte = 25
Public Const DESC_LENGTH As Byte = 150

' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' Map constants
Public Const MAX_MAPS As Long = 100
Public Const MAX_MAPX As Byte = 24
Public Const MAX_MAPY As Byte = 18
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1
Public Const MAP_MORAL_BOSS As Byte = 2

' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_KEY As Byte = 5
Public Const TILE_TYPE_KEYOPEN As Byte = 6
Public Const TILE_TYPE_RESOURCE As Byte = 7
Public Const TILE_TYPE_DOOR As Byte = 8
Public Const TILE_TYPE_NPCSPAWN As Byte = 9
Public Const TILE_TYPE_SHOP As Byte = 10
Public Const TILE_TYPE_BANK As Byte = 11
Public Const TILE_TYPE_HEAL As Byte = 12
Public Const TILE_TYPE_TRAP As Byte = 13
Public Const TILE_TYPE_SLIDE As Byte = 14
Public Const TILE_TYPE_CHAT As Byte = 15
Public Const TILE_TYPE_APPEAR As Byte = 16
Public Const TILE_TYPE_SOUND As Byte = 17

' Item constants
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_WEAPON As Byte = 1
Public Const ITEM_TYPE_ARMOR As Byte = 2
Public Const ITEM_TYPE_HELMET As Byte = 3
Public Const ITEM_TYPE_SHIELD As Byte = 4
Public Const ITEM_TYPE_PANTS As Byte = 5
Public Const ITEM_TYPE_FEET As Byte = 6
Public Const ITEM_TYPE_CONSUME As Byte = 7
Public Const ITEM_TYPE_KEY As Byte = 8
Public Const ITEM_TYPE_CURRENCY As Byte = 9
Public Const ITEM_TYPE_SPELL As Byte = 10
Public Const ITEM_TYPE_UNIQUE As Byte = 11
Public Const ITEM_TYPE_FOOD As Byte = 12

' Direction constants
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3
Public Const DIR_UP_LEFT As Byte = 4
Public Const DIR_UP_RIGHT As Byte = 5
Public Const DIR_DOWN_LEFT As Byte = 6
Public Const DIR_DOWN_RIGHT As Byte = 7

' Constants for player movement
Public Const MOVING_WALKING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2

' Tile size constants
Public Const PIC_X As Long = 32
Public Const PIC_Y As Long = 32

' Admin constants
Public Const ADMIN_MONITOR As Byte = 1
Public Const ADMIN_MAPPER As Byte = 2
Public Const ADMIN_DEVELOPER As Byte = 3
Public Const ADMIN_CREATOR As Byte = 4

' NPC constants
Public Const NPC_BEHAVIOUR_ATTACKONSIGHT As Byte = 0
Public Const NPC_BEHAVIOUR_ATTACKWHENATTACKED As Byte = 1
Public Const NPC_BEHAVIOUR_FRIENDLY As Byte = 2
Public Const NPC_BEHAVIOUR_SHOPKEEPER As Byte = 3
Public Const NPC_BEHAVIOUR_GUARD As Byte = 4

' Spell constants
Public Const SPELL_TYPE_DAMAGEHP As Byte = 0
Public Const SPELL_TYPE_DAMAGEMP As Byte = 1
Public Const SPELL_TYPE_HEALHP As Byte = 2
Public Const SPELL_TYPE_HEALMP As Byte = 3
Public Const SPELL_TYPE_WARP As Byte = 4
Public Const SPELL_TYPE_PROJECTILE As Byte = 5

' Game editor constants
Public Const EDITOR_ITEM As Byte = 1
Public Const EDITOR_NPC As Byte = 2
Public Const EDITOR_SPELL As Byte = 3
Public Const EDITOR_SHOP As Byte = 4

' Target type constants
Public Const TARGET_TYPE_NONE As Byte = 0
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_NPC As Byte = 2

' Default starting location [Server Only]
Public Const START_MAP As Long = 1
Public Const START_X As Long = 30
Public Const START_Y As Long = 10

' Scrolling action message constants
Public Const ACTIONMSG_STATIC As Long = 0
Public Const ACTIONMSG_SCROLL As Long = 1
Public Const ACTIONMSG_SCREEN As Long = 2

' Do Events
Public Const nLng As Long = (&H80 Or &H1 Or &H4 Or &H20) + (&H8 Or &H40)

' dialogue alert strings
Public Const DIALOGUE_MSG_CONNECTION As Byte = 1
Public Const DIALOGUE_MSG_BANNED As Byte = 2
Public Const DIALOGUE_MSG_KICKED As Byte = 3
Public Const DIALOGUE_MSG_OUTDATED As Byte = 4
Public Const DIALOGUE_MSG_USERLENGTH As Byte = 5
Public Const DIALOGUE_MSG_ILLEGALNAME As Byte = 6
Public Const DIALOGUE_MSG_REBOOTING As Byte = 7
Public Const DIALOGUE_MSG_NAMETAKEN As Byte = 8
Public Const DIALOGUE_MSG_NAMELENGTH As Byte = 9
Public Const DIALOGUE_MSG_NAMEILLEGAL As Byte = 10
Public Const DIALOGUE_MSG_MYSQL As Byte = 11
Public Const DIALOGUE_MSG_WRONGPASS As Byte = 12
Public Const DIALOGUE_MSG_ACTIVATED As Byte = 13
Public Const DIALOGUE_MSG_MERGE As Byte = 14
Public Const DIALOGUE_MSG_MAXCHARS As Byte = 15
Public Const DIALOGUE_MSG_MERGENAME As Byte = 16
Public Const DIALOGUE_MSG_DELCHAR As Byte = 17
Public Const DIALOGUE_ACCOUNT_CREATED As Byte = 18

' Menu
Public Const MENU_MAIN As Byte = 1
Public Const MENU_LOGIN As Byte = 2
Public Const MENU_REGISTER As Byte = 3
Public Const MENU_CREDITS As Byte = 4
Public Const MENU_CLASS As Byte = 5
Public Const MENU_NEWCHAR As Byte = 6
Public Const MENU_CHARS As Byte = 7
Public Const MENU_MERGE As Byte = 8

' values
Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767
Public Const MAX_LONG As Long = 2147483647

Public Const DegreeToRadian As Single = 0.0174532919296
Public Const RadianToDegree As Single = 57.2958300962816
