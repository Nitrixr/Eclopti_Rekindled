Attribute VB_Name = "modConstants"
Option Explicit

' API
Public Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

' path constants
Public Const ADMIN_LOG As String = "admin.log"
Public Const PLAYER_LOG As String = "player.log"

' Version constants
Public Const CLIENT_MAJOR As Byte = 1
Public Const CLIENT_MINOR As Byte = 3
Public Const CLIENT_REVISION As Byte = 0
Public Const MAX_LINES As Long = 500 ' Used for frmServer.txtText

' ********************************************************
' * The values below must match with the client's values *
' ********************************************************
' General constants
Public Const MAX_PLAYERS As Long = 70
Public Const MAX_ITEMS As Long = 255
Public Const MAX_NPCS As Long = 255
Public Const MAX_ANIMATIONS As Long = 255
Public Const MAX_INV As Long = 35
Public Const MAX_MAP_ITEMS As Long = 255
Public Const MAX_MAP_NPCS As Long = 30
Public Const MAX_SHOPS As Long = 50
Public Const MAX_PLAYER_SPELLS As Long = 35
Public Const MAX_SPELLS As Long = 255
Public Const MAX_TRADES As Long = 30
Public Const MAX_RESOURCES As Long = 100
Public Const MAX_LEVELS As Long = 100
Public Const MAX_BANK As Long = 99
Public Const MAX_HOTBAR As Long = 12
Public Const MAX_PARTYS As Long = 35
Public Const MAX_PARTY_MEMBERS As Long = 4
Public Const MAX_SWITCHES As Long = 1000
Public Const MAX_VARIABLES As Long = 1000

' server-side stuff
Public Const ITEM_SPAWN_TIME As Long = 30000 ' 30 seconds
Public Const ITEM_DESPAWN_TIME As Long = 90000 ' 1:30 seconds
Public Const MAX_DOTS As Long = 30

' Boolean constants
Public Const NO As Byte = 0
Public Const YES As Byte = 1

' String constants
Public Const NAME_LENGTH As Byte = 20
Public Const MUSIC_LENGTH As Byte = 40
Public Const ACCOUNT_LENGTH As Byte = 12

' Map constants
Public Const MAX_MAPS As Long = 100
Public Const MAX_MAPX As Byte = 14
Public Const MAX_MAPY As Byte = 11

' ********************************************
' Default starting location [Server Only]
Public Const START_MAP As Long = 1
Public Const START_X As Long = 12
Public Const START_Y As Long = 3

' Do Events
Public Const nLng As Long = (&H80 Or &H1 Or &H4 Or &H20) + (&H8 Or &H40)

'- Pathfinding Constant -
'1 is the old method, faster but not smart at all
'2 is the new method, smart but can slow the server down if maps are huge and alot of npcs have targets.
Public Const PathfindingType As Long = 2

