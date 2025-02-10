Attribute VB_Name = "modCostanti"
Option Explicit

Rem ' *****************************************************************************
Rem ' *****************************************************************************
Rem ' **                                                                         **
Rem ' **                                                                         **
Rem ' ** CONTIENE DEFINIZIONE DI VARIABILI USATE PER CAMBIARE LE PROPRIETA'      **
Rem ' ** DEGLI OGGETTI: QUALI COLORI, COLORI DI SISTEMA, ALLINEAMENTI ED INOLTRE **
Rem ' ** TASTI FUNZIONE                                                          **
Rem ' **                                                                         **
Rem ' *****************************************************************************
Rem ' *****************************************************************************

' Colori
Global Const BLACK = &H0&
Global Const RED = &HFF&
Global Const GREEN = &HFF00&
Global Const YELLOW = &HFFFF&
Global Const BLUE = &HFF0000
Global Const MAGENTA = &HFF00FF
Global Const CYAN = &HFFFF00
Global Const WHITE = &HFFFFFF
Global Const GRAY = &HC0C0C0
Global Const BROWN = &H80&
Global Const VIOLET = &H800080
Global Const DARK_RED = &H80&
Global Const DARK_YELLOW = &H8080&
Global Const DARK_GREEN = &H8000&
Global Const DARK_BLUE = &H800000
Global Const DARK_CYAN = &HC0C000
Global Const DARK_GRAY = &H808080

' ASCII Characters
Global Const ASCII_TAB = 9
Global Const ASCII_LF = 10
Global Const ASCII_CR = 13
Global Const ASCII_ESC = 27

' Align (picture box)
Global Const NONE = 0
Global Const ALIGN_TOP = 1
Global Const ALIGN_BOTTOM = 2

' Alignment
Global Const left1_JUSTIFY = 0  ' 0 - left1 Justify
Global Const right1_JUSTIFY = 1 ' 1 - right1 Justify
Global Const CENTER = 2        ' 2 - Center

' BorderStyle (form)
'Global Const NONE = 0          ' 0 - None
Global Const FIXED_SINGLE = 1   ' 1 - Fixed Single
Global Const SIZABLE = 2        ' 2 - Sizable (Forms only)
Global Const FIXED_DOUBLE = 3   ' 3 - Fixed Double (Forms only)

' MousePointer
Global Const DEFAULT = 0        ' 0 - Default
Global Const ARROW = 1          ' 1 - Arrow
Global Const CROSSHAIR = 2      ' 2 - Cross
Global Const IBEAM = 3          ' 3 - I-Beam
Global Const ICON_POINTER = 4   ' 4 - Icon
Global Const SIZE_POINTER = 5   ' 5 - Size
Global Const SIZE_NE_SW = 6     ' 6 - Size NE SW
Global Const SIZE_N_S = 7       ' 7 - Size N S
Global Const SIZE_NW_SE = 8     ' 8 - Size NW SE
Global Const SIZE_W_E = 9       ' 9 - Size W E
Global Const UP_ARROW = 10      ' 10 - Up Arrow
Global Const HOURGLASS = 11     ' 11 - Hourglass
Global Const NO_DROP = 12       ' 12 - No drop

' DragMode
Global Const MANUAL = 0    ' 0 - Manual
Global Const AUTOMATIC = 1 ' 1 - Automatic

' Arrange Method
' for MDI Forms
Global Const CASCADE = 0
Global Const TILE_HORIZONTAL = 1
Global Const TILE_VERTICAL = 2
Global Const ARRANGE_ICONS = 3

' ScaleMode
Global Const user = 0        ' 0 - User
Global Const TWIPS = 1       ' 1 - Twip
Global Const POINTS = 2      ' 2 - Point
Global Const PIXELS = 3      ' 3 - Pixel
Global Const CHARACTERS = 4  ' 4 - Character
Global Const INCHES = 5      ' 5 - Inch
Global Const MILLIMETERS = 6 ' 6 - Millimeter
Global Const CENTIMETERS = 7 ' 7 - Centimeter

' ScrollBar
' Global Const NONE     = 0 ' 0 - None
Global Const HORIZONTAL = 1 ' 1 - Horizontal
Global Const VERTICAL = 2   ' 2 - Vertical
Global Const BOTH = 3       ' 3 - Both

' WindowState
Global Const NORMAL = 0    ' 0 - Normal
Global Const MINIMIZED = 1 ' 1 - Minimized
Global Const MAXIMIZED = 2 ' 2 - Maximized

' Function Parameters
' MsgBox parameters
Global Const MB_OK = 0                 ' OK button only
Global Const MB_OKCANCEL = 1           ' OK and Cancel buttons
Global Const MB_ABORTRETRYIGNORE = 2   ' Abort, Retry, and Ignore buttons
Global Const MB_YESNOCANCEL = 3        ' Yes, No, and Cancel buttons
Global Const MB_YESNO = 4              ' Yes and No buttons
Global Const MB_RETRYCANCEL = 5        ' Retry and Cancel buttons

Global Const MB_ICONSTOP = 16          ' Critical message
Global Const MB_ICONQUESTION = 32      ' Warning query
Global Const MB_ICONEXCLAMATION = 48   ' Warning message
Global Const MB_ICONINFORMATION = 64   ' Information message

Global Const MB_APPLMODAL = 0          ' Application Modal Message Box
Global Const MB_DEFBUTTON1 = 0         ' First button is default
Global Const MB_DEFBUTTON2 = 256       ' Second button is default
Global Const MB_DEFBUTTON3 = 512       ' Third button is default
Global Const MB_SYSTEMMODAL = 4096      'System Modal

' MsgBox return values
Global Const IDOK = 1                  ' OK button pressed
Global Const IDCANCEL = 2              ' Cancel button pressed
Global Const IDABORT = 3               ' Abort button pressed
Global Const IDRETRY = 4               ' Retry button pressed
Global Const IDIGNORE = 5              ' Ignore button pressed
Global Const IDYES = 6                 ' Yes button pressed
Global Const IDNO = 7                  ' No button pressed

'Grid
'ColAlignment,FixedAlignment Properties
Global Const GRID_ALIGNleft1 = 0
Global Const GRID_ALIGNright1 = 1
Global Const GRID_ALIGNCENTER = 2


'---------------------------------------------------------
'      Table of Contents for Visual Basic Professional
'
'       1.  3-D Controls
'           (Frame/Panel/Option/Check/Command/Group Push)
'       2.  Animated Button
'       3.  Common Dialog Section
'       4.  Gauge Control
'       5.  Graph Control Section
'       6.  Key Status Control
'       7.  Spin Button
'       8.  MCI Control (Multimedia)
'       9.  Masked Edit Control
'       10. Outline
'---------------------------------------------------------


'-------------------------------------------------------------------
'3D Controls
'-------------------------------------------------------------------
'Alignment (Check Box)
Global Const SSCB_TEXT_right1 = 0         '0 - Text to the right1
Global Const SSCB_TEXT_left1 = 1          '1 - Text to the left1

'Alignment (Option Button)
Global Const SSOB_TEXT_right1 = 0         '0 - Text to the right1
Global Const SSOB_TEXT_left1 = 1          '1 - Text to the left1

'Alignment (Frame)
Global Const SSFR_left1_JUSTIFY = 0       '0 - left1 justify text
Global Const SSFR_right1_JUSTIFY = 1      '1 - right1 justify text
Global Const SSFR_CENTER = 2             '2 - Center text

'Alignment (Panel)
Global Const SSPN_left1_TOP = 0           '0 - Text to left1 and top
Global Const SSPN_left1_MIDDLE = 1        '1 - Text to left1 and middle
Global Const SSPN_left1_BOTTOM = 2        '2 - Text to left1 and bottom
Global Const SSPN_right1_TOP = 3          '3 - Text to right1 and top
Global Const SSPN_right1_MIDDLE = 4       '4 - Text to right1 and middle
Global Const SSPN_right1_BOTTOM = 5       '5 - Text to right1 and bottom
Global Const SSPN_CENTER_TOP = 6         '6 - Text to center and top
Global Const SSPN_CENTER_MIDDLE = 7      '7 - Text to center and middle
Global Const SSPN_CENTER_BOTTOM = 8      '8 - Text to center and bottom

'Autosize (Command Button)
Global Const SS_AUTOSIZE_NONE = 0        '0 - No Autosizing
Global Const SSPB_AUTOSIZE_PICTOBUT = 1  '0 - Autosize Picture to Button
Global Const SSPB_AUTOSIZE_BUTTOPIC = 2  '0 - Autosize Button to Picture

'Autosize (Ribbon Button)
'Global Const SS_AUTOSIZE_NONE      = 0  '0 - No Autosizing
Global Const SSRI_AUTOSIZE_PICTOBUT = 1  '0 - Autosize Picture to Button
Global Const SSRI_AUTOSIZE_BUTTOPIC = 2  '0 - Autosize Button to Picture

'Autosize (Panel)
'Global Const SS_AUTOSIZE_NONE    = 0    '0 - No Autosizing
Global Const SSPN_AUTOSIZE_WIDTH = 1     '1 - Autosize Panel width to Caption
Global Const SSPN_AUTOSIZE_HEIGHT = 2    '2 - Autosize Panel height to Caption
Global Const SSPN_AUTOSIZE_CHILD = 3     '3 - Autosize Child to Panel

'BevelInner (Panel)
Global Const SS_BEVELINNER_NONE = 0      '0 - No Inner Bevel
Global Const SS_BEVELINNER_INSET = 1     '1 - Inset Inner Bevel
Global Const SS_BEVELINNER_RAISED = 2    '2 - Raised Inner Bevel

'BevelOuter (Panel)
Global Const SS_BEVELOUTER_NONE = 0      '0 - No Outer Bevel
Global Const SS_BEVELOUTER_INSET = 1     '1 - Inset Outer Bevel
Global Const SS_BEVELOUTER_RAISED = 2    '2 - Raised Outer Bevel

'FloodType (Panel)
Global Const SS_FLOODTYPE_NONE = 0       '0 - No flood
Global Const SS_FLOODTYPE_L_TO_R = 1     '1 - left1 to light
Global Const SS_FLOODTYPE_R_TO_L = 2     '2 - right1 to left1
Global Const SS_FLOODTYPE_T_TO_B = 3     '3 - Top to bottom
Global Const SS_FLOODTYPE_B_TO_T = 4     '4 - Bottom to top
Global Const SS_FLOODTYPE_CIRCLE = 5     '5 - Widening circle

'Font3D (Panel, Command Button, Option Button, Check Box, Frame)
Global Const SS_FONT3D_NONE = 0          '0 - No 3-D text
Global Const SS_FONT3D_RAISED_LIGHT = 1  '1 - Raised with light shading
Global Const SS_FONT3D_RAISED_HEAVY = 2  '2 - Raised with heavy shading
Global Const SS_FONT3D_INSET_LIGHT = 3   '3 - Inset with light shading
Global Const SS_FONT3D_INSET_HEAVY = 4   '4 - Inset with heavy shading

'PictureDnChange (Ribbon Button)
Global Const SS_PICDN_NOCHANGE = 0       '0 - Use 'Up'bitmap with no change
Global Const SS_PICDN_DITHER = 1         '1 - Dither 'Up'bitmap
Global Const SS_PICDN_INVERT = 2         '2 - Invert 'Up'bitmap

'ShadowColor (Panel, Frame)
Global Const SS_SHADOW_DARKGREY = 0      '0 - Dark grey shadow
Global Const SS_SHADOW_BLACK = 1         '1 - Black shadow

'ShadowStyle (Frame)
Global Const SS_SHADOW_INSET = 0         '0 - Shadow inset
Global Const SS_SHADOW_RAISED = 1        '1 - Shadow raised


'---------------------------------------
'Animated Button
'---------------------------------------
'Cycle property
Global Const ANI_ANIMATED = 0
Global Const ANI_MULTISTATE = 1
Global Const ANI_TWO_STATE = 2

'Click Filter property
Global Const ANI_ANYWHERE = 0
Global Const ANI_IMAGE_AND_TEXT = 1
Global Const ANI_IMAGE = 2
Global Const ANI_TEXT = 3

'PicDrawMode Property
Global Const ANI_XPOS_YPOS = 0
Global Const ANI_AUTOSIZE = 1
Global Const ANI_STRETCH = 2

'SpecialOp Property
Global Const ANI_CLICK = 1

'TextPosition Property
Global Const ANI_CENTER = 0
Global Const ANI_left1 = 1
Global Const ANI_right1 = 2
Global Const ANI_BOTTON = 3
Global Const ANI_TOP = 4



'-------------------------------------------------------------------
'Common Dialog
'-------------------------------------------------------------------

'Action Property
Global Const DLG_FILE_OPEN = 1
Global Const DLG_FILE_SAVE = 2
Global Const DLG_COLOR = 3
Global Const DLG_FONT = 4
Global Const DLG_PRINT = 5
Global Const DLG_HELP = 6

'File Open/Save Dialog Flags
Global Const OFN_READONLY = &H1&
Global Const OFN_OVERWRITEPROMPT = &H2&
Global Const OFN_HIDEREADONLY = &H4&
Global Const OFN_NOCHANGEDIR = &H8&
Global Const OFN_SHOWHELP = &H10&
Global Const OFN_NOVALIDATE = &H100&
Global Const OFN_ALLOWMULTISELECT = &H200&
Global Const OFN_EXTENTIONDIFFERENT = &H400&
Global Const OFN_PATHMUSTEXIST = &H800&
Global Const OFN_FILEMUSTEXIST = &H1000&
Global Const OFN_CREATEPROMPT = &H2000&
Global Const OFN_SHAREAWARE = &H4000&
Global Const OFN_NOREADONLYRETURN = &H8000&

'Color Dialog Flags
Global Const CC_RGBINIT = &H1&
Global Const CC_FULLOPEN = &H2&
Global Const CC_PREVENTFULLOPEN = &H4&
Global Const CC_SHOWHELP = &H8&

'Fonts Dialog Flags
Global Const CF_SCREENFONTS = &H1&
Global Const CF_PRINTERFONTS = &H2&
Global Const CF_BOTH = &H3&
Global Const CF_SHOWHELP = &H4&
Global Const CF_INITTOLOGFONTSTRUCT = &H40&
Global Const CF_USESTYLE = &H80&
Global Const CF_EFFECTS = &H100&
Global Const CF_APPLY = &H200&
Global Const CF_ANSIONLY = &H400&
Global Const CF_NOVECTORFONTS = &H800&
Global Const CF_NOSIMULATIONS = &H1000&
Global Const CF_LIMITSIZE = &H2000&
Global Const CF_FIXEDPITCHONLY = &H4000&
Global Const CF_WYSIWYG = &H8000&                        'must also have CF_SCREENFONTS & CF_formFONTS
Global Const CF_FORCEFONTEXIST = &H10000
Global Const CF_SCALABLEONLY = &H20000
Global Const CF_TTONLY = &H40000
Global Const CF_NOFACESEL = &H80000
Global Const CF_NOSTYLESEL = &H100000
Global Const CF_NOSIZESEL = &H200000

'form Dialog Flags
Global Const PD_ALLPAGES = &H0&
Global Const PD_SELECTION = &H1&
Global Const PD_PAGENUMS = &H2&
Global Const PD_NOSELECTION = &H4&
Global Const PD_NOPAGENUMS = &H8&
Global Const PD_COLLATE = &H10&
Global Const PD_PRINTTOFILE = &H20&
Global Const PD_PRINTSETUP = &H40&
Global Const PD_NOWARNING = &H80&
Global Const PD_RETURNDC = &H100&
Global Const PD_RETURNIC = &H200&
Global Const PD_RETURNDEFAULT = &H400&
Global Const PD_SHOWHELP = &H800&
Global Const PD_USEDEVMODECOPIES = &H40000
Global Const PD_DISABLEPRINTTOFILE = &H80000
Global Const PD_HIDEPRINTTOFILE = &H100000

'Help Constants
Global Const HELP_CONTEXT = &H1           'Display topic in ulTopic
Global Const HELP_QUIT = &H2              'Terminate help
Global Const HELP_INDEX = &H3             'Display index
Global Const HELP_CONTENTS = &H3
Global Const HELP_HELPONHELP = &H4        'Display help on using help
Global Const HELP_SETINDEX = &H5          'Set the current Index for multi index help
Global Const HELP_SETCONTENTS = &H5
Global Const HELP_CONTEXTPOPUP = &H8
Global Const HELP_FORCEFILE = &H9
Global Const HELP_KEY = &H101             'Display topic for keyword in offabData
Global Const HELP_COMMAND = &H102
Global Const HELP_PARTIALKEY = &H105      'call the search engine in winhelp

'Error Constants
Global Const CDERR_DIALOGFAILURE = &HFFFF

Global Const CDERR_GENERALCODES = &H0
Global Const CDERR_STRUCTSIZE = &H1
Global Const CDERR_INITIALIZATION = &H2
Global Const CDERR_NOTEMPLATE = &H3
Global Const CDERR_NOHINSTANCE = &H4
Global Const CDERR_LOADSTRFAILURE = &H5
Global Const CDERR_FINDRESFAILURE = &H6
Global Const CDERR_LOADRESFAILURE = &H7
Global Const CDERR_LOCKRESFAILURE = &H8
Global Const CDERR_MEMALLOCFAILURE = &H9
Global Const CDERR_MEMLOCKFAILURE = &HA
Global Const CDERR_NOHOOK = &HB

'Added for CMDLG.VBX
Global Const CDERR_CANCEL = &HC
Global Const CDERR_NODLL = &HD
Global Const CDERR_ERRPROC = &HE
Global Const CDERR_ALLOC = &HF
Global Const CDERR_HELP = &H10

Global Const PDERR_PRINTERCODES = &H1000
Global Const PDERR_SETUPFAILURE = &H1001
Global Const PDERR_PARSEFAILURE = &H1002
Global Const PDERR_RETDEFFAILURE = &H1003
Global Const PDERR_LOADDRVFAILURE = &H1004
Global Const PDERR_GETDEVMODEFAIL = &H1005
Global Const PDERR_INITFAILURE = &H1006
Global Const PDERR_NODEVICES = &H1007
Global Const PDERR_NODEFAULTPRN = &H1008
Global Const PDERR_DNDMMISMATCH = &H1009
Global Const PDERR_CREATEICFAILURE = &H100A
Global Const PDERR_PRINTERNOTFOUND = &H100B

Global Const CFERR_CHOOSEFONTCODES = &H2000
Global Const CFERR_NOFONTS = &H2001

Global Const FNERR_FILENAMECODES = &H3000
Global Const FNERR_SUBCLASSFAILURE = &H3001
Global Const FNERR_INVALIDFILENAME = &H3002
Global Const FNERR_BUFFERTOOSMALL = &H3003

Global Const FRERR_FINDREPLACECODES = &H4000
Global Const CCERR_CHOOSECOLORCODES = &H5000


'---------------------------------------
'Gauge
'---------------------------------------
'Style Property
Global Const GAUGE_HORIZ = 0
Global Const GAUGE_VERT = 1
Global Const GAUGE_SEMI = 2
Global Const GAUGE_FULL = 3


'----------------------------------------
'Graph Control
'----------------------------------------
'General
Global Const G_NONE = 0
Global Const G_DEFAULT = 0

Global Const G_OFF = 0
Global Const G_ON = 1

Global Const G_MONO = 0
Global Const G_COLOR = 1

'Graph Types
Global Const G_PIE2D = 1
Global Const G_PIE3D = 2
Global Const G_BAR2D = 3
Global Const G_BAR3D = 4
Global Const G_GANTT = 5
Global Const G_LINE = 6
Global Const G_LOGLIN = 7
Global Const G_AREA = 8
Global Const G_SCATTER = 9
Global Const G_POLAR = 10
Global Const G_HLC = 11

'Colors
Global Const G_BLACK = 0
Global Const G_BLUE = 1
Global Const G_GREEN = 2
Global Const G_CYAN = 3
Global Const G_RED = 4
Global Const G_MAGENTA = 5
Global Const G_BROWN = 6
Global Const G_LIGHT_GRAY = 7
Global Const G_DARK_GRAY = 8
Global Const G_LIGHT_BLUE = 9
Global Const G_LIGHT_GREEN = 10
Global Const G_LIGHT_CYAN = 11
Global Const G_LIGHT_RED = 12
Global Const G_LIGHT_MAGENTA = 13
Global Const G_YELLOW = 14
Global Const G_WHITE = 15
Global Const G_AUTOBW = 16

'Patterns
Global Const G_SOLID = 0
Global Const G_HOLLOW = 1
Global Const G_HATCH1 = 2
Global Const G_HATCH2 = 3
Global Const G_HATCH3 = 4
Global Const G_HATCH4 = 5
Global Const G_HATCH5 = 6
Global Const G_HATCH6 = 7
Global Const G_BITMAP1 = 16
Global Const G_BITMAP2 = 17
Global Const G_BITMAP3 = 18
Global Const G_BITMAP4 = 19
Global Const G_BITMAP5 = 20
Global Const G_BITMAP6 = 21
Global Const G_BITMAP7 = 22
Global Const G_BITMAP8 = 23
Global Const G_BITMAP9 = 24
Global Const G_BITMAP10 = 25
Global Const G_BITMAP11 = 26
Global Const G_BITMAP12 = 27
Global Const G_BITMAP13 = 28
Global Const G_BITMAP14 = 29
Global Const G_BITMAP15 = 30
Global Const G_BITMAP16 = 31

'Symbols
Global Const G_CROSS_PLUS = 0
Global Const G_CROSS_TIMES = 1
Global Const G_TRIANGLE_UP = 2
Global Const G_SOLID_TRIANGLE_UP = 3
Global Const G_TRIANGLE_DOWN = 4
Global Const G_SOLID_TRIANGLE_DOWN = 5
Global Const G_SQUARE = 6
Global Const G_SOLID_SQUARE = 7
Global Const G_DIAMOND = 8
Global Const G_SOLID_DIAMOND = 9

'Line Styles
'Global Const G_SOLID = 0
Global Const G_DASH = 1
Global Const G_DOT = 2
Global Const G_DASHDOT = 3
Global Const G_DASHDOTDOT = 4

'Grids
Global Const G_HORIZONTAL = 1
Global Const G_VERTICAL = 2

'Statistics
Global Const G_MEAN = 1
Global Const G_MIN_MAX = 2
Global Const G_STD_DEV = 4
Global Const G_BEST_FIT = 8

'Data Arrays
Global Const G_GRAPH_DATA = 1
Global Const G_COLOR_DATA = 2
Global Const G_EXTRA_DATA = 3
Global Const G_LABEL_TEXT = 4
Global Const G_LEGEND_TEXT = 5
Global Const G_PATTERN_DATA = 6
Global Const G_SYMBOL_DATA = 7
Global Const G_XPOS_DATA = 8
Global Const G_ALL_DATA = 9

'Draw Mode
Global Const G_NO_ACTION = 0
Global Const G_CLEAR = 1
Global Const G_DRAW = 2
Global Const G_BLIT = 3
Global Const G_COPY = 4
Global Const G_PRINT = 5
Global Const G_WRITE = 6

'Print Options
Global Const G_BORDER = 2

'Pie Chart Options             '
Global Const G_NO_LINES = 1
Global Const G_COLORED = 2
Global Const G_PERCENTS = 4

'Bar Chart Options             '
'Global Const G_HORIZONTAL = 1
Global Const G_STACKED = 2
Global Const G_PERCENTAGE = 4
Global Const G_Z_CLUSTERED = 6

'Gantt Chart Options           '
Global Const G_SPACED_BARS = 1

'Line/Polar Chart Options      '
Global Const G_SYMBOLS = 1
Global Const G_STICKS = 2
Global Const G_LINES = 4

'Area Chart Options            '
Global Const G_ABSOLUTE = 1
Global Const G_PERCENT = 2

'HLC Chart Options             '
Global Const G_NO_CLOSE = 1
Global Const G_NO_HIGH_LOW = 2


'---------------------------------------
'Key Status Control
'---------------------------------------
'Style
Global Const KEYSTAT_CAPSLOCK = 0
Global Const KEYSTAT_NUMLOCK = 1
Global Const KEYSTAT_INSERT = 2
Global Const KEYSTAT_SCROLLLOCK = 3


'---------------------------------------
'MCI Control (Multimedia)
'---------------------------------------
'Mode Property
Global Const MCI_MODE_NOT_OPEN = 11
Global Const MCI_MODE_STOP = 12
Global Const MCI_MODE_PLAY = 13
Global Const MCI_MODE_RECORD = 14
Global Const MCI_MODE_SEEK = 15
Global Const MCI_MODE_PAUSE = 16
Global Const MCI_MODE_READY = 17

'NotifyValue Property
Global Const MCI_NOTIFY_SUCCESSFUL = 1
Global Const MCI_NOTIFY_SUPERSEDED = 2
Global Const MCI_ABORTED = 4
Global Const MCI_FAILURE = 8

'Orientation Property
Global Const MCI_ORIENT_HORZ = 0
Global Const MCI_ORIENT_VERT = 1

'RecordMode Porperty
Global Const MCI_RECORD_INSERT = 0
Global Const MCI_RECORD_OVERWRITE = 1

'TimeFormat Property
Global Const MCI_FORMAT_MILLISECONDS = 0
Global Const MCI_FORMAT_HMS = 1
Global Const MCI_FORMAT_FRAMES = 3
Global Const MCI_FORMAT_SMPTE_24 = 4
Global Const MCI_FORMAT_SMPTE_25 = 5
Global Const MCI_FORMAT_SMPTE_30 = 6
Global Const MCI_FORMAT_SMPTE_30DROP = 7
Global Const MCI_FORMAT_BYTES = 8
Global Const MCI_FORMAT_SAMPLES = 9
Global Const MCI_FORMAT_TMSF = 10


'---------------------------------------
'Spin Button
'---------------------------------------
'SpinOrientation
Global Const SPIN_VERTICAL = 0
Global Const SPIN_HORIZONTAL = 1


'---------------------------------------
'Masked Edit Control
'---------------------------------------
'ClipMode
Global Const ME_INCLIT = 0
Global Const ME_EXCLIT = 1


'-------------------------------------------------
'  Outline
'-------------------------------------------------
' PictureType
Global Const MSOUTLINE_PICTURE_CLOSED = 0
Global Const MSOUTLINE_PICTURE_OPEN = 1
Global Const MSOUTLINE_PICTURE_LEAF = 2

'Outline Control Error Constants
Global Const MSOUTLINE_BADPICFORMAT = 32000
Global Const MSOUTLINE_BADINDENTATION = 32001
Global Const MSOUTLINE_MEM = 32002
Global Const MSOUTLINE_PARENTNOTEXPANDED = 32003


'========================================================================

' Dichiarazione funzione per form sempre ON TOP e relative costanti
Declare Function SetWindowPos Lib "user32" (ByVal h As Long, ByVal hb As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal cy As Long, ByVal f As Long) As Long
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
' Utilizzo :
' per mettere una form xxx sempre in primo piano:
' succes%=SetWindowPos (xxxx.hwnd,HWND_TOPMOST,0,0,0,0,FLAGS)
' SUCCES% è <> da 0 quando è OK
' per disattivare l'opzione primo piano si usa al posto di TOPMOST NOTOPMOST

'Alby Maggio 2014
Global Const LIMITE_MEDIA_GIORNALIERA = 0
Global Const LIMITE_MEDIA_ORARIA = 1

Global Const SOGLIA_ATTENZIONE = 1
Global Const SOGLIA_ALLARME = 2
Global Const DATO_AUSILIARIO = 3
'luca 07/10/2015
Global Const SOGLIA_ATTENZIONE_GIORNALIERA = 4
Global Const SOGLIA_ALLARME_GIORNALIERA = 5

Global Const SUPERO_GIORNO_IN_CORSO = 0
Global Const SUPERO_GIORNO_ULTIMO = 1
Global Const SUPERO_ORA_IN_CORSO = 2
Global Const SUPERO_ORA_ULTIMO = 3
Global Const SUPERO_ISTANTANEO = 4
Global Const SUPERO_MINUTO_IN_CORSO = 5
Global Const SUPERO_MINUTO_ULTIMO = 6

'Federica novembre 2017
Public Const TIPO_MEDIE_ORARIE = 1
Public Const TIPO_MEDIE_SEMIORARIE = 2
