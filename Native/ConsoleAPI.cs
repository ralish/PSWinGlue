/*
    Implements P/Invoke signatures for the majority of the Windows
    Console API. Still need to add the Read* and Write* functions.

    https://docs.microsoft.com/en-us/windows/console/console-functions
*/

[Flags]
public enum ConsoleDisplayModeGetFlags
{
    CONSOLE_FULLSCREEN_MODE = 1,
    CONSOLE_WINDOWED_MODE   = 2
}

[Flags]
public enum ConsoleDisplayModeSetFlags
{
    CONSOLE_FULLSCREEN          = 1,
    CONSOLE_FULLSCREEN_HARDWARE = 2
}

[Flags]
public enum ConsoleModeInputFlags
{
    ENABLE_PROCESSED_INPUT          = 0x001,
    ENABLE_LINE_INPUT               = 0x002,
    ENABLE_ECHO_INPUT               = 0x004,
    ENABLE_WINDOW_INPUT             = 0x008,
    ENABLE_MOUSE_INPUT              = 0x010,
    ENABLE_INSERT_MODE              = 0x020,
    ENABLE_QUICK_EDIT_MODE          = 0x040,
    ENABLE_EXTENDED_FLAGS           = 0x080,
    ENABLE_AUTO_POSITION            = 0x100,
    ENABLE_VIRTUAL_TERMINAL_INPUT   = 0x200
}

[Flags]
public enum ConsoleModeOutputFlags
{
    ENABLE_PROCESSED_OUTPUT             = 0x01,
    ENABLE_WRAP_AT_EOL_OUTPUT           = 0x02,
    ENABLE_VIRTUAL_TERMINAL_PROCESSING  = 0x04,
    DISABLE_NEWLINE_AUTO_RETURN         = 0x08,
    ENABLE_LVB_GRID_WORLDWIDE           = 0x10
}

[Flags]
public enum ConsoleScreenBufferFlags {
    CONSOLE_TEXTMODE_BUFFER = 1
}

[Flags]
public enum HandlerRoutineCtrlTypes
{
    CTRL_C_EVENT        = 0,
    CTRL_BREAK_EVENT    = 1,
    CTRL_CLOSE_EVENT    = 2,
    CTRL_LOGOFF_EVENT   = 5,
    CTRL_SHUTDOWN_EVENT = 6,
}

[Flags]
public enum PseudoConsoleFlags {
    PSEUDOCONSOLE_INHERIT_CURSOR = 1
}

[Flags]
public enum StdHandleDevices
{
    STD_INPUT_HANDLE    = -10,
    STD_OUTPUT_HANDLE   = -11,
    STD_ERROR_HANDLE    = -12
}

[StructLayout(LayoutKind.Sequential)]
public struct COLORREF
{
    public byte R;
    public byte G;
    public byte B;
}

[StructLayout(LayoutKind.Sequential)]
public struct CONSOLE_CURSOR_INFO
{
    public uint dwSize;
    public bool bVisible;
}

[StructLayout(LayoutKind.Sequential)]
public struct CONSOLE_FONT_INFO
{
    public uint nFont;
    public COORD dwFontSize;
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct CONSOLE_FONT_INFOEX
{
    public uint cbSize;
    public uint nFont;
    public COORD dwFontSize;
    public uint FontFamily;
    public uint FontWeight;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 32)]
    public char[] FaceName;
}
[StructLayout(LayoutKind.Sequential)]
public struct CONSOLE_HISTORY_INFO
{
    public uint cbSize;
    public uint HistoryBufferSize;
    public uint NumberOfHistoryBuffers;
    public uint dwFlags;
}

[StructLayout(LayoutKind.Sequential)]
public struct CONSOLE_SCREEN_BUFFER_INFO
{
    public COORD dwSize;
    public COORD dwCursorPosition;
    public ushort wAttributes;
    public SMALL_RECT srWindow;
    public COORD dwMaximumWindowSize;
}

[StructLayout(LayoutKind.Sequential)]
public struct CONSOLE_SCREEN_BUFFER_INFOEX
{
    public uint cbSize;
    public COORD dwSize;
    public COORD dwCursorPosition;
    public ushort wAttributes;
    public SMALL_RECT srWindow;
    public COORD dwMaximumWindowSize;
    public ushort wPopupAttributes;
    public bool bFullscreenSupported;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 16)]
    public COLORREF[] ColorTable;
}

[StructLayout(LayoutKind.Sequential)]
public struct CONSOLE_SELECTION_INFO
{
    public uint dwFlags;
    public COORD dwSelectionAnchor;
    public SMALL_RECT srSelection;
}

[StructLayout(LayoutKind.Sequential)]
public struct COORD
{
    public short X;
    public short Y;
}

[StructLayout(LayoutKind.Sequential)]
public struct FOCUS_EVENT_RECORD
{
    public bool bSetFocus;
}

[StructLayout(LayoutKind.Explicit)]
public struct INPUT_RECORD
{
    [FieldOffset(0)]
    public ushort EventType;
    [FieldOffset(4)]
    public KEY_EVENT_RECORD KeyEvent;
    [FieldOffset(4)]
    public MOUSE_EVENT_RECORD MouseEvent;
    [FieldOffset(4)]
    public WINDOW_BUFFER_SIZE_RECORD WindowBufferSizeRecord;
    [FieldOffset(4)]
    public MENU_EVENT_RECORD MenuEvent;
    [FieldOffset(4)]
    public FOCUS_EVENT_RECORD FocusEvent;
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct KEY_EVENT_RECORD
{
    public bool bKeyDown;
    public ushort wRepeatCount;
    public ushort wVirtualKeyCode;
    public ushort wVirtualScanCode;
    public char cChar;
    public uint dwControlKeyState;
}

[StructLayout(LayoutKind.Sequential)]
public struct MENU_EVENT_RECORD
{
    public uint dwCommandId;
}

[StructLayout(LayoutKind.Sequential)]
public struct MOUSE_EVENT_RECORD
{
    public COORD dwMousePosition;
    public uint dwButtonState;
    public uint dwControlKeyState;
    public uint dwEventFlags;
}

[StructLayout(LayoutKind.Sequential)]
public struct SMALL_RECT
{
    public short Left;
    public short Top;
    public short Right;
    public short Bottom;
}

[StructLayout(LayoutKind.Sequential)]
public struct WINDOW_BUFFER_SIZE_RECORD
{
    public COORD dwSize;
}

[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public extern static bool AddConsoleAlias(
    [MarshalAs(UnmanagedType.LPTStr)] string Source,
    [MarshalAs(UnmanagedType.LPTStr)] string Target,
    [MarshalAs(UnmanagedType.LPTStr)] string ExeName
);

[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public extern static bool AddConsoleAlias(
    [MarshalAs(UnmanagedType.LPTStr)] string Source,
    IntPtr Target,
    [MarshalAs(UnmanagedType.LPTStr)] string ExeName
);

[DllImport("kernel32.dll", SetLastError = true)]
public extern static bool AllocConsole();

[DllImport("kernel32.dll", SetLastError = true)]
public extern static bool AttachConsole(
    int dwProcessId
);

[DllImport("kernel32.dll")]
public extern static void ClosePseudoConsole(
    IntPtr hPC
);

[DllImport("kernel32.dll")]
public extern static int CreatePseudoConsole(
    COORD size,
    IntPtr hInput,
    IntPtr hOutput,
    uint dwFlags,
    out IntPtr phPC
);

[DllImport("kernel32.dll", SetLastError = true)]
public extern static int CreateConsoleScreenBuffer(
    int dwDesiredAccess,
    int dwShareMode,
    IntPtr lpSecurityAttributes,
    uint dwFlags,
    IntPtr lpScreenBufferData
);

[DllImport("kernel32.dll", SetLastError = true)]
public extern static bool FillConsoleOutputAttribute(
    IntPtr hConsoleOutput,
    ushort wAttribute,
    uint nLength,
    COORD dwWriteCoord,
    out uint lpNumberOfAttrsWritten
);

[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public extern static bool FillConsoleOutputCharacter(
    IntPtr hConsoleOutput,
    char cCharacter,
    uint nLength,
    COORD dwWriteCoord,
    out uint lpNumberOfCharsWritten
);

[DllImport("kernel32.dll", SetLastError = true)]
public extern static bool FlushConsoleInputBuffer(
    IntPtr hConsoleInput
);

[DllImport("kernel32.dll", SetLastError = true)]
public extern static bool FreeConsole();

[DllImport("kernel32.dll", SetLastError = true)]
public extern static bool GenerateConsoleCtrlEvent(
    uint dwCtrlEvent,
    uint dwProcessGroupID
);

[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern int GetConsoleAlias(
    [MarshalAs(UnmanagedType.LPTStr)] string lpSource,
    System.Text.StringBuilder lpTargetBuffer,
    uint TargetBufferLength,
    [MarshalAs(UnmanagedType.LPTStr)] string lpExeName
);

[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern int GetConsoleAliases(
    System.Text.StringBuilder lpAliasBuffer,
    uint AliasBufferLength,
    [MarshalAs(UnmanagedType.LPTStr)] string lpExeName
);

[DllImport("kernel32.dll", CharSet = CharSet.Auto)]
public static extern int GetConsoleAliasesLength(
    [MarshalAs(UnmanagedType.LPTStr)] string lpExeName
);

[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern int GetConsoleAliasExes(
    System.Text.StringBuilder lpExeNameBuffer,
    uint ExeNameBufferLength
);

[DllImport("kernel32.dll", CharSet = CharSet.Auto)]
public static extern int GetConsoleAliasExesLength();

[DllImport("kernel32.dll", SetLastError = true)]
public static extern int GetConsoleCP();

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool GetConsoleCursorInfo(
    IntPtr hConsoleOutput,
    out CONSOLE_CURSOR_INFO lpConsoleCursorInfo
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool GetConsoleDisplayMode(
    out uint lpModeFlags
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern COORD GetConsoleFontSize(
    IntPtr hConsoleOutput,
    uint nFont
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool GetConsoleHistoryInfo(
    out CONSOLE_HISTORY_INFO lpConsoleHistoryInfo
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool GetConsoleMode(
    IntPtr hConsoleHandle,
    out uint lpMode
);

[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern int GetConsoleOriginalTitle(
    System.Text.StringBuilder lpConsoleTitle,
    uint nSize
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern int GetConsoleOutputCP();

[DllImport("kernel32.dll", SetLastError = true)]
public static extern uint GetConsoleProcessList(
    out int lpdwProcessList, // LPDWORD
    uint dwProcessCount
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool GetConsoleScreenBufferInfo(
    IntPtr hConsoleOutput,
    out CONSOLE_SCREEN_BUFFER_INFO lpConsoleScreenBufferInfo
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool GetConsoleScreenBufferInfoEx(
    IntPtr hConsoleOutput,
    out CONSOLE_SCREEN_BUFFER_INFOEX lpConsoleScreenBufferInfoEx
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool GetConsoleSelectionInfo(
    out CONSOLE_SELECTION_INFO lpConsoleSelectionInfo
);

[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern int GetConsoleTitle(
    System.Text.StringBuilder lpConsoleTitle,
    uint nSize
);

[DllImport("kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool GetCurrentConsoleFont(
    IntPtr hConsoleOutput,
    bool bMaximumWindow,
    out CONSOLE_FONT_INFO lpConsoleCurrentFont
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool GetCurrentConsoleFontEx(
    IntPtr hConsoleOutput,
    bool bMaximumWindow,
    out CONSOLE_FONT_INFOEX lpConsoleCurrentFontEx
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern COORD GetLargestConsoleWindowSize(
    IntPtr hConsoleOutput
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool GetNumberOfConsoleInputEvents(
    IntPtr hConsoleInput,
    out uint lpcNumberOfEvents
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool GetNumberOfConsoleMouseButtons(
    out uint lpNumberOfMouseButtons
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern IntPtr GetStdHandle(
    int nStdHandle
);

[DllImport("kernel32.dll")]
public static extern bool HandlerRoutine(
    uint dwCtrlType
);

[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern bool PeekConsoleInput(
    IntPtr hConsoleInput,
    [MarshalAs(UnmanagedType.LPArray)]
    out INPUT_RECORD[] lpBuffer,
    uint nLength,
    out uint lpNumberOfEventsRead
);

[DllImport("kernel32.dll")]
public static extern int ResizePseudoConsole(
    IntPtr hPC,
    COORD size
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetConsoleActiveScreenBuffer(
    IntPtr hConsoleOutput
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetConsoleCP(
    uint wCodePageID
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetConsoleCtrlHandler(
    IntPtr HandlerRoutine,
    bool Add
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetConsoleCursorInfo(
    IntPtr hConsoleOutput,
    CONSOLE_CURSOR_INFO lpConsoleCursorInfo
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetConsoleCursorPosition(
    IntPtr hConsoleOutput,
    COORD dwCursorPosition
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetConsoleDisplayMode(
    IntPtr hConsoleOutput,
    uint dwFlags
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetConsoleDisplayMode(
    IntPtr hConsoleOutput,
    uint dwFlags,
    out COORD lpNewScreenBufferDimensions
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetConsoleHistoryInfo(
    CONSOLE_HISTORY_INFO lpConsoleHistoryInfo
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetConsoleMode(
    IntPtr hConsoleHandle,
    uint dwMode
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetConsoleOutputCP(
    uint wCodePageID
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetConsoleScreenBufferInfoEx(
    IntPtr hConsoleOutput,
    CONSOLE_SCREEN_BUFFER_INFOEX lpConsoleScreenBufferInfoEx
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetConsoleScreenBufferSize(
    IntPtr hConsoleOutput,
    COORD dwSize
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetConsoleTextAttribute(
    IntPtr hConsoleOutput,
    ushort wAttributes
);

[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern bool SetConsoleTitle(
    [MarshalAs(UnmanagedType.LPTStr)] string lpConsoleTitle
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetConsoleWindowInfo(
    IntPtr hConsoleOutput,
    bool bAbsolute,
    SMALL_RECT lpConsoleWindow
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetCurrentConsoleFontEx(
    IntPtr hConsoleOutput,
    bool bMaximumWindow,
    CONSOLE_FONT_INFOEX lpConsoleCurrentFontEx
);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetStdHandle(
    int nStdHandle,
    IntPtr hHandle
);
