/*
    Implements P/Invoke signatures for the Windows Console API
    https://docs.microsoft.com/en-us/windows/console/console-functions
*/

using System;
using System.Runtime.InteropServices;
using System.Text;

namespace PSWinGlue {
    public static class ConsoleAPI {
        #region Constants

        // AttachConsole
        public const uint ATTACH_PARENT_PROCESS = 4294967295; // -1

        // CONSOLE_FONT_INFOEX
        private const int LF_FACESIZE = 32;

        #endregion

        #region Delegates

        public delegate bool HandlerRoutine(ControlEvent dwCtrlType);

        #endregion

        #region Enumerations

        [Flags]
        public enum AccessRightsFlags : uint {
            GENERIC_WRITE = 0x40000000,
            GENERIC_READ  = 0x80000000
        }

        [Flags]
        public enum CharacterAttributes : ushort {
            FOREGROUND_BLUE            = 0x1,
            FOREGROUND_GREEN           = 0x2,
            FOREGROUND_RED             = 0x4,
            FOREGROUND_INTENSITY       = 0x8,
            BACKGROUND_BLUE            = 0x10,
            BACKGROUND_GREEN           = 0x20,
            BACKGROUND_RED             = 0x40,
            BACKGROUND_INTENSITY       = 0x80,
            COMMON_LVB_LEADING_BYTE    = 0x100,
            COMMON_LVB_TRAILING_BYTE   = 0x200,
            COMMON_LVB_GRID_HORIZONTAL = 0x400,
            COMMON_LVB_GRID_LVERTICAL  = 0x800,
            COMMON_LVB_GRID_RVERTICAL  = 0x1000,
            COMMON_LVB_REVERSE_VIDEO   = 0x4000,
            COMMON_LVB_UNDERSCORE      = 0x8000
        }

        public enum ControlEvent : uint {
            CTRL_C_EVENT        = 0,
            CTRL_BREAK_EVENT    = 1,
            CTRL_CLOSE_EVENT    = 2,
            CTRL_LOGOFF_EVENT   = 5,
            CTRL_SHUTDOWN_EVENT = 6
        }

        [Flags]
        public enum ControlKeyStates : uint {
            RIGHT_ALT_PRESSED  = 0x1,
            LEFT_ALT_PRESSED   = 0x2,
            RIGHT_CTRL_PRESSED = 0x4,
            LEFT_CTRL_PRESSED  = 0x8,
            SHIFT_PRESSED      = 0x10,
            NUMLOCK_ON         = 0x20,
            SCROLLLOCK_ON      = 0x40,
            CAPSLOCK_ON        = 0x80,
            ENHANCED_KEY       = 0x100
        }

        [Flags]
        public enum DisplayModeGetFlags : uint {
            CONSOLE_FULLSCREEN          = 0x1,
            CONSOLE_FULLSCREEN_HARDWARE = 0x2
        }

        [Flags]
        public enum DisplayModeSetFlags : uint {
            CONSOLE_FULLSCREEN_MODE = 0x1,
            CONSOLE_WINDOWED_MODE   = 0x2
        }

        [Flags]
        public enum EventType : ushort {
            KEY_EVENT                = 0x1,
            MOUSE_EVENT              = 0x2,
            WINDOW_BUFFER_SIZE_EVENT = 0x4,
            MENU_EVENT               = 0x8,
            FOCUS_EVENT              = 0x10
        }

        [Flags]
        public enum HistoryInfoFlags : uint {
            HISTORY_NO_DUP_FLAG = 0x1
        }

        [Flags]
        public enum InputModeFlags : uint {
            ENABLE_PROCESSED_INPUT        = 0x1,
            ENABLE_LINE_INPUT             = 0x2,
            ENABLE_ECHO_INPUT             = 0x4,
            ENABLE_WINDOW_INPUT           = 0x8,
            ENABLE_MOUSE_INPUT            = 0x10,
            ENABLE_INSERT_MODE            = 0x20,
            ENABLE_QUICK_EDIT_MODE        = 0x40,
            ENABLE_EXTENDED_FLAGS         = 0x80,
            ENABLE_AUTO_POSITION          = 0x100,
            ENABLE_VIRTUAL_TERMINAL_INPUT = 0x200
        }

        [Flags]
        public enum MouseButtonStates : uint {
            FROM_LEFT_1ST_BUTTON_PRESSED = 0x1,
            RIGHTMOST_BUTTON_PRESSED     = 0x2,
            FROM_LEFT_2ND_BUTTON_PRESSED = 0x4,
            FROM_LEFT_3RD_BUTTON_PRESSED = 0x8,
            FROM_LEFT_4TH_BUTTON_PRESSED = 0x10
        }

        [Flags]
        public enum MouseEventFlags : uint {
            MOUSE_MOVED    = 0x1,
            DOUBLE_CLICK   = 0x2,
            MOUSE_WHEELED  = 0x4,
            MOUSE_HWHEELED = 0x8
        }

        [Flags]
        public enum PseudoConsoleFlags : uint {
            PSEUDOCONSOLE_INHERIT_CURSOR = 0x1
        }

        [Flags]
        public enum ScreenBufferFlags : uint {
            CONSOLE_TEXTMODE_BUFFER = 0x1
        }

        [Flags]
        public enum ScreenBufferModeFlags : uint {
            ENABLE_PROCESSED_OUTPUT            = 0x1,
            ENABLE_WRAP_AT_EOL_OUTPUT          = 0x2,
            ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x4,
            DISABLE_NEWLINE_AUTO_RETURN        = 0x8,
            ENABLE_LVB_GRID_WORLDWIDE          = 0x10
        }

        [Flags]
        public enum SelectionInfoFlags : uint {
            CONSOLE_NO_SELECTION          = 0x0,
            CONSOLE_SELECTION_IN_PROGRESS = 0x1,
            CONSOLE_SELECTION_NOT_EMPTY   = 0x2,
            CONSOLE_MOUSE_SELECTION       = 0x4,
            CONSOLE_MOUSE_DOWN            = 0x8
        }

        [Flags]
        public enum ShareModeFlags : uint {
            FILE_SHARE_READ  = 0x1,
            FILE_SHARE_WRITE = 0x2
        }

        public enum StandardDevice : uint {
            STD_INPUT_HANDLE  = 4294967286, // -10
            STD_OUTPUT_HANDLE = 4294967285, // -11
            STD_ERROR_HANDLE  = 4294967284  // -12
        }

        #endregion

        #region Functions

        // Introduced in Windows XP
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "AddConsoleAliasW", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool AddConsoleAlias(
            [MarshalAs(UnmanagedType.LPWStr)]
            string Source,

            [MarshalAs(UnmanagedType.LPWStr)]
            string Target,

            [MarshalAs(UnmanagedType.LPWStr)]
            string ExeName
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool AllocConsole();

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool AttachConsole(
            uint dwProcessId
        );

        // Introduced in Windows 10 version 1809 / Windows Server 2019
        [DllImport("kernel32.dll")]
        public static extern void ClosePseudoConsole(
            IntPtr hPC
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr CreateConsoleScreenBuffer(
            AccessRightsFlags dwDesiredAccess,
            ShareModeFlags dwShareMode,
            IntPtr lpSecurityAttributes, // TODO
            ScreenBufferFlags dwFlags,
            IntPtr lpScreenBufferData // NULL
        );

        // Introduced in Windows 10 version 1809 / Windows Server 2019
        [DllImport("kernel32.dll")]
        public static extern int CreatePseudoConsole(
            COORD size,
            IntPtr hInput,
            IntPtr hOutput,
            PseudoConsoleFlags dwFlags,
            out IntPtr phPC
        );

        // Present in Windows headers but undocumented
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "ExpungeConsoleCommandHistoryW", ExactSpelling = true, SetLastError = true)]
        public static extern void ExpungeConsoleCommandHistory(
            [MarshalAs(UnmanagedType.LPWStr)]
            string ExeName
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool FillConsoleOutputAttribute(
            IntPtr hConsoleOutput,
            CharacterAttributes wAttribute,
            uint nLength,
            COORD dwWriteCoord,
            out uint lpNumberOfAttrsWritten
        );

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "FillConsoleOutputCharacterW", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool FillConsoleOutputCharacter(
            IntPtr hConsoleOutput,
            char cCharacter,
            uint nLength,
            COORD dwWriteCoord,
            out uint lpNumberOfCharsWritten
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool FlushConsoleInputBuffer(
            IntPtr hConsoleInput
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool FreeConsole();

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GenerateConsoleCtrlEvent(
            ControlEvent dwCtrlEvent,
            uint dwProcessGroupId
        );

        // Introduced in Windows XP
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetConsoleAliasW", ExactSpelling = true, SetLastError = true)]
        public static extern uint GetConsoleAlias(
            [MarshalAs(UnmanagedType.LPWStr)]
            string lpSource,

            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)]
            out char[] lpTargetBuffer,

            uint TargetBufferLength,

            [MarshalAs(UnmanagedType.LPWStr)]
            string lpExeName
        );

        // Introduced in Windows XP
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetConsoleAliasesW", ExactSpelling = true, SetLastError = true)]
        public static extern uint GetConsoleAliases(
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)]
            out char[] lpAliasBuffer,

            uint AliasBufferLength,

            [MarshalAs(UnmanagedType.LPWStr)]
            string lpExeName
        );

        // Introduced in Windows XP
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetConsoleAliasesLengthW", ExactSpelling = true)]
        public static extern uint GetConsoleAliasesLength(
            [MarshalAs(UnmanagedType.LPWStr)]
            string lpExeName
        );

        // Introduced in Windows XP
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetConsoleAliasExesW", ExactSpelling = true, SetLastError = true)]
        public static extern uint GetConsoleAliasExes(
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)]
            out char[] lpExeNameBuffer,

            uint ExeNameBufferLength
        );

        // Introduced in Windows XP
        [DllImport("kernel32.dll", EntryPoint = "GetConsoleAliasExesLengthW")]
        public static extern uint GetConsoleAliasExesLength();

        // Present in Windows headers but undocumented
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetConsoleCommandHistoryW", ExactSpelling = true, SetLastError = true)]
        public static extern uint GetConsoleCommandHistory(
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)]
            out char[] Commands,

            uint CommandBufferLength,

            [MarshalAs(UnmanagedType.LPWStr)]
            string ExeName
        );

        // Present in Windows headers but undocumented
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetConsoleCommandHistoryLengthW", ExactSpelling = true, SetLastError = true)]
        public static extern uint GetConsoleCommandHistoryLength(
            [MarshalAs(UnmanagedType.LPWStr)]
            string ExeName
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern uint GetConsoleCP();

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetConsoleCursorInfo(
            IntPtr hConsoleOutput,
            out CONSOLE_CURSOR_INFO lpConsoleCursorInfo
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetConsoleDisplayMode(
            out DisplayModeGetFlags lpModeFlags
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern COORD GetConsoleFontSize(
            IntPtr hConsoleOutput,
            uint nFont
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetConsoleHistoryInfo(
            out CONSOLE_HISTORY_INFO lpConsoleHistoryInfo
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetConsoleMode(
            IntPtr hConsoleHandle,
            out InputModeFlags lpMode
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetConsoleMode(
            IntPtr hConsoleHandle,
            out ScreenBufferModeFlags lpMode
        );

        // Introduced in Windows Vista / Windows Server 2008
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetConsoleOriginalTitleW", ExactSpelling = true, SetLastError = true)]
        public static extern uint GetConsoleOriginalTitle(
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)]
            out char[] lpConsoleTitle,

            uint nSize
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern uint GetConsoleOutputCP();

        // Introduced in Windows XP
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern uint GetConsoleProcessList(
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)]
            out uint[] lpdwProcessList,

            uint dwProcessCount
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetConsoleScreenBufferInfo(
            IntPtr hConsoleOutput,
            out CONSOLE_SCREEN_BUFFER_INFO lpConsoleScreenBufferInfo
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetConsoleScreenBufferInfoEx(
            IntPtr hConsoleOutput,
            out CONSOLE_SCREEN_BUFFER_INFOEX lpConsoleScreenBufferInfoEx
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetConsoleSelectionInfo(
            out CONSOLE_SELECTION_INFO lpConsoleSelectionInfo
        );

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetConsoleTitleW", ExactSpelling = true, SetLastError = true)]
        public static extern uint GetConsoleTitle(
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)]
            out char[] lpConsoleTitle,

            uint nSize
        );

        [DllImport("kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetCurrentConsoleFont(
            IntPtr hConsoleOutput,

            [MarshalAs(UnmanagedType.Bool)]
            bool bMaximumWindow,

            out CONSOLE_FONT_INFO lpConsoleCurrentFont
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetCurrentConsoleFontEx(
            IntPtr hConsoleOutput,

            [MarshalAs(UnmanagedType.Bool)]
            bool bMaximumWindow,

            out CONSOLE_FONT_INFOEX lpConsoleCurrentFontEx
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern COORD GetLargestConsoleWindowSize(
            IntPtr hConsoleOutput
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetNumberOfConsoleInputEvents(
            IntPtr hConsoleInput,
            out uint lpcNumberOfEvents
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetNumberOfConsoleMouseButtons(
            out uint lpNumberOfMouseButtons
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr GetStdHandle(
            StandardDevice nStdHandle
        );

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "PeekConsoleInputW", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool PeekConsoleInput(
            IntPtr hConsoleInput,

            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)]
            out INPUT_RECORD[] lpBuffer,

            uint nLength,
            out uint lpNumberOfEventsRead
        );

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "ReadConsoleW", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ReadConsole(
            IntPtr hConsoleInput,

            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)]
            out char[] lpBuffer,

            uint nNumberOfCharsToRead,
            out uint lpNumberOfCharsRead,
            IntPtr pInputControl // NULL
        );

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "ReadConsoleW", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ReadConsole(
            IntPtr hConsoleInput,

            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)]
            out char[] lpBuffer,

            uint nNumberOfCharsToRead,
            out uint lpNumberOfCharsRead,
            CONSOLE_READCONSOLE_CONTROL pInputControl
        );

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "ReadConsoleInputW", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ReadConsoleInput(
            IntPtr hConsoleInput,

            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)]
            out INPUT_RECORD[] lpBuffer,

            uint nLength,
            out uint lpNumberOfEventsRead
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ReadConsoleOutputAttribute(
            IntPtr hConsoleOutput,

            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)]
            out CharacterAttributes[] lpAttribute,

            uint nLength,
            COORD dwReadCoord,
            out uint lpNumberOfAttrsRead
        );

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "ReadConsoleOutputCharacterW", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ReadConsoleOutputCharacter(
            IntPtr hConsoleOutput,

            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)]
            out char[] lpCharacter,

            uint nLength,
            COORD dwReadCoord,
            out uint lpNumberOfCharsRead
        );

        // Introduced in Windows 10 version 1809 / Windows Server 2019
        [DllImport("kernel32.dll")]
        public static extern int ResizePseudoConsole(
            IntPtr hPC,
            COORD size
        );

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "ScrollConsoleScreenBufferW", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ScrollConsoleScreenBuffer(
            IntPtr hConsoleOutput,
            ref SMALL_RECT lpScrollRectangle,
            IntPtr lpClipRectangle, // NULL
            COORD dwDestinationOrigin,
            ref CHAR_INFO lpFill
        );

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "ScrollConsoleScreenBufferW", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ScrollConsoleScreenBuffer(
            IntPtr hConsoleOutput,
            ref SMALL_RECT lpScrollRectangle,
            ref SMALL_RECT lpClipRectangle,
            COORD dwDestinationOrigin,
            ref CHAR_INFO lpFill
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetConsoleActiveScreenBuffer(
            IntPtr hConsoleOutput
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetConsoleCP(
            uint wCodePageID
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetConsoleCtrlHandler(
            IntPtr HandlerRoutine, // NULL

            [MarshalAs(UnmanagedType.Bool)]
            bool Add
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetConsoleCtrlHandler(
            HandlerRoutine HandlerRoutine,

            [MarshalAs(UnmanagedType.Bool)]
            bool Add
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetConsoleCursorInfo(
            IntPtr hConsoleOutput,
            ref CONSOLE_CURSOR_INFO lpConsoleCursorInfo
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetConsoleCursorPosition(
            IntPtr hConsoleOutput,
            COORD dwCursorPosition
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetConsoleDisplayMode(
            IntPtr hConsoleOutput,
            DisplayModeSetFlags dwFlags,
            IntPtr lpNewScreenBufferDimensions // NULL
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetConsoleDisplayMode(
            IntPtr hConsoleOutput,
            DisplayModeSetFlags dwFlags,
            out COORD lpNewScreenBufferDimensions
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetConsoleHistoryInfo(
            CONSOLE_HISTORY_INFO lpConsoleHistoryInfo
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetConsoleMode(
            IntPtr hConsoleHandle,
            InputModeFlags dwMode
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetConsoleMode(
            IntPtr hConsoleHandle,
            ScreenBufferModeFlags dwMode
        );

        // Present in Windows headers but undocumented
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "SetConsoleNumberOfCommandsW", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetConsoleNumberOfCommands(
            uint Number,

            [MarshalAs(UnmanagedType.LPWStr)]
            string ExeName
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetConsoleOutputCP(
            uint wCodePageID
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetConsoleScreenBufferInfoEx(
            IntPtr hConsoleOutput,
            CONSOLE_SCREEN_BUFFER_INFOEX lpConsoleScreenBufferInfoEx
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetConsoleScreenBufferSize(
            IntPtr hConsoleOutput,
            COORD dwSize
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetConsoleTextAttribute(
            IntPtr hConsoleOutput,
            CharacterAttributes wAttributes
        );

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "SetConsoleTitleW", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetConsoleTitle(
            [MarshalAs(UnmanagedType.LPWStr)]
            string lpConsoleTitle
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetConsoleWindowInfo(
            IntPtr hConsoleOutput,

            [MarshalAs(UnmanagedType.Bool)]
            bool bAbsolute,

            ref SMALL_RECT lpConsoleWindow
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetCurrentConsoleFontEx(
            IntPtr hConsoleOutput,

            [MarshalAs(UnmanagedType.Bool)]
            bool bMaximumWindow,

            CONSOLE_FONT_INFOEX lpConsoleCurrentFontEx
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetStdHandle(
            StandardDevice nStdHandle,
            IntPtr hHandle
        );

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "WriteConsoleW", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool WriteConsole(
            IntPtr hConsoleOutput,

            [MarshalAs(UnmanagedType.LPWStr)]
            string lpBuffer,

            uint nNumberOfCharsToWrite,
            IntPtr lpNumberOfCharsWritten, // NULL
            IntPtr lpReserved // NULL
        );

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "WriteConsoleW", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool WriteConsole(
            IntPtr hConsoleOutput,

            [MarshalAs(UnmanagedType.LPWStr)]
            string lpBuffer,

            uint nNumberOfCharsToWrite,
            out uint lpNumberOfCharsWritten,
            IntPtr lpReserved // NULL
        );

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "WriteConsoleInputW", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool WriteConsoleInput(
            IntPtr hConsoleInput,

            [MarshalAs(UnmanagedType.LPArray)]
            ref INPUT_RECORD[] lpBuffer,

            uint nLength,
            out uint lpNumberOfEventsWritten
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool WriteConsoleOutputAttribute(
            IntPtr hConsoleOutput,
            ref CharacterAttributes lpAttribute,
            uint nLength,
            COORD dwWriteCoord,
            out uint lpNumberOfAttrsWritten
        );

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "WriteConsoleOutputCharacterW", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool WriteConsoleOutputCharacter(
            IntPtr hConsoleOutput,

            [MarshalAs(UnmanagedType.LPWStr)]
            string lpCharacter,

            uint nLength,
            COORD dwWriteCoord,
            out uint lpNumberOfCharsWritten
        );

        #endregion

        #region Structures

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct CHAR_INFO {
            public char Char;
            public CharacterAttributes Attributes;
        }

        public struct COLORREF {
            public byte R;
            public byte G;
            public byte B;
        }

        public struct CONSOLE_CURSOR_INFO {
            public uint dwSize;
            public bool bVisible;
        }

        public struct CONSOLE_FONT_INFO {
            public uint nFont;
            public COORD dwFontSize;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public class CONSOLE_FONT_INFOEX {
            public uint cbSize;
            public uint nFont;
            public COORD dwFontSize;
            public uint FontFamily;
            public uint FontWeight;

            [MarshalAs(UnmanagedType.ByValArray, ArraySubType = UnmanagedType.LPWStr, SizeConst = LF_FACESIZE)]
            public char[] FaceName;

            public CONSOLE_FONT_INFOEX() {
                cbSize = (uint)Marshal.SizeOf(typeof(CONSOLE_FONT_INFOEX));
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        public class CONSOLE_HISTORY_INFO {
            public uint cbSize;
            public uint HistoryBufferSize;
            public uint NumberOfHistoryBuffers;
            public HistoryInfoFlags dwFlags;

            public CONSOLE_HISTORY_INFO() {
                cbSize = (uint)Marshal.SizeOf(typeof(CONSOLE_HISTORY_INFO));
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        public class CONSOLE_READCONSOLE_CONTROL {
            public uint nLength;
            public uint nInitialChars;
            public uint dwCtrlWakeupMask;
            public ControlKeyStates dwControlKeyState;

            public CONSOLE_READCONSOLE_CONTROL() {
                nLength = (uint)Marshal.SizeOf(typeof(CONSOLE_READCONSOLE_CONTROL));
            }
        }

        public struct CONSOLE_SCREEN_BUFFER_INFO {
            public COORD dwSize;
            public COORD dwCursorPosition;
            public CharacterAttributes wAttributes;
            public SMALL_RECT srWindow;
            public COORD dwMaximumWindowSize;
        }

        [StructLayout(LayoutKind.Sequential)]
        public class CONSOLE_SCREEN_BUFFER_INFOEX {
            public uint cbSize;
            public COORD dwSize;
            public COORD dwCursorPosition;
            public CharacterAttributes wAttributes;
            public SMALL_RECT srWindow;
            public COORD dwMaximumWindowSize;
            public ushort wPopupAttributes; // TODO
            public bool bFullscreenSupported;

            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 16)]
            public COLORREF[] ColorTable;

            public CONSOLE_SCREEN_BUFFER_INFOEX() {
                cbSize = (uint)Marshal.SizeOf(typeof(CONSOLE_SCREEN_BUFFER_INFOEX));
            }
        }

        public struct CONSOLE_SELECTION_INFO {
            public SelectionInfoFlags dwFlags;
            public COORD dwSelectionAnchor;
            public SMALL_RECT srSelection;
        }

        public struct COORD {
            public short X;
            public short Y;
        }

        public struct FOCUS_EVENT_RECORD {
            public bool bSetFocus;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct INPUT_RECORD {
            [FieldOffset(0)]
            public EventType EventType;

            [FieldOffset(4)]
            public KEY_EVENT_RECORD KeyEvent;

            [FieldOffset(4)]
            public MOUSE_EVENT_RECORD MouseEvent;

            [FieldOffset(4)]
            public WINDOW_BUFFER_SIZE_RECORD WindowBufferSizeEvent;

            [FieldOffset(4)]
            public MENU_EVENT_RECORD MenuEvent;

            [FieldOffset(4)]
            public FOCUS_EVENT_RECORD FocusEvent;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct KEY_EVENT_RECORD {
            public bool bKeyDown;
            public ushort wRepeatCount;
            public ushort wVirtualKeyCode;
            public ushort wVirtualScanCode;
            public char uChar;
            public ControlKeyStates dwControlKeyState;
        }

        public struct MENU_EVENT_RECORD {
            public uint dwCommandId;
        }

        public struct MOUSE_EVENT_RECORD {
            public COORD dwMousePosition;
            public MouseButtonStates dwButtonState;
            public ControlKeyStates dwControlKeyState;
            public MouseEventFlags dwEventFlags;
        }

        public struct SMALL_RECT {
            public short Left;
            public short Top;
            public short Right;
            public short Bottom;
        }

        public struct WINDOW_BUFFER_SIZE_RECORD {
            public COORD dwSize;
        }

        #endregion
    }
}
