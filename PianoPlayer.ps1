# Put VP sheet here

$sheet = ''


$speed = 1
$sheet = $sheet.Replace("`n","").Replace("`r", "").Replace(" - ", " | ").Replace("-","") -replace "[{}]+"

<#---------------------------------------------------------------------------------------#>
$keySchedule = @()
$timeSchedule = @()
$spTime = 250 * $speed
$vertTime = 500 * $speed
$nonTime = 175 * $speed
$fastTime = 100 * $speed

for ($i = 0; $i -lt $sheet.Length; $i++){
    if ( -not ($sheet[$i] -eq " " -or $sheet[$i] -eq "[" -or $sheet[$i] -eq "]")) {
        if (($sheet.Length -gt $i+3) -and (($sheet.Substring($i+1,3) -eq " | ") -or ($sheet.Substring($i+1,3) -eq " - "))) {
            $keySchedule += $sheet[$i]
            $timeSchedule += $vertTime
            $i += 3
        } elseif ($sheet[$i+1] -eq " "){
            $keySchedule += $sheet[$i]
            $timeSchedule += $spTime
            $i++
        } elseif ($sheet[$i+1] -ne " " ) {
            $keySchedule += $sheet[$i]
            $timeSchedule += $nonTime
        }
    } elseif ($sheet.Substring($i) -match "^\[.+?]") {
        $s = $Matches[0].Trim("[]")
        if ($s -match "\s") {
            for ($o; $o -lt $s.Length; $o++) {
                if ($s[$o] -eq " ") { return }
                $keySchedule += $s[$o]
                $timeSchedule += $fastTime
            }
        } else {
            $keySchedule += $s
        }
        $i += ($s.Length + 1)
        if (($sheet.Length -gt $i+3) -and $sheet.Substring($i+1,3) -eq " | ") {
            $timeSchedule += $vertTime
            $i += 3
        } elseif ($sheet[$i+1] -eq " "){
            $timeSchedule += $spTime
            $i++
        } elseif ($sheet[$i+1] -ne " " ) {
            $timeSchedule += $nonTime
        }
    }
}


$keydownFlag = 0x0000
$keyupFlag = 0x0002;

$source = @"
using System;
using System.Runtime.InteropServices;

public class Keyer{

    const int INPUT_MOUSE = 0;
    const int INPUT_KEYBOARD = 1;
    const int INPUT_HARDWARE = 2;
    const int SCANCODE = 0x0008;

    [StructLayout(LayoutKind.Sequential)]
    struct MOUSEINPUT
    {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct KEYBDINPUT 
    {
         public ushort wVk;
         public ushort wScan;
         public uint dwFlags;
         public uint time;
         public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct HARDWAREINPUT
    {
         public int uMsg;
         public short wParamL;
         public short wParamH;
    }

    [StructLayout(LayoutKind.Explicit)]
    struct MouseKeybdHardwareInputUnion
    {
        [FieldOffset(0)]
        public MOUSEINPUT mi;

        [FieldOffset(0)]
        public KEYBDINPUT ki;

        [FieldOffset(0)]
        public HARDWAREINPUT hi;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct INPUT
    {
        public uint type;
        public MouseKeybdHardwareInputUnion u;
    }
    
    public enum DirectXKeyStrokes
    {
        DIK_ESCAPE = 0x01,
        DIK_1 = 0x02,
        DIK_2 = 0x03,
        DIK_3 = 0x04,
        DIK_4 = 0x05,
        DIK_5 = 0x06,
        DIK_6 = 0x07,
        DIK_7 = 0x08,
        DIK_8 = 0x09,
        DIK_9 = 0x0A,
        DIK_0 = 0x0B,
        DIK_MINUS = 0x0C,
        DIK_EQUALS = 0x0D,
        DIK_BACK = 0x0E,
        DIK_TAB = 0x0F,
        DIK_Q = 0x10,
        DIK_W = 0x11,
        DIK_E = 0x12,
        DIK_R = 0x13,
        DIK_T = 0x14,
        DIK_Y = 0x15,
        DIK_U = 0x16,
        DIK_I = 0x17,
        DIK_O = 0x18,
        DIK_P = 0x19,
        DIK_LBRACKET = 0x1A,
        DIK_RBRACKET = 0x1B,
        DIK_RETURN = 0x1C,
        DIK_LCONTROL = 0x1D,
        DIK_A = 0x1E,
        DIK_S = 0x1F,
        DIK_D = 0x20,
        DIK_F = 0x21,
        DIK_G = 0x22,
        DIK_H = 0x23,
        DIK_J = 0x24,
        DIK_K = 0x25,
        DIK_L = 0x26,
        DIK_SEMICOLON = 0x27,
        DIK_APOSTROPHE = 0x28,
        DIK_GRAVE = 0x29,
        DIK_LSHIFT = 0x2A,
        DIK_BACKSLASH = 0x2B,
        DIK_Z = 0x2C,
        DIK_X = 0x2D,
        DIK_C = 0x2E,
        DIK_V = 0x2F,
        DIK_B = 0x30,
        DIK_N = 0x31,
        DIK_M = 0x32,
        DIK_COMMA = 0x33,
        DIK_PERIOD = 0x34,
        DIK_SLASH = 0x35,
        DIK_RSHIFT = 0x36,
        DIK_MULTIPLY = 0x37,
        DIK_LMENU = 0x38,
        DIK_SPACE = 0x39,
        DIK_CAPITAL = 0x3A,
        DIK_F1 = 0x3B,
        DIK_F2 = 0x3C,
        DIK_F3 = 0x3D,
        DIK_F4 = 0x3E,
        DIK_F5 = 0x3F,
        DIK_F6 = 0x40,
        DIK_F7 = 0x41,
        DIK_F8 = 0x42,
        DIK_F9 = 0x43,
        DIK_F10 = 0x44,
        DIK_NUMLOCK = 0x45,
        DIK_SCROLL = 0x46,
        DIK_NUMPAD7 = 0x47,
        DIK_NUMPAD8 = 0x48,
        DIK_NUMPAD9 = 0x49,
        DIK_SUBTRACT = 0x4A,
        DIK_NUMPAD4 = 0x4B,
        DIK_NUMPAD5 = 0x4C,
        DIK_NUMPAD6 = 0x4D,
        DIK_ADD = 0x4E,
        DIK_NUMPAD1 = 0x4F,
        DIK_NUMPAD2 = 0x50,
        DIK_NUMPAD3 = 0x51,
        DIK_NUMPAD0 = 0x52,
        DIK_DECIMAL = 0x53,
        DIK_F11 = 0x57,
        DIK_F12 = 0x58,
        DIK_F13 = 0x64,
        DIK_F14 = 0x65,
        DIK_F15 = 0x66,
        DIK_KANA = 0x70,
        DIK_CONVERT = 0x79,
        DIK_NOCONVERT = 0x7B,
        DIK_YEN = 0x7D,
        DIK_NUMPADEQUALS = 0x8D,
        DIK_CIRCUMFLEX = 0x90,
        DIK_AT = 0x91,
        DIK_COLON = 0x92,
        DIK_UNDERLINE = 0x93,
        DIK_KANJI = 0x94,
        DIK_STOP = 0x95,
        DIK_AX = 0x96,
        DIK_UNLABELED = 0x97,
        DIK_NUMPADENTER = 0x9C,
        DIK_RCONTROL = 0x9D,
        DIK_NUMPADCOMMA = 0xB3,
        DIK_DIVIDE = 0xB5,
        DIK_SYSRQ = 0xB7,
        DIK_RMENU = 0xB8,
        DIK_HOME = 0xC7,
        DIK_UP = 0xC8,
        DIK_PRIOR = 0xC9,
        DIK_LEFT = 0xCB,
        DIK_RIGHT = 0xCD,
        DIK_END = 0xCF,
        DIK_DOWN = 0xD0,
        DIK_NEXT = 0xD1,
        DIK_INSERT = 0xD2,
        DIK_DELETE = 0xD3,
        DIK_LWIN = 0xDB,
        DIK_RWIN = 0xDC,
        DIK_APPS = 0xDD,
        DIK_BACKSPACE = DIK_BACK,
        DIK_NUMPADSTAR = DIK_MULTIPLY,
        DIK_LALT = DIK_LMENU,
        DIK_CAPSLOCK = DIK_CAPITAL,
        DIK_NUMPADMINUS = DIK_SUBTRACT,
        DIK_NUMPADPLUS = DIK_ADD,
        DIK_NUMPADPERIOD = DIK_DECIMAL,
        DIK_NUMPADSLASH = DIK_DIVIDE,
        DIK_RALT = DIK_RMENU,
        DIK_UPARROW = DIK_UP,
        DIK_PGUP = DIK_PRIOR,
        DIK_LEFTARROW = DIK_LEFT,
        DIK_RIGHTARROW = DIK_RIGHT,
        DIK_DOWNARROW = DIK_DOWN,
        DIK_PGDN = DIK_NEXT,

        // Mined these out of nowhere.
        DIK_LEFTMOUSEBUTTON = 0x100,
        DIK_RIGHTMOUSEBUTTON  = 0x101,
        DIK_MIDDLEWHEELBUTTON = 0x102,
        DIK_MOUSEBUTTON3 = 0x103,
        DIK_MOUSEBUTTON4 = 0x104,
        DIK_MOUSEBUTTON5 = 0x105,
        DIK_MOUSEBUTTON6 = 0x106,
        DIK_MOUSEBUTTON7 = 0x107,
        DIK_MOUSEWHEELUP = 0x108,
        DIK_MOUSEWHEELDOWN = 0x109,
    }

    [DllImport("user32.dll", SetLastError = true)]
    static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    public static void KeyString(string keys, uint flags) {
        char[] chars = keys.ToCharArray();
        INPUT[] inputs = new INPUT[chars.Length];

        for (int i = 0; i < chars.Length; i++){
            string ind = "DIK_" + chars[i];
            DirectXKeyStrokes kSv = (DirectXKeyStrokes) Enum.Parse(typeof(DirectXKeyStrokes), ind);
            inputs[i].type = INPUT_KEYBOARD;
            inputs[i].u.ki.wScan = (ushort) kSv;
            inputs[i].u.ki.dwFlags = flags | SCANCODE;
        };

		SendInput((uint) chars.Length, inputs, Marshal.SizeOf(inputs[0].GetType()));
    }

    public static void KeyScan(ushort key, uint flags) {
        INPUT[] inputs = new INPUT[1];

        inputs[0].type = INPUT_KEYBOARD;
        inputs[0].u.ki.wScan = key;
        inputs[0].u.ki.dwFlags = flags | SCANCODE;

		SendInput((uint) 1, inputs, Marshal.SizeOf(inputs[0].GetType()));
    }
}
"@

Add-Type -TypeDefinition $source

$conv = @{"!"="1"; "@"="2";"\$"="4"; "%"="5"; "\^"="6"; "\*"="8"; "\("="9"}

sleep 1;
for ($i = 0; $i -lt $keySchedule.Length; $i++) {
    $keys = $keySchedule[$i].ToString().ToUpper()
    $up = $keySchedule[$i] -creplace "[^A-Z!$%^*(@]+"
    $down = $keySchedule[$i] -creplace "[A-Z!$%^*(@]+"
    foreach ($e in $conv.GetEnumerator()) {
        $up = $up -replace $e.Name, $e.Value
        $keys = $keys -replace $e.Name, $e.Value
    }

    [Keyer]::KeyScan([Keyer+DirectXKeyStrokes]::DIK_LSHIFT, $keydownFlag)
        if ($up) { [Keyer]::KeyString($up.ToUpper(), $keydownFlag) }
    [Keyer]::KeyScan([Keyer+DirectXKeyStrokes]::DIK_LSHIFT, $keyupFlag)
    if ($down) { [Keyer]::KeyString($down.ToUpper(), $keydownFlag) }

    sleep -Milliseconds 50
    [Keyer]::KeyString($keys, $keyupFlag)

    sleep -Milliseconds $timeSchedule[$i] #(Get-Random -Maximum ($timeSchedule[$i] + 10) -Minimum ($timeSchedule[$i] - 10))
}
