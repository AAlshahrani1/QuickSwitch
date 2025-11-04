using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Linq;

namespace QuickSwitch;

public partial class MainForm : Form
{

    private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN = 0x0100;
    private const int WM_KEYUP = 0x0101;
    private const int WM_SYSKEYDOWN = 0x0104;
    private const int WM_SYSKEYUP = 0x0105;

    private const int VK_CAPITAL = 0x14;
    private const int VK_LSHIFT = 0xA0;
    private const int VK_RSHIFT = 0xA1;
    private const int VK_LMENU = 0xA4;

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, 
        IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string lpModuleName);

    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);

    private const int KEYEVENTF_KEYUP = 0x0002;

    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    private readonly LowLevelKeyboardProc _proc;
    private static IntPtr _hookID = IntPtr.Zero;
    private static bool _shiftPressed = false;
    private NotifyIcon? _trayIcon;
    private CheckBox _startupCheckbox = null!;

    public MainForm()
    {
        _proc = HookCallback;
        InitializeComponent();
        _hookID = SetHook(_proc);

        this.Icon = LoadApplicationIcon();

        CheckStartupMode();
    }

    private void CheckStartupMode()
    {

        string[] args = Environment.GetCommandLineArgs();

        bool isStartup = args.Length > 1 && args[1] == "--startup";
        
        if (isStartup)
        {

            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            this.Opacity = 0;

            this.Load += (s, e) =>
            {
                this.Hide();
                this.Opacity = 100;
            };
        }
    }

    private static Icon LoadApplicationIcon()
    {
        try
        {

            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            var resourceName = assembly.GetManifestResourceNames()
                .FirstOrDefault(name => name.EndsWith("caps.ico"));
            
            if (resourceName != null)
            {
                using var stream = assembly.GetManifestResourceStream(resourceName);
                if (stream != null)
                {
                    return new Icon(stream);
                }
            }

            if (File.Exists("caps.ico"))
            {
                return new Icon("caps.ico");
            }

            var paths = new[]
            {
                Path.Combine(Application.StartupPath, "caps.ico"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "caps.ico"),
                Path.Combine(Directory.GetCurrentDirectory(), "caps.ico")
            };
            
            foreach (var path in paths)
            {
                if (File.Exists(path))
                {
                    return new Icon(path);
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Could not load custom icon: {ex.Message}");
        }

        return SystemIcons.Application;
    }

    private void InitializeComponent()
    {
        this.Text = "Quick Switch";
        this.Size = new System.Drawing.Size(440, 320);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedSingle;
        this.MaximizeBox = false;
        this.MinimizeBox = true;
        this.BackColor = System.Drawing.Color.White;

        var titleLabel = new Label
        {
            Text = "Quick Switch",
            Location = new System.Drawing.Point(20, 20),
            Size = new System.Drawing.Size(380, 30),
            Font = new System.Drawing.Font("Segoe UI", 14, System.Drawing.FontStyle.Bold),
            ForeColor = System.Drawing.Color.FromArgb(50, 50, 50)
        };
        this.Controls.Add(titleLabel);

        var separator1 = new Panel
        {
            Location = new System.Drawing.Point(20, 55),
            Size = new System.Drawing.Size(380, 1),
            BackColor = System.Drawing.Color.FromArgb(220, 220, 220)
        };
        this.Controls.Add(separator1);

        var infoLabel = new Label
        {
            Text = "Key Mappings:",
            Location = new System.Drawing.Point(20, 70),
            Size = new System.Drawing.Size(380, 25),
            Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold),
            ForeColor = System.Drawing.Color.FromArgb(70, 70, 70)
        };
        this.Controls.Add(infoLabel);

        var mapping1 = new Label
        {
            Text = "CapsLock  →  Shift+Alt",
            Location = new System.Drawing.Point(40, 100),
            Size = new System.Drawing.Size(200, 25),
            Font = new System.Drawing.Font("Segoe UI", 10),
            ForeColor = System.Drawing.Color.FromArgb(80, 80, 80)
        };
        this.Controls.Add(mapping1);

        var mapping1Desc = new Label
        {
            Text = "(Switch keyboard language)",
            Location = new System.Drawing.Point(250, 100),
            Size = new System.Drawing.Size(180, 25),
            Font = new System.Drawing.Font("Segoe UI", 9),
            ForeColor = System.Drawing.Color.FromArgb(120, 120, 120)
        };
        this.Controls.Add(mapping1Desc);

        var mapping2 = new Label
        {
            Text = "Shift+CapsLock  →  CapsLock",
            Location = new System.Drawing.Point(40, 130),
            Size = new System.Drawing.Size(220, 25),
            Font = new System.Drawing.Font("Segoe UI", 10),
            ForeColor = System.Drawing.Color.FromArgb(80, 80, 80)
        };
        this.Controls.Add(mapping2);

        var mapping2Desc = new Label
        {
            Text = "(Toggle CapsLock on/off)",
            Location = new System.Drawing.Point(270, 130),
            Size = new System.Drawing.Size(160, 25),
            Font = new System.Drawing.Font("Segoe UI", 9),
            ForeColor = System.Drawing.Color.FromArgb(120, 120, 120)
        };
        this.Controls.Add(mapping2Desc);

        var separator2 = new Panel
        {
            Location = new System.Drawing.Point(20, 170),
            Size = new System.Drawing.Size(380, 1),
            BackColor = System.Drawing.Color.FromArgb(220, 220, 220)
        };
        this.Controls.Add(separator2);

        _startupCheckbox = new CheckBox
        {
            Text = "Run on startup",
            Location = new System.Drawing.Point(20, 186),
            Size = new System.Drawing.Size(200, 25),
            Checked = IsInStartup(),
            Font = new System.Drawing.Font("Segoe UI", 10),
            ForeColor = System.Drawing.Color.FromArgb(70, 70, 70)
        };
        _startupCheckbox.CheckedChanged += StartupCheckbox_CheckedChanged;
        this.Controls.Add(_startupCheckbox);

        var buttonPanel = new Panel
        {
            Location = new System.Drawing.Point(20, 230),
            Size = new System.Drawing.Size(380, 40)
        };

        var minimizeButton = new Button
        {
            Text = "Minimize to Tray",
            Location = new System.Drawing.Point(0, 0),
            Size = new System.Drawing.Size(130, 35),
            UseVisualStyleBackColor = true,
            FlatStyle = FlatStyle.System,
            Font = new System.Drawing.Font("Segoe UI", 9)
        };
        minimizeButton.Click += (s, e) => HideToTray();

        var exitButton = new Button
        {
            Text = "Exit",
            Location = new System.Drawing.Point(280, 0),
            Size = new System.Drawing.Size(100, 35),
            UseVisualStyleBackColor = true,
            FlatStyle = FlatStyle.System,
            Font = new System.Drawing.Font("Segoe UI", 9)
        };
        exitButton.Click += (s, e) => ExitApplication();

        buttonPanel.Controls.Add(minimizeButton);
        buttonPanel.Controls.Add(exitButton);
        this.Controls.Add(buttonPanel);

        SetupSystemTray();

        this.FormClosing += (s, e) => 
        {
            if (e?.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                HideToTray();
            }
        };

        string[] args = Environment.GetCommandLineArgs();
        bool isStartup = args.Length > 1 && args[1] == "--startup";
        
        if (!isStartup)
        {
            _trayIcon?.ShowBalloonTip(3000, 
                "Quick Switch",
                "Application is running. CapsLock → Shift+Alt active.",
                ToolTipIcon.Info);
        }
    }

    private void SetupSystemTray()
    {
        _trayIcon = new NotifyIcon
        {
            Icon = LoadApplicationIcon(),
            Text = "Quick Switch - Active",
            Visible = true
        };

        _trayIcon.DoubleClick += (s, e) => ShowFromTray();

        var trayMenu = new ContextMenuStrip();
        trayMenu.Items.Add("Show Window", null, (s, e) => ShowFromTray());
        trayMenu.Items.Add("-");
        trayMenu.Items.Add("Exit", null, (s, e) => ExitApplication());
        
        _trayIcon.ContextMenuStrip = trayMenu;
    }

    private void HideToTray()
    {
        this.Hide();
        this.ShowInTaskbar = false;
    }

    private void ShowFromTray()
    {
        this.Show();
        this.ShowInTaskbar = true;
        this.WindowState = FormWindowState.Normal;
        this.BringToFront();
    }

    private void ExitApplication()
    {
        _trayIcon?.Dispose();
        Application.Exit();
    }

    private static bool IsInStartup()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", false);
            var value = key?.GetValue("QuickSwitch");
            return value != null;
        }
        catch
        {
            return false;
        }
    }

    private static void AddToStartup()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);
            var exePath = Application.ExecutablePath;

            key?.SetValue("QuickSwitch", $"\"{exePath}\" --startup");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Could not add to startup: {ex.Message}", 
                "Startup Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private static void RemoveFromStartup()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);
            key?.DeleteValue("QuickSwitch", false);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Could not remove from startup: {ex.Message}", 
                "Startup Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void StartupCheckbox_CheckedChanged(object? sender, EventArgs e)
    {
        if (_startupCheckbox == null) return;

        if (_startupCheckbox.Checked)
        {
            var result = MessageBox.Show(
                "After enabling startup, do not move or rename this exe file.\n\n" +
                "Make sure the program is in a permanent location (like C:\\ or a dedicated folder) before enabling startup.\n\n" +
                "Continue?",
                "Startup Location Warning",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);

            if (result == DialogResult.Yes)
            {
                AddToStartup();
            }
            else
            {
                _startupCheckbox.CheckedChanged -= StartupCheckbox_CheckedChanged;
                _startupCheckbox.Checked = false;
                _startupCheckbox.CheckedChanged += StartupCheckbox_CheckedChanged;
            }
        }
        else
        {
            RemoveFromStartup();
        }
    }

    private static IntPtr SetHook(LowLevelKeyboardProc proc)
    {
        using var curProcess = Process.GetCurrentProcess();
        using var curModule = curProcess.MainModule;
        return SetWindowsHookEx(WH_KEYBOARD_LL, proc,
            GetModuleHandle(curModule?.ModuleName ?? string.Empty), 0);
    }

    private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0)
        {
            int vkCode = Marshal.ReadInt32(lParam);
            bool keyDown = (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN);
            bool keyUp = (wParam == (IntPtr)WM_KEYUP || wParam == (IntPtr)WM_SYSKEYUP);

            if (vkCode == VK_LSHIFT || vkCode == VK_RSHIFT)
            {
                if (keyDown)
                    _shiftPressed = true;
                else if (keyUp)
                    _shiftPressed = false;
            }

            if (vkCode == VK_CAPITAL)
            {
                if (keyDown)
                {
                    if (_shiftPressed)
                    {

                        return CallNextHookEx(_hookID, nCode, wParam, lParam);
                    }
                    else
                    {

                        Task.Run(() => SendShiftAlt());
                        return (IntPtr)1;
                    }
                }
                else if (keyUp && !_shiftPressed)
                {

                    return (IntPtr)1;
                }
            }
        }

        return CallNextHookEx(_hookID, nCode, wParam, lParam);
    }

    private static void SendShiftAlt()
    {

        Thread.Sleep(10);

        keybd_event(VK_LSHIFT, 0, 0, 0);
        keybd_event(VK_LMENU, 0, 0, 0);
        Thread.Sleep(50);
        keybd_event(VK_LMENU, 0, KEYEVENTF_KEYUP, 0);
        keybd_event(VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0);
    }

    protected override void Dispose(bool disposing)
    {
        if (_hookID != IntPtr.Zero)
        {
            UnhookWindowsHookEx(_hookID);
            _hookID = IntPtr.Zero;
        }
        _trayIcon?.Dispose();
        base.Dispose(disposing);
    }
}

public static class Program
{
    [STAThread]
    public static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        var currentProcess = Process.GetCurrentProcess();
        var processes = Process.GetProcessesByName(currentProcess.ProcessName);
        
        if (processes.Length > 1)
        {
            MessageBox.Show("Quick Switch is already running!", "Already Running", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        try
        {
            Application.Run(new MainForm());
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error starting application: {ex.Message}", "Error", 
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}