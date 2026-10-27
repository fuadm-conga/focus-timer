using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Media;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Threading;
using Shapes = System.Windows.Shapes;

namespace FocusTimer;

static class Program
{
    [STAThread]
    static void Main() => new Application().Run(new TimerWindow());
}

enum Corner { TopLeft, TopRight, BottomLeft, BottomRight }

// Always-on-top focus timer that sits in a screen corner.
// Drag it to any corner of any monitor and it snaps there; right-click for options.
class TimerWindow : Window
{
    static readonly SolidColorBrush Surface = Hex("#151A35"), Ink = Hex("#EEF0FB"), Muted = Hex("#9AA0C2"),
        Track = Hex("#232A4D"), Hover = Hex("#2E3566"), Accent = Hex("#8C8FFF"), Done = Hex("#6FD3A8");
    const double RingSize = 128, Stroke = 9, Gap = 8; // Gap sits outside the 10px shadow margin
    const string PlayGlyph = "", PauseGlyph = "", ResetGlyph = ""; // Segoe Fluent Icons
    static readonly int[] Presets = { 15, 25, 50 };

    static readonly string SettingsFile = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "FocusTimer", "settings.txt");

    readonly Shapes.Path arc = new()
    {
        Stroke = Accent, StrokeThickness = Stroke,
        StrokeStartLineCap = PenLineCap.Round, StrokeEndLineCap = PenLineCap.Round,
    };
    readonly TextBlock time = new()
    {
        FontFamily = new FontFamily("Segoe UI Variable Display, Segoe UI"), FontWeight = FontWeights.Light,
        FontSize = 30, Foreground = Ink, HorizontalAlignment = HorizontalAlignment.Center,
    };
    readonly TextBlock state = new()
    {
        FontSize = 10, FontWeight = FontWeights.SemiBold, Foreground = Muted,
        HorizontalAlignment = HorizontalAlignment.Center,
    };
    readonly TextBlock toggleIcon = Glyph(PlayGlyph, Surface, 14);
    readonly List<(Border Chip, int Minutes)> chips = new();
    readonly TextBox custom = new()
    {
        Width = 44, Height = 30, MaxLength = 3, FontSize = 12, Background = Track, Foreground = Ink, CaretBrush = Ink,
        BorderThickness = new Thickness(0), HorizontalContentAlignment = HorizontalAlignment.Center,
        VerticalContentAlignment = VerticalAlignment.Center, ToolTip = "Custom length in minutes (1–180), then Enter",
    };
    UIElement lengthRow = null!;
    Border closeButton = null!;
    readonly DispatcherTimer ticker =new() { Interval = TimeSpan.FromMilliseconds(33) };

    TimeSpan length = TimeSpan.FromMinutes(25), remaining;
    DateTime endAt;
    bool running, placed;
    Corner corner = Corner.BottomRight;
    POINT? savedCenter;

    IntPtr Hwnd => new WindowInteropHelper(this).Handle;
    bool Idle => !running && remaining == length;

    public TimerWindow()
    {
        Title = "Focus Timer";
        WindowStyle = WindowStyle.None;
        AllowsTransparency = true;
        Background = Brushes.Transparent;
        ResizeMode = ResizeMode.NoResize;
        SizeToContent = SizeToContent.WidthAndHeight;
        Topmost = true;
        ShowInTaskbar = false;
        Opacity = 0.97;
        System.Windows.Documents.Typography.SetNumeralAlignment(time, FontNumeralAlignment.Tabular);

        Load();
        remaining = length;
        Content = BuildCard();
        ContextMenu = BuildMenu();

        ticker.Tick += (_, _) => Tick();
        Loaded += (_, _) =>
        {
            var monitor = savedCenter is POINT p ? MonitorFromPoint(p, MONITOR_DEFAULTTONEAREST)
                                                 : MonitorFromPoint(default, MONITOR_DEFAULTTOPRIMARY);
            UpdateHover();
            SnapTo(monitor, corner);
            placed = true;
            Paint();
            SetState("READY");
        };
        // Hover-revealed rows or a DPI change on another monitor resize the window from its top-left,
        // so re-anchor to the corner.
        SizeChanged += (_, _) =>
        {
            if (placed) SnapTo(CurrentMonitor(), corner); // synchronous, so hover-revealed rows don't flicker past the screen edge
        };
        MouseEnter += (_, _) => UpdateHover();
        MouseLeave += (_, _) => UpdateHover();
        KeyDown += (_, e) =>
        {
            if (e.OriginalSource is TextBox) return;
            if (e.Key == Key.Space) Toggle();
            if (e.Key == Key.R) Reset();
        };
    }

    UIElement BuildCard()
    {
        var face = new StackPanel { VerticalAlignment = VerticalAlignment.Center, Children = { time, state } };
        var ring = new Grid
        {
            Width = RingSize, Height = RingSize, Background = Brushes.Transparent,
            ToolTip = "Scroll to change length (Shift = 5 min)\nDrag to any corner · Right-click for options",
            Children =
            {
                new Shapes.Ellipse { Stroke = Track, StrokeThickness = Stroke, Margin = new Thickness(Stroke / 2) },
                arc,
                face,
            },
        };
        ring.MouseWheel += (_, e) =>
        {
            if (!Idle) return; // a stray scroll can't wreck a running session
            int step = Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) ? 5 : 1;
            SetLength((int)length.TotalMinutes + Math.Sign(e.Delta) * step);
        };

        lengthRow = BuildLengthRow();
        closeButton = IconButton(Glyph(((char)0xE711).ToString(), Muted, 10), "Close", Brushes.Transparent, Track, Close);
        closeButton.Width = closeButton.Height = 24;
        closeButton.HorizontalAlignment = HorizontalAlignment.Right;
        closeButton.VerticalAlignment = VerticalAlignment.Top;
        closeButton.Margin = new Thickness(0, -8, -8, 0); // tuck into the card's top-right padding

        var reset = IconButton(Glyph(ResetGlyph, Ink, 14), "Reset (R)", Track, Hover, Reset);
        reset.Margin = new Thickness(10, 0, 0, 0);
        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center,
            Margin = new Thickness(0, 12, 0, 0),
            Children = { IconButton(toggleIcon, "Start / pause (Space)", Accent, Accent, Toggle), reset },
        };

        var card = new Border
        {
            Width = 192, Background = Surface, CornerRadius = new CornerRadius(20), Padding = new Thickness(16),
            Margin = new Thickness(10), // room for the shadow inside the transparent window
            Effect = new DropShadowEffect { BlurRadius = 14, ShadowDepth = 2, Opacity = 0.35 },
            Child = new Grid
            {
                Children =
                {
                    new StackPanel { Children = { ring, buttons, lengthRow } },
                    closeButton,
                },
            },
        };
        card.MouseLeftButtonDown += (_, _) =>
        {
            DragMove();
            SnapToNearestCorner();
        };
        return card;
    }

    // Close button and length presets stay out of the way until the pointer is over the card
    // (or you're still typing a custom length).
    void UpdateHover()
    {
        bool show = IsMouseOver || custom.IsKeyboardFocused;
        Opacity = show ? 1 : 0.97;
        closeButton.Visibility = show ? Visibility.Visible : Visibility.Hidden;
        lengthRow.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
    }

    UIElement BuildLengthRow()
    {
        var row = new StackPanel
        {
            Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center,
            Margin = new Thickness(0, 12, 0, 0),
        };
        foreach (int m in Presets)
        {
            var chip = new Border
            {
                Width = 32, Height = 30, CornerRadius = new CornerRadius(15), Margin = new Thickness(0, 0, 6, 0),
                Cursor = Cursors.Hand, ToolTip = $"{m} minutes",
                Child = new TextBlock
                {
                    Text = m.ToString(), FontSize = 12, Foreground = Ink,
                    HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center,
                },
            };
            chip.MouseLeftButtonDown += (_, e) => { e.Handled = true; if (Idle) SetLength(m); };
            chips.Add((chip, m));
            row.Children.Add(chip);
        }

        custom.PreviewTextInput += (s, e) => e.Handled = !e.Text.All(char.IsDigit);
        custom.KeyDown += (_, e) =>
        {
            if (e.Key != Key.Enter) return;
            ApplyCustom();
            Keyboard.ClearFocus();
        };
        custom.LostKeyboardFocus += (_, _) => { ApplyCustom(); UpdateHover(); };
        row.Children.Add(new Border { CornerRadius = new CornerRadius(8), ClipToBounds = true, Child = custom });
        return row;
    }

    ContextMenu BuildMenu()
    {
        var menu = new ContextMenu();
        var cornerMenu = new MenuItem { Header = "Corner" };
        foreach (Corner c in Enum.GetValues<Corner>())
        {
            var item = new MenuItem { Header = c switch
            {
                Corner.TopLeft => "Top left", Corner.TopRight => "Top right",
                Corner.BottomLeft => "Bottom left", _ => "Bottom right",
            }, Tag = c };
            item.Click += (_, _) => SnapTo(CurrentMonitor(), c);
            cornerMenu.Items.Add(item);
        }
        var exit = new MenuItem { Header = "Exit" };
        exit.Click += (_, _) => Close();

        menu.Items.Add(cornerMenu);
        menu.Items.Add(new Separator());
        menu.Items.Add(exit);
        menu.Opened += (_, _) =>
        {
            foreach (MenuItem i in cornerMenu.Items) i.IsChecked = (Corner)i.Tag == corner;
        };
        return menu;
    }

    // ---- Timer ----

    void Toggle()
    {
        if (running) PauseTimer();
        else Start();
    }

    void Start()
    {
        if (remaining <= TimeSpan.Zero) remaining = length;
        arc.Stroke = Accent;
        endAt = DateTime.UtcNow + remaining;
        running = true;
        ticker.Start();
        SetState("FOCUSING");
    }

    void PauseTimer()
    {
        remaining = endAt - DateTime.UtcNow;
        running = false;
        ticker.Stop();
        Paint();
        SetState("PAUSED");
    }

    void Reset()
    {
        running = false;
        ticker.Stop();
        remaining = length;
        arc.Stroke = Accent;
        Paint();
        SetState("READY");
    }

    void Tick()
    {
        remaining = endAt - DateTime.UtcNow;
        if (remaining > TimeSpan.Zero) { Paint(); return; }

        running = false;
        ticker.Stop();
        remaining = TimeSpan.Zero;
        arc.Stroke = Done;
        DrawArc(1); // full green ring marks the finished session
        time.Text = Format(TimeSpan.Zero);
        SetState("COMPLETE");
        SystemSounds.Asterisk.Play();
    }

    void SetLength(int minutes)
    {
        length = TimeSpan.FromMinutes(Math.Clamp(minutes, 1, 180));
        Reset();
        Save();
    }

    void ApplyCustom()
    {
        if (Idle && int.TryParse(custom.Text, out int m) && m != (int)length.TotalMinutes) SetLength(m);
        else custom.Text = ((int)length.TotalMinutes).ToString();
    }

    void SetState(string text)
    {
        state.Text = text;
        state.Foreground = text == "COMPLETE" ? Done : Muted;
        toggleIcon.Text = running ? PauseGlyph : PlayGlyph;

        // Length can only change between sessions, like the web version.
        int minutes = (int)length.TotalMinutes;
        foreach (var (chip, m) in chips)
        {
            chip.Background = m == minutes ? Hover : Track;
            chip.BorderBrush = Accent;
            chip.BorderThickness = new Thickness(m == minutes ? 1 : 0);
            chip.Opacity = Idle ? 1 : 0.45;
        }
        custom.IsEnabled = Idle;
        custom.Opacity = Idle ? 1 : 0.45;
        if (!custom.IsKeyboardFocused) custom.Text = minutes.ToString();
    }

    void Paint()
    {
        time.Text = Format(remaining);
        DrawArc(remaining / length);
    }

    static string Format(TimeSpan t)
    {
        int s = (int)Math.Ceiling(Math.Max(0, t.TotalSeconds));
        return s >= 3600 ? $"{s / 3600}:{s % 3600 / 60:00}:{s % 60:00}" : $"{s / 60:00}:{s % 60:00}";
    }

    // Clockwise arc from 12 o'clock covering `frac` of the ring.
    void DrawArc(double frac)
    {
        if (frac <= 0) { arc.Data = null; return; }
        frac = Math.Min(frac, 0.99999); // a full 360° ArcSegment degenerates to nothing
        double c = RingSize / 2, r = (RingSize - Stroke) / 2, a = frac * 2 * Math.PI;
        var fig = new PathFigure { StartPoint = new Point(c, c - r) };
        fig.Segments.Add(new ArcSegment(new Point(c + r * Math.Sin(a), c - r * Math.Cos(a)),
            new Size(r, r), 0, frac > 0.5, SweepDirection.Clockwise, true));
        arc.Data = new PathGeometry(new[] { fig });
    }

    // ---- Placement (physical pixels, any monitor) ----

    IntPtr CurrentMonitor() => MonitorFromWindow(Hwnd, MONITOR_DEFAULTTONEAREST);

    void SnapToNearestCorner()
    {
        var monitor = CurrentMonitor();
        var wa = WorkArea(monitor);
        GetWindowRect(Hwnd, out var r);
        bool left = (r.Left + r.Right) / 2 < (wa.Left + wa.Right) / 2;
        bool top = (r.Top + r.Bottom) / 2 < (wa.Top + wa.Bottom) / 2;
        SnapTo(monitor, top ? (left ? Corner.TopLeft : Corner.TopRight) : (left ? Corner.BottomLeft : Corner.BottomRight));
    }

    void SnapTo(IntPtr monitor, Corner c)
    {
        corner = c;
        var wa = WorkArea(monitor); // excludes that monitor's taskbar
        GetWindowRect(Hwnd, out var r);
        int w = r.Right - r.Left, h = r.Bottom - r.Top;
        int gap = (int)(Gap * GetDpiForWindow(Hwnd) / 96.0);
        int x = c is Corner.TopLeft or Corner.BottomLeft ? wa.Left + gap : wa.Right - w - gap;
        int y = c is Corner.TopLeft or Corner.TopRight ? wa.Top + gap : wa.Bottom - h - gap;
        SetWindowPos(Hwnd, IntPtr.Zero, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
        savedCenter = new POINT { X = x + w / 2, Y = y + h / 2 };
        Save();
    }

    static RECT WorkArea(IntPtr monitor)
    {
        var mi = new MONITORINFO { cbSize = Marshal.SizeOf<MONITORINFO>() };
        GetMonitorInfo(monitor, ref mi);
        return mi.rcWork;
    }

    // ---- Settings: corner,minutes,centerX,centerY ----

    void Load()
    {
        try
        {
            var parts = File.ReadAllText(SettingsFile).Split(',');
            Enum.TryParse(parts[0], out corner);
            length = TimeSpan.FromMinutes(Math.Clamp(int.Parse(parts[1]), 1, 180));
            if (parts.Length >= 4) savedCenter = new POINT { X = int.Parse(parts[2]), Y = int.Parse(parts[3]) };
        }
        catch { } // first run or unreadable file: keep defaults
    }

    void Save()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsFile)!);
            var c = savedCenter ?? default;
            File.WriteAllText(SettingsFile, $"{corner},{(int)length.TotalMinutes},{c.X},{c.Y}");
        }
        catch { } // settings are a convenience; never crash the timer over them
    }

    // ---- Small UI helpers ----

    static Border IconButton(TextBlock icon, string tip, Brush bg, Brush hover, Action onClick)
    {
        var b = new Border
        {
            Width = 36, Height = 36, CornerRadius = new CornerRadius(18), Background = bg,
            Child = icon, ToolTip = tip, Cursor = Cursors.Hand,
        };
        b.MouseEnter += (_, _) => b.Background = hover;
        b.MouseLeave += (_, _) => b.Background = bg;
        b.MouseLeftButtonDown += (_, e) => { e.Handled = true; onClick(); }; // don't start a window drag
        return b;
    }

    static TextBlock Glyph(string glyph, Brush fg, double size) => new()
    {
        Text = glyph, FontFamily = new FontFamily("Segoe Fluent Icons, Segoe MDL2 Assets"), FontSize = size,
        Foreground = fg, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center,
    };

    static SolidColorBrush Hex(string hex)
    {
        var b = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
        b.Freeze();
        return b;
    }

    // ---- Win32 ----

    const uint MONITOR_DEFAULTTOPRIMARY = 1, MONITOR_DEFAULTTONEAREST = 2;
    const uint SWP_NOSIZE = 0x1, SWP_NOZORDER = 0x4, SWP_NOACTIVATE = 0x10;

    [StructLayout(LayoutKind.Sequential)] struct POINT { public int X, Y; }
    [StructLayout(LayoutKind.Sequential)] struct RECT { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)] struct MONITORINFO { public int cbSize; public RECT rcMonitor, rcWork; public uint dwFlags; }

    [DllImport("user32.dll")] static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint flags);
    [DllImport("user32.dll")] static extern IntPtr MonitorFromPoint(POINT pt, uint flags);
    [DllImport("user32.dll")] static extern bool GetMonitorInfo(IntPtr monitor, ref MONITORINFO info);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);
    [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
}
