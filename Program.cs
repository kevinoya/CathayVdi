using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;

class Program
{
    static void Main(string[] args)
    {
        // 1. 擷取作業系統版本資訊
        CaptureSystemVersion();

        // 2. 擷取作業系統更新紀錄
        CaptureSystemUpdates();

        // 3. 啟動 Windows Defender 的病毒與威脅防護
        CaptureWindowsDefender();

        // 6. 取得 MAC 位址
        CaptureMacAddress();
    }

    static void CaptureSystemVersion()
    {
        Console.WriteLine("啟動 winver...");
        Process winverProcess = Process.Start("winver");

        Thread.Sleep(3000); // 等待視窗載入

        IntPtr winverHandle = winverProcess.MainWindowHandle;
        if (winverHandle != IntPtr.Zero)
        {
            CaptureWindow(winverHandle, "winver_screenshot.png");
            Console.WriteLine("winver 截圖已儲存為 winver_screenshot.png");
        }
        else
        {
            Console.WriteLine("無法取得 winver 視窗，截圖失敗。");
        }

        if (winverProcess != null && !winverProcess.HasExited)
        {
            winverProcess.Kill();
            Console.WriteLine("已關閉 winver 視窗。");
        }
        CloseWindow(winverHandle);
    }

    static void CaptureSystemUpdates()
    {
        Console.WriteLine("啟動 PowerShell 視窗以執行指令...");
        Process powerShellProcess = new Process();
        powerShellProcess.StartInfo.FileName = "powershell.exe";
        powerShellProcess.StartInfo.Arguments = "-NoExit -Command \"Get-Hotfix | Sort-Object -Property InstalledOn\"";
        powerShellProcess.StartInfo.UseShellExecute = true;
        powerShellProcess.Start();

        Thread.Sleep(5000); // 等待視窗載入

        IntPtr powerShellHandle = powerShellProcess.MainWindowHandle;
        if (powerShellHandle != IntPtr.Zero)
        {
            SetWindowPos(powerShellHandle, HWND_TOPMOST, 0, 0, 800, 600, SWP_SHOWWINDOW);
            Thread.Sleep(2000);
            CaptureWindow(powerShellHandle, "hotfix_screenshot.png");
            Console.WriteLine("更新資訊截圖已儲存為 hotfix_screenshot.png");
        }
        else
        {
            Console.WriteLine("無法取得 PowerShell 視窗，截圖失敗。");
        }

        if (powerShellProcess != null && !powerShellProcess.HasExited)
        {
            powerShellProcess.Kill();
            Console.WriteLine("已關閉 PowerShell 視窗。");
        }
        CloseWindow(powerShellHandle);
    }

    static void CaptureWindowsDefender()
    {
        Console.WriteLine("啟動 Windows Defender...");
        Process defenderProcess = new Process();
        defenderProcess.StartInfo.FileName = "cmd.exe";
        defenderProcess.StartInfo.Arguments = "/c start windowsdefender://threat";
        defenderProcess.StartInfo.UseShellExecute = true;
        defenderProcess.Start();

        Thread.Sleep(5000); // 等待視窗載入

        IntPtr defenderHandle = GetForegroundWindow();
        if (defenderHandle != IntPtr.Zero)
        {
            SetWindowPos(defenderHandle, HWND_TOPMOST, 0, -150, 800, 1800, SWP_SHOWWINDOW);
            Thread.Sleep(2000);
            CaptureWindow(defenderHandle, "defender_screenshot.png");
            Console.WriteLine("Windows Defender 截圖已儲存為 defender_screenshot.png");
        }
        else
        {
            Console.WriteLine("無法取得 Windows Defender 視窗，截圖失敗。");
        }


        if (defenderProcess != null && !defenderProcess.HasExited)
        {
            defenderProcess.Kill();
            Console.WriteLine("已關閉 Windows Defender 視窗。");
        }
        CloseWindow(defenderHandle);
    }

    static void CaptureMacAddress()
    {
        Console.WriteLine("啟動 ipconfig /all...");
        Process ipconfigProcess = new Process();
        ipconfigProcess.StartInfo.FileName = "cmd.exe";
        ipconfigProcess.StartInfo.Arguments = "/k ipconfig /all";
        ipconfigProcess.StartInfo.UseShellExecute = true;
        ipconfigProcess.Start();

        Thread.Sleep(5000); // 等待視窗載入

        IntPtr ipconfigHandle = GetForegroundWindow();
        if (ipconfigHandle != IntPtr.Zero)
        {
            SetWindowPos(ipconfigHandle, HWND_TOPMOST, 0, 0, 800, 600, SWP_SHOWWINDOW);
            CaptureWindow(ipconfigHandle, "ipconfig_screenshot.png");
            Console.WriteLine("ipconfig /all 截圖已儲存為 ipconfig_screenshot.png");
        }
        else
        {
            Console.WriteLine("無法取得 cmd 視窗，截圖失敗。");
        }

        if (ipconfigProcess != null && !ipconfigProcess.HasExited)
        {
            ipconfigProcess.Kill();
            Console.WriteLine("已關閉 cmd 視窗。");
        }
        CloseWindow(ipconfigHandle);
    }

    static void CaptureWindow(IntPtr handle, string fileName)
    {
        try
        {
            RECT rect;
            GetWindowRect(handle, out rect);

            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;

            using (Bitmap screenshot = new Bitmap(width, height))
            {
                using (Graphics g = Graphics.FromImage(screenshot))
                {
                    g.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(width, height));
                }

                screenshot.Save(fileName, ImageFormat.Png);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("擷取視窗失敗: " + ex.Message);
        }
    }

    static void CloseWindow(IntPtr handle)
    {
        const int WM_CLOSE = 0x0010;
        PostMessage(handle, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        Console.WriteLine("已發送關閉視窗的訊息。");
    }

    // 匯入 Windows API 函數
    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll")]
    private static extern bool PostMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

    private const uint SWP_SHOWWINDOW = 0x0040;
    private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);

    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
}
