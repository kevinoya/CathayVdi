using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Reflection.Metadata;
using System.Runtime.InteropServices;
using System.Threading;

class Program
{
    static void Main(string[] args)
    {
        // 1. 擷取作業系統版本資訊
        //Console.WriteLine("啟動 winver...");

        // 啟動 winver 指令
        Process winverProcess = Process.Start("winver");

        // 等待 3 秒，讓視窗完全載入
        Thread.Sleep(3000);

        // 擷取 winver 視窗截圖
        IntPtr winverHandle = winverProcess.MainWindowHandle;
        if (winverHandle != IntPtr.Zero)
        {
            CaptureWindow(winverHandle, "winver_screenshot.png");
            Console.WriteLine("winver 截圖已儲存為 winver_screenshot.png");
        }
        else
        {
            Console.WriteLine("無法取得 winver 視窗句柄，截圖失敗。");
        }

        // 關閉 winver 視窗
        if (winverProcess != null && !winverProcess.HasExited)
        {
            winverProcess.Kill();
            Console.WriteLine("已關閉 winver 視窗。");
        }

        // 2. 擷取作業系統更新紀錄
        Console.WriteLine("啟動 PowerShell 視窗以執行指令...");

        // 啟動 PowerShell 並執行指令
        Process powerShellProcess = new Process();
        powerShellProcess.StartInfo.FileName = "powershell.exe";
        powerShellProcess.StartInfo.Arguments = "-NoExit -Command \"Get-Hotfix | Sort-Object -Property InstalledOn \"";
        powerShellProcess.StartInfo.UseShellExecute = true; // 使用系統 shell 開啟新視窗
        powerShellProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal;

        // 啟動指令
        powerShellProcess.Start();

        // 等待 PowerShell 視窗完全顯示
        Thread.Sleep(5000);

        // 將 PowerShell 視窗置頂
        IntPtr powerShellHandle = powerShellProcess.MainWindowHandle;
        if (powerShellHandle != IntPtr.Zero)
        {
            SetWindowPos(powerShellHandle, HWND_TOPMOST, 0, 0, 800, 600, SWP_SHOWWINDOW);

            // 等待 2 秒以確保置頂成功
            Thread.Sleep(2000);

            // 擷取 PowerShell 視窗截圖
            CaptureWindow(powerShellHandle, "hotfix_screenshot.png");
            Console.WriteLine("更新資訊截圖已儲存為 hotfix_screenshot.png");
        }
        else
        {
            Console.WriteLine("無法取得 PowerShell 視窗句柄，截圖失敗。");
        }

        // 關閉 PowerShell 視窗
        if (powerShellProcess != null && !powerShellProcess.HasExited)
        {
            powerShellProcess.Kill();
            Console.WriteLine("已關閉 PowerShell 視窗。");
        }

        // 3. 啟動 Windows Defender 的病毒與威脅防護
        Console.WriteLine("啟動 Windows Defender...");

        // 啟動 Windows 安全性病毒與威脅防護視窗
        Process defenderProcess = new Process();
        defenderProcess.StartInfo.FileName = "cmd.exe";
        defenderProcess.StartInfo.Arguments = "/c start windowsdefender://threat";
        defenderProcess.StartInfo.UseShellExecute = true;
        defenderProcess.Start();

        // 等待視窗完全顯示
        Thread.Sleep(5000);

        // 取得 Windows Defender 視窗句柄（可能需要手動定位）
        IntPtr defenderHandle = GetForegroundWindow(); // 獲取目前前景視窗
        if (defenderHandle != IntPtr.Zero)
        {
            // 將 Defender 視窗置頂
            //SetWindowPos(defenderHandle, HWND_TOPMOST, -1500, -200, 800, 1800, SWP_SHOWWINDOW);
            SetWindowPos(defenderHandle, HWND_TOPMOST, 0, -150, 800, 1800, SWP_SHOWWINDOW);

            // 捕捉視窗上半部分截圖
            Thread.Sleep(2000);
            CaptureWindow(defenderHandle, "defender_screenshot.png");
            Console.WriteLine("Windows Defender 上半部截圖已儲存為 defender_screenshot_top.png");
        }
        else
        {
            Console.WriteLine("無法取得 Windows Defender 視窗句柄，截圖失敗。");
        }

        // 嘗試關閉視窗
        CloseWindow(defenderHandle);

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

                // 儲存為 PNG 檔案
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

        // 發送關閉視窗訊息
        PostMessage(handle, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        Console.WriteLine("已發送關閉 Windows Defender 視窗的訊息。");
    }

    // 匯入 Windows API 函數
    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow(); // 取得前景視窗

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll")]
    private static extern bool PostMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

    private const uint SWP_SHOWWINDOW = 0x0040; // 顯示視窗
    private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1); // 置頂

    // 定義 RECT 結構
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
}
