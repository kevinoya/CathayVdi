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
        Console.WriteLine("啟動 winver...");

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

    // 匯入 Windows API 函數
    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

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
