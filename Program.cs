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
        Process.Start("winver");

        // 等待 3 秒，讓視窗完全載入
        Thread.Sleep(3000);

        // 擷取螢幕截圖
        CaptureScreenshot("winver_screenshot.png");

        Console.WriteLine("截圖已儲存為 winver_screenshot.png");
        Console.WriteLine("請檢查程式執行目錄中的檔案。");

        // 2. 擷取作業系統更新紀錄
        Console.WriteLine("執行 PowerShell 指令以取得作業系統更新...");

        // 啟動 PowerShell 並執行指令
        Process process = new Process();
        process.StartInfo.FileName = "cmd.exe";
        process.StartInfo.Arguments = "/c powershell Get-Hotfix | Sort-Object -Property InstalledOn -Descending";
        process.StartInfo.UseShellExecute = true;
        process.StartInfo.WindowStyle = ProcessWindowStyle.Normal;

        // 啟動指令
        process.Start();

        // 等待視窗完全顯示
        Thread.Sleep(3000);

        // 將 cmd 視窗置頂
        IntPtr handle = process.MainWindowHandle;
        SetWindowPos(handle, HWND_TOPMOST, 0, 0, 800, 600, SWP_SHOWWINDOW);

        // 等待 2 秒以確保置頂成功
        Thread.Sleep(2000);

        // 擷取螢幕截圖
        CaptureScreenshot("hotfix_screenshot.png");

        Console.WriteLine("更新資訊截圖已儲存為 hotfix_screenshot.png");
        Console.WriteLine("請檢查程式執行目錄中的檔案。");
    }

    static void CaptureScreenshot(string fileName)
    {
        try
        {
            // 獲取螢幕大小
            int screenWidth = GetSystemMetrics(SM_CXSCREEN);
            int screenHeight = GetSystemMetrics(SM_CYSCREEN);

            // 建立與螢幕相同大小的 Bitmap
            using (Bitmap screenshot = new Bitmap(screenWidth, screenHeight))
            {
                using (Graphics g = Graphics.FromImage(screenshot))
                {
                    // 擷取整個螢幕
                    g.CopyFromScreen(0, 0, 0, 0, screenshot.Size);
                }

                // 儲存為 PNG 檔案
                screenshot.Save(fileName, ImageFormat.Png);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("擷取螢幕失敗: " + ex.Message);
        }
    }

    // 匯入 Windows API 函數取得螢幕大小
    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int nIndex);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    private const int SM_CXSCREEN = 0; // 螢幕寬度
    private const int SM_CYSCREEN = 1; // 螢幕高度

    private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1); // 置頂
    private const uint SWP_SHOWWINDOW = 0x0040; // 顯示視窗
}
