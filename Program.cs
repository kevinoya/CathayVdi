using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
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

        //產出Word
        CreateWord();
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

    static void CreateWord()
    {
        // 創建或打開一個 Word 文件
        string filePath = "00587992.docx";
        // 創建 Word 文件
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
        {
            // 添加主要部分
            MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();

            // 初始化 Body
            Body body = new Body();
            mainPart.Document.Append(body); // 正確附加 Body

            // 添加標題
            //AddParagraph(body, "系統資訊報告", true, 18, JustificationValues.Center);

            // 1. 作業系統版本
            AddParagraph(body, "1.作業系統版本", true, 14);
            AddImageToWord(mainPart, body, "winver_screenshot.png");

            // 2. 作業系統更新
            AddParagraph(body, "2.作業系統更新", true, 14);
            AddImageToWord(mainPart, body, "hotfix_screenshot.png");

            // 3. 防毒軟體版本與掃毒
            AddParagraph(body, "3.防毒軟體版本", true, 14);
            AddParagraph(body, "4.防毒軟體病毒碼更新日期", true, 14);
            AddParagraph(body, "5.防毒軟體全機完整掃毒", true, 14);
            AddImageToWord(mainPart, body, "defender_screenshot.png");

            // 6. MAC 位址
            AddParagraph(body, "6.MAC位址", true, 14);
            AddImageToWord(mainPart, body, "ipconfig_screenshot.png");

            // 儲存文件
            mainPart.Document.Save();
        }

        Console.WriteLine($"Word 文件已生成：{filePath}");
    }

    static void AddParagraph(Body body, string text, bool isBold = false, int fontSize = 12, JustificationValues? alignment = null)
    {
        Paragraph paragraph = new Paragraph();
        Run run = new Run();
        RunProperties runProps = new RunProperties();

        if (isBold)
        {
            runProps.AppendChild(new Bold());
        }

        runProps.AppendChild(new FontSize() { Val = (fontSize * 2).ToString() });
        run.AppendChild(runProps);
        run.AppendChild(new Text(text));
        paragraph.AppendChild(run);

        // 設定對齊方式，若未指定則使用預設對齊方式
        ParagraphProperties paragraphProps = new ParagraphProperties();
        paragraphProps.AppendChild(new Justification() { Val = alignment ?? JustificationValues.Left });
        paragraph.InsertAt(paragraphProps, 0);

        body.AppendChild(paragraph);
    }

    static void AddImageToWord(MainDocumentPart mainPart, Body body, string imagePath)
    {
        if (!File.Exists(imagePath))
        {
            AddParagraph(body, $"未找到檔案：{imagePath}");
            return;
        }

        ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
        using (FileStream stream = new FileStream(imagePath, FileMode.Open))
        {
            imagePart.FeedData(stream);
        }

        AddImageToBody(mainPart.GetIdOfPart(imagePart), body);
    }

    static void AddImageToBody(string relationshipId, Body body)
    {
        long width = 6 * 914400; // 寬度設為 6 英寸（1 英寸 = 914400 單位）
        long height = 4 * 914400; // 高度設為 4 英寸

        Drawing drawing = new Drawing(
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = width, Cy = height }, // 圖片大小
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent()
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                },
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties()
                {
                    Id = (UInt32Value)1U,
                    Name = "Picture"
                },
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                    new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks() { NoChangeAspect = true }),
                new DocumentFormat.OpenXml.Drawing.Graphic(
                    new DocumentFormat.OpenXml.Drawing.GraphicData(
                        new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties()
                                {
                                    Id = (UInt32Value)0U,
                                    Name = "New Bitmap Image"
                                },
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                            new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                new DocumentFormat.OpenXml.Drawing.Blip() { Embed = relationshipId },
                                new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                            new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                new DocumentFormat.OpenXml.Drawing.Transform2D(
                                    new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                                    new DocumentFormat.OpenXml.Drawing.Extents() { Cx = width, Cy = height }),
                                new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                    new DocumentFormat.OpenXml.Drawing.AdjustValueList()
                                )
                                { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }))
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
            {
                DistanceFromTop = (UInt32Value)0U,
                DistanceFromBottom = (UInt32Value)0U,
                DistanceFromLeft = (UInt32Value)0U,
                DistanceFromRight = (UInt32Value)0U
            });

        Paragraph paragraph = new Paragraph(new Run(drawing));
        body.AppendChild(paragraph);
    }

}
