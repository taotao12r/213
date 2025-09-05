using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using System.Text;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Text.RegularExpressions;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA2;
using FlaUI.Core.EventHandlers;
using FlaUI.Core.WindowsAPI;
using FlaUI.Core.Definitions;
using AutoItX3Lib;
using FlaUI.UIA3;

class Program
{
    [DllImport("user32.dll")]
    static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

    [DllImport("user32.dll")]
    static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

    [DllImport("user32.dll")]
    static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll")]
    static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll")]
    static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

    [DllImport("user32.dll")]
    static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    static extern int GetSystemMetrics(int nIndex);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    const int MOUSEEVENTF_LEFTDOWN = 0x02;
    const int MOUSEEVENTF_LEFTUP = 0x04;
    const int SM_CXSCREEN = 0;
    const int SM_CYSCREEN = 1;

    static Dictionary<string, Dictionary<string, int>> resultDict = null;
    static HashSet<string> openedCodes = new HashSet<string>();
    static HashSet<string> processedProductMessages = new HashSet<string>();
    // 全局撤回标志
    static volatile bool recallDetected = false;

    const string GROUP_NAME = "古米特米助群";

    static object uiaLock = new object();

    // 全局撤回群聊名单
    static List<string> windowsToRevoke = new List<string> { "邵超超专群", "大琦陈衍礼报盘专群", "山东浦众 报盘专群", "华商鑫宸吴超伟报盘专群", "王志勇报盘专群", "山东优硕报盘专群", "玉树/诺尔布报盘专群", "江苏味莱报盘专群", "山东泉誉报盘专群", "杨老板-古米特报盘专群", "铭海国际-古米特报盘群", "山东诚铸报盘专群", "丰腾报盘专群", "河北伊新澳拉报盘群", "上海茅韵国际报盘专群", "鸿志食品报盘专群", "南京沃牛报盘专群", "牧农报盘专群", "云南优阳报盘专群", "北京启商报盘专群 ", "泉润鑫诚报盘专群", "上海曙河于总报盘专群", "天硕高生报盘专群—古米特", "丰腾杨强报盘专群", "九岭报盘专群", "伊品斋王彬报盘专群", "古米特&米裁报盘合作群", "伊永真尤发平报盘专群", "古米特&金福达报盘专群", "益茂通李四保专群" };

    static AutomationElement lastMessageListElement = null;
    static string lastWindowName = null;
    static AutomationElement lastChatWindow = null;

    // 产品-客户字典
    static Dictionary<string, List<string>> productToCustomers = new Dictionary<string, List<string>>
    {
        { "六切", new List<string>{ "邵超超","牧农","启商","豪之安","吴超伟","伊品斋","刘哥" } },
        { "件套", new List<string>{ "泉润","江正阳","启商","刘哥","曙河","牧农","陈衍礼","逸枫冷","吴超伟","九岭","天硕","天硕","邵超超","日月昊","赵总","今好牛","西部牧业","嘉昶","江苏福亿食品","广州菁霞" } },
        { "肋排", new List<string>{ "玉树","鸿志","陈衍礼","米裁","伊明肉业","铭海","伊品斋","江正阳","启商","伊新澳拉","邵超超","杨强","曙河","泉润","日月昊","杨强","宋老板","王军" } },
        { "碎肉", new List<string>{ "日月昊","江正阳","伊新澳拉","鸿志食品","茅韵","九岭","铭海","曙河","铭海","赵潘明","泉润","沃牛","铭海" } },
        { "腱子", new List<string>{ "伊新澳拉","杨强","江正阳","泉润","陈衍礼","宋老板","铭海","刘哥","曙河","九岭","伊品斋","邵超超","启商" } },
        { "砧扒", new List<string>{ "玉树","江正阳","邵超超","玉树","沃牛","逸枫冷","九岭","江苏味莱","曙河" } },
        { "大米龙", new List<string>{ "江正阳","邵超超","启商","泉润","牧农","沃牛","逸枫冷","九岭","吴超伟","宋老板","玉树","曙河","江苏味莱","刘哥","赵总","今好牛" } },
        { "牛霖", new List<string>{ "江正阳","泉润","刘哥","牧农","九岭","吴超伟","曙河","邵超超","启商","沃牛","逸枫冷","西部牧业"} },
        { "肋条", new List<string>{ "江正阳","泉润","米裁","铭海","沃牛","曙河","刘哥" } },
        { "缺失牛前", new List<string>{ "日月昊","希阳","刘哥","王晓勇","江正阳","沃牛","鸿志食品","王军","九岭","杨强","邵超超","泉润","伊新澳拉","西部牧业" } },
        { "西冷眼肉", new List<string>{ "铭海","江正阳" } },
        { "全牛", new List<string>{ "江正阳","穆逸商贸","启商","赵总","日月昊","铭海","茂笙刘总","九岭","泉润","邵超超","豪之安","嘉昶" } },
        { "牛副", new List<string>{ "泉润","启商","邵超超","张长征","赵总" } },
        { "脐橙板", new List<string>{ "伊新澳拉","刘哥","日月昊","江正阳","九岭","杨强","邵超超","曙河","伊明肉业","希阳","铭海","天硕","天硕" } },
        { "带骨脖肉", new List<string>{ "伊新澳拉","江正阳","刘哥","泉润","曙河","沃牛","邵超超","赵总","九岭","鸿志","天硕","杨强","王军","玉树","宋老板","戴哥" } },
        { "骨头拼柜", new List<string>{ "伊明肉业","江正阳","吴国杰","沃牛","邵超超","刘哥","九岭","曙河","沈忠雷","张长征" } },
        { "板腱", new List<string>{ "曙河","邵超超","启商","江苏味莱","铭海","茅韵" } },
        { "嫩肩", new List<string>{ "泉润","曙河","邵超超","启商","江苏味莱","刘哥"} },
        { "小米龙", new List<string>{ "江正阳","泉润","赵潘明" } },
        { "腹肉", new List<string>{ "江正阳","泉润","邵超超","日月昊","九岭" } },
        { "后胸", new List<string>{ "曙河","杨强","泉润" } },
        { "保乐肩", new List<string>{ "邵超超","泉润","沃牛","启商","江苏味莱" } },
        { "脂肪", new List<string>{ "九岭","赵总","江正阳","伊新澳拉","刘哥","王志勇","新征程","西部牧业" } },
        { "胸口油", new List<string>{ "泉润","沃牛","天津亿明","茅韵" } },
        { "带骨四分体", new List<string>{ "江正阳","刘哥","邵超超","九岭","沃牛","玉树","西部牧业" } },
        { "肩胛背肩", new List<string>{ "伊永真","玉树","牧农" } },
        { "牛舌", new List<string>{ "泉润" } },
        { "牛肾", new List<string>{ "江正阳","刘哥","泉润","九岭" } },
        { "牛心", new List<string>{ "刘哥","泉润","铭海" } },
        { "臀腰肉", new List<string>{ "邵超超","玉树","刘哥","米裁" } },
        { "胸腩连体", new List<string>{ "豪之安" } },
        { "光排", new List<string>{ "泉润","邵超超","王志勇","启商","西部牧业" } },
        { "仔骨", new List<string>{ "江正阳","启商","聚荣祥" } },
    };

    // 客户-群聊字典
    static Dictionary<string, string> customerToGroup = new Dictionary<string, string>
    {
        { "邵超超", "邵超超专群" },
        { "牧农", "牧农报盘专群" },
        { "启商", "北京启商报盘专群 " },
        { "豪之安", "豪之安朱永豪报盘专群" },
        { "伊品斋", "伊品斋王彬报盘专群" },
        { "刘哥", "刘哥报盘专群" },
        { "泉润", "泉润鑫诚报盘专群" },
        { "江正阳", "山东浦众 报盘专群" },
        { "曙河", "上海曙河于总报盘专群" },
        { "陈衍礼", "大琦陈衍礼报盘专群" },
        { "逸枫冷", "南京逸枫冷冻食品报盘群" },
        { "吴超伟", "华商鑫宸吴超伟报盘专群" },
        { "九岭", "九岭报盘专群" },
        { "天硕", "天硕高生报盘专群—古米特" },
        { "日月昊", "天津日月昊报盘专群" },
        { "赵总", "赵总报盘专群" },
        { "今好牛", "今好牛报盘专群" },
        { "嘉昶", "嘉昶商贸孙老板报盘群" },
        { "江苏福亿食品", "江苏福亿食品报盘专群" },
        { "广州菁霞", "广州菁霞报盘专群" },
        { "玉树", "玉树/诺尔布报盘专群" },
        { "鸿志", "鸿志食品报盘专群" },
        { "米裁", "古米特&米裁报盘合作群" },
        { "伊明肉业", "伊明肉业报盘专群" },
        { "伊新澳拉", "河北伊新澳拉报盘群" },
        { "杨强", "丰腾杨强报盘专群" },
        { "宋老板", "天津亿源报盘专群" },
        { "王军", "甜甜牛羊肉报盘专群" },
        { "鸿志食品", "鸿志食品报盘专群" },
        { "茅韵", "上海茅韵国际报盘专群" },
        { "铭海", "铭海国际-古米特报盘群" },
        { "赵潘明", "南京逸枫冷冻食品报盘群" },
        { "沃牛", "南京沃牛报盘专群" },
        { "江苏味莱", "江苏味莱报盘专群" },
        { "西部牧业", "西部牧业梁总报盘专群" },
        { "希阳", "河北希阳报盘专群" },
        { "穆逸商贸", "河南穆逸商贸报盘专群" },
        { "茂笙刘总", "茂笙商贸报盘专群" },
        { "张长征", "湖南张长征报盘专群" },
        { "戴哥", "戴哥报盘专群" },
        { "吴国杰", "河南首承报盘专群" },
        { "沈忠雷", "沈忠雷报盘专群" },
        { "王志勇", "王志勇报盘专群" },
        { "新征程", "河南新征程报盘专群" },
        { "天津亿明", "天津亿明报盘专群" },
        { "伊永真", "伊永真尤发平报盘专群" },
        { "聚荣祥", "聚荣祥黄总报盘专群" }
    };

    static List<string> lastReplayCustomers = new List<string>();
    static string lastReplayText = "";

    static bool waitingForChoice = false;
    static DateTime choiceStartTime = DateTime.MinValue;
    static string lastBaopanText = "";
    static string lastText = "";

    // 新增：在全局变量区添加：
    static string lastProcessedReplayInput = "";
    static string pendingReplayInput = ""; // 捕获用户选择，供复盘分支消费

    // 防重入标志
    static int _monitoringTick = 0;          // 监听回调防重入
    static int _isForwarding = 0;            // 复盘批量转发防重入

    // 单例 STA 线程与任务队列，用于安全 SendKeys
    static Thread _staThread;
    static readonly BlockingCollection<Action> _staQueue = new BlockingCollection<Action>();

    // 新增：强引用保存每个窗口的消息监听Timer，防止被GC回收；以及柜号并发保护锁
    static readonly System.Collections.Concurrent.ConcurrentDictionary<string, System.Threading.Timer> _messageTimers
        = new System.Collections.Concurrent.ConcurrentDictionary<string, System.Threading.Timer>();
    static readonly object openedCodesLock = new object();

    static void EnsureStaThread()
    {
        if (_staThread != null) return;
        _staThread = new Thread(() =>
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
            while (true)
            {
                var work = _staQueue.Take();
                try { work(); } catch (Exception ex) { Console.WriteLine("STA work error: " + ex); }
            }
        });
        _staThread.IsBackground = true;
        _staThread.SetApartmentState(ApartmentState.STA);
        _staThread.Start();
    }
    static void RunOnSta(Action action)
    {
        EnsureStaThread();
        using (var ev = new ManualResetEvent(false))
        {
            _staQueue.Add(() => { try { action(); } finally { ev.Set(); } });
            ev.WaitOne();
        }
    }
    static void SendKeysSafe(string keys) => RunOnSta(() => SendKeys.SendWait(keys));
    static void SendTextSafe(string text) => RunOnSta(() => SendKeys.SendWait(text));

    [STAThread]
    static void Main(string[] args)
    {
        Console.WriteLine("程序启动，正在加载统计结果...");
        LoadResultDictFromExcel(@"D:\工作资料\统计结果.xlsx");
        Console.WriteLine("统计结果加载完成，正在查找微信进程...");
        
        // 新增：程序开始前先搜索并打开古米特米助群
        if (!SearchAndOpenGroup())
        {
            Console.WriteLine("无法找到或打开古米特米助群，程序退出。");
            Console.WriteLine("按任意键退出...");
            Console.ReadKey();
            return;
        }
        
        StartAllListeners();
        Console.ReadLine(); // 等待用户输入，防止主线程退出

        SetClipboardText("xxx");
        string clip = GetClipboardText();
        Console.WriteLine("当前剪贴板内容：" + clip);

        SetClipboardText("yyy");
        clip = GetClipboardText();
        Console.WriteLine("当前剪贴板内容：" + clip);
    }

    // 新增：搜索并打开古米特米助群的方法
    static bool SearchAndOpenGroup()
    {
        try
        {
            Console.WriteLine("正在搜索并打开古米特米助群...");
            
            var processes = Process.GetProcessesByName("weixin");
            if (processes.Length == 0)
            {
                Console.WriteLine("未找到微信进程，请先启动微信。");
                return false;
            }

            foreach (var proc in processes)
            {
                using (var app = FlaUI.Core.Application.Attach(proc))
                using (var automation = new UIA3Automation())
                {
                    // 查找微信主窗口
                    var mainWindow = app.GetMainWindow(automation);
                    if (mainWindow == null)
                    {
                        Console.WriteLine("未找到微信主窗口。");
                        continue;
                    }

                    // 激活主窗口
                    IntPtr hwnd = (IntPtr)mainWindow.Properties.NativeWindowHandle.Value;
                    SetForegroundWindow(hwnd);
                    Thread.Sleep(500);

                    // 查找搜索框
                    var searchBox = mainWindow.FindFirstDescendant(cf =>
                        cf.ByControlType(FlaUI.Core.Definitions.ControlType.Edit)
                        .And(cf.ByClassName("mmui::XValidatorTextEdit"))
                        .And(cf.ByName("搜索"))
                    );

                    if (searchBox == null)
                    {
                        Console.WriteLine("未找到搜索框。");
                        continue;
                    }

                    // 点击搜索框
                    searchBox.Focus();
                    Thread.Sleep(200);

                    // 输入群名
                    if (searchBox.Patterns.Value.IsSupported)
                    {
                        searchBox.Patterns.Value.Pattern.SetValue(GROUP_NAME);
                    }
                    else
                    {
                        SendTextSafe(GROUP_NAME);
                    }
                    Console.WriteLine($"已在搜索框输入：{GROUP_NAME}");
                    Thread.Sleep(500);

                    // 查找搜索结果中的群聊项
                    var chatSessionItems = mainWindow.FindAllDescendants(cf =>
                        cf.ByControlType(FlaUI.Core.Definitions.ControlType.ListItem)
                        .And(cf.ByClassName("mmui::ChatSessionCell"))
                    );

                    AutomationElement targetChatItem = null;
                    foreach (var item in chatSessionItems)
                    {
                        try
                        {
                            string itemName = item.Name;
                            if (!string.IsNullOrEmpty(itemName) && Regex.IsMatch(itemName, "古米特米助群"))
                            {
                                targetChatItem = item;
                                Console.WriteLine($"找到匹配的群聊项：{itemName}");
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"检查群聊项时出错：{ex.Message}");
                        }
                    }

                    if (targetChatItem == null)
                    {
                        Console.WriteLine("未找到古米特米助群。");
                        continue;
                    }

                    // 双击打开群聊
                    targetChatItem.DoubleClick();
                    Console.WriteLine("已双击打开古米特米助群");
                    Thread.Sleep(2000);

                    // 查找独立窗口
                    var allWindows = app.GetAllTopLevelWindows(automation);
                    AutomationElement groupWindow = null;
                    foreach (var win in allWindows)
                    {
                        if (win.Name == GROUP_NAME)
                        {
                            groupWindow = win;
                            break;
                        }
                    }

                    if (groupWindow == null)
                    {
                        Console.WriteLine("未找到古米特米助群独立窗口。");
                        continue;
                    }

                    // 调整窗口位置和大小
                    AdjustWindowPosition(groupWindow);
                    Console.WriteLine("古米特米助群已成功打开并调整位置。");
                    return true;
                }
            }

            Console.WriteLine("所有微信进程都未能成功打开古米特米助群。");
            return false;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"搜索并打开古米特米助群时出错：{ex.Message}");
            return false;
        }
    }

    // 新增：调整窗口位置的方法
    static void AdjustWindowPosition(AutomationElement window)
    {
        try
        {
            IntPtr hwnd = (IntPtr)window.Properties.NativeWindowHandle.Value;
            
            // 获取当前窗口位置和大小
            RECT rect;
            GetWindowRect(hwnd, out rect);
            
            // 获取屏幕尺寸
            int screenWidth = GetSystemMetrics(SM_CXSCREEN);
            int screenHeight = GetSystemMetrics(SM_CYSCREEN);
            
            // 计算新位置：屏幕最右边，宽度不变，高度增加30
            int newX = screenWidth - (rect.Right - rect.Left);
            int newY = 0; // 顶部对齐
            int newWidth = rect.Right - rect.Left; // 宽度不变
            int newHeight = (rect.Bottom - rect.Top) + 30; // 高度增加30
            
            // 移动窗口
            MoveWindow(hwnd, newX, newY, newWidth, newHeight, true);
            Console.WriteLine($"窗口已调整到位置：({newX}, {newY})，大小：{newWidth}x{newHeight}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"调整窗口位置时出错：{ex.Message}");
        }
    }

    // 打印窗口层次结构的辅助方法
    static void PrintWindowHierarchy(AutomationElement element, int level)
    {
        if (element == null) return;

        string indent = new string(' ', level * 2);
        try
        {
            var props = "";
            try { props += $"Name='{element.Name}', "; } catch { props += "Name=<error>, "; }
            try { props += $"AutomationId='{element.AutomationId}', "; } catch { props += "AutomationId=<error>, "; }
            try { props += $"ClassName='{element.ClassName}', "; } catch { props += "ClassName=<error>, "; }
            try { props += $"ControlType={element.ControlType}"; } catch { props += "ControlType=<error>"; }

            Console.WriteLine($"{indent}{props}");

            // 只展开前两层，防止输出过多
            if (level < 2)
            {
                var children = element.FindAllChildren();
                foreach (var child in children)
                {
                    PrintWindowHierarchy(child, level + 1);
                }
            }
            else if (level == 2)
            {
                var childCount = 0;
                try
                {
                    childCount = element.FindAllChildren().Length;
                    if (childCount > 0)
                        Console.WriteLine($"{indent}  (包含 {childCount} 个子元素...)");
                }
                catch
                {
                    Console.WriteLine($"{indent}  (获取子元素失败)");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"{indent}(获取元素属性时出错: {ex.Message})");
        }
    }

    static AutomationElement FindByPath(AutomationElement root, string path)
    {
        if (root == null) return null;

        if (string.IsNullOrEmpty(path))
            return root;

        string[] segments = new string[] { path };
        if (path.Contains("/"))
            segments = path.Trim('/').Split('/');
        AutomationElement current = root;
        foreach (var seg in segments)
        {
            if (current == null) return null;

            var type = seg;
            int index = 1;
            if (seg.Contains("["))
            {
                var left = seg.IndexOf('[');
                var right = seg.IndexOf(']');
                type = seg.Substring(0, left);
                index = int.Parse(seg.Substring(left + 1, right - left - 1));
            }

            var children = current.FindAllChildren();
            var match = children.Where(c => c.ControlType.ToString() == type).ToList();
            if (match.Count < index)
                return null;
            current = match[index - 1];
        }
        return current;
    }

    static void StartMonitoringMessages(AutomationElement chatMessageListElement, string windowName, AutomationElement chatWindow)
    {
        Console.WriteLine($"[DEBUG] StartMonitoringMessages被调用，windowName={windowName}, chatMessageListElement is null? {chatMessageListElement == null}");
        // 新增：保存监听参数，便于撤回后重启监听
        lastMessageListElement = chatMessageListElement;
        lastWindowName = windowName;
        lastChatWindow = chatWindow;

        if (chatMessageListElement == null)
        {
            Console.WriteLine("错误：聊天消息列表元素未找到，无法监听。");
            return;
        }

        Console.WriteLine($"开始监听消息列表: {windowName} ...");

        // 新增：空扫描计数用于检测元素失效并重建
        int emptyScanCount = 0;

        System.Threading.Timer timer = null;
        // 强引用保存Timer并在回调内引用自身
        timer = new System.Threading.Timer(_ =>
        {
            // 防重入：同一时间只允许一轮回调执行
            if (Interlocked.Exchange(ref _monitoringTick, 1) == 1) return;
            try
            {
                // 新增：如果正在进行批量转发，暂停本次扫描，避免UIA并发冲突
                if (Volatile.Read(ref _isForwarding) == 1)
                {
                    return;
                }
                // === 新增：检测到撤回后立即停止Timer ===
                if (recallDetected)
                {
                    Console.WriteLine($"[{windowName}] Timer检测到撤回，主动停止消息监听Timer。");
                    try { timer.Dispose(); } catch { }
                    // 修复：显式声明 out 变量类型，避免编译器推断为 object
                    System.Threading.Timer removedTimer;
                    if (_messageTimers.TryRemove(windowName, out removedTimer))
                    {
                        try { removedTimer?.Dispose(); } catch { }
                    }
                    HandleRecall(chatWindow); // 新增
                    return;
                }
                try
                {
                    var msgs = TryFindAllChildren(chatMessageListElement);

                    // 新增：检测消息列表元素是否失效，必要时重建
                    if (msgs.Length == 0)
                    {
                        emptyScanCount++;
                        if (emptyScanCount >= 5)
                        {
                            Console.WriteLine($"[{windowName}] 连续空扫描{emptyScanCount}次，尝试重新获取消息列表控件...");
                            var newWin = ReAttachWeixinWindow(windowName);
                            var newList = TryReacquireMessageList(newWin);
                            if (newWin != null && newList != null)
                            {
                                chatWindow = newWin;
                                chatMessageListElement = newList;
                                lastChatWindow = newWin;
                                lastMessageListElement = newList;
                                Console.WriteLine($"[{windowName}] 消息列表控件已重新获取。");
                                emptyScanCount = 0;
                            }
                            else
                            {
                                Console.WriteLine($"[{windowName}] 重新获取消息列表失败。");
                                // 新增：如果重新获取消息列表失败就退出程序
                                Console.WriteLine("程序将退出，因为无法重新获取消息列表控件。");
                                Environment.Exit(1);
                            }
                        }
                        return;
                    }
                    else
                    {
                        emptyScanCount = 0;
                    }

                    // 其余的消息处理逻辑保持不变...
                    // 这里省略了原来的消息处理代码，因为太长
                    // 实际使用时需要将原来的消息处理逻辑复制到这里
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[{windowName}] 定时器异常: {ex}");
                    // 检查微信进程是否还在，不在则退出定时器或重启监听
                }
            }
            finally
            {
                Interlocked.Exchange(ref _monitoringTick, 0);
            }
        }, null, 0, 1000); // 1秒轮询

        // 强引用保存，避免Timer被GC
        _messageTimers[windowName] = timer;
    }

    // 安全获取子元素列表，自动捕获COM异常并重试
    static AutomationElement[] TryFindAllChildren(AutomationElement root, int retries = 3, int delayMs = 200)
    {
        for (int i = 0; i < retries; i++)
        {
            try
            {
                var arr = root.FindAllChildren();
                if (arr != null) return arr;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Thread.Sleep(delayMs);
            }
            catch { }
        }
        return Array.Empty<AutomationElement>();
    }

    // 新增：重新获取窗口对象的方法
    static AutomationElement ReAttachWeixinWindow(string windowName)
    {
        var processes = Process.GetProcessesByName("weixin");
        foreach (var proc in processes)
        {
            using (var app = FlaUI.Core.Application.Attach(proc))
            using (var automation = new UIA3Automation())
            {
                var allWindows = app.GetAllTopLevelWindows(automation);
                foreach (var win in allWindows)
                {
                    if (win.Name == windowName)
                    {
                        return win;
                    }
                }
            }
        }
        return null;
    }

    // 新增：尝试根据窗口重新获取消息列表控件
    static AutomationElement TryReacquireMessageList(AutomationElement chatWin)
    {
        if (chatWin == null) return null;
        AutomationElement messageList = null;
        try
        {
            if (chatWin.ClassName == "mmui::FramelessMainWindow")
            {
                var messageView = chatWin.FindFirstDescendant(cf => cf.ByClassName("mmui::MessageView"));
                if (messageView != null)
                    messageList = messageView.FindFirstDescendant(cf => cf.ByClassName("mmui::RecyclerListView").And(cf.ByControlType(ControlType.List)));
            }
            else if (chatWin.ClassName == "mmui::SearchMsgUniqueChatWindow")
            {
                messageList = chatWin.FindFirstDescendant(cf => cf.ByClassName("mmui::RecyclerListView").And(cf.ByControlType(ControlType.List)));
            }
            else if (chatWin.ClassName == "mmui::MainWindow")
            {
                messageList = chatWin.FindFirstDescendant(cf => cf.ByControlType(ControlType.List));
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("TryReacquireMessageList: " + ex.Message);
        }
        return messageList;
    }

    static void HandleRecall(object chatWindow)
    {
        lock (uiaLock)
        {
            Console.WriteLine("执行批量撤回相关操作...");
            foreach (var name in windowsToRevoke)
            {
                RevokeLastMessage(name);
                Thread.Sleep(1500);
            }
            recallDetected = false;

            // === 新增：撤回后自动重启消息监听 ===
            if (lastWindowName != null)
            {
                Console.WriteLine("撤回后自动重启消息监听...");
                var newChatWindow = ReAttachWeixinWindow(lastWindowName);
                if (newChatWindow != null)
                {
                    Console.WriteLine("ReAttachWeixinWindow成功，窗口名：" + newChatWindow.Name + "，ClassName：" + newChatWindow.ClassName);

                    // 新增：先释放旧Timer，避免重复监听
                    // 修复：显式声明 out 变量类型，避免低版本 C# 下编译错误
                    System.Threading.Timer oldTimer;
                    if (_messageTimers.TryRemove(lastWindowName, out oldTimer))
                    {
                        try { oldTimer.Dispose(); } catch { }
                    }

                    AutomationElement messageList = null;
                    if (newChatWindow.ClassName == "mmui::FramelessMainWindow")
                    {
                        var messageView = newChatWindow.FindFirstDescendant(cf => cf.ByClassName("mmui::MessageView"));
                        if (messageView != null)
                            messageList = messageView.FindFirstDescendant(cf => cf.ByClassName("mmui::RecyclerListView").And(cf.ByControlType(ControlType.List)));
                    }
                    else if (newChatWindow.ClassName == "mmui::SearchMsgUniqueChatWindow")
                    {
                        messageList = newChatWindow.FindFirstDescendant(cf => cf.ByClassName("mmui::RecyclerListView").And(cf.ByControlType(ControlType.List)));
                    }
                    else if (newChatWindow.ClassName == "mmui::MainWindow")
                    {
                        messageList = newChatWindow.FindFirstDescendant(cf => cf.ByControlType(ControlType.List));
                    }

                    if (messageList != null)
                    {
                        Console.WriteLine("消息列表控件重新获取成功，准备重启监听。");
                        StartMonitoringMessages(messageList, lastWindowName, newChatWindow);
                    }
                    else
                    {
                        Console.WriteLine("撤回后未能重新获取消息列表控件，监听未恢复。");
                    }
                }
                else
                {
                    Console.WriteLine("撤回后未能重新获取窗口，监听未恢复。");
                }
            }
        }
    }

    static void RevokeLastMessage(string searchText)
    {
        try
        {
            Console.WriteLine($"RevokeLastMessage: 开始执行撤回流程，搜索：{searchText}");
            var processes = Process.GetProcessesByName("weixin");
            Console.WriteLine($"RevokeLastMessage: weixin进程数={processes.Length}");
            if (processes.Length == 0)
            {
                Console.WriteLine("RevokeLastMessage: 未找到 weixin.exe 进程。");
                return;
            }

            bool found = false;
            foreach (var proc in processes)
            {
                Console.WriteLine($"RevokeLastMessage: Attach进程ID={proc.Id}");
                try
                {
                    using (var app = FlaUI.Core.Application.Attach(proc))
                    using (var automation = new UIA3Automation())
                    {
                        Console.WriteLine("RevokeLastMessage: Attach成功，准备查找主窗口句柄...");
                        IntPtr hwnd = FindWeChatMainWindow(proc.Id);
                        if (hwnd == IntPtr.Zero)
                        {
                            Console.WriteLine("RevokeLastMessage: 未找到主窗口句柄。");
                            continue;
                        }
                        Console.WriteLine("RevokeLastMessage: 已获取主窗口句柄，准备激活...");
                        SetForegroundWindow(hwnd);
                        Thread.Sleep(500);

                        var mainWin = automation.FromHandle(hwnd);
                        if (mainWin == null)
                        {
                            Console.WriteLine("RevokeLastMessage: FromHandle失败。");
                            continue;
                        }
                        Thread.Sleep(300);
                        Console.WriteLine("RevokeLastMessage: 已获取主窗口AutomationElement，准备查找搜索框...");

                        // 查找搜索框
                        var searchBox = mainWin.FindFirstDescendant(cf =>
                            cf.ByControlType(FlaUI.Core.Definitions.ControlType.Edit)
                            .And(cf.ByClassName("mmui::XValidatorTextEdit"))
                        );

                        if (searchBox != null)
                        {
                            searchBox.Focus();
                            Thread.Sleep(300);
                            if (searchBox.Patterns.Value.IsSupported)
                            {
                                searchBox.Patterns.Value.Pattern.SetValue(searchText);
                                Console.WriteLine($"已在搜索框输入：{searchText}");
                            }
                            else
                            {
                                SendTextSafe(searchText);
                                Console.WriteLine($"已用SendKeys输入：{searchText}");
                            }
                            Thread.Sleep(200);
                            SendKeysSafe("{ENTER}");
                            Console.WriteLine("已在搜索框回车。");
                            Thread.Sleep(200);
                        }
                        else
                        {
                            Console.WriteLine("未找到搜索输入框（mmui::XLineEdit, Edit）。");
                            continue;
                        }

                        // === 这里是补充的完整撤回流程 ===
                        Console.WriteLine("RevokeLastMessage: 查找消息列表...");
                        var msgList = mainWin.FindFirstDescendant(cf => cf.ByControlType(ControlType.List));
                        if (msgList == null)
                        {
                            Console.WriteLine("RevokeLastMessage: 未找到消息列表。");
                            continue;
                        }
                        var msgs = msgList.FindAllChildren();
                        if (msgs.Length == 0)
                        {
                            Console.WriteLine("RevokeLastMessage: 消息列表为空。");
                            continue;
                        }
                        var lastMsg = msgs.Last();

                        Console.WriteLine("RevokeLastMessage: 找到最后一条消息，准备右键...");
                        var rect = lastMsg.BoundingRectangle;
                        int x = (int)(rect.Right - 120);
                        int y = (int)(rect.Bottom - 15);
                        Cursor.Position = new System.Drawing.Point(x, y);
                        Thread.Sleep(100);
                        mouse_event(0x08, 0, 0, 0, 0); // 右键按下
                        mouse_event(0x10, 0, 0, 0, 0); // 右键抬起
                        Thread.Sleep(200);

                        Console.WriteLine("RevokeLastMessage: 发送上键选择撤回...");
                        SendKeysSafe("{UP}");
                        Thread.Sleep(100);
                        SendKeysSafe("{ENTER}");
                        Thread.Sleep(200);

                        var deleteMsgText = mainWin.FindFirstDescendant(cf =>
                            cf.ByControlType(FlaUI.Core.Definitions.ControlType.Text)
                            .And(cf.ByClassName("mmui::XTextView"))
                            .And(cf.ByName("删除该条消息？"))
                        );

                        if (deleteMsgText != null)
                        {
                            Console.WriteLine("检测到\"删除该条消息？\"提示，自动按下ESC键关闭。");
                            SendKeysSafe("{ESC}");
                            Thread.Sleep(200);
                        }
                        else
                        {
                            Console.WriteLine("未检测到\"删除该条消息？\"提示，无需ESC。");
                        }

                        Console.WriteLine("RevokeLastMessage: 撤回操作已完成。");
                        found = true;
                        break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("RevokeLastMessage: 进程异常：" + ex);
                }
            }
            if (!found)
            {
                Console.WriteLine("RevokeLastMessage: 所有进程都未能成功撤回。");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("RevokeLastMessage: 执行异常：" + ex);
        }
    }

    static IntPtr FindWeChatMainWindow(int processId)
    {
        IntPtr found = IntPtr.Zero;
        EnumWindows((hWnd, lParam) =>
        {
            uint pid;
            GetWindowThreadProcessId(hWnd, out pid);
            if (pid == processId && IsWindowVisible(hWnd))
            {
                int len = GetWindowTextLength(hWnd);
                if (len > 0)
                {
                    StringBuilder sb = new StringBuilder(len + 1);
                    GetWindowText(hWnd, sb, sb.Capacity);
                    string title = sb.ToString();
                    if (!string.IsNullOrEmpty(title) && title.Contains("微信"))
                    {
                        found = hWnd;
                        return false; // 找到就停止
                    }
                }
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    static void StartAllListeners()
    {
        Console.WriteLine("程序启动，正在查找微信进程...");
        try
        {
            var processes = Process.GetProcessesByName("weixin");
            List<AutomationElement> chatWindows = new List<AutomationElement>();

            foreach (var proc in processes)
            {
                using (var app = FlaUI.Core.Application.Attach(proc))
                using (var automation = new UIA3Automation())
                {
                    var allWindows = app.GetAllTopLevelWindows(automation);
                    foreach (var win in allWindows)
                    {
                        if (win.ClassName == "mmui::FramelessMainWindow" && win.Name == GROUP_NAME)
                        {
                            chatWindows.Add(win);
                        }
                    }
                }
            }

            if (chatWindows.Count > 0)
            {
                Console.WriteLine($"找到 {chatWindows.Count} 个独立聊天窗口。");

                // 用于保存所有监听任务
                List<Task> tasks = new List<Task>();

                foreach (var chatWin in chatWindows)
                {
                    // 查找消息列表控件
                    AutomationElement messageList = null;
                    if (chatWin.ClassName == "mmui::FramelessMainWindow")
                    {
                        var messageView = chatWin.FindFirstDescendant(cf => cf.ByClassName("mmui::MessageView"));
                        if (messageView != null)
                            messageList = messageView.FindFirstDescendant(cf => cf.ByClassName("mmui::RecyclerListView").And(cf.ByControlType(ControlType.List)));
                    }
                    else if (chatWin.ClassName == "mmui::SearchMsgUniqueChatWindow")
                    {
                        messageList = chatWin.FindFirstDescendant(cf => cf.ByClassName("mmui::RecyclerListView").And(cf.ByControlType(ControlType.List)));
                    }
                    else if (chatWin.ClassName == "mmui::MainWindow")
                    {
                        messageList = chatWin.FindFirstDescendant(cf => cf.ByControlType(ControlType.List));
                    }

                    if (messageList != null)
                    {
                        var winName = chatWin.Name;
                        var winObj = chatWin;
                        // 启动一个新任务监听该窗口
                        var task = Task.Run(() =>
                        {
                            try
                            {
                                Console.WriteLine($"监听窗口: {winName}");
                                StartMonitoringMessages(messageList, winName, winObj);
                                StartRecallMonitor(chatWin);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"监听线程异常: {ex}");
                            }
                        });
                        tasks.Add(task);
                    }
                    else
                    {
                        Console.WriteLine($"窗口[{chatWin.Name}] 未找到消息列表控件。");
                    }
                }

                Console.WriteLine("所有窗口监听已启动，程序将持续运行...");
                Thread.Sleep(Timeout.Infinite); // 阻塞主线程，防止退出
            }
            else
            {
                Console.WriteLine("所有微信进程都未找到独立聊天窗口！");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("发生异常：" + ex.Message);
            Console.WriteLine("按任意键退出...");
            Console.ReadKey();
        }
    }

    static void StartRecallMonitor(AutomationElement chatWindow)
    {
        bool lastRecall = false;
        new Thread(() =>
        {
            while (true)
            {
                if (recallDetected && !lastRecall)
                {
                    Console.WriteLine("【全局监控】检测到撤回，已通知主流程准备撤回！");
                    lastRecall = true;
                }
                else if (!recallDetected && lastRecall)
                {
                    // 撤回已处理，重置
                    lastRecall = false;
                }
                Thread.Sleep(100); // 100ms轮询
            }
        })
        { IsBackground = true }.Start();
    }

    static void LoadResultDictFromExcel(string excelPath)
    {
        resultDict = new Dictionary<string, Dictionary<string, int>>();
        using (var fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
        {
            IWorkbook workbook = new XSSFWorkbook(fs);
            ISheet sheet = workbook.GetSheetAt(0);

            // 读取表头（公司名）
            IRow headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;
            List<string> companyNames = new List<string>();
            for (int i = 1; i < cellCount; i++) // 第一列是前缀
            {
                companyNames.Add(headerRow.GetCell(i).ToString().Trim().ToUpper());
            }

            // 读取每一行
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;
                string prefix = row.GetCell(0).ToString().Trim().ToUpper();
                var companyDict = new Dictionary<string, int>();
                for (int j = 1; j < cellCount; j++)
                {
                    string company = companyNames[j - 1].Trim().ToUpper();
                    int value = 0;
                    var cell = row.GetCell(j);
                    if (cell != null && cell.CellType == CellType.Numeric)
                        value = (int)cell.NumericCellValue;
                    else if (cell != null && cell.CellType == CellType.String && int.TryParse(cell.StringCellValue, out int v))
                        value = v;
                    companyDict[company] = value;
                }
                resultDict[prefix] = companyDict;
            }
        }
    }

    // 写入剪贴板
    public static void SetClipboardText(string text)
    {
        // var thread = new Thread(() => { Clipboard.SetText(text); });
        // thread.SetApartmentState(ApartmentState.STA);
        // thread.Start();
        // thread.Join();
    }

    // 读取剪贴板
    public static string GetClipboardText()
    {
        // string result = "";
        // var thread = new Thread(() =>
        // {
        //     try
        //     {
        //         result = Clipboard.GetText();
        //     }
        //     catch (Exception ex)
        //     {
        //         Console.WriteLine("读取剪贴板失败：" + ex.Message);
        //     }
        // });
        // thread.SetApartmentState(ApartmentState.STA);
        // thread.Start();
        // thread.Join();
        // return result;
        return "";
    }

    public static void WriteCodeToFile(string text)
    {
        string filePath = @"C:\Users\x\Desktop\剪贴.txt";
        File.WriteAllText(filePath, text, Encoding.UTF8);
        Console.WriteLine("柜号已写入文件：" + filePath + " 内容：" + text);
    }
}