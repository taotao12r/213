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

    // 新增：用于移动与获取窗口矩形
    [DllImport("user32.dll", SetLastError = true)]
    static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

    [DllImport("user32.dll", SetLastError = true)]
    static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [StructLayout(LayoutKind.Sequential)]
    struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    const int MOUSEEVENTF_LEFTDOWN = 0x02;
    const int MOUSEEVENTF_LEFTUP = 0x04;

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

        // 新增：启动时执行靠左+搜索+双击+独立窗靠右增高
        OpenGroupAtStartup();

        StartAllListeners();
        Console.ReadLine(); // 等待用户输入，防止主线程退出

        SetClipboardText("xxx");
        string clip = GetClipboardText();
        Console.WriteLine("当前剪贴板内容：" + clip);

        SetClipboardText("yyy");
        clip = GetClipboardText();
        Console.WriteLine("当前剪贴板内容：" + clip);
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

                    // 新逻辑：连续空扫描>=5次则直接退出程序
                    if (msgs.Length == 0)
                    {
                        emptyScanCount++;
                        if (emptyScanCount >= 5)
                        {
                            Console.WriteLine($"[{windowName}] 连续空扫描{emptyScanCount}次，程序即将退出。");
                            Environment.Exit(0);
                        }
                        return;
                    }
                    else
                    {
                        emptyScanCount = 0;
                    }

                    if (msgs.Length > 0)
                    {
                        var last = msgs.Last();
                        var text = TryExtractMessageText(last);
                        var normText = Normalize(text);

                        // 1. 跳过系统消息，不更新lastText
                        if (text.Contains("图片") || text.Contains("已发送："))
                        {
                            // 直接return，不更新lastText
                            return;
                        }

                        // 只在检测到新消息时处理
                        if (!string.IsNullOrEmpty(normText) && normText != Normalize(lastText))
                        {
                            lastText = text; // 只在这里更新
                            Console.WriteLine($"[{windowName}] 检测到新消息: {text}");

                            // 检查是否是"匹配到是xxx客户"
                            if (text.StartsWith("匹配到") && text.Contains("客户"))
                            {
                                // 只在未等待选择时才设置
                                if (!waitingForChoice)
                                {
                                    waitingForChoice = true;
                                    // 解析客户列表
                                    var matchedProducts = ExtractProductsFromText(text, productToCustomers.Keys);
                                    var allCustomers = new List<string>();
                                    foreach (var prod in matchedProducts)
                                        if (productToCustomers.ContainsKey(prod))
                                            allCustomers.AddRange(productToCustomers[prod]);
                                    allCustomers = allCustomers.Distinct().ToList();
                                    lastReplayCustomers = allCustomers;
                                    lastReplayText = text;
                                    // 发送提示输入
                                }
                                return;
                            }

                            // 检查是否是用户输入选择
                            if (waitingForChoice && lastReplayCustomers.Count > 0 && text.Trim() == "1" && text != lastProcessedReplayInput)
                            {
                                // 记录已处理输入，但不清空状态、不提前返回，交由后续复盘批量转发逻辑统一处理
                                lastProcessedReplayInput = text;
                                pendingReplayInput = text.Trim();
                                // 不 return
                            }

                            // 只对新消息里的柜号
                            var matches = Regex.Matches(text, @"[A-Z]{4}\d{7}");
                            foreach (Match m in matches)
                            {
                                string code = m.Value;
                                Console.WriteLine($"[DEBUG] 正在处理柜号: {code}");
                                // 先写入文件
                                WriteCodeToFile(code);
                                // 再查船期（异步且去重）
                                bool shouldStart = false;
                                lock (openedCodesLock)
                                {
                                    if (!openedCodes.Contains(code))
                                    {
                                        openedCodes.Add(code);
                                        shouldStart = true;
                                    }
                                }
                                if (shouldStart)
                                {
                                    Task.Run(() =>
                                    {
                                        try { OpenTopCompanySites(code, 3); }
                                        catch (Exception ex) { Console.WriteLine("OpenTopCompanySites 异常: " + ex); }
                                    });
                                }
                            }

                            // 只对新消息设置撤回标志
                            if ((text.Trim() == "撤回" || text.Trim() == "撤回信息") && !recallDetected)
                            {
                                recallDetected = true;
                                Console.WriteLine("检测到撤回，等待主流程结束后执行撤回。");
                            }
                        }

                        // 2. 在监听到"撤回"消息时，调用批量撤回
                        if (normText == "撤回")
                        {
                            if (!recallDetected)
                            {
                                // Console.WriteLine("检测到撤回，等待主流程结束后执行撤回。");
                            }
                        }
                        else if (normText == "正在撤回")
                        {
                            // 忽略"正在撤回"消息，不做任何处理
                            return;
                        }
                        else
                        {
                            recallDetected = false;
                        }

                        if (windowName.Trim() == GROUP_NAME)
                        {
                            // 检查是否等待选择超时
                            if (waitingForChoice && (DateTime.Now - choiceStartTime).TotalSeconds > 25)
                            {
                                Console.WriteLine($"[{windowName}] 25秒未选择，已自动取消本次操作");
                                SendReplyInWindow(chatWindow, "25秒未选择，已自动取消本次操作");
                                waitingForChoice = false;
                                lastBaopanText = ""; // 清空
                            }

                            // 检查是否包含关键字并未处于等待状态
                            if (
                                !waitingForChoice &&
                                (
                                    (text.Contains("报盘") &&
                                     (text.Contains("美金") || text.Contains("欧元") || text.Contains("澳元") || text.Contains("澳币")))
                                )
                            )
                            {
                                // 原报盘逻辑保持不变
                                string menu =
                                    "输入1 发送给牛肉客户\n" +
                                    "输入2 发送给骨头客户\n" +
                                    "输入3 发送给羊肉客户\n" +
                                    "输入4 发送给猪肉客户\n" +
                                    "输入5 发送新西兰客户\n" +
                                    "输入6 发给碎肉客户\n" +
                                    "输入7 发送给牛肉客户带上梅佳客户\n" +
                                    "输入8 发送给圆切\n" +
                                    "输入0 取消本次操作\n" +
                                    "输入撤回 中断发送完成后在发报盘3分钟以后";
                                SendReplyInWindow(chatWindow, menu);
                                Console.WriteLine($"[{windowName}] 已发送选择菜单");
                                waitingForChoice = true;
                                choiceStartTime = DateTime.Now;
                                lastBaopanText = text; // 存储本次报盘内容
                            }
                            // 新增：复盘逻辑，单独处理
                            else if (!waitingForChoice && text.Contains("复盘"))
                            {
                                var matchedProducts = ExtractProductsFromText(text, productToCustomers.Keys);
                                var allCustomers = new List<string>();
                                string productName = matchedProducts.Count > 0 ? matchedProducts[0] : "";
                                foreach (var prod in matchedProducts)
                                    if (productToCustomers.ContainsKey(prod))
                                        allCustomers.AddRange(productToCustomers[prod]);
                                allCustomers = allCustomers.Distinct().ToList();

                                if (allCustomers.Count == 0)
                                {
                                    SendReplyInWindow(chatWindow, "未匹配到产品客户，请检查产品名称。");
                                    return;
                                }

                                string menu = BuildReplayMenu(allCustomers, matchedProducts);
                                SendReplyInWindow(chatWindow, menu);
                                Console.WriteLine($"[{windowName}] 已发送选择菜单");
                                waitingForChoice = true;
                                choiceStartTime = DateTime.Now;
                                lastReplayCustomers = allCustomers;
                                lastReplayText = text;
                            }
                            // 检测用户输入选择
                            if (waitingForChoice)
                            {
                                string trimmed = text.Trim();
                                if (trimmed == "0")
                                {
                                    Console.WriteLine($"[{windowName}] 用户选择了：0，已取消本次操作");
                                    SendReplyInWindow(chatWindow, "已取消本次操作");
                                    waitingForChoice = false;
                                    lastBaopanText = ""; // 清空
                                    lastReplayCustomers.Clear();
                                    lastReplayText = "";
                                    pendingReplayInput = "";
                                }
                                // 复盘专属选择处理
                                else if (lastReplayCustomers.Count > 0)
                                {
                                    // 防御：如果已处理过，直接return，防止死循环
                                    if (!waitingForChoice || lastReplayCustomers.Count == 0)
                                        return;
                                    Console.WriteLine($"[调试] 进入复盘批量转发分支，lastReplayCustomers.Count={lastReplayCustomers.Count}, waitingForChoice={waitingForChoice}");

                                    // 防重入：一次只跑一遍复盘流程
                                    if (Interlocked.Exchange(ref _isForwarding, 1) == 1) return;

                                    try
                                    {
                                        // 过滤掉系统自动回复和空消息（必须是新用户输入才处理）
                                        if ((string.IsNullOrWhiteSpace(pendingReplayInput) && (text.StartsWith("匹配到") || text.StartsWith("输入无效") || text.StartsWith("未找到客户") || text.StartsWith("20秒未选择") || string.IsNullOrWhiteSpace(text) || text == lastReplayText)))
                                            return;

                                        var inputToUse = string.IsNullOrWhiteSpace(pendingReplayInput) ? text.Trim() : pendingReplayInput.Trim();
                                        var selectedCustomers = ParseUserInputToCustomers(inputToUse, lastReplayCustomers);
                                        if (selectedCustomers.Count == 0)
                                        {
                                            SendReplyInWindow(chatWindow, "输入无效，请重新输入。");
                                            return;
                                        }

                                        // 有效输入，退出等待状态，进入实际转发
                                        waitingForChoice = false;
                                        lastProcessedReplayInput = inputToUse;
                                        pendingReplayInput = "";
                                        // 分组，每9个一组
                                        var batches = SplitByBatch(selectedCustomers, 9);
                                        int rightClickCount = 0;
                                        foreach (var batch in batches)
                                        {
                                            rightClickCount++;
                                            Console.WriteLine($"[调试] 本批次右键最新复盘消息，第 {rightClickCount} 次（本轮应只执行一次）");
                                            // 1. 查找消息列表List控件
                                            var list = FindByPath(chatWindow, "/Group/Group/Group/Custom/Group[2]/List");
                                            if (list == null) { Console.WriteLine("未找到消息列表，跳过本轮"); continue; }
                                            var items = list.FindAllChildren();
                                            AutomationElement replayItem = null;
                                            foreach (var item in items.Reverse())
                                            {
                                                var msgText = TryExtractMessageText(item, 0);
                                                if (!string.IsNullOrEmpty(msgText) && msgText.Contains("复盘"))
                                                {
                                                    replayItem = item;
                                                    break;
                                                }
                                            }
                                            if (replayItem == null) { Console.WriteLine("未找到复盘消息，跳过本轮"); continue; }

                                            // 2. 右键最新"复盘"消息（每一轮只执行一次）
                                            var boundingRect = replayItem.BoundingRectangle;
                                            int x = (int)(boundingRect.Left - 2 + 100);
                                            int y = (int)(boundingRect.Bottom - 2 - 50);
                                            Cursor.Position = new System.Drawing.Point(x, y);
                                            Thread.Sleep(100);
                                            mouse_event(0x08, 0, 0, 0, 0); // 右键按下
                                            mouse_event(0x10, 0, 0, 0, 0); // 右键抬起
                                            Thread.Sleep(200);
                                            for (int i = 0; i < 5; i++) { SendKeysSafe("{DOWN}"); Thread.Sleep(10); }
                                            SendKeysSafe("{ENTER}");
                                            Thread.Sleep(300);

                                            // 3. 查找转发弹窗（限定同一进程，且必须包含"搜索"编辑框）
                                            AutomationElement forwardDialog = null;
                                            IntPtr forwardDlgHwnd = IntPtr.Zero;
                                            var targetPid = chatWindow.Properties.ProcessId.Value;
                                            for (int i = 0; i < 12; i++)
                                            {
                                                var allWindows = GetDesktopWindowsSafe(chatWindow);
                                                foreach (var w in allWindows)
                                                {
                                                    try
                                                    {
                                                        if (w.ControlType != ControlType.Window) continue;
                                                        if (w.Properties.ProcessId.Value != targetPid) continue;
                                                        var sb = w.FindFirstDescendant(cf =>
                                                            cf.ByControlType(FlaUI.Core.Definitions.ControlType.Edit)
                                                              .And(cf.ByClassName("mmui::XValidatorTextEdit"))
                                                              .And(cf.ByName("搜索"))
                                                        );
                                                        if (sb != null)
                                                        {
                                                            forwardDialog = w;
                                                            forwardDlgHwnd = (IntPtr)w.Properties.NativeWindowHandle.Value;
                                                            break;
                                                        }
                                                    }
                                                    catch { }
                                                }
                                                if (forwardDialog != null) break;
                                                Thread.Sleep(200);
                                            }
                                            if (forwardDialog == null)
                                            {
                                                Console.WriteLine("未找到转发弹窗，跳过本轮");
                                                lastReplayCustomers.Clear();
                                                lastReplayText = "";
                                                continue;
                                            }

                                            // 4. 用 forwardDialog 判断内容一致性
                                            if (!HasExactBaopanText(forwardDialog, lastReplayText))
                                            {
                                                SendKeysSafe("{TAB}");
                                                SendKeysSafe("{TAB}");
                                                SendKeysSafe("{TAB}");
                                                SendKeysSafe("{TAB}");
                                                SendKeysSafe("{TAB}");
                                                SendKeysSafe("{ENTER}");
                                                Console.WriteLine("复盘内容不一致，已按下ESC键关闭弹窗。");
                                                continue;
                                            }

                                            // 5. 在弹窗里循环输入客户、勾选CheckBox
                                            foreach (var cust in batch)
                                            {
                                                try
                                                {
                                                    if (!customerToGroup.ContainsKey(cust))
                                                    {
                                                        Console.WriteLine($"[调试] 未找到客户{cust}的群聊名，已跳过。");
                                                        continue;
                                                    }
                                                    string groupName = customerToGroup[cust];
                                                    Console.WriteLine($"[调试] 当前客户: {cust}, 群聊名: {groupName}");

                                                    // 只找弹窗里的"搜索"输入框
                                                    var searchBox = forwardDialog.FindFirstDescendant(cf =>
                                                        cf.ByControlType(FlaUI.Core.Definitions.ControlType.Edit)
                                                        .And(cf.ByClassName("mmui::XValidatorTextEdit"))
                                                        .And(cf.ByName("搜索"))
                                                    );

                                                    if (searchBox != null)
                                                    {
                                                        // 确保弹窗在前台
                                                        IntPtr dlgHwnd = (IntPtr)forwardDialog.Properties.NativeWindowHandle.Value;
                                                        SetForegroundWindow(dlgHwnd);
                                                        Thread.Sleep(200);

                                                        // 先输入群名
                                                        searchBox.Focus();
                                                        Thread.Sleep(150);
                                                        if (searchBox.Patterns.Value.IsSupported)
                                                        {
                                                            searchBox.Patterns.Value.Pattern.SetValue(groupName);
                                                            Console.WriteLine($"已在搜索框输入内容：{groupName}");
                                                        }
                                                        else
                                                        {
                                                            SendTextSafe(groupName);
                                                            Console.WriteLine($"已用SendKeys在搜索框输入内容：{groupName}");
                                                        }
                                                        Thread.Sleep(300);

                                                        // 以键盘方式选择第一项：方向下 + 回车
                                                        SendKeysSafe("{DOWN}");
                                                        Thread.Sleep(20);
                                                        SendKeysSafe("{ENTER}");
                                                        Console.WriteLine("已按下方向下并回车选择第一项。");
                                                        Thread.Sleep(100);

                                                        // 清空搜索框内容，降低结果堆积导致卡顿
                                                        try
                                                        {
                                                            if (searchBox.Patterns.Value.IsSupported)
                                                            {
                                                                searchBox.Patterns.Value.Pattern.SetValue("");
                                                            }
                                                            else
                                                            {
                                                                searchBox.Focus();
                                                                Thread.Sleep(100);
                                                                SendKeysSafe("^a");
                                                                Thread.Sleep(50);
                                                                SendKeysSafe("{DEL}");
                                                            }
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            Console.WriteLine("清空搜索框异常：" + ex.Message);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        Console.WriteLine("未找到搜索输入框，跳过该客户。");
                                                        continue;
                                                    }

                                                    // 这里可以加：查找CheckBox并勾选的逻辑（如果需要自动勾选）
                                                    // ...

                                                }
                                                catch (Exception ex)
                                                {
                                                    Console.WriteLine($"处理客户{cust}时异常: {ex.Message}");
                                                    continue;
                                                }
                                                finally
                                                {
                                                    // 每个客户之间稍作停顿，降低UI阻塞风险
                                                    Thread.Sleep(350);
                                                }
                                            }

                                            // 找不到就用 TAB + ENTER 兜底
                                            SendKeysSafe("{TAB}");
                                            SendKeysSafe("{TAB}");
                                            SendKeysSafe("{TAB}");
                                            SendKeysSafe("{TAB}");
                                            SendKeysSafe("{TAB}");
                                            SendKeysSafe("{ENTER}");
                                            Console.WriteLine("未找到确认控件，已用TAB+ENTER确认。");


                                        }
                                        // 处理完所有批次后，立刻清空状态并 return，防止重复
                                        Console.WriteLine($"[调试] 复盘批量转发处理完毕，清空状态");
                                        lastReplayCustomers.Clear();
                                        lastReplayText = "";
                                        pendingReplayInput = "";
                                        return;
                                    }
                                    finally
                                    {
                                        Interlocked.Exchange(ref _isForwarding, 0);
                                    }
                                }
                                else if (trimmed == "1" || trimmed == "2" || trimmed == "3" ||
                                         trimmed == "4" || trimmed == "5" || trimmed == "6" ||
                                         trimmed == "7" || trimmed == "8")
                                {
                                    Console.WriteLine($"[{windowName}] 用户选择了：{trimmed}，将发送报盘到米助待定9");
                                    waitingForChoice = false;
                                    // 只有此时才发送
                                    if (!string.IsNullOrEmpty(lastBaopanText))
                                    {
                                        try
                                        {
                                            if (Process.GetProcessesByName("weixin").Length == 0)
                                            {
                                                Console.WriteLine("微信已退出，自动化流程终止。");
                                                return;
                                            }
                                            switch (trimmed)
                                            {
                                                case "1":
                                                    recallDetected = false; // 进入流程前先清空
                                                    if (!SendToWeixinMainWindow(chatWindow, "壹★牛", lastBaopanText)) return;
                                                    Thread.Sleep(300);

                                                    if (!SendToWeixinMainWindow(chatWindow, "贰★牛", lastBaopanText)) return;
                                                    Thread.Sleep(300);

                                                    if (!SendToWeixinMainWindow(chatWindow, "叁★牛", lastBaopanText)) return;
                                                    Thread.Sleep(300);
                                                    if (!SendToWeixinMainWindow(chatWindow, "肆★牛", lastBaopanText)) return;
                                                    Thread.Sleep(300);
                                                    if (!SendToWeixinMainWindow(chatWindow, "伍★牛", lastBaopanText)) return;
                                                    Thread.Sleep(300);
                                                    if (!SendToWeixinMainWindow(chatWindow, "陆★牛", lastBaopanText)) return;
                                                    Thread.Sleep(300);
                                                    if (!SendToWeixinMainWindow(chatWindow, "柒★牛", lastBaopanText)) return;
                                                    Thread.Sleep(300);
                                                    if (!SendToWeixinMainWindow(chatWindow, "捌★牛 ", lastBaopanText)) return;
                                                    Thread.Sleep(300);
                                                    if (!SendToWeixinMainWindow(chatWindow, "玖★牛 ", lastBaopanText)) return;
                                                    Thread.Sleep(300);
                                                    break;
                                                case "2":
                                                    recallDetected = false; // 进入流程前先清空
                                                    if (!SendToWeixinMainWindow(chatWindow, "壹★骨", lastBaopanText)) return;
                                                    Thread.Sleep(300);

                                                    if (!SendToWeixinMainWindow(chatWindow, "贰★骨", lastBaopanText)) return;
                                                    Thread.Sleep(300);

                                                    if (!SendToWeixinMainWindow(chatWindow, "叁★骨", lastBaopanText)) return;
                                                    Thread.Sleep(300);

                                                    if (!SendToWeixinMainWindow(chatWindow, "肆★骨", lastBaopanText)) return;
                                                    Thread.Sleep(300);

                                                    if (!SendToWeixinMainWindow(chatWindow, "伍★骨", lastBaopanText)) return;
                                                    Thread.Sleep(300);

                                                    break;
                                                case "3":
                                                    recallDetected = false; // 进入流程前先清空
                                                    if (!SendToWeixinMainWindow(chatWindow, "壹★羊", lastBaopanText)) return;
                                                    Thread.Sleep(300);

                                                    //if (!SendToWeixinMainWindow(chatWindow, "贰★羊", lastBaopanText)) return;
                                                    // Thread.Sleep(300);

                                                    break;
                                                case "4":
                                                    recallDetected = false; // 进入流程前先清空
                                                    if (!SendToWeixinMainWindow(chatWindow, "壹★猪", lastBaopanText)) return;
                                                    Thread.Sleep(300);

                                                    if (!SendToWeixinMainWindow(chatWindow, "贰★猪", lastBaopanText)) return;
                                                    Thread.Sleep(300);

                                                    if (!SendToWeixinMainWindow(chatWindow, "叁★猪", lastBaopanText)) return;
                                                    Thread.Sleep(300);

                                                    break;
                                                case "5":
                                                    recallDetected = false; // 进入流程前先清空
                                                    if (!SendToWeixinMainWindow(chatWindow, "壹★新", lastBaopanText)) return;
                                                    Thread.Sleep(300);

                                                    // if (!SendToWeixinMainWindow(chatWindow, "贰★新", lastBaopanText)) return;
                                                    // Thread.Sleep(300);

                                                    break;
                                                case "6":
                                                    recallDetected = false; // 进入流程前先清空
                                                    if (!SendToWeixinMainWindow(chatWindow, "壹★碎", lastBaopanText)) return;
                                                    Thread.Sleep(300);

                                                    if (!SendToWeixinMainWindow(chatWindow, "贰★碎", lastBaopanText)) return;
                                                    Thread.Sleep(300);

                                                    break;
                                                case "7":
                                                    recallDetected = false; // 进入流程前先清空
                                                    if (!SendToWeixinMainWindow(chatWindow, "壹★牛", lastBaopanText)) return;
                                                    Thread.Sleep(300);

                                                    if (!SendToWeixinMainWindow(chatWindow, "贰★牛", lastBaopanText)) return;
                                                    Thread.Sleep(300);

                                                    if (!SendToWeixinMainWindow(chatWindow, "叁★牛", lastBaopanText)) return;
                                                    Thread.Sleep(300);
                                                    if (!SendToWeixinMainWindow(chatWindow, "肆★牛", lastBaopanText)) return;
                                                    Thread.Sleep(300);
                                                    if (!SendToWeixinMainWindow(chatWindow, "伍★牛", lastBaopanText)) return;
                                                    Thread.Sleep(300);
                                                    if (!SendToWeixinMainWindow(chatWindow, "陆★牛", lastBaopanText)) return;
                                                    Thread.Sleep(300);
                                                    if (!SendToWeixinMainWindow(chatWindow, "柒★牛", lastBaopanText)) return;
                                                    Thread.Sleep(300);
                                                    if (!SendToWeixinMainWindow(chatWindow, "捌★牛 ", lastBaopanText)) return;
                                                    Thread.Sleep(300);
                                                    if (!SendToWeixinMainWindow(chatWindow, "玖★牛 ", lastBaopanText)) return;
                                                    Thread.Sleep(300);
                                                    if (!SendToWeixinMainWindow(chatWindow, "壹☆牛", lastBaopanText)) return;
                                                    Thread.Sleep(300);

                                                    if (!SendToWeixinMainWindow(chatWindow, "贰☆牛", lastBaopanText)) return;
                                                    Thread.Sleep(300);
                                                    // if (!SendToWeixinMainWindow(chatWindow, "叁☆牛", lastBaopanText)) return;
                                                    // Thread.Sleep(300);
                                                    break; ;
                                                case "8":
                                                    recallDetected = false; // 进入流程前先清空
                                                    if (!SendToWeixinMainWindow(chatWindow, "壹★圆", lastBaopanText)) return;
                                                    Thread.Sleep(300);

                                                    //if (!SendToWeixinMainWindow(chatWindow, "贰★圆", lastBaopanText)) return;
                                                    // Thread.Sleep(300);

                                                    break;
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"SendToWeixinMainWindow异常: {ex}");
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("没有可发送的报盘内容，已跳过。");
                                    }
                                }
                                else if (trimmed == "撤回" || trimmed == "撤回信息")
                                {
                                    Console.WriteLine($"[{windowName}] 用户选择了：撤回，后续动作待实现");
                                    waitingForChoice = false;
                                    lastBaopanText = "";
                                    EnsurePopupClosed(chatWindow);
                                    HandleRecall(chatWindow);
                                }
                            }
                            if (windowName.Trim() == "111周")
                            {
                                if (text.Contains("再见"))
                                {
                                    SendReplyInWindow(chatWindow, "你好");
                                    Console.WriteLine($"[{windowName}] 自动回复：你好");
                                }
                            }
                            // 超时检测（防止用户一直不发消息）
                            if (waitingForChoice && (DateTime.Now - choiceStartTime).TotalSeconds > 25)
                            {
                                Console.WriteLine($"[{windowName}] 25秒未选择，已自动取消本次操作");
                                SendReplyInWindow(chatWindow, "25秒未选择，已自动取消本次操作");
                                waitingForChoice = false;
                                lastBaopanText = "";
                            }
                        }

                        if (Regex.IsMatch(text.Trim(), @"^([\u4e00-\u9fa5]+)?(\d+)厂([\u4e00-\u9fa5A-Za-z0-9]+)$"))
                        {
                            string uniqueKey = text.Trim();

                            // 只保留最近N条，防止HashSet无限增长
                            if (processedProductMessages.Count > 100)
                            {
                                processedProductMessages.Clear();
                            }

                            if (processedProductMessages.Contains(uniqueKey))
                            {
                                // 已处理过，跳过
                                return;
                            }
                            processedProductMessages.Add(uniqueKey);

                            if (processedProductMessages.Count > 100)
                            {
                                processedProductMessages.Clear();
                            }

                            var match = Regex.Match(text.Trim(), @"^([\u4e00-\u9fa5]+)?(\d+)厂([\u4e00-\u9fa5A-Za-z0-9]+)$");
                            if (match.Success)
                            {
                                string country = match.Groups[1].Success ? match.Groups[1].Value.Trim() : "";
                                string factory = match.Groups[2].Value.Trim();
                                string product = match.Groups[3].Value.Trim();
                                Console.WriteLine($"匹配到国家: {country}, 厂号: {factory}, 产品: {product}");
                                string baseDir = @"Z:\\产品图片";
                                string countryDir = string.IsNullOrEmpty(country) ? baseDir : System.IO.Path.Combine(baseDir, country);
                                string factoryDir = System.IO.Path.Combine(countryDir, factory);
                                Console.WriteLine($"国家目录: {countryDir}");
                                Console.WriteLine($"厂号目录: {factoryDir}");
                                if (!System.IO.Directory.Exists(countryDir))
                                {
                                    SendReplyInWindow(chatWindow, $"未找到国家目录：{countryDir}");
                                    return;
                                }
                                if (!System.IO.Directory.Exists(factoryDir))
                                {
                                    SendReplyInWindow(chatWindow, $"没有存这个厂号的图片：{factory}");
                                    return;
                                }
                                var matchingDirs = System.IO.Directory.GetDirectories(factoryDir)
                                    .Where(d => d.IndexOf(product, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
                                Console.WriteLine($"匹配到的产品目录数量: {matchingDirs.Count}");
                                if (matchingDirs.Count == 0)
                                {
                                    // 获取所有产品目录名
                                    var allProductDirs = System.IO.Directory.GetDirectories(factoryDir);
                                    if (allProductDirs.Length == 0)
                                    {
                                        SendReplyInWindow(chatWindow, $"该厂号下没有任何产品目录。");
                                        return;
                                    }
                                    // 提取产品名（文件夹名）
                                    var allProductNames = allProductDirs.Select(d => System.IO.Path.GetFileName(d)).ToList();
                                    string allProductsText = "没存，该厂号有存的产品图片为：\n" + string.Join("\n", allProductNames);
                                    SendReplyInWindow(chatWindow, allProductsText);
                                    return;
                                }
                                foreach (var dir in matchingDirs)
                                {
                                    var files = System.IO.Directory.GetFiles(dir).ToList();
                                    Console.WriteLine($"目录 {dir} 下文件数量: {files.Count}");
                                    if (files.Count == 0)
                                    {
                                        SendReplyInWindow(chatWindow, $"未找到匹配的产品文件：{product}");
                                        continue;
                                    }
                                    SendFilesInBatches(chatWindow, files);
                                    SendReplyInWindow(chatWindow, $"已发送：{System.IO.Path.GetFileName(dir)}");
                                }
                                return;
                            }
                            else
                            {
                                Console.WriteLine("未能正确匹配国家、厂号、产品，请检查输入格式。");
                            }
                        }
                    }
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

    static string TryExtractMessageText(AutomationElement messageElement, int depth = 0)
    {
        if (messageElement == null || depth > 10) return null; // 限制最大递归深度

        // 优先递归拼接所有Text类型控件内容
        if (messageElement.ControlType == FlaUI.Core.Definitions.ControlType.Text)
        {
            var text = messageElement.Name;
            if (!string.IsNullOrEmpty(text)) return text;
        }

        // ValuePattern
        if (messageElement.Patterns.Value.IsSupported)
        {
            var text = messageElement.Patterns.Value.Pattern.Value;
            if (!string.IsNullOrEmpty(text)) return text;
        }

        // TextPattern
        if (messageElement.Patterns.Text.IsSupported)
        {
            var text = messageElement.Patterns.Text.Pattern.DocumentRange.GetText(-1)?.Trim();
            if (!string.IsNullOrEmpty(text)) return text;
        }

        // 递归所有子控件
        var children = messageElement.FindAllChildren();
        StringBuilder sb = new StringBuilder();
        foreach (var child in children)
        {
            var childText = TryExtractMessageText(child, depth + 1);
            if (!string.IsNullOrEmpty(childText))
            {
                sb.Append(childText);
            }
        }
        if (sb.Length > 0) return sb.ToString();

        // 最后再尝试Name属性
        var name = messageElement.Name;
        if (!string.IsNullOrEmpty(name)) return name;

        return null;
    }

    static void PrintAllControls(AutomationElement element, int indent = 0)
    {
        var children = element.FindAllChildren();
        foreach (var child in children)
        {
            Console.WriteLine($"{new string(' ', indent * 2)} Name[{child.Name}], ClassName {child.ClassName}, ControlType {child.ControlType}");
            PrintAllControls(child, indent + 1);
        }
    }

    static void SendReplyInWindow(AutomationElement chatWindow, string replyText)
    {
        IntPtr oldForeground = GetForegroundWindow();
        var hwnd = (IntPtr)chatWindow.Properties.NativeWindowHandle.Value;
        SetForegroundWindow(hwnd);
        Thread.Sleep(200);

        AutomationElement inputBox = null;
        for (int i = 0; i < 5; i++)
        {
            inputBox = chatWindow.FindFirstDescendant(cf => cf.ByControlType(ControlType.Edit));
            if (inputBox != null) break;
            Thread.Sleep(200);
        }
        if (inputBox == null)
        {
            Console.WriteLine("未找到输入框，无法自动回复。");
            if (oldForeground != IntPtr.Zero) SetForegroundWindow(oldForeground);
            return;
        }

        try
        {
            if (inputBox.Patterns.Value.IsSupported)
            {
                inputBox.Patterns.Value.Pattern.SetValue(replyText);
            }
            else
            {
                // 兜底方案：用SendKeys粘贴
                inputBox.Focus();
                Thread.Sleep(200);
                SendKeysSafe("^v");
            }

            // 查找"发送"按钮
            var sendButton = chatWindow.FindFirstDescendant(cf =>
                cf.ByControlType(FlaUI.Core.Definitions.ControlType.Button)
                .And(cf.ByClassName("mmui::XOutlineButton"))
                .And(cf.ByName("发送"))
            );

            if (sendButton != null)
            {
                sendButton.AsButton().Invoke();
                Console.WriteLine("已后台点击发送按钮。");
            }
            else
            {
                // 兜底方案
                SendKeysSafe("^{ENTER}");
                Console.WriteLine("未找到发送按钮，已用快捷键发送。");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("后台发送失败：" + ex.Message);
        }

        if (oldForeground != IntPtr.Zero)
        {
            Thread.Sleep(100);
            SetForegroundWindow(oldForeground);
        }
    }

    static bool SendToWeixinMainWindow(AutomationElement chatWindow, string inputText, string baopanText)
    {
        // 首先激活古米特米助群的窗口
        SetForegroundWindow(FindWindow(null, "古米特米助群"));
        // 检查微信是否已退出
        if (Process.GetProcessesByName("weixin").Length == 0)
        {
            Console.WriteLine("微信已退出，自动化流程终止。");
            return false;
        }

        // 1. 找到消息列表List控件
        var list = FindByPath(chatWindow, "/Group/Group/Group/Custom/Group[2]/List");
        if (list == null)
        {
            Console.WriteLine("SendToWeixinMainWindow: 未找到List控件，跳过本次操作。");
            return false;
        }
        var items = list.FindAllChildren();
        AutomationElement baopanItem = null;
        foreach (var item in items)
        {
            // 让每个item滚动到可见
            if (item.Patterns.ScrollItem.IsSupported)
            {
                item.Patterns.ScrollItem.Pattern.ScrollIntoView();
                Thread.Sleep(500); // 等待滚动动画完全结束
            }

            var msgText = TryExtractMessageText(item);
            var rect = item.BoundingRectangle;
            bool isVisible = !item.Properties.IsOffscreen.Value && rect.Width > 0 && rect.Height > 0;

            if (!string.IsNullOrEmpty(msgText) && (msgText.Contains("报盘") || msgText.Contains("复盘")) && isVisible)
            {
                baopanItem = item;
                break;
            }
        }

        if (baopanItem != null)
        {
            // 1. 滚动到可见
            if (baopanItem.Patterns.ScrollItem.IsSupported)
            {
                baopanItem.Patterns.ScrollItem.Pattern.ScrollIntoView();
                Thread.Sleep(500); // 等待滚动动画完全结束
            }

            // 2. 获取BoundingRectangle
            var boundingRect = baopanItem.BoundingRectangle;
            int x = (int)(boundingRect.Left - 2 + 100); // 右下角往左偏移100像素
            int y = (int)(boundingRect.Bottom - 2 - 50); // 右下角往上偏移50像素

            // 3. 激活窗口
            IntPtr hwnd = (IntPtr)chatWindow.Properties.NativeWindowHandle.Value;
            SetForegroundWindow(hwnd);
            Thread.Sleep(200); // 极短等待，保证窗口激活

            // 4. 鼠标移动并右键
            Cursor.Position = new System.Drawing.Point(x, y);
            Thread.Sleep(100);
            mouse_event(0x08, 0, 0, 0, 0); // 右键按下
            mouse_event(0x10, 0, 0, 0, 0); // 右键抬起
            Thread.Sleep(200);

            // 5. 按5次下方向键
            for (int i = 0; i < 5; i++)
            {
                SendKeysSafe("{DOWN}");
                Thread.Sleep(10);
            }
            // 6. 回车
            SendKeysSafe("{ENTER}");
            Thread.Sleep(300); // 等待UI反应

            // 7. 判断报盘内容是否一致
            if (HasExactBaopanText(chatWindow, baopanText))
            {
                // 一致，继续后续操作
                SelectAllCheckBoxAndSend(chatWindow, inputText, baopanText);
                return true;
            }
            else
            {
                // 不一致，直接按ESC键退出弹窗
                SendKeysSafe("{TAB}");
                SendKeysSafe("{TAB}");
                SendKeysSafe("{TAB}");
                SendKeysSafe("{TAB}");
                SendKeysSafe("{TAB}");
                SendKeysSafe("{ENTER}");
                Console.WriteLine("报盘内容不一致，已按下ESC键关闭弹窗。");
                return false;
            }
        }
        else
        {
            Console.WriteLine("未找到可见且包含\"报盘\"的ListItem。");
        }

        if (recallDetected) { HandleRecall(chatWindow); return false; }
        return true;
    }

    static void FindAllXTableCellListItems(AutomationElement root, List<AutomationElement> result)
    {
        if (root.ClassName == "mmui::XTableCell" && root.ControlType == FlaUI.Core.Definitions.ControlType.ListItem)
        {
            result.Add(root);
        }
        var children = root.FindAllChildren();
        foreach (var child in children)
        {
            FindAllXTableCellListItems(child, result);
        }
    }

    static void FindAllSearchContactCellCheckBoxes(AutomationElement root, List<AutomationElement> result)
    {
        if (root.ClassName == "mmui::SearchContactCellView" && root.ControlType == FlaUI.Core.Definitions.ControlType.CheckBox)
        {
            result.Add(root);
        }
        var children = root.FindAllChildren();
        foreach (var child in children)
        {
            FindAllSearchContactCellCheckBoxes(child, result);
        }
    }

    /// <summary>
    /// 勾选所有 CheckBox，按下方向键和回车，然后输入指定内容
    /// </summary>
    /// <param name="chatWindow">聊天窗口 AutomationElement</param>
    /// <param name="inputText">要输入的内容</param>
    static void SelectAllCheckBoxAndSend(AutomationElement chatWindow, string inputText, string baopanText)
    {
        lock (uiaLock)
        {

            if (!HasExactBaopanText(chatWindow, baopanText))
            {
                Console.WriteLine("未检测到内容与报盘完全一致的 XTextView Text 控件，跳过后续操作。");
                return;
            }

            // 1. 先在搜索框输入内容
            var searchBox = chatWindow.FindFirstDescendant(cf =>
                cf.ByControlType(FlaUI.Core.Definitions.ControlType.Edit)
                .And(cf.ByClassName("mmui::XValidatorTextEdit"))
                .And(cf.ByName("搜索"))
            );

            if (searchBox != null)
            {
                searchBox.Focus();
                Thread.Sleep(200);

                if (searchBox.Patterns.Value.IsSupported)
                {
                    searchBox.Patterns.Value.Pattern.SetValue(inputText);
                    Console.WriteLine($"已在搜索框输入内容：{inputText}");
                }
                else
                {
                    SendTextSafe(inputText);
                    Console.WriteLine($"已用SendKeys在搜索框输入内容：{inputText}");
                }
                Thread.Sleep(200); // 等待输入完成

                SendKeysSafe("{ENTER}"); // 回车
                Console.WriteLine("已在搜索框回车。");
                Thread.Sleep(600); // 等待UI刷新
            }
            else
            {
                Console.WriteLine("未找到搜索输入框，无法输入内容。");
            }

            // 2. 查找并点击第二个 mmui::XTableCell ListItem
            var allTargets = new List<AutomationElement>();
            FindAllXTableCellListItems(chatWindow, allTargets);

            if (allTargets.Count >= 2)
            {
                allTargets[1].Click();
                Console.WriteLine("已点击第2个 ClassName=mmui::XTableCell, ControlType=ListItem 的控件。");
            }
            else
            {
                Console.WriteLine($"未找到足够的 ClassName=mmui::XTableCell, ControlType=ListItem 控件（共找到 {allTargets.Count} 个）。");
            }

            // 3. 再查找并点击所有 CheckBox
            AutomationElement[] allCheckBoxes = null;
            for (int i = 0; i < 10; i++)
            {
                allCheckBoxes = chatWindow.FindAllDescendants(cf =>
                    cf.ByControlType(FlaUI.Core.Definitions.ControlType.CheckBox)
                    .And(cf.ByClassName("mmui::SearchContactCellView"))
                );
                if (allCheckBoxes.Length > 0)
                    break;
                Thread.Sleep(300);

            }
            foreach (var cb in allCheckBoxes)
            {
                Console.WriteLine($"CheckBox: Name={cb.Name}, ClassName={cb.ClassName}");
            }
            if (allCheckBoxes.Length > 0)
            {
                foreach (var cb in allCheckBoxes)
                {

                    var rect = cb.BoundingRectangle;
                    int clickX = (int)(rect.Left + rect.Width / 2);
                    int clickY = (int)(rect.Top + rect.Height / 2);
                    Cursor.Position = new System.Drawing.Point(clickX, clickY);
                    mouse_event(0x02, 0, 0, 0, 0); // 左键按下
                    mouse_event(0x04, 0, 0, 0, 0); // 左键抬起
                    Console.WriteLine($"已点击 CheckBox，位置：({clickX},{clickY})");
                    Thread.Sleep(10);
                }

                // 点击确认
                Thread.Sleep(300);
                SendKeysSafe("{TAB}");
                SendKeysSafe("{TAB}");
                SendKeysSafe("{TAB}");
                SendKeysSafe("{TAB}");
                SendKeysSafe("{ENTER}");
                Console.WriteLine("已按下确认键关闭弹窗。");
                if (CheckRecallAndReturn(chatWindow)) return;
            }
            else
            {
                Console.WriteLine("未找到任何 SearchContactCellView 类型的 CheckBox。");
                PrintAllControls(chatWindow);
            }
            if (CheckRecallAndReturn(chatWindow)) return;
            // 查找"发送"按钮
            var sendButton = chatWindow.FindFirstDescendant(cf =>
                cf.ByControlType(FlaUI.Core.Definitions.ControlType.Button)
                .And(cf.ByClassName("mmui::XOutlineButton"))
                .And(cf.ByName("发送"))
            );
            if (sendButton != null)
            {
                sendButton.AsButton().Invoke();
                Console.WriteLine("已后台点击发送按钮。");
            }
            else
            {
                SendKeysSafe("^{ENTER}");
                Console.WriteLine("未找到发送按钮，已用快捷键发送。");
            }
            if (CheckRecallAndReturn(chatWindow)) return;
        }
    }

    static void DoRecallAndReturn(AutomationElement chatWindow)
    {
        Thread.Sleep(1500);
        if (Process.GetProcessesByName("weixin").Length > 0)
        {
            EnsurePopupClosed(chatWindow);
            Thread.Sleep(500); // 等待UI稳定
            HandleRecall(chatWindow);
        }
        recallDetected = false;
    }

    /// <summary>
    /// 判断窗口中是否存在内容与指定内容完全一致的 mmui::XTextView Text 控件
    /// </summary>
    static bool HasExactBaopanText(AutomationElement root, string baopanText)
    {
        var allTextViews = new List<AutomationElement>();
        FindAllXTextViewTextControls(root, allTextViews);
        foreach (var tv in allTextViews)
        {
            string text = tv.Name?.Trim();
            if (!string.IsNullOrEmpty(text) && text == baopanText.Trim())
            {
                return true;
            }
        }
        return false;
    }

    /// <summary>
    /// 递归查找所有 ClassName=mmui::XTextView, ControlType=Text 的控件
    /// </summary>
    static void FindAllXTextViewTextControls(AutomationElement root, List<AutomationElement> result)
    {
        // 新增：递归前检测撤回
        if (recallDetected)
        {
            Console.WriteLine("FindAllXTextViewTextControls: 检测到撤回，递归中断。");
            return;
        }
        try
        {
            string className = "";
            try
            {
                className = root.ClassName;
            }
            catch (Exception ex)
            {
                Console.WriteLine("FindAllXTextViewTextControls: 获取ClassName异常，已跳过该节点：" + ex.Message);
                className = "";
            }

            if (className == "mmui::XTextView" && root.ControlType == FlaUI.Core.Definitions.ControlType.Text)
            {
                result.Add(root);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("FindAllXTextViewTextControls: 节点异常，已跳过：" + ex.Message);
        }

        AutomationElement[] children = null;
        try
        {
            children = root.FindAllChildren();
        }
        catch (Exception ex)
        {
            Console.WriteLine("FindAllXTextViewTextControls: 获取子节点异常，已跳过：" + ex.Message);
            return;
        }
        if (children != null)
        {
            foreach (var child in children)
            {
                // 新增：递归前检测撤回
                if (recallDetected)
                {
                    Console.WriteLine("FindAllXTextViewTextControls: 检测到撤回，递归中断。");
                    return;
                }
                FindAllXTextViewTextControls(child, result);
            }
        }
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

    static string Normalize(string s)
    {
        return string.IsNullOrEmpty(s) ? "" : Regex.Replace(s.Trim(), @"\s+", "");
    }

    static void RevokeLastMessagesBatch(List<string> windowNames)
    {
        foreach (var winName in windowNames)
        {
            Console.WriteLine($"开始撤回窗口[{winName}]的最后一条消息...");
            bool success = RevokeLastMessageForWindow(winName);
            if (success)
            {
                Console.WriteLine($"[{winName}] 撤回成功！");
            }
            else
            {
                Console.WriteLine($"[{winName}] 撤回失败或未找到窗口！");
            }
            Thread.Sleep(1500); // 等待1.5秒，防止UI卡死
        }
    }

    static bool RevokeLastMessageForWindow(string windowName)
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
                        // 下面的逻辑和你现有的 RevokeLastMessage 一样
                        // 1. 激活窗口
                        // 2. 查找消息列表
                        // 3. 找到最后一条消息，右键撤回
                        // 4. 检查"删除该条消息？"弹窗，自动ESC
                        // ...（可直接复用你现有的撤回流程代码）
                        // 成功返回 true
                        return true;
                    }
                }
            }
        }
        return false;
    }

    static void SendFilesInBatches(AutomationElement chatWindow, List<string> filePaths)
    {
        try
        {
            int total = filePaths.Count;
            int batchCount = (int)Math.Ceiling(total / 9.0);

            for (int batch = 0; batch < batchCount; batch++)
            {
                var filesThisBatch = filePaths.Skip(batch * 9).Take(9).ToList();
                Console.WriteLine($"准备发送第{batch + 1}批，共{filesThisBatch.Count}个文件：");
                foreach (var f in filesThisBatch)
                {
                    Console.WriteLine($"  {f}");
                }

                // 1. 清空剪贴板（只清一次，不重试）
                Console.WriteLine("清空剪贴板...");
                var clearThread = new Thread(() =>
                {
                    try
                    {
                        Clipboard.Clear();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("清空剪贴板异常：" + ex);
                    }
                });
                clearThread.SetApartmentState(ApartmentState.STA);
                clearThread.Start();
                clearThread.Join(500); // 最多等0.5秒
                // 不管成不成功，直接往下走

                // 2. 复制本批文件到剪贴板
                Console.WriteLine("复制文件到剪贴板...");
                var thread = new Thread(() =>
                {
                    try
                    {
                        var fileCollection = new System.Collections.Specialized.StringCollection();
                        foreach (var file in filesThisBatch)
                        {
                            fileCollection.Add(file);
                        }
                        Clipboard.SetFileDropList(fileCollection);
                        Console.WriteLine("已写入剪贴板。");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("写入剪贴板异常：" + ex);
                    }
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join(500); // 最多等0.5秒
                // 不管成不成功，直接往下走

                // 3. 激活窗口
                IntPtr hwnd = (IntPtr)chatWindow.Properties.NativeWindowHandle.Value;
                SetForegroundWindow(hwnd);
                Thread.Sleep(200); // 极短等待，保证窗口激活

                // 4. 找到输入框
                var inputBox = chatWindow.FindFirstDescendant(cf => cf.ByControlType(ControlType.Edit));
                if (inputBox == null)
                {
                    Console.WriteLine("未找到输入框，无法发送文件。");
                    continue;
                }

                // 5. 模拟鼠标点击输入框，确保焦点
                var rect = inputBox.BoundingRectangle;
                System.Windows.Forms.Cursor.Position = new System.Drawing.Point((int)(rect.Left + rect.Width / 2), (int)(rect.Top + rect.Height / 2));
                mouse_event(0x02, 0, 0, 0, 0); // 左键按下
                mouse_event(0x04, 0, 0, 0, 0); // 左键抬起
                Thread.Sleep(100); // 极短等待

                inputBox.Focus();
                Thread.Sleep(100); // 极短等待

                // 6. 粘贴
                Console.WriteLine("准备粘贴...");
                SendKeysSafe("^v");
                Thread.Sleep(1000);
                Console.WriteLine($"已粘贴第{batch + 1}批文件到输入框，准备发送...");

                // 7. 发送
                Console.WriteLine("准备发送...");
                SendKeysSafe("^{ENTER}");
                Console.WriteLine($"已发送第{batch + 1}批。");

                // 8. 清空输入框（Ctrl+A, Delete）
                SendKeysSafe("^a");
                Thread.Sleep(100);
                SendKeysSafe("{DEL}");
                Thread.Sleep(100);

                // 9. 直接进入下一批，不等待
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("SendFilesInBatches异常：" + ex);
        }
    }

    static Dictionary<string, List<string>> GroupFilesByType(List<string> files)
    {
        var groups = new Dictionary<string, List<string>>();
        foreach (var file in files)
        {
            string ext = Path.GetExtension(file).ToLower();
            string type = "other";
            if (new[] { ".jpg", ".jpeg", ".png", ".bmp", ".gif" }.Contains(ext))
                type = "image";
            else if (new[] { ".mp4", ".avi", ".mov", ".wmv" }.Contains(ext))
                type = "video";
            else if (new[] { ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".txt" }.Contains(ext))
                type = "doc";
            if (!groups.ContainsKey(type))
                groups[type] = new List<string>();
            groups[type].Add(file);
        }
        return groups;
    }

    static void AutoItDragDropDemo()
    {
        var autoit = new AutoItX3();
        autoit.WinActivate(GROUP_NAME);
        autoit.WinWaitActive(GROUP_NAME);
        // 下面的坐标需要你用AutoIt Window Info工具测量
        autoit.MouseClickDrag("left", 200, 200, 800, 600, 1);
        autoit.Sleep(2000);
        autoit.Send("^({ENTER})");
    }

    static void AutoItDragDropBatch(List<string> files, string wechatWindowTitle)
    {
        var autoit = new AutoItX3();
        autoit.WinActivate(wechatWindowTitle);
        autoit.WinWaitActive(wechatWindowTitle);

        foreach (var file in files)
        {
            // 1. 打开资源管理器并选中文件（可用Shell命令或人工辅助）
            // 2. 拖拽到微信输入框
            // 下面的坐标需用AutoIt Window Info工具测量
            autoit.MouseClickDrag("left", 200, 200, 800, 600, 1);
            autoit.Sleep(2000);
            autoit.Send("^({ENTER})");
            autoit.Sleep(3000);
        }
    }

    // 1. 船公司查柜网址映射
    static Dictionary<string, string> companyUrls = new Dictionary<string, string>
    {
        { "COSCO", "https://elines.coscoshipping.com/ebusiness/cargoTracking?" },
        { "MSC", "https://www.msccargo.cn/en/track-a-shipment?containerNo={0}" },
        { "CMA", "https://www.cma-cgm.com/" },
        { "ONE", "https://ecomm.one-line.com/one-ecom/manage-shipment/cargo-tracking?ctrack-field={0}&trakNoParam={0}" },
        { "HMM", "https://www.hmm21.com/e-service/general/trackNTrace/TrackNTrace.do?trackingNo={0}" },
        { "YML", "https://e-solution.yangming.com/e-service/track_trace/track_trace_cargo_tracking.aspx?containerNo={0}" },
        { "HPL", "https://www.hapag-lloyd.com/en/online-business/track/track-by-container.html?container={0}" },
        //{ "OOCL", "https://iac.oocl.com/chatui/#/app/online?containerNo={0}" },
        { "ZIM", "https://www.zim.com/tools/track-a-shipment?consnumber={0}" },
        { "PIL", "https://www.pilship.com/digital-solutions/?tab=customer&id=track-trace&label=containerTandT&module=TrackContStatus&refNo={0}" },
        { "EMC", "https://www.evergreen-shipping.cn/servlet/TDB1_CargoTracking.do?containerNo={0}" },
        { "SITC", "https://cargo.sitc.com/track/trackDetail?containerNo={0}" },
        { "MAERSK", "https://www.maersk.com.cn/tracking/{0}" },
        { "MOL", "https://www.molpower.com/ebusiness/trackTrace.do?containerNo={0}" },
        { "HYUNDAI PARAMOUNT", "https://www.hmm21.com/e-service/general/trackNTrace/TrackNTrace.do?trackingNo={0}" }
        // 其他公司可继续补充
    };

    // 3. 自动打开高概率船公司查柜网站
    public static void OpenTopCompanySites(string code, int topN = 4)
    {
        Console.WriteLine($"[调试] OpenTopCompanySites 被调用，code={code}, topN={topN}, 时间={DateTime.Now:HH:mm:ss.fff}");
        if (!Regex.IsMatch(code, @"^[A-Z]{4}\d{7}$"))
            return;

        string prefix = code.Substring(0, 4);
        if (resultDict.ContainsKey(prefix))
        {
            var row = resultDict[prefix];
            Console.WriteLine("row内容（含ASCII码和长度）：");
            foreach (var kv in row)
            {
                var key = kv.Key;
                Console.WriteLine($"{key} | 长度:{key.Length} | ASCII:{string.Join(",", key.Select(c => (int)c))}");
            }
            var topCompanies = new List<KeyValuePair<string, int>>(row);
            topCompanies.Sort((a, b) => b.Value.CompareTo(a.Value));
            int opened = 0;
            Console.WriteLine($"【调试】{code} 前4高概率公司：");
            var topCompaniesWithUrl = row
                .Where(kv => kv.Value > 0 && companyUrls.ContainsKey(kv.Key.Trim().ToUpper()))
                .OrderByDescending(kv => kv.Value)
                .Take(4)
                .ToList();

            // 把最大概率的公司移到最后
            if (topCompaniesWithUrl.Count > 1)
            {
                var maxItem = topCompaniesWithUrl[0];
                topCompaniesWithUrl.RemoveAt(0);
                topCompaniesWithUrl.Add(maxItem);
            }

            Console.WriteLine("topCompaniesWithUrl内容：");
            foreach (var kv in topCompaniesWithUrl)
            {
                Console.WriteLine($"{kv.Key} | 概率:{kv.Value} | 有网址:{companyUrls.ContainsKey(kv.Key.Trim().ToUpper())}");
            }

            Console.WriteLine($"==== 开始打开 {code} 的所有公司 ====");
            try
            {
                foreach (var kv in topCompaniesWithUrl)
                {
                    try
                    {
                        if (Process.GetProcessesByName("weixin").Length == 0)
                        {
                            Console.WriteLine("微信已退出，自动化流程终止。");
                            return;
                        }
                        string company = kv.Key.Trim().ToUpper();
                        Console.WriteLine($"公司: {company}, 概率: {kv.Value}, 有网址: {companyUrls.ContainsKey(company)}");
                        string url = string.Format(companyUrls[company], code);
                        Console.WriteLine($"准备打开：{company} {url}");

                        // 1. 批量打开所有高概率公司网站
                        ProcessStartInfo psi = new ProcessStartInfo("cmd", $"/c start \"\" \"{url}\"");
                        psi.CreateNoWindow = true;
                        psi.UseShellExecute = false;
                        psi.WindowStyle = ProcessWindowStyle.Hidden;
                        Process.Start(psi);
                        Thread.Sleep(800); // 避免浏览器卡死
                        opened++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"打开 {kv.Key} 时异常: {ex.Message}");
                    }
                }
                Console.WriteLine("所有高概率公司网站已全部打开。");

                // 2. 等待Automa脚本截图（比如等待15秒）
                string imgPath = @"C:\\Users\\x\\Downloads\\cqjt.png";
                int maxWaitSeconds = 50;
                bool found = false;
                for (int i = 0; i < maxWaitSeconds; i++)
                {
                    if (System.IO.File.Exists(imgPath))
                    {
                        found = true;
                        break;
                    }
                    Thread.Sleep(1000); // 每秒检测一次
                }
                if (found)
                {
                    Console.WriteLine("检测到图片，准备关闭浏览器...");
                    // 先关闭所有Edge浏览器
                    try
                    {
                        foreach (var proc in Process.GetProcessesByName("msedge"))
                        {
                            proc.Kill();
                        }
                        Console.WriteLine("已关闭所有Edge浏览器进程。");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("关闭Edge浏览器失败：" + ex.Message);
                    }

                    // 再发送图片
                    IntPtr chatHwnd = FindWindow(null, GROUP_NAME);
                    if (chatHwnd == IntPtr.Zero)
                    {
                        Console.WriteLine("未找到群聊窗口");
                        return;
                    }
                    Console.WriteLine($"[调试] 调用SendImageToWeixin，imgPath={imgPath}");
                    Thread thread2 = new Thread(() =>
                    {
                        SendImageToWeixin(chatHwnd, imgPath);
                    });
                    thread2.SetApartmentState(ApartmentState.STA);
                    thread2.Start();
                    thread2.Join(); // 如果需要等待执行完毕
                }
                else
                {
                    Console.WriteLine("超时未检测到图片，流程结束。");
                }

                // 4. 关闭所有Edge浏览器
                try
                {
                    foreach (var proc in Process.GetProcessesByName("msedge"))
                    {
                        proc.Kill();
                    }
                    Console.WriteLine("已关闭所有Edge浏览器进程。");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("关闭Edge浏览器失败：" + ex.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Task异常: {ex.Message}");
            }
            Console.WriteLine($"==== 完成打开 {code} 的所有公司 ====");
        }
        else
        {
            Console.WriteLine($"{code}（前缀{prefix}）：未找到相关船公司数据");
        }

        Console.WriteLine("companyUrls 所有key：");
        foreach (var k in companyUrls.Keys)
        {
            Console.WriteLine($"'{k}'");
        }
    }

    static void SendImageToWeixin(IntPtr chatHwnd, string imgPath)
    {
        // 先激活群聊窗口
        SetForegroundWindow(chatHwnd);
        Thread.Sleep(200);

        // 复制图片到剪贴板
        using (Image img = Image.FromFile(imgPath))
        {
            Clipboard.SetImage(img);
            // 这里img用完后会自动释放
        }
        Thread.Sleep(200);

        // 粘贴并发送
        SendKeysSafe("^v");
        Thread.Sleep(200);
        // 新增：发送提示文本"船期截图请查收"
        SendTextSafe("船期截图请查收");
        Thread.Sleep(200);
        SendKeysSafe("^{ENTER}");
        Thread.Sleep(200);

        try
        {
            File.Delete(imgPath);
            Console.WriteLine("图片已删除：" + imgPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("删除图片失败：" + ex.Message);
        }
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

    // 新增：查找群聊窗口的方法
    static IntPtr FindChatWindow(string chatTitle, int retry = 5)
    {
        for (int i = 0; i < retry; i++)
        {
            var hwnd = FindWindow(null, chatTitle);
            if (hwnd != IntPtr.Zero)
                return hwnd;
            Thread.Sleep(1000);
        }
        return IntPtr.Zero;
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
        string filePath = @"C:\\Users\\x\\Desktop\\剪贴.txt";
        File.WriteAllText(filePath, text, Encoding.UTF8);
        Console.WriteLine("柜号已写入文件：" + filePath + " 内容：" + text);
    }

    static void EnsurePopupClosed(AutomationElement chatWindow)
    {
        try
        {
            // 查找"取消"按钮时加异常保护
            var cancelButton = chatWindow.FindFirstDescendant(cf =>
                cf.ByControlType(FlaUI.Core.Definitions.ControlType.Button)
                .And(cf.ByClassName("mmui::XOutlineButton"))
                .And(cf.ByName("取消"))
            );
            if (cancelButton != null)
            {
                SendKeysSafe("{ESC}");
                Thread.Sleep(300);
                Console.WriteLine("检测到弹窗，已按ESC关闭。");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("EnsurePopupClosed异常：" + ex.Message);
            // 可以选择直接return，防止线程崩溃
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

    static bool CheckRecallAndReturn(AutomationElement chatWindow)
    {
        if (recallDetected)
        {
            Console.WriteLine("主流程检测到撤回，主动退出并执行撤回流程！");
            HandleRecall(chatWindow);
            return true;
        }
        return false;
    }

    static List<string> ExtractProductsFromText(string text, IEnumerable<string> productNames)
    {
        var found = new List<string>();
        foreach (var prod in productNames)
        {
            if (text.Contains(prod))
                found.Add(prod);
        }
        return found;
    }

    static string BuildReplayMenu(List<string> customers, IEnumerable<string> matchedProducts = null)
    {
        var sb = new StringBuilder();
        var productList = matchedProducts == null ? new List<string>() : matchedProducts.Where(p => !string.IsNullOrWhiteSpace(p)).Distinct().ToList();
        if (productList != null && productList.Count > 0)
        {
            sb.AppendLine($"匹配到是{string.Join("、", productList)}客户");
            sb.AppendLine("输入1发送给下面客户");
        }
        else
        {
            sb.AppendLine("匹配到的客户：");
        }
        int idx = 0;
        // 紧凑格式：a邵超超 b牧农 c启商 ...
        foreach (var cust in customers)
        {
            string code = GetAlphaCode(idx);
            sb.Append($"{code}{cust} ");
            idx++;
        }
        sb.AppendLine();
        sb.AppendLine("输入0取消本次操作");
        sb.AppendLine("输入-a就是发送除a以外的客户，-a-b就是除ab以外的客户");
        return sb.ToString();
    }

    // a, b, ..., z, aa, ab, ... az, ba, ...
    static string GetAlphaCode(int idx)
    {
        var code = "";
        do
        {
            code = (char)('a' + (idx % 26)) + code;
            idx = idx / 26 - 1;
        } while (idx >= 0);
        return code;
    }

    static List<string> ParseUserInputToCustomers(string input, List<string> customers)
    {
        input = input.Trim();
        if (input == "1")
        {
            // 输入1时返回全部客户
            return customers;
        }
        if (input.StartsWith("-"))
        {
            var excludes = input.Substring(1).Split('-').ToList();
            var excludeIdx = excludes.Select(code => AlphaCodeToIndex(code)).ToHashSet();
            return customers.Where((c, i) => !excludeIdx.Contains(i)).ToList();
        }
        int idx;
        if (int.TryParse(input, out idx) && idx > 0 && idx <= customers.Count)
            return new List<string> { customers[idx - 1] };
        return new List<string>();
    }

    static int AlphaCodeToIndex(string code)
    {
        int idx = 0;
        foreach (var c in code)
        {
            idx = idx * 26 + (c - 'a' + 1);
        }
        return idx - 1;
    }

    static List<List<string>> SplitByBatch(List<string> customers, int batchSize)
    {
        var result = new List<List<string>>();
        for (int i = 0; i < customers.Count; i += batchSize)
        {
            var batch = new List<string>();
            for (int j = i; j < i + batchSize && j < customers.Count; j++)
            {
                batch.Add(customers[j]);
            }
            result.Add(batch);
        }
        return result;
    }

    // 辅助：查找名称包含关键词的任意控件
    static AutomationElement FindElementByNameContains(AutomationElement root, string keyword)
    {
        if (root == null || string.IsNullOrEmpty(keyword)) return null;
        AutomationElement[] all = Array.Empty<AutomationElement>();
        for (int i = 0; i < 3; i++)
        {
            try
            {
                all = root.FindAllDescendants();
                break;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Thread.Sleep(200);
            }
            catch { break; }
        }
        foreach (var e in all)
        {
            string name = null;
            try { name = e.Name; } catch { }
            if (!string.IsNullOrEmpty(name) && name.Contains(keyword)) return e;
        }
        return null;
    }

    // 辅助：向上寻找可点击祖先（Invoke/Button/LegacyIAccessible），否则返回自身
    static AutomationElement FindClickableAncestor(AutomationElement element, int maxHops = 6)
    {
        var current = element;
        for (int i = 0; i < maxHops && current != null; i++)
        {
            try
            {
                if (current.Patterns.Invoke.IsSupported) return current;
                if (current.ControlType == ControlType.Button) return current;
                if (current.Patterns.LegacyIAccessible.IsSupported) return current;
            }
            catch { }
            try { current = current.Parent; } catch { current = null; }
        }
        return element;
    }

    // 辅助：尝试触发点击（优先Invoke/LegacyIAccessible，否则鼠标点击）
    static bool TryInvokeOrClick(AutomationElement element)
    {
        if (element == null) return false;
        var target = FindClickableAncestor(element);
        try
        {
            if (target.Patterns.Invoke.IsSupported)
            {
                target.Patterns.Invoke.Pattern.Invoke();
                return true;
            }
        }
        catch { }
        try
        {
            if (target.ControlType == ControlType.Button)
            {
                target.AsButton().Invoke();
                return true;
            }
        }
        catch { }
        try
        {
            if (target.Patterns.LegacyIAccessible.IsSupported)
            {
                target.Patterns.LegacyIAccessible.Pattern.DoDefaultAction();
                return true;
            }
        }
        catch { }
        try
        {
            var rect = target.BoundingRectangle;
            int clickX = (int)(rect.Left + rect.Width / 2);
            int clickY = (int)(rect.Top + rect.Height / 2);
            System.Windows.Forms.Cursor.Position = new System.Drawing.Point(clickX, clickY);
            mouse_event(0x02, 0, 0, 0, 0);
            mouse_event(0x04, 0, 0, 0, 0);
            return true;
        }
        catch { }
        return false;
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

    // 安全枚举桌面顶层窗口
    static AutomationElement[] GetDesktopWindowsSafe(AutomationElement anyElement, int retries = 3, int delayMs = 200)
    {
        for (int i = 0; i < retries; i++)
        {
            try
            {
                return anyElement.Automation.GetDesktop().FindAllChildren();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Thread.Sleep(delayMs);
            }
            catch { }
        }
        return Array.Empty<AutomationElement>();
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

    // 新增：双击、靠边工具
    static void DoubleClickElement(AutomationElement el)
    {
        if (el == null) return;
        var rect = el.BoundingRectangle;
        int x = (int)(rect.Left + rect.Width / 2);
        int y = (int)(rect.Top + rect.Height / 2);
        Cursor.Position = new System.Drawing.Point(x, y);
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
        Thread.Sleep(120);
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
    }

    static void SnapWindowLeft(IntPtr hWnd)
    {
        if (hWnd == IntPtr.Zero) return;
        RECT r; if (!GetWindowRect(hWnd, out r)) return;
        int width = r.Right - r.Left;
        int height = r.Bottom - r.Top;
        var wa = Screen.PrimaryScreen.WorkingArea;
        MoveWindow(hWnd, wa.Left, wa.Top, width, height, true);
    }

    static void SnapWindowRightKeepWidthIncreaseHeight(IntPtr hWnd, int deltaHeight)
    {
        if (hWnd == IntPtr.Zero) return;
        RECT r; if (!GetWindowRect(hWnd, out r)) return;
        int width = r.Right - r.Left;
        int height = r.Bottom - r.Top + deltaHeight;
        var wa = Screen.PrimaryScreen.WorkingArea;
        int left = wa.Right - width;
        MoveWindow(hWnd, left, r.Top, width, height, true);
    }

    // 新增：启动时自动靠左、搜索并打开群聊、独立窗靠右并增高
    static void OpenGroupAtStartup()
    {
        try
        {
            var processes = Process.GetProcessesByName("weixin");
            if (processes.Length == 0)
            {
                Console.WriteLine("未找到微信进程。");
                return;
            }
            foreach (var proc in processes)
            {
                using (var app = FlaUI.Core.Application.Attach(proc))
                using (var automation = new UIA3Automation())
                {
                    IntPtr mainHwnd = FindWeChatMainWindow(proc.Id);
                    if (mainHwnd == IntPtr.Zero) continue;

                    // 主窗口靠左
                    SnapWindowLeft(mainHwnd);
                    SetForegroundWindow(mainHwnd);
                    Thread.Sleep(300);

                    var mainWin = automation.FromHandle(mainHwnd);
                    if (mainWin == null) continue;

                    // 搜索框：Class=mmui::XValidatorTextEdit Name=搜索
                    var searchBox = mainWin.FindFirstDescendant(cf =>
                        cf.ByControlType(FlaUI.Core.Definitions.ControlType.Edit)
                          .And(cf.ByClassName("mmui::XValidatorTextEdit"))
                          .And(cf.ByName("搜索"))
                    ) ?? mainWin.FindFirstDescendant(cf =>
                        cf.ByControlType(FlaUI.Core.Definitions.ControlType.Edit)
                          .And(cf.ByClassName("mmui::XValidatorTextEdit"))
                    );

                    if (searchBox != null)
                    {
                        searchBox.Focus();
                        Thread.Sleep(150);
                        if (searchBox.Patterns.Value.IsSupported)
                            searchBox.Patterns.Value.Pattern.SetValue(GROUP_NAME);
                        else
                            SendTextSafe(GROUP_NAME);
                        Thread.Sleep(200);
                        SendKeysSafe("{ENTER}");
                        Thread.Sleep(400);
                    }
                    else
                    {
                        Console.WriteLine("未找到主窗口搜索框。");
                    }

                    // 双击会话列表里 Name 包含群名的 ListItem
                    var candidates = mainWin.FindAllDescendants(cf =>
                        cf.ByControlType(FlaUI.Core.Definitions.ControlType.ListItem)
                          .And(cf.ByClassName("mmui::ChatSessionCell"))
                    );
                    var target = candidates.FirstOrDefault(e =>
                    {
                        try { return (e.Name ?? "").Contains(GROUP_NAME); } catch { return false; }
                    });
                    if (target != null)
                    {
                        DoubleClickElement(target);
                        Thread.Sleep(600);
                    }
                    else
                    {
                        Console.WriteLine("未在会话列表中找到目标群聊项。");
                    }

                    // 等待独立聊天窗口并靠右 + 高度+30
                    AutomationElement chatWin = null;
                    for (int i = 0; i < 20; i++)
                    {
                        var allWins = app.GetAllTopLevelWindows(automation);
                        chatWin = allWins.FirstOrDefault(w => w.ClassName == "mmui::FramelessMainWindow" && w.Name == GROUP_NAME);
                        if (chatWin != null) break;
                        Thread.Sleep(200);
                    }
                    if (chatWin != null)
                    {
                        var h = (IntPtr)chatWin.Properties.NativeWindowHandle.Value;
                        SnapWindowRightKeepWidthIncreaseHeight(h, 30);
                        Console.WriteLine("已将独立群聊窗口靠右并增高30。");
                    }
                    else
                    {
                        Console.WriteLine("未找到独立群聊窗口。");
                    }
                    break;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("OpenGroupAtStartup异常：" + ex);
        }
    }
}

