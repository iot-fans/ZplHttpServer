using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Net;
using System.Net.Sockets;
using Zebra.Sdk.Comm;
using Zebra.Sdk.Printer;
using Zebra.Sdk.Printer.Discovery;

namespace ZplHttpServer
{
    public partial class ZplHttpServer : IDisposable
    {
        #region initial

        private TcpTransport transHB = null;
        private System.Timers.Timer hbDetecter = new System.Timers.Timer();
        public event HandleStringMsg OnMsg;

        private RollingFileWriterManager rfwManager = null;
        private bool enableWriteFile = false;
        private int maxFileKiloByte = 10000;
        private int maxFileNum = 10;

        public string currDir = "";
        public string ZPLDirectory = "";

        static HttpListener httpobj;
        private string httpip = "127.0.0.1";
        private int httpport = 9037;

        private string PrinterConectionType = "USB";
        private string printerip = "127.0.0.1";
        private int printerport = 9100;

        private System.Timers.Timer PrintDetecter = new System.Timers.Timer();
        private List<string> Printers = new List<string>(); //设备列表
        private object lockObject = new object();

        private DiscoveryHandlerImpl discoveryHandlerImpl = new DiscoveryHandlerImpl();

        private int refreshListen = 0;

        #endregion

        public ZplHttpServer()
        {
            try
            {
                enableWriteFile = bool.Parse(ConfigurationManager.AppSettings["EnableWriteFile"]);
            }
            catch (Exception) { }
            if (enableWriteFile)
            {
                try
                {
                    maxFileKiloByte = int.Parse(ConfigurationManager.AppSettings["MaxFileKiloByte"]);
                }
                catch (Exception) { }

                try
                {
                    maxFileNum = int.Parse(ConfigurationManager.AppSettings["MaxFileNum"]);
                }
                catch (Exception) { }

                rfwManager = new RollingFileWriterManager(maxFileKiloByte, maxFileNum);

                try
                {
                    rfwManager.WriteInterval = int.Parse(ConfigurationManager.AppSettings["WriteFileInterval"]);
                }
                catch (Exception) { }

                try
                {
                    rfwManager.FileNameSuffix = ConfigurationManager.AppSettings["FileNameSuffix"];
                }
                catch (Exception) { }

                try
                {
                    rfwManager.OutputPath = ConfigurationManager.AppSettings["OutputPath"];
                }
                catch (Exception) { }

                rfwManager.OnErrMsg += rfwManagerOnErrMsg;
                rfwManager.StartWrite();
            }

            try
            {
                httpip = ConfigurationManager.AppSettings["HttpIP"];
                httpport = int.Parse(ConfigurationManager.AppSettings["HttpPort"]);
            }
            catch (Exception) { }

            try
            {
                PrinterConectionType = ConfigurationManager.AppSettings["PrinterConectionType"];
                printerip = ConfigurationManager.AppSettings["PrinterIP"];
                printerport = int.Parse(ConfigurationManager.AppSettings["PrinterPort"]);
            }
            catch (Exception) { }

            currDir = Assembly.GetExecutingAssembly().Location.ToString();
            string[] dirs = currDir.Split('\\', '/');
            currDir = "";
            for (int i = 0; i < dirs.Length - 1; i++)
            {
                if (i > 0)
                {
                    currDir += "\\";
                }
                currDir += dirs[i];
            }
            DirectoryInfo dirInfo = new DirectoryInfo(currDir + "\\Data\\ZPL\\");
            if (!dirInfo.Exists)
                dirInfo.Create();
            ZPLDirectory = currDir + "\\Data\\ZPL\\";

            //查找设备
            lock (lockObject)
            {
                Printers.Clear();
                try
                {
                    foreach (DiscoveredUsbPrinter printer in UsbDiscoverer.GetZebraUsbPrinters(new ZebraPrinterFilter()))
                    {
                        Printers.Add(printer.Address);
                    }
                }
                catch
                {

                }
            }

            //定时搜索打印机
            PrintDetecter.Elapsed += PrintDetecter_Elapsed;
            PrintDetecter.Interval = 120000;
        }

        #region base

        private DateTime lastHeartbeatTime = DateTime.Now;// 上次心跳时间
        private void transHBOnReceivedData(byte[] data, int dataLen)
        {
            lastHeartbeatTime = DateTime.Now;// 收到心跳信息 更新上次心跳时间
        }
        private void hbDetecterElapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            TimeSpan ts = DateTime.Now - lastHeartbeatTime;
            if (ts.TotalSeconds > 12)
            {
                hbDetecter.Stop();
                hbDetecter.Elapsed -= hbDetecterElapsed;
                if (transHB != null)
                {
                    transHB.Stop();
                    transHB.OnReceivedData -= transHBOnReceivedData;
                }
                Stop();
                Environment.Exit(-1);
            }
            if (transHB != null)
            {
                transHB.AsyncSendData(new byte[] { 0x0A, 0x03 });
            }
        }
        public void Init(ushort heartbeatPort)
        {
            if (heartbeatPort > 0)
            {
                transHB = new TcpTransport
                {
                    IP = "127.0.0.1",
                    Port = heartbeatPort
                };
                transHB.OnReceivedData += transHBOnReceivedData;
                transHB.Start();

                lastHeartbeatTime = DateTime.Now;
                hbDetecter.Elapsed += hbDetecterElapsed;
                hbDetecter.Interval = 1000;
                hbDetecter.Start();
            }
        }
        public void Start()
        {
            try
            {
                //提供一个简单的、可通过编程方式控制的 HTTP 协议侦听器。此类不能被继承。
                httpobj = new HttpListener();
                //定义url及端口号，通常设置为配置文件
                string httpaddr = "http://" + httpip + ":" + httpport.ToString() + "/";
                //string httpaddr = $"http://{httpip}:{httpport}/";
                httpobj.Prefixes.Add(httpaddr);

                httpobj.TimeoutManager.IdleConnection = TimeSpan.FromSeconds(10);

                //启动监听器
                httpobj.Start();
                //异步监听客户端请求，当客户端的网络请求到来时会自动执行Result委托
                //该委托没有返回值，有一个IAsyncResult接口的参数，可通过该参数获取context对象
                httpobj.BeginGetContext(Result, null);
                ShowMsg(StringMsgType.Info, "ALL", "Start listen to : " + httpaddr);

                PrintDetecter.Start();
            }
            catch (Exception ex)
            {
                ShowMsg(StringMsgType.Error, "ALL", "Failed to do auto start. " + ex.Message);
                Environment.Exit(-1);
            }
        }
        public void Stop()
        {
            ShowMsg(StringMsgType.Info, "Base", "即将关闭后台进程");
            try
            {
                httpobj.Stop();
                httpobj.Abort();
                httpobj.Close();
            }
            catch
            { }

            PrintDetecter.Stop();

            Thread.Sleep(2000);
            if (rfwManager != null)
            {
                rfwManager.StopWrite();
                Thread.Sleep(rfwManager.WriteInterval + 1000);
            }
        }
        public void Dispose()
        {
            try
            {
                Stop();
            }
            catch { }
        }
        public void ShowMsg(StringMsgType msgType, string machineID, string msg)
        {
            if (msgType == StringMsgType.Info || msgType == StringMsgType.Warning || msgType == StringMsgType.Error)
            {
                try
                {
                    if (OnMsg != null)
                    {
                        try { OnMsg(msgType, machineID + ": " + msg + "\r\n"); }
                        catch { }
                    }
                }
                catch
                {

                }
            }
            if (rfwManager != null)
            {
                if (msgType == StringMsgType.Data)
                {
                    return;
                }
                rfwManager.AddData(machineID, msgType.ToString(), DateTime.Now.ToString("yyyyMMdd HH:mm:ss.fff") + ": " + msg + "\r\n");
            }
        }
        private void rfwManagerOnErrMsg(string errMsg)
        {
            ShowMsg(StringMsgType.Error, "FileWriter", errMsg);
        }

        #endregion

        #region http

        private void Result(IAsyncResult ar)
        {
            //当接收到请求后程序流会走到这里

            var context = httpobj.EndGetContext(ar);

            //继续异步监听
            try
            {
                httpobj.BeginGetContext(Result, null);
            }
            catch (Exception ex)
            {
                ShowMsg(StringMsgType.Error, "Http", $"BeginGetContext Error：{ex.Message}");
            }

            var guid = Guid.NewGuid().ToString();
            ShowMsg(StringMsgType.Info, "Http", $"接到新的请求:{guid}.");
            //获得context对象
            var request = context.Request;
            var response = context.Response;
            ////如果是js的ajax请求，还可以设置跨域的ip地址与参数
            //context.Response.AppendHeader("Access-Control-Allow-Origin", "*");//后台跨域请求，通常设置为配置文件
            //context.Response.AppendHeader("Access-Control-Allow-Headers", "ID,PW");//后台跨域参数设置，通常设置为配置文件
            //context.Response.AppendHeader("Access-Control-Allow-Method", "post");//后台跨域请求设置，通常设置为配置文件
            context.Response.ContentType = "text/plain;charset=UTF-8";//告诉客户端返回的ContentType类型为纯文本格式，编码为UTF-8
            context.Response.AddHeader("Content-type", "text/plain");//添加响应头信息
            context.Response.ContentEncoding = Encoding.UTF8;
            string returnObj = "ERROR REQUEST";//定义返回客户端的信息
            if (request.HttpMethod == "POST" && request.InputStream != Stream.Null)
            {
                string tempurl = request.RawUrl.ToUpper();
                if (tempurl == "/ZPL" || tempurl == "/ZPL/")
                {
                    returnObj = HandlePostZPL(request, response);
                }
                if (tempurl == "/FILE" || tempurl == "/FILE/")
                {
                    returnObj = HandlePostFile(request, response);
                }
            }
            if (request.HttpMethod == "GET")
            {
                string tempurl = request.RawUrl.ToUpper();
                if (tempurl == "/")
                {
                    returnObj = HandleGet(request, response);
                }
            }
            var returnByteArr = Encoding.UTF8.GetBytes(returnObj);//设置客户端返回信息的编码
            try
            {
                using (var stream = response.OutputStream)
                {
                    //把处理信息返回到客户端
                    stream.Write(returnByteArr, 0, returnByteArr.Length);
                }
            }
            catch (Exception ex)
            {
                ShowMsg(StringMsgType.Error, "Http", $"网络异常：{ex.Message}");
            }
            ShowMsg(StringMsgType.Info, "Http", $"请求[{request.HttpMethod}][{request.RawUrl}]处理完成:{guid}.");
        }

        private string HandlePostZPL(HttpListenerRequest request, HttpListenerResponse response)
        {
            string data = null;
            string Printer = "";
            try
            {
                var byteList = new List<byte>();

                if (request.ContentType != null)
                {
                    if (request.ContentType.Contains("multipart/form-data"))
                    {
                        Dictionary<string, string> args = new Dictionary<string, string>();
                        Dictionary<string, HttpFile> files = RequestMultipartExtensions.ParseMultipartForm(request, args);
                        if (args.ContainsKey("Printer"))
                        {
                            Printer = args["Printer"];
                        }
                        else
                        {
                            lock (lockObject)
                            {
                                if (Printers.Count > 0)
                                {
                                    Printer = Printers[0];
                                }
                            }
                        }
                        if (args.ContainsKey("ZPL"))
                        {
                            data = args["ZPL"];
                            byte[] byteArr = request.ContentEncoding.GetBytes(data);
                            byteList.AddRange(byteArr);
                        }
                        else
                        {
                            throw new Exception("Not found 'zpl' in multipart/form-data .");
                        }
                    }
                    if (request.ContentType.Contains("application/x-www-form-urlencoded"))
                    {
                        Dictionary<string, string> args = new Dictionary<string, string>();
                        if (RequestFormExtensions.ParseForm(request, args))
                        {
                            if (args.ContainsKey("Printer"))
                            {
                                Printer = args["Printer"];
                            }
                            else
                            {
                                lock (lockObject)
                                {
                                    if (Printers.Count > 0)
                                    {
                                        Printer = Printers[0];
                                    }
                                }
                            }
                            if (args.ContainsKey("ZPL"))
                            {
                                data = args["ZPL"];
                                byte[] byteArr = request.ContentEncoding.GetBytes(data);
                                byteList.AddRange(byteArr);
                            }
                            else
                            {
                                throw new Exception("Not found 'tspl' in application/x-www-form-urlencoded .");
                            }
                        }
                        else
                        {
                            throw new Exception("Parse application/x-www-form-urlencoded Error");
                        }
                    }
                    if (request.ContentType.Contains("text/plain"))
                    {
                        var byteArr = new byte[2048];
                        int readLen = 0;
                        int len = 0;
                        //接收客户端传过来的数据并转成字符串类型
                        do
                        {
                            readLen = request.InputStream.Read(byteArr, 0, byteArr.Length);
                            len += readLen;
                            byteList.AddRange(byteArr);
                        } while (readLen != 0);
                        data = Encoding.UTF8.GetString(byteList.ToArray(), 0, len);
                    }
                }

                //获取得到数据data可以进行其他操作
                Connection printerConnection = null;
                try
                {
                    if(PrinterConectionType == "USB")
                    {
                        printerConnection = new UsbConnection($"{Printer}"/*SymbolicName*/);
                    }
                    else
                    {
                        printerConnection = new TcpConnection(printerip, printerport);
                    }
                    printerConnection.Open();
                    printerConnection.Write(byteList.ToArray());
                }
                catch (ConnectionException ex)
                {
                    throw new Exception($"Communications Error:[{ex.Message}]");
                }
                catch (ZebraPrinterLanguageUnknownException)
                {
                    throw new Exception("Invalid Printer Language");
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
                finally
                {
                    try
                    {
                        try
                        {
                            if (printerConnection != null)
                                printerConnection.Close();
                        }
                        catch (ConnectionException) { }
                    }
                    catch (ConnectionException)
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                response.StatusDescription = "404";
                response.StatusCode = 404;
                ShowMsg(StringMsgType.Error, "Printer", $"发送ZPL失败:{ex.Message}.");
                return $"ERROR,{ex.Message}";
            }
            response.StatusDescription = "200";//获取或设置返回给客户端的 HTTP 状态代码的文本说明。
            response.StatusCode = 200;// 获取或设置返回给客户端的 HTTP 状态代码。
            ShowMsg(StringMsgType.Info, "Printer", $"发送ZPL完成:{data.Trim()}.");
            return "OK";
        }

        private string HandlePostFile(HttpListenerRequest request, HttpListenerResponse response)
        {
            try
            {
                string Printer = "";
                //转存文件
                List<string> filepathlist = new List<string>();
                if (request.ContentType != null)
                {
                    if (request.ContentType.Contains("multipart/form-data"))
                    {
                        Dictionary<string, string> args = new Dictionary<string, string>();
                        Dictionary<string, HttpFile> files = RequestMultipartExtensions.ParseMultipartForm(request, args);
                        string TempID = "";
                        foreach (var f in files.Values)
                        {
                            f.Save(Path.Combine(ZPLDirectory, f.FileName), true);
                            filepathlist.Add(Path.Combine(ZPLDirectory, f.FileName));
                            TempID = f.FileName;
                            ShowMsg(StringMsgType.Info, "Printer", $"下载标签[{TempID}]完成.");
                        }
                        if (args.ContainsKey("Printer"))
                        {
                            Printer = args["Printer"];
                        }
                        else
                        {
                            lock (lockObject)
                            {
                                if (Printers.Count > 0)
                                {
                                    Printer = Printers[0];
                                }
                            }
                        }
                    }
                    else
                    {
                        throw new Exception("ContentType Error");
                    }
                }
                else
                {
                    throw new Exception("ContentType NULL Error");
                }
                //using (var stream = request.InputStream)
                //{
                //    using (var br = new BinaryReader(stream))
                //    {
                //        byte[] fileContents = new byte[] { };
                //        var bytes = new byte[request.ContentLength64];
                //        int i = 0;
                //        while ((i = br.Read(bytes, 0, (int)request.ContentLength64)) != 0)
                //        {
                //            byte[] arr = new byte[fileContents.LongLength + i];
                //            fileContents.CopyTo(arr, 0);
                //            Array.Copy(bytes, 0, arr, fileContents.Length, i);
                //            fileContents = arr;
                //        }

                //        using (var fs = new FileStream(filePath, FileMode.Create))
                //        {
                //            using (var bw = new BinaryWriter(fs))
                //            {
                //                bw.Write(fileContents);
                //            }
                //        }
                //    }
                //}

                if (Printer == "")
                {
                    throw new Exception("Printer Not Found");
                }

                Connection printerConnection = null;
                try
                {
                    if (PrinterConectionType == "USB")
                    {
                        printerConnection = new UsbConnection($"{Printer}"/*SymbolicName*/);
                    }
                    else
                    {
                        printerConnection = new TcpConnection(printerip, printerport);
                    }
                    printerConnection.Open();
                    ZebraPrinter printer = ZebraPrinterFactory.GetInstance(printerConnection);
                    for (int i = 0; i < filepathlist.Count; i++)
                    {
                        printer.SendFileContents(filepathlist[i]);
                    }
                }
                catch (ConnectionException)
                {
                    throw new Exception("Connection Error");
                }
                catch (IOException)
                {
                    throw new Exception("IO Error");
                }
                catch (ZebraPrinterLanguageUnknownException)
                {
                    throw new Exception("Connection Error");
                }
                catch (Exception)
                {
                    throw new Exception("Send File Error");
                }
                finally
                {
                    try
                    {
                        if (printerConnection != null)
                            printerConnection.Close();
                    }
                    catch (ConnectionException) { }

                    //if (filePath != null)
                    //{
                    //    try
                    //    {
                    //        new FileInfo(filePath).Delete();
                    //    }
                    //    catch { }
                    //}
                }
            }
            catch (Exception ex)
            {
                response.StatusDescription = "404";
                response.StatusCode = 404;
                ShowMsg(StringMsgType.Error, "Printer", $"发送FILE失败:{ex.Message}.");
                return $"ERROR,{ex.Message}";
            }
            response.StatusDescription = "200";//获取或设置返回给客户端的 HTTP 状态代码的文本说明。
            response.StatusCode = 200;// 获取或设置返回给客户端的 HTTP 状态代码。
            ShowMsg(StringMsgType.Info, "Printer", $"发送FILE完成.");
            return "OK";
        }

        private string HandleGet(HttpListenerRequest request, HttpListenerResponse response)
        {
            StringBuilder ps = new StringBuilder();
            ps.Append("{\"Printers\":[");
            bool isfirst = true;
            try
            {
                lock (lockObject)
                {
                    Printers.Clear();//刷新设备列表
                    if (PrinterConectionType == "USB")
                    {
                        foreach (DiscoveredUsbPrinter printer in UsbDiscoverer.GetZebraUsbPrinters(new ZebraPrinterFilter()))
                        {
                            if (isfirst)
                            {
                                isfirst = false;
                            }
                            else
                            {
                                ps.Append(",");
                            }
                            ps.Append(printer.Address);
                            Printers.Add(printer.Address);
                        }
                    }
                    if (PrinterConectionType == "TCP")
                    {
                        discoveryHandlerImpl.isFinding = true;
                        NetworkDiscoverer.FindPrinters(discoveryHandlerImpl);
                        while (discoveryHandlerImpl.isFinding)
                        {
                            Thread.Sleep(1000);
                        }
                        foreach (DiscoveredPrinter printer in discoveryHandlerImpl.foundprinters.Values)
                        {
                            if (isfirst)
                            {
                                isfirst = false;
                            }
                            else
                            {
                                ps.Append(",");
                            }
                            ps.Append(printer.Address);
                            Printers.Add(printer.Address);
                        }
                    }
                }
            }
            catch
            {

            }
            ps.Append("]}");
            response.StatusDescription = "200";
            response.StatusCode = 200;
            return ps.ToString();
        }

        #endregion

        #region Printer

        private class DiscoveryHandlerImpl : DiscoveryHandler
        {
            public Dictionary<string, DiscoveredPrinter> foundprinters = new Dictionary<string, DiscoveredPrinter>();
            public bool isFinding = true;
            public DiscoveryHandlerImpl() { }

            public void DiscoveryError(string message)
            {
                Console.WriteLine("DiscoveryError");
                Console.WriteLine(message);
            }

            public void DiscoveryFinished()
            {
                isFinding = false;
                Console.WriteLine("DiscoveryFinished");
            }

            public void FoundPrinter(DiscoveredPrinter printer)
            {
                if(foundprinters.ContainsKey(printer.Address))
                {
                    foundprinters[printer.Address] = printer;
                }
                else
                {
                    foundprinters.Add(printer.Address, printer);
                }
            }
        }

        private void PrintDetecter_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            lock (lockObject)
            {
                Printers.Clear();
                try
                {
                    if (PrinterConectionType == "USB")
                    {
                        foreach (DiscoveredUsbPrinter printer in UsbDiscoverer.GetZebraUsbPrinters(new ZebraPrinterFilter()))
                        {
                            Printers.Add(printer.Address);
                        }
                    }
                    if (PrinterConectionType == "TCP")
                    {
                        discoveryHandlerImpl.isFinding = true;
                        NetworkDiscoverer.FindPrinters(discoveryHandlerImpl);
                        while (discoveryHandlerImpl.isFinding)
                        {
                            Thread.Sleep(1000);
                        }
                        foreach (DiscoveredPrinter printer in discoveryHandlerImpl.foundprinters.Values)
                        {
                            Printers.Add(printer.Address);
                        }
                    }
                }
                catch
                {

                }
            }
            try
            {
                if (refreshListen > 120)
                {
                    refreshListen = 0;

                    //开启异步监听
                    try
                    {
                        httpobj.BeginGetContext(Result, null);
                    }
                    catch (Exception ex)
                    {
                        ShowMsg(StringMsgType.Error, "Http", $"BeginGetContext Error：{ex.Message}");
                    }
                }
            }
            catch
            {

            }
            refreshListen++;
        }

        #endregion
    }
}
