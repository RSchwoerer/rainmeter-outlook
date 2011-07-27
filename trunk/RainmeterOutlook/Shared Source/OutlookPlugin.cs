using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookPlugin
{

    class OutlookPlugin
    {

        private static Outlook.Application _app;

        private static int status = 0;
        private static string statusMsg = "";

        private static Outlook.Application GetOutlook()
        {
            if (status == 0)
            {
                try
                {

                    _app = new Outlook.Application();
                    status = 1;
                    statusMsg = "Connected";
                }
                catch (Exception e)
                {
                    status = -1;
                    statusMsg = e.Message;
                }
            }
            return _app;
        }

        private Outlook.Application outlook;

        public static int resetId = 0;

        //private static Dictionary<String, MeasureResult> cache = new Dictionary<string,MeasureResult>();

        public OutlookPlugin() 
        {
            outlook = GetOutlook();
        }

        #region Rainmeter Interface

        public UInt32 Update(Rainmeter.Settings.InstanceSettings Instance)
        {
            // not used
            return 0;
        }

        public double Update2(Rainmeter.Settings.InstanceSettings Instance)
        {
            try
            {
                MeasureResult result = Measure(Instance);
                return result.AsDouble(Instance);
            }
            catch (Exception e)
            {
                Rainmeter.Log(Rainmeter.LogLevel.Error, "Sorry, " + e.ToString());
                return double.NaN;
            }
        }

        public string GetString(Rainmeter.Settings.InstanceSettings Instance)
        {
            try
            {
                MeasureResult result = Measure(Instance);
                return result.AsString(Instance);
            }
            catch (Exception e)
            {
                Rainmeter.Log(Rainmeter.LogLevel.Error, "Sorry, " +  e.ToString());
                return "Sorry, " + e.ToString();
            }
        }

        // 'ExecuteBang' is a way of Rainmeter telling your plugin to do something *right now*.
        // What it wants to do can be defined by the 'Command' parameter.
        public void ExecuteBang(Rainmeter.Settings.InstanceSettings Instance, string Command)
        {
            //string[] args = Command.Split(' ');
            //Command = args[0];
            try
            {
                switch (Command)
                {
                    case "ClearCache":
                        resetId++;
                        return;
                }
                MeasureResult mr = Measure(Instance);
                mr.Bang(GetOutlook(), Command);
            }
            catch (Exception e)
            {
                Rainmeter.Log(Rainmeter.LogLevel.Error, e.ToString());
            }
            return;
        }

        #endregion

        #region Measure

        private MeasureResult Measure(Rainmeter.Settings.InstanceSettings Instance)
        {
            lock (Instance)
            {
                MeasureResult cached = GetCached(Instance);
                if (cached != null)
                {
                    int age = (int) Instance.GetTempValue("Age", 0);
                    Instance.SetTempValue("Age", age+1);
                    return cached;
                }
                return Evaluate(Instance);
            }
        }

        private MeasureResult GetCached(Rainmeter.Settings.InstanceSettings Instance)
        {
            MeasureResult cached;
            cached = (MeasureResult) Instance.GetTempValue("Cached", null);
            if (cached == null || !cached.checkAge(Instance))
            {
                return null;
            }
            return cached;
        }

        private MeasureResult Evaluate(Rainmeter.Settings.InstanceSettings Instance)
        {
            MeasureResult result = null;
            try
            {
                result = GetResource(Instance);

                result = result.Select(Instance);
                
                result = result.Filter(Instance);

                string strIndex = Instance.INI_Value("Index");
                if (strIndex.Length > 0)
                {
                    int index = int.Parse(strIndex);
                    result = result.Index(index, Instance);
                }
            }
            catch (Exception e)
            {
                result = new ErrorResult(-1, e.Message);
            }
            finally
            {
                Instance.SetTempValue("Age", 0);
                Instance.SetTempValue("Cached", result);
                Instance.SetTempValue("resetId", OutlookPlugin.resetId);
            }
            return result;
        }

        private MeasureResult GetResource(Rainmeter.Settings.InstanceSettings Instance)
        {
            string resourceKey = Instance.INI_Value("Resource");
            if (resourceKey.Length == 0)
            {
                return new ErrorResult(-1, "Resource required");
            }
            else if (resourceKey.StartsWith("["))
            {
                MeasureResult r;
                if (TryUpdateOtherMeasure(resourceKey, Instance, out r)) return r;
                return new ErrorResult(-1, "Unknown measure " + resourceKey);
            }
            else if (resourceKey == "Status")
            {
                return new StatusResult(status, statusMsg);
            }
            else if (resourceKey == "MAPIFolder" || resourceKey == "EmailFolder")
            {
                return GetMAPIFolders(Instance);
            }
            return new ErrorResult(-1, "Unknown resource '" + resourceKey + "'");
        }

        private bool TryUpdateOtherMeasure(string name, Rainmeter.Settings.InstanceSettings Instance, out MeasureResult result)
        {
            string section = name;
            if (section.StartsWith("["))
            {
                section = section.Substring(1, section.Length - 2);
            }
            Rainmeter.Settings.InstanceSettings other = Instance.GetSection(section);
            if (other == null)
            {
                result = null;
                return false;
            }

            // a measure that is not directly used by a meter does not age,
            // we have to work around this
            int age = Math.Max((int)other.GetTempValue("Age", 0), (int)Instance.GetTempValue("Age", 0));
            other.SetTempValue("Age", age);

            result = Measure(other);
            Instance.SetTempValue("Base", other);
            return true;
        }

        #endregion

        private MeasureResult GetMAPIFolders(Rainmeter.Settings.InstanceSettings Instance)
        {
            MAPIFolderListResult result = new MAPIFolderListResult();
            string rootList = Instance.INI_Value("Root");
            if (rootList.Length == 0) rootList = "Inbox";
            Outlook.NameSpace nsMapi = outlook.GetNamespace("MAPI");
            foreach (string root in rootList.Split('|'))
            {
                if (!root.StartsWith("\\"))
                {
                    switch (root)
                    {
                        case "Inbox":
                            Outlook.MAPIFolder inbox = nsMapi.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                            result.AddRoot(inbox);
                            break;
                        default:
                            return new ErrorResult(-1, root + " not implemented");
                    }
                }
                else
                {
                    Outlook.MAPIFolder folder = FindRoot(nsMapi.Folders, root);
                    if (folder == null)
                    {
                        return new ErrorResult(-1, root + " not found");
                    }
                    result.AddRoot(folder);
                }            
            }
            return result;
        }

        private Outlook.MAPIFolder FindRoot(Outlook.Folders folders, string root)
        {
            string[] path = root.Substring(2).Split('\\');
            return FindRoot(folders, path, 0);
        }

        private Outlook.MAPIFolder FindRoot(Outlook.Folders folders, string[] path, int n)
        {
            foreach (Outlook.Folder f in folders)
            {
                if (f.Name == path[n])
                {
                    if (n + 1 == path.Length)
                    {
                        return f;
                    }
                    return FindRoot(f.Folders, path, n + 1);
                }
            }
            return null;
        }
    }

    abstract class MeasureResult
    {

        private DateTime created = DateTime.Now;

        public bool checkAge(Rainmeter.Settings.InstanceSettings Instance)
        {
            int myResetId = (int) Instance.GetTempValue("resetId", OutlookPlugin.resetId);
            if (myResetId != OutlookPlugin.resetId) return false;

            string strUpdateRate = virtual_INI_value(Instance, "UpdateRate");
            int updateRate;
            if (!int.TryParse(strUpdateRate, out updateRate))
            {
                updateRate = 300;
            }

            int age = (int)Instance.GetTempValue("Age", (int)0);

            return age < updateRate;
        }

        protected virtual string GetResultKey()
        {
            return "Result";
        }

        protected string virtual_INI_value(Rainmeter.Settings.InstanceSettings Instance, String key)
        {
            string r = Instance.INI_Value(key);
            if (r.Length > 0) return r;
            if (Instance.INI_Value("Override") == "1") return "";
            Rainmeter.Settings.InstanceSettings other = (Rainmeter.Settings.InstanceSettings)Instance.GetTempValue("Base", null);
            if (other == null) return "";
            return virtual_INI_value(other, key);
        }

        public double AsDouble(Rainmeter.Settings.InstanceSettings Instance)
        {
            string result = virtual_INI_value(Instance, GetResultKey());
            if (result.StartsWith("%"))
            {
                double? d = GetDouble(result, Instance);
                if (d != null) return (double)d;
                return double.NaN;
            }
            else
            {
                double d = double.NaN;
                double.TryParse(result, out d);
                return d;
            }
        }

        protected virtual double? GetDouble(string key, Rainmeter.Settings.InstanceSettings Instance)
        {
            return null;
        }

        public string AsString(Rainmeter.Settings.InstanceSettings Instance)
        {
            string result = virtual_INI_value(Instance, GetResultKey());
            Regex regex = new Regex("%[a-zA-Z]+");
            return regex.Replace(result, delegate(Match match)
            {
                return GetString(match.ToString(), Instance);
            });
        }

        protected virtual string GetString(string key, Rainmeter.Settings.InstanceSettings Instance)
        {
            double? d = GetDouble(key, Instance);
            if (d != null) return d.ToString();
            return "";
        }

        public virtual MeasureResult Select(Rainmeter.Settings.InstanceSettings Instance)
        {
            return this;
        }

        public virtual MeasureResult Filter(Rainmeter.Settings.InstanceSettings Instance)
        {
            return this;
        }

        public virtual MeasureResult Index(int i, Rainmeter.Settings.InstanceSettings Instance)
        {
            return NullResult.Instance;
        }

        public virtual void Bang(Outlook.Application App, string Command)
        {
            throw new Exception("Unknown command " + Command);
        }
    }

    #region System Results

    class NullResult : MeasureResult
    {
        public static NullResult Instance = new NullResult();

        protected override string GetResultKey()
        {
            return "Default";
        }
    }

    class ErrorResult : MeasureResult
    {
        private int code;
        private string message;

        public ErrorResult(int code, string message)
        {
            this.code = code;
            this.message = message;
        }

        protected override string GetResultKey()
        {
            return "OnError";
        }

        protected override double? GetDouble(string key, Rainmeter.Settings.InstanceSettings Instance)
        {
            if (key == "%Code")
            {
                return code;
            }
            return base.GetDouble(key, Instance);
        }

        protected override string GetString(string key, Rainmeter.Settings.InstanceSettings Instance)
        {
            if (key == "%Message")
            {
                return message;
            }
            return base.GetString(key, Instance);
        }

        public override MeasureResult Index(int i, Rainmeter.Settings.InstanceSettings Instance)
        {
            // don't hide errors after selecting
            return this;
        }
    }

    class StatusResult : MeasureResult
    {
        private int code;
        private string message;

        public StatusResult(int code, string message)
        {
            this.code = code;
            this.message = message;
        }

        private bool IsOk()
        {
            return code >= 0;
        }

        protected override double? GetDouble(string key, Rainmeter.Settings.InstanceSettings Instance)
        {
            if (key == "%Code")
            {
                return code;
            }
            else if (key == "%IsOk")
            {
                return IsOk() ? 1 : 0;
            }
            return null;
        }

        protected override string GetString(string key, Rainmeter.Settings.InstanceSettings Instance)
        {
            if (key == "%Message")
            {
                if (IsOk())
                {
                    string okMsg = virtual_INI_value(Instance, "OkMessage");
                    if (okMsg.Length > 0) return okMsg;
                }
                return message;
            }
            return base.GetString(key, Instance);
        }
    }

    #endregion

    class MAPIFolderListResult : MeasureResult
    {
        private List<MAPIFolderResult> roots;

        protected List<MAPIFolderResult> folders;
        public List<MAPIFolderResult> Folders { get { return folders; } }

        public MAPIFolderListResult()
        {
            this.roots = new List<MAPIFolderResult>();
            this.folders = new List<MAPIFolderResult>();
        }

        private MAPIFolderListResult(List<MAPIFolderResult> folders)
        {
            this.roots = new List<MAPIFolderResult>();
            this.folders = folders;
        }

        public void AddRoot(Outlook.MAPIFolder folder, bool includeRoot = true)
        {
            MAPIFolderResult root = new MAPIFolderResult(folder, 0);
            roots.Add(root);
            if (includeRoot) folders.Add(root);
            fillList(root);
        }

        public void Add(MAPIFolderResult folder)
        {
            this.folders.Add(folder);
        }

        private void fillList(MAPIFolderResult folder)
        {
            foreach (MAPIFolderResult f in folder.Folders)
            {
                folders.Add(f);
                fillList(f);
            }
        }

        protected override double? GetDouble(string key, Rainmeter.Settings.InstanceSettings Instance)
        {
            switch (key)
            {
                case "%Count": return folders.Count;
                case "%TotalUnreadItemCount":
                    Rainmeter.Log(Rainmeter.LogLevel.Error, "0");
                    if (roots.Count > 0)
                    {
                        int total = 0;
                        Rainmeter.Log(Rainmeter.LogLevel.Error, "1");
                        foreach (MAPIFolderResult root in roots)
                        {
                            Rainmeter.Log(Rainmeter.LogLevel.Error, "2");
                            total += root.TotalUnreadItemCount;
                        }
                        Rainmeter.Log(Rainmeter.LogLevel.Error, "3 " + total);
                        return total;
                    }
                    else
                    {
                        int total = 0;
                        Rainmeter.Log(Rainmeter.LogLevel.Error, "4");
                        foreach (MAPIFolderResult f in folders)
                        {
                            Rainmeter.Log(Rainmeter.LogLevel.Error, "5");
                            total += f.UnreadItemCount;
                        }
                        Rainmeter.Log(Rainmeter.LogLevel.Error, "6 " + total);
                        return total;
                    }
                default: return base.GetDouble(key, Instance);
            }
        }

        protected override string GetString(string key, Rainmeter.Settings.InstanceSettings Instance)
        {
            return base.GetString(key, Instance);
        }

        public override MeasureResult Select(Rainmeter.Settings.InstanceSettings Instance)
        {
            string select = Instance.INI_Value("Select");
            if (select == "Root")
            {
                MAPIFolderListResult result = new MAPIFolderListResult();
                foreach (MAPIFolderResult root in roots)
                {
                    result.Add(root);
                }
                return result;
            }
            return this;
        }

        public override MeasureResult Filter(Rainmeter.Settings.InstanceSettings Instance)
        {
            List<MAPIFolderResult> list = folders;
            
            string filter = Instance.INI_Value("Filter");
            if (filter.Length > 0)
            {
                list = list.FindAll(delegate(MAPIFolderResult f)
                {
                    return f.testFilter(filter, Instance);
                });
            }

            string include = Instance.INI_Value("Include");
            if (include.Length > 0)
            {
                include = include.Replace(".", "\\.").Replace("*", ".*");
                Regex regex = new Regex("^(" + include + ")$");
                list = list.FindAll(delegate(MAPIFolderResult f)
                {
                    return regex.IsMatch(f.Name);
                });
            }

            string exclude = Instance.INI_Value("Exclude");
            if (exclude.Length > 0)
            {
                exclude = exclude.Replace(".","\\.").Replace("*", ".*");
                Regex regex = new Regex("^(" + exclude + ")$");
                list = list.FindAll(delegate(MAPIFolderResult f)
                {
                    return ! regex.IsMatch(f.Name);
                });
            }

            if (list == folders) return this;
            return new MAPIFolderListResult(list);
        }

        public override MeasureResult Index(int i, Rainmeter.Settings.InstanceSettings Instance)
        {
            if (0 <= i && i < folders.Count)
            {
                return folders[i];
            }
            return NullResult.Instance;
        }

        public override void Bang(Outlook.Application App, string Command)
        {
            folders[0].Bang(App, Command);
        }
    }

    class MAPIFolderResult : MAPIFolderListResult
    {
        private Outlook.MAPIFolder folder;

        private int depth = -1;
        public int Depth { get { return depth; } }

        private string name = null;
        public string Name { get { if (name == null) name = folder.Name; return name; } }

        private string path = null;
        public string Path { get { if (path == null) path = folder.FolderPath; return path; } }

        private int unread = -1;
        public int UnreadItemCount { get { if (unread == -1) unread = folder.UnReadItemCount; return unread; } }

        private int totalUnread = -1;
        public int TotalUnreadItemCount { get { if (totalUnread == -1) InitTotalUnread(); return totalUnread; } }

        private int itemCount = -1;
        public int ItemCount { get { if (itemCount == -1) itemCount = folder.Items.Count; return itemCount; } }

        public MAPIFolderResult(Outlook.MAPIFolder folder, int depth) : base()
        {
            this.folder = folder;
            this.depth = depth;

            foreach (Outlook.MAPIFolder f in folder.Folders)
            {
                folders.Add(new MAPIFolderResult(f, depth+1));
            }
            folders.Sort(delegate(MAPIFolderResult a, MAPIFolderResult b)
            {
                int c = a.Path.CompareTo(b.Path);
                if (c == 0) c = a.Name.CompareTo(b.Name);
                return c;
            });
        }

        private void InitTotalUnread()
        {
            totalUnread = folder.UnReadItemCount;
            foreach (MAPIFolderResult f in Folders)
            {
                totalUnread += f.UnreadItemCount;
            }
        }

        public bool testFilter(string filter, Rainmeter.Settings.InstanceSettings Instance)
        {
            double? d = GetDouble(filter, Instance);
            return d == 1;
        }

        protected override double? GetDouble(string key, Rainmeter.Settings.InstanceSettings Instance)
        {
            switch (key)
            {
                case "%TotalUnreadItemCount": return TotalUnreadItemCount;
                case "%UnreadItemCount": return UnreadItemCount;
                case "%HasUnreadItems": return UnreadItemCount > 0 ? 1 : 0;
                case "%ItemCount": return ItemCount;
                default: return base.GetDouble(key, Instance);
            }
        }

        protected override string GetString(string key, Rainmeter.Settings.InstanceSettings Instance)
        {
            switch (key)
            {
                case "%Name": return Name;
                case "%Path": return Path;
                case "%Indent": return Indent(Instance);
                default: return base.GetString(key, Instance);
            }
        }

        private string Indent(Rainmeter.Settings.InstanceSettings Instance)
        {
            string indent = virtual_INI_value(Instance, "Indent");
            if (indent == "") indent = "  ";

            if (indent.Length > 1)
            {
                int start = indent.StartsWith("\"") ? 1 : 0;
                int end = indent.EndsWith("\"") ? 1 : 0;
                if (start + end > 0) indent = indent.Substring(start, indent.Length - start - end);
            }

            string result = "";
            for (int i = 0; i < depth; i++)
                result += indent;

            return result;
        }

        public override MeasureResult Select(Rainmeter.Settings.InstanceSettings Instance)
        {
            string select = Instance.INI_Value("Select");
            if (select == "Subfolders")
            {
                MAPIFolderListResult result = new MAPIFolderListResult();
                result.AddRoot(folder, false);
                return result;
            }
            return base.Select(Instance);
        }

        public override void Bang(Outlook.Application App, string Command)
        {
            switch (Command)
            {
                case "Display":
                    folder.Display();
                    return;
            }
        }
    }

}
