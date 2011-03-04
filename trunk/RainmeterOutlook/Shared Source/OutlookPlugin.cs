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

        private static Dictionary<String, MeasureResult> cache = new Dictionary<string,MeasureResult>();

        public OutlookPlugin() 
        {
            outlook = GetOutlook();
        }

        public UInt32 Update(Rainmeter.Settings.InstanceSettings Instance)
        {
            // not used
            return 0;
        }

        public double Update2(Rainmeter.Settings.InstanceSettings Instance)
        {
            MeasureResult result = Measure(Instance);
            return result.AsDouble(Instance);
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
                return "Sorry, " + e.ToString();
            }
        }


        // 'ExecuteBang' is a way of Rainmeter telling your plugin to do something *right now*.
        // What it wants to do can be defined by the 'Command' parameter.
        public void ExecuteBang(Rainmeter.Settings.InstanceSettings Instance, string Command)
        {
            return;
        }

        private MeasureResult Measure(Rainmeter.Settings.InstanceSettings Instance)
        {
            MeasureResult cached = GetCached(Instance.Section, Instance);
            if (cached != null) return cached;
            return Evaluate(Instance);
        }

        private MeasureResult GetCached(string section, Rainmeter.Settings.InstanceSettings Instance)
        {
            MeasureResult cached;
            if (!section.StartsWith("[")) section = "[" + section + "]";
            if (!cache.TryGetValue(Instance.INI_File + section, out cached))
            {
                return null;
            }

            string strUpdateRate = Instance.INI_Value("UpdateRate").Trim();
            int updateRate;
            if (!int.TryParse(strUpdateRate, out updateRate))
            {
                updateRate = 300;
            }

            if (cached.Age > updateRate) return null;

            return cached;
        }

        private MeasureResult Evaluate(Rainmeter.Settings.InstanceSettings Instance)
        {
            MeasureResult result = null;
            try
            {
                result = GetResource(Instance);
                result = result.Filter(Instance);
                string strIndex = Instance.INI_Value("Index").Trim();
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
                cache[Instance.INI_File + "[" + Instance.Section + "]"] = result;
            }
            return result;
        }

        private MeasureResult GetResource(Rainmeter.Settings.InstanceSettings Instance)
        {
            string resourceKey = Instance.INI_Value("Resource").Trim();
            if (resourceKey.Length == 0)
            {
                return new ErrorResult(-1, "Resource required");
            }
            else if (resourceKey.StartsWith("["))
            {
                MeasureResult r = GetCached(resourceKey, Instance);
                if (r != null) return r;
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
            result = Measure(other);
            Instance.SetTempValue("Base", other);
            return true;
        }

        private MeasureResult GetMAPIFolders(Rainmeter.Settings.InstanceSettings Instance)
        {
            string root = Instance.INI_Value("Root").Trim();
            if (root.Length == 0) root = "Inbox";
            if (!root.StartsWith("\\"))
            {
                Outlook.NameSpace nsMapi = outlook.GetNamespace("MAPI");
                if (root == "Inbox")
                {
                    Outlook.MAPIFolder inbox = nsMapi.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                    return new MAPIFolderListResult(inbox);
                }
            }
            return new ErrorResult(-1, root + " not implemented");
        }
    }

    abstract class MeasureResult
    {

        private DateTime created = DateTime.Now;

        public double Age
        {
            get { return (DateTime.Now - created).TotalSeconds; }
        }

        protected virtual string GetResultKey()
        {
            return "Result";
        }

        protected string virtual_INI_value(Rainmeter.Settings.InstanceSettings Instance, String key)
        {
            string r = Instance.INI_Value(key);
            if (r.Length > 0) return r;
            if (Instance.INI_Value("Override").Trim() == "1") return "";
            Rainmeter.Settings.InstanceSettings other = (Rainmeter.Settings.InstanceSettings)Instance.GetTempValue("Base", null);
            if (other == null) return "";
            return virtual_INI_value(other, key);
        }

        public double AsDouble(Rainmeter.Settings.InstanceSettings Instance)
        {
            string result = virtual_INI_value(Instance, GetResultKey()).Trim();
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

        public virtual MeasureResult Filter(Rainmeter.Settings.InstanceSettings Instance)
        {
            return this;
        }

        public virtual MeasureResult Index(int i, Rainmeter.Settings.InstanceSettings Instance)
        {
            return NullResult.Instance;
        }
    }

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

    class MAPIFolderListResult : MeasureResult
    {
        private MAPIFolderResult root;
        private List<MAPIFolderResult> folders;

        public MAPIFolderListResult(Outlook.MAPIFolder folder)
        {
            root = new MAPIFolderResult(folder, 0);
            this.folders = new List<MAPIFolderResult>();
            fillList(root);
        }

        private MAPIFolderListResult(List<MAPIFolderResult> folders)
        {
            this.folders = folders;
        }

        private void fillList(MAPIFolderResult folder)
        {
            folders.Add(folder);
            foreach (MAPIFolderResult f in folder.Folders)
            {
                fillList(f);
            }
        }

        protected override double? GetDouble(string key, Rainmeter.Settings.InstanceSettings Instance)
        {
            switch (key)
            {
                case "%Count": return folders.Count;
                case "%TotalUnreadItemCount": return root.TotalUnreadItemCount;
                default: return base.GetDouble(key, Instance);
            }
        }

        protected override string GetString(string key, Rainmeter.Settings.InstanceSettings Instance)
        {
            return base.GetString(key, Instance);
        }

        public override MeasureResult Filter(Rainmeter.Settings.InstanceSettings Instance)
        {
            List<MAPIFolderResult> list = folders;
            string filter = Instance.INI_Value("Filter").Trim();
            if (filter.Length > 0)
            {
                list = list.FindAll(delegate(MAPIFolderResult f)
                {
                    return f.testFilter(filter, Instance);
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
    }

    class MAPIFolderResult : MeasureResult
    {
        private Outlook.MAPIFolder folder;
        
        private List<MAPIFolderResult> folders;
        public List<MAPIFolderResult> Folders { get { return folders; } }

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

        public MAPIFolderResult(Outlook.MAPIFolder folder, int depth)
        {
            this.folder = folder;
            this.depth = depth;

            folders = new List<MAPIFolderResult>();
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
            if (indent.EndsWith(".")) indent = indent.Substring(0, indent.Length - 1);

            string result = "";
            for (int i = 0; i < depth; i++)
                result += indent;

            return result;
        }
    }

}
