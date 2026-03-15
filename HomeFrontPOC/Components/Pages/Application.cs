using System;
using System.Data;
using System.Diagnostics;
using System.Dynamic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using HomeFront.Data;
using Microsoft.JSInterop;

namespace HomeFront.Components.Pages
{
    public class Application
    {
        private const string SRCFILE = "CApp::";
        private const uint SEM_NOGPFAULTERRORBOX = 0x2;
        private const int DbCount = 8;

        [DllImport("kernel32.dll")]
        private static extern uint SetErrorMode(uint mode);

        public enum Connections
        {
            dbHomefront,
            dbEstimating,
            dbAccounting,
            dbReceivables,
            dbAccountingDictionary,
            dbAccountingCustom,
            dbReceivablesDictionary,
            dbTimberlinePVdata
        }

        public enum EstimatingSystems
        {
            esNone = 0,
            esPipeline = 1,
            esTimberline = 2
        }

        public enum AccountingSystems
        {
            asNone = 0,
            asTimberline = 1,
            asMasterBuilder = 2,
            asQuickBooks = 3,
            asSimply = 4,
            asMYOB = 5,
            asPeachtree = 6,
            asXero = 7,
            asSpectrum = 8,
            asIntacct = 9,
            asQuickbooksOnline = 10
        }

        public enum SalesSystems
        {
            asNone = 0,
            asHomeFront = 1,
            asBuilder1440 = 2
        }

        public enum XMLFieldTypes
        {
            a, b, c, d, f, g, m, n, o, p, r, st, t, Y, z
        }

        private enum DbQuoteType
        {
            Str,
            Num,
            Date
        }

        private bool _isProduction;
        private readonly DbWrapperSqlServer?[] _databases = new DbWrapperSqlServer[DbCount];
        private readonly string[] _connectionStrings = new string[DbCount];
        private dynamic _license = new LicenseStub();
        private string _sqlEdition = "";

        private dynamic _simplyOdbcUtilities;

        private long _userLevel;
        private string _loginId = "";
        private string _loginPswd = "";
        private string _divisionId = "";

        public dynamic Options { get; private set; } = new OptionsStub();

        public static Application? HFApp { get; private set; }

        /// <summary>
        /// Optional JSRuntime for browser interop (e.g. alert dialogs).
        /// Set this from the DI container or hosting layer if JS interop is needed.
        /// </summary>
        public static IJSRuntime? JSRuntime { get; set; }

        private string _productName = "";
        private string _productCode = "";
        private string _moduleCode = "";
        private string _moduleName = "";
        private string _moduleVersion = "";
        private string _systemVersion = "";
        private string _exePath = "";
        private string _trustedConnection = "";
        private bool _trustedSkipAccounting;
        private string _trustedUser = "";
        private bool _closing;
        private bool _jobSimplicity;

        public void Initialize()
        {
            HFApp = this;
            Options = CreateOptions();
            _license = CreateLicense();
            _databases[(int)Connections.dbHomefront] ??= new DbWrapperSqlServer();
        }

        // ─────────────────────────────────────────────────────────────────────────
        // Helpers
        // ─────────────────────────────────────────────────────────────────────────
        private dynamic CreateLicense()
        {
            var type = Type.GetType("CLicense") ?? Type.GetType("HomeFront.CLicense");
            return type != null ? Activator.CreateInstance(type)! : new LicenseStub();
        }

        private dynamic CreateOptions()
        {
            var type = Type.GetType("Options") ?? Type.GetType("HomeFront.Options");
            return type != null ? Activator.CreateInstance(type)! : new OptionsStub();
        }

        private void MsgBox(string message)
        {
            Console.WriteLine(message);
            try { _ = JSRuntime?.InvokeVoidAsync("alert", message); } catch { }
        }

        private DbWrapperSqlServer GetDb(Connections d)
        {
            var idx = (int)d;
            if (_databases[idx] == null)
            {
                var cs = _connectionStrings[idx];
                _databases[idx] = string.IsNullOrWhiteSpace(cs) ? new DbWrapperSqlServer() : new DbWrapperSqlServer(cs);
            }
            return _databases[idx]!;
        }

        private string OptionValue(string name) => Options?.ValueByName(name) ?? "";
        private bool OptionBool(string name) => string.Equals(OptionValue(name), "true", StringComparison.OrdinalIgnoreCase);
        private int OptionInt(string name) => int.TryParse(OptionValue(name), out var v) ? v : 0;

        private static string Parse(string text, int index = 0, string delimiter = ",")
        {
            var parts = (text ?? "").Split(new[] { delimiter }, StringSplitOptions.None);
            if (index <= 0) return parts.Length.ToString(CultureInfo.InvariantCulture);
            return index - 1 < parts.Length ? parts[index - 1] : "";
        }

        private static string PathAppend(params string[] parts) => Path.Combine(parts);
        private static string TempFile(string extension) => Path.ChangeExtension(Path.GetTempFileName(), extension);
        private static bool FileExists(string path) => File.Exists(path);
        private static void FileCopy(string source, string dest) => File.Copy(source, dest, true);
        private static void CreatePath(string label, string path) => Directory.CreateDirectory(path);
        private static string FilePath(string fileName) => Path.GetDirectoryName(fileName) ?? "";
        private static string MachineName() => Environment.MachineName;

        private static string FixXml(string s)
        {
            s = s.Replace("&", "&amp;");
            s = s.Replace("'", "&apos;");
            s = s.Replace("\"", "&quot;");
            s = s.Replace("<", "&lt;");
            s = s.Replace(">", "&gt;");
            return s;
        }

        private static Version AppVersion => Assembly.GetEntryAssembly()?.GetName().Version ?? new Version(1, 0, 0, 0);
        private static string AppPath => AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar);
        private static string AppExeName => Path.GetFileNameWithoutExtension(Process.GetCurrentProcess().MainModule?.FileName ?? "app");
        private static string AppProductName => Assembly.GetEntryAssembly()?.GetCustomAttribute<AssemblyProductAttribute>()?.Product ?? "HomeFront";

        private static string DbQuote(DbQuoteType type, object? value, int length = 0)
        {
            switch (type)
            {
                case DbQuoteType.Str:
                    var s = Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
                    if (length > 0 && s.Length > length) s = s[..length];
                    return $"'{s.Replace("'", "''")}'";
                case DbQuoteType.Num:
                    if (value == null || value == DBNull.Value) return "0";
                    if (double.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
                        return d.ToString(CultureInfo.InvariantCulture);
                    return "0";
                case DbQuoteType.Date:
                    if (value is DateTime dt)
                        return $"'{dt:yyyy-MM-dd HH:mm:ss}'";
                    if (DateTime.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), out var parsed))
                        return $"'{parsed:yyyy-MM-dd HH:mm:ss}'";
                    return "NULL";
                default:
                    return "NULL";
            }
        }

        private static string StripFormating(string s)
        {
            const string FORMATCHRS = "+=_-)(*&^%$#@!~`[]{}\\|/?'\"<>;:,";
            foreach (var ch in FORMATCHRS)
            {
                s = s.Replace(ch.ToString(), "");
            }
            s = s.Replace(((char)189).ToString(), "1/2");
            return s;
        }

        private static string FormatPadded(string value, int length)
        {
            value ??= "";
            if (value.Length >= length) return value;
            return value.PadRight(length, '@');
        }

        private static string GetFolderPath(int csidl) =>
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

        private static void ShellAndLoop(string command)
        {
            try
            {
                var process = Process.Start(new ProcessStartInfo
                {
                    FileName = command,
                    CreateNoWindow = true,
                    UseShellExecute = true
                });
                process?.WaitForExit();
            }
            catch
            {
            }
        }

        private static bool IsIn(string value, params string[] options) =>
            options.Any(o => string.Equals(o, value, StringComparison.OrdinalIgnoreCase));

        // ─────────────────────────────────────────────────────────────────────────
        // Public API
        // ─────────────────────────────────────────────────────────────────────────
        public dynamic License => _license;

        public void SendFAX(
            string? SenderName = null, string? SenderCompany = null, string? SenderFaxNumber = null, string? SenderVoiceNumber = null, string? SenderEmailAddress = null,
            string? RecipientName = null, string? RecipientCompany = null, string? RecipientFaxNumber = null, string? RecipientVoiceNumber = null, string? RecipientEmailAddress = null,
            string? Subject = null, string? BillingCode = null,
            string? Attachment1 = null, string? Attachment2 = null, string? Attachment3 = null, string? Attachment4 = null, string? Attachment5 = null,
            string? Attachment6 = null, string? Attachment7 = null, string? Attachment8 = null, string? Attachment9 = null, string? Attachment10 = null)
        {
            var billing = string.IsNullOrWhiteSpace(BillingCode) ? OptionValue("FaxBillingCode") : BillingCode;
            var provider = OptionValue("FaxProvider");

            switch (provider)
            {
                case "GFI FAXmaker":
                    GFIFAXmaker(SenderName, SenderCompany, SenderFaxNumber, SenderVoiceNumber, SenderEmailAddress,
                        RecipientName, RecipientCompany, RecipientFaxNumber, RecipientVoiceNumber, RecipientEmailAddress,
                        Subject, billing, Attachment1, Attachment2, Attachment3, Attachment4, Attachment5, Attachment6, Attachment7, Attachment8, Attachment9, Attachment10);
                    break;
                case "Windows Fax Services":
                    WindowsFaxServices(SenderName, SenderCompany, SenderFaxNumber, SenderVoiceNumber, SenderEmailAddress,
                        RecipientName, RecipientCompany, RecipientFaxNumber, RecipientVoiceNumber, RecipientEmailAddress,
                        Subject, billing, Attachment1, Attachment2, Attachment3, Attachment4, Attachment5, Attachment6, Attachment7, Attachment8, Attachment9, Attachment10);
                    break;
                case "":
                case "none":
                    MsgBox("No fax delivery provider has been configured. Please check your system settings.");
                    break;
                default:
                    throw new NotSupportedException("Unsupported Provider");
            }
        }

        public bool SendMail(bool preview, string sendTo, string ccTo, string subject, string body, string attachments, string? fromAddress = null)
        {
            try
            {
                if (preview && (OptionBool("SendPO_ViaSMTP") || OptionBool("SendPO_Via365SMTP")))
                {
                    if (!EditMailDialog(sendTo, ccTo, subject, body, attachments))
                        return false;
                }

                sendTo = TrimTrailingSemicolon(sendTo);
                ccTo = TrimTrailingSemicolon(ccTo);

                if (string.IsNullOrWhiteSpace(fromAddress))
                {
                    if (OptionBool("SendPO_FromCurrentUsersEmail"))
                        fromAddress = UserEmail;
                    if (string.IsNullOrWhiteSpace(fromAddress))
                        fromAddress = OptionValue("Estimator_Email");
                }

                var xml = BuildMailXml(preview, sendTo, ccTo, subject, body, attachments, fromAddress ?? "");
                var fileName = TempFile("xml");
                File.WriteAllText(fileName, xml, Encoding.Unicode);

                var lastEmailPath = PathAppend(GetFolderPath(0), "HomeFront", "LastEmailMessage.xml");
                CreatePath("", Path.GetDirectoryName(lastEmailPath) ?? "");
                File.WriteAllText(lastEmailPath, xml, Encoding.Unicode);

                var exePath = PathAppend(AppPath, "hfmailer.exe");
                ShellAndLoop($"{exePath} {fileName}");
                return true;
            }
            catch (Exception ex)
            {
                MsgBox("Error sending email" + Environment.NewLine + ex.Message);
                return false;
            }
        }

        private static string TrimTrailingSemicolon(string value)
        {
            value = (value ?? "").Trim();
            while (value.EndsWith(";", StringComparison.Ordinal))
                value = value[..^1];
            return value;
        }

        private string BuildMailXml(bool preview, string sendTo, string ccTo, string subject, string body, string attachments, string from)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<?xml version='1.0' encoding=\"UTF-16\" ?>");
            sb.AppendLine("<xml>");
            sb.AppendLine("  <system>");

            if (OptionBool("SendPO_ViaOutlook"))
            {
                sb.AppendLine("    <type>outlook</type>");
            }
            else if (OptionBool("SendPO_ViaSMTP"))
            {
                sb.AppendLine("    <type>smtp</type>");
                var smtp = OptionValue("SendPO_SMTPServer");
                var server = Parse(smtp, 1, ":");
                var port = Parse(smtp, 2, ":");
                if (string.IsNullOrWhiteSpace(port)) port = OptionValue("SmtpPort");
                sb.AppendLine($"    <server>{FixXml(server)}</server>");
                sb.AppendLine($"    <port>{FixXml(port)}</port>");

                if (OptionBool("Email_Authenticate"))
                {
                    sb.AppendLine($"    <user>{FixXml(OptionValue("Email_User"))}</user>");
                    sb.AppendLine($"    <pswd>{FixXml(OptionValue("Email_Pswd"))}</pswd>");
                }
                if (OptionBool("Email_SSL"))
                    sb.AppendLine("    <ssl>true</ssl>");
                if (OptionBool("Email_STARTTLS"))
                    sb.AppendLine("    <starttls>true</starttls>");

                sb.AppendLine($"    <from>{FixXml(from)}</from>");
                sb.AppendLine($"    <reply>{FixXml(from)}</reply>");
            }
            else if (OptionBool("SendPO_Via365SMTP"))
            {
                sb.AppendLine("    <type>smtp365</type>");
                sb.AppendLine($"    <appid>{FixXml(OptionValue("Smtp365EmailAppID"))}</appid>");
                sb.AppendLine($"    <tenantid>{FixXml(OptionValue("Smtp365EmailTenantID"))}</tenantid>");
                sb.AppendLine($"    <secret>{FixXml(OptionValue("Smtp365EmailSecret"))}</secret>");
                sb.AppendLine($"    <user>{FixXml(OptionValue("Smtp365EmailUser"))}</user>");
                sb.AppendLine($"    <pswd>{FixXml(OptionValue("Smtp365EmailPswd"))}</pswd>");
                sb.AppendLine($"    <from>{FixXml(from)}</from>");
                sb.AppendLine($"    <reply>{FixXml(from)}</reply>");
            }
            else
            {
                sb.AppendLine("    <type>mapi</type>");
            }

            sb.AppendLine("  </system>");
            sb.AppendLine("  <mail>");
            sb.AppendLine($"    <to>{FixXml(sendTo)}</to>");
            sb.AppendLine($"    <cc>{FixXml(ccTo)}</cc>");
            sb.AppendLine($"    <subject>{FixXml(subject)}</subject>");
            sb.AppendLine($"    <body>{FixXml(body)}</body>");
            sb.AppendLine($"    <attachment>{FixXml(attachments)}</attachment>");
            sb.AppendLine($"    <display>{(preview ? "yes" : "no")}</display>");
            sb.AppendLine("  </mail>");
            sb.AppendLine("</xml>");
            return sb.ToString();
        }

        public string DivisionID => int.TryParse(_divisionId, out var v) ? v.ToString(CultureInfo.InvariantCulture) : _divisionId;
        public string LoginPswd => _loginPswd;
        public string LoginID => _loginId;

        public void SetLoginID(string uid, string pwd)
        {
            _loginId = uid ?? "";
            _loginPswd = pwd ?? "";
        }

        public void SetUserLevel(long userLevel) => _userLevel = userLevel;

        public void SetDivision(string divisionId) => SetDivisionID(divisionId);

        public void SetDivisionID(string divisionId) => _divisionId = divisionId ?? "";

        public string LoginDSN => GetConnectionStringValue(ConnectionString(Connections.dbHomefront), "DSN");
        public string LoginDBName => GetConnectionStringValue(ConnectionString(Connections.dbHomefront), "Database");
        public string LoginDBUID => GetConnectionStringValue(ConnectionString(Connections.dbHomefront), "UID");
        public string LoginDBPWD => GetConnectionStringValue(ConnectionString(Connections.dbHomefront), "PWD");

        public void Activations()
        {
        }

        public void About()
        {
            MsgBox("HomeFront");
        }

        public void Logout()
        {
            _closing = true;
            try { Options?.SaveData(); } catch { }

            SetErrorMode(SEM_NOGPFAULTERRORBOX);
            for (var i = 0; i < DbCount; i++) _databases[i] = null;
            HFApp = null;
        }

        public bool Login(string productCode, string productName, string productVersion, string exePath)
        {
            _productName = productName;
            _productCode = productCode;
            _moduleCode = productCode;
            _moduleName = productName;
            _moduleVersion = productVersion;
            _systemVersion = $"{AppVersion.Major}.{AppVersion.Minor}";
            _exePath = (exePath ?? "").TrimEnd('\\');

            try
            {
                CreatePath("working folder", PathAppend(GetFolderPath(0), "Homefront", Parse(_exePath, int.Parse(Parse(_exePath, 0, "\\")), "\\")));
                CreatePath("working folder", PathAppend(GetFolderPath(0), "HomeFront Software Ltd", "HFSales"));
                _license.ReadLicense(LicenseFile);
            }
            catch { }

            if (string.IsNullOrWhiteSpace(_trustedConnection))
            {
                // if (ShowLoginDialog())
                // {
                //      WriteUserSession();
                //      return true;
                //  }
                //  return false;
            }

            try
            {
                ConnectionString(Connections.dbHomefront, _trustedConnection);
                _loginId = string.IsNullOrWhiteSpace(_trustedUser) ? "system" : _trustedUser;
                _loginPswd = "";
                Options.ReadData();
                return true;
            }
            catch (Exception ex)
            {
                MsgBox("Unable to open the Homefront database" + Environment.NewLine + ex.Message);
                return false;
            }
        }

        public string SQLEdition
        {
            get
            {
                if (string.IsNullOrWhiteSpace(_sqlEdition))
                {
                    var dt = SqlExec("SELECT cast(SERVERPROPERTY('edition') as varchar(150))", Connections.dbHomefront);
                    _sqlEdition = dt.Rows.Count > 0 ? Convert.ToString(dt.Rows[0][0]) ?? "" : "";
                }
                return _sqlEdition;
            }
        }

        private void WriteUserSession()
        {
            try
            {
                var s = "delete from UserSessions where LoginDate <= getdate()-14 and WorkstationName=host_name()";
                SqlExec(s);

                s = "insert into UserSessions(Host_ID,LoginName,WorkstationName,LoginDate,HFLoginID)" + Environment.NewLine +
                    "select Host_ID(),System_User,host_name(),GetDate()," + DbQuote(DbQuoteType.Str, _loginId);
                SqlExec(s, Connections.dbHomefront);

                s = "UPDATE UserSessions" + Environment.NewLine +
                    "   SET LoginName=System_User" + Environment.NewLine +
                    "      ,LoginDate=GetDate()" + Environment.NewLine +
                    "      ,HFLoginID=" + DbQuote(DbQuoteType.Str, _loginId) + Environment.NewLine +
                    "WHERE Host_ID=Host_ID() AND WorkstationName=host_name()";
                SqlExec(s, Connections.dbHomefront);
            }
            catch { }
        }

        public string LicenseFile => PathAppend(SystemFolder, "license.lic");

        public string SystemFolder
        {
            get
            {
                var path = RegGetKey(@"SOFTWARE\HomeFront", "HomeFrontLicensePath");
                if (string.IsNullOrWhiteSpace(path)) path = AppPath;
                if (!path.EndsWith(Path.DirectorySeparatorChar)) path += Path.DirectorySeparatorChar;
                return path;
            }
        }

        public string ConnectionString(Connections d) => _connectionStrings[(int)d];

        public void ConnectionString(Connections d, string value) => _connectionStrings[(int)d] = value ?? "";

        public DbWrapperSqlServer Databases(Connections d) => GetDb(d);

        public DataTable SqlExec(string sql, Connections database = Connections.dbHomefront, int recordsAffected = 0)
        {
            sql = ConvertSQL(sql);
            return GetDb(database).SqlExec(sql);
        }

        private static string ConvertSQL(string sql)
        {
            if (string.IsNullOrWhiteSpace(sql)) return sql;
            var pieces = sql.Split('\'');
            for (var i = 0; i < pieces.Length; i += 2)
            {
                var s = pieces[i];
                s = ReplaceIgnoreCase(s, "IFNULL", "ISNULL");
                s = ReplaceIgnoreCase(s, "CURDATE", "GETDATE");
                s = ReplaceIgnoreCase(s, "LONGVARCHAR", "VARCHAR(MAX)");
                s = ReplaceIgnoreCase(s, " DATETIMETIME", " DATETIME");
                s = ReplaceIgnoreCase(s, " DATETIMEDIFF", " DATEDIFF");
                s = ReplaceIgnoreCase(s, " TIMESTAMP", " DATETIME");
                s = ReplaceIgnoreCase(s, "UCASE(", "UPPER(");
                s = ReplaceIgnoreCase(s, "LCASE(", "LOWER(");
                s = ReplaceIgnoreCase(s, "CURTIME(", "GETDATE(");
                s = ReplaceIgnoreCase(s, ", SQL_VARCHAR)", " AS VARCHAR)");
                s = ReplaceIgnoreCase(s, ",SQL_VARCHAR)", " AS VARCHAR)");
                s = ReplaceIgnoreCase(s, ", SQL_INTEGER)", " AS INTEGER)");
                s = ReplaceIgnoreCase(s, ",SQL_INTEGER)", " AS INTEGER)");
                s = ReplaceIgnoreCase(s, "CAST(VARCHAR", "CONVERT(VARCHAR");
                s = ReplaceIgnoreCase(s, "INTEGER IDENTITY_INSERT", "IDENTITY_INSERT");
                pieces[i] = s;
            }
            return string.Join("'", pieces);
        }

        private static string ReplaceIgnoreCase(string input, string search, string replacement) =>
            string.IsNullOrEmpty(input) ? input : System.Text.RegularExpressions.Regex.Replace(input, System.Text.RegularExpressions.Regex.Escape(search), replacement, System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        public long SqlIdentity(string tableName, Connections database = Connections.dbHomefront)
        {
            try
            {
                var dt = SqlExec($"SELECT IDENT_CURRENT('{tableName}')", database);
                return dt.Rows.Count > 0 ? Convert.ToInt64(dt.Rows[0][0]) : 0;
            }
            catch
            {
                return 0;
            }
        }

        public void WriteAuditLog(string task, string description, string? xref1 = null, string? xref2 = null, string? xref3 = null, string? xref4 = null, string? xref5 = null, string? xref6 = null)
        {
            var s = "INSERT INTO UserLog(Workstation,UserID,LogDate,LogTime,Task,Description)" + Environment.NewLine +
                    "VALUES(" + DbQuote(DbQuoteType.Str, MachineName(), 30) + Environment.NewLine +
                    "      ," + DbQuote(DbQuoteType.Str, LoginID, 10) + Environment.NewLine +
                    "      ,GETDATE()" + Environment.NewLine +
                    "      ," + DbQuote(DbQuoteType.Str, DateTime.Now.ToString("hh:mm tt"), 10) + Environment.NewLine +
                    "      ," + DbQuote(DbQuoteType.Str, task, 50) + Environment.NewLine +
                    "      ," + DbQuote(DbQuoteType.Str, description, 200) + ")";
            SqlExec(s);
        }

        public void LockRecord(string tableName, object recordKey)
        {
            var s = "SELECT * FROM TableLog WHERE TableName=" + DbQuote(DbQuoteType.Str, tableName) +
                    " AND KeyRecord=" + DbQuote(DbQuoteType.Str, recordKey);
            var dt = SqlExec(s);
            if (dt.Rows.Count > 0) return;

            s = "INSERT INTO TableLog(TableName,KeyRecord) VALUES(" +
                DbQuote(DbQuoteType.Str, tableName) + "," + DbQuote(DbQuoteType.Str, recordKey) + ")";
            SqlExec(s);

            s = "UPDATE TableLog SET RecordLocked=1,User_ID=" + DbQuote(DbQuoteType.Str, LoginID) +
                ",WorkStation_Name=" + DbQuote(DbQuoteType.Str, MachineName()) +
                ",LockDate=GETDATE() WHERE TableName=" + DbQuote(DbQuoteType.Str, tableName) +
                " AND KeyRecord=" + DbQuote(DbQuoteType.Str, recordKey);
            SqlExec(s);
        }

        public void UnLockRecord(string tableName, object recordKey)
        {
            var s = "DELETE FROM TableLog WHERE TableName=" + DbQuote(DbQuoteType.Str, tableName) +
                    " AND KeyRecord=" + DbQuote(DbQuoteType.Str, recordKey);
            SqlExec(s);
        }

        public string RecordLockedBy(string tableName, object recordKey)
        {
            var s = "SELECT User_ID FROM TableLog WHERE TableName=" + DbQuote(DbQuoteType.Str, tableName) +
                    " AND RecordLocked=1 AND KeyRecord=" + DbQuote(DbQuoteType.Str, recordKey) +
                    " AND User_ID<>" + DbQuote(DbQuoteType.Str, LoginID);
            var dt = SqlExec(s);
            return dt.Rows.Count > 0 ? Convert.ToString(dt.Rows[0][0]) ?? "" : "";
        }

        private string UserEmail
        {
            get
            {
                var dt = SqlExec("SELECT email FROM User_Manager WHERE User_ID=" + DbQuote(DbQuoteType.Str, LoginID), Connections.dbHomefront);
                return dt.Rows.Count > 0 ? Convert.ToString(dt.Rows[0][0]) ?? "" : "";
            }
        }

        public bool UserPermission(string permission)
        {
            var dt = SqlExec("SELECT * FROM User_Manager WHERE User_ID=" + DbQuote(DbQuoteType.Str, LoginID), Connections.dbHomefront);
            var perms = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (dt.Rows.Count > 0)
            {
                var row = dt.Rows[0];
                foreach (DataColumn col in dt.Columns)
                {
                    if (col.DataType == typeof(bool) && row[col] is bool b && b)
                    {
                        perms.Add(col.ColumnName.ToUpperInvariant());
                    }
                }
            }

            return permission switch
            {
                "AdminUser" => _userLevel == 1,
                "SalesManagerUser" => _userLevel == 2,
                "AccountingUser" => _userLevel == 3,
                "SalesPersonUser" => _userLevel == 4,
                "DCSalesPersonUser" => _userLevel == 5,
                "ReadOnlyUser" => _userLevel == 6,
                "EstimatingUser" => _userLevel == 7,
                _ => perms.Contains(permission.ToUpperInvariant())
            };
        }

        public void POInquiry(string? po = null) { }

        public void EditJob(string? job = null) { }

        public void EditVendor(string? vendor = null) { }

        private void EditCustomer(string? customer = null) { }

        public void EditPOIndexList() => RunTask("EditPOIndexes");

        public void EditCommunities() { }

        public void EditProjectManagerList() { }

        private void EditSecurity()
        {
            if (UserPermission("ToolsUsrAdmin"))
                return;
            MsgBox("Your security permissions don't allow you to do that. Sorry.");
        }

        public void EditOptions()
        {
            if (UserPermission("ToolsSysAdmin"))
                return;
            MsgBox("Your security permissions don't allow you to do that. Sorry.");
        }

        public bool ImportData(string tableName, string keyColumns, string title, string? requiredColumns = null, string? columnMappings = null)
        {
            return false;
        }

        public void RunTask(string taskName)
        {
            var task = Parse(taskName, 1, "|");
            var p1 = Parse(taskName, 2, "|");
            var p2 = Parse(taskName, 3, "|");
            var p3 = Parse(taskName, 4, "|");

            switch (task)
            {
                case "SendBuildProJobs":
                    SendBuildProJobs(p1);
                    break;
                case "CancelBuildProPO":
                    CancelBuildProPO(p1, p2);
                    break;
                case "SendBuildProPOs":
                    SendBuildProPOs(p1, p2);
                    break;
                case "SendBuildProPOIndexes":
                    SendBuildProPOIndexes();
                    break;
                case "SendBuildProVendors":
                    SendBuildProVendors();
                    break;
                case "SendBuildProCommunities":
                    SendBuildProCommunities();
                    break;
                case "SendBuildProRemittances":
                    SendBuildProVouchers();
                    SendBuildProChecks();
                    break;
                case "SendBuildProChecks":
                    SendBuildProChecks();
                    break;
                case "SendBuildProVouchers":
                    SendBuildProVouchers();
                    break;
                case "EditOptions":
                    EditOptions();
                    break;
                case "EditSecurity":
                    EditSecurity();
                    break;
                case "EditJCCategories":
                    EditJCCategories();
                    break;
                case "EditJCCostCodes":
                    EditJCCostCodes();
                    break;
                case "POInquiry":
                    POInquiry(p1);
                    break;
                case "EditCustomer":
                    EditCustomer(p1);
                    break;
                case "EditVendor":
                    EditVendor(p1);
                    break;
                case "EditJob":
                    EditJob(p1);
                    break;
                case "EditProjectManagerList":
                    EditProjectManagerList();
                    break;
                case "EditCommunities":
                    EditCommunities();
                    break;
                case "WriteJobToAccounting":
                    WriteJobToAccounting(p1);
                    break;
                case "WriteCustomerToAccounting":
                    WriteCustomerToAccounting(p1, p2);
                    break;
                case "TrustedLogin":
                    _trustedConnection = p1;
                    _trustedSkipAccounting = p2 == "1";
                    _trustedUser = p3;
                    break;
                case "ExportData":
                    ExportData(p1, p2);
                    break;
                default:
                    MsgBox($"Unable to execute task \"{taskName}\" - it has not been implemented in this version.");
                    break;
            }
        }

        private void EditJCCostCodes()
        {
            // UI grid not implemented
        }

        private void EditJCCategories()
        {
            // UI grid not implemented
        }

        // VB6: Public Function Encrypt(Text As String) As String
        //         Set hff = CreateObject("ZYBFunctions.HomeFrontFunctions")
        //         Encrypt = hff.MyCrypt(Text, "Fazlul")
        public string Encrypt(string text)
        {
            return FLogin.MyCrypt(text, "Fazlul");
        }

        public string Decrypt(string text)
        {
            return FLogin.MyDeCrypt(text, "Fazlul");
        }

        // ─────────────────────────────────────────────────────────────────────────
        // External System Stubs (QuickBooks, Sage, etc.)
        // ─────────────────────────────────────────────────────────────────────────
        private void WriteJobToSage300(string jobNumber) { }
        private void WriteJobToSage100(string jobNumber) { }
        private void WriteCustomerToQuickBooks(string jobNumber, string customerNumber) { }
        private void WriteCustomerToQuickBooksMF(string jobNumber, string customerNumber) { }
        private void WriteCustomerToQuickBooksSF(string jobNumber, string customerNumber) { }
        private string getCustEditSeq(string listId) => "";
        private void WriteCustomerToSage100(string jobNumber, string customerNumber) { }
        private void WriteCustomerToSage50(string jobNumber, object customerNumber, bool arCustomers = false) { }
        private void WriteJobToSage50(string jobNumber) { }
        private void WriteCustomerToIntacct(string jobNumber, string customerNumber) { }
        private void WriteJobToIntacct(string jobNumber) { }
        private void WriteCustomerToQBO(string jobNumber, string customerNumber) { }

        public void ExportData(string selectStatement, string dialogTitle) { }

        public int EstFieldSize(string field) => field switch
        {
            "Phase" => 20,
            "PhaseDesc" => 150,
            "Item" => 15,
            "ItemDesc" => 150,
            "ItemNotes" => 1520,
            "AssemblyDesc" => 70,
            "AssemblyNotes" => 1520,
            _ => throw new ArgumentOutOfRangeException(nameof(field))
        };

        public string FormatPhase(string phase)
        {
            if (OptionInt("EstimatingMajorVersion") >= 13)
                return phase?.Trim() ?? "";
            return FormatPadded(phase?.Trim() ?? "", EstFieldSize("Phase"));
        }

        public string FormatItem(string item)
        {
            if (string.Equals(item, "AppVersion", StringComparison.OrdinalIgnoreCase))
                return $"{AppVersion.Major}.{AppVersion.Minor}.{AppVersion.Build}";
            if (string.Equals(item, "TitlebarCaption", StringComparison.OrdinalIgnoreCase))
            {
                var divCodeDt = SqlExec("Select DivisionCode from Divisions where DivisionID = " + DivisionID, Connections.dbHomefront);
                var divCode = divCodeDt.Rows.Count > 0 ? Convert.ToString(divCodeDt.Rows[0][0]) ?? "" : "";
                return $" ({LoginDSN}, Login User: {LoginID}) - version {AppVersion.Major}.{AppVersion.Minor} Division: {divCode}";
            }
            if (OptionInt("EstimatingMajorVersion") >= 13)
                return item?.Trim() ?? "";
            return FormatPadded(item?.Trim() ?? "", EstFieldSize("Item"));
        }

        public string FormatJob(string job)
        {
            return OptionInt("AccountingSystem") == (int)AccountingSystems.asTimberline
                ? string.Format(CultureInfo.InvariantCulture, "{0}", StripFormating(job ?? ""))
                : job ?? "";
        }

        public string FormatCostCode(string costCode)
        {
            return OptionInt("AccountingSystem") switch
            {
                (int)AccountingSystems.asTimberline => StripFormating(costCode ?? ""),
                (int)AccountingSystems.asMasterBuilder => string.Format(CultureInfo.InvariantCulture, "{0:0.000}", Convert.ToDecimal(costCode)),
                _ => costCode ?? ""
            };
        }

        public string XmlMbStart(string company, string userName) => "";
        public string XmlMBEnd() => "";
        public string XmlMBAdd(XMLFieldTypes fieldType, double size, string name, object value) => "";
        public bool XmlMbSubmit(string xml, string pswd) => false;

        public string XmlQBSubmit(string xml) => "";
        public string XmlQBStart() => "";
        public string XmlQBEnd() => "";
        public string XmlQBAdd(XMLFieldTypes fieldType, double size, string name, object value)
        {
            var s = Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
            s = s.Replace(((char)188).ToString(), "1/4")
                 .Replace(((char)189).ToString(), "1/2")
                 .Replace(((char)190).ToString(), "3/4")
                 .Replace(((char)150).ToString(), "-")
                 .Replace(((char)145).ToString(), "'")
                 .Replace(((char)146).ToString(), "'")
                 .Replace(((char)147).ToString(), "\"")
                 .Replace(((char)148).ToString(), "\"")
                 .Replace(((char)152).ToString(), "\"")
                 .Replace("&lt;", "<")
                 .Replace("&gt;", ">");
            return s;
        }

        public void LogInfoToFile(string filename, string s)
        {
            try
            {
                CreatePath("", FilePath(filename));
                File.AppendAllText(filename, new string('=', 50) + Environment.NewLine + s + Environment.NewLine);
            }
            catch { }
        }

        public bool AcquireLicense()
        {
            if (!FileExists(LicenseFile))
            {
                MsgBox("Access denied.\n\nNo license could be found.");
                return false;
            }

            try
            {
                if (!_license.ReadLicense(LicenseFile))
                {
                    MsgBox("Access denied.\n\nInvalid license. Please contact your HomeFront solution provider for assistance.");
                    return false;
                }

                if (_license.Leased && _license.LeaseExpiryDate < DateTime.Today)
                {
                    MsgBox("Your software lease has expired. Please contact your HomeFront solution provider for assistance.");
                    return false;
                }

                if (_jobSimplicity) _productCode = "JobSimplicity";

                if (_license.ModuleDemo(_license.ModuleIndex(_productCode)))
                {
                    var daysLeft = _license.ModuleDemoExpiry(_license.ModuleIndex(_productCode)) - DateTime.Today.Day;
                    if (daysLeft < 1)
                    {
                        MsgBox("The evaluation period has expired. Please contact your HomeFront solution provider for assistance.");
                        return false;
                    }
                }

                _productName = _license.ModuleName(_license.ModuleIndex(_productCode));
                var seats = _license.ModuleSeats(_license.ModuleIndex(_productCode));
                if (_productCode == "HFStaging" || _productCode == "HFSalesMgmt" || _productCode == "HFSynch")
                    seats = 9999;

                string msg;
                var inUse = _license.ModuleSeatsInUse(_license.ModuleIndex(_productCode), out msg);
                if (seats > inUse) return true;

                if (string.IsNullOrEmpty(msg))
                    MsgBox($"All {seats} licenses are in use.\n\nPlease contact your HomeFront solution provider if you require additional seats.");
                else
                    MsgBox("SQL Server configuration error.\nConnecting user requires VIEW SERVER STATE. Please contact your system administrator.");
            }
            catch
            {
                MsgBox("Access denied.");
            }

            return false;
        }

        public string LicenseNumber => Convert.ToString(_license.SerialNumber) ?? "";
        public string ClientID => Convert.ToString(_license.ClientID) ?? "";
        public string LicensedOwner => Convert.ToString(_license.CompanyName) ?? "";
        //public long LicensesFree => _license.ModuleSeats(_license.ModuleIndex(_productCode)) - _license.ModuleSeatsInUse(_license.ModuleIndex(_productCode), out _);
        public long LicensedUsers => _license.ModuleSeats(_license.ModuleIndex(_productCode));
        public long LicensedSeats(string productCode) => _license.ModuleSeats(_license.ModuleIndex(productCode));

        public string LicenseLimit => "";

        public void WriteCustomerToAccounting(string jobNumber, string customerNumber)
        {
            switch ((AccountingSystems)OptionInt("AccountingSystem"))
            {
                case AccountingSystems.asTimberline:
                    break;
                case AccountingSystems.asIntacct:
                    WriteCustomerToIntacct(jobNumber, customerNumber);
                    break;
                case AccountingSystems.asMasterBuilder:
                    WriteCustomerToSage100(jobNumber, customerNumber);
                    break;
                case AccountingSystems.asQuickBooks:
                    WriteCustomerToQuickBooks(jobNumber, customerNumber);
                    break;
                case AccountingSystems.asQuickbooksOnline:
                    WriteCustomerToQBO(jobNumber, customerNumber);
                    break;
                case AccountingSystems.asSimply:
                    WriteCustomerToSage50(jobNumber, customerNumber, string.IsNullOrWhiteSpace(jobNumber));
                    break;
            }
        }

        public void WriteJobToAccounting(string jobNumber)
        {
            var dt = SqlExec("select isquote from tbljobs where job_no=" + DbQuote(DbQuoteType.Str, jobNumber), Connections.dbHomefront);
            if (dt.Rows.Count > 0 && string.Equals(Convert.ToString(dt.Rows[0][0]), "True", StringComparison.OrdinalIgnoreCase))
                return;

            WriteCustomerToAccounting(jobNumber, "");

            switch ((AccountingSystems)OptionInt("AccountingSystem"))
            {
                case AccountingSystems.asTimberline:
                    WriteJobToSage300(jobNumber);
                    break;
                case AccountingSystems.asIntacct:
                    WriteJobToIntacct(jobNumber);
                    break;
                case AccountingSystems.asMasterBuilder:
                    WriteJobToSage100(jobNumber);
                    break;
                case AccountingSystems.asSimply:
                    WriteJobToSage50(jobNumber);
                    break;
            }
        }

        public void OpenSimplyODBC()
        {
            MsgBox("Simply ODBC connection not implemented in Blazor.");
        }

        public void CloseSimplyODBC()
        {
            _simplyOdbcUtilities = null;
        }

        public bool CreateDivisionUsers()
        {
            try
            {
                var s = "CREATE TABLE [dbo].[DivisionUsers](" +
                        "   [DivisionID] [int] NOT NULL," +
                        "   [USERID] [varchar](10) NOT NULL," +
                        " CONSTRAINT [PK_DivisionUsers] PRIMARY KEY CLUSTERED " +
                        "([DivisionID] ASC, [USERID] ASC))";
                SqlExec(s);

                s = "insert into DivisionUsers(DivisionID,UserID) Select d.divisionid,u.user_id from user_manager u join divisions d on(1=1)";
                SqlExec(s);
                return true;
            }
            catch
            {
                return true;
            }
        }

        // ─────────────────────────────────────────────────────────────────────────
        // External Provider Stubs
        // ─────────────────────────────────────────────────────────────────────────
        private void GFIFAXmaker(params object?[] _)
        {
        }

        private void WindowsFaxServices(params object?[] _)
        {
        }

        private bool EditMailDialog(string sendTo, string ccTo, string subject, string body, string attachments) => true;

        private void SendBuildProJobs(string job) { }
        private void CancelBuildProPO(string job, string po) { }
        private void SendBuildProPOs(string job, string pos) { }
        private void SendBuildProPOIndexes() { }
        private void SendBuildProVendors() { }
        private void SendBuildProCommunities() { }
        private void SendBuildProVouchers() { }
        private void SendBuildProChecks() { }

        private string GetConnectionStringValue(string connectionString, string key)
        {
            if (string.IsNullOrWhiteSpace(connectionString)) return "";
            var parts = connectionString.Split(';', StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts)
            {
                var kv = part.Split('=', 2);
                if (kv.Length == 2 && kv[0].Trim().Equals(key, StringComparison.OrdinalIgnoreCase))
                    return kv[1].Trim();
            }
            return "";
        }

        private string RegGetKey(string subKey, string valueName)
        {
            try
            {
                using var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(subKey);
                return key?.GetValue(valueName)?.ToString() ?? "";
            }
            catch
            {
                return "";
            }
        }

        private sealed class LicenseStub
        {
            public bool ReadLicense(string file) => false;
            public bool Leased => false;
            public DateTime LeaseExpiryDate => DateTime.MaxValue;
            public bool ModuleDemo(int idx) => false;
            public int ModuleDemoExpiry(int idx) => 0;
            public string ModuleName(int idx) => "";
            public int ModuleIndex(string code) => 0;
            public int ModuleSeats(int idx) => 0;
            public int ModuleSeatsInUse(int idx, out string msg) { msg = ""; return 0; }
            public string SerialNumber => "";
            public string ClientID => "";
            public string CompanyName => "";
        }

        private sealed class OptionsStub
        {
            private readonly Dictionary<string, string> _values = new(StringComparer.OrdinalIgnoreCase);
            public string ValueByName(string name) => _values.TryGetValue(name, out var v) ? v : "";
            public string Value(string name) => ValueByName(name);
            public string this[string name] { get => ValueByName(name); set => _values[name] = value; }
            public void ReadData() { }
            public void SaveData() { }
        }
    }
}
