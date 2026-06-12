using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Services;
using System.Web.Script.Services;
using System.Web.Script.Serialization;

// Consolidated code-behind for ASMDaily.aspx.
//
// Two server-side endpoints live on this single page:
//   1) GetWaferCounts  - ASP.NET page method (called as ASMDaily.aspx/GetWaferCounts).
//                        Reads W/C from SQL Server using the GPTDB_EAS connection string.
//   2) ?op=chat        - AI gateway proxy (merged from the old AiChat.aspx.cs).
//                        Keeps the api-key on the server; the browser never sees it.
public partial class ASMDaily : System.Web.UI.Page
{
    // ======================= AI chat proxy =======================
    // Handles ASMDaily.aspx?op=chat. Any request without ?op falls through
    // to the normal page lifecycle and renders the dashboard UI.
    protected void Page_Load(object sender, EventArgs e)
    {
        string op = Request.QueryString["op"];
        if (string.IsNullOrEmpty(op)) return; // fall through to the ASPX UI
        Response.ContentType = "application/json; charset=utf-8";
        Response.Cache.SetCacheability(HttpCacheability.NoCache);
        try
        {
            if (string.Equals(op, "chat", StringComparison.OrdinalIgnoreCase))
            {
                HandleChat();
            }
            else
            {
                Response.StatusCode = 400;
                Response.Write("{\"ok\":false,\"error\":\"unknown op\"}");
            }
        }
        catch (Exception ex)
        {
            Response.StatusCode = 500;
            Response.Write("{\"ok\":false,\"error\":\"" + JsonEscape(ex.Message) + "\"}");
        }
        Response.End();
    }

    private void HandleChat()
    {
        string url = ConfigurationManager.AppSettings["AiGatewayUrl"];
        string apiKey = ConfigurationManager.AppSettings["AiApiKey"];
        string userId = ConfigurationManager.AppSettings["AiUserId"];
        string systemPrompt = ConfigurationManager.AppSettings["AiSystemPrompt"];
        // Fallback prompt is ASCII to keep this source file encoding-safe.
        // For Chinese / domain-specific prompts, set AiSystemPrompt in web.config.
        if (string.IsNullOrEmpty(systemPrompt)) systemPrompt = "You are a helpful assistant.";

        if (string.IsNullOrEmpty(url) || string.IsNullOrEmpty(apiKey))
        {
            Response.StatusCode = 500;
            Response.Write("{\"ok\":false,\"error\":\"AiGatewayUrl / AiApiKey not configured in web.config\"}");
            return;
        }

        string body = ReadBody();
        var ser = NewSerializer();
        Dictionary<string, object> clientReq;
        try { clientReq = ser.Deserialize<Dictionary<string, object>>(body) ?? new Dictionary<string, object>(); }
        catch
        {
            Response.StatusCode = 400;
            Response.Write("{\"ok\":false,\"error\":\"invalid json\"}");
            return;
        }

        // Prepend the server-controlled system prompt, then forward whatever
        // the client sent.  Client cannot override the system prompt because
        // the server always inserts its own first.
        var messages = new System.Collections.ArrayList();
        messages.Add(new Dictionary<string, object> {
            { "role", "system" },
            { "content", systemPrompt }
        });
        object clientMessages;
        if (clientReq.TryGetValue("messages", out clientMessages) && clientMessages is System.Collections.ArrayList)
        {
            foreach (var m in (System.Collections.ArrayList)clientMessages)
            {
                if (m != null) messages.Add(m);
            }
        }
        var payload = new Dictionary<string, object> { { "messages", messages } };
        byte[] payloadBytes = Encoding.UTF8.GetBytes(ser.Serialize(payload));

        var req = (HttpWebRequest)WebRequest.Create(url);
        req.Method = "POST";
        req.Accept = "*/*";
        req.ContentType = "application/json";
        req.Headers["api-key"] = apiKey;
        if (!string.IsNullOrEmpty(userId)) req.Headers["user-id"] = userId;
        req.Timeout = 120000;
        req.ReadWriteTimeout = 120000;
        req.ContentLength = payloadBytes.Length;
        try
        {
            using (var s = req.GetRequestStream()) s.Write(payloadBytes, 0, payloadBytes.Length);
            using (var resp = (HttpWebResponse)req.GetResponse())
            using (var sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
            {
                Response.Write(sr.ReadToEnd());
            }
        }
        catch (WebException wex)
        {
            string detail = "";
            int status = 502;
            var httpResp = wex.Response as HttpWebResponse;
            if (httpResp != null)
            {
                status = (int)httpResp.StatusCode;
                try
                {
                    using (var sr = new StreamReader(httpResp.GetResponseStream(), Encoding.UTF8))
                        detail = sr.ReadToEnd();
                }
                catch { }
            }
            Response.StatusCode = status;
            Response.Write("{\"ok\":false,\"error\":\"" + JsonEscape(wex.Message) + "\",\"detail\":" + ser.Serialize(detail) + "}");
        }
    }

    private string ReadBody()
    {
        using (var reader = new StreamReader(Request.InputStream, Encoding.UTF8))
            return reader.ReadToEnd();
    }
    private static JavaScriptSerializer NewSerializer()
    {
        return new JavaScriptSerializer { MaxJsonLength = 200 * 1024 * 1024 };
    }
    private static string JsonEscape(string s)
    {
        if (s == null) return "";
        return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "");
    }

    // ======================= W/C page method =======================
    // Server: UMCESIDB02   Table: [GPTDB_EAS].[dbo].[XSITEUSAGEMETER]
    // Connection string lives in web.config <connectionStrings name="GPTDB_EAS">.
    [WebMethod]
    [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
    public static Dictionary<string, object> GetWaferCounts()
    {
        var result = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        var connStrSettings = System.Configuration.ConfigurationManager.ConnectionStrings["GPTDB_EAS"];

        if (connStrSettings == null || string.IsNullOrWhiteSpace(connStrSettings.ConnectionString))
            throw new Exception("Missing connection string: GPTDB_EAS");

        var connStr = connStrSettings.ConnectionString;

        var sql = @"
      SELECT EQPID, DATA_VAL
      FROM [dbo].[XSITEUSAGEMETER]
      WHERE METERTYPE = 'CH_WAFER_COUNT'
    ";

        using (var conn = new SqlConnection(connStr))
        using (var cmd = new SqlCommand(sql, conn))
        {
            conn.Open();
            using (var rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    var eqpid = (rdr["EQPID"] == DBNull.Value) ? "" : rdr["EQPID"].ToString().Trim();
                    if (string.IsNullOrEmpty(eqpid)) continue;

                    object val = (rdr["DATA_VAL"] == DBNull.Value) ? null : rdr["DATA_VAL"];
                    result[eqpid] = val;
                }
            }
        }

        return result;
    }
}
