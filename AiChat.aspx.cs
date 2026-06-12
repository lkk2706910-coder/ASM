using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Script.Serialization;

// Standalone AI chat proxy.
//
// Reads LLM endpoint + api-key from web.config <appSettings> so the key
// never leaves the server.  The client posts { messages: [...] } to
// ?op=chat and gets back the LLM's verbatim response.  Drop this folder
// into any IIS site, fill in the AppSettings, and you have a working
// AI assistant page at AiChat.aspx.
public partial class AiChat : System.Web.UI.Page
{
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
}
