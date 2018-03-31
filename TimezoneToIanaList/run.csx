#r ".\TimeZoneNames.dll"

using System.Net;
using System.Linq;
using TimeZoneNames;

public async static Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    log.Info("C# HTTP trigger function processed a request.");

    // parse query parameter
    string name = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "name", true) == 0)
        .Value;

    var res = string.IsNullOrEmpty(name) ? Enumerable.Empty<string>().ToArray(): TZNames.GetIdsForCldrName(name);
    return req.CreateResponse(HttpStatusCode.OK, res);
}
