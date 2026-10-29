/* Task Clock sign-in and Planner calls.
   Uses Outlook's built-in sign-in (nested app authentication) so Task Clock acts as the person
   using Outlook. Needs msal-browser loaded first (global `msal`).
   Planner task chat is a Microsoft Graph preview (beta) API and may change. */
(function () {
  "use strict";

  var CLIENT_ID = "a1f45e77-9af1-4174-abc9-b65603fad236";
  var TENANT = "1324f058-3838-4c96-a4cf-7ea32c3d191e";
  var SCOPES = ["Tasks.ReadWrite"];
  var GRAPH = "https://graph.microsoft.com";
  var pca = null;

  function available() {
    try {
      return !!(window.msal && msal.createNestablePublicClientApplication &&
        Office.context.requirements.isSetSupported("NestedAppAuth", "1.1"));
    } catch (e) {
      return false;
    }
  }

  function init() {
    if (pca) return Promise.resolve(pca);
    return msal.createNestablePublicClientApplication({
      auth: { clientId: CLIENT_ID, authority: "https://login.microsoftonline.com/" + TENANT },
      cache: { cacheLocation: "localStorage" }
    }).then(function (p) { pca = p; return p; });
  }

  function getToken() {
    if (!available()) return Promise.reject(new Error("This version of Outlook doesn't support Task Clock sign-in."));
    return init().then(function (p) {
      var req = { scopes: SCOPES };
      return p.acquireTokenSilent(req).catch(function () {
        return p.acquireTokenPopup(req);
      });
    }).then(function (r) { return r.accessToken; });
  }

  function graph(method, path, body, extraHeaders) {
    return getToken().then(function (token) {
      var headers = { Authorization: "Bearer " + token };
      if (body !== undefined) headers["Content-Type"] = "application/json";
      for (var k in (extraHeaders || {})) headers[k] = extraHeaders[k];
      return fetch(GRAPH + path, {
        method: method,
        headers: headers,
        body: body === undefined ? undefined : JSON.stringify(body)
      });
    }).then(function (res) {
      if (res.ok) return res.status === 204 ? null : res.json().catch(function () { return null; });
      return res.text().then(function (t) {
        var msg = t;
        try { msg = JSON.parse(t).error.message || t; } catch (e) { /* keep text */ }
        throw new Error(res.status + " " + msg);
      });
    });
  }

  function escapeHtml(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  // Posts a message to the task's chat as the signed-in person.
  function postChat(taskId, text) {
    return graph("POST", "/beta/planner/tasks/" + encodeURIComponent(taskId) + "/messages",
      { content: escapeHtml(text) });
  }

  // Marks the task complete. Resolves true if it changed, false if it was already complete.
  function completeTask(taskId) {
    var path = "/v1.0/planner/tasks/" + encodeURIComponent(taskId);
    return graph("GET", path).then(function (t) {
      if (t && t.percentComplete === 100) return false;
      return graph("PATCH", path, { percentComplete: 100 },
        { "If-Match": t["@odata.etag"], Prefer: "return=minimal" }).then(function () { return true; });
    });
  }

  window.TaskClockAuth = { available: available, getToken: getToken, postChat: postChat, completeTask: completeTask };
})();
