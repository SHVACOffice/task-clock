/* Task Clock sign-in and Planner calls.
   Uses Outlook's built-in sign-in (nested app authentication) so Task Clock acts as the person
   using Outlook. Needs msal-browser loaded first (global `msal`).
   Planner task chat is a Microsoft Graph preview (beta) API and may change. */
(function () {
  "use strict";

  var CLIENT_ID = "a1f45e77-9af1-4174-abc9-b65603fad236";
  var TENANT = "1324f058-3838-4c96-a4cf-7ea32c3d191e";
  var SCOPES = ["Tasks.ReadWrite"];
  var CAL_SCOPES = ["Calendars.ReadWrite"];
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

  // silentOnly: never show a sign-in window (used where a pop-up would be in the way).
  function getToken(scopes, silentOnly) {
    if (!available()) return Promise.reject(new Error("This version of Outlook doesn't support Task Clock sign-in."));
    return init().then(function (p) {
      var req = { scopes: scopes || SCOPES };
      return p.acquireTokenSilent(req).catch(function (e) {
        if (silentOnly) throw e;
        return p.acquireTokenPopup(req);
      });
    }).then(function (r) { return r.accessToken; });
  }

  function graph(method, path, body, extraHeaders, scopes, silentOnly) {
    return getToken(scopes, silentOnly).then(function (token) {
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

  function linesToHtml(lines) {
    return lines.map(function (l) {
      return escapeHtml(l).replace(/(https:\/\/[^\s<]+)/g, function (u) {
        return '<a href="' + u.replace(/"/g, "&quot;") + '">' + u + "</a>";
      });
    }).join("<br>");
  }

  function graphTime(d) {
    return { dateTime: new Date(d).toISOString().replace("Z", ""), timeZone: "UTC" };
  }

  // Saves a calendar entry straight to the signed-in person's calendar, without opening a form.
  function createEvent(opts) {
    return graph("POST", "/v1.0/me/events", {
      subject: opts.subject,
      body: { contentType: "HTML", content: linesToHtml(opts.lines) },
      start: graphTime(opts.start),
      end: graphTime(opts.end),
      showAs: "busy",
      isReminderOn: false
    }, null, CAL_SCOPES, true);
  }

  // Saves the entry directly; if that isn't possible, opens the filled-in form so nothing is lost.
  // Resolves "saved" or "form"; rejects only if the form couldn't open either.
  function saveEntry(opts) {
    function openForm() {
      Office.context.mailbox.displayNewAppointmentForm({
        requiredAttendees: [], optionalAttendees: [], location: "", resources: [],
        start: new Date(opts.start), end: new Date(opts.end), subject: opts.subject, body: opts.lines.join("\n")
      });
      return "form";
    }
    if (opts.formOnly) return Promise.resolve().then(openForm);
    return createEvent(opts).then(function () { return "saved"; }, function () { return openForm(); });
  }

  // Deletes this task's planned calendar spots (events carrying "PlannerTask:<id>") that start
  // between `from` and `to`. Resolves with the number removed.
  function removeSpots(taskId, from, to) {
    var marker = "PlannerTask:" + taskId;
    var path = "/v1.0/me/calendarView?startDateTime=" + encodeURIComponent(new Date(from).toISOString()) +
      "&endDateTime=" + encodeURIComponent(new Date(to).toISOString()) + "&$select=id,body&$top=100";
    var removed = 0;
    function page(p) {
      return graph("GET", p, undefined, { Prefer: 'outlook.body-content-type="text"' }, CAL_SCOPES, true).then(function (r) {
        var hits = (r && r.value || []).filter(function (ev) {
          return ev.body && ev.body.content && ev.body.content.indexOf(marker) !== -1;
        });
        return hits.reduce(function (prev, ev) {
          return prev.then(function () {
            return graph("DELETE", "/v1.0/me/events/" + encodeURIComponent(ev.id), undefined, null, CAL_SCOPES, true)
              .then(function () { removed++; }, function () { /* leave it for the scheduler */ });
          });
        }, Promise.resolve()).then(function () {
          var next = r && r["@odata.nextLink"];
          return next ? page(next.replace(GRAPH, "")) : removed;
        });
      });
    }
    return page(path);
  }

  window.TaskClockAuth = { removeSpots: removeSpots, available: available, getToken: getToken, postChat: postChat, completeTask: completeTask,
    createEvent: createEvent, saveEntry: saveEntry };
})();
