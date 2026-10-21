/* Task Clock toolbar commands: Clock In / Clock Out run directly from the event's toolbar.
   State lives in the add-in's roaming settings (per mailbox, synced across Outlook clients).
   Clock Out opens a pre-filled "Worked:" appointment; the user clicks Save. */
(function () {
  "use strict";

  var KEY_CLOCK = "clock";      // { title, taskId, link, start }
  var KEY_LAST = "lastSession"; // { title, taskId, link, start, end }
  var NOTE_ID = "taskclock";

  function fmtTime(d) {
    return d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  }

  function short(s, n) {
    s = s || "";
    return s.length > n ? s.slice(0, n - 1) + "…" : s;
  }

  function notify(text, done) {
    var item = Office.context.mailbox && Office.context.mailbox.item;
    if (!item || !item.notificationMessages) return done();
    item.notificationMessages.replaceAsync(NOTE_ID, {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: short(text, 150),
      icon: "icon16",
      persistent: false
    }, function () { done(); });
  }

  function decodeEntities(s) {
    return (s || "").replace(/&amp;/g, "&").replace(/&quot;/g, "\"").replace(/&#39;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">");
  }

  function parseBody(html) {
    var out = { taskId: "", link: "" };
    var m = /PlannerTask:([A-Za-z0-9_-]{28})/.exec(html || "");
    if (m) out.taskId = m[1];
    var l = /href=["'](https:\/\/planner\.cloud\.microsoft[^"']+)["']/i.exec(html || "");
    if (l) out.link = decodeEntities(l[1]);
    return out;
  }

  function cleanTitle(subject) {
    return (subject || "").replace(/^\s*Fire:\s*/i, "").trim() || "(untitled event)";
  }

  function readItem(cb) {
    var item = Office.context.mailbox.item;
    var subjectDone = function (subject) {
      if (!item.body || !item.body.getAsync) return cb(subject, "");
      item.body.getAsync(Office.CoercionType.Html, function (r) {
        cb(subject, r.status === Office.AsyncResultStatus.Succeeded ? r.value : "");
      });
    };
    if (typeof item.subject === "string") return subjectDone(item.subject);
    if (item.subject && item.subject.getAsync) {
      return item.subject.getAsync(function (r) {
        subjectDone(r.status === Office.AsyncResultStatus.Succeeded ? r.value : "");
      });
    }
    subjectDone("");
  }

  function save(cb) {
    Office.context.roamingSettings.saveAsync(function (r) {
      cb(r.status === Office.AsyncResultStatus.Succeeded, r.error && r.error.message);
    });
  }

  function clockIn(event) {
    var rs = Office.context.roamingSettings;
    var finish = function () { event.completed(); };
    var c = rs.get(KEY_CLOCK);
    if (c) {
      notify("Already clocked in to “" + short(c.title, 60) + "” since " + fmtTime(new Date(c.start)) + ". Clock out first.", finish);
      return;
    }
    readItem(function (subject, html) {
      var p = parseBody(html);
      var title = cleanTitle(subject);
      var start = new Date();
      rs.set(KEY_CLOCK, { title: title, taskId: p.taskId, link: p.link, start: start.toISOString() });
      save(function (ok, err) {
        if (!ok) return notify("Couldn't clock in: " + (err || "save failed"), finish);
        notify("Clocked in at " + fmtTime(start) + ": " + short(title, 90), finish);
      });
    });
  }

  function clockOut(event) {
    var rs = Office.context.roamingSettings;
    var finish = function () { event.completed(); };
    var c = rs.get(KEY_CLOCK);
    if (!c) {
      notify("You're not clocked in.", finish);
      return;
    }
    var start = new Date(c.start);
    var end = new Date();
    if (end - start < 60000) end = new Date(start.getTime() + 60000);
    var mins = Math.max(1, Math.round((end - start) / 60000));
    var session = { title: c.title, taskId: c.taskId, link: c.link, start: start.toISOString(), end: end.toISOString() };
    rs.set(KEY_LAST, session);
    rs.remove(KEY_CLOCK);
    save(function () {
      var lines = [
        "Time worked on Planner task: " + c.title,
        fmtTime(start) + " – " + fmtTime(end) + " (" + mins + " min)"
      ];
      if (c.link) lines.push("Open the task in Planner: " + c.link);
      if (c.taskId) lines.push("TimeLog:" + c.taskId);
      try {
        Office.context.mailbox.displayNewAppointmentForm({
          requiredAttendees: [],
          optionalAttendees: [],
          start: start,
          end: end,
          location: "",
          resources: [],
          subject: "Worked: " + c.title,
          body: lines.join("\n")
        });
      } catch (e) {
        return notify("Clocked out " + fmtTime(start) + "–" + fmtTime(end) + " (" + mins + " min), but the entry didn't open: " + e.message, finish);
      }
      notify("Clocked out: " + fmtTime(start) + "–" + fmtTime(end) + " (" + mins + " min). Click Save on the entry that opened.", finish);
    });
  }

  window.clockIn = clockIn;
  window.clockOut = clockOut;

  Office.onReady(function () {
    if (Office.actions && Office.actions.associate) {
      Office.actions.associate("clockIn", clockIn);
      Office.actions.associate("clockOut", clockOut);
    }
  });
})();
