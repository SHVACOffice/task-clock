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
    var m = /(?:PlannerTask|TimeLog|TaskDone):([A-Za-z0-9_-]{28})/.exec(html || "");
    if (m) out.taskId = m[1];
    var l = /href=["'](https:\/\/planner\.cloud\.microsoft[^"']+)["']/i.exec(html || "") ||
            /(https:\/\/planner\.cloud\.microsoft[^\s"'<]+)/i.exec(html || "");
    if (l) out.link = decodeEntities(l[1]);
    return out;
  }

  function cleanTitle(subject) {
    return (subject || "").replace(/^\s*(Fire|Worked|Done):\s*/i, "").trim() || "(untitled event)";
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

  function sameTask(c, t) {
    if (c.fire || t.fire) return !!c.fire && !!t.fire;
    if (c.taskId || t.taskId) return c.taskId === t.taskId;
    return c.title === t.title;
  }

  // Ends the running clock c at `now`; returns { start, end, mins, opened, error }.
  function endSession(c, now) {
    var rs = Office.context.roamingSettings;
    var start = new Date(c.start);
    var end = now;
    if (end - start < 60000) end = new Date(start.getTime() + 60000);
    var mins = Math.max(1, Math.round((end - start) / 60000));
    rs.set(KEY_LAST, { title: c.title, taskId: c.taskId, link: c.link, start: start.toISOString(), end: end.toISOString() });
    rs.remove(KEY_CLOCK);
    var lines = c.fire ? [
      "Fire clocked with Task Clock. Type what the fire was after “Fire:” in the title, then Save.",
      fmtTime(start) + " – " + fmtTime(end) + " (" + mins + " min)",
      "If it turns into a multi-day task, add “cont.” to the end of the Planner task's title so it isn't logged twice."
    ] : [
      "Time worked on Planner task: " + c.title,
      fmtTime(start) + " – " + fmtTime(end) + " (" + mins + " min)"
    ];
    if (c.link) lines.push("Open the task in Planner: " + c.link);
    if (c.taskId) lines.push("TimeLog:" + c.taskId);
    var res = { start: start, end: end, mins: mins, opened: true, error: "" };
    try {
      Office.context.mailbox.displayNewAppointmentForm({
        requiredAttendees: [],
        optionalAttendees: [],
        start: start,
        end: end,
        location: "",
        resources: [],
        subject: c.fire ? "Fire: " : "Worked: " + c.title,
        body: lines.join("\n")
      });
    } catch (e) {
      res.opened = false;
      res.error = e.message;
    }
    return res;
  }

  function span(r) {
    return fmtTime(r.start) + "–" + fmtTime(r.end) + " (" + r.mins + " min)";
  }

  function startClock(t, finish) {
    var rs = Office.context.roamingSettings;
    var c = rs.get(KEY_CLOCK);
    if (c && sameTask(c, t)) {
      notify((t.fire ? "Already clocked in to a fire since " : "Already clocked in to this task since ") + fmtTime(new Date(c.start)) + ".", finish);
      return;
    }
    var now = new Date();
    var prev = c ? endSession(c, now) : null;
    rs.set(KEY_CLOCK, { title: t.title, taskId: t.taskId, link: t.link, fire: !!t.fire, start: now.toISOString() });
    save(function (ok, err) {
      if (!ok) return notify("Couldn't clock in: " + (err || "save failed"), finish);
      var label = t.fire ? "a fire" : short(t.title, 40);
      if (!prev) return notify("Clocked in at " + fmtTime(now) + " to " + (t.fire ? "a fire" : short(t.title, 90)) + ".", finish);
      notify("Switched. " + short(c.fire ? "Fire" : c.title, 40) + ": " + span(prev) +
        (prev.opened ? ". Save the entry that opened." : " (entry didn't open).") +
        " Now on " + label + ".", finish);
    });
  }

  function clockIn(event) {
    var finish = function () { event.completed(); };
    readItem(function (subject, html) {
      var p = parseBody(html);
      startClock({ title: cleanTitle(subject), taskId: p.taskId, link: p.link, fire: false }, finish);
    });
  }

  function clockInFire(event) {
    var finish = function () { event.completed(); };
    startClock({ title: "Fire", taskId: "", link: "", fire: true }, finish);
  }

  function clockOut(event) {
    var rs = Office.context.roamingSettings;
    var finish = function () { event.completed(); };
    var c = rs.get(KEY_CLOCK);
    if (!c) {
      notify("You're not clocked in.", finish);
      return;
    }
    var r = endSession(c, new Date());
    save(function () {
      if (!r.opened) return notify("Clocked out " + span(r) + ", but the entry didn't open: " + r.error, finish);
      notify("Clocked out: " + span(r) + ". Click Save on the entry that opened.", finish);
    });
  }

  window.clockIn = clockIn;
  window.clockInFire = clockInFire;
  window.clockOut = clockOut;

  Office.onReady(function () {
    if (Office.actions && Office.actions.associate) {
      Office.actions.associate("clockIn", clockIn);
      Office.actions.associate("clockInFire", clockInFire);
      Office.actions.associate("clockOut", clockOut);
    }
  });
})();
