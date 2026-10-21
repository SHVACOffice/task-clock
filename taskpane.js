/* Task Clock — clock in and out of Planner tasks from their Outlook calendar events.
   State lives in the add-in's roaming settings (per mailbox, synced across Outlook clients).
   Clocking out opens a pre-filled appointment form; the user clicks Save. */
(function () {
  "use strict";

  var KEY_CLOCK = "clock";      // { title, taskId, link, start }
  var KEY_LAST = "lastSession"; // { title, taskId, link, start, end }
  var current = { title: "", taskId: "", link: "" };
  var tick = null;

  function $(id) { return document.getElementById(id); }

  function setMsg(text, isErr) {
    var m = $("msg");
    m.textContent = text || "";
    m.className = "msg" + (isErr ? " err" : "");
  }

  function fmtTime(d) {
    return d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  }

  function fmtDay(d) {
    var today = new Date();
    if (d.toDateString() === today.toDateString()) return "today";
    return d.toLocaleDateString([], { weekday: "short", month: "short", day: "numeric" });
  }

  function fmtElapsed(ms) {
    var mins = Math.max(0, Math.floor(ms / 60000));
    var h = Math.floor(mins / 60), m = mins % 60;
    return h + ":" + (m < 10 ? "0" : "") + m;
  }

  function settings() { return Office.context.roamingSettings; }

  function getClock() { return settings().get(KEY_CLOCK) || null; }

  function save(cb) {
    settings().saveAsync(function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) {
        setMsg("Couldn't save the clock state: " + (r.error && r.error.message), true);
      }
      if (cb) cb(r.status === Office.AsyncResultStatus.Succeeded);
    });
  }

  function decodeEntities(s) {
    var t = document.createElement("textarea");
    t.innerHTML = s;
    return t.value;
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

  function readSubject(item, cb) {
    if (typeof item.subject === "string") return cb(item.subject);
    if (item.subject && item.subject.getAsync) {
      return item.subject.getAsync(function (r) { cb(r.status === Office.AsyncResultStatus.Succeeded ? r.value : ""); });
    }
    cb("");
  }

  function readBody(item, cb) {
    if (!item.body || !item.body.getAsync) return cb("");
    item.body.getAsync(Office.CoercionType.Html, function (r) {
      cb(r.status === Office.AsyncResultStatus.Succeeded ? r.value : "");
    });
  }

  function button(label, onClick, secondary) {
    var b = document.createElement("button");
    b.textContent = label;
    if (secondary) b.className = "secondary";
    b.addEventListener("click", function () { b.disabled = true; onClick(function () { b.disabled = false; }); });
    return b;
  }

  function openLogForm(s, done) {
    var start = new Date(s.start);
    var end = new Date(s.end);
    var mins = Math.max(1, Math.round((end - start) / 60000));
    var lines = [
      "Time worked on Planner task: " + s.title,
      fmtTime(start) + " – " + fmtTime(end) + " (" + mins + " min)"
    ];
    if (s.link) lines.push("Open the task in Planner: " + s.link);
    if (s.taskId) lines.push("TimeLog:" + s.taskId);
    try {
      Office.context.mailbox.displayNewAppointmentForm({
        requiredAttendees: [],
        optionalAttendees: [],
        start: start,
        end: end,
        location: "",
        resources: [],
        subject: "Worked: " + s.title,
        body: lines.join("\n")
      });
      setMsg("A filled-in calendar entry opened. Click Save on it to keep the time.");
    } catch (e) {
      setMsg("Couldn't open the calendar entry: " + e.message, true);
    }
    if (done) done();
  }

  function clockIn(done) {
    if (getClock()) { setMsg("You're already clocked in. Clock out first.", true); render(); return done(); }
    settings().set(KEY_CLOCK, {
      title: current.title, taskId: current.taskId, link: current.link, start: new Date().toISOString()
    });
    save(function (ok) { if (ok) setMsg(""); render(); done(); });
  }

  function clockOut(done) {
    var c = getClock();
    if (!c) { render(); return done(); }
    var start = new Date(c.start);
    var end = new Date();
    if (end - start < 60000) end = new Date(start.getTime() + 60000);
    var session = { title: c.title, taskId: c.taskId, link: c.link, start: start.toISOString(), end: end.toISOString() };
    settings().set(KEY_LAST, session);
    settings().remove(KEY_CLOCK);
    save(function () {
      render();
      openLogForm(session, done);
    });
  }

  function reopenLast(done) {
    var s = settings().get(KEY_LAST);
    if (!s) { setMsg("No recent clock-out to reopen.", true); return done(); }
    openLogForm(s, done);
  }

  function render() {
    if (tick) { clearInterval(tick); tick = null; }
    var c = getClock();
    var actions = $("actions");
    actions.innerHTML = "";
    var card = $("statusCard"), st = $("statusText"), el = $("elapsed");
    card.hidden = false;
    el.hidden = true;
    st.className = "status";

    var sameTask = c && ((c.taskId && c.taskId === current.taskId) || (!c.taskId && !current.taskId && c.title === current.title));

    if (!c) {
      st.textContent = "Not clocked in.";
      actions.appendChild(button("Clock in", clockIn));
    } else {
      var since = new Date(c.start);
      st.className = "status on";
      st.textContent = (sameTask ? "Clocked in since " : "Clocked in to “" + c.title + "” since ") +
        fmtTime(since) + (fmtDay(since) === "today" ? "" : " (" + fmtDay(since) + ")");
      el.hidden = false;
      var upd = function () { el.textContent = fmtElapsed(Date.now() - since.getTime()); };
      upd();
      tick = setInterval(upd, 15000);
      actions.appendChild(button(sameTask ? "Clock out" : "Clock out of that task", clockOut));
    }
    if (settings().get(KEY_LAST)) {
      actions.appendChild(button("Reopen last clock-out entry", reopenLast, true));
    }
  }

  Office.onReady(function () {
    var item = Office.context.mailbox && Office.context.mailbox.item;
    if (!item) { $("taskTitle").textContent = "Open a calendar event to use Task Clock."; return; }
    readSubject(item, function (subject) {
      readBody(item, function (html) {
        var p = parseBody(html);
        current = { title: cleanTitle(subject), taskId: p.taskId, link: p.link };
        $("taskTitle").textContent = current.title;
        var sub = $("taskSub");
        sub.innerHTML = "";
        if (current.link) {
          var a = document.createElement("a");
          a.href = current.link; a.target = "_blank"; a.rel = "noopener";
          a.textContent = "Open this task in Planner";
          sub.appendChild(a);
        } else if (!current.taskId) {
          sub.textContent = "This event isn't linked to a Planner task. Time will be logged under the event title.";
        }
        render();
      });
    });
  });
})();
