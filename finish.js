/* Task Clock — Finish panel.
   Signs in as the person using Outlook, posts the optional note to the Planner task's chat under
   their name, and marks the task complete in Planner right away.
   - Clocked in to this task: the clock stops and a "Worked:" entry opens to record the time.
     It carries "TaskDone:<id>" so the scheduler clears the task's remaining planned spots.
   - If sign-in or Planner isn't reachable, it falls back to the older route: a calendar entry with
     "TaskDone:<id>" (and a "Planner note:" line) that the scheduler acts on at its next run. */
(function () {
  "use strict";

  var KEY_CLOCK = "clock";      // { title, taskId, link, fire, start }
  var KEY_LAST = "lastSession"; // { title, taskId, link, start, end }
  var current = { title: "", taskId: "", link: "" };

  function $(id) { return document.getElementById(id); }

  function setMsg(text, kind) {
    var m = $("msg");
    m.textContent = text || "";
    m.className = "msg" + (kind ? " " + kind : "");
  }

  function fmtTime(d) {
    return d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  }

  function decodeEntities(s) {
    var t = document.createElement("textarea");
    t.innerHTML = s;
    return t.value;
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

  function settings() { return Office.context.roamingSettings; }

  function cleanNote(s) {
    return (s || "").replace(/\s+/g, " ").trim().slice(0, 2000);
  }

  // For the fallback "Planner note:" line, which the flow reads up to the next "<".
  function flowNote(s) {
    return s.replace(/[<>]/g, "").slice(0, 900);
  }

  function saveSettings() {
    return new Promise(function (resolve, reject) {
      settings().saveAsync(function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(r.error && r.error.message));
      });
    });
  }

  function finish() {
    var btn = $("finishBtn");
    btn.disabled = true;
    setMsg("Finishing…");
    var note = cleanNote($("note").value);
    var c = settings().get(KEY_CLOCK) || null;
    var clockedHere = !!(c && !c.fire && c.taskId && c.taskId === current.taskId);
    var now = new Date();
    var start, end, mins = 0;

    if (clockedHere) {
      start = new Date(c.start);
      end = now;
      if (end - start < 60000) end = new Date(start.getTime() + 60000);
      mins = Math.max(1, Math.round((end - start) / 60000));
      settings().set(KEY_LAST, { title: c.title, taskId: c.taskId, link: c.link, start: start.toISOString(), end: end.toISOString() });
      settings().remove(KEY_CLOCK);
    }

    var result = { chat: !note, chatErr: "", done: false, doneErr: "" };
    var chatText = "Finished" + (clockedHere ? " (" + mins + " min)" : "") + ": " + note;

    (clockedHere ? saveSettings() : Promise.resolve()).then(function () {
      var chatP = note ? TaskClockAuth.postChat(current.taskId, chatText).then(
        function () { result.chat = true; }, function (e) { result.chatErr = e.message; }) : Promise.resolve();
      return chatP.then(function () {
        return TaskClockAuth.completeTask(current.taskId).then(
          function () { result.done = true; }, function (e) { result.doneErr = e.message; });
      }).then(function () {
        if (!result.done) return;
        // The task is finished, so its remaining planned spots come off the calendar now.
        var from = new Date(); from.setHours(0, 0, 0, 0); from.setDate(from.getDate() - 1);
        var to = new Date(); to.setDate(to.getDate() + 14);
        return TaskClockAuth.removeSpots(current.taskId, from, to).then(
          function (n) { result.removed = n; }, function () { result.removed = 0; });
      });
    }).then(function () {
      var needEntry = clockedHere || !result.done || (note && !result.chat);
      if (needEntry) {
        var lines, subject, s, e;
        if (clockedHere) {
          s = start; e = end;
          subject = "Worked: " + current.title;
          lines = [
            "Time worked on Planner task: " + current.title,
            fmtTime(start) + " – " + fmtTime(end) + " (" + mins + " min)",
            "Finished with Task Clock."
          ];
        } else {
          s = now; e = new Date(now.getTime() + 60000);
          subject = "Done: " + current.title;
          lines = [
            "Finished with Task Clock (no time clocked).",
            "This entry is removed automatically once the task is marked complete in Planner."
          ];
        }
        if (note && !result.chat) lines.push("Planner note: " + flowNote(note));
        if (current.link) lines.push("Open the task in Planner: " + current.link);
        if (clockedHere) lines.push("TimeLog:" + current.taskId);
        lines.push("TaskDone:" + current.taskId);
        if (!clockedHere) lines.push("DoneMarker");
        return TaskClockAuth.saveEntry({ start: s, end: e, subject: subject, lines: lines }).then(
          function (how) { return how; }, function () { return "failed"; });
      }
      return "none";
    }).then(function (how) {
      var parts = [];
      parts.push(result.done ? "Marked complete in Planner." :
        "Couldn't mark it complete directly (" + result.doneErr + "). The half-hour update will finish it from the calendar entry.");
      if (note) parts.push(result.chat ? "Your note is in the task chat." :
        "The note couldn't go to the task chat (" + result.chatErr + "), so it will be added to the task's Notes instead.");
      var entryMsg = how === "saved" ? (clockedHere ? "Clocked out; your time is saved to your calendar." : "") :
        how === "form" ? (clockedHere ? "Clocked out. Click Save on the time entry that opened." : "Click Save on the entry that opened.") :
        how === "failed" ? "The calendar entry couldn't be saved or opened." : "";
      if (entryMsg) parts.push(entryMsg);
      if (result.removed) parts.push("Removed " + result.removed + " planned spot" + (result.removed === 1 ? "" : "s") + " from your calendar.");
      if (c && !clockedHere) parts.push("You're still clocked in to “" + (c.fire ? "a fire" : c.title) + "”.");
      setMsg(parts.join(" "), result.done && result.chat && how !== "failed" ? "ok" : "err");
    }).catch(function (e) {
      setMsg("Couldn't finish: " + e.message, "err");
      btn.disabled = false;
    });
  }

  Office.onReady(function () {
    var item = Office.context.mailbox && Office.context.mailbox.item;
    if (!item) { $("taskTitle").textContent = "Open a calendar event to finish its task."; return; }
    readSubject(item, function (subject) {
      readBody(item, function (html) {
        var p = parseBody(html);
        current = { title: cleanTitle(subject), taskId: p.taskId, link: p.link };
        $("taskTitle").textContent = current.title;
        if (!current.taskId) {
          $("taskSub").textContent = "This event isn't linked to a Planner task, so there's nothing to mark complete.";
          return;
        }
        var c = settings().get(KEY_CLOCK) || null;
        var card = $("clockCard");
        if (c && !c.fire && c.taskId === current.taskId) {
          card.textContent = "Clocked in since " + fmtTime(new Date(c.start)) + ". Finishing also clocks you out and saves the time.";
          card.hidden = false;
        } else if (c) {
          card.textContent = "You're clocked in to “" + (c.fire ? "a fire" : c.title) + "”. Finishing this task won't stop that clock.";
          card.hidden = false;
        }
        $("form").hidden = false;
        $("finishBtn").addEventListener("click", finish);
      });
    });
  });
})();
