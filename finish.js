/* Task Clock — Finish panel.
   Marks a Planner task as finished from its calendar event. The add-in cannot write to Planner,
   so it opens a pre-filled calendar entry carrying "TaskDone:<taskId>" (and an optional
   "Planner note:" line). The scheduler flow reads that entry at its next run, marks the task
   complete in Planner, adds the note to the task, and removes the task's planned spots.
   - Clocked in to this task: the clock stops and the entry is the "Worked:" record of the time.
   - Not clocked in to this task: the entry is a 1-minute "Done:" marker that the flow deletes
     after it has finished the task. */
(function () {
  "use strict";

  var KEY_CLOCK = "clock";      // { title, taskId, link, fire, start }
  var KEY_LAST = "lastSession"; // { title, taskId, link, start, end }
  var current = { title: "", taskId: "", link: "" };

  function $(id) { return document.getElementById(id); }

  function setMsg(text, isErr) {
    var m = $("msg");
    m.textContent = text || "";
    m.className = "msg" + (isErr ? " err" : "");
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
    // One line, no characters the flow uses to find the end of the note.
    return (s || "").replace(/[\r\n]+/g, " ").replace(/[<>]/g, "").replace(/\s+/g, " ").trim().slice(0, 900);
  }

  function openEntry(opts) {
    Office.context.mailbox.displayNewAppointmentForm({
      requiredAttendees: [],
      optionalAttendees: [],
      start: opts.start,
      end: opts.end,
      location: "",
      resources: [],
      subject: opts.subject,
      body: opts.lines.join("\n")
    });
  }

  function finish() {
    var btn = $("finishBtn");
    btn.disabled = true;
    var note = cleanNote($("note").value);
    var c = settings().get(KEY_CLOCK) || null;
    var clockedHere = !!(c && !c.fire && c.taskId && c.taskId === current.taskId);
    var now = new Date();
    var lines, subject, start, end;

    if (clockedHere) {
      start = new Date(c.start);
      end = now;
      if (end - start < 60000) end = new Date(start.getTime() + 60000);
      var mins = Math.max(1, Math.round((end - start) / 60000));
      subject = "Worked: " + current.title;
      lines = [
        "Time worked on Planner task: " + current.title,
        fmtTime(start) + " – " + fmtTime(end) + " (" + mins + " min)",
        "Finished with Task Clock."
      ];
      settings().set(KEY_LAST, { title: c.title, taskId: c.taskId, link: c.link, start: start.toISOString(), end: end.toISOString() });
      settings().remove(KEY_CLOCK);
    } else {
      start = now;
      end = new Date(now.getTime() + 60000);
      subject = "Done: " + current.title;
      lines = [
        "Finished with Task Clock (no time clocked).",
        "This entry is removed automatically once the task is marked complete in Planner."
      ];
    }
    if (note) lines.push("Planner note: " + note);
    if (current.link) lines.push("Open the task in Planner: " + current.link);
    if (clockedHere) lines.push("TimeLog:" + current.taskId);
    lines.push("TaskDone:" + current.taskId);
    if (!clockedHere) lines.push("DoneMarker");

    var proceed = function () {
      try {
        openEntry({ start: start, end: end, subject: subject, lines: lines });
        setMsg("Click Save on the entry that opened. Planner will show the task complete within about 30 minutes." +
          (c && !clockedHere ? " You're still clocked in to “" + (c.fire ? "a fire" : c.title) + "”." : ""));
      } catch (e) {
        setMsg("Couldn't open the calendar entry: " + e.message, true);
        btn.disabled = false;
      }
    };
    if (clockedHere) {
      settings().saveAsync(function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          setMsg("Couldn't save the clock state: " + (r.error && r.error.message), true);
          btn.disabled = false;
          return;
        }
        proceed();
      });
    } else {
      proceed();
    }
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
