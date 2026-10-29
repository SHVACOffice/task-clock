/* Task Clock — Clock Out panel. Stops the running clock, opens the "Worked:" calendar entry,
   and posts the optional note to the Planner task's chat as the signed-in person. */
(function () {
  "use strict";

  var KEY_CLOCK = "clock";      // { title, taskId, link, fire, start }
  var KEY_LAST = "lastSession"; // { title, taskId, link, start, end }

  function $(id) { return document.getElementById(id); }

  function setMsg(text, kind) {
    var m = $("msg");
    m.textContent = text || "";
    m.className = "msg" + (kind ? " " + kind : "");
  }

  function fmtTime(d) {
    return d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  }

  function settings() { return Office.context.roamingSettings; }

  function cleanNote(s) {
    return (s || "").replace(/\s+/g, " ").trim().slice(0, 2000);
  }

  function workedEntry(c, start, end, mins) {
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
    // A fire needs its title typed in, so it always opens the form.
    return { start: start, end: end, subject: c.fire ? "Fire: " : "Worked: " + c.title, lines: lines, formOnly: !!c.fire };
  }

  function clockOut() {
    var btn = $("outBtn");
    btn.disabled = true;
    var c = settings().get(KEY_CLOCK);
    if (!c) { setMsg("You're not clocked in.", "err"); return; }
    var note = cleanNote($("note").value);
    var start = new Date(c.start);
    var end = new Date();
    if (end - start < 60000) end = new Date(start.getTime() + 60000);
    var mins = Math.max(1, Math.round((end - start) / 60000));
    settings().set(KEY_LAST, { title: c.title, taskId: c.taskId, link: c.link, start: start.toISOString(), end: end.toISOString() });
    settings().remove(KEY_CLOCK);
    settings().saveAsync(function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) {
        setMsg("Couldn't save the clock state: " + (r.error && r.error.message), "err");
        btn.disabled = false;
        return;
      }
      TaskClockAuth.saveEntry(workedEntry(c, start, end, mins)).then(function (how) {
        return clearSpots().then(function (n) {
          finishUp((how === "saved" ? "Your time is saved to your calendar." : "Click Save on the time entry that opened.") +
            (n ? " Removed " + n + " planned spot" + (n === 1 ? "" : "s") + " for this task." : ""), true);
        });
      }, function () {
        finishUp("The time entry couldn't be saved or opened.", false);
      });
    });

    // Removes this task's planned spots from the day before the work through the end of today,
    // so the task doesn't show twice. Anything left is handled by the scheduler.
    function clearSpots() {
      if (!c.taskId) return Promise.resolve(0);
      var from = new Date(start); from.setHours(0, 0, 0, 0); from.setDate(from.getDate() - 1);
      var to = new Date(end); to.setHours(23, 59, 59, 0);
      return TaskClockAuth.removeSpots(c.taskId, from, to).catch(function () { return 0; });
    }

    function finishUp(entryMsg, ok) {
      var base = "Clocked out: " + fmtTime(start) + "–" + fmtTime(end) + " (" + mins + " min). " + entryMsg;
      if (!note) { setMsg(base, ok ? "ok" : "err"); return; }
      if (!c.taskId) {
        setMsg(base + " The note wasn't posted because this clock isn't tied to a Planner task.", "err");
        return;
      }
      setMsg(base + " Posting your note…");
      TaskClockAuth.postChat(c.taskId, "Stopped work (" + mins + " min): " + note).then(function () {
        setMsg(base + " Your note is in the task chat.", "ok");
      }, function (e) {
        setMsg(base + " The note couldn't be posted (" + e.message + "). Add it in Planner instead.", "err");
      });
    }
  }

  Office.onReady(function () {
    var c = settings().get(KEY_CLOCK);
    if (!c) {
      $("taskTitle").textContent = "You're not clocked in.";
      return;
    }
    $("taskTitle").textContent = c.fire ? "Fire" : c.title;
    $("taskSub").textContent = "Clocked in since " + fmtTime(new Date(c.start)) + ".";
    $("form").hidden = false;
    $("outBtn").addEventListener("click", clockOut);
    if (c.fire) { $("note").placeholder = "Notes can't be posted for fires (no Planner task)"; $("note").disabled = true; }
  });
})();
