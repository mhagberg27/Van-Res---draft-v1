/*
  W&L Athletics Van Reservations
  Front end for GitHub Pages. Data lives in a SharePoint list, reached through
  two Power Automate HTTP flows (see README.md). Sign-in uses W&L Microsoft 365.
*/
(function () {
  "use strict";

  const CFG = window.VAN_CONFIG || {};
  const VANS = CFG.vans || [];
  const AUTH = CFG.auth || {};
  const AUTH_ON = !!AUTH.clientId;
  const DEMO = !CFG.listUrl || !CFG.bookUrl;
  const DOMAIN = (CFG.emailDomain || "").toLowerCase();

  const DEMO_KEY = "wlu-van-demo-v1";
  const PREFS_KEY = "wlu-van-prefs-v1";
  const PENDING_KEY = "wlu-van-pending-v1";
  const REFRESH_MS = 60000;

  const state = {
    reservations: [],
    view: "list",
    listDate: todayISO(),
    month: todayISO().slice(0, 8) + "01",
    selectedDay: todayISO(),
    preselectVan: null,
    loadPromise: null
  };

  let msalApp = null;
  let account = null;
  let demoMemory = null;

  /* ================= Helpers ================= */

  function $(sel) { return document.querySelector(sel); }
  function pad(n) { return String(n).padStart(2, "0"); }
  function toISO(d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); }
  function fromISO(s) { const p = s.split("-").map(Number); return new Date(p[0], p[1] - 1, p[2]); }
  function todayISO() { return toISO(new Date()); }
  function addDays(s, n) { const d = fromISO(s); d.setDate(d.getDate() + n); return toISO(d); }
  function addMonths(s, n) { const d = fromISO(s); d.setDate(1); d.setMonth(d.getMonth() + n); return toISO(d); }
  function fmt(s, opts) { return fromISO(s).toLocaleDateString("en-US", opts); }
  function fmtShort(s) { return fmt(s, { month: "short", day: "numeric" }); }
  function fmtLong(s) { return fmt(s, { weekday: "long", month: "long", day: "numeric", year: "numeric" }); }
  function fmtRange(a, b) {
    const full = { month: "short", day: "numeric", year: "numeric" };
    if (a === b) return fmt(a, full);
    const sameYear = a.slice(0, 4) === b.slice(0, 4);
    return fmt(a, sameYear ? { month: "short", day: "numeric" } : full) + " to " + fmt(b, full);
  }
  function dayCount(a, b) { return Math.round((fromISO(b) - fromISO(a)) / 86400000) + 1; }
  function overlaps(a, b) { return a.startDate <= b.endDate && a.endDate >= b.startDate; }
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function vanName(id) { const v = VANS.find(function (x) { return x.id === id; }); return v ? v.name : "Van " + id; }
  function reservationOn(vanId, date) {
    return state.reservations.find(function (r) { return r.van === vanId && r.startDate <= date && r.endDate >= date; });
  }
  function isAvailable(vanId, start, end) {
    return !state.reservations.some(function (r) { return r.van === vanId && overlaps(r, { startDate: start, endDate: end }); });
  }
  function nextFreeDay(vanId, date) {
    let d = date;
    for (let i = 0; i < 730; i++) { if (!reservationOn(vanId, d)) return d; d = addDays(d, 1); }
    return null;
  }
  function storageGet(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } }
  function storageSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* storage blocked */ } }
  function storageDel(k) { try { localStorage.removeItem(k); } catch (e) { /* storage blocked */ } }

  let toastTimer = null;
  function toast(msg) {
    const t = $("#toast");
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.hidden = true; }, 7000);
  }
  function setStatus(msg) { $("#status").textContent = msg; }

  /* ================= W&L sign-in (Microsoft 365) ================= */

  function signedInEmail() { return account ? String(account.username || "").toLowerCase() : ""; }

  async function initAuth() {
    if (!AUTH_ON) { renderAuth(); return; }
    if (!window.msal) { setStatus("The W&L sign-in service did not load. Refresh the page to try again."); return; }

    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: AUTH.clientId,
        authority: "https://login.microsoftonline.com/" + (AUTH.tenantId || "organizations"),
        redirectUri: AUTH.redirectUri || (location.origin + location.pathname)
      },
      cache: { cacheLocation: "localStorage" }
    });

    try {
      const result = await msalApp.handleRedirectPromise();
      if (result && result.account) account = result.account;
    } catch (e) {
      console.error(e);
      toast("Sign-in did not finish. Please try again.");
    }
    if (!account) {
      const all = msalApp.getAllAccounts();
      if (all.length) account = all[0];
    }
    if (account) msalApp.setActiveAccount(account);
    renderAuth();

    // Resume a reservation the person started before signing in
    const pending = storageGet(PENDING_KEY);
    if (account && pending) {
      storageDel(PENDING_KEY);
      await load(true);
      openBooking(pending.van, pending.date);
    }
  }

  function signIn(pending) {
    if (pending) storageSet(PENDING_KEY, pending);
    msalApp.loginRedirect({ scopes: ["User.Read"], prompt: "select_account" });
  }

  function signOut() {
    msalApp.logoutRedirect({ account: account, postLogoutRedirectUri: AUTH.redirectUri || (location.origin + location.pathname) });
  }

  async function getFlowToken() {
    if (!AUTH_ON || !AUTH.flowScope || !account) return null;
    const req = { scopes: [AUTH.flowScope], account: account };
    try {
      return (await msalApp.acquireTokenSilent(req)).accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) { msalApp.acquireTokenRedirect(req); }
      throw e;
    }
  }

  function renderAuth() {
    const area = $("#authArea");
    if (!AUTH_ON) { area.innerHTML = ""; return; }
    if (account) {
      area.innerHTML = '<span class="who">Signed in as ' + esc(account.name || account.username) + "</span>" +
        '<button type="button" class="linklike" id="signOutBtn">Sign out</button>';
    } else {
      area.innerHTML = '<button type="button" class="linklike" id="signInBtn">Sign in with W&amp;L</button>';
    }
  }

  /* ================= Data ================= */

  function normalize(r) {
    function pick() {
      for (let i = 0; i < arguments.length; i++) {
        let v = r[arguments[i]];
        if (v && typeof v === "object" && "Value" in v) v = v.Value;
        if (v !== undefined && v !== null && v !== "") return String(v);
      }
      return "";
    }
    return {
      id: pick("id", "ID", "Id"),
      van: pick("van", "Van"),
      startDate: pick("startDate", "StartDate").slice(0, 10),
      endDate: pick("endDate", "EndDate").slice(0, 10),
      team: pick("team", "Team"),
      driver: pick("driver", "DriverName"),
      destination: pick("destination", "Destination")
    };
  }

  function demoSeed() {
    const t = todayISO();
    return [
      { id: "d1", van: "208", startDate: t, endDate: t, team: "Men's Tennis", driver: "Sample Coach", destination: "Recruiting" },
      { id: "d2", van: "140", startDate: addDays(t, 1), endDate: addDays(t, 3), team: "Field Hockey", driver: "Sample Coach", destination: "Away game" },
      { id: "d3", van: "203", startDate: addDays(t, 5), endDate: addDays(t, 6), team: "Football", driver: "Sample Coach", destination: "Scouting" }
    ];
  }
  function demoLoad() {
    let rows = storageGet(DEMO_KEY);
    if (!Array.isArray(rows)) rows = demoMemory || demoSeed();
    demoMemory = rows;
    storageSet(DEMO_KEY, rows);
    return rows.slice();
  }
  function demoSave(rows) { demoMemory = rows; storageSet(DEMO_KEY, rows); }

  async function fetchReservations() {
    if (DEMO) return demoLoad();
    const res = await fetch(CFG.listUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ today: todayISO() })
    });
    if (!res.ok) throw new Error("List request failed (" + res.status + ")");
    const data = await res.json();
    const rows = Array.isArray(data) ? data : (data.reservations || data.value || []);
    return rows.map(normalize).filter(function (r) { return r.van && r.startDate && r.endDate; });
  }

  async function createReservation(payload) {
    if (DEMO) {
      const rows = demoLoad();
      if (rows.some(function (r) { return r.van === payload.van && overlaps(r, payload); })) {
        return { ok: false, conflict: true };
      }
      rows.push(Object.assign({ id: "d" + Date.now() }, payload));
      demoSave(rows);
      return { ok: true };
    }
    const headers = { "Content-Type": "application/json" };
    const token = await getFlowToken();
    if (token) headers.Authorization = "Bearer " + token;

    const res = await fetch(CFG.bookUrl, { method: "POST", headers: headers, body: JSON.stringify(payload) });
    let body = {};
    try { body = await res.json(); } catch (e) { /* empty body */ }
    if (res.status === 409) return { ok: false, conflict: true, message: body.message };
    if (res.status === 400) return { ok: false, message: body.message || "The reservation details were not accepted." };
    if (!res.ok) throw new Error(body.message || "Booking request failed (" + res.status + ")");
    return { ok: true, id: body.id };
  }

  function load(quiet) {
    if (state.loadPromise) return state.loadPromise;
    if (!quiet) setStatus("Loading reservations...");
    state.loadPromise = fetchReservations()
      .then(function (rows) {
        rows.sort(function (a, b) { return a.startDate.localeCompare(b.startDate) || Number(a.van) - Number(b.van); });
        state.reservations = rows;
        setStatus("Last updated " + new Date().toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" }));
        render();
      })
      .catch(function (e) {
        console.error(e);
        setStatus("Could not load reservations. Check your connection and refresh the page.");
      })
      .finally(function () { state.loadPromise = null; });
    return state.loadPromise;
  }

  /* ================= Rendering ================= */

  function render() {
    $("#listView").hidden = state.view !== "list";
    $("#calendarView").hidden = state.view !== "calendar";
    document.querySelectorAll(".toggle").forEach(function (b) {
      b.setAttribute("aria-pressed", String(b.dataset.view === state.view));
    });
    if (state.view === "list") renderList(); else { renderCalendar(); renderDay(); }
  }

  function renderList() {
    const date = state.listDate;
    const today = todayISO();
    $("#listDate").value = date;

    $("#vanList").innerHTML = VANS.map(function (v) {
      const cur = reservationOn(v.id, date);
      const upcoming = state.reservations.filter(function (r) { return r.van === v.id && r.endDate >= today; });

      let status;
      if (cur) {
        const free = nextFreeDay(v.id, date);
        status =
          '<p class="van-status"><span class="pill busy">Reserved</span>' +
          "<span><strong>" + esc(cur.team) + "</strong><br>" + esc(fmtRange(cur.startDate, cur.endDate)) + "</span></p>" +
          (free ? '<p class="hint">Available again ' + esc(fmt(free, { weekday: "short", month: "short", day: "numeric" })) + "</p>" : "");
      } else {
        status =
          '<p class="van-status"><span class="pill ok">Available</span>' +
          "<span>on " + esc(fmt(date, { weekday: "short", month: "short", day: "numeric" })) + "</span></p>";
      }

      const list = upcoming.length
        ? '<ul class="res-list">' + upcoming.map(function (r) {
            const isNow = r.startDate <= date && r.endDate >= date;
            return '<li class="' + (isNow ? "now" : "") + '"><span>' + esc(fmtRange(r.startDate, r.endDate)) +
              '</span><span class="res-team">' + esc(r.team) + "</span></li>";
          }).join("") + "</ul>"
        : '<p class="hint">No upcoming reservations.</p>';

      const btn = (!cur && date >= today)
        ? '<button type="button" class="btn small" data-reserve="' + esc(v.id) + '" data-date="' + date + '">Reserve ' + esc(v.name) + "</button>"
        : "";

      return '<article class="van-card ' + (cur ? "is-busy" : "is-ok") + '">' +
        "<h3>" + esc(v.name) + "</h3>" + status +
        "<h4>Upcoming reservations</h4>" + list + btn + "</article>";
    }).join("");
  }

  function renderCalendar() {
    const m = state.month;
    const today = todayISO();
    const first = fromISO(m);
    const daysInMonth = new Date(first.getFullYear(), first.getMonth() + 1, 0).getDate();
    const cellCount = Math.ceil((first.getDay() + daysInMonth) / 7) * 7;
    const start = addDays(m, -first.getDay());

    $("#calHeading").textContent = fmt(m, { month: "long", year: "numeric" });

    let html = "";
    for (let i = 0; i < cellCount; i++) {
      const d = addDays(start, i);
      const res = state.reservations.filter(function (r) { return r.startDate <= d && r.endDate >= d; });
      const classes = ["day"];
      if (d.slice(0, 7) !== m.slice(0, 7)) classes.push("out");
      if (d < today) classes.push("past");
      if (d === today) classes.push("today");
      if (d === state.selectedDay) classes.push("selected");

      const label = fmtLong(d) + ": " + (res.length
        ? res.length + " of " + VANS.length + " vans reserved, " + res.map(function (r) { return vanName(r.van) + " by " + r.team; }).join("; ")
        : "all vans available");

      const chips = res.length
        ? res.map(function (r) {
            return '<span class="chip" title="' + esc(vanName(r.van) + ", " + r.team + ", " + fmtRange(r.startDate, r.endDate)) + '">' +
              esc(r.van) + " " + esc(r.team) + "</span>";
          }).join("")
        : (d >= today ? '<span class="open">All open</span>' : "");

      const count = res.length
        ? '<span class="count busy">' + res.length + " out</span>"
        : (d >= today ? '<span class="count ok">Open</span>' : "");

      html += '<button type="button" class="' + classes.join(" ") + '" data-day="' + d + '" aria-label="' + esc(label) +
        '" aria-pressed="' + (d === state.selectedDay) + '">' +
        '<span class="num" aria-hidden="true">' + fromISO(d).getDate() + "</span>" +
        '<span class="chips" aria-hidden="true">' + chips + "</span>" +
        '<span aria-hidden="true">' + count + "</span></button>";
    }
    $("#calGrid").innerHTML = html;
  }

  function renderDay() {
    const d = state.selectedDay;
    const today = todayISO();
    $("#dayHeading").textContent = fmtLong(d);
    $("#dayDetail").innerHTML = '<ul class="day-list">' + VANS.map(function (v) {
      const r = reservationOn(v.id, d);
      if (r) {
        return '<li><span class="pill busy">Reserved</span><span class="dl-van">' + esc(v.name) + "</span>" +
          '<span class="dl-info"><strong>' + esc(r.team) + "</strong>, " + esc(fmtRange(r.startDate, r.endDate)) +
          (r.destination ? " (" + esc(r.destination) + ")" : "") + "</span></li>";
      }
      return '<li><span class="pill ok">Available</span><span class="dl-van">' + esc(v.name) + "</span>" +
        (d >= today ? '<button type="button" class="btn small" data-reserve="' + esc(v.id) + '" data-date="' + d + '">Reserve</button>' : "") +
        "</li>";
    }).join("") + "</ul>";
  }

  /* ================= Booking ================= */

  function fillTeams() {
    const sel = $("#team");
    sel.innerHTML = '<option value="">Choose a team</option>' +
      (CFG.teams || []).map(function (t) { return '<option value="' + esc(t) + '">' + esc(t) + "</option>"; }).join("") +
      '<option value="__other">Other</option>';
  }

  function selectedVan() {
    const c = document.querySelector('input[name="van"]:checked');
    return c ? c.value : null;
  }

  function updateChoices() {
    const s = $("#startDate").value;
    let e = $("#endDate").value;
    $("#endDate").min = s || todayISO();
    if (s && (!e || e < s)) { e = s; $("#endDate").value = s; }

    const prev = selectedVan() || state.preselectVan;
    if (!s || !e) {
      $("#vanChoices").innerHTML = "";
      $("#dayCount").textContent = "";
      $("#unavailNote").textContent = "Choose your dates to see which vans are open.";
      return;
    }
    const n = dayCount(s, e);
    $("#dayCount").textContent = n + (n === 1 ? " day: " : " days: ") + fmtRange(s, e);

    const avail = VANS.filter(function (v) { return isAvailable(v.id, s, e); });
    const taken = VANS.length - avail.length;

    $("#vanChoices").innerHTML = avail.length
      ? avail.map(function (v) {
          return '<label class="choice"><input type="radio" name="van" value="' + esc(v.id) + '"' +
            (v.id === prev ? " checked" : "") + "><span>" + esc(v.name) + "</span></label>";
        }).join("")
      : '<p class="error">No vans are open for these dates. Try different dates.</p>';

    $("#unavailNote").textContent = taken
      ? taken + (taken === 1 ? " van is" : " vans are") + " already reserved for some or all of these dates and not shown."
      : "All vans are open for these dates.";
  }

  function openBooking(vanId, date) {
    if (AUTH_ON && !account) { signIn({ van: vanId || null, date: date || null }); return; }

    const today = todayISO();
    const start = date && date >= today ? date : today;
    state.preselectVan = vanId || null;

    $("#bookForm").reset();
    $("#startDate").min = today;
    $("#startDate").value = start;
    $("#endDate").value = start;
    $("#formError").textContent = "";

    const prefs = storageGet(PREFS_KEY) || {};
    if (prefs.team) $("#team").value = prefs.team;
    if (prefs.otherTeam) $("#otherTeam").value = prefs.otherTeam;
    $("#otherTeamWrap").hidden = $("#team").value !== "__other";

    if (AUTH_ON && account) {
      $("#emailWrap").hidden = true;
      $("#signedInAs").hidden = false;
      $("#signedInAs").textContent = "Reserving as " + (account.name || "") + " (" + signedInEmail() + ")";
      $("#driver").value = prefs.driver || account.name || "";
    } else {
      $("#emailWrap").hidden = false;
      $("#signedInAs").hidden = true;
      $("#driver").value = prefs.driver || "";
      $("#email").value = prefs.email || "";
    }

    updateChoices();
    $("#bookDialog").showModal();
  }

  async function submitBooking(ev) {
    ev.preventDefault();
    const err = $("#formError");
    const btn = $("#submitBtn");
    err.textContent = "";

    const s = $("#startDate").value;
    const e = $("#endDate").value;
    const van = selectedVan();
    let team = $("#team").value;
    if (team === "__other") team = $("#otherTeam").value.trim();
    const driver = $("#driver").value.trim();
    const email = AUTH_ON && account ? signedInEmail() : $("#email").value.trim().toLowerCase();
    const today = todayISO();

    const problems = [];
    if (!s || !e) problems.push("Choose a start and end date.");
    else if (s < today) problems.push("The start date cannot be in the past.");
    else if (e < s) problems.push("The end date must be on or after the start date.");
    else if (CFG.maxDays && dayCount(s, e) > CFG.maxDays) problems.push("Reservations can be at most " + CFG.maxDays + " days.");
    if (!van) problems.push("Choose a van.");
    if (!team) problems.push("Choose your team.");
    if (!driver) problems.push("Enter the driver's name.");
    if (!/^[^@\s]+@[^@\s]+$/.test(email) || (DOMAIN && !email.endsWith("@" + DOMAIN))) {
      problems.push(AUTH_ON ? "Please sign in with your W&L account." : "Enter a valid @" + DOMAIN + " email address.");
    }
    if (problems.length) { err.textContent = problems.join(" "); return; }

    btn.disabled = true;
    btn.textContent = "Checking...";
    try {
      // Get the latest reservations so nobody books over someone else
      await load(true);
      if (!isAvailable(van, s, e)) {
        err.textContent = "Sorry, " + vanName(van) + " was just reserved for some of those dates. Please choose another van or different dates.";
        updateChoices();
        return;
      }

      btn.textContent = "Reserving...";
      const payload = {
        van: van,
        vanName: vanName(van),
        startDate: s,
        endDate: e,
        team: team,
        driver: driver,
        email: email,
        reservedBy: account ? (account.name || email) : driver,
        destination: $("#destination").value.trim(),
        notes: $("#notes").value.trim()
      };
      const result = await createReservation(payload);
      if (!result.ok) {
        err.textContent = result.message || ("Sorry, " + vanName(van) + " was just reserved for some of those dates. Please choose another van or different dates.");
        await load(true);
        updateChoices();
        return;
      }

      storageSet(PREFS_KEY, {
        team: $("#team").value,
        otherTeam: $("#otherTeam").value.trim(),
        driver: driver,
        email: AUTH_ON ? "" : email
      });
      $("#bookDialog").close();
      toast("Reserved: " + vanName(van) + " for " + team + ", " + fmtRange(s, e) + ".");
      await load(true);
    } catch (x) {
      console.error(x);
      err.textContent = "Something went wrong and the reservation was not saved. Please try again.";
    } finally {
      btn.disabled = false;
      btn.textContent = "Reserve";
    }
  }

  /* ================= Events ================= */

  function wire() {
    document.addEventListener("click", function (ev) {
      const reserve = ev.target.closest("[data-reserve]");
      if (reserve) { openBooking(reserve.dataset.reserve, reserve.dataset.date); return; }

      const day = ev.target.closest("[data-day]");
      if (day) {
        state.selectedDay = day.dataset.day;
        if (state.selectedDay.slice(0, 7) !== state.month.slice(0, 7)) state.month = state.selectedDay.slice(0, 8) + "01";
        renderCalendar();
        renderDay();
        const again = document.querySelector('[data-day="' + state.selectedDay + '"]');
        if (again) again.focus();
        return;
      }

      const toggle = ev.target.closest(".toggle");
      if (toggle) {
        state.view = toggle.dataset.view;
        storageSet(PREFS_KEY + "-view", state.view);
        render();
        return;
      }

      if (ev.target.id === "signInBtn") signIn();
      if (ev.target.id === "signOutBtn") signOut();
    });

    $("#reserveBtn").addEventListener("click", function () {
      openBooking(null, state.view === "calendar" ? state.selectedDay : state.listDate);
    });
    $("#listDate").addEventListener("change", function (e) { if (e.target.value) { state.listDate = e.target.value; renderList(); } });
    $("#prevMonth").addEventListener("click", function () { state.month = addMonths(state.month, -1); renderCalendar(); });
    $("#nextMonth").addEventListener("click", function () { state.month = addMonths(state.month, 1); renderCalendar(); });
    $("#todayBtn").addEventListener("click", function () {
      state.selectedDay = todayISO();
      state.month = state.selectedDay.slice(0, 8) + "01";
      renderCalendar();
      renderDay();
    });

    $("#startDate").addEventListener("change", updateChoices);
    $("#endDate").addEventListener("change", updateChoices);
    $("#team").addEventListener("change", function () {
      $("#otherTeamWrap").hidden = $("#team").value !== "__other";
    });
    $("#bookForm").addEventListener("submit", submitBooking);
    $("#cancelBtn").addEventListener("click", function () { $("#bookDialog").close(); });
    $("#closeDialog").addEventListener("click", function () { $("#bookDialog").close(); });

    // Keep the page current
    setInterval(function () { if (!document.hidden) load(true); }, REFRESH_MS);
    document.addEventListener("visibilitychange", function () { if (!document.hidden) load(true); });
  }

  /* ================= Start ================= */

  async function start() {
    $("#demoBanner").hidden = !DEMO;
    const savedView = storageGet(PREFS_KEY + "-view");
    if (savedView === "list" || savedView === "calendar") state.view = savedView;
    fillTeams();
    wire();
    render();
    await load(false);
    await initAuth();
  }

  start();
})();
