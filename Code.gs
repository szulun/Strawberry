/**
 * Strawberry — Google Apps Script backend.
 *
 * 貼到「Strawberry data」的 Apps Script，存檔，選 setup 按執行，授權。
 * 然後：部署 → 新增部署作業 → 網頁應用程式。
 * 執行身分：我。誰可以存取：所有人。
 * 把結尾是 /exec 的網址貼回 strawberry.html 的 API_URL。
 *
 * 試算表分頁：Rooms、People、Checkins。Tokens 會藏起來。
 * 破冰題要用 AI 時：專案設定 → 指令碼屬性，新增 ANTHROPIC_API_KEY。
 * 選用 MODEL，不填就用 claude-haiku-5-5。貼上新程式後先執行一次 testIcebreaker，才部署。
 */

var TABS = {
  Rooms: ["room_id", "code", "activity", "when_text", "place", "created_at", "host_person_id", "icebreakers", "when_at", "status", "meet2_activity", "meet2_place", "meet2_at", "meet3_activity", "meet3_place", "meet3_at", "reminded", "meet2_fixed", "meet3_fixed"],
  People: ["room_id", "person_id", "joined_at", "name", "email", "from_place", "standing", "hook", "home", "secret", "answers_json", "intro", "capsule", "wish", "vote", "slot", "bye", "contact", "guess_target", "guess_topic", "guess_text", "guess_verdict", "left", "next_activity", "meet3_at", "meet3_activity", "meet2_ok", "meet3_ok"],
  Tokens: ["room_id", "person_id", "token", "is_host"],
  Checkins: ["room_id", "person_id", "name", "again", "met", "reasons", "connected", "why", "group_name", "at"],
  Users: ["email", "name", "from_place", "standing", "home", "answers_json", "updated_at"],
  UserTokens: ["email", "token", "created_at"],
  Messages: ["room_id", "message_id", "person_id", "name", "text", "at"],
};

var CARD_KEYS = {
  drink: ["coffee", "boba", "matcha", "smoothie"],
  movie: ["horror", "romcom", "action", "animated"],
  music: ["karaoke", "playlist", "concert", "podcast"],
  hangout: ["boardgames", "foodout", "escape", "thrift"],
  trip: ["mountains", "city", "beach", "themepark"],
  clock: ["early", "night", "brunch", "philosopher"],
  food: ["cook", "streetfood", "potluck", "fancy"],
  vibe: ["quiet", "lively", "picnic", "sports"],
};
var STANDINGS = ["undergrad", "exchange", "grad"];
var HOMES = ["dish", "tradition", "expression", "other"];
var OPTION_LABELS = {
  coffee: "coffee", boba: "boba", matcha: "matcha", smoothie: "smoothies",
  horror: "horror movies", romcom: "rom-coms", action: "action movies", animated: "animated movies",
  karaoke: "karaoke", playlist: "playlists", concert: "concerts", podcast: "podcasts",
  boardgames: "board games", foodout: "eating out", escape: "escape rooms", thrift: "thrifting",
  mountains: "mountains", city: "cities", beach: "the beach", themepark: "theme parks",
  early: "early mornings", night: "late nights", brunch: "brunch", philosopher: "late-night talks",
  cook: "cooking", streetfood: "street food", potluck: "potlucks", fancy: "a nice meal",
  quiet: "quiet plans", lively: "lively plans", picnic: "picnics", sports: "sports",
};
var CODE_CHARS = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
// Test phase only. Set to false before a real Berkeley round.
var ALLOW_GMAIL = true;
var MAX_PEOPLE = 6;
var COLORS = ["#E2553A", "#1E9A60", "#6A4BE8", "#0E7490", "#C2410C", "#7C3AED"];

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty("SHEET_ID", ss.getId());
  Object.keys(TABS).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    var head = TABS[name];
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold");
    sh.setFrozenRows(1);
  });
  var tokens = ss.getSheetByName("Tokens");
  if (tokens) tokens.hideSheet();
  var userTokens = ss.getSheetByName("UserTokens");
  if (userTokens) userTokens.hideSheet();
  var blank = ss.getSheetByName("Sheet1") || ss.getSheetByName("工作表1");
  if (blank && blank.getLastRow() <= 1 && ss.getSheets().length > 1) ss.deleteSheet(blank);
  installReminders_();
  Logger.log("Setup done. Deploy as a web app. Execute as: Me. Access: Anyone.");
}

function doGet() {
  return json_({ ok: true, app: "Strawberry" });
}

function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    out = route_(req);
    out.ok = true;
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return json_(out);
}

function route_(r) {
  switch (r.action) {
    case "createRoom": return lock_(function () { return createRoom_(r); });
    case "join": return lock_(function () { return join_(r); });
    case "getState": return getState_(r);
    case "save": return lock_(function () { return save_(r); });
    case "checkin": return lock_(function () { return checkin_(r); });
    case "icebreaker": return icebreaker_(r);
    case "updateRoom": return lock_(function () { return updateRoom_(r); });
    case "cancelRoom": return lock_(function () { return cancelRoom_(r); });
    case "leaveRoom": return lock_(function () { return leaveRoom_(r); });
    case "listTables": return listTables_(r);
    case "postMessage": return lock_(function () { return postMessage_(r); });
    case "judge": return lock_(function () { return judge_(r); });
    case "proposePlan": return lock_(function () { return proposePlan_(r); });
    case "answerPlan": return lock_(function () { return answerPlan_(r); });
    case "sendCode": return lock_(function () { return sendCode_(r); });
    case "verifyCode": return lock_(function () { return verifyCode_(r); });
    case "me": return me_(r);
    default: throw new Error("Unknown action.");
  }
}

function createRoom_(r) {
  var activity = clip_(r.activity, 40);
  var whenAt = iso_(r.whenAt);
  var place = clip_(r.place, 80);
  if (activity.length < 2) throw new Error("Name the activity. That is the hook.");
  if (!whenAt) throw new Error("Add a date and time.");
  if (new Date(whenAt).getTime() < Date.now() - 2 * 60 * 1000) throw new Error("Pick a time that has not already passed.");
  if (place.length < 2) throw new Error("Add where you meet.");
  var whenText = clip_(r.when, 60) || prettyWhen_(whenAt);
  var profile = profile_(r);
  requireUser_(r, profile.email);
  upsertUser_(profile);
  var code = newCode_();
  var roomId = Utilities.getUuid();
  var personId = Utilities.getUuid();
  var token = Utilities.getUuid();
  var now = now_();
  append_("Rooms", {
    room_id: roomId, code: code, activity: activity, when_text: whenText, place: place,
    created_at: now, host_person_id: personId, when_at: whenAt, status: "open",
  });
  append_("People", personRow_(roomId, personId, profile, activity, now));
  append_("Tokens", { room_id: roomId, person_id: personId, token: token, is_host: "1" });
  try { installReminders_(); } catch (e) {}
  var room = rows_("Rooms").filter(function (x) { return x.room_id === roomId; })[0];
  remind_(room, [{ email: profile.email }], "m1");
  var packed = pack_(room, personId);
  return { token: token, person_id: personId, room: packed.room, people: packed.people, you: packed.you, messages: packed.messages };
}

function join_(r) {
  var code = String(r.code || "").trim().toUpperCase();
  var room = rows_("Rooms").filter(function (x) { return x.code === code; })[0];
  if (!room) throw new Error("No table with that code.");
  if (room.status === "cancelled") throw new Error("That table was cancelled.");
  var profile = profile_(r);
  requireUser_(r, profile.email);
  upsertUser_(profile);
  var all = rows_("People").filter(function (x) { return x.room_id === room.room_id; });
  var existing = all.filter(function (p) { return p.email === profile.email; })[0];
  var active = all.filter(function (p) { return String(p.left) !== "1"; });
  if (existing && String(existing.left) !== "1") {
    var again = Utilities.getUuid();
    append_("Tokens", { room_id: room.room_id, person_id: existing.person_id, token: again, is_host: existing.person_id === room.host_person_id ? "1" : "" });
    var back = pack_(room, existing.person_id);
    return { token: again, person_id: existing.person_id, room: back.room, people: back.people, you: back.you, messages: back.messages, rejoined: true };
  }
  if (active.length >= MAX_PEOPLE) throw new Error("This table is full.");
  if (existing) {
    setCells_("People", existing._row, { left: "" });
    var backToken = Utilities.getUuid();
    append_("Tokens", { room_id: room.room_id, person_id: existing.person_id, token: backToken, is_host: existing.person_id === room.host_person_id ? "1" : "" });
    var returned = pack_(room, existing.person_id);
    return { token: backToken, person_id: existing.person_id, room: returned.room, people: returned.people, you: returned.you, messages: returned.messages, rejoined: true };
  }
  var personId = Utilities.getUuid();
  var token = Utilities.getUuid();
  append_("People", personRow_(room.room_id, personId, profile, room.activity, now_()));
  append_("Tokens", { room_id: room.room_id, person_id: personId, token: token, is_host: "" });
  var state = pack_(room, personId);
  return { token: token, person_id: personId, room: state.room, people: state.people, you: state.you, messages: state.messages };
}

function getState_(r) {
  var found = find_(r.code, r.token);
  return pack_(found.room, found.person.person_id);
}

function save_(r) {
  var found = find_(r.code, r.token);
  if (found.room.status === "cancelled") throw new Error("This table was cancelled.");
  var patch = {};
  if (r.intro != null) patch.intro = clip_(r.intro, 120);
  if (r.capsule != null) patch.capsule = clip_(r.capsule, 120);
  if (r.wish != null) patch.wish = clip_(r.wish, 80);
  if (r.bye != null) patch.bye = clip_(r.bye, 40);
  if (r.slot != null) patch.slot = iso_(r.slot) || clip_(r.slot, 40);
  if (r.vote != null && r.vote !== "") patch.vote = String(r.vote) === "1" ? "1" : "0";
  if (r.contact != null) patch.contact = clip_(r.contact, 80);
  if (r.guess_target != null) patch.guess_target = clip_(r.guess_target, 80);
  if (r.guess_topic != null) patch.guess_topic = clip_(r.guess_topic, 80);
  if (r.guess_text != null) patch.guess_text = clip_(r.guess_text, 80);
  if (r.next_activity != null) patch.next_activity = clip_(r.next_activity, 40);
  if (r.meet3_at != null) patch.meet3_at = iso_(r.meet3_at) || "";
  if (r.meet3_activity != null) patch.meet3_activity = clip_(r.meet3_activity, 40);
  if (!Object.keys(patch).length) throw new Error("Nothing to save.");
  setCells_("People", found.person._row, patch);
  var room = rows_("Rooms").filter(function (x) { return x.room_id === found.room.room_id; })[0];
  return pack_(room, found.person.person_id);
}

function updateRoom_(r) {
  var found = find_(r.code, r.token);
  assertHost_(found);
  if (found.room.status === "cancelled") throw new Error("This table was cancelled.");
  var patch = {};
  if (r.activity != null) {
    var activity = clip_(r.activity, 40);
    if (activity.length < 2) throw new Error("Name the activity.");
    patch.activity = activity;
  }
  if (r.place != null) {
    var place = clip_(r.place, 80);
    if (place.length < 2) throw new Error("Add where you meet.");
    patch.place = place;
  }
  if (r.whenAt != null) {
    var whenAt = iso_(r.whenAt);
    if (!whenAt) throw new Error("Add a date and time.");
    if (new Date(whenAt).getTime() < Date.now() - 2 * 60 * 1000) throw new Error("Pick a time that has not already passed.");
    patch.when_at = whenAt;
    patch.when_text = clip_(r.when, 60) || prettyWhen_(whenAt);
  }
  if (!Object.keys(patch).length) throw new Error("Nothing to change.");
  setCells_("Rooms", found.room._row, patch);
  var room = rows_("Rooms").filter(function (x) { return x.room_id === found.room.room_id; })[0];
  return pack_(room, found.person.person_id);
}

function cancelRoom_(r) {
  var found = find_(r.code, r.token);
  assertHost_(found);
  setCells_("Rooms", found.room._row, { status: "cancelled" });
  var room = rows_("Rooms").filter(function (x) { return x.room_id === found.room.room_id; })[0];
  return pack_(room, found.person.person_id);
}

function leaveRoom_(r) {
  var found = find_(r.code, r.token);
  setCells_("People", found.person._row, { left: "1" });
  if (isHost_(found)) setCells_("Rooms", found.room._row, { status: "cancelled" });
  return { left: true };
}

function listTables_(r) {
  var user = userByToken_(r.userToken);
  var profile = rows_("Users").filter(function (u) { return u.email === user.email; })[0];
  var mine = softAnswers_(r.answers);
  if (!Object.keys(mine).length && profile) mine = parseAnswers_(profile.answers_json);
  var tables = [];
  rows_("Rooms").forEach(function (room) {
    if (room.status === "cancelled" || !room.code) return;
    if (room.when_at) {
      var at = new Date(room.when_at).getTime();
      if (!isNaN(at) && at < Date.now() - 2 * 60 * 60 * 1000) return;
    }
    var people = activePeople_(room.room_id);
    if (!people.length || people.length >= MAX_PEOPLE) return;
    var yours = people.some(function (p) { return p.email === user.email; });
    tables.push({
      code: room.code,
      activity: room.activity,
      when: room.when_text,
      when_at: room.when_at,
      place: room.place,
      seats: MAX_PEOPLE - people.length,
      yours: yours,
      people: people.map(function (p) { return { name: p.name, from: p.from_place }; }),
      share: shareLine_(mine, people, user.email),
    });
  });
  tables.sort(function (a, b) { return String(a.when_at || "9999").localeCompare(String(b.when_at || "9999")); });
  return { tables: tables.slice(0, 12) };
}

function postMessage_(r) {
  var found = find_(r.code, r.token);
  if (found.room.status === "cancelled") throw new Error("This table was cancelled.");
  var text = clip_(r.text, 200);
  if (!text) throw new Error("Write a message first.");
  append_("Messages", {
    room_id: found.room.room_id,
    message_id: Utilities.getUuid(),
    person_id: found.person.person_id,
    name: found.person.name,
    text: text,
    at: now_(),
  });
  return pack_(found.room, found.person.person_id);
}

function judge_(r) {
  var found = find_(r.code, r.token);
  var guesserId = String(r.guesserId || "");
  var verdict = r.verdict === "right" ? "right" : r.verdict === "wrong" ? "wrong" : "";
  if (!verdict) throw new Error("Say if it was right or wrong.");
  var guesser = rows_("People").filter(function (p) {
    return p.person_id === guesserId && p.room_id === found.room.room_id;
  })[0];
  if (!guesser || guesser.guess_target !== found.person.person_id) throw new Error("That guess is not about you.");
  setCells_("People", guesser._row, { guess_verdict: verdict });
  return pack_(found.room, found.person.person_id);
}

function icebreaker_(r) {
  var found = find_(r.code, r.token);
  if (found.room.status === "cancelled") throw new Error("This table was cancelled.");
  if (!r.refresh && found.room.icebreakers) {
    try {
      var cached = JSON.parse(found.room.icebreakers);
      if (cached && cached.length) return { prompts: cached, cached: true };
    } catch (e) {}
  }
  var people = rows_("People").filter(function (x) { return x.room_id === found.room.room_id; });
  var made = claudeIcebreakers_(found.room, people, r.shares || []);
  if (made.prompts.length) {
    lock_(function () { setCells_("Rooms", found.room._row, { icebreakers: JSON.stringify(made.prompts) }); });
  }
  return made;
}

/** Run once from the editor to allow Strawberry to send email. */
function testMail() {
  var email = Session.getActiveUser().getEmail();
  MailApp.sendEmail(email, "Strawberry mail test", "Mail from Strawberry works.");
  Logger.log("Sent to " + email);
}

/** Run once from the editor after adding ANTHROPIC_API_KEY. Check the log for 3 prompts. */
function testIcebreaker() {
  var made = claudeIcebreakers_(
    { activity: "Boba crawl", place: "Telegraph", when_text: "Thursday 5:30" },
    [{ name: "Alex", from_place: "Taipei", standing: "exchange" }, { name: "Jordan", from_place: "Oakland", standing: "undergrad" }],
    ["Boba"]
  );
  Logger.log(JSON.stringify(made));
}

function claudeIcebreakers_(room, people, shares) {
  var key = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");
  if (!key) return { prompts: [], error: "Add ANTHROPIC_API_KEY in Apps Script, then deploy again." };
  var model = PropertiesService.getScriptProperties().getProperty("MODEL") || "claude-haiku-5-5";
  var table = people.map(function (p) {
    return { name: p.name, from: p.from_place, standing: p.standing };
  });
  var prompt = "A small group of UC Berkeley students is together to do one activity, not to date or swap profiles.\n"
    + "Activity: " + room.activity + "\nPlace: " + (room.place || "") + "\nWhen: " + (room.when_text || "") + "\n"
    + "People: " + JSON.stringify(table) + "\n"
    + "Things they marked in common: " + JSON.stringify(shares || []) + "\n\n"
    + "Write 3 icebreakers they can read out loud while doing the activity.\n"
    + "Rules:\n"
    + "- Every line must use this activity, a place someone is from, or something they share. No generic 'what is your hobby'.\n"
    + "- Ask for a short story, not a fact quiz.\n"
    + "- One line should make it easy to talk about home without interviewing anyone.\n"
    + "- Under 22 words each. Warm and specific.\n"
    + "- Do not ask about politics, gender, religion, or age. Also skip stereotypes, dating, money, and appearance.\n"
    + "Reply with only JSON: {\"prompts\":[\"...\",\"...\",\"...\"]}";
  var data = claudeMessage_(key, model, prompt, true);
  if (data.error) data = claudeMessage_(key, model, prompt, false);
  if (data.error) return { prompts: [], error: data.error };
  var blocks = data.content || [];
  var text = blocks.filter(function (b) { return b && b.type === "text" && b.text; }).map(function (b) { return b.text; }).join("\n");
  var prompts = promptsFrom_(text);
  if (!prompts.length) {
    var types = blocks.map(function (b) { return b && b.type; }).join(",") || "none";
    return { prompts: [], error: "No icebreaker came back. stop=" + (data.stop_reason || "") + " types=" + types };
  }
  return { prompts: prompts };
}

function claudeMessage_(key, model, prompt, disableThinking) {
  var body = { model: model, max_tokens: disableThinking ? 1200 : 4000, messages: [{ role: "user", content: prompt }] };
  if (disableThinking) body.thinking = { type: "disabled" };
  var res = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: { "x-api-key": key, "anthropic-version": "2023-06-01" },
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  var raw = res.getContentText();
  if (res.getResponseCode() !== 200) return { error: "The icebreaker model did not answer (" + res.getResponseCode() + "): " + raw.slice(0, 180) };
  return JSON.parse(raw);
}

function promptsFrom_(text) {
  var found = [];
  var match = String(text || "").match(/\{[\s\S]*\}/);
  if (match) {
    try {
      var obj = JSON.parse(match[0]);
      var list = obj.prompts || obj.icebreakers || obj.questions || [];
      if (!list.length && obj[0]) list = obj;
      found = list.map(function (p) { return typeof p === "string" ? p : (p && (p.text || p.prompt)) || ""; });
    } catch (e) {}
  }
  if (!found.length) {
    found = String(text || "").split("\n").map(function (line) { return line.replace(/^\s*(?:\d+[\.\)]|[-*])\s*/, ""); });
  }
  return found.filter(function (p) { return p && String(p).trim().length > 8; }).slice(0, 3).map(function (p) { return clip_(p, 160); });
}

function checkin_(r) {
  var found = find_(r.code, r.token);
  var why = clip_(r.connectWhy || r.why, 160);
  if (!why) throw new Error("Add one line on why.");
  var again = r.again === true || r.again === "Yes" || r.again === "true" ? "Yes" : "No";
  var met = r.met === true || r.met === "Yes" || r.met === "true" ? "Yes" : "No";
  var reasons = clip_(Array.isArray(r.reasons) ? r.reasons.join(", ") : r.reasons, 300);
  var connected = clip_(Array.isArray(r.connected) ? r.connected.join(", ") : r.connected, 200);
  if (!reasons) throw new Error("Pick at least one reason.");
  if (!connected) throw new Error("Say who you connected with.");
  var existing = rows_("Checkins").filter(function (x) {
    return x.room_id === found.room.room_id && x.person_id === found.person.person_id;
  })[0];
  var row = {
    room_id: found.room.room_id,
    person_id: found.person.person_id,
    name: found.person.name,
    again: again,
    met: met,
    reasons: reasons,
    connected: connected,
    why: why,
    group_name: clip_(r.group, 80),
    at: now_(),
  };
  if (existing) setCells_("Checkins", existing._row, row);
  else append_("Checkins", row);
  return { saved: true };
}

function personRow_(roomId, personId, profile, activity, joined) {
  return {
    room_id: roomId,
    person_id: personId,
    joined_at: joined,
    name: profile.name,
    email: profile.email,
    from_place: profile.from,
    standing: profile.standing,
    hook: activity,
    home: profile.home,
    secret: profile.secret,
    answers_json: JSON.stringify(profile.answers),
    intro: "",
    capsule: "",
    wish: "",
    vote: "",
    slot: "",
    bye: "",
  };
}

function emailOk_(email) {
  if (/^[^\s@]+@berkeley\.edu$/.test(email)) return true;
  return ALLOW_GMAIL && /^[^\s@]+@gmail\.com$/.test(email);
}

function normEmail_(value) {
  return String(value || "").trim().toLowerCase();
}

function sendCode_(r) {
  var email = normEmail_(r.email);
  if (!emailOk_(email)) throw new Error(ALLOW_GMAIL ? "Use a @berkeley.edu or Gmail address." : "Use your @berkeley.edu email.");
  var cache = CacheService.getScriptCache();
  if (cache.get("sent:" + email)) throw new Error("Wait a minute, then ask for a new code.");
  var code = ("000" + Math.floor(Math.random() * 10000)).slice(-4);
  cache.put("code:" + email, code, 600);
  cache.put("tries:" + email, "0", 600);
  cache.put("sent:" + email, "1", 60);
  MailApp.sendEmail(email, "Your Strawberry code", "Your code is " + code + ".\n\nIt expires in 10 minutes. If you did not ask for this, ignore this email.");
  return { sent: true };
}

function verifyCode_(r) {
  var email = normEmail_(r.email);
  if (!emailOk_(email)) throw new Error(ALLOW_GMAIL ? "Use a @berkeley.edu or Gmail address." : "Use your @berkeley.edu email.");
  var code = String(r.code || "").replace(/\D/g, "");
  var cache = CacheService.getScriptCache();
  var tries = Number(cache.get("tries:" + email) || "0");
  if (tries >= 5) throw new Error("Too many tries. Ask for a new code.");
  var saved = cache.get("code:" + email);
  if (!saved || saved !== code) {
    cache.put("tries:" + email, String(tries + 1), 600);
    throw new Error("That code does not match. Check the email and try again.");
  }
  cache.remove("code:" + email);
  var token = Utilities.getUuid();
  append_("UserTokens", { email: email, token: token, created_at: now_() });
  var user = rows_("Users").filter(function (u) { return u.email === email; })[0];
  return { token: token, email: email, profile: user ? publicUser_(user) : null };
}

function me_(r) {
  var row = userByToken_(r.userToken);
  var user = rows_("Users").filter(function (u) { return u.email === row.email; })[0];
  return { email: row.email, profile: user ? publicUser_(user) : null };
}

function userByToken_(token) {
  token = String(token || "");
  var row = rows_("UserTokens").filter(function (x) { return x.token === token; })[0];
  if (!row) throw new Error("Sign in again, then come back to the table.");
  return row;
}

function requireUser_(r, email) {
  var row = userByToken_(r.userToken);
  if (row.email !== email) throw new Error("Sign in again, then come back to the table.");
  return row;
}

function publicUser_(u) {
  return {
    name: u.name,
    from: u.from_place,
    standing: u.standing,
    home: u.home,
    answers: parseAnswers_(u.answers_json),
  };
}

function upsertUser_(profile) {
  var existing = rows_("Users").filter(function (u) { return u.email === profile.email; })[0];
  var row = {
    email: profile.email,
    name: profile.name,
    from_place: profile.from,
    standing: profile.standing,
    home: profile.home,
    answers_json: JSON.stringify(profile.answers || {}),
    updated_at: now_(),
  };
  if (existing) setCells_("Users", existing._row, row);
  else append_("Users", row);
}

function profile_(r) {
  var name = clip_(r.name, 40);
  if (name.length < 2) throw new Error("Add your name.");
  if (name.toLowerCase() === "you") throw new Error("Pick a name other than You.");
  var email = String(r.email || "").trim().toLowerCase();
  if (!emailOk_(email)) throw new Error(ALLOW_GMAIL ? "Use a @berkeley.edu or Gmail address." : "Use your @berkeley.edu email.");
  var from = clip_(r.from, 40);
  if (from.length < 2) throw new Error("Add where you're from.");
  if (STANDINGS.indexOf(r.standing) < 0) throw new Error("Say if you're an undergrad, an exchange student, or a grad student.");
  if (HOMES.indexOf(r.home) < 0) throw new Error("Pick what you want to talk about from home.");
  var secret = clip_(r.secret, 80);
  return {
    name: name, email: email, from: from, standing: r.standing, home: r.home, secret: secret,
    answers: cleanAnswers_(r.answers),
  };
}

function cleanAnswers_(raw) {
  var answers = {};
  var src = raw || {};
  Object.keys(CARD_KEYS).forEach(function (id) {
    var key = String(src[id] || "");
    if (CARD_KEYS[id].indexOf(key) < 0) throw new Error("Answer every card before joining.");
    answers[id] = key;
  });
  return answers;
}

function activePeople_(roomId) {
  return rows_("People").filter(function (p) { return p.room_id === roomId && String(p.left) !== "1"; });
}

function isHost_(found) {
  return String(found.tok.is_host) === "1" || found.room.host_person_id === found.person.person_id;
}

function assertHost_(found) {
  if (!isHost_(found)) throw new Error("Only the host can change this table.");
}

function softAnswers_(raw) {
  var answers = {};
  var src = raw || {};
  Object.keys(CARD_KEYS).forEach(function (id) {
    var key = String(src[id] || "");
    if (CARD_KEYS[id].indexOf(key) >= 0) answers[id] = key;
  });
  return answers;
}

function shareLine_(mine, people, email) {
  var keys = Object.keys(CARD_KEYS);
  for (var i = 0; i < keys.length; i++) {
    var id = keys[i];
    if (!mine[id]) continue;
    for (var j = 0; j < people.length; j++) {
      if (people[j].email === email) continue;
      var answers = parseAnswers_(people[j].answers_json);
      if (answers[id] && answers[id] === mine[id]) {
        return "You and " + people[j].name + " both like " + (OPTION_LABELS[mine[id]] || mine[id]) + ".";
      }
    }
  }
  return "You showed up for the same activity.";
}

function planKind_(kind) {
  if (kind === "meet3") return { at: "meet3_at", activity: "meet3_activity", place: "meet3_place", fixed: "meet3_fixed", ok: "meet3_ok", remind: "m3" };
  return { at: "meet2_at", activity: "meet2_activity", place: "meet2_place", fixed: "meet2_fixed", ok: "meet2_ok", remind: "m2" };
}

function proposePlan_(r) {
  var found = find_(r.code, r.token);
  if (found.room.status === "cancelled") throw new Error("This table was cancelled.");
  var spec = planKind_(r.kind);
  var activity = clip_(r.activity, 40);
  var whenAt = iso_(r.whenAt);
  if (activity.length < 2) throw new Error("Name the activity.");
  if (!whenAt) throw new Error("Add a date and time.");
  if (new Date(whenAt).getTime() < Date.now() - 2 * 60 * 1000) throw new Error("Pick a time that has not already passed.");
  if (String(found.room[spec.fixed]) === "1" && !r.reopen) throw new Error("That time is already set.");
  var roomPatch = {};
  roomPatch[spec.at] = whenAt;
  roomPatch[spec.activity] = activity;
  roomPatch[spec.place] = found.room.place || "";
  roomPatch[spec.fixed] = "";
  roomPatch.reminded = dropFlag_(found.room, spec.remind);
  setCells_("Rooms", found.room._row, roomPatch);
  activePeople_(found.room.room_id).forEach(function (p) {
    var mark = {};
    mark[spec.ok] = p.person_id === found.person.person_id ? "yes" : "";
    setCells_("People", p._row, mark);
  });
  settlePlan_(found.room.room_id, spec);
  var room = rows_("Rooms").filter(function (x) { return x.room_id === found.room.room_id; })[0];
  return pack_(room, found.person.person_id);
}

function answerPlan_(r) {
  var found = find_(r.code, r.token);
  if (found.room.status === "cancelled") throw new Error("This table was cancelled.");
  var spec = planKind_(r.kind);
  if (!found.room[spec.at]) throw new Error("Someone still has to suggest a time.");
  if (String(found.room[spec.fixed]) === "1") throw new Error("That time is already set.");
  var answer = r.answer === "yes" ? "yes" : r.answer === "no" ? "no" : "";
  if (!answer) throw new Error("Say if this time works.");
  var mark = {};
  mark[spec.ok] = answer;
  setCells_("People", found.person._row, mark);
  settlePlan_(found.room.room_id, spec);
  var room = rows_("Rooms").filter(function (x) { return x.room_id === found.room.room_id; })[0];
  return pack_(room, found.person.person_id);
}

function settlePlan_(roomId, spec) {
  var room = rows_("Rooms").filter(function (x) { return x.room_id === roomId; })[0];
  if (!room || !room[spec.at]) return;
  var people = activePeople_(roomId);
  var ready = people.length >= 2 && people.every(function (p) { return p[spec.ok] === "yes"; });
  if (ready && String(room[spec.fixed]) !== "1") {
    var patch = {};
    patch[spec.fixed] = "1";
    setCells_("Rooms", room._row, patch);
    room[spec.fixed] = "1";
    remind_(room, people, spec.remind);
  }
}

function dropFlag_(room, flag) {
  return String(room.reminded || "").split(",").filter(function (x) { return x && x !== flag; }).join(",");
}

function remind_(room, people, kind) {
  var already = String(room.reminded || "").split(",").filter(Boolean);
  if (already.indexOf(kind) >= 0) return;
  var when = (kind === "m1" || kind === "soon1") ? (room.when_text || prettyWhen_(room.when_at))
    : (kind === "m2" || kind === "soon2") ? prettyWhen_(room.meet2_at)
    : prettyWhen_(room.meet3_at);
  var what = (kind === "m2" || kind === "soon2") ? (room.meet2_activity || room.activity)
    : (kind === "m3" || kind === "soon3") ? (room.meet3_activity || room.activity)
    : room.activity;
  var body = "See you " + when + ".\n" + (what || "") + (room.place ? "\n" + room.place : "") + "\n\nOpen Strawberry at that time. This email is your reminder.";
  (people || []).forEach(function (p) {
    if (!p.email) return;
    try { MailApp.sendEmail(p.email, "Strawberry: see you " + when, body); } catch (e) {}
  });
  already.push(kind);
  if (room._row) setCells_("Rooms", room._row, { reminded: already.join(",") });
  room.reminded = already.join(",");
}

function prettyWhen_(iso) {
  var d = new Date(iso);
  if (!iso || isNaN(d.getTime())) return String(iso || "");
  return Utilities.formatDate(d, "America/Los_Angeles", "EEE, MMM d, h:mm a");
}

function iso_(value) {
  var s = String(value || "").trim();
  if (!s) return "";
  var d = new Date(s);
  if (isNaN(d.getTime())) return "";
  return d.toISOString();
}

function installReminders_() {
  var triggers = ScriptApp.getProjectTriggers();
  var has = triggers.some(function (t) { return t.getHandlerFunction() === "sendDueReminders"; });
  if (!has) ScriptApp.newTrigger("sendDueReminders").timeBased().everyHours(1).create();
}

function sendDueReminders() {
  var now = Date.now();
  rows_("Rooms").forEach(function (room) {
    if (room.status === "cancelled") return;
    var people = activePeople_(room.room_id);
    [["soon1", room.when_at], ["soon2", room.meet2_at], ["soon3", room.meet3_at]].forEach(function (pair) {
      if (!pair[1]) return;
      var t = new Date(pair[1]).getTime();
      if (isNaN(t)) return;
      var delta = t - now;
      if (delta > 0 && delta <= 90 * 60 * 1000) remind_(room, people, pair[0]);
    });
  });
}

function roomMessages_(room, people) {
  var list = rows_("Messages").filter(function (m) { return m.room_id === room.room_id; });
  list.sort(function (a, b) { return String(a.at).localeCompare(String(b.at)); });
  if (!list.length) {
    people.forEach(function (p) {
      if (p.intro) list.push({ person_id: p.person_id, name: p.name, text: p.intro, at: p.joined_at || "" });
    });
  }
  return list.slice(-40).map(function (m) {
    return { person_id: m.person_id, name: m.name, text: m.text, at: m.at };
  });
}

function pack_(room, viewerId) {
  var people = activePeople_(room.room_id);
  people.sort(function (a, b) { return String(a.joined_at).localeCompare(String(b.joined_at)); });
  var names = people.map(function (p) { return p.name; });
  var viewer = null;
  var others = [];
  var checked = rows_("Checkins").some(function (c) {
    return c.room_id === room.room_id && c.person_id === viewerId;
  });
  people.forEach(function (p, i) {
    var display = p.name;
    if (names.filter(function (n) { return n === p.name; }).length > 1) display = p.name + " " + String(p.person_id).slice(0, 4);
    var pub = {
      person_id: p.person_id,
      name: display,
      from: p.from_place,
      standing: p.standing,
      hook: p.hook || room.activity,
      color: COLORS[i % COLORS.length],
      answers: parseAnswers_(p.answers_json),
      intro: p.intro,
      wish: p.wish,
      vote: p.vote,
      slot: p.slot,
      bye: p.bye,
      has_capsule: p.capsule ? true : false,
      capsule: "",
      contact: checked ? (p.contact || "") : "",
      guess_target: p.guess_target || "",
      guess_topic: p.guess_topic || "",
      guess_text: p.guess_text || "",
      guess_verdict: p.guess_verdict || "",
      next_activity: p.next_activity || "",
      meet3_at: p.meet3_at || "",
      meet3_activity: p.meet3_activity || "",
      meet2_ok: p.meet2_ok || "",
      meet3_ok: p.meet3_ok || "",
    };
    if (p.person_id === viewerId) viewer = p;
    else others.push(pub);
  });
  if (!viewer) throw new Error("This phone is not in that table.");
  if (viewer.capsule) {
    others.forEach(function (pub) {
      var row = people.filter(function (p) { return p.person_id === pub.person_id; })[0];
      pub.capsule = row && row.capsule ? row.capsule : "";
    });
  }
  return {
    room: {
      code: room.code, activity: room.activity, when: room.when_text, place: room.place,
      when_at: room.when_at || "", status: room.status || "",
      meet2_activity: room.meet2_activity || "", meet2_place: room.meet2_place || "", meet2_at: room.meet2_at || "", meet2_fixed: room.meet2_fixed || "",
      meet3_activity: room.meet3_activity || "", meet3_place: room.meet3_place || "", meet3_at: room.meet3_at || "", meet3_fixed: room.meet3_fixed || "",
    },
    you: {
      person_id: viewer.person_id,
      name: viewer.name,
      from: viewer.from_place,
      home: viewer.home,
      standing: viewer.standing,
      answers: parseAnswers_(viewer.answers_json),
      intro: viewer.intro,
      capsule: viewer.capsule,
      wish: viewer.wish,
      vote: viewer.vote,
      slot: viewer.slot,
      bye: viewer.bye,
      contact: viewer.contact || "",
      guess_target: viewer.guess_target || "",
      guess_topic: viewer.guess_topic || "",
      guess_text: viewer.guess_text || "",
      guess_verdict: viewer.guess_verdict || "",
      next_activity: viewer.next_activity || "",
      meet3_at: viewer.meet3_at || "",
      meet3_activity: viewer.meet3_activity || "",
      meet2_ok: viewer.meet2_ok || "",
      meet3_ok: viewer.meet3_ok || "",
      is_host: room.host_person_id === viewer.person_id,
    },
    people: others,
    messages: roomMessages_(room, people),
  };
}

function find_(code, token) {
  code = String(code || "").trim().toUpperCase();
  token = String(token || "");
  var room = rows_("Rooms").filter(function (x) { return x.code === code; })[0];
  if (!room) throw new Error("No table with that code.");
  var tok = rows_("Tokens").filter(function (x) { return x.token === token && x.room_id === room.room_id; })[0];
  if (!tok) throw new Error("This phone is not in that table. Join again.");
  var person = rows_("People").filter(function (x) { return x.person_id === tok.person_id && x.room_id === room.room_id; })[0];
  if (!person) throw new Error("This phone is not in that table. Join again.");
  return { room: room, person: person, tok: tok };
}

function parseAnswers_(text) {
  try { return JSON.parse(text || "{}"); } catch (e) { return {}; }
}

function newCode_() {
  var taken = {};
  rows_("Rooms").forEach(function (x) { taken[x.code] = true; });
  var code = "";
  do {
    code = "";
    for (var i = 0; i < 4; i++) code += CODE_CHARS.charAt(Math.floor(Math.random() * CODE_CHARS.length));
  } while (taken[code]);
  return code;
}

function ss_() {
  var id = PropertiesService.getScriptProperties().getProperty("SHEET_ID");
  if (id) return SpreadsheetApp.openById(id);
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error("Run setup once from the script editor.");
  return active;
}

function sheet_(name) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  var head = TABS[name];
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold");
    sh.setFrozenRows(1);
    if (name === "Tokens" || name === "UserTokens") sh.hideSheet();
  } else {
    var width = Math.max(sh.getLastColumn(), 1);
    var first = sh.getRange(1, 1, 1, width).getValues()[0];
    var dirty = false;
    for (var h = 0; h < head.length; h++) if (String(first[h] || "") !== head[h]) dirty = true;
    if (dirty) sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold");
  }
  return sh;
}

function rows_(name) {
  var sh = sheet_(name);
  var values = sh.getDataRange().getValues();
  var head = TABS[name];
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = { _row: i + 1 };
    var empty = true;
    for (var c = 0; c < head.length; c++) {
      var v = values[i][c];
      row[head[c]] = v == null ? "" : String(v);
      if (row[head[c]] !== "") empty = false;
    }
    if (!empty) out.push(row);
  }
  return out;
}

function append_(name, obj) {
  var head = TABS[name];
  sheet_(name).appendRow(head.map(function (h) { return obj[h] == null ? "" : obj[h]; }));
}

function setCells_(name, row, patch) {
  var sh = sheet_(name);
  var head = TABS[name];
  Object.keys(patch).forEach(function (k) {
    var c = head.indexOf(k);
    if (c >= 0) sh.getRange(row, c + 1).setValue(patch[k]);
  });
}

function lock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function clip_(s, n) {
  return String(s == null ? "" : s).replace(/\s+/g, " ").trim().slice(0, n);
}

function now_() {
  return new Date().toISOString();
}
