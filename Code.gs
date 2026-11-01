/**
 * Inventory Tracker — Google Apps Script backend.
 *
 * Uses the existing "ES Full Inventory Counts" spreadsheet as the database:
 *   "Full Inventory Counts Master Sheet" — one row per item (per assigned job).
 *   The app reads/writes "Available Today" as the on-hand quantity.
 *   "App Log" tab (created by the app) — audit trail of every change made from the phone.
 *   "Item Comments" tab (created by the app) — comments left on item pages, one row per comment.
 *   "Item Documents" tab (created by the app) — PDFs attached to items, one row per file.
 * Attached PDFs are stored in a Drive folder named "Inventory Documents".
 *   "Orders" tab (created by the app) — one row per order (Draft / Shipped / Cancelled).
 *   "Order Lines" tab (created by the app) — one row per item on an order.
 *   "Order Documents" tab (created by the app) — PDFs attached to orders; the files live in a Drive folder "Order Documents".
 *   "Order Photos" tab (created by the app) — photos attached to orders; the files live in a Drive folder "Order Photos".
 *   "Order Links" tab (created by the app) — web links saved on orders (tracking, quotes, portals…), one row per link.
 * Item photos are stored as files in a Drive folder named "Inventory Photos".
 *
 * Columns are found by their header text, so inserting new monthly
 * "Starting Qty" columns doesn't break anything (the app adds them itself: see "Monthly starting counts").
 */

const SPREADSHEET_ID = '1YXqhyazHSftJiWo4IsorNAjrT_oO21TkayfxepZ9d6U';
const MASTER_SHEET_ID = 1463870932;            // "Full Inventory Counts Master Sheet" (by ID, so renaming the tab is safe)
const LOG_SHEET = 'App Log';
const PHOTO_FOLDER = 'Inventory Photos';
const COMMENTS_SHEET = 'Item Comments';
const DOCS_SHEET = 'Item Documents';
const ORDERS_SHEET = 'Orders';
const LINES_SHEET = 'Order Lines';
// Order fields as typed in the app -> column headers on the Orders tab.
const ORDER_FIELDS = [
  ['orderDate', 'Order Date'], ['projectName', 'Project Name'], ['projectNo', 'Project #'], ['activityType', 'Activity Type'],
  ['fromLocation', 'From Location'], ['toLocation', 'To Location'], ['contact', 'Point of Contact'],
  ['freight', 'Freight Service Company'], ['vendor', 'Company / Vendor'], ['pm', 'PM'], ['notes', 'Notes'],
];
// Columns added later go at the END of the Orders tab, so tabs created by an earlier version keep lining up.
const ORDER_EXTRA_FIELDS = [['trackingUrl', 'Tracking Link'], ['shippingCost', 'Shipping Cost'], ['shipTo', 'Ship To Address'], ['tags', 'Tags'], ['endpoints', 'Endpoints'], ['deliveryDate', 'Delivery Date'], ['shipDate', 'Pickup / Ship Date'], ['pmNote', 'PM Note Sent'], ['tracker', 'Tracker Logged']];
const MAX_ADDRESS = 300;
const MAX_TAGS = 20, MAX_TAG_LENGTH = 30;
const ORDER_HEADERS = ['Order #', 'Status'].concat(ORDER_FIELDS.map(f => f[1]),
  ['Total Ship Qty', 'Our Cost', 'Lines', 'Created', 'Created By', 'Updated', 'Updated By', 'Shipped', 'Shipped By', 'Cancelled', 'Cancelled By'],
  ORDER_EXTRA_FIELDS.map(f => f[1]));
const ORDER_DOCS_SHEET = 'Order Documents';
const ORDER_DOCS_FOLDER = 'Order Documents';
const ORDER_DOC_HEADERS = ['Timestamp', 'User', 'Order #', 'Project', 'File Name', 'Size (KB)', 'Drive File ID', 'Link'];
const ORDER_PHOTOS_SHEET = 'Order Photos';
const ORDER_PHOTOS_FOLDER = 'Order Photos';
// Position (added later, so it's the last column) is the order photos are shown in; blank = after the arranged ones, oldest first.
const ORDER_PHOTO_HEADERS = ['Timestamp', 'User', 'Order #', 'Project', 'Drive File ID', 'Link', 'Position'];
const MAX_ORDER_PHOTOS = 30;
const ORDER_LINKS_SHEET = 'Order Links';
// Position (added later, so it's the last column) is the order links are shown in; blank = after the arranged ones, oldest first.
const ORDER_LINK_HEADERS = ['Timestamp', 'User', 'Order #', 'Label', 'URL', 'Link ID', 'Position'];
const MAX_ORDER_LINKS = 50;
const ORDER_COSTS_SHEET = 'Order Costs';
const ORDER_COST_HEADERS = ['Timestamp', 'User', 'Order #', 'Description', 'Amount', 'Cost ID'];
const MAX_ORDER_COSTS = 50;
// Endpoints: the endpoints that line's item goes to (comma-separated); the order's own Endpoints column sums up its lines.
// Endpoint Serials: each endpoint's serial #, in the same order (blank between commas when an endpoint has none).
// Each endpoint's optional details, same comma format as Endpoint Serials (blank between commas when not entered):
// rotary dial setting, amp draw, breaker size, and the other endpoint on the same order it load-shares with.
// [field in the app, column header, max length, name in messages]
const EP_INFO = [['rotary', 'Endpoint Rotary Dial', 10, 'Rotary dial'], ['amps', 'Endpoint Amp Draw', 10, 'Amp draw'],
  ['breaker', 'Endpoint Breaker Size', 10, 'Breaker size'], ['share', 'Endpoint Load Sharing', 30, 'Load sharing']];
const LINE_HEADERS = ['Order #', 'Item Name', 'Part Number', 'Assigned Job', 'Ship Quantity', 'Cost Price', 'Our Cost', 'Item Key', 'Endpoints', 'Endpoint Serials']
  .concat(EP_INFO.map(f => f[1]));
const MAX_SERIAL = 60;
const DOCS_FOLDER = 'Inventory Documents';
// Position (added later, so it's the last column) is the order an item's PDFs are shown in; blank = after the arranged ones, newest first.
const DOC_HEADERS = ['Timestamp', 'User', 'Item Name', 'Part Number', 'Assigned Job', 'File Name', 'Size (KB)', 'Drive File ID', 'Item Key', 'Link', 'Position'];
const MAX_DOC_MB = 10;   // per PDF; larger uploads from a phone tend to time out
// "Item Key" links a comment to its item (name | part # | job); it is updated if the item is renamed in the app.
const COMMENT_HEADERS = ['Timestamp', 'User', 'Item Name', 'Part Number', 'Assigned Job', 'Comment', 'Item Key', 'Comment ID'];
const MAX_COMMENT = 2000;   // the box allows 1000 typed characters; @mentions are saved as longer @{email} tags
// Every stock change made in the app (the Adjust stock form and the quick +/− buttons), with vendor and PO # when given.
const ADJ_SHEET = 'Inventory Adjustments';
// Returned From / Order # / Return Reason / Point of Contact are filled in for returning inventory (see receiveReturn);
// Point of Contact (added later, so it's the last column) is who brought it back, picked from Contacts in the app.
// Endpoints / Endpoint Serials: endpoints entered for an EVSE item while receiving it (same format as Order Lines).
const ADJ_HEADERS = ['Timestamp', 'User', 'Type', 'Item Name', 'Part Number', 'Assigned Job', 'Change', 'Qty Before', 'Qty After',
  'Vendor', 'PO #', 'Note', 'Item Key', 'Adjustment ID', 'Returned From', 'Order #', 'Return Reason', 'Endpoints', 'Endpoint Serials', 'Point of Contact', 'Tracker Logged', 'PM', 'Links', 'PDFs'];
const LOG_HEADERS = ['Timestamp', 'User', 'Action', 'Row', 'Item Name', 'Part Number', 'Assigned Job', 'Change', 'New Qty', 'Note'];

// App field -> header text in the master sheet (or a list of accepted names, first one preferred).
// Matching ignores case and extra spaces.
const FIELDS = {
  name: 'Item Name',
  category: 'Catagory',
  location: 'Location',
  shelf: 'Shelf',
  voltage: 'Voltage',
  phase: 'Phase',
  amperage: ['Amperage', 'Amprage'],   // the sheet has used both spellings
  partNumber: 'Part Number/SKEW',
  manufacturer: 'Manufacturer',
  job: 'Assigned Job',
  quantity: 'Available Today',
  rating: 'Rating',
  dimensions: 'Dimensions (Box)',
  weight: 'Approx. Weight',
  connector: 'Connector Type',
  cost: 'Cost Price',          // the unit price: only Admins can change it in the app (saveItem)
  extended: 'Extended Price',  // sheet formula: Cost Price × Available Today
  barcode: ['Barcode / QR Code', 'Barcode/QR Code', 'Barcode'],
  photoId: 'Photo ID',
};
// Fields setup() adds if missing, using the first empty placeholder table column ("Column 4", "Column 5", ...).
const ADDED_FIELDS = ['photoId', 'barcode'];
const SPARE_COLUMN = /^column \d+$/;
// Fields the phone app is allowed to edit (everything else in the row is left alone).
const EDITABLE = ['name', 'category', 'location', 'shelf', 'partNumber', 'manufacturer', 'job', 'quantity', 'barcode',
  'voltage', 'phase', 'amperage', 'dimensions', 'weight'];
// Everything typed in the app except the quantity is stored as plain text, so Sheets can't turn shelf "8-1"
// into a date, "012345" into 12345, or "1/3" into January 3.
const TEXT_FIELDS = EDITABLE.filter(f => f !== 'quantity');

// ---------- Users & permissions ----------
// Only people on the "App Users" tab can open the app. Roles:
//   Admin  — everything, plus managing users        Editor — can change anything
//   Viewer — can look at everything but change nothing
// The app's owner (the account that deployed it) is always an Admin and doesn't need to be on the list.

const USERS_SHEET = 'App Users';
const USER_HEADERS = ['Email', 'Role', 'Name / Note', 'Added', 'Added By', 'Hidden Screens'];

// Screens an admin can hide from someone (Users & permissions). Admins always see every screen. Stored as the hidden
// ones ("orders, reports"), so a screen added later shows for everyone until an admin hides it.
const SCREENS = {
  all: 'Items', in: 'In Stock', scan: 'Scan', orders: 'Orders', reports: 'Reports', log: 'Activity',
  supply: 'Supply Chain requests', stock: 'Receive & adjust stock', contacts: 'Contacts',
};
const SCREEN_IDS_ = Object.keys(SCREENS);
const MAIN_SCREENS = ['all', 'in', 'orders', 'reports', 'log'];   // the bottom tabs: at least one must stay
function cleanHidden_(list) {
  const out = [...new Set((Array.isArray(list) ? list : String(list || '').split(',')).map(x => String(x).trim()).filter(x => SCREENS[x]))];
  if (MAIN_SCREENS.every(t => out.includes(t))) throw new Error('Leave at least one of Items, In Stock, Orders, Reports or Activity showing.');
  return out;
}
const ROLE_RANK = { Viewer: 1, Editor: 2, Manager: 3, Admin: 4 };
const ROLE_NAMES = { Admin: 'Admin', Manager: 'Manager', Editor: 'Editor', Viewer: 'View only' };
// 🧑‍💼 Managers: everything an Editor can do, plus managing the people THEY added (as View only or Editor). They never see
// or change anyone else on the list, can't touch an admin's (or another manager's) access, can't give anyone a screen
// that's hidden from themselves, and see nothing admin-only (every item, unit prices, monthly counts…).
const MANAGER_CAN_GIVE_ = ['Viewer', 'Editor'];
const isMyUser_ = u => !!u && String(u.addedBy || '').trim().toLowerCase() === currentUser_().toLowerCase();
/** Admins manage everyone; a manager only the View only / Editor people they added. Throws otherwise. */
function requireCanManage_(u) {
  if (!canManage_(u)) throw new Error(myRole_() === 'Admin' ? 'Only the app owner or the admin who added them can change another admin.' : !u || !isMyUser_(u) ? 'Managers can only change people they added.' : 'Only admins can change an admin’s or a manager’s access.');
}
/** 🛡 Can I change this person? The owner: anyone. An admin: anyone except another admin they didn't add.
 *  A manager: only the View only / Editor people they added. */
function canManage_(u) {
  if (!u) return false;
  if (currentUser_().toLowerCase() === ownerEmail_()) return true;
  if (myRole_() === 'Admin') return u.role !== 'Admin' || isMyUser_(u);
  return isMyUser_(u) && MANAGER_CAN_GIVE_.indexOf(u.role) >= 0;
}

function usersSheet_() { return tabWithHeaders_(USERS_SHEET, USER_HEADERS); }

function ownerEmail_() {
  try { return String(Session.getEffectiveUser().getEmail() || '').toLowerCase(); } catch (e) { return ''; }
}

function readUsers_() {
  const sheet = usersSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, USER_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2, email: String(r[0]).trim().toLowerCase(), role: String(r[1]).trim(), note: String(r[2]),
    added: r[3] instanceof Date ? r[3].toISOString() : String(r[3]), addedBy: String(r[4]),
    hidden: String(r[5] || '').split(',').map(x => x.trim()).filter(x => SCREENS[x]),
  })).filter(u => u.email);
}

let roleCache_ = null, hiddenCache_ = null;   // one lookup per request
/** 🪜 Screens trickle down: someone loses every screen hidden from them AND every screen hidden from whoever added them,
 *  and from whoever added that person, up the chain, until it reaches an admin or the owner (who see everything).
 *  So an admin hiding Reports from a manager hides it from everyone that manager added, straight away. */
function effectiveHidden_(u, all) {
  const out = new Set(u.hidden), seen = new Set([u.email]), owner = ownerEmail_();
  let by = String(u.addedBy || '').trim().toLowerCase();
  while (by && by !== owner && !seen.has(by)) {
    seen.add(by);
    const up = all.find(x => x.email === by);
    if (!up || up.role === 'Admin') break;
    up.hidden.forEach(h => out.add(h));
    by = String(up.addedBy || '').trim().toLowerCase();
  }
  return SCREEN_IDS_.filter(h => out.has(h));
}
/** 'Admin' | 'Editor' | 'Viewer' for the person using the app, or '' if they're not allowed in. */
function myRole_() {
  if (roleCache_ !== null) return roleCache_;
  const me = String(currentUser_() || '').toLowerCase();
  let role = '';
  if (me && me === ownerEmail_()) role = 'Admin';
  else if (me && me !== 'unknown') {
    const all = readUsers_(), u = all.find(x => x.email === me);
    if (u && ROLE_RANK[u.role]) role = u.role;
    hiddenCache_ = u && role !== 'Admin' ? effectiveHidden_(u, all) : [];
  }
  roleCache_ = role;
  return role;
}

/** The screens hidden from the person using the app ([] for admins). */
function myHiddenScreens_() {
  myRole_();
  return myRole_() === 'Admin' ? [] : hiddenCache_ || [];
}
/** For server functions that belong to one screen: refuses if an admin has hidden that screen from this person. */
function requireScreen_(id) {
  if (myHiddenScreens_().includes(id)) throw new Error('An admin has turned off ' + SCREENS[id] + ' for you.');
}

/** Every server function starts with this. Viewer = can read; Editor = can change; Admin = can manage users. */
function requireRole_(min) {
  const role = myRole_();
  if (!role) throw new Error("You don't have access to Supplybot. Ask an admin (" + ownerEmail_() + ') to add you.');
  if (ROLE_RANK[role] < ROLE_RANK[min]) {
    throw new Error(min === 'Admin' ? 'Only admins can do that.' : "You have view-only access, so you can't make changes.");
  }
  return role;
}

/** Who am I, and what can I do? Called when the app opens. Allowed for anyone (it's how the app knows to say "no access"). */
function getMe() {
  const role = myRole_();
  return { email: currentUser_(), role, roleName: ROLE_NAMES[role] || '', isOwner: currentUser_().toLowerCase() === ownerEmail_(), ownerEmail: ownerEmail_(),
    hidden: role ? myHiddenScreens_() : [], resume: role ? myResume_() : '', appUrl: appUrl_(), sbAi: !!role && sbAiAllowed_() };
}
/** The app's own web address, for the "Reload" button the app shows if its connection was lost while the phone slept. */
function appUrl_() { try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } }

// ---------- Pick up where you left off ----------
// The phone keeps where each person was (tab, search, open page, an unsaved order or item form) on the phone itself, but
// phones don't always keep that for a page inside Google's frame. So the app also sends it here now and then, kept for
// 6 hours, and gets it back with getMe when it starts.
const RESUME_MAX_CHARS = 90000;
const resumeCacheKey_ = () => 'resume:' + currentUser_().toLowerCase();
function myResume_() { try { return CacheService.getScriptCache().get(resumeCacheKey_()) || ''; } catch (e) { return ''; } }
function saveResume(json) {
  requireRole_('Viewer');
  json = String(json || '');
  if (json.length > RESUME_MAX_CHARS) return false;
  CacheService.getScriptCache().put(resumeCacheKey_(), json, 21600);
  return true;
}

// ---------- My profile ----------
// Each person can store a small portrait for their profile. The pictures live in a Drive folder "Profile Photos" and
// the "User Profiles" tab says whose is whose. Anyone allowed into the app (View only too) can change their own.
// Thumbnail: a tiny copy (64 px, ~3 KB) kept right in the tab, shown next to that person's comments everywhere.
// Prefs: that person's display settings, so they follow them to any phone or computer (see setMyPrefs).
const PROFILES_SHEET = 'User Profiles';
const PROFILE_HEADERS = ['Email', 'Photo File ID', 'Updated', 'Thumbnail', 'Prefs', 'Emoji Pics'];
const MAX_THUMB_CHARS = 20000;
const PROFILE_FOLDER = 'Profile Photos';
const MAX_PROFILE_KB = 400;   // the phone shrinks the picture to a small square first, so it's normally ~20 KB
const profilesSheet_ = () => tabWithHeaders_(PROFILES_SHEET, PROFILE_HEADERS);

function readProfiles_() {
  const sheet = profilesSheet_(), last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, PROFILE_HEADERS.length).getValues()
    .map((r, i) => ({ row: i + 2, email: String(r[0]).trim().toLowerCase(), fileId: String(r[1]).trim(), thumb: String(r[3] || ''), prefs: prefsOf_(r[4]) }))
    .filter(p => p.email);
}

/** The signed-in person: email, role and their portrait (a data URL, or '' when none). */
function getMyProfile() {
  const role = requireRole_('Viewer');
  const me = currentUser_().toLowerCase();
  const p = readProfiles_().find(x => x.email === me);
  let photo = '';
  if (p && p.fileId) { try { photo = dataUrl_(DriveApp.getFileById(p.fileId).getBlob()); } catch (e) { /* file deleted from Drive */ } }
  return { email: currentUser_(), role, roleName: ROLE_NAMES[role] || '', photo, prefs: p ? p.prefs : {}, emojiPics: p ? emojiPicsOf_(profileEmojiPicsCell_(p.row)) : [] };
}
// 🖼 Picture emojis (My emoji → Upload a picture): small pictures (about 96 px), up to 8, kept in their own column of the
// person's profile row so they're the same on every device. Index n matches the [emo:pic.n…] token in their emoji list.
const MAX_EMOJI_PIC = 9000;
const emojiPicsOf_ = v => { let a = []; try { a = JSON.parse(String(v || '[]')); } catch (e) { a = []; }
  return (Array.isArray(a) ? a : []).slice(0, 8).map(x => typeof x === 'string' && x.length <= MAX_EMOJI_PIC && /^data:image\/(webp|jpeg|png);base64,[A-Za-z0-9+\/=]+$/.test(x) ? x : ''); };
function profileEmojiPicsCell_(row) { try { return profilesSheet_().getRange(row, 6).getValue(); } catch (e) { return ''; } }
function setMyEmojiPics(list) {
  requireRole_('Viewer');
  const me = currentUser_().toLowerCase(), clean = emojiPicsOf_(JSON.stringify(list || []));
  return withLock_(() => {
    const sheet = profilesSheet_(), p = readProfiles_().find(x => x.email === me);
    if (p) sheet.getRange(p.row, 6).setValue(JSON.stringify(clean));
    else sheet.appendRow([me, '', new Date(), '', '{}', JSON.stringify(clean)]);
    return clean.length;
  });
}

// Display settings a person can save to their profile: [name, allowed values]. itemPics / orderPics: how photos look
// in the Items lists and on order cards (the picture style picker on each list).
const PIC_STYLES = ['small', 'large', 'round', 'banner', 'grid', 'polaroid', 'side', 'none', 'hex', 'stack', 'cover', 'gridcover'];
const PICK_VIEWS = ['list', 'large', 'round', 'grid', 'cover', 'gridcover', 'polaroid', 'none'];   // the Add items picker on an order
const PREF_FIELDS = { itemPics: PIC_STYLES, orderPics: PIC_STYLES, pickPics: PICK_VIEWS, loaderLogo: ['mark', 'bot', 'truck'],
  orderPhotos: ['bento', 'carousel', 'full', 'masonry', 'snapshot', 'showcase', 'classic', 'mini'], itemPhotos: ['standard', 'mini'],
  font: ['theme', 'system', 'inter', 'poppins', 'nunito', 'montserrat', 'grotesk', 'slab', 'comfortaa', 'mono',
    'outfit', 'manrope', 'dmsans', 'sora', 'lexend', 'rubik', 'orbitron', 'exo', 'quicksand', 'merriweather', 'atkinson'] };
const prefsOf_ = v => {
  let o = {};
  try { o = JSON.parse(String(v || '{}')) || {}; } catch (e) { o = {}; }
  const out = {};
  Object.keys(PREF_FIELDS).forEach(k => { if (PREF_FIELDS[k].includes(o[k])) out[k] = o[k]; });
  // 😀 My emoji collection (up to 8), so it's the same on every device: each one as its [emo:…] token (the same text comments use)
  if (Array.isArray(o.emojis)) out.emojis = o.emojis.filter(t => typeof t === 'string' && t.length < 260 && /^\[emo:[a-z0-9.~-]+\]$/.test(t)).slice(0, 8);
  return out;
};

/** Saves the signed-in person's display settings (only the ones given change); returns all of them. */
function setMyPrefs(prefs) {
  requireRole_('Viewer');
  const me = currentUser_().toLowerCase();
  return withLock_(() => {
    const p = readProfiles_().find(x => x.email === me);
    const next = Object.assign({}, p ? p.prefs : {}, prefsOf_(JSON.stringify(prefs || {})));
    const sheet = profilesSheet_();
    if (p) sheet.getRange(p.row, 5).setValue(JSON.stringify(next));
    else sheet.appendRow([me, '', new Date(), '', JSON.stringify(next)]);
    return next;
  });
}

/** Everyone's comment thumbnail: { email: data URL }. */
function getAvatars() {
  requireRole_('Viewer');
  const out = {};
  readProfiles_().forEach(p => { if (/^data:image\/jpeg;base64,/.test(p.thumb)) out[p.email] = p.thumb; });
  return out;
}

/**
 * Saves the signed-in person's portrait (a JPEG / PNG data URL, already shrunk on the phone), replacing any earlier one.
 * thumb: the tiny copy for comments (a JPEG data URL).
 */
function setMyPhoto(dataUrl, thumb) {
  requireRole_('Viewer');
  const m = /^data:(image\/(?:jpeg|png|webp));base64,(.+)$/.exec(dataUrl || '');
  if (!m) throw new Error('That picture couldn’t be read. Try a JPEG or PNG photo.');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > MAX_PROFILE_KB * 1024) throw new Error('That picture is too large.');
  thumb = String(thumb || '');
  if (thumb && (!/^data:image\/jpeg;base64,[A-Za-z0-9+\/=]+$/.test(thumb) || thumb.length > MAX_THUMB_CHARS)) thumb = '';
  const me = currentUser_().toLowerCase();
  return withLock_(() => {
    const it = DriveApp.getFoldersByName(PROFILE_FOLDER);
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder(PROFILE_FOLDER);
    const file = folder.createFile(Utilities.newBlob(bytes, m[1], me + ' profile.' + (m[1] === 'image/png' ? 'png' : 'jpg')));
    const sheet = profilesSheet_(), p = readProfiles_().find(x => x.email === me);
    if (p) {
      if (p.fileId) { try { DriveApp.getFileById(p.fileId).setTrashed(true); } catch (e) { /* already gone */ } }
      sheet.getRange(p.row, 2, 1, 3).setValues([[file.getId(), new Date(), thumb]]);
    } else sheet.appendRow([me, file.getId(), new Date(), thumb]);
    return getMyProfile();
  });
}

/** Removes the signed-in person's portrait (the file goes to Drive's trash). */
function removeMyPhoto() {
  requireRole_('Viewer');
  const me = currentUser_().toLowerCase();
  return withLock_(() => {
    const p = readProfiles_().find(x => x.email === me);
    if (p && p.fileId) {
      try { DriveApp.getFileById(p.fileId).setTrashed(true); } catch (e) { /* already gone */ }
      profilesSheet_().getRange(p.row, 2, 1, 3).setValues([['', new Date(), '']]);
    }
    return getMyProfile();
  });
}

// ---------- 📝 My notes: private notes for the signed-in person (title, text, links and PDFs) ----------
// Each person only ever gets their own notes back from these functions. The rows live in a hidden "My Notes" tab;
// attached PDFs go in a Drive folder "My Notes Files", shared only with the note's owner (not with the domain).
const NOTES_SHEET = 'My Notes';
const NOTES_HEADERS = ['Owner', 'ID', 'Title', 'Text', 'Links', 'Files', 'Updated', 'Created', 'Color', 'Tags', 'Media', 'Value', 'Tally', 'Wallpaper'];
// 🧮 A note's tally: rows of { label, n, u } (u: ft in yd mi m km = distance, ea = quantity, usd = value) and du, the unit the distance total shows in
const TALLY_UNITS = ['ft', 'in', 'yd', 'mi', 'm', 'km', 'ea', 'usd', 'time'];   // time: s / e are start and end ("HH:MM")
function cleanTally_(t) {
  t = t && typeof t === 'object' ? t : {};
  const hm = v => /^([01]\d|2[0-3]):[0-5]\d$/.test(String(v || '')) ? String(v) : '';
  const rows = (Array.isArray(t.rows) ? t.rows : []).slice(0, 150).map(r => { const u = TALLY_UNITS.indexOf(r && r.u) !== -1 ? r.u : 'ea';
    const o = { label: String(r && r.label || '').slice(0, 80), n: r && r.n !== '' && r.n != null && isFinite(Number(r.n)) ? Number(r.n) : null, u };
    if (u === 'time') { o.s = hm(r.s); o.e = hm(r.e); o.n = null;
      const iso = v => { if (!v) return ''; const d = new Date(String(v)); return isNaN(d.getTime()) ? '' : d.toISOString(); };   // ▶ / ■ buttons: the exact date and time
      if (iso(r.sd)) o.sd = iso(r.sd); if (iso(r.ed)) o.ed = iso(r.ed); }
    return o; });
  // vs: the ⚡ smart value's settings (which live sources add up, and the hourly rate for time)
  const vs = t.vs && typeof t.vs === 'object' ? { on: !!t.vs.on, src: (Array.isArray(t.vs.src) ? t.vs.src : []).filter(x => ['tally', 'orders', 'items', 'time'].indexOf(x) !== -1), rate: isFinite(Number(t.vs.rate)) ? Math.max(0, Number(t.vs.rate)) : 0 } : null;
  const out = { rows, du: ['ft', 'yd', 'mi', 'm', 'km'].indexOf(t.du) !== -1 ? t.du : 'ft' };
  if (vs) out.vs = vs;
  if (t.ao) out.ao = true;   // ➕ add this note's value to the totals of the orders it's tagged to (only its owner sees that)
  return out;
}
function tallyOf_(v) { try { return cleanTally_(JSON.parse(String(v || '{}'))); } catch (e) { return { rows: [], du: 'ft' }; } }

// ⏱🔒 Timer lock: below Manager (Editors, View only), a ⏱ time row is filled only by ▶ Start and ■ Stop, stamped when tapped.
// Once started it can't be restarted, removed or retimed; once stopped it's final. A mistake goes to their manager as a
// ✋ correction request (Timer Corrections), and the manager (or an admin) makes the change. Managers and admins aren't locked.
const TIMER_SLACK_MS_ = 15 * 60 * 1000;   // a tap is saved within seconds; this leaves room for a slow connection
const timersLocked_ = () => (ROLE_RANK[myRole_()] || 0) < ROLE_RANK.Manager;
function checkTimerLock_(old, neu) {
  const ASK = ' Use ✋ Ask for a correction on the timer and your manager will fix it.';
  const oldT = old.rows.filter(r => r.u === 'time'), newT = neu.rows.filter(r => r.u === 'time');
  const now = Date.now(), fresh = v => Math.abs(Date.parse(v) - now) <= TIMER_SLACK_MS_;
  oldT.filter(r => r.sd).forEach(o => {
    const n = newT.find(r => r.sd === o.sd);
    if (!n) throw new Error('A started timer can’t be restarted or removed.' + ASK);
    if (n.s !== o.s) throw new Error('A started timer’s start time can’t be changed.' + ASK);
    if (o.ed) { if (n.ed !== o.ed || n.e !== o.e) throw new Error('A stopped timer can’t be changed or started again.' + ASK); }
    else if (n.ed && !fresh(n.ed)) throw new Error('A timer stops with ■ Stop, at the moment you tap it.' + ASK);
    else if (!n.ed && n.e) throw new Error('A timer’s end is set with ■ Stop.' + ASK);
  });
  newT.forEach(n => {
    if (n.sd && !oldT.some(o => o.sd === n.sd) && (!fresh(n.sd) || (n.ed && !fresh(n.ed)))) throw new Error('Timers start with ▶ Start, at the moment you tap it.' + ASK);
    if (!n.sd && (n.s || n.e) && !oldT.some(o => !o.sd && o.s === n.s && o.e === n.e)) throw new Error('Times are set with ▶ Start and ■ Stop, not typed in.');
  });
}

// ✋ Timer correction requests. The request goes to whoever added the person, if that's a manager or an admin; otherwise
// the app owner. They see it under ⏱ Timer corrections and approve (the timer in the note changes) or decline.
const TFIX_SHEET = 'Timer Corrections';
const TFIX_HEADERS = ['ID', 'Requested', 'By', 'Note ID', 'Note Title', 'Timer', 'Timer Start', 'Timer End', 'Wanted Start', 'Wanted End',
  'Wanted Start Clock', 'Wanted End Clock', 'Reason', 'To', 'Status', 'Decided', 'Decided By', 'Decision Note'];
function tfixSheet_() { const sh = tabWithHeaders_(TFIX_SHEET, TFIX_HEADERS); try { if (!sh.isSheetHidden()) sh.hideSheet(); } catch (e) { /* stays visible */ } return sh; }
function readTfix_() {
  const sh = tfixSheet_(), last = sh.getLastRow(); if (last < 2) return [];
  const iso = v => v instanceof Date ? v.toISOString() : String(v || '');
  return sh.getRange(2, 1, last - 1, TFIX_HEADERS.length).getValues().map((r, i) => ({ row: i + 2, id: String(r[0]), requested: iso(r[1]), by: String(r[2]).toLowerCase(),
    noteId: String(r[3]), noteTitle: String(r[4]), label: String(r[5]), sd: String(r[6]), ed: String(r[7]), wantSd: String(r[8]), wantEd: String(r[9]),
    wantS: String(r[10]), wantE: String(r[11]), reason: String(r[12]), to: String(r[13]).toLowerCase(), status: String(r[14]), decided: iso(r[15]), decidedBy: String(r[16]), decisionNote: String(r[17]) })).filter(f => f.id);
}
function approverFor_(email) {
  const all = readUsers_(), u = all.find(x => x.email === email), by = u && String(u.addedBy || '').trim().toLowerCase(), up = by && all.find(x => x.email === by);
  return up && (up.role === 'Manager' || up.role === 'Admin') ? by : ownerEmail_();
}
/** ✋ Ask for a timer to be corrected. want: { sd, ed, s, e, reason } (ISO date-times, and the clock times as "HH:MM" on the asker's phone). */
function requestTimerFix(noteId, timerStart, want) {
  requireRole_('Viewer');
  want = want || {};
  const me = currentUser_().toLowerCase(), iso = v => { const d = new Date(String(v || '')); return v && !isNaN(d.getTime()) ? d.toISOString() : ''; };
  const hm = v => /^([01]\d|2[0-3]):[0-5]\d$/.test(String(v || '')) ? String(v) : '';
  const wSd = iso(want.sd), wEd = iso(want.ed), reason = String(want.reason || '').trim().slice(0, 300);
  if (!wSd) throw new Error('Pick the correct start.');
  if (wEd && Date.parse(wEd) <= Date.parse(wSd)) throw new Error('The end has to be after the start.');
  if (!reason) throw new Error('Say briefly what happened, so your manager knows.');
  return withLock_(() => {
    const n = readMyNotes_(notesSheet_()).find(x => x.id === String(noteId));
    if (!n) throw new Error('That note was deleted.');
    const r = n.tally.rows.find(x => x.u === 'time' && x.sd === iso(timerStart));
    if (!r) throw new Error('That timer changed. Reopen the note and try again.');
    const sh = tfixSheet_(), to = approverFor_(me);
    readTfix_().filter(f => f.by === me && f.noteId === n.id && f.sd === r.sd && f.status === 'Pending').forEach(f => sh.getRange(f.row, 15, 1, 3).setValues([['Replaced', new Date(), me]]));
    sh.appendRow([Utilities.getUuid(), new Date(), me, n.id, asText_(n.title || 'Untitled note'), asText_(r.label || ''), r.sd, r.ed || '', wSd, wEd, hm(want.s), hm(want.e), asText_(reason), to, 'Pending', '', '', '']);
    sh.getRange(sh.getLastRow(), 7, 1, 6).setNumberFormat('@');
    bumpPulse_('@{' + to + '}');
    return getTimerFixes();
  });
}
/** { mine: my requests (newest first, last 60 days), toMe: requests waiting on me (admins: every pending one), canDecide } */
function getTimerFixes() {
  requireRole_('Viewer');
  const me = currentUser_().toLowerCase(), admin = myRole_() === 'Admin', mgr = !timersLocked_(), cut = Date.now() - 60 * 864e5;
  const all = readTfix_().sort((a, b) => b.requested.localeCompare(a.requested));
  const out = f => { const o = Object.assign({}, f); delete o.row; return o; };
  return { mine: all.filter(f => f.by === me && Date.parse(f.requested) > cut).map(out), canDecide: mgr,
    toMe: mgr ? all.filter(f => f.status === 'Pending' && (f.to === me || admin)).map(out) : [] };
}
/** Managers / admins: approve (the timer in their note is changed to what they asked for) or decline a request. */
function decideTimerFix(id, approve, note) {
  requireRole_('Manager');
  const me = currentUser_().toLowerCase();
  return withLock_(() => {
    const f = readTfix_().find(x => x.id === String(id));
    if (!f || f.status !== 'Pending') throw new Error('That request was already handled.');
    if (f.to !== me && myRole_() !== 'Admin') throw new Error('That request is for ' + f.to + '.');
    if (approve) {
      const sh = notesSheet_(), last = sh.getLastRow(), rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, NOTES_HEADERS.length).getValues();
      const i = rows.findIndex(r => String(r[0]).trim().toLowerCase() === f.by && String(r[1]) === f.noteId);
      if (i < 0) throw new Error('Their note was deleted, so there’s nothing to correct. Decline it instead.');
      const t = tallyOf_(rows[i][12]), r = t.rows.find(x => x.u === 'time' && x.sd === f.sd);
      if (!r) throw new Error('That timer was changed or removed since they asked. Decline it instead.');
      r.sd = f.wantSd; r.s = f.wantS || r.s; if (f.wantEd) { r.ed = f.wantEd; r.e = f.wantE || r.e; } else { delete r.ed; r.e = ''; }
      sh.getRange(i + 2, 13).setValue(JSON.stringify(cleanTally_(t)));
    }
    tfixSheet_().getRange(f.row, 15, 1, 4).setValues([[approve ? 'Approved' : 'Declined', new Date(), me, asText_(String(note || '').trim().slice(0, 300))]]);
    logSheet_().appendRow([new Date(), currentUser_(), approve ? 'Timer corrected' : 'Timer correction declined', '', f.by, '', '', 0, '', (f.label || 'Timer') + ' · ' + f.noteTitle]);
    bumpPulse_('@{' + f.by + '}');
    return getTimerFixes();
  });
}
// 📷 Photos and videos on notes: the files sit in "My Notes Files" (shared only with the note's owner); a small preview of
// each (made on the phone) is kept in the hidden "Note Media" tab so galleries open fast without fetching the full files.
const NOTE_MEDIA_SHEET = 'Note Media';
const NOTE_MEDIA_HEADERS = ['Note ID', 'File ID', 'Preview', 'Added'];
const MAX_NOTE_MEDIA = 40, MAX_NOTE_PHOTO_MB = 8, MAX_NOTE_VIDEO_MB = 20;
function noteMediaSheet_() { const sh = tabWithHeaders_(NOTE_MEDIA_SHEET, NOTE_MEDIA_HEADERS); try { if (!sh.isSheetHidden()) sh.hideSheet(); } catch (e) { /* visible */ } return sh; }
// 📣 Someone @mentioned in a note gets a notification and can read that one note (read-only) while they're still mentioned in it.
const NOTE_MENTIONS_SHEET = 'Note Mentions';
const NOTE_MENTION_HEADERS = ['Timestamp', 'From', 'To', 'Note ID', 'Title', 'Text', 'ID'];
const NOTE_TAG_KINDS = ['item', 'order', 'contact', 'project', 'company', 'catalog'];
function noteMentionsSheet_() { const sh = tabWithHeaders_(NOTE_MENTIONS_SHEET, NOTE_MENTION_HEADERS); try { if (!sh.isSheetHidden()) sh.hideSheet(); } catch (e) { /* visible */ } return sh; }
function cleanNoteTags_(tags) {
  const seen = {};
  return (Array.isArray(tags) ? tags : []).slice(0, 40).map(t => ({ t: String(t && t.t || ''), ref: String(t && t.ref || '').trim().slice(0, 200), label: String(t && t.label || '').trim().slice(0, 160) }))
    .filter(t => NOTE_TAG_KINDS.indexOf(t.t) !== -1 && t.ref && !seen[t.t + '|' + t.ref] && (seen[t.t + '|' + t.ref] = true));
}
const noteMentioned_ = text => { const out = {}; String(text || '').replace(/@\{([^}\s]+@[^}\s]+)\}/g, (m, e) => { out[e.toLowerCase()] = true; return m; }); return Object.keys(out); };
/** Records (and so notifies) people newly @mentioned in a note. Returns their emails. */
function notifyNoteMentions_(note, me) {
  const who = noteMentioned_(note.text).filter(e => e !== me);
  if (!who.length) return [];
  const sh = noteMentionsSheet_(), last = sh.getLastRow();
  const had = last < 2 ? [] : sh.getRange(2, 1, last - 1, NOTE_MENTION_HEADERS.length).getValues().filter(r => String(r[3]) === note.id).map(r => String(r[2]).toLowerCase());
  const fresh = who.filter(e => had.indexOf(e) === -1);
  fresh.forEach(e => {
    const at = String(note.text).toLowerCase().indexOf('@{' + e + '}'), from = Math.max(0, at - 160);
    const snip = (from ? '…' : '') + String(note.text).slice(from, from + 420) + (from + 420 < String(note.text).length ? '…' : '');
    sh.appendRow([new Date(), me, e, note.id, note.title || '', snip, Utilities.getUuid()]);
  });
  return fresh;
}
/** Note mentions for the notifications list (yours as the person mentioned). */
function noteMentionRows_(me, since) {
  const sh = noteMentionsSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  const ns = notesSheet_(), nl = ns.getLastRow(), still = {};   // only notes that still mention you (deleted or un-mentioned ones drop out)
  if (nl >= 2) ns.getRange(2, 2, nl - 1, 3).getValues().forEach(r => { if (mentions_(r[2], me)) still[String(r[0])] = true; });
  return sh.getRange(2, 1, last - 1, NOTE_MENTION_HEADERS.length).getValues().filter(r => String(r[2]).toLowerCase() === me.toLowerCase() && r[6] && still[String(r[3])]).map(r => ({
    id: String(r[6]), timestamp: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), user: String(r[1]), itemName: String(r[4] || 'A note'), job: '', text: String(r[5]),
    noteId: String(r[3]), kind: 'note', isMine: false, mentionsMe: true, isNew: r[0] instanceof Date && r[0] > since }));
}
/** A note someone shared with you by @mentioning you (read-only), or your own. */
function getSharedNote(id) {
  requireRole_('Viewer');
  const n = sharedNote_(id);
  return { id: n.id, title: n.title, text: n.text, color: n.color, wp: n.wp, links: n.links, files: n.files.map(f => ({ id: f.id, name: f.name, kb: f.kb })), media: n.media, value: n.value, tally: n.tally, owner: n.owner, updated: n.updated, mine: n.owner === currentUser_().toLowerCase() };
}
function getSharedNoteFile(id, fileId) {
  requireRole_('Viewer');
  const n = sharedNote_(id), f = n.files.find(x => x.id === fileId);
  if (!f) throw new Error('That PDF isn’t on this note any more.');
  return { name: f.name, dataUrl: 'data:application/pdf;base64,' + Utilities.base64Encode(DriveApp.getFileById(fileId).getBlob().getBytes()) };
}
function sharedNote_(id) {
  const me = currentUser_().toLowerCase(), sh = notesSheet_(), last = sh.getLastRow();
  const r = last < 2 ? null : sh.getRange(2, 1, last - 1, NOTES_HEADERS.length).getValues().find(x => String(x[1]) === String(id));
  if (!r) throw new Error('That note was deleted.');
  const n = { owner: String(r[0]).trim().toLowerCase(), id: String(r[1]), title: String(r[2] || ''), text: String(r[3] || ''), color: String(r[8] || ''), links: jsonList_(r[4]), files: jsonList_(r[5]),
    updated: r[6] instanceof Date ? r[6].toISOString() : String(r[6] || ''), media: jsonList_(r[10]), value: r[11] === '' || r[11] == null || isNaN(Number(r[11])) ? null : Number(r[11]), tally: tallyOf_(r[12]), wp: String(r[13] || '') };
  if (n.owner !== me && !mentions_(n.text, me)) throw new Error('That note isn’t shared with you any more.');
  return n;
}
const NOTES_FOLDER = 'My Notes Files';
const MAX_NOTES = 300, MAX_NOTE_TEXT = 20000, MAX_NOTE_LINKS = 25, MAX_NOTE_FILES = 15;
// Aa formatting marks in a note's text ({red|…}, **bold**, ++underline++, ~~strike~~, _italic_, [center] / [right]) taken
// out, for places that show plain text (the note PDF). Lists, quotes and indents stay as typed: they read fine as text.
const FMT_KEYS_ = 'red|orange|yellow|green|teal|blue|purple|pink|gray|serif|mono|wide|narrow|comic|small|large|huge|mark';
function fmtPlain_(t) {
  t = String(t || '');
  const span = new RegExp('\\{(?:' + FMT_KEYS_ + ')(?: (?:' + FMT_KEYS_ + '))*\\|([^{}\\n]*)\\}');
  for (let k = 0, b; k < 40 && b !== t; k++) { b = t; t = t.replace(span, '$1'); }
  return t.replace(/\*\*(?=\S)([^\n]*?\S)\*\*/g, '$1').replace(/\+\+(?=\S)([^\n]*?\S)\+\+/g, '$1').replace(/~~(?=\S)([^\n]*?\S)~~/g, '$1')
    .replace(/(^|[^\w])_(?=\S)([^\n]*?\S)_(?!\w)/g, '$1$2').replace(/^((?: {2})*)\[(?:center|right)\] /gm, '$1');
}
function notesSheet_() {
  const sh = tabWithHeaders_(NOTES_SHEET, NOTES_HEADERS);
  try { if (!sh.isSheetHidden()) sh.hideSheet(); } catch (e) { /* stays visible */ }
  return sh;
}
function jsonList_(v) { try { const a = JSON.parse(String(v || '[]')); return Array.isArray(a) ? a : []; } catch (e) { return []; } }
function readMyNotes_(sheet) {
  const me = currentUser_().toLowerCase(), last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, NOTES_HEADERS.length).getValues().map((r, i) => ({ row: i + 2, owner: String(r[0]).trim().toLowerCase(), id: String(r[1]),
    title: String(r[2] || ''), text: String(r[3] || ''), links: jsonList_(r[4]), files: jsonList_(r[5]),
    updated: r[6] instanceof Date ? r[6].toISOString() : String(r[6] || ''), created: r[7] instanceof Date ? r[7].toISOString() : String(r[7] || ''), color: String(r[8] || ''), tags: jsonList_(r[9]), media: jsonList_(r[10]), value: r[11] === '' || r[11] == null || isNaN(Number(r[11])) ? null : Number(r[11]), tally: tallyOf_(r[12]), wp: String(r[13] || '') }))
    .filter(n => n.owner === me && n.id);
}
const noteOut_ = n => ({ id: n.id, title: n.title, text: n.text, links: n.links, files: n.files, updated: n.updated, created: n.created, color: n.color, tags: n.tags, media: n.media, value: n.value, tally: n.tally, wp: n.wp });
/** Your notes, newest change first. */
function getMyNotes() {
  requireRole_('Viewer');
  return { notes: readMyNotes_(notesSheet_()).map(noteOut_).sort((a, b) => String(b.updated).localeCompare(String(a.updated))) };
}
function cleanNoteLinks_(links) {
  return (Array.isArray(links) ? links : []).slice(0, MAX_NOTE_LINKS).map(l => ({ url: String(l && l.url || '').trim().slice(0, 2000), label: String(l && l.label || '').trim().slice(0, 120) }))
    .filter(l => /^https?:\/\/[^\s]+$/i.test(l.url));
}
/** Creates (no id) or updates one of your notes: { id?, title, text, links: [{ url, label }], color }. Returns { notes, id }. */
function saveMyNote(note) {
  requireRole_('Viewer');
  note = note || {};
  const title = String(note.title || '').trim().slice(0, 200), text = String(note.text || '').slice(0, MAX_NOTE_TEXT), links = cleanNoteLinks_(note.links), tags = cleanNoteTags_(note.tags);
  const color = /^(?:[a-z]{0,12}|[a-z]{2,10}-[0-9a-f]{6})$/.test(String(note.color || '')) ? String(note.color || '') : '';   // a preset, or "<pattern>-<hex>" (🎨 custom background)
  const value = note.value === '' || note.value == null || !isFinite(Number(note.value)) ? '' : Math.round(Number(note.value) * 100) / 100;   // 💲 a dollar amount, or blank
  const me = currentUser_().toLowerCase();
  return withLock_(() => {
    const sheet = notesSheet_(), mine = readMyNotes_(sheet), now = new Date();
    let id = String(note.id || '');
    const n = id && mine.find(x => x.id === id);
    if (id && !n) throw new Error('That note was deleted (maybe on another device).');
    if (n && Array.isArray(note.fileOrder)) {   // PDFs arranged in the editor: that order, anything not listed after
      const at = id => { const i = note.fileOrder.indexOf(id); return i === -1 ? 1e6 : i; };
      n.files = n.files.slice().sort((a, b) => at(a.id) - at(b.id));
    }
    if (n && Array.isArray(note.mediaOrder)) {   // photos / videos arranged in the gallery
      const at = id => { const i = note.mediaOrder.indexOf(id); return i === -1 ? 1e6 : i; };
      n.media = n.media.slice().sort((a, b) => at(a.id) - at(b.id));
      sheet.getRange(n.row, 11).setValue(JSON.stringify(n.media));
    }
    const tally = cleanTally_(note.tally);
    if (timersLocked_()) checkTimerLock_(n ? n.tally : { rows: [] }, tally);
    if (n) { sheet.getRange(n.row, 3, 1, 8).setValues([[title, text, JSON.stringify(links), JSON.stringify(n.files), now, n.created || now, color, JSON.stringify(tags)]]); sheet.getRange(n.row, 12, 1, 2).setValues([[value, JSON.stringify(tally)]]); }
    else {
      if (mine.length >= MAX_NOTES) throw new Error('You have ' + MAX_NOTES + ' notes. Delete some old ones first.');
      id = Utilities.getUuid();
      sheet.appendRow([me, id, title, text, JSON.stringify(links), '[]', now, now, color, JSON.stringify(tags), '[]', value, JSON.stringify(tally)]);
      sheet.getRange(sheet.getLastRow(), 3, 1, 2).setNumberFormat('@');
    }
    const notified = notifyNoteMentions_({ id, title, text }, me);
    if (notified.length) bumpPulse_(notified.map(e => '@{' + e + '}').join(' '));
    return { notes: getMyNotes().notes, id, notified };
  });
}
/** Deletes one of your notes; its PDFs go to Drive's trash. */
function deleteMyNote(id) {
  requireRole_('Viewer');
  return withLock_(() => {
    const sheet = notesSheet_(), n = readMyNotes_(sheet).find(x => x.id === String(id));
    if (n) {
      n.files.concat(n.media).forEach(f => { try { DriveApp.getFileById(f.id).setTrashed(true); } catch (e) { /* gone */ } });
      if (n.wp) { try { DriveApp.getFileById(n.wp).setTrashed(true); } catch (e) { /* gone */ } }
      sheet.deleteRow(n.row);
      if (n.media.length) dropNotePreviews_(n.media.map(m => m.id));
    }
    return getMyNotes();
  });
}
/** Attaches a PDF (data URL) to one of your notes. Shared only with you. */
function addMyNoteFile(id, fileName, dataUrl) {
  requireRole_('Viewer');
  const pdf = parsePdf_(fileName, dataUrl), me = currentUser_().toLowerCase();
  return withLock_(() => {
    const sheet = notesSheet_(), n = readMyNotes_(sheet).find(x => x.id === String(id));
    if (!n) throw new Error('Save the note first, then attach PDFs.');
    if (n.files.length >= MAX_NOTE_FILES) throw new Error('A note can hold up to ' + MAX_NOTE_FILES + ' PDFs.');
    const it = DriveApp.getFoldersByName(NOTES_FOLDER), folder = it.hasNext() ? it.next() : DriveApp.createFolder(NOTES_FOLDER);
    const file = folder.createFile(Utilities.newBlob(pdf.bytes, 'application/pdf', pdf.name));
    file.setDescription('Private note attachment for ' + me + ' (Supplybot · My notes)');
    try { if (String(file.getOwner() && file.getOwner().getEmail()).toLowerCase() !== me) file.addViewer(me); } catch (e) { /* the app can still hand you a copy */ }
    n.files.push({ id: file.getId(), name: pdf.name, kb: Math.max(1, Math.round(pdf.bytes.length / 1024)), url: file.getUrl() });
    sheet.getRange(n.row, 6, 1, 2).setValues([[JSON.stringify(n.files), new Date()]]);
    return getMyNotes();
  });
}
function removeMyNoteFile(id, fileId) {
  requireRole_('Viewer');
  return withLock_(() => {
    const sheet = notesSheet_(), n = readMyNotes_(sheet).find(x => x.id === String(id));
    if (!n) throw new Error('That note was deleted.');
    if (n.files.some(f => f.id === fileId)) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* gone */ } }
    sheet.getRange(n.row, 6, 1, 2).setValues([[JSON.stringify(n.files.filter(f => f.id !== fileId)), new Date()]]);
    return getMyNotes();
  });
}
/** Adds a photo or video (data URL) to one of your notes, with its small preview (a JPEG data URL made on the phone). */
function addMyNoteMedia(id, fileName, dataUrl, preview, meta) {
  requireRole_('Viewer');
  const m = /^data:((image\/(jpeg|png|webp|gif))|(video\/(mp4|quicktime|webm|3gpp|x-m4v)));base64,(.+)$/.exec(dataUrl || '');
  if (!m) throw new Error('Only photos (JPEG, PNG, WebP, GIF) and videos (MP4, MOV, WebM) can be added.');
  const isVideo = !!m[4], bytes = Utilities.base64Decode(m[6]), limit = (isVideo ? MAX_NOTE_VIDEO_MB : MAX_NOTE_PHOTO_MB) * 1024 * 1024;
  if (bytes.length > limit) throw new Error('That ' + (isVideo ? 'video' : 'photo') + ' is over ' + (isVideo ? MAX_NOTE_VIDEO_MB : MAX_NOTE_PHOTO_MB) + ' MB.');
  preview = String(preview || '');
  if (!/^data:image\/jpeg;base64,[A-Za-z0-9+\/=]+$/.test(preview) || preview.length > 40000) preview = '';
  meta = meta || {};
  const me = currentUser_().toLowerCase();
  let name = String(fileName || (isVideo ? 'video' : 'photo')).replace(/[\\/:*?"<>|]+/g, ' ').trim().slice(0, 120) || 'media';
  return withLock_(() => {
    const sheet = notesSheet_(), n = readMyNotes_(sheet).find(x => x.id === String(id));
    if (!n) throw new Error('Save the note first, then add photos and videos.');
    if (n.media.length >= MAX_NOTE_MEDIA) throw new Error('A note can hold up to ' + MAX_NOTE_MEDIA + ' photos and videos.');
    const it = DriveApp.getFoldersByName(NOTES_FOLDER), folder = it.hasNext() ? it.next() : DriveApp.createFolder(NOTES_FOLDER);
    const file = folder.createFile(Utilities.newBlob(bytes, m[1], name));
    file.setDescription('Private note ' + (isVideo ? 'video' : 'photo') + ' for ' + me + ' (Supplybot · My notes)');
    try { if (String(file.getOwner() && file.getOwner().getEmail()).toLowerCase() !== me) file.addViewer(me); } catch (e) { /* the app still shows it */ }
    n.media.push({ id: file.getId(), kind: isVideo ? 'video' : 'photo', name, kb: Math.max(1, Math.round(bytes.length / 1024)),
      w: Math.round(Number(meta.w) || 0), h: Math.round(Number(meta.h) || 0), dur: Math.round(Number(meta.dur) || 0) });
    sheet.getRange(n.row, 11).setValue(JSON.stringify(n.media));
    sheet.getRange(n.row, 7).setValue(new Date());
    noteMediaSheet_().appendRow([n.id, file.getId(), preview, new Date()]);
    return getMyNotes();
  });
}
function removeMyNoteMedia(id, fileId) {
  requireRole_('Viewer');
  return withLock_(() => {
    const sheet = notesSheet_(), n = readMyNotes_(sheet).find(x => x.id === String(id));
    if (!n) throw new Error('That note was deleted.');
    if (n.media.some(f => f.id === fileId)) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* gone */ } }
    sheet.getRange(n.row, 11).setValue(JSON.stringify(n.media.filter(f => f.id !== fileId)));
    dropNotePreviews_([fileId]);
    return getMyNotes();
  });
}
function dropNotePreviews_(ids) {
  const sh = noteMediaSheet_(), last = sh.getLastRow();
  if (last < 2) return;
  const vals = sh.getRange(2, 2, last - 1, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) if (ids.indexOf(String(vals[i][0])) !== -1) sh.deleteRow(i + 2);
}
/** The small previews for a note's photos and videos: { fileId: dataUrl }. Yours, or a note that @mentions you. */
function getNoteMediaPreviews(id) {
  requireRole_('Viewer');
  const n = sharedNote_(id), want = {}; n.media.forEach(m => want[m.id] = true);
  const sh = noteMediaSheet_(), last = sh.getLastRow(), out = {};
  if (last >= 2) sh.getRange(2, 1, last - 1, 3).getValues().forEach(r => { if (String(r[0]) === n.id && want[String(r[1])] && r[2]) out[String(r[1])] = String(r[2]); });
  return out;
}
/** One photo or video from a note, as a data URL. Yours, or a note that @mentions you. */
// 🖼 A note's own wallpaper: a picture uploaded for its background (shrunk on the phone first), kept in the private "My Notes
// Files" folder like its photos. The owner and anyone the note @mentions see it (getNoteWallpaper checks, like shared notes).
const MAX_NOTE_WP_KB = 900;
function setNoteWallpaper(id, dataUrl) {
  requireRole_('Viewer');
  const me = currentUser_().toLowerCase();
  let bytes = null, type = '';
  if (dataUrl) {
    const m = /^data:(image\/(jpeg|png|webp));base64,(.+)$/.exec(String(dataUrl));
    if (!m) throw new Error('Wallpapers can be JPEG, PNG or WebP pictures.');
    bytes = Utilities.base64Decode(m[3]); type = m[1];
    if (bytes.length > MAX_NOTE_WP_KB * 1024) throw new Error('That picture is too big for a wallpaper (over ' + MAX_NOTE_WP_KB + ' KB).');
  }
  return withLock_(() => {
    const sheet = notesSheet_(), n = readMyNotes_(sheet).find(x => x.id === String(id));
    if (!n) throw new Error('Save the note first, then add a wallpaper.');
    if (n.wp) { try { DriveApp.getFileById(n.wp).setTrashed(true); } catch (e) { /* already gone */ } }
    let fid = '';
    if (bytes) {
      const it = DriveApp.getFoldersByName(NOTES_FOLDER), folder = it.hasNext() ? it.next() : DriveApp.createFolder(NOTES_FOLDER);
      const file = folder.createFile(Utilities.newBlob(bytes, type, 'wallpaper-' + n.id));
      file.setDescription('Private note wallpaper for ' + me + ' (Supplybot · My notes)');
      fid = file.getId();
    }
    sheet.getRange(n.row, 14).setValue(fid);
    sheet.getRange(n.row, 7).setValue(new Date());
    return { wp: fid };
  });
}
/** The wallpaper picture of a note that's yours or @mentions you: { wp, dataUrl } ('' when it has none). */
function getNoteWallpaper(id) {
  requireRole_('Viewer');
  const n = sharedNote_(id);
  if (!n.wp) return { wp: '', dataUrl: '' };
  const blob = DriveApp.getFileById(n.wp).getBlob();
  return { wp: n.wp, dataUrl: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
}
function getNoteMediaFile(id, fileId) {
  requireRole_('Viewer');
  const n = sharedNote_(id), f = n.media.find(x => x.id === fileId);
  if (!f) throw new Error('That photo or video isn’t on this note any more.');
  const blob = DriveApp.getFileById(fileId).getBlob();
  return { name: f.name, kind: f.kind, dataUrl: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
}

/**
 * 📄 A note as a PDF (yours, or one that @mentions you): title, value, tags, the text, the tally with its totals, links,
 * photos (and video stills), and the names of attached PDFs. Returns { name, dataUrl } for the phone to save.
 * opts.photos: false leaves the photos out (smaller file); opts.logo: false leaves off the EnviroSpark logo header.
 */
function exportNotePdf(id, opts) {
  requireRole_('Viewer');
  opts = opts || {};
  const n = sharedNote_(id), e = escHtml_, tz = Session.getScriptTimeZone();
  const sh = notesSheet_(), last = sh.getLastRow();
  const row = last < 2 ? null : sh.getRange(2, 1, last - 1, NOTES_HEADERS.length).getValues().find(r => String(r[1]) === n.id);
  const tags = row ? jsonList_(row[9]) : [], tally = row ? tallyOf_(row[12]) : { rows: [], du: 'ft' }, color = row ? String(row[8] || '') : '';
  const usd = v => (v < 0 ? '-$' : '$') + Math.abs(Math.round(Number(v) * 100) / 100).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const num = (v, d) => (Math.round(Number(v) * Math.pow(10, d == null ? 2 : d)) / Math.pow(10, d == null ? 2 : d)).toLocaleString('en-US');
  const nice = t => String(t || '').replace(/@\{([^}\s]+)\}/g, (m, em) => '@' + em.split('@')[0].replace(/[._]+/g, ' ').replace(/\b\w/g, c => c.toUpperCase()))
    .replace(/\[emo:[^\]]+\]/g, '🙂');
  const title = n.title || nice(n.text).trim().split('\n')[0].slice(0, 60) || 'Untitled note';
  if (opts.value != null && isFinite(Number(opts.value))) n.value = Number(opts.value);   // ⚡ a smart value, worked out live on the phone
  const ICON = { item: '📦', order: '🚚', contact: '👤', project: '🏗', company: '🏢', catalog: '📒' };
  // tally
  const U = { ft: ['dist', 'ft', 1], in: ['dist', 'in', 1 / 12], yd: ['dist', 'yd', 3], mi: ['dist', 'mi', 5280], m: ['dist', 'm', 3.280839895], km: ['dist', 'km', 3280.839895], ea: ['qty', 'ea', 1], usd: ['val', '$', 1], time: ['time', 'hrs', 1] };
  const tm = v => { const m = /^(\d{1,2}):(\d{2})$/.exec(v || ''); return m ? Number(m[1]) * 60 + Number(m[2]) : null; };
  const mins = r => { if (r.sd && r.ed) { const d = (new Date(r.ed) - new Date(r.sd)) / 60000; if (d >= 0) return d; }
    const a = tm(r.s), b = tm(r.e); return a == null || b == null ? null : (b - a + 1440) % 1440; };
  const stamp = v => { try { return Utilities.formatDate(new Date(v), Session.getScriptTimeZone(), 'MMM d, h:mm a'); } catch (x) { return ''; } };
  const hm = m => { m = Math.round(m); const h = Math.floor(m / 60), r = m % 60; return h ? h + 'h' + (r ? ' ' + r + 'm' : '') : r + 'm'; };
  const clock = v => { const m = tm(v); if (m == null) return ''; const h = Math.floor(m / 60); return ((h % 12) || 12) + ':' + String(m % 60).padStart(2, '0') + (h < 12 ? ' AM' : ' PM'); };
  const tot = { dist: 0, distN: 0, qty: 0, qtyN: 0, val: 0, valN: 0, time: 0, timeN: 0 };
  const tallyRows = tally.rows.map(r => {
    const u = U[r.u] || U.ea;
    if (u[0] === 'time') { const m = mins(r); if (m != null) { tot.time += m; tot.timeN++; } return '<tr><td>⏱ ' + e(r.label || '—') + '</td><td class="mut">' + e([r.sd ? stamp(r.sd) : clock(r.s), r.ed ? stamp(r.ed) : clock(r.e)].filter(Boolean).join(' – ')) + '</td><td class="r">' + (m == null ? '—' : hm(m)) + '</td></tr>'; }
    const v = r.n == null ? null : Number(r.n);
    if (v != null) { if (u[0] === 'dist') { tot.dist += v * u[2]; tot.distN++; } else if (u[0] === 'qty') { tot.qty += v; tot.qtyN++; } else { tot.val += v; tot.valN++; } }
    return '<tr><td>' + ({ dist: '📏 ', qty: '🔢 ', val: '💲 ' }[u[0]]) + e(r.label || '—') + '</td><td class="mut">' + ({ dist: 'Distance', qty: 'Quantity', val: 'Value' }[u[0]]) + '</td><td class="r">' + (v == null ? '—' : u[0] === 'val' ? usd(v) : num(v) + ' ' + u[1]) + '</td></tr>';
  }).join('');
  const totBoxes = [tot.distN && ['📏 Distance', num(tot.dist / U[tally.du][2]) + ' ' + U[tally.du][1], tally.du === 'ft' ? num(tot.dist / 5280, 3) + ' mi' : num(tot.dist) + ' ft'],
    tot.timeN && ['⏱ Time', hm(tot.time), num(tot.time / 60) + ' hours'], tot.qtyN && ['🔢 Quantity', num(tot.qty), ''], tot.valN && ['💲 Value', usd(tot.val), '']].filter(Boolean)
    .map(([l, v, sub]) => '<td class="tot"><div class="tl">' + l + '</div><div class="tv">' + e(v) + '</div>' + (sub ? '<div class="ts">' + e(sub) + '</div>' : '') + '</td>').join('');
  // photos / videos
  let gallery = '';
  if (opts.photos !== false && n.media.length) {
    const prev = {};
    const ms = noteMediaSheet_(), ml = ms.getLastRow();
    if (ml >= 2) ms.getRange(2, 1, ml - 1, 3).getValues().forEach(r => { if (String(r[0]) === n.id && r[2]) prev[String(r[1])] = String(r[2]); });
    const cells = n.media.slice(0, 24).map(m => {
      let src = '';
      if (m.kind === 'photo') { try { const b = DriveApp.getFileById(m.id).getBlob(); src = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()); } catch (er) { src = prev[m.id] || ''; } }
      else src = prev[m.id] || '';
      return '<td class="ph">' + (src ? '<img src="' + src + '">' : '<div class="noph">' + (m.kind === 'video' ? '🎬' : '🖼') + '</div>') +
        '<div class="cap">' + (m.kind === 'video' ? '▶ Video' + (m.dur ? ' · ' + Math.floor(m.dur / 60) + ':' + String(m.dur % 60).padStart(2, '0') : '') : e(m.name || '')) + '</div></td>';
    });
    gallery = '<h2>📷 Photos &amp; videos <span class="n">' + n.media.length + '</span></h2><table class="gal">' +
      Array.from({ length: Math.ceil(cells.length / 3) }, (_, i) => '<tr>' + cells.slice(i * 3, i * 3 + 3).join('') + '<td class="ph empty"></td>'.repeat(Math.max(0, 3 - cells.slice(i * 3, i * 3 + 3).length)) + '</tr>').join('') + '</table>' +
      (n.media.length > 24 ? '<p class="mut">' + (n.media.length - 24) + ' more in the app.</p>' : '');
  }
  const accent = { yellow: '#f59e0b', green: '#22c55e', blue: '#3b82f6', pink: '#ec4899', purple: '#8b5cf6' }[color] || '#94ba65';
  const html = '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,Helvetica,sans-serif;color:#1f2937;font-size:12px;margin:0}table,td{font-size:12px}' +
    '.bar{height:8px;background:' + accent + '}.wrap{padding:22px 28px}' +
    'h1{font-size:24px;margin:10px 0 4px}h2{font-size:14px;margin:20px 0 8px;padding-bottom:4px;border-bottom:2px solid #e5e7eb}h2 .n{color:#6b7280;font-weight:normal}' +
    '.meta{color:#6b7280;font-size:11px}.val{display:inline-block;margin-top:8px;padding:5px 14px;border-radius:6px;background:#fcd34d;color:#3b2a00;font-weight:bold;font-size:16px}' +
    '.tags span{display:inline-block;margin:6px 6px 0 0;padding:3px 9px;border-radius:10px;background:#f3f4f6;font-size:11px}' +
    '.text{white-space:pre-wrap;line-height:1.55;font-size:13px}' +
    'table.tal{width:100%;border-collapse:collapse}table.tal td{padding:6px 8px;border-bottom:1px solid #eee}.r{text-align:right;font-weight:bold}.mut{color:#6b7280}' +
    'table.tots{width:100%;border-collapse:separate;border-spacing:8px 0;margin:10px -8px 0}td.tot{background:#f9fafb;border:1px solid #e5e7eb;border-radius:8px;padding:8px 10px;vertical-align:top}' +
    '.tl{font-size:11px;color:#6b7280;font-weight:bold}.tv{font-size:18px;font-weight:bold;margin-top:2px}.ts{font-size:10px;color:#6b7280}' +
    'a{color:#2563eb}.links div,.pdfs div{padding:4px 0}' +
    'table.gal{width:100%;border-collapse:separate;border-spacing:6px}td.ph{width:33%;vertical-align:top;text-align:center}td.ph img{width:100%;max-height:200px;object-fit:cover;border-radius:6px}' +
    '.noph{height:120px;background:#f3f4f6;border-radius:6px;font-size:30px;line-height:120px}.cap{font-size:10px;color:#6b7280;margin-top:3px;word-break:break-all}td.empty{border:0}' +
    '.foot{margin-top:24px;padding-top:8px;border-top:1px solid #e5e7eb;color:#9ca3af;font-size:10px}' +
    'table.lh{width:100%;border-collapse:collapse;border-bottom:3px solid #8fb85f;margin-bottom:12px}table.lh td{vertical-align:bottom;padding-bottom:10px}.logo{width:150px}.lt{text-align:right;font-size:18px;font-weight:bold;color:#374151}' +
    '</style></head><body><div class="bar"></div><div class="wrap">' +
    (opts.logo !== false ? '<table class="lh"><tr><td><img class="logo" src="' + PDF_LOGO + '"></td><td class="lt">Note</td></tr></table>' : '') +
    '<div class="meta">📝 Note · updated ' + e(Utilities.formatDate(new Date(n.updated || new Date()), tz, 'MMM d, yyyy h:mm a')) + '</div>' +
    '<h1>' + e(title) + '</h1>' +
    (n.value != null ? '<div class="val">' + usd(n.value) + '</div>' : '') +
    (tags.length ? '<div class="tags">' + tags.map(t => '<span>' + (ICON[t.t] || '🏷') + ' ' + e(t.label) + '</span>').join('') + '</div>' : '') +
    (String(n.text || '').trim() ? '<h2>Notes</h2><div class="text">' + e(nice(fmtPlain_(n.text))) + '</div>' : '') +
    (tally.rows.length ? '<h2>🧮 Tally</h2><table class="tal">' + tallyRows + '</table>' + (totBoxes ? '<table class="tots"><tr>' + totBoxes + '</tr></table>' : '') : '') +
    (n.links.length ? '<h2>🔗 Links</h2><div class="links">' + n.links.map(l => '<div><a href="' + e(l.url) + '">' + e(l.label || l.url) + '</a>' + (l.label ? ' <span class="mut">' + e(l.url) + '</span>' : '') + '</div>').join('') + '</div>' : '') +
    gallery +
    (n.files.length ? '<h2>📄 Attached PDFs</h2><div class="pdfs">' + n.files.map(f => '<div>📄 ' + e(f.name) + ' <span class="mut">(' + (f.kb >= 1024 ? (f.kb / 1024).toFixed(1) + ' MB' : f.kb + ' KB') + ')</span></div>').join('') + '</div>' : '') +
    '<div class="foot">Exported from Supplybot · My notes · ' + e(Utilities.formatDate(new Date(), tz, "MMM d, yyyy 'at' h:mm a")) + ' by ' + e(currentUser_()) + '</div></div></body></html>';
  const bytes = Utilities.newBlob(html, 'text/html', 'note.html').getAs('application/pdf').getBytes();
  const name = (title.replace(/[\\/:*?"<>|]+/g, ' ').trim().slice(0, 80) || 'Note') + '.pdf';
  return { name, dataUrl: 'data:application/pdf;base64,' + Utilities.base64Encode(bytes) };
}

/**
 * ▦ A sheet of QR labels as a PDF: 3 across, each with the QR (a PNG made on the phone), the item's name, part #,
 * location and its code. labels: [{ name, part, loc, code, png }] (up to 150). opts.logo: the EnviroSpark logo on top.
 */
function makeQrLabelsPdf(labels, opts) {
  requireRole_('Viewer');
  opts = opts || {};
  labels = (Array.isArray(labels) ? labels : []).slice(0, 150).filter(l => l && /^data:image\/png;base64,[A-Za-z0-9+\/=]+$/.test(String(l.png || '')) && String(l.png).length < 60000);
  if (!labels.length) throw new Error('No QR labels to print.');
  const e = escHtml_, tz = Session.getScriptTimeZone();
  const cell = l => '<td class="lb"><img class="qr" src="' + l.png + '"><div class="nm">' + e(String(l.name || '').slice(0, 80)) + '</div>' +
    (l.part ? '<div class="sub">Part # ' + e(String(l.part).slice(0, 40)) + '</div>' : '') + (l.loc ? '<div class="sub">📍 ' + e(String(l.loc).slice(0, 50)) + '</div>' : '') +
    '<div class="code">' + e(String(l.code || '').slice(0, 20)) + '</div></td>';
  const rows = Array.from({ length: Math.ceil(labels.length / 3) }, (_, r) => '<tr>' + labels.slice(r * 3, r * 3 + 3).map(cell).join('') + '<td class="lb empty"></td>'.repeat(Math.max(0, 3 - labels.slice(r * 3, r * 3 + 3).length)) + '</tr>').join('');
  const html = '<html><head><meta charset="utf-8"><style>body{font-family:Arial,Helvetica,sans-serif;color:#111;margin:0}.wrap{padding:16px 18px}' +
    'table.lh{width:100%;border-collapse:collapse;border-bottom:3px solid #8fb85f;margin-bottom:10px}table.lh td{vertical-align:bottom;padding-bottom:8px}.logo{width:130px}.lt{text-align:right;font-size:15px;font-weight:bold;color:#374151}' +
    'table.sheet{width:100%;border-collapse:separate;border-spacing:8px}td.lb{width:33%;border:1px dashed #9ca3af;border-radius:8px;padding:10px 8px;text-align:center;vertical-align:top;page-break-inside:avoid}td.empty{border:0}' +
    '.qr{width:130px;height:130px}.nm{font-size:12px;font-weight:bold;margin-top:4px;line-height:1.25}.sub{font-size:10px;color:#4b5563;margin-top:2px}.code{font-family:Courier New,monospace;font-size:13px;font-weight:bold;letter-spacing:1px;margin-top:5px}' +
    '.foot{color:#9ca3af;font-size:9px;margin-top:6px}</style></head><body><div class="wrap">' +
    (opts.logo !== false ? '<table class="lh"><tr><td><img class="logo" src="' + PDF_LOGO + '"></td><td class="lt">QR labels · ' + labels.length + '</td></tr></table>' : '') +
    '<table class="sheet">' + rows + '</table><div class="foot">Scan with Supplybot ▦ Scan to open the item · made ' + e(Utilities.formatDate(new Date(), tz, 'MMM d, yyyy')) + '</div></div></body></html>';
  const bytes = Utilities.newBlob(html, 'text/html', 'labels.html').getAs('application/pdf').getBytes();
  const name = (labels.length === 1 ? 'QR label ' + String(labels[0].name || '').replace(/[\\/:*?"<>|]+/g, ' ').trim().slice(0, 60) : 'QR labels (' + labels.length + ')') + '.pdf';
  return { name, dataUrl: 'data:application/pdf;base64,' + Utilities.base64Encode(bytes) };
}

/** A PDF from one of your notes as a data URL (for when Drive won't open it for you). */
function getMyNoteFile(id, fileId) {
  requireRole_('Viewer');
  const n = readMyNotes_(notesSheet_()).find(x => x.id === String(id)), f = n && n.files.find(x => x.id === fileId);
  if (!f) throw new Error('That PDF isn’t on this note any more.');
  return { name: f.name, dataUrl: 'data:application/pdf;base64,' + Utilities.base64Encode(DriveApp.getFileById(fileId).getBlob().getBytes()) };
}

/** Admins: the user list (the owner first, marked as always Admin). */
function getUsers() {
  requireRole_('Manager');
  const owner = ownerEmail_(), admin = myRole_() === 'Admin', all = readUsers_();
  return { owner: admin ? owner : '', manager: !admin, give: admin ? Object.keys(ROLE_RANK) : MANAGER_CAN_GIVE_, myHidden: admin ? [] : myHiddenScreens_(),
    screens: SCREENS, users: all.filter(u => u.email !== owner && (admin || isMyUser_(u))).map(u => ({ email: u.email, role: u.role, note: u.note, added: u.added, addedBy: u.addedBy, hidden: u.hidden,
      inherited: u.role === 'Admin' ? [] : effectiveHidden_(u, all).filter(h => u.hidden.indexOf(h) < 0), locked: !canManage_(u) })) };
}

/** Everyone who can use the app (for @mentions in comments): email and the users list's Name / Note. */
function getPeople() {
  requireRole_('Viewer');
  const owner = ownerEmail_().toLowerCase();
  const people = readUsers_().map(u => ({ email: u.email, note: u.note }));
  if (owner && !people.some(p => p.email === owner)) people.push({ email: owner, note: '' });
  return people.sort((a, b) => a.email.localeCompare(b.email));
}

/** Admins: add someone, or change their role / note, and (when given) which screens are hidden from them. */
function saveUser(email, role, note, hidden) {
  requireRole_('Manager');
  email = String(email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Enter a valid email address.');
  if (!ROLE_RANK[role]) throw new Error('Choose a role: Admin, Manager, Editor or View only.');
  if (myRole_() !== 'Admin' && MANAGER_CAN_GIVE_.indexOf(role) < 0) throw new Error('Managers can add people as View only or Editor.');
  if (email === ownerEmail_()) throw new Error("That's the app's owner, who is always an Admin.");
  if (email === currentUser_().toLowerCase() && role !== 'Admin') throw new Error("You can't lower your own role. Ask another admin.");
  return withLock_(() => {
    const sheet = usersSheet_();
    const u = readUsers_().find(x => x.email === email);
    if (u) requireCanManage_(u);
    const hid = hidden === undefined || hidden === null ? (u ? u.hidden : []) : cleanHidden_(hidden);
    // (a manager can't give a screen they don't have: screens hidden from them are hidden from their people too, live — effectiveHidden_)
    if (u) sheet.getRange(u.row, 2, 1, 2).setValues([[role, asText_(String(note || '').trim())]]);
    else sheet.appendRow([email, role, asText_(String(note || '').trim()), new Date(), currentUser_(), '']);
    const row = u ? u.row : sheet.getLastRow();
    sheet.getRange(row, USER_HEADERS.indexOf('Hidden Screens') + 1).setValue(hid.join(', '));
    const screensChanged = u && hid.join(',') !== u.hidden.join(',');
    logSheet_().appendRow([new Date(), currentUser_(), !u ? 'User added' : screensChanged && u.role === role ? 'User screens changed' : 'User role changed', '', email, '', '', 0, '',
      ROLE_NAMES[role] + (hid.length && role !== 'Admin' ? ' · hidden: ' + hid.map(h => SCREENS[h]).join(', ') : '')]);
    return getUsers();
  });
}

/** Admins: remove someone's access. */
function removeUser(email) {
  requireRole_('Manager');
  email = String(email || '').trim().toLowerCase();
  if (email === currentUser_().toLowerCase()) throw new Error("You can't remove yourself. Ask another admin.");
  return withLock_(() => {
    const u = readUsers_().find(x => x.email === email);
    if (!u) throw new Error('That person was already removed.');
    requireCanManage_(u);
    usersSheet_().deleteRow(u.row);
    logSheet_().appendRow([new Date(), currentUser_(), 'User removed', '', email, '', '', 0, '', '']);
    return getUsers();
  });
}

/** Serves the web app. People not on the users list get a short "no access" page instead. */
function doGet() {
  if (!myRole_()) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:-apple-system,Segoe UI,Roboto,sans-serif;max-width:420px;margin:15vh auto;padding:24px;text-align:center">' +
      '<h2>Supplybot</h2><p>You don’t have access yet.</p>' +
      '<p style="color:#6b7280">Signed in as <b>' + escHtml_(currentUser_()) + '</b>. Ask an admin (' + escHtml_(ownerEmail_()) + ') to add you, then reopen this page.</p></div>')
      .setTitle('Supplybot')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  const page = HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Supplybot')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  // 📲 Opens full screen like an app when started from the phone's home screen (Add to Home Screen)
  ['apple-mobile-web-app-capable', 'mobile-web-app-capable'].forEach(n => { try { page.addMetaTag(n, 'yes'); } catch (e) { /* not allowed here: fine */ } });
  // 📲 The installed Supplybot app (https://sunshineufo.github.io/supplybot-app/) shows this page inside itself, so phones
  // don't add a script.google.com bar on top. Apps Script can only allow framing for every site, so the page guards itself:
  // Script.html (frameGuard) covers the app with "open Supplybot directly" when anything but that address frames it.
  // Approved by the app's owner, Nov 2026.
  try { page.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); } catch (e) { /* older runtime: browser only */ }
  let logo = 'mark';
  try { const me = currentUser_().toLowerCase(), p = readProfiles_().find(x => x.email === me); if (p && p.prefs.loaderLogo) logo = p.prefs.loaderLogo; } catch (e) { /* default icon */ }
  setFavicon_(page, logo);
  return page;
}

// The page is split into several HTML files because Apps Script cuts off any single file at about 1 MB (the end of the app
// code once went missing, so it never started): Index (the page itself), Styles (+ Styles2) for the CSS, and Script,
// Script2 (+ Script3) for the app code. Index is a template: inside ONE <script> block it pastes Script, Script2 and
// Script3 one after another, so the browser still sees a single script, just stored in parts. Each part stays well under
// 1 MB (the build cuts the code into parts of about 500 KB), leaving room to grow. Script2, Script3 and Styles2 are
// optional: a part that isn't in the project yet is simply skipped (the app's own check then says if code is missing).
const PAGE_PARTS_ = ['Styles', 'Styles2', 'Script', 'Script2', 'Script3'];
const OPTIONAL_PARTS_ = ['Styles2', 'Script2', 'Script3'];
function includeFile(name) {
  if (PAGE_PARTS_.indexOf(name) < 0) throw new Error('Not a page part: ' + name);
  // The Script parts are plain code (no <script> tags of their own), so they're read as raw text: reading them as a page
  // would make Apps Script check them as HTML and refuse ("Malformed HTML content"). Styles is a normal page part.
  try { return /^Script/.test(name) ? HtmlService.createTemplateFromFile(name).getRawContent() : HtmlService.createHtmlOutputFromFile(name).getContent(); }
  catch (e) { if (OPTIONAL_PARTS_.indexOf(name) >= 0 && /not found|No HTML file/i.test(String(e && e.message))) return ''; throw e; }
}

// ---------- Favicon (browser-tab icon) ----------
// The page runs inside Google's frame, so a <link rel="icon"> in Index.html is ignored, and setFaviconUrl won't take
// an embedded (data:) image: it needs the web address of an image file. So the first time each icon is needed, the app
// saves it as a PNG in the "Inventory Photos" Drive folder, shares it view-only by link, and remembers the file. The
// tab shows the same logo as the person's loading logo (My profile → Loading logo): the EnviroSpark mark or Supplybot.
// The image is linked through lh3.googleusercontent.com/d/<id>, which serves the picture itself (the older
// drive.google.com/uc?export=view link now answers with a redirect page that browsers won't use as an icon, which left
// the default globe in the tab). If none of this works, the app still opens, just with the browser's default tab icon.

// Optional: a public https address of an icon (ending in .png or .ico) to use instead of the saved Drive file.
const FAVICON_URL = '';

// The EnviroSpark bolt mark in logo green, 128 x 128 PNG (rendered from the app's vector logo).
const FAVICON_PNG_BASE64 = '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';
// The Supplybot robot on a dark rounded tile, 128 x 128 PNG (for people whose loading logo is Supplybot).
const FAVICON_BOT_PNG_BASE64 = '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';
// Change the _v2 to make the app save a fresh copy. (v1 pointed at an empty copy saved on Oct 28, which Google answers
// with "not found", so the tab showed the default globe.)
const FAVICON_FILE_PROP = 'faviconFileId_v2';
const FAVICON_BOT_FILE_PROP = 'faviconBotFileId_v2';

function faviconUrl_(logo) {
  if (FAVICON_URL) return FAVICON_URL;
  const bot = logo === 'bot' || logo === 'truck', key = bot ? FAVICON_BOT_FILE_PROP : FAVICON_FILE_PROP;
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(key);
  if (!id) {
    const blob = Utilities.newBlob(Utilities.base64Decode(bot ? FAVICON_BOT_PNG_BASE64 : FAVICON_PNG_BASE64), 'image/png',
      bot ? 'Supplybot tab icon (robot).png' : 'Supplybot tab icon.png');
    let file = photoFolder_().createFile(blob);
    if (!(file.getSize() > 0)) {   // an empty copy would leave the globe: try once more, and give up if that's empty too
      file.setTrashed(true);
      file = photoFolder_().createFile(blob.copyBlob());
      if (!(file.getSize() > 0)) { file.setTrashed(true); throw new Error('Couldn’t save the tab icon.'); }
    }
    file.setDescription('Browser-tab icon for the Supplybot app. Keep it shared by link, or the tab icon disappears.');
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e2) { /* Workspace doesn't allow link sharing */ }
    }
    id = file.getId();
    props.setProperty(key, id);
  }
  return 'https://lh3.googleusercontent.com/d/' + id + '?.png';   // the trailing .png tells Google it's an image
}

// Each person's own tab icon: drawn by the app in their floating menu's colour with their loading logo (setMyFavicon),
// saved as a PNG next to the shared ones, and used instead of them. Script property "favicon:<email>" = its file id.
const FAVICON_MINE_PREFIX = 'favicon:';
function myFaviconUrl_() {
  const id = PropertiesService.getScriptProperties().getProperty(FAVICON_MINE_PREFIX + currentUser_().toLowerCase());
  return id ? 'https://lh3.googleusercontent.com/d/' + id + '?.png' : '';
}
/** Saves the tab icon the app drew for this person (a small PNG data URL) and replaces their previous one. */
function setMyFavicon(dataUrl) {
  requireRole_('Viewer');
  const m = /^data:image\/png;base64,([A-Za-z0-9+/=]+)$/.exec(String(dataUrl || ''));
  if (!m || m[1].length > 200000) throw new Error('That tab icon couldn’t be read.');
  const me = currentUser_().toLowerCase(), props = PropertiesService.getScriptProperties(), key = FAVICON_MINE_PREFIX + me;
  const file = photoFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(m[1]), 'image/png', 'Supplybot tab icon · ' + me + '.png'));
  if (!(file.getSize() > 0)) { file.setTrashed(true); throw new Error('Couldn’t save the tab icon.'); }
  file.setDescription('Browser-tab icon for ' + me + ' in the Supplybot app (their floating menu colours). Keep it shared by link.');
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch (e) { try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e2) { /* link sharing not allowed */ } }
  const old = props.getProperty(key);
  props.setProperty(key, file.getId());
  if (old) try { DriveApp.getFileById(old).setTrashed(true); } catch (e) { /* already gone */ }
  return true;
}

function setFavicon_(page, logo) {
  try {
    page.setFaviconUrl(myFaviconUrl_() || faviconUrl_(logo));
  } catch (e) {
    // not accepted, or Drive unavailable: the app opens with the default tab icon
  }
}

/** One-time setup: adds the Photo ID and Barcode / QR Code columns if missing, and the App Log tab. Safe to run repeatedly. */
function setup() {
  requireRole_('Admin');
  const sheet = master_();
  const headers = headerRow_(sheet);
  ADDED_FIELDS.forEach(field => {
    if (findCol_(headers, FIELDS[field]) !== -1) return;
    const title = [].concat(FIELDS[field])[0];
    let col = headers.findIndex(h => SPARE_COLUMN.test(norm_(h)));
    if (col === -1) col = headers.length;
    sheet.getRange(1, col + 1).setValue(title);
    headers[col] = title;
  });
  logSheet_();
  commentsSheet_();
  docsSheet_();
  adjSheet_();
  tabWithHeaders_(ORDERS_SHEET, ORDER_HEADERS);
  tabWithHeaders_(LINES_SHEET, LINE_HEADERS);
  tabWithHeaders_(ORDER_DOCS_SHEET, ORDER_DOC_HEADERS);
  tabWithHeaders_(ORDER_PHOTOS_SHEET, ORDER_PHOTO_HEADERS);
  tabWithHeaders_(ORDER_LINKS_SHEET, ORDER_LINK_HEADERS);
  tabWithHeaders_(ORDER_COSTS_SHEET, ORDER_COST_HEADERS);
  const missing = Object.keys(FIELDS).filter(f => findCol_(headerRow_(sheet), FIELDS[f]) === -1);
  Logger.log(missing.length ? 'Missing columns: ' + missing.map(f => [].concat(FIELDS[f])[0]).join(', ') : 'Setup complete.');
}

/**
 * Optional one-time cleanup (run from the Apps Script editor): Shelf cells that Sheets turned into dates
 * (e.g. "8-1" stored as August 1) are rewritten as the plain text they show. The app displays them correctly
 * either way; this just makes the spreadsheet itself hold "8-1" as text, so a format change can't garble it.
 */
function convertShelvesToText() {
  requireRole_('Admin');
  const sheet = master_();
  const cols = columns_(sheet);
  if (cols.shelf === -1) return Logger.log('No Shelf column.');
  const last = lastItemRow_(sheet, cols);
  const range = sheet.getRange(2, cols.shelf + 1, last - 1, 1);
  const values = range.getValues(), shown = range.getDisplayValues();
  let fixed = 0;
  values.forEach((r, i) => {
    if (r[0] instanceof Date) {
      sheet.getRange(i + 2, cols.shelf + 1).setNumberFormat('@').setValue(shown[i][0]);
      fixed++;
    }
  });
  Logger.log(fixed + ' shelf value(s) converted from dates to text.');
}

const escHtml_ = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

// ---------- Sheet helpers ----------

function master_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheets().find(s => s.getSheetId() === MASTER_SHEET_ID);
  if (!sheet) throw new Error('Master inventory tab not found.');
  return sheet;
}

function docsSheet_() { return tabWithHeaders_(DOCS_SHEET, DOC_HEADERS); }

/** All document rows as objects, with their sheet row number. */
function readDocs_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, DOC_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2,
    timestamp: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
    user: String(r[1]),
    itemName: String(r[2]),
    name: String(r[5]),
    sizeKB: Number(r[6]) || 0,
    fileId: String(r[7]),
    key: String(r[8]),
    url: String(r[9]),
    position: Number(r[10]) || 0,
  })).filter(d => d.fileId);
}

/** An item's PDFs in display order: arranged ones by Position, then the rest newest first. */
function docsInOrder_(docs) {
  return docs.slice().sort((a, b) => (a.position > 0) !== (b.position > 0) ? (a.position > 0 ? -1 : 1)
    : a.position > 0 ? a.position - b.position : b.row - a.row);
}

/** Saves the order an item's PDFs are shown in (item page and search results). fileIds = every PDF, in the new order. */
function reorderDocuments(key, fileIds) {
  requireRole_('Editor');
  fileIds = (fileIds || []).map(String);
  return withLock_(() => {
    const sheet = docsSheet_();
    const mine = readDocs_(sheet).filter(d => d.key === key);
    if (mine.length !== fileIds.length || !mine.every(d => fileIds.includes(d.fileId))) throw new Error('The PDFs changed. Refresh and try again.');
    const col = DOC_HEADERS.indexOf('Position') + 1;
    mine.forEach(d => sheet.getRange(d.row, col).setValue(fileIds.indexOf(d.fileId) + 1));
    return getDocuments(key);
  });
}

function commentsSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(COMMENTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(COMMENTS_SHEET, ss.getNumSheets());
    sheet.appendRow(COMMENT_HEADERS);
    sheet.getRange(1, 1, 1, COMMENT_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** All comment rows as objects, with their sheet row number. */
function readComments_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, COMMENT_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2,
    timestamp: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
    user: String(r[1]),
    text: String(r[5]),
    key: String(r[6]),
    id: String(r[7]),
  })).filter(c => c.id);
}

function logSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET, ss.getNumSheets());
    sheet.appendRow(LOG_HEADERS);
    sheet.getRange(1, 1, 1, LOG_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

const norm_ = s => String(s).trim().replace(/\s+/g, ' ').toLowerCase();

function headerRow_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/** Index of the first header matching any of the given names, or -1. */
function findCol_(headers, names) {
  for (const name of [].concat(names)) {
    const i = headers.findIndex(h => norm_(h) === norm_(name));
    if (i !== -1) return i;
  }
  return -1;
}

/** An item's barcode cell can hold several codes (e.g. a UPC and a QR), separated by commas. */
function splitCodes_(value) {
  return String(value ?? '').split(/[,;\n]+/).map(s => s.trim()).filter(Boolean);
}

/** Comparison key for codes: no spaces/dashes/case; all-digit codes ignore leading zeros (UPC-A vs EAN-13). */
function codeKey_(code) {
  const c = String(code).trim().toLowerCase().replace(/[\s-]/g, '');
  return /^\d+$/.test(c) ? c.replace(/^0+(?=\d)/, '') : c;
}

/** Map of field -> 0-based column index (or -1 when the column doesn't exist). */
function columns_(sheet) {
  const headers = headerRow_(sheet);
  const cols = {};
  Object.keys(FIELDS).forEach(f => cols[f] = findCol_(headers, FIELDS[f]));
  // Every other column with a header (except the monthly Starting Qty columns): searched in the app, not shown
  const used = new Set(Object.values(cols));
  Object.defineProperty(cols, 'extra', { enumerable: false,
    value: headers.map((h, i) => i).filter(i => String(headers[i] || '').trim() && !used.has(i) && !/starting qty/i.test(headers[i])) });
  ['name', 'quantity'].forEach(f => {
    if (cols[f] === -1) throw new Error('Column "' + [].concat(FIELDS[f])[0] + '" not found in the master sheet.');
  });
  return cols;
}

/** Last row that has an Item Name (the table extends past the data with empty rows). */
function lastItemRow_(sheet, cols) {
  const names = sheet.getRange(1, cols.name + 1, sheet.getLastRow(), 1).getValues();
  for (let i = names.length - 1; i >= 1; i--) if (String(names[i][0]).trim()) return i + 1;
  return 1;
}

function cell_(row, col) {
  if (col === -1) return '';
  const v = row[col];
  return v instanceof Date ? v.toISOString() : String(v ?? '').trim();
}

/**
 * Reads rows twice: raw values (for numbers like Available Today and prices) and the text exactly as the sheet shows it.
 * Text fields use the shown text, because Sheets silently turns entries like shelf "8-1" into a date (Aug 1)
 * that only looks right in the sheet because of its "m-d" format.
 */
function readRows_(sheet, startRow, numRows) {
  const range = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn());
  return { values: range.getValues(), shown: range.getDisplayValues() };
}

/** Cost Price as a number: cells may hold 420 (currency-formatted) or text like "$1,234.50". Blank -> null. */
function money_(value) {
  if (typeof value === 'number') return value;
  const n = parseFloat(String(value ?? '').replace(/[$,\s]/g, ''));
  return isFinite(n) ? n : null;
}

function rowToItem_(row, rowNum, cols, shown) {
  const item = { id: String(rowNum) };
  Object.keys(FIELDS).forEach(f => item[f] = cell_(shown || row, cols[f]));
  item.quantity = Number(row[cols.quantity]) || 0;
  item.cost = cols.cost === -1 ? null : money_(row[cols.cost]);
  // Extended Price comes from the sheet's formula; rows without one get cost × Available Today.
  // A blank cost makes the formula show $0, which would be misleading, so that counts as "not set".
  let ext = cols.extended === -1 ? null : money_(row[cols.extended]);
  if (item.cost == null && !ext) ext = null;
  item.extended = ext != null ? ext : item.cost != null ? Math.round(item.cost * item.quantity * 100) / 100 : null;
  item.key = keyOf_(item);
  // sx: the row's other columns as the sheet shows them (breakers, AIC rating, QB name, notes…), so search finds them too
  if (cols.extra) item.sx = cols.extra.map(c => cell_(shown || row, c)).filter(v => v && v !== '0').join(' · ').slice(0, 1500);
  return item;
}

/** Fingerprint used to make sure a row still holds the same item before writing (rows can be sorted/inserted). */
function keyOf_(item) {
  return [item.name, item.partNumber, item.job].map(norm_).join('|');
}

/**
 * Returns the current sheet row for an item the phone loaded earlier.
 * Checks the row still holds the same item; if the sheet was re-sorted, finds it again by fingerprint.
 */
function locate_(sheet, cols, id, key) {
  const last = lastItemRow_(sheet, cols);
  const rowNum = Number(id);
  if (rowNum >= 2 && rowNum <= last) {
    const one = readRows_(sheet, rowNum, 1);
    if (rowToItem_(one.values[0], rowNum, cols, one.shown[0]).key === key) return rowNum;
  }
  const all = readRows_(sheet, 2, last - 1);
  const hits = [];
  all.values.forEach((row, i) => { if (rowToItem_(row, i + 2, cols, all.shown[i]).key === key) hits.push(i + 2); });
  if (hits.length === 1) return hits[0];
  throw new Error('The sheet changed since you loaded it. Pull down / tap ⟳ to refresh, then try again.');
}

function readItem_(sheet, cols, rowNum) {
  const one = readRows_(sheet, rowNum, 1);
  return Object.assign(rowToItem_(one.values[0], rowNum, cols, one.shown[0]),
    { qtyColor: cellFill_(sheet.getRange(rowNum, cols.quantity + 1).getBackground()) });
}

/** A cell's fill as lowercase #rrggbb, or '' for no fill (white). */
function cellFill_(c) {
  const v = String(c || '').toLowerCase();
  return /^#[0-9a-f]{6}$/.test(v) && v !== '#ffffff' ? v : '';
}

// ---------- Contacts: vendors, subcontractors, electricians, freight, PMs, customers ----------
// One list of the people and companies the team deals with, in the "Contacts" tab, so an order's Point of Contact,
// Company / Vendor and Freight can be picked from it instead of typed again (and spelled differently) each time.
// Duplicates are refused: the same email, the same phone number, or the same name at the same company is already in
// the list. Near-matches (same name at another company, same company with no name…) are shown as a warning first,
// and can be saved anyway (force). Editors and Admins add, change, merge and delete; everyone can look them up.
const CONTACTS_SHEET = 'Contacts';
// 🏗 Project contacts: type "Project", Name = the project's name, plus its Project #, and what an order for it usually
// has (from / to location, activity type, freight). Any other contact can be linked to projects ("Projects", names
// separated by ;), so picking a project on an order fills in its PM, point of contact, vendor and freight too.
const CONTACT_HEADERS = ['ID', 'Type', 'Name', 'Company', 'Title / Role', 'Phone', 'Email', 'Address', 'Notes',
  'Created', 'Created By', 'Updated', 'Updated By', 'Project #', 'Projects', 'Order Defaults'];
const CONTACT_TYPES = ['Vendor', 'Subcontractor', 'Electrician', 'Freight / Carrier', 'PM', 'Customer / Site', 'Project', 'Other'];
const CONTACT_FIELDS = [['type', 40], ['name', 80], ['company', 100], ['title', 60], ['phone', 40], ['email', 120], ['address', 300], ['notes', 1000],
  ['projectNo', 40], ['projects', 800], ['defaults', 1500]];
// A project's order defaults: only these keys, short text each
const PROJECT_DEFAULT_KEYS = ['fromLocation', 'toLocation', 'activityType', 'freight', 'pm', 'contact', 'vendor'];
function cleanProjectDefaults_(v) {
  let o = {}; try { o = typeof v === 'string' ? (v ? JSON.parse(v) : {}) : (v || {}); } catch (e) { o = {}; }
  const out = {}; PROJECT_DEFAULT_KEYS.forEach(k => { const t = String(o && o[k] || '').replace(/\s+/g, ' ').trim().slice(0, 120); if (t) out[k] = t; });
  return out;
}
const contactsSheet_ = () => tabWithHeaders_(CONTACTS_SHEET, CONTACT_HEADERS);

function readContacts_() {
  const sheet = contactsSheet_(), last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, CONTACT_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2, id: String(r[0]), type: String(r[1]), name: String(r[2]), company: String(r[3]), title: String(r[4]),
    phone: String(r[5]), email: String(r[6]), address: String(r[7]), notes: String(r[8]),
    created: fmtTime_(r[9]), createdBy: String(r[10]), updated: fmtTime_(r[11]), updatedBy: String(r[12]),
    projectNo: String(r[13] || ''), projects: String(r[14] || ''), defaults: cleanProjectDefaults_(String(r[15] || '')),
  })).filter(c => c.id && (c.name || c.company));
}

// How two contacts are compared: names and companies without case, punctuation or "Inc / LLC / Co…"; phone numbers by
// their last 10 digits; emails without case.
const contactKey_ = s => String(s || '').toLowerCase().replace(/&/g, ' and ').replace(/[^a-z0-9]+/g, ' ')
  .replace(/\b(inc|llc|l l c|co|corp|corporation|company|ltd|the|pllc|lp)\b/g, ' ').replace(/\s+/g, ' ').trim();
const contactPhone_ = s => { const d = String(s || '').replace(/\D/g, ''); return d.length >= 7 ? d.slice(-10) : ''; };
const contactEmail_ = s => String(s || '').trim().toLowerCase();

/** { same: the contact this one duplicates (or null), similar: [near-matches] } among the others. */
function contactMatches_(c, others) {
  const n = contactKey_(c.name), co = contactKey_(c.company), ph = contactPhone_(c.phone), em = contactEmail_(c.email);
  let same = null; const similar = [];
  others.forEach(o => {
    const on = contactKey_(o.name), oco = contactKey_(o.company);
    const why = em && em === contactEmail_(o.email) ? 'same email'
      : ph && ph === contactPhone_(o.phone) ? 'same phone number'
      : (n || co) && n === on && co === oco ? (n ? 'same name' + (co ? ' and company' : '') : 'same company') : '';
    if (why) { if (!same) same = Object.assign({ why }, o); return; }
    const near = n && n === on ? 'same name, different company'
      : co && co === oco && (!n || !on) ? 'same company'
      : n && on && n.split(' ')[0] === on.split(' ')[0] && n.split(' ').pop() === on.split(' ').pop() && n.split(' ').length > 1 ? 'similar name'
      : '';
    if (near) similar.push(Object.assign({ why: near }, o));
  });
  return { same, similar: similar.slice(0, 5) };
}

/** Every contact, A–Z by name (or company). */
function getContacts() {
  requireRole_('Viewer');
  return readContacts_().map(c => { delete c.row; return c; })
    .sort((a, b) => (a.name || a.company).localeCompare(b.name || b.company));
}

/**
 * Adds a contact (no id) or changes one (id). Returns { contact } once saved, { duplicate } when it's already in the
 * list (nothing saved), or { similar } for near-matches when force isn't set (nothing saved yet).
 */
function saveContact(c, force) {
  requireRole_('Editor');
  requireScreen_('contacts');
  c = c || {};
  const clean = {};
  CONTACT_FIELDS.forEach(([k, max]) => {
    const v = k === 'defaults' ? JSON.stringify(cleanProjectDefaults_(c[k])) : String(c[k] ?? '').replace(/\r/g, '').trim();
    if (v.length > max) throw new Error('The ' + k + ' is too long (up to ' + max + ' characters).');
    clean[k] = k === 'notes' || k === 'address' || k === 'defaults' ? v : v.replace(/\s+/g, ' ');
  });
  clean.projects = clean.projects.split(/\s*;\s*/).filter(Boolean).filter((x, i, a) => a.findIndex(y => y.toLowerCase() === x.toLowerCase()) === i).join('; ');
  if (clean.type !== 'Project') clean.defaults = '{}';
  if (!clean.name && !clean.company) throw new Error('Enter a name or a company.');
  if (!CONTACT_TYPES.includes(clean.type)) clean.type = 'Other';
  if (clean.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clean.email)) throw new Error('That email address doesn’t look right.');
  return withLock_(() => {
    const all = readContacts_(), id = String(c.id || '');
    const old = id ? all.find(x => x.id === id) : null;
    if (id && !old) throw new Error('That contact was deleted. Refresh and try again.');
    const m = contactMatches_(clean, all.filter(x => x.id !== id));
    if (m.same) { delete m.same.row; return { duplicate: m.same }; }
    if (m.similar.length && !force) return { similar: m.similar.map(x => { delete x.row; return x; }) };
    const sheet = contactsSheet_(), me = currentUser_(), now = new Date();
    const cells = [clean.type, clean.name, clean.company, clean.title, clean.phone, clean.email, clean.address, clean.notes].map(asText_);
    let saved;
    if (old) {
      sheet.getRange(old.row, 2, 1, 8).setValues([cells]);
      sheet.getRange(old.row, 12, 1, 5).setValues([[now, me].concat([clean.projectNo, clean.projects, clean.defaults].map(asText_))]);
      saved = Object.assign({}, old, clean, { updated: now.toISOString(), updatedBy: me });
    } else {
      const newId = Utilities.getUuid();
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, CONTACT_HEADERS.length).setValues([[newId].concat(cells, [now, me, now, me], [clean.projectNo, clean.projects, clean.defaults].map(asText_))]);
      saved = Object.assign({ id: newId, created: now.toISOString(), createdBy: me, updated: now.toISOString(), updatedBy: me }, clean);
    }
    saved.defaults = cleanProjectDefaults_(saved.defaults);
    delete saved.row;
    logSheet_().appendRow([now, me, old ? 'Contact changed' : 'Contact added', '', saved.name || saved.company, '', '', 0, '',
      [saved.type, saved.company && saved.name ? saved.company : ''].filter(Boolean).join(' · ')]);
    return { contact: saved };
  });
}

/** Deletes a contact (orders that mention it keep their text). */
function deleteContact(id) {
  requireRole_('Editor');
  requireScreen_('contacts');
  return withLock_(() => {
    const c = readContacts_().find(x => x.id === String(id));
    if (!c) return true;
    contactsSheet_().deleteRow(c.row);
    logSheet_().appendRow([new Date(), currentUser_(), 'Contact deleted', '', c.name || c.company, '', '', 0, '', c.type]);
    return true;
  });
}

/**
 * Merges duplicates into one: keepId keeps its details, fills any blanks from the others (notes are added together),
 * and the others are deleted. Returns the merged contact.
 */
function mergeContacts(keepId, dropIds) {
  requireRole_('Editor');
  requireScreen_('contacts');
  return withLock_(() => {
    const all = readContacts_(), keep = all.find(x => x.id === String(keepId));
    if (!keep) throw new Error('That contact was deleted. Refresh and try again.');
    const drops = all.filter(x => (dropIds || []).map(String).includes(x.id) && x.id !== keep.id);
    if (!drops.length) throw new Error('Pick at least two contacts to merge.');
    const merged = Object.assign({}, keep);
    drops.forEach(d => ['type', 'name', 'company', 'title', 'phone', 'email', 'address', 'projectNo'].forEach(k => { if (!merged[k] && d[k]) merged[k] = d[k]; }));
    merged.projects = [keep.projects].concat(drops.map(d => d.projects)).join(';').split(/\s*;\s*/).filter(Boolean)
      .filter((x, i, a) => a.findIndex(y => y.toLowerCase() === x.toLowerCase()) === i).join('; ').slice(0, 800);
    drops.forEach(d => Object.keys(d.defaults || {}).forEach(k => { if (!merged.defaults[k]) merged.defaults[k] = d.defaults[k]; }));
    merged.notes = [keep.notes].concat(drops.map(d => d.notes)).filter(Boolean).join('\n').slice(0, 1000);
    const sheet = contactsSheet_(), me = currentUser_(), now = new Date();
    sheet.getRange(keep.row, 2, 1, 8).setValues([[merged.type, merged.name, merged.company, merged.title, merged.phone, merged.email, merged.address, merged.notes].map(asText_)]);
    sheet.getRange(keep.row, 12, 1, 5).setValues([[now, me].concat([merged.projectNo, merged.projects, JSON.stringify(merged.defaults)].map(asText_))]);
    drops.map(d => d.row).sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));   // bottom first, so row numbers stay right
    logSheet_().appendRow([now, me, 'Contacts merged', '', merged.name || merged.company, '', '', 0, '',
      'Merged in: ' + drops.map(d => d.name || d.company).join(', ')]);
    delete merged.row;
    return Object.assign(merged, { updated: now.toISOString(), updatedBy: me });
  });
}

// ---------- Labels: names for the Available Today highlight colours ----------
// The sheet colours some Available Today cells; the app names each colour (e.g. yellow = "On order") so the Labels
// filter above the Items list reads in the team's own words, and which colours appear there at all can be picked
// (Hidden = Yes leaves a colour out of that row). Shared by everyone and kept in the "Color Labels" tab (one row per
// colour), where they can also be edited by hand. A colour without a name shows its colour name (Light yellow, Green…).
const LABELS_SHEET = 'Color Labels';
const LABEL_HEADERS = ['Color', 'Label', 'Updated', 'Updated By', 'Hidden', 'Monthly Count'];
const MAX_LABEL = 30, MAX_LABELS = 60;
const labelsSheet_ = () => tabWithHeaders_(LABELS_SHEET, LABEL_HEADERS);
const isYes_ = v => /^(yes|y|true|x|1|hidden)$/i.test(String(v).trim());

/** The tab's rows: [{ hex, label, updated, by, hidden }]. */
function readLabelRows_() {
  const sheet = labelsSheet_(), last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, LABEL_HEADERS.length).getValues()
    .map(r => ({ hex: cellFill_(String(r[0]).trim()), label: String(r[1]).trim().slice(0, MAX_LABEL), updated: r[2], by: String(r[3]),
      hidden: isYes_(r[4]), monthly: String(r[5]).trim() === '' ? null : isYes_(r[5]) }))
    .filter(r => r.hex && (r.label || r.hidden || r.monthly !== null));
}

/**
 * { labels: { '#rrggbb': 'Label' }, hidden: ['#rrggbb', …], monthly: ['#rrggbb', …] }: the colours' names, those left
 * out of the Labels row, and those set to get a monthly starting count (see monthlyColors_) — monthly is null until
 * someone picks, meaning "the green ones".
 */
function getColorLabels() {
  requireRole_('Viewer');
  const rows = readLabelRows_(), labels = {};
  rows.forEach(r => { if (r.label) labels[r.hex] = r.label; });
  const picked = rows.some(r => r.monthly !== null);
  return { labels, hidden: rows.filter(r => r.hidden).map(r => r.hex), monthly: picked ? rows.filter(r => r.monthly).map(r => r.hex) : null };
}

/**
 * Names colours and picks which show: labels = { '#rrggbb': 'Label' } ('' takes a name away), hidden = { '#rrggbb':
 * true (leave out of the Labels row) / false (show) }, monthly = { '#rrggbb': true / false } (gets a monthly starting
 * count). Colours not mentioned keep what they had. Returns getColorLabels().
 */
function setColorLabels(labels, hidden, monthly) {
  requireRole_('Editor');
  const hexOf = c => { const h = cellFill_(String(c).trim()); if (!h) throw new Error('"' + c + '" isn’t a colour.'); return h; };
  const names = {}, hide = {};
  Object.keys(labels || {}).forEach(c => {
    const name = String(labels[c] ?? '').replace(/\s+/g, ' ').trim();
    if (name.length > MAX_LABEL) throw new Error('Labels can be up to ' + MAX_LABEL + ' characters ("' + name.slice(0, 20) + '…").');
    names[hexOf(c)] = name;
  });
  Object.keys(hidden || {}).forEach(c => { hide[hexOf(c)] = !!hidden[c]; });
  const month = {};
  Object.keys(monthly || {}).forEach(c => { month[hexOf(c)] = !!monthly[c]; });
  return withLock_(() => {
    const sheet = labelsSheet_(), last = sheet.getLastRow(), rows = readLabelRows_();
    const me = currentUser_(), now = new Date(), notes = [];
    // the first time anyone picks monthly colours, the green ones that were counted by default are written down too
    if (Object.keys(month).length && !rows.some(r => r.monthly !== null)) monthlyColors_(rows).forEach(h => { if (!(h in month)) month[h] = true; });
    [...new Set(Object.keys(names).concat(Object.keys(hide), Object.keys(month)))].forEach(hex => {
      let r = rows.find(x => x.hex === hex);
      const label = hex in names ? names[hex] : r ? r.label : '', hid = hex in hide ? hide[hex] : r ? r.hidden : false;
      const mon = hex in month ? month[hex] : r ? r.monthly : null;
      if (r && r.label === label && r.hidden === hid && r.monthly === mon) return;
      if (!r) { r = { hex, monthly: null }; rows.push(r); }
      if ((r.label || '') !== label) notes.push(hex + (label ? ' = ' + label : ' (name removed)'));
      if (!!r.hidden !== hid) notes.push(hex + (hid ? ' hidden' : ' shown'));
      if (r.monthly !== mon && mon !== null) notes.push(hex + (mon ? ' counted monthly' : ' not counted monthly'));
      Object.assign(r, { label, hidden: hid, monthly: mon, updated: now, by: me });
    });
    const keep = rows.filter(r => r.label || r.hidden || r.monthly !== null);
    if (keep.length > MAX_LABELS) throw new Error('Up to ' + MAX_LABELS + ' labels.');
    if (notes.length) {
      if (last >= 2) sheet.getRange(2, 1, last - 1, LABEL_HEADERS.length).clearContent();
      if (keep.length) sheet.getRange(2, 1, keep.length, LABEL_HEADERS.length)
        .setValues(keep.map(r => [r.hex, asText_(r.label), r.updated, r.by, r.hidden ? 'Yes' : '', r.monthly === null ? '' : r.monthly ? 'Yes' : 'No']));
      sheet.getRange(2, 1, Math.max(keep.length, 1), 1).setNumberFormat('@');
      logSheet_().appendRow([now, me, 'Labels changed', '', '', '', '', 0, '', notes.join('; ')]);
    }
    return getColorLabels();
  });
}

// ---------- Monthly starting counts ----------
// The master sheet keeps one "MM.DD.YY / Starting Qty" column per month, just before Available Today: each month's
// starting count of every row, newest last. The newest is filled dark green with white text, older ones grey. On the 1st
// of each month (a daily time trigger, turned on once: see setMonthlyCountsAuto), the app adds that month's column the
// same way it's done by hand: inserted before Available Today, formatted like last month's (whose column then turns
// grey like the older ones), titled with the date, and filled with Available Today for the rows whose Available Today
// colour is set to get a monthly count (Labels → ✎ Edit → 📅 Monthly; the green ones until that's been picked). Other
// rows are left blank. It never adds a second column for the same month.
const MONTH_COL_RE = /^(\d{2})\.(\d{2})\.(\d{2})\s*\/\s*Starting Qty$/i;
const MONTHLY_TRIGGER_FN = 'monthlyCountsTick';

/** The owner and admins see every item; everyone else only the green-labelled ones. */
function seesAllItems_() { return myRole_() === 'Admin' || currentUser_().toLowerCase() === ownerEmail_(); }
/** Is this a green (any shade, not grey)? Same idea as the app's colour names. */
function isGreen_(hex) {
  const n = parseInt(String(hex).slice(1), 16), r = (n >> 16 & 255) / 255, g = (n >> 8 & 255) / 255, b = (n & 255) / 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2, d = mx - mn;
  const sat = d === 0 ? 0 : d / (1 - Math.abs(2 * l - 1));
  if (sat < 0.15 || mx !== g) return false;
  const h = ((b - r) / d + 2) * 60;
  return h >= 70 && h < 165;
}

/** The colours that get a monthly count: the ones picked in Labels, or (until someone picks) every green in use. */
function monthlyColors_(rows) {
  rows = rows || readLabelRows_();
  if (rows.some(r => r.monthly !== null)) return rows.filter(r => r.monthly).map(r => r.hex);
  const sheet = master_(), cols = columns_(sheet), last = lastItemRow_(sheet, cols);
  if (last < 2) return [];
  return [...new Set(sheet.getRange(2, cols.quantity + 1, last - 1, 1).getBackgrounds().map(r => cellFill_(r[0])).filter(h => h && isGreen_(h)))];
}

/** The month columns ({ col, key: 'yyyy-mm' }, 1-based, left to right) and Available Today's column. */
function monthColumns_(sheet) {
  const headers = headerRow_(sheet), avail = findCol_(headers, FIELDS.quantity) + 1, list = [];
  headers.forEach((h, i) => { const m = MONTH_COL_RE.exec(String(h).trim()); if (m) list.push({ col: i + 1, key: '20' + m[3] + '-' + m[1], title: String(h).trim() }); });
  return { avail, list };
}

/**
 * Adds this month's Starting Qty column (when it isn't there yet). Returns { added, title, counted, skipped } —
 * added = false when this month already has one. who: who did it, for the App Log.
 */
function addMonthColumn_(who) {
  return withLock_(() => {
    const sheet = master_(), tz = Session.getScriptTimeZone(), now = new Date();
    const key = Utilities.formatDate(now, tz, 'yyyy-MM'), title = Utilities.formatDate(now, tz, 'MM.dd.yy') + ' / Starting Qty';
    let { avail, list } = monthColumns_(sheet);
    if (avail < 1) throw new Error('No "Available Today" column in the master sheet.');
    const had = list.find(m => m.key === key);
    if (had) return { added: false, title: had.title };
    const prev = list.length ? list[list.length - 1] : null, older = list.length > 1 ? list[list.length - 2] : null;
    const rows = sheet.getMaxRows();
    // the new column goes just before Available Today, looking like last month's (the newest look)
    sheet.insertColumnBefore(avail);
    const at = avail;   // the new column; Available Today is now one to the right
    avail += 1;
    if (prev) {
      const prevCol = prev.col < at ? prev.col : prev.col + 1;
      sheet.getRange(1, prevCol, rows, 1).copyTo(sheet.getRange(1, at, rows, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      sheet.setColumnWidth(at, sheet.getColumnWidth(prevCol));
      // last month's column now looks like the older ones
      if (older) sheet.getRange(1, older.col < at ? older.col : older.col + 1, rows, 1).copyTo(sheet.getRange(1, prevCol, rows, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }
    sheet.getRange(1, at).setNumberFormat('@').setValue(title);
    const cols = columns_(sheet), last = lastItemRow_(sheet, cols);
    const counted = new Set(monthlyColors_());
    let n = 0, skipped = 0;
    if (last >= 2) {
      const qty = sheet.getRange(2, avail, last - 1, 1).getValues(), fills = sheet.getRange(2, avail, last - 1, 1).getBackgrounds();
      const vals = qty.map((r, i) => {
        if (!counted.has(cellFill_(fills[i][0]))) { skipped++; return ['']; }
        n++; return [r[0] === '' ? 0 : r[0]];
      });
      sheet.getRange(2, at, last - 1, 1).setValues(vals);
    }
    PropertiesService.getScriptProperties().setProperty('MONTHLY_LAST', JSON.stringify({ title, at: now.toISOString(), counted: n, by: who }));
    logSheet_().appendRow([now, who, 'Monthly counts', '', '', '', '', 0, '', title + ' added: ' + n + ' row' + (n === 1 ? '' : 's') + ' counted'
      + (skipped ? ', ' + skipped + ' left blank (not a monthly colour)' : '')]);
    return { added: true, title, counted: n, skipped };
  });
}

/** The daily trigger: on the first run of a new month, adds that month's column. */
function monthlyCountsTick() {
  try { addMonthColumn_('Automatic (monthly counts)'); }
  catch (e) { logSheet_().appendRow([new Date(), 'Automatic (monthly counts)', 'Monthly counts failed', '', '', '', '', 0, '', String(e && e.message || e)]); }
}

const monthlyTriggers_ = () => ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === MONTHLY_TRIGGER_FN);

/** Admins: { auto, last: { title, at, counted, by }, thisMonth: title or '', colors: [...] } for the Labels screen. */
function getMonthlyCountsStatus() {
  requireRole_('Admin');
  let auto = null;
  try { auto = monthlyTriggers_().length > 0; } catch (e) { auto = null; }   // null: the script hasn't been allowed to manage triggers yet
  let last = null;
  try { last = JSON.parse(PropertiesService.getScriptProperties().getProperty('MONTHLY_LAST') || 'null'); } catch (e) { last = null; }
  const { list } = monthColumns_(master_()), key = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const mine = list.find(m => m.key === key);
  return { auto, last, thisMonth: mine ? mine.title : '', newest: list.length ? list[list.length - 1].title : '', colors: monthlyColors_() };
}

/** Admins: turns the automatic monthly column on (a daily trigger, early morning) or off. */
function setMonthlyCountsAuto(on) {
  requireRole_('Admin');
  return withLock_(() => {
    monthlyTriggers_().forEach(t => ScriptApp.deleteTrigger(t));
    if (on) ScriptApp.newTrigger(MONTHLY_TRIGGER_FN).timeBased().everyDays(1).atHour(1).create();
    logSheet_().appendRow([new Date(), currentUser_(), 'Monthly counts ' + (on ? 'turned on' : 'turned off'), '', '', '', '', 0, '', '']);
    return getMonthlyCountsStatus();
  });
}

/** Admins: adds this month's column now (if it isn't there yet). */
function addMonthlyCountsNow() {
  requireRole_('Admin');
  const r = addMonthColumn_(currentUser_());
  return Object.assign(r, { status: getMonthlyCountsStatus() });
}

/** Run once from the Apps Script editor (select it, ▶ Run) to allow and turn on the automatic monthly column. */
function turnOnMonthlyCounts() {
  monthlyTriggers_().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger(MONTHLY_TRIGGER_FN).timeBased().everyDays(1).atHour(1).create();
  Logger.log('Monthly counts are on: each month\'s Starting Qty column is added early on the 1st.');
}

/** Text typed in the app is stored as text: a leading apostrophe stops Sheets reading "=...", "+...", "-...", "@..." as a formula. */
function asText_(value) {
  const v = String(value);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function currentUser_() {
  return Session.getActiveUser().getEmail() || 'unknown';
}

function log_(action, item, change, note) {
  logSheet_().appendRow([
    new Date(), currentUser_(), action, Number(item.id), item.name, item.partNumber, item.job, change, item.quantity, note || '',
  ]);
}

/** Serializes writes so two people editing at once don't clobber each other. */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function photoFolder_() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}

// ---------- API called from the web app via google.script.run ----------

function getItems() {
  requireRole_('Viewer');
  const sheet = master_();
  const cols = columns_(sheet);
  const last = lastItemRow_(sheet, cols);
  if (last < 2) return [];
  const counts = {}, docsByKey = {};
  readComments_(commentsSheet_()).forEach(c => counts[c.key] = (counts[c.key] || 0) + 1);
  // Each item's PDFs (newest first, like its Documents section) so the list can show them as buttons
  const tagsByKey = {}, minByKey = {};
  readItemTags_().forEach(t => tagsByKey[t.key] = t);
  readMinStock_().forEach(m => minByKey[m.key] = m);
  const epByKey = {};   // each item's endpoints, from every order line it's on
  const snByKey = {};   // and their serial #s, so items can be found by serial
  const rcvEp = {}, rcvSn = {};   // the same from receipts (Receive a PO / Returning inventory)
  readLineRows_().forEach(l => {
    if (l.endpoints.length) epByKey[l.key] = unionTags_([epByKey[l.key], l.endpoints]);
    if (l.serials.some(Boolean)) snByKey[l.key] = unionTags_([snByKey[l.key], l.serials.filter(Boolean)]);
  });
  readAdjustments_().forEach(a => {
    if (a.endpoints.length) rcvEp[a.key] = unionTags_([rcvEp[a.key], a.endpoints]);
    if (a.serials.some(Boolean)) rcvSn[a.key] = unionTags_([rcvSn[a.key], a.serials.filter(Boolean)]);
  });
  docsInOrder_(readDocs_(docsSheet_())).forEach(d => (docsByKey[d.key] = docsByKey[d.key] || []).push({ fileId: d.fileId, name: d.name, url: d.url }));
  const all = readRows_(sheet, 2, last - 1);
  // The Available Today cell's fill colour in the sheet (e.g. green), so the app can show just one colour's rows
  const fills = sheet.getRange(2, cols.quantity + 1, last - 1, 1).getBackgrounds();
  // Only the owner and admins see every item; everyone else sees just the rows whose Available Today cell is filled green.
  const seeAll = seesAllItems_();
  return all.values
    .map((row, i) => Object.assign(rowToItem_(row, i + 2, cols, all.shown[i]), { qtyColor: cellFill_(fills[i][0]) }))
    .filter(item => item.name && (seeAll || isGreen_(item.qtyColor)))
    .map(item => Object.assign(item, { tags: (tagsByKey[item.key] || {}).tags || [], endpoints: unionTags_([epByKey[item.key], rcvEp[item.key]]), serials: unionTags_([snByKey[item.key], rcvSn[item.key]]),
      minTrack: !!(minByKey[item.key] || {}).track, minQty: (minByKey[item.key] || {}).min ?? null,
      rcvEndpoints: rcvEp[item.key] || [], rcvSerials: rcvSn[item.key] || [], commentCount: counts[item.key] || 0, documentCount: (docsByKey[item.key] || []).length, documents: docsByKey[item.key] || [] }));
}

/** Sets "Available Today" by adding delta (negative to remove). Refuses to go below zero. */
/** The quick +/− buttons. Recorded in Inventory Adjustments as a "Quick adjust". */
function adjustQuantity(id, key, delta, note) {
  requireRole_('Editor');
  return changeStock_(id, key, delta, { type: 'Quick adjust', note: note }).item;
}

/** The Adjust stock form: add or remove stock with a type, vendor, PO # and note. Returns { item, adjustment }. */
function adjustStock(id, key, delta, info) {
  requireRole_('Editor');
  info = cleanAdjInfo_(info);
  if (!info.type) info.type = Number(delta) > 0 ? 'Received' : 'Removed';
  return changeStock_(id, key, delta, info);
}

function cleanAdjInfo_(info) {
  info = info || {};
  const clip = (v, max, what) => {
    const t = String(v || '').trim().replace(/\s+/g, ' ');
    if (t.length > max) throw new Error(what + ' is too long (' + max + ' characters max).');
    return t;
  };
  return { type: clip(info.type, 40, 'Type'), vendor: clip(info.vendor, 100, 'Vendor'), po: clip(info.po, 60, 'PO #'), note: clip(info.note, 300, 'Note'),
    from: clip(info.from, 100, 'Returned from'), orderNo: clip(info.orderNo, 40, 'Order #'), reason: clip(info.reason, 60, 'Reason'),
    contact: clip(info.contact, 100, 'Point of contact'), pm: clip(info.pm, 100, 'PM') };
}

/**
 * Receives several items against one PO in one step: lines = [{ id, key, qty }], info = { vendor, po, note }.
 * Every item is found first, so nothing changes unless all of them can be received. Returns { items, adjustments }.
 */
function receivePo(lines, info) {
  requireRole_('Editor');
  requireScreen_('stock');
  info = cleanAdjInfo_(info);
  info.type = 'Received from vendor';
  info.from = info.orderNo = info.reason = '';
  if (!info.po && !info.vendor) throw new Error('Enter a PO # or a vendor.');
  return receiveLines_(lines, info);
}

/**
 * Returning inventory: stock coming back from a job, project or customer (e.g. unused items from a shipped order).
 * info = { from, orderNo, reason, contact, note }. Adds every quantity to Available Today and records one "Returned" adjustment per item.
 */
function receiveReturn(lines, info) {
  requireRole_('Editor');
  requireScreen_('stock');
  info = cleanAdjInfo_(info);
  info.type = 'Returned';
  info.po = '';   // a return can still name the vendor / company it's tied to (picked from contacts)
  if (!info.from && !info.orderNo) throw new Error('Enter where it came back from or an order #.');
  return receiveLines_(lines, info);
}

/**
 * Adds several items' quantities in one locked step (Receive a PO and Returning inventory). Each line may carry
 * endpoints + serials (EVSE items only), saved with its adjustment.
 */
function receiveLines_(lines, info) {
  const merged = {};
  (lines || []).forEach(l => {
    const qty = Number(l.qty);
    if (!(qty > 0) || Math.floor(qty) !== qty || qty > 100000) throw new Error('Each quantity must be a whole number of 1 or more.');
    const k = String(l.key), m = merged[k];
    merged[k] = m ? Object.assign(m, { qty: m.qty + qty, endpoints: m.endpoints.concat(l.endpoints || []), serials: m.serials.concat(l.serials || []) })
      : { id: l.id, key: k, qty, endpoints: (l.endpoints || []).slice(), serials: (l.serials || []).slice() };
  });
  Object.values(merged).forEach(m => Object.assign(m, pairEndpoints_(m.endpoints, m.serials)));
  const withEps = Object.values(merged).some(m => m.endpoints.length), byKey = withEps ? itemsByKey_() : {};
  if (withEps) Object.values(merged).forEach(m => checkEndpointsAllowed_(byKey[m.key], m.endpoints, []));
  const list = Object.values(merged);
  if (!list.length) throw new Error('Add at least one item.');
  if (list.length > 200) throw new Error('Receive up to 200 different items at a time.');
  return withLock_(() => {
    if (withEps) checkEndpointsUnique_(list.map(m => ({ name: (byKey[m.key] || {}).name || '', endpoints: m.endpoints })), endpointUses_());
    const sheet = master_();
    const cols = columns_(sheet);
    const rows = list.map(l => locate_(sheet, cols, l.id, l.key));   // throws before any change if an item has moved or gone
    const out = { items: [], adjustments: [] };
    list.forEach((l, i) => {
      const res = applyStockChange_(sheet, cols, rows[i], l.qty, Object.assign({}, info, { endpoints: l.endpoints, serials: l.serials }));
      out.items.push(res.item); out.adjustments.push(res.adjustment);
    });
    return out;
  });
}

function changeStock_(id, key, delta, info) {
  delta = Number(delta);
  if (!delta || Math.floor(delta) !== delta) throw new Error('Change must be a whole number other than zero.');
  return withLock_(() => {
    const sheet = master_();
    const cols = columns_(sheet);
    return applyStockChange_(sheet, cols, locate_(sheet, cols, id, key), delta, info);
  });
}

/** Changes one row's Available Today and records it (Inventory Adjustments + App Log). Call inside withLock_. */
function applyStockChange_(sheet, cols, rowNum, delta, info) {
  const qtyCell = sheet.getRange(rowNum, cols.quantity + 1);
  const current = Number(qtyCell.getValue()) || 0;
  const next = current + delta;
  if (next < 0) throw new Error('Not enough available (have ' + current + ').');
  qtyCell.setValue(next);
  const item = readItem_(sheet, cols, rowNum);
  const adjustment = {
    id: Utilities.getUuid(), timestamp: new Date().toISOString(), user: currentUser_(), type: info.type || '',
    change: delta, before: current, after: next, vendor: info.vendor || '', po: info.po || '', note: String(info.note || ''),
    from: info.from || '', orderNo: info.orderNo || '', reason: info.reason || '', contact: info.contact || '', pm: info.pm || '',
    endpoints: info.endpoints || [], serials: info.serials || [],
  };
  adjSheet_().appendRow([new Date(), adjustment.user, asText_(adjustment.type), asText_(item.name), asText_(item.partNumber), asText_(item.job),
    delta, current, next, asText_(adjustment.vendor), asText_(adjustment.po), asText_(adjustment.note), asText_(item.key), adjustment.id,
    asText_(adjustment.from), asText_(adjustment.orderNo), asText_(adjustment.reason)].concat(endpointCells_(adjustment), [asText_(adjustment.contact), '', asText_(adjustment.pm)]));
  const details = [info.type !== 'Quick adjust' && info.type, info.vendor && 'Vendor: ' + info.vendor, info.po && 'PO #: ' + info.po,
    info.from && 'From: ' + info.from, info.orderNo && 'Order #: ' + info.orderNo, info.reason, info.contact && 'Contact: ' + info.contact,
    (info.endpoints || []).length && 'Endpoints: ' + info.endpoints.map((p, i) => p + (info.serials[i] ? ' (S/N ' + info.serials[i] + ')' : '')).join(', '),
    info.note].filter(Boolean).join(' · ');
  log_(delta > 0 ? 'Stock In' : 'Stock Out', item, delta, details);
  return { item, adjustment };
}

const adjSheet_ = () => tabWithHeaders_(ADJ_SHEET, ADJ_HEADERS);

function readAdjustments_() {
  const sheet = adjSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, ADJ_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2, timestamp: fmtTime_(r[0]), user: String(r[1]), type: String(r[2]), name: String(r[3]), partNumber: String(r[4]),
    job: String(r[5]), change: Number(r[6]) || 0, before: Number(r[7]) || 0, after: Number(r[8]) || 0, vendor: String(r[9]),
    po: String(r[10]), note: String(r[11]), key: String(r[12]), id: String(r[13]),
    from: String(r[14] ?? ''), orderNo: String(r[15] ?? ''), reason: String(r[16] ?? ''),
    ...readEndpoints_(r[17], r[18]), contact: String(r[19] ?? ''), tracker: String(r[20] ?? ''), pm: String(r[21] ?? ''), links: jsonList_(r[22]), pdfs: jsonList_(r[23]),
  })).filter(a => a.change);
}

/** Stock adjustments, newest first. key: only that item's. Without a key, the latest `limit` (default 500) for the whole inventory. */
function getAdjustments(key, limit) {
  requireRole_('Viewer');
  let list = readAdjustments_();
  if (key) list = list.filter(a => a.key === key);
  return list.reverse().slice(0, Math.min(Number(limit) || 500, 2000)).map(a => { delete a.row; return a; });
}

/** Keeps an item's adjustment history attached when the item is renamed in the app. */
function moveAdjustments_(oldKey, item) {
  const sheet = adjSheet_();
  readAdjustments_().filter(a => a.key === oldKey).forEach(a => {
    sheet.getRange(a.row, 4, 1, 3).setValues([[asText_(item.name), asText_(item.partNumber), asText_(item.job)]]);
    sheet.getRange(a.row, 13).setValue(asText_(item.key));
  });
}

/**
 * Updates an existing row (when item.id is set) or adds a new row after the last item. Only EDITABLE fields are written,
 * plus item.cost (the unit price, "Cost Price") when it's sent: only Admins (the app's owner is always one) may change it.
 */
function saveItem(item) {
  requireRole_('Editor');
  if (!item || !String(item.name || '').trim()) throw new Error('Item Name is required.');
  item.quantity = Math.max(0, Number(item.quantity) || 0);
  const priceText = item.cost === undefined || item.cost === null ? '' : String(item.cost).trim();
  const price = priceText === '' ? null : money_(priceText);
  if (priceText !== '' && (price === null || price < 0 || price > 1e8)) throw new Error('Enter the unit price as an amount, like 420 or $1,250.50.');
  return withLock_(() => {
    const sheet = master_();
    const cols = columns_(sheet);
    if (String(item.barcode || '').trim() && cols.barcode === -1) throw new Error(NO_BARCODE_COLUMN);
    let rowNum, before = null;
    if (item.id) {
      rowNum = locate_(sheet, cols, item.id, item.key);
      before = readItem_(sheet, cols, rowNum);
    } else {
      rowNum = lastItemRow_(sheet, cols) + 1;
      if (rowNum > sheet.getMaxRows()) sheet.insertRowAfter(sheet.getMaxRows());
    }
    // Unit price: checked before anything is written, so a refused change leaves the row untouched
    const was = before ? before.cost : null, next = price === null ? null : Math.round(price * 100) / 100;
    const priceChanged = item.cost !== undefined && cols.cost !== -1 && next !== was;
    if (priceChanged && myRole_() !== 'Admin') throw new Error('Only admins can change the unit price.');
    EDITABLE.forEach(f => {
      if (cols[f] === -1 || item[f] === undefined) return;
      // Plain-text cells can't become formulas, so they get the text exactly as typed (no protective apostrophe).
      const text = String(item[f]).trim();
      const value = f === 'quantity' ? item.quantity : TEXT_FIELDS.includes(f) ? text : asText_(text);
      if (before && String(before[f]) === String(value)) return;   // only touch cells that changed
      const cell = sheet.getRange(rowNum, cols[f] + 1);
      if (TEXT_FIELDS.includes(f)) cell.setNumberFormat('@');
      cell.setValue(value);
    });
    if (!before && cols.extended !== -1 && cols.cost !== -1) {
      const a1 = c => sheet.getRange(rowNum, c + 1).getA1Notation();
      sheet.getRange(rowNum, cols.extended + 1).setFormula('=' + a1(cols.cost) + '*' + a1(cols.quantity));
    }
    // Unit price: written only when it was sent and differs from what's there (Extended Price follows by its formula)
    let priceNote = '';
    if (priceChanged) {
      const cell = sheet.getRange(rowNum, cols.cost + 1);
      if (next === null) cell.clearContent(); else { cell.setValue(next); if (!before) cell.setNumberFormat('$#,##0.00'); }
      const usd = v => v === null ? 'not set' : '$' + v.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
      priceNote = 'Unit price ' + usd(was) + ' → ' + usd(next);
    }
    const saved = readItem_(sheet, cols, rowNum);
    if (before && before.key !== saved.key) { moveComments_(before.key, saved); moveDocs_(before.key, saved); moveOrderLines_(before.key, saved); moveAdjustments_(before.key, saved); moveItemTags_(before.key, saved); moveMinStock_(before.key, saved); }
    if (priceChanged && before) {
      saved.repriced = repriceDrafts_(saved.key, next);
      if (saved.repriced.length) priceNote += ' (also on draft ' + saved.repriced.join(', ') + ')';
    }
    log_(before ? 'Edit' : 'Create', saved, saved.quantity - (before ? before.quantity : 0), priceNote);
    return saved;
  });
}

/**
 * A changed unit price flows into every Draft order's lines for that item (their Cost Price and Our Cost) and those
 * orders' Our Cost totals. Shipped and cancelled orders keep the price they had. Returns the draft order numbers changed.
 */
function repriceDrafts_(key, price) {
  const drafts = new Set(readOrderRows_().filter(o => o.status === 'Draft').map(o => o.orderNo));
  const ls = linesSheet_(), costCol = LINE_HEADERS.indexOf('Cost Price') + 1, touched = new Set();
  readLineRows_().forEach(l => {
    if (l.key !== key || !drafts.has(l.orderNo)) return;
    ls.getRange(l.row, costCol, 1, 2).setValues([[price === null ? '' : price, price === null ? '' : roundMoney_(price * l.qty)]]);
    touched.add(l.orderNo);
  });
  if (!touched.size) return [];
  const lines = readLineRows_(), os = ordersSheet_(), H = h => ORDER_HEADERS.indexOf(h) + 1;
  readOrderRows_().filter(o => touched.has(o.orderNo)).forEach(o => {
    const total = lines.filter(l => l.orderNo === o.orderNo).reduce((n, l) => n + (l.ourCost || 0), 0);
    os.getRange(o.row, H('Our Cost')).setValue(roundMoney_(total));
    os.getRange(o.row, H('Updated'), 1, 2).setValues([[new Date(), currentUser_()]]);
  });
  return [...touched];
}

const NO_BARCODE_COLUMN = 'No "Barcode / QR Code" column found in the master sheet. Run setup() in Apps Script.';

/**
 * Attaches a scanned code to an item by adding it to the item's "Barcode / QR Code" cell.
 * Refuses if another row already has that code, so each code leads to exactly one item.
 */
function attachBarcode(id, key, code) {
  requireRole_('Editor');
  code = String(code || '').trim();
  if (!code) throw new Error('No code to attach.');
  return withLock_(() => {
    const sheet = master_();
    const cols = columns_(sheet);
    if (cols.barcode === -1) throw new Error(NO_BARCODE_COLUMN);
    const rowNum = locate_(sheet, cols, id, key);
    const owner = getItems().find(i => i.id !== String(rowNum) && splitCodes_(i.barcode).some(c => codeKey_(c) === codeKey_(code)));
    if (owner) throw new Error('Code ' + code + ' is already attached to "' + owner.name + '" (row ' + owner.id + ').');
    const cell = sheet.getRange(rowNum, cols.barcode + 1);
    const codes = splitCodes_(cell.getDisplayValue());
    if (!codes.some(c => codeKey_(c) === codeKey_(code))) codes.push(code);
    cell.setNumberFormat('@').setValue(codes.join(', '));
    const item = readItem_(sheet, cols, rowNum);
    log_('Barcode', item, 0, 'Attached ' + code);
    return item;
  });
}

const MAX_PHOTOS = 5;
const NO_PHOTO_COLUMN = 'No "Photo ID" column found in the master sheet. Run setup() in Apps Script.';

/**
 * Runs fn(photoIds) against an item's "Photo ID" cell, which holds up to MAX_PHOTOS Drive file IDs
 * separated by commas (the first one is the main photo), then writes back whatever fn returns.
 */
function editPhotos_(id, key, action, fn) {
  return withLock_(() => {
    const sheet = master_();
    const cols = columns_(sheet);
    if (cols.photoId === -1) throw new Error(NO_PHOTO_COLUMN);
    const rowNum = locate_(sheet, cols, id, key);
    const cell = sheet.getRange(rowNum, cols.photoId + 1);
    const photos = fn(splitCodes_(cell.getDisplayValue()), readItem_(sheet, cols, rowNum));
    cell.setNumberFormat('@').setValue(photos.join(', '));
    const item = readItem_(sheet, cols, rowNum);
    log_(action, item, 0, photos.length + ' photo' + (photos.length === 1 ? '' : 's'));
    return item;
  });
}

/** Saves a photo (data URL from the phone) to Drive and adds it to the item's gallery. */
function uploadPhoto(id, key, dataUrl) {
  requireRole_('Editor');
  const m = /^data:(image\/[\w+.-]+);base64,(.+)$/.exec(dataUrl || '');
  if (!m) throw new Error('Invalid image.');
  return editPhotos_(id, key, 'Photo added', (photos, item) => {
    if (photos.length >= MAX_PHOTOS) throw new Error('This item already has ' + MAX_PHOTOS + ' photos. Delete one first.');
    const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], item.name.slice(0, 60) + ' ' + (photos.length + 1) + '.jpg');
    return photos.concat(photoFolder_().createFile(blob).getId());
  });
}

/** Removes a photo from the item's gallery and moves the file to Drive's trash (restorable for 30 days). */
function deletePhoto(id, key, fileId) {
  requireRole_('Editor');
  return editPhotos_(id, key, 'Photo deleted', photos => {
    if (!photos.includes(fileId)) throw new Error('That photo is no longer on this item. Refresh and try again.');
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* already gone */ }
    return photos.filter(p => p !== fileId);
  });
}

/** Makes a photo the item's main photo (first in the gallery). */
function setMainPhoto(id, key, fileId) {
  requireRole_('Editor');
  return editPhotos_(id, key, 'Main photo set', photos => {
    if (!photos.includes(fileId)) throw new Error('That photo is no longer on this item. Refresh and try again.');
    return [fileId].concat(photos.filter(p => p !== fileId));
  });
}

/**
 * Returns a photo as a data URL, so photos stay private in Drive (no public sharing needed).
 * Only serves files listed in the Photo ID column: the app runs as its owner, so it must not hand out other Drive files.
 */
function getPhoto(fileId) {
  requireRole_('Viewer');
  if (!fileId) return '';
  if (!listedPhotos_().has(fileId)) throw new Error('Photo not found.');
  return dataUrl_(DriveApp.getFileById(fileId).getBlob());
}

/**
 * Small previews for the lists: { fileId: dataUrl }, up to 12 photos per call. size 'lg' = a bigger copy (640 px on its
 * long side) for the Banner / Grid picture styles; otherwise about 220 px. Drive makes these copies at whatever size is
 * asked for, so a big phone photo (tall or wide) arrives as a few KB instead of the whole picture, and all the photos
 * asked for are fetched at the same time (UrlFetchApp.fetchAll). Each copy is kept in the script cache for 6 hours, so
 * the next phone to ask gets it at once. If Drive's copy can't be had, Drive's own small thumbnail is used; a photo with
 * neither (Drive may still be making them for one just uploaded) comes back '' unless the photo itself is small, so a
 * list is never held up downloading full-size pictures.
 */
const THUMB_PX = { sm: 220, lg: 640 };
const THUMB_CACHE_S = 21600;
function getThumbnails(fileIds, size) {
  requireRole_('Viewer');
  const px = THUMB_PX[size] || THUMB_PX.sm;
  const ids = [...new Set([].concat(fileIds || []).map(String))].slice(0, 12);
  const allowed = thumbAllowed_(ids);
  const out = {}, cache = CacheService.getScriptCache(), key = id => 'th' + px + ':' + id;
  let hits = {};
  try { hits = cache.getAll(ids.map(key)); } catch (e) { /* cache unavailable: fetch them all */ }
  const need = ids.filter(id => { if (!allowed.has(id)) return false; if (hits[key(id)]) { out[id] = hits[key(id)]; return false; } return true; });
  if (need.length) {
    driveThumbs_(need, px, out);
    need.forEach(id => { if (out[id] === undefined) out[id] = thumbFallback_(id); });
    const put = {};
    need.forEach(id => { if (out[id] && out[id].length < 95000) put[key(id)] = out[id]; });   // the cache holds up to 100 KB per entry
    try { cache.putAll(put, THUMB_CACHE_S); } catch (e) { /* not cached this time */ }
  }
  return out;
}

/** The photo IDs the lists may show (items' Photo IDs and order photos); kept in the cache for 5 minutes. */
function thumbAllowed_(ids) {
  const cache = CacheService.getScriptCache();
  let list = null;
  try { list = JSON.parse(cache.get('thumbAllowed') || 'null'); } catch (e) { list = null; }
  let set = new Set(list || []);
  if (!list || ids.some(id => !set.has(id))) {   // a photo added since: look again
    set = listedPhotos_();
    readOrderPhotos_().forEach(p => set.add(p.fileId));
    try { const json = JSON.stringify([...set]); if (json.length < 95000) cache.put('thumbAllowed', json, 300); } catch (e) { /* not cached */ }
  }
  return set;
}

/** Drive's resized copies (px on the long side) of several photos at once, into out[id]; leaves out the ones it can't get. */
function driveThumbs_(ids, px, out) {
  try {
    const auth = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
    const metas = UrlFetchApp.fetchAll(ids.map(id => ({ url: 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(id)
      + '?fields=thumbnailLink&supportsAllDrives=true', headers: auth, muteHttpExceptions: true })));
    const links = metas.map(r => { try { return r.getResponseCode() === 200 ? JSON.parse(r.getContentText()).thumbnailLink || '' : ''; } catch (e) { return ''; } });
    const want = ids.map((id, n) => links[n] ? n : -1).filter(n => n >= 0);
    if (!want.length) return;
    const imgs = UrlFetchApp.fetchAll(want.map(n => ({ url: links[n].replace(/=s\d+[^/=]*$/, '') + '=s' + px, headers: auth, muteHttpExceptions: true })));
    want.forEach((n, k) => {
      const r = imgs[k], blob = r.getResponseCode() === 200 ? r.getBlob() : null;
      if (blob && /^image\//.test(blob.getContentType() || '')) out[ids[n]] = dataUrl_(blob);
    });
  } catch (e) { /* Drive's API not reachable: thumbFallback_ still works */ }
}

/** Drive's own small thumbnail; or the photo itself only when it's small (≤ 300 KB); otherwise ''. */
function thumbFallback_(id) {
  try {
    const file = DriveApp.getFileById(id), t = file.getThumbnail();
    if (t) return dataUrl_(t);
    return file.getSize() <= 300 * 1024 ? dataUrl_(file.getBlob()) : '';
  } catch (e) {
    return '';
  }
}

/** Every Drive file ID in the Photo ID column. */
function listedPhotos_() {
  const sheet = master_();
  const cols = columns_(sheet);
  if (cols.photoId === -1) throw new Error(NO_PHOTO_COLUMN);
  const last = lastItemRow_(sheet, cols);
  const ids = new Set();
  if (last >= 2) sheet.getRange(2, cols.photoId + 1, last - 1, 1).getDisplayValues().forEach(r => splitCodes_(r[0]).forEach(id => ids.add(id)));
  return ids;
}

function dataUrl_(blob) {
  return 'data:' + (blob.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

// ---------- Comments ----------

/** Comments for one item, newest first, plus the viewer's email (so the app can offer delete on their own comments). */
function getComments(key) {
  requireRole_('Viewer');
  const list = readComments_(commentsSheet_()).filter(c => c.key === key).reverse();
  const rx = reactionsByComment_();
  return {
    me: currentUser_(),
    comments: list.map(c => ({ id: c.id, timestamp: c.timestamp, user: c.user, text: c.text, reactions: rx[c.id] || [] })),
  };
}

// ----- Comments on orders -----
// Same idea as item comments: stored in the "Order Comments" tab, newest first, people delete only their own.
const ORDER_COMMENTS_SHEET = 'Order Comments';
const ORDER_COMMENT_HEADERS = ['Timestamp', 'User', 'Order #', 'Project', 'Comment', 'Comment ID'];
const orderCommentsSheet_ = () => tabWithHeaders_(ORDER_COMMENTS_SHEET, ORDER_COMMENT_HEADERS);

function readOrderComments_() {
  const sheet = orderCommentsSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, ORDER_COMMENT_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2, timestamp: fmtTime_(r[0]), at: r[0] instanceof Date ? r[0].getTime() : 0, user: String(r[1]), orderNo: String(r[2]),
    project: String(r[3]), text: String(r[4]), id: String(r[5]),
  })).filter(c => c.id);
}

/** Comments on one order, newest first, each with its emoji reactions. */
function getOrderComments(orderNo) {
  requireRole_('Viewer');
  const rx = reactionsByComment_();
  return {
    me: currentUser_(),
    comments: readOrderComments_().filter(c => c.orderNo === orderNo).reverse()
      .map(c => ({ id: c.id, timestamp: c.timestamp, user: c.user, text: c.text, reactions: rx[c.id] || [] })),
  };
}

function addOrderComment(orderNo, text) {
  requireRole_('Viewer');   // everyone with access can comment (and delete their own)
  text = String(text || '').trim();
  if (!text) throw new Error('Write a comment first.');
  if (text.length > MAX_COMMENT) throw new Error('Comments can be up to ' + MAX_COMMENT + ' characters.');
  return withLock_(() => {
    const order = readOrderRows_().find(o => o.orderNo === orderNo);
    if (!order) throw new Error('Order ' + orderNo + ' was not found.');
    const project = [order.projectName, order.projectNo && '#' + order.projectNo].filter(String).join(' ');
    orderCommentsSheet_().appendRow([new Date(), currentUser_(), orderNo, asText_(project), asText_(text), Utilities.getUuid()]);
    bumpPulse_(text);
    return getOrderComments(orderNo);
  });
}

/** Deletes a comment on an order (only your own), along with its reactions. */
function deleteOrderComment(commentId) {
  requireRole_('Viewer');   // everyone with access can comment (and delete their own)
  return withLock_(() => {
    const c = readOrderComments_().find(x => x.id === commentId);
    if (!c) throw new Error('That comment was already deleted.');
    if (c.user !== currentUser_()) throw new Error('You can only delete your own comments.');
    orderCommentsSheet_().deleteRow(c.row);
    const rs = reactionsSheet_();
    readReactions_().filter(x => x.commentId === commentId).map(x => x.row).sort((a, b) => b - a).forEach(row => rs.deleteRow(row));
    return getOrderComments(c.orderNo);
  });
}

// ----- Emoji reactions on comments -----
// One row per person per emoji per comment in the "Comment Reactions" tab. Tapping an emoji again removes your reaction.
const REACTIONS_SHEET = 'Comment Reactions';
const REACTION_HEADERS = ['Timestamp', 'User', 'Comment ID', 'Emoji'];
// Any single emoji from the phone's emoji keyboard: faces, people with skin tones, families (joined with ZWJ), flags, keycaps.
const EMOJI_RE = /^(?:\p{Extended_Pictographic}|\p{Emoji_Modifier}|\p{Regional_Indicator}|[\u200D\uFE0F\u20E3#*0-9]|[\u{E0020}-\u{E007F}])+$/u;
const EMOJI_CORE_RE = /\p{Extended_Pictographic}|\p{Regional_Indicator}|\u20E3/u;
const isEmoji_ = e => e.length > 0 && e.length <= 32 && EMOJI_RE.test(e) && EMOJI_CORE_RE.test(e);
// 😀 one of a person's own built emojis (from My emoji), as the same [emo:…] token comments use
const isCemo_ = e => e.length > 0 && e.length < 260 && /^\[emo:[a-z0-9.~-]+\]$/.test(e);
const reactionsSheet_ = () => tabWithHeaders_(REACTIONS_SHEET, REACTION_HEADERS);

function readReactions_() {
  const sheet = reactionsSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, REACTION_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2, user: String(r[1]), commentId: String(r[2]), emoji: String(r[3]),
  })).filter(x => x.commentId && x.emoji);
}

/** comment id -> [{ emoji, users: [email, ...] }], emojis in the order they were first used, people in the order they reacted. */
function reactionsByComment_(rows) {
  const out = {};
  (rows || readReactions_()).forEach(x => {
    const list = out[x.commentId] || (out[x.commentId] = []);
    let e = list.find(y => y.emoji === x.emoji);
    if (!e) list.push(e = { emoji: x.emoji, users: [] });
    if (e.users.indexOf(x.user) === -1) e.users.push(x.user);
  });
  return out;
}

/** Adds your reaction (or removes it if it's already there). Returns { commentId, reactions } for that comment. */
function toggleReaction(commentId, emoji) {
  requireRole_('Editor');
  commentId = String(commentId || ''); emoji = String(emoji || '');
  if (!isEmoji_(emoji) && !isCemo_(emoji)) throw new Error('Reactions must be a single emoji.');
  return withLock_(() => {
    if (!readComments_(commentsSheet_()).some(c => c.id === commentId) && !readOrderComments_().some(c => c.id === commentId)) {
      throw new Error('That comment was deleted.');
    }
    const me = currentUser_(), sheet = reactionsSheet_();
    const rows = readReactions_();
    const mine = rows.filter(x => x.commentId === commentId && x.emoji === emoji && x.user === me);
    if (mine.length) mine.map(x => x.row).sort((a, b) => b - a).forEach(row => sheet.deleteRow(row));
    else sheet.appendRow([new Date(), me, commentId, emoji]);
    bumpPulse_('');
    return { commentId, reactions: reactionsByComment_()[commentId] || [] };
  });
}

// ---------- 🔴 Live notifications ----------
// Apps Script can't push to the phone, so the app asks "anything new?" every few seconds with getPulse(), which only reads
// the script cache (no spreadsheet), and loads the full notifications only when the answer changes. Every new comment,
// reaction or note mention bumps the shared counter; whoever is @mentioned also gets their own counter bumped.
function bumpPulse_(text) {
  try {
    const c = CacheService.getScriptCache(), now = String(Date.now()), put = { sbPulse: now };
    String(text || '').replace(/@\{([^}\s]+)\}/g, (m, e) => { put['sbPulse:' + e.toLowerCase()] = now; return m; });
    c.putAll(put, 21600);
  } catch (e) { /* the app's slower check still picks it up */ }
}
/** { all, me }: change counters (just numbers). Cheap on purpose: no spreadsheet reads. */
function getPulse() {
  const me = String(currentUser_() || '').toLowerCase();
  if (!me || me === 'unknown') return { all: '0', me: '0' };
  const got = CacheService.getScriptCache().getAll(['sbPulse', 'sbPulse:' + me]);
  return { all: got.sbPulse || '0', me: got['sbPulse:' + me] || '0' };
}

function addComment(id, key, text) {
  requireRole_('Viewer');   // everyone with access can comment (and delete their own)
  text = String(text || '').trim();
  if (!text) throw new Error('Write a comment first.');
  if (text.length > MAX_COMMENT) throw new Error('Comments can be up to ' + MAX_COMMENT + ' characters.');
  return withLock_(() => {
    const sheet = master_();
    const cols = columns_(sheet);
    const item = readItem_(sheet, cols, locate_(sheet, cols, id, key));
    commentsSheet_().appendRow([new Date(), currentUser_(), asText_(item.name), asText_(item.partNumber), asText_(item.job), asText_(text), asText_(item.key), Utilities.getUuid()]);
    bumpPulse_(text);
    return getComments(item.key);
  });
}

/** Deletes a comment. People can only delete their own. */
function deleteComment(commentId) {
  requireRole_('Viewer');   // everyone with access can comment (and delete their own)
  return withLock_(() => {
    const sheet = commentsSheet_();
    const c = readComments_(sheet).find(x => x.id === commentId);
    if (!c) throw new Error('That comment was already deleted.');
    if (c.user !== currentUser_()) throw new Error('You can only delete your own comments.');
    sheet.deleteRow(c.row);
    const rs = reactionsSheet_();
    readReactions_().filter(x => x.commentId === commentId).map(x => x.row).sort((a, b) => b - a).forEach(row => rs.deleteRow(row));
    return getComments(c.key);
  });
}

// ---------- Documents (PDFs attached to items) ----------

/** PDFs attached to one item, newest first. */
function getDocuments(key) {
  requireRole_('Viewer');
  return docsInOrder_(readDocs_(docsSheet_()).filter(d => d.key === key))
    .map(d => ({ fileId: d.fileId, name: d.name, sizeKB: d.sizeKB, user: d.user, timestamp: d.timestamp, url: d.url }));
}

/**
 * Saves a PDF (data URL from the phone) to the "Inventory Documents" Drive folder and attaches it to the item.
 * The file is shared view-only with people in your Google Workspace domain who have the link, so coworkers can open it.
 */
/** Checks an uploaded PDF (data URL from the phone): { bytes, name }. */
function parsePdf_(fileName, dataUrl) {
  const m = /^data:application\/pdf;base64,(.+)$/.exec(dataUrl || '');
  if (!m) throw new Error('Only PDF files can be attached.');
  const bytes = Utilities.base64Decode(m[1]);
  if (bytes.length > MAX_DOC_MB * 1024 * 1024) throw new Error('That PDF is over ' + MAX_DOC_MB + ' MB.');
  if (String.fromCharCode.apply(null, bytes.slice(0, 4).map(b => b & 0xff)) !== '%PDF') throw new Error('That file isn\'t a valid PDF.');
  let name = String(fileName || 'document.pdf').replace(/[\\/:*?"<>|]+/g, ' ').trim().slice(0, 120) || 'document.pdf';
  if (!/\.pdf$/i.test(name)) name += '.pdf';
  return { bytes, name };
}

/** Saves a PDF in a Drive folder, shared view-only with your Workspace domain (people with the link). */
function savePdf_(folderName, pdf, description) {
  const it = DriveApp.getFoldersByName(folderName);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(folderName);
  const file = folder.createFile(Utilities.newBlob(pdf.bytes, 'application/pdf', pdf.name));
  file.setDescription(description);
  try {
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    // Your Workspace may not allow this; the file is still saved, but only the app's owner can open it.
  }
  return file;
}

function uploadDocument(id, key, fileName, dataUrl) {
  requireRole_('Editor');
  const { bytes, name } = parsePdf_(fileName, dataUrl);
  return withLock_(() => {
    const sheet = master_();
    const cols = columns_(sheet);
    const item = readItem_(sheet, cols, locate_(sheet, cols, id, key));
    const file = savePdf_(DOCS_FOLDER, { bytes, name },
      'Attached to "' + item.name + '" (' + [item.partNumber, item.job].filter(String).join(', ') + ') from Supplybot');
    const mine = readDocs_(docsSheet_()).filter(d => d.key === item.key);
    const arranged = mine.some(d => d.position > 0);   // once arranged, new PDFs go at the end of that order
    docsSheet_().appendRow([new Date(), currentUser_(), asText_(item.name), asText_(item.partNumber), asText_(item.job),
      asText_(name), Math.max(1, Math.round(bytes.length / 1024)), file.getId(), asText_(item.key), file.getUrl(),
      arranged ? Math.max(...mine.map(d => d.position)) + 1 : '']);
    log_('Document added', item, 0, name);
    return getDocuments(item.key);
  });
}

/** Removes a PDF from its item and moves the file to Drive's trash (restorable for 30 days). */
function deleteDocument(fileId) {
  requireRole_('Editor');
  return withLock_(() => {
    const sheet = docsSheet_();
    const d = readDocs_(sheet).find(x => x.fileId === fileId);
    if (!d) throw new Error('That document was already removed.');
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* already gone */ }
    sheet.deleteRow(d.row);
    logSheet_().appendRow([new Date(), currentUser_(), 'Document removed', '', asText_(d.itemName), '', '', 0, '', asText_(d.name)]);
    return getDocuments(d.key);
  });
}

/** After an item's name, part # or job is edited, re-point its documents at the new key. */
function moveDocs_(oldKey, item) {
  const sheet = docsSheet_();
  readDocs_(sheet).filter(d => d.key === oldKey).forEach(d => {
    sheet.getRange(d.row, 3, 1, 3).setValues([[asText_(item.name), asText_(item.partNumber), asText_(item.job)]]);
    sheet.getRange(d.row, 9).setValue(asText_(item.key));
  });
}

// ---------- Orders ----------
// An order is a Draft until "Mark as shipped", which subtracts each Ship Quantity from Available Today.
// Cancelling a shipped order puts the quantities back. Our Cost = Ship Quantity × the item's Cost Price.

function tabWithHeaders_(name, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name, ss.getNumSheets());
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    // A tab created by an earlier version: add any newer columns at the end. Count the header row itself, not the
    // widest data row, so a header is still added if a value already landed in its column.
    const row1 = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
    let have = row1.length;
    while (have && String(row1[have - 1]).trim() === '') have--;
    if (have < headers.length) {
      sheet.getRange(1, have + 1, 1, headers.length - have).setValues([headers.slice(have)]).setFontWeight('bold');
    }
  }
  return sheet;
}
const orderDocsSheet_ = () => tabWithHeaders_(ORDER_DOCS_SHEET, ORDER_DOC_HEADERS);
const orderPhotosSheet_ = () => tabWithHeaders_(ORDER_PHOTOS_SHEET, ORDER_PHOTO_HEADERS);
const orderLinksSheet_ = () => tabWithHeaders_(ORDER_LINKS_SHEET, ORDER_LINK_HEADERS);
const orderCostsSheet_ = () => tabWithHeaders_(ORDER_COSTS_SHEET, ORDER_COST_HEADERS);
const ordersSheet_ = () => tabWithHeaders_(ORDERS_SHEET, ORDER_HEADERS);
const linesSheet_ = () => tabWithHeaders_(LINES_SHEET, LINE_HEADERS);

const fmtDay_ = d => d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(d || '');
const fmtTime_ = d => d instanceof Date ? d.toISOString() : String(d || '');
const roundMoney_ = n => Math.round(n * 100) / 100;

function readOrderRows_() {
  const sheet = ordersSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const H = i => ORDER_HEADERS.indexOf(i);
  return sheet.getRange(2, 1, last - 1, ORDER_HEADERS.length).getValues().map((r, i) => {
    const o = { row: i + 2, orderNo: String(r[0]), status: String(r[1]) };
    ORDER_FIELDS.concat(ORDER_EXTRA_FIELDS).forEach(([k, h]) => o[k] = k === 'orderDate' ? fmtDay_(r[H(h)]) : String(r[H(h)] ?? ''));
    o.shippingCost = money_(r[H('Shipping Cost')]);   // number, or null when not entered
    o.deliveryDate = fmtDay_(r[H('Delivery Date')]);   // 'yyyy-MM-dd', or '' when not set
    o.shipDate = fmtDay_(r[H('Pickup / Ship Date')]);  // expected pickup / ship day, same format
    o.tags = cleanTags_(o.tags, true);
    o.endpoints = cleanTags_(o.endpoints, true);   // summary of the lines' endpoints (getOrders rebuilds it from the lines)
    Object.assign(o, {
      totalQty: Number(r[H('Total Ship Qty')]) || 0, ourCost: Number(r[H('Our Cost')]) || 0,
      created: fmtTime_(r[H('Created')]), createdBy: String(r[H('Created By')]),
      updated: fmtTime_(r[H('Updated')]), updatedBy: String(r[H('Updated By')]),
      shipped: fmtTime_(r[H('Shipped')]), shippedBy: String(r[H('Shipped By')]),
      cancelled: fmtTime_(r[H('Cancelled')]), cancelledBy: String(r[H('Cancelled By')]),
    });
    return o;
  }).filter(o => o.orderNo);
}

function readLineRows_() {
  const sheet = linesSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, LINE_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2, orderNo: String(r[0]), name: String(r[1]), partNumber: String(r[2]), job: String(r[3]),
    qty: Number(r[4]) || 0, cost: r[5] === '' ? null : Number(r[5]), ourCost: r[6] === '' ? null : Number(r[6]), key: String(r[7]),
    ...readEndpoints_(r[8], r[9], r.slice(10, 10 + EP_INFO.length)),
  })).filter(l => l.orderNo);
}

/** Every order (newest first) with its lines. */
function getOrders() {
  requireRole_('Viewer');
  if (myHiddenScreens_().includes('orders')) return [];   // Orders turned off for this person: no orders anywhere in the app
  const lines = readLineRows_();
  const docCounts = {};
  const docsByOrder = {};   // each order's PDFs (name and link), so the app can find an order by a PDF's name
  readOrderDocs_().forEach(d => { docCounts[d.orderNo] = (docCounts[d.orderNo] || 0) + 1; (docsByOrder[d.orderNo] = docsByOrder[d.orderNo] || []).push({ name: d.name, url: d.url }); });
  const linksByOrder = {};
  readOrderLinks_().forEach(k => (linksByOrder[k.orderNo] = linksByOrder[k.orderNo] || []).push({ id: k.id, label: k.label, url: k.url }));
  const costsByOrder = {};
  readOrderCosts_().forEach(c => (costsByOrder[c.orderNo] = costsByOrder[c.orderNo] || []).push({ id: c.id, label: c.label, amount: c.amount, by: c.user, at: c.timestamp }));
  const commentCounts = {};
  readOrderComments_().forEach(c => commentCounts[c.orderNo] = (commentCounts[c.orderNo] || 0) + 1);
  const photoCounts = {};
  const firstPhotos = {};   // each order's first photo (in the arranged order), shown on its card in the Orders list
  readOrderPhotos_().forEach(p => { photoCounts[p.orderNo] = (photoCounts[p.orderNo] || 0) + 1; if (!firstPhotos[p.orderNo]) firstPhotos[p.orderNo] = p.fileId; });
  const returned = returnedByOrder_();
  return readOrderRows_().reverse().map(o => Object.assign(o, {
    documentCount: docCounts[o.orderNo] || 0,
    docs: docsByOrder[o.orderNo] || [],
    commentCount: commentCounts[o.orderNo] || 0,
    photoCount: photoCounts[o.orderNo] || 0,
    firstPhoto: firstPhotos[o.orderNo] || '',
    links: linksByOrder[o.orderNo] || [],
    extraCosts: costsByOrder[o.orderNo] || [],
    extraCost: roundMoney_((costsByOrder[o.orderNo] || []).reduce((n, c) => n + c.amount, 0)),
    lines: lines.filter(l => l.orderNo === o.orderNo).map(l => ({ name: l.name, partNumber: l.partNumber, job: l.job, qty: l.qty, cost: l.cost, ourCost: l.ourCost, key: l.key, endpoints: l.endpoints, serials: l.serials,
      rotary: l.rotary, amps: l.amps, breaker: l.breaker, share: l.share })),
  })).map(o => Object.assign(o, { endpoints: unionTags_(o.lines.map(l => l.endpoints)) }))
    .map(o => o.status === 'Shipped' ? netOfReturns_(o, returned[o.orderNo.trim().toLowerCase()]) : o);
}

// ---------- Returns taken off shipped orders ----------
// A return saved with an order # (Receive → Returning inventory) takes those units off that shipped order everywhere
// in the app: each line's qty and Our Cost become what was kept (shippedQty / shippedOurCost / returned keep what went
// out), and the order's totalQty and ourCost follow. Returned units go to the order's lines with that item, first line first.
/** { order # (lower case): { item key: units returned } } */
function returnedByOrder_() {
  const out = {};
  readAdjustments_().forEach(a => {
    if (a.type !== RECEIPT_TYPES.return || !a.orderNo || !(a.change > 0)) return;
    const m = out[a.orderNo.trim().toLowerCase()] = out[a.orderNo.trim().toLowerCase()] || {};
    m[a.key] = (m[a.key] || 0) + a.change;
  });
  return out;
}
/** The order's lines with each one's returned units, without changing anything. */
function allocateReturns_(lines, byKey) {
  const left = Object.assign({}, byKey || {});
  return lines.map(l => { const r = Math.min(Math.max(0, left[l.key] || 0), Number(l.qty) || 0); left[l.key] = (left[l.key] || 0) - r; return r; });
}
function netOfReturns_(o, byKey) {
  if (!byKey) return o;
  const back = allocateReturns_(o.lines, byKey);
  if (!back.some(Boolean)) return o;
  let rq = 0, rc = 0;
  o.lines.forEach((l, i) => {
    const r = back[i], q = Number(l.qty) || 0;
    l.shippedQty = q; l.returned = r; l.qty = q - r;
    if (l.ourCost != null && l.ourCost !== '') { l.shippedOurCost = Number(l.ourCost); l.ourCost = roundMoney_(q ? l.shippedOurCost * (q - r) / q : 0); rc += l.shippedOurCost - l.ourCost; }
    rq += r;
  });
  const lineSum = o.lines.reduce((n, l) => n + (Number(l.shippedOurCost ?? l.ourCost) || 0), 0);
  o.shippedQty = Number(o.totalQty) || o.lines.reduce((n, l) => n + (l.shippedQty || 0), 0);
  o.shippedOurCost = roundMoney_(Number(o.ourCost) || lineSum);   // the order's Our Cost, or its lines' when that's blank
  o.returnedQty = rq; o.returnedCost = roundMoney_(rc);
  o.totalQty = Math.max(0, o.shippedQty - rq); o.ourCost = roundMoney_(Math.max(0, o.shippedOurCost - rc));
  return o;
}

/** Suggestions for the order form: names already used on past orders, plus the PM list from the master sheet's "PM Request" dropdown. */
function getOrderOptions() {
  requireRole_('Viewer');
  const orders = readOrderRows_();
  const used = k => [...new Set(orders.map(o => o[k]).filter(Boolean))].sort();
  const pms = new Set(used('pm'));
  try {
    const sheet = master_();
    const col = findCol_(headerRow_(sheet), 'PM Request');
    const rule = col === -1 ? null : sheet.getRange(2, col + 1).getDataValidation();
    if (rule && rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) rule.getCriteriaValues()[0].forEach(v => pms.add(String(v)));
  } catch (e) { /* no dropdown: past orders only */ }
  const out = {};
  ORDER_FIELDS.forEach(([k]) => out[k] = used(k));
  out.pm = [...pms].filter(Boolean).sort();
  return out;
}

function nextOrderNo_(orders) {
  const year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy');
  const max = orders.reduce((m, o) => {
    const x = new RegExp('^ORD-' + year + '-(\\d+)$').exec(o.orderNo);
    return x ? Math.max(m, Number(x[1])) : m;
  }, 0);
  return 'ORD-' + year + '-' + String(max + 1).padStart(4, '0');
}

/** Item lookup by key from a fresh read of the master sheet. */
function itemsByKey_() {
  const map = {};
  getItems().forEach(i => { if (!map[i.key]) map[i.key] = i; });
  return map;
}

/**
 * Creates or updates a Draft order. order = { orderNo?, orderDate, projectName, ..., lines: [{ key, qty }] }.
 * Cost Price is taken from the item now, and Our Cost = qty × cost.
 */
function saveOrder(order) {
  requireRole_('Editor');
  if (!order) throw new Error('Nothing to save.');
  const lines = (order.lines || []).map(l => Object.assign({ key: String(l.key), qty: Math.floor(Number(l.qty)) }, pairEndpoints_(l.endpoints, l.serials, false, l)));
  if (!lines.length) throw new Error('Add at least one item to the order.');
  if (lines.some(l => !(l.qty > 0))) throw new Error('Every item needs a Ship Quantity of 1 or more.');
  if (!String(order.projectName || '').trim() && !String(order.projectNo || '').trim()) throw new Error('Enter a Project Name or Project #.');
  return withLock_(() => {
    const items = itemsByKey_();
    const missing = lines.filter(l => !items[l.key]);
    if (missing.length) throw new Error(missing.length + ' item(s) on this order are no longer in the sheet. Refresh and try again.');
    const orders = readOrderRows_();
    const existing = order.orderNo ? orders.find(o => o.orderNo === order.orderNo) : null;
    if (order.orderNo && !existing) throw new Error('Order ' + order.orderNo + ' was not found.');
    if (existing) requireOrderOwner_(existing.orderNo);
    if (existing && existing.status !== 'Draft') throw new Error('Only Draft orders can be edited (' + existing.orderNo + ' is ' + existing.status + ').');
    // Endpoints go on EVSE items only. Ones already on another item (from before this rule) may stay, but none can be added.
    const before = existing ? readLineRows_().filter(l => l.orderNo === existing.orderNo) : [];
    const had = l => unionTags_(before.filter(b => b.key === l.key).map(b => b.endpoints));
    lines.forEach(l => checkEndpointsAllowed_(items[l.key], l.endpoints, had(l)));
    if (existing && !orderOwnerOf_(existing)) {   // no Created By: its endpoints stay exactly as they are
      const eps = ls => JSON.stringify(ls.flatMap(l => (l.endpoints || []).map((e, j) => l.key + '|' + String(e).toLowerCase() + '|' + String((l.serials || [])[j] || ''))).sort());
      if (eps(lines) !== eps(before)) throw new Error(existing.orderNo + ' has no Created By, so its endpoints can’t be changed.');
    }
    if (lines.some(l => l.endpoints.length))
      checkEndpointsUnique_(lines.map(l => ({ name: items[l.key].name, endpoints: l.endpoints, had: had(l) })), endpointUses_(existing ? existing.orderNo : ''));
    checkShares_(lines);
    const orderNo = existing ? existing.orderNo : nextOrderNo_(orders);
    const me = currentUser_(), now = new Date();

    // Lines: replace this order's lines with the new set.
    const ls = linesSheet_();
    readLineRows_().filter(l => l.orderNo === orderNo).map(l => l.row).sort((a, b) => b - a).forEach(row => ls.deleteRow(row));
    let totalQty = 0, ourCost = 0;
    lines.forEach(l => {
      const it = items[l.key];
      const lineCost = it.cost == null ? '' : roundMoney_(it.cost * l.qty);
      totalQty += l.qty; ourCost += lineCost || 0;
      ls.appendRow([orderNo, asText_(it.name), asText_(it.partNumber), asText_(it.job), l.qty, it.cost == null ? '' : it.cost, lineCost, asText_(it.key)].concat(lineEpCells_(l)));
    });

    // Order row
    const H = h => ORDER_HEADERS.indexOf(h);
    const row = existing ? ordersSheet_().getRange(existing.row, 1, 1, ORDER_HEADERS.length).getValues()[0] : new Array(ORDER_HEADERS.length).fill('');
    row[H('Order #')] = orderNo;
    row[H('Status')] = 'Draft';
    row[H('Tracking Link')] = cleanUrl_(order.trackingUrl);
    row[H('Shipping Cost')] = shippingCost_(order.shippingCost);
    row[H('Ship To Address')] = asText_(cleanAddress_(order.shipTo));
    row[H('Delivery Date')] = deliveryDate_(order.deliveryDate);
    row[H('Pickup / Ship Date')] = deliveryDate_(order.shipDate, 'Pickup / Ship Date');
    row[H('Tags')] = asText_(cleanTags_(order.tags).join(', '));
    row[H('Endpoints')] = asText_(unionTags_(lines.map(l => l.endpoints)).join(', '));
    ORDER_FIELDS.forEach(([k, h]) => {
      const v = String(order[k] ?? '').trim();
      if (k === 'orderDate') {
        const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v);
        row[H(h)] = m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : asText_(v);
      } else row[H(h)] = asText_(v);
    });
    row[H('Total Ship Qty')] = totalQty;
    row[H('Our Cost')] = roundMoney_(ourCost);
    row[H('Lines')] = lines.length;
    if (!existing) { row[H('Created')] = now; row[H('Created By')] = me; }
    row[H('Updated')] = now; row[H('Updated By')] = me;
    const os = ordersSheet_();
    if (existing) os.getRange(existing.row, 1, 1, ORDER_HEADERS.length).setValues([row]);
    else os.appendRow(row);
    os.getRange(existing ? existing.row : os.getLastRow(), ORDER_HEADERS.indexOf('Shipping Cost') + 1).setNumberFormat('$#,##0.00');
    os.getRange(existing ? existing.row : os.getLastRow(), ORDER_HEADERS.indexOf('Delivery Date') + 1).setNumberFormat(DELIVERY_FORMAT);
    os.getRange(existing ? existing.row : os.getLastRow(), ORDER_HEADERS.indexOf('Pickup / Ship Date') + 1).setNumberFormat(DELIVERY_FORMAT);
    return getOrders().find(o => o.orderNo === orderNo);
  });
}

/** Ships a Draft: checks every item has enough Available Today, then subtracts all of them and logs each change. */
function shipOrder(orderNo) {
  requireRole_('Editor');
  requireOrderOwner_(orderNo);
  return withLock_(() => {
    const order = readOrderRows_().find(o => o.orderNo === orderNo);
    if (!order) throw new Error('Order ' + orderNo + ' was not found.');
    if (order.status !== 'Draft') throw new Error(orderNo + ' is already ' + order.status + '.');
    const lines = readLineRows_().filter(l => l.orderNo === orderNo);
    const sheet = master_(), cols = columns_(sheet);
    const items = itemsByKey_();
    // Same item on several lines counts once with the total quantity.
    const need = {};
    lines.forEach(l => need[l.key] = (need[l.key] || 0) + l.qty);
    const problems = [];
    Object.keys(need).forEach(k => {
      const it = items[k];
      if (!it) problems.push('an item is no longer in the sheet');
      else if (it.quantity < need[k]) problems.push(it.name + ': ships ' + need[k] + ', only ' + it.quantity + ' available');
    });
    if (problems.length) throw new Error('Not enough stock to ship:\n' + problems.join('\n'));
    Object.keys(need).forEach(k => {
      const it = items[k];
      const cell = sheet.getRange(Number(it.id), cols.quantity + 1);
      cell.setValue((Number(cell.getValue()) || 0) - need[k]);
      log_('Order shipped', readItem_(sheet, cols, Number(it.id)), -need[k], orderNo);
    });
    setOrderStatus_(order, 'Shipped');
    return getOrders().find(o => o.orderNo === orderNo);
  });
}

/** Cancels an order. If it had shipped, the quantities go back into Available Today. */
function cancelOrder(orderNo) {
  requireRole_('Editor');
  requireOrderOwner_(orderNo);
  return withLock_(() => {
    const order = readOrderRows_().find(o => o.orderNo === orderNo);
    if (!order) throw new Error('Order ' + orderNo + ' was not found.');
    if (order.status === 'Cancelled') throw new Error(orderNo + ' is already cancelled.');
    if (order.status === 'Shipped') {
      const lines = readLineRows_().filter(l => l.orderNo === orderNo);
      const sheet = master_(), cols = columns_(sheet);
      const items = itemsByKey_();
      const missing = lines.filter(l => !items[l.key]);
      if (missing.length) throw new Error("Can't return stock: " + missing.map(l => l.name).join(', ') + ' is no longer in the sheet.');
      const back = allocateReturns_(lines, returnedByOrder_()[orderNo.trim().toLowerCase()]);   // already returned: not added twice
      lines.forEach((l, i) => {
        const it = items[l.key], qty = l.qty - back[i];
        if (!(qty > 0)) return;
        const cell = sheet.getRange(Number(it.id), cols.quantity + 1);
        cell.setValue((Number(cell.getValue()) || 0) + qty);
        log_('Order cancelled (stock returned)', readItem_(sheet, cols, Number(it.id)), qty, orderNo);
      });
    }
    setOrderStatus_(order, 'Cancelled');
    return getOrders().find(o => o.orderNo === orderNo);
  });
}

/** Links must be web addresses; "ups.com/track?…" gets https:// added. Anything else (e.g. javascript:) is refused. */
/** Shipping cost as a number ("$1,250.50" -> 1250.5), or '' when left blank. Negative or non-numbers are refused. */
function shippingCost_(value) {
  const v = String(value ?? '').trim();
  if (!v) return '';
  const n = Number(v.replace(/[$,\s]/g, ''));
  if (!isFinite(n) || n < 0) throw new Error('Shipping Cost must be an amount like 250 or $1,250.50.');
  return roundMoney_(n);
}

/** Sets or clears an order's Shipping Cost. Allowed in any status, since the freight bill often comes after shipping. */
function setOrderShipping(orderNo, value) {
  requireRole_('Editor');
  requireOrderOwner_(orderNo);
  const cost = shippingCost_(value);
  return withLock_(() => {
    const order = readOrderRows_().find(o => o.orderNo === orderNo);
    if (!order) throw new Error('Order ' + orderNo + ' was not found.');
    const sheet = ordersSheet_();
    const H = h => ORDER_HEADERS.indexOf(h) + 1;
    sheet.getRange(order.row, H('Shipping Cost')).setNumberFormat('$#,##0.00').setValue(cost);
    sheet.getRange(order.row, H('Updated'), 1, 2).setValues([[new Date(), currentUser_()]]);
    return getOrders().find(o => o.orderNo === orderNo);
  });
}

/** Trims an address, keeps its line breaks (max 4 lines) and caps its length. */
function cleanAddress_(value) {
  const v = String(value || '').replace(/\r/g, '').split('\n').map(l => l.trim().replace(/\s+/g, ' ')).filter(Boolean).slice(0, 4).join('\n');
  if (v.length > MAX_ADDRESS) throw new Error('That address is too long (' + MAX_ADDRESS + ' characters max).');
  return v;
}

/** Sets or clears an order's Ship To Address. Allowed in any status, like the shipping cost. */
/** A delivery date from the app ('yyyy-mm-dd') as a sheet date, or '' when left blank. Anything else is refused. */
const DELIVERY_FORMAT = 'mmm d, yyyy';
function deliveryDate_(value, label) {
  const v = String(value ?? '').trim();
  if (!v) return '';
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v);
  if (!m) throw new Error((label || 'Delivery Date') + ' must be a date.');
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

/** Sets or clears an order's Delivery Date. Allowed in any status, since the delivery often happens after shipping. */
function setOrderDeliveryDate(orderNo, value) {
  return setOrderDate_(orderNo, 'Delivery Date', value);
}

/** Sets or clears an order's expected Pickup / Ship Date. Allowed in any status, like the delivery date. */
function setOrderShipDate(orderNo, value) {
  return setOrderDate_(orderNo, 'Pickup / Ship Date', value);
}

function setOrderDate_(orderNo, header, value) {
  requireRole_('Editor');
  requireOrderOwner_(orderNo);
  const date = deliveryDate_(value, header);
  return withLock_(() => {
    const order = readOrderRows_().find(o => o.orderNo === orderNo);
    if (!order) throw new Error('Order ' + orderNo + ' was not found.');
    const sheet = ordersSheet_();
    const H = h => ORDER_HEADERS.indexOf(h) + 1;
    sheet.getRange(order.row, H(header)).setNumberFormat(DELIVERY_FORMAT).setValue(date);
    sheet.getRange(order.row, H('Updated'), 1, 2).setValues([[new Date(), currentUser_()]]);
    return getOrders().find(o => o.orderNo === orderNo);
  });
}

function setOrderAddress(orderNo, address) {
  requireRole_('Editor');
  requireOrderOwner_(orderNo);
  const v = cleanAddress_(address);
  return withLock_(() => {
    const order = readOrderRows_().find(o => o.orderNo === orderNo);
    if (!order) throw new Error('Order ' + orderNo + ' was not found.');
    const sheet = ordersSheet_();
    const H = h => ORDER_HEADERS.indexOf(h) + 1;
    sheet.getRange(order.row, H('Ship To Address')).setValue(asText_(v));
    sheet.getRange(order.row, H('Updated'), 1, 2).setValues([[new Date(), currentUser_()]]);
    return getOrders().find(o => o.orderNo === orderNo);
  });
}

/**
 * Tags as a list: from an array or "a, b" text. Trims spaces, drops duplicates (ignoring capitals) and blanks.
 * lenient: when reading the sheet, keep whatever is there instead of throwing.
 */
function cleanTags_(value, lenient, what) {
  what = what || 'tags';
  const list = Array.isArray(value) ? value : String(value || '').split(',');
  const seen = {}, out = [];
  list.forEach(t => {
    const v = String(t ?? '').replace(/,/g, ' ').replace(/\s+/g, ' ').trim();
    if (!v || seen[v.toLowerCase()]) return;
    if (!lenient && v.length > MAX_TAG_LENGTH) throw new Error(what.charAt(0).toUpperCase() + what.slice(1) + ' can be up to ' + MAX_TAG_LENGTH + ' characters ("' + v + '" is too long).');
    seen[v.toLowerCase()] = true;
    out.push(v);
  });
  if (!lenient && out.length > MAX_TAGS) throw new Error('Up to ' + MAX_TAGS + ' ' + what + ' each.');
  return out;
}

/** Tag field name in the app -> column header. (Endpoints are set per order line instead: setLineEndpoints.) */
const TAG_COLUMNS = { tags: 'Tags' };
function tagColumn_(field) {
  const h = TAG_COLUMNS[field || 'tags'];
  if (!h) throw new Error('Unknown tag type: ' + field);
  return h;
}

/**
 * A line's Endpoints, Endpoint Serials (and, for order lines, the EP_INFO detail cells) -> { endpoints: [...],
 * serials: [...], rotary: [...], ... } all the same length ('' = not entered).
 */
function readEndpoints_(epCell, snCell, infoCells) {
  const names = String(epCell || '').split(','), sns = String(snCell || '').split(',');
  const info = {};
  EP_INFO.forEach(([k], n) => info[k] = String((infoCells || [])[n] ?? '').split(','));
  return pairEndpoints_(names, sns, true, info);
}

/**
 * Pairs endpoints with their serials, dropping blanks and repeats (ignoring capitals) and keeping them lined up.
 * lenient: when reading the sheet, keep whatever is there instead of throwing.
 */
function pairEndpoints_(names, serials, lenient, info) {
  names = Array.isArray(names) ? names : []; serials = Array.isArray(serials) ? serials : []; info = info || {};
  const clean = v => String(v ?? '').replace(/,/g, ' ').replace(/\s+/g, ' ').trim();
  const seen = {}, out = { endpoints: [], serials: [] };
  EP_INFO.forEach(([k]) => out[k] = []);
  names.forEach((n, i) => {
    const name = clean(n), sn = clean(serials[i]);
    if (!name || seen[name.toLowerCase()]) return;
    if (!lenient && name.length > MAX_TAG_LENGTH) throw new Error('Endpoints can be up to ' + MAX_TAG_LENGTH + ' characters ("' + name + '" is too long).');
    if (!lenient && sn.length > MAX_SERIAL) throw new Error('Serial numbers can be up to ' + MAX_SERIAL + ' characters.');
    seen[name.toLowerCase()] = true;
    out.endpoints.push(name); out.serials.push(sn);
    EP_INFO.forEach(([k, , max, label]) => {   // all optional
      const v = clean((Array.isArray(info[k]) ? info[k] : [])[i]);
      if (!lenient && v.length > max) throw new Error(label + ' can be up to ' + max + ' characters (' + name + ').');
      out[k].push(v);
    });
  });
  if (!lenient && out.endpoints.length > MAX_TAGS) throw new Error('Up to ' + MAX_TAGS + ' endpoints each.');
  return out;
}

/** The Endpoints and Endpoint Serials cells for a line (serials left blank when none are entered). */
const endpointCells_ = p => [asText_(p.endpoints.join(', ')), p.serials.some(Boolean) ? asText_(p.serials.join(', ')) : ''];
/** An order line's endpoint cells: Endpoints, Endpoint Serials, then the EP_INFO details (each left blank when none are entered). */
const lineEpCells_ = p => endpointCells_(p).concat(EP_INFO.map(([k]) => (p[k] || []).some(Boolean) ? asText_(p[k].join(', ')) : ''));

/**
 * Load sharing on an order: each endpoint's partner must be another endpoint on the same order, and each endpoint can be
 * chosen only once (the app sets both sides of a pair, A ⇄ B, which passes this check).
 */
function checkShares_(lines) {
  const names = {}, chosen = {};
  lines.forEach(l => l.endpoints.forEach(e => names[e.toLowerCase()] = e));
  lines.forEach(l => l.endpoints.forEach((e, j) => {
    const s = (l.share || [])[j];
    if (!s) return;
    const k = s.toLowerCase();
    if (!names[k]) throw new Error('Load sharing for ' + e + ' must be another endpoint on this order ("' + s + '" isn’t).');
    if (k === e.toLowerCase()) throw new Error(e + ' can’t load-share with itself.');
    if (chosen[k]) throw new Error(names[k] + ' is already chosen for load sharing with ' + chosen[k] + '.');
    chosen[k] = e;
  }));
}

/**
 * Saves the endpoints and their details on one or more item lines of an order in one step (load sharing pairs can
 * span two lines). changes: [{ index, key, endpoints, serials, rotary, amps, breaker, share }], index = the line's place
 * in the order's item list, key = that line's item (to make sure the order hasn't changed meanwhile). Any status.
 */
function setOrderEndpointDetails(orderNo, changes) {
  requireRole_('Editor');
  requireOrderOwner_(orderNo);
  requireEndpointOwner_(orderNo);
  changes = (changes || []).map(c => Object.assign({ index: Number(c.index), key: String(c.key) }, pairEndpoints_(c.endpoints, c.serials, false, c)));
  if (!changes.length) throw new Error('Nothing to save.');
  return withLock_(() => {
    const order = readOrderRows_().find(o => o.orderNo === orderNo);
    if (!order) throw new Error('Order ' + orderNo + ' was not found.');
    const lines = readLineRows_().filter(l => l.orderNo === orderNo), items = itemsByKey_();
    const after = lines.map(l => Object.assign({}, l));
    changes.forEach(c => {
      const line = lines[c.index];
      if (!line || line.key !== c.key) throw new Error('The items on ' + orderNo + ' have changed. Refresh and try again.');
      checkEndpointsAllowed_(items[line.key], c.endpoints, line.endpoints);
      Object.assign(after[c.index], c);
    });
    checkEndpointsUnique_(after.map((l, n) => ({ name: l.name, endpoints: l.endpoints, had: lines[n].endpoints })), endpointUses_(orderNo));
    checkShares_(after);
    const col = LINE_HEADERS.indexOf('Endpoints') + 1, width = 2 + EP_INFO.length;
    changes.forEach(c => linesSheet_().getRange(lines[c.index].row, col, 1, width).setValues([lineEpCells_(c)]));
    const sheet = ordersSheet_();
    const H = h => ORDER_HEADERS.indexOf(h) + 1;
    sheet.getRange(order.row, H('Endpoints')).setValue(asText_(unionTags_(after.map(l => l.endpoints)).join(', ')));
    sheet.getRange(order.row, H('Updated'), 1, 2).setValues([[new Date(), currentUser_()]]);
    return getOrders().find(o => o.orderNo === orderNo);
  });
}

// ---------- Who may change an order ----------
// Only the person who created an order can change it (its items, endpoints, dates, costs, links, photos, PDFs, tags,
// address, shipping, ship / cancel). Anyone can still read it, comment on it, make its PDF and log it in the tracker.
// Orders with no Created By (from before it was recorded) can be changed by any editor.
const orderOwnerOf_ = o => String(o && o.createdBy || '').trim().toLowerCase();
// Orders with no Created By can't have their endpoints changed at all (no one owns them).
function requireEndpointOwner_(orderNo) {
  const o = readOrderRows_().find(x => x.orderNo === String(orderNo || ''));
  if (o && !orderOwnerOf_(o)) throw new Error(o.orderNo + ' has no Created By, so its endpoints can’t be changed.');
}
function requireOrderOwner_(orderNo) {
  const o = readOrderRows_().find(x => x.orderNo === String(orderNo || ''));
  if (!o || !orderOwnerOf_(o) || orderOwnerOf_(o) === currentUser_().toLowerCase()) return;
  throw new Error('Only ' + o.createdBy + ', who created ' + o.orderNo + ', can change it.');
}

const isEvse_ = item => !!item && String(item.category || '').trim().replace(/\s+/g, ' ').toLowerCase() === 'evse';

/** Endpoints can only be added to EVSE items. had: what the line already had (those may stay, or be removed). */
function checkEndpointsAllowed_(item, endpoints, had) {
  if (isEvse_(item)) return;
  const old = (had || []).map(t => t.toLowerCase());
  const added = (endpoints || []).filter(t => !old.includes(t.toLowerCase()));
  if (added.length) throw new Error('Endpoints can only be added to EVSE items' + (item ? ' ("' + item.name + '" is ' + (item.category || 'not in a category') + ').' : '.'));
}

/**
 * Where each endpoint is in use: lower-case name -> e.g. "ORD-2026-0003 (Autel AC Single 50A)" or "PO 4512 (…)".
 * Endpoints on cancelled orders are free again. skipOrder: an order whose lines to leave out (it's being re-saved);
 * skipRow: one Order Lines row to leave out (it's being changed).
 */
function endpointUses_(skipOrder, skipRow) {
  const status = {}, used = {};
  readOrderRows_().forEach(o => status[o.orderNo] = o.status);
  const add = (e, where) => { if (!used[e.toLowerCase()]) used[e.toLowerCase()] = where; };
  readLineRows_().forEach(l => {
    if (status[l.orderNo] === 'Cancelled' || (skipOrder && l.orderNo === skipOrder) || l.row === skipRow) return;
    l.endpoints.forEach(e => add(e, l.orderNo + ' (' + l.name + ')'));
  });
  readAdjustments_().forEach(a => a.endpoints.forEach(e =>
    add(e, (a.po ? 'PO ' + a.po : a.type === 'Returned' ? 'a return' : 'a receipt') + ' (' + a.name + ')')));
  return used;
}

/**
 * Endpoints are unique: each one can be on only one item line or receipt (ignoring capitals).
 * lines: [{ name, endpoints, had }]; had = what that line already had, which may stay as it is.
 */
function checkEndpointsUnique_(lines, used) {
  const seen = {};
  lines.forEach(l => {
    const had = (l.had || []).map(t => t.toLowerCase());
    l.endpoints.forEach(e => {
      const k = e.toLowerCase(), kept = had.includes(k);
      if (seen[k] && !kept) throw new Error('Endpoint "' + e + '" is on two items here (' + seen[k] + ' and ' + l.name + '). Each endpoint can only be used once.');
      seen[k] = l.name;
      if (used[k] && !kept) throw new Error('Endpoint "' + e + '" is already used on ' + used[k] + '. Each endpoint can only be used once.');
    });
  });
}

/** Several tag lists as one, without repeats (ignoring capitals); the first spelling wins. */
function unionTags_(lists) {
  const seen = {}, out = [];
  lists.forEach(list => (list || []).forEach(t => { if (!seen[t.toLowerCase()]) { seen[t.toLowerCase()] = true; out.push(t); } }));
  return out;
}

/**
 * Replaces the endpoints (and their serial #s, same order) on one item line of an order (index = its place in the order's item list; key = that line's
 * item, to make sure the order hasn't changed meanwhile). Allowed in any status, like tags.
 */
function setLineEndpoints(orderNo, index, key, endpoints, serials) {
  requireRole_('Editor');
  requireOrderOwner_(orderNo);
  requireEndpointOwner_(orderNo);
  pairEndpoints_(endpoints, serials);   // checks lengths before anything is read
  return withLock_(() => {
    const order = readOrderRows_().find(o => o.orderNo === orderNo);
    if (!order) throw new Error('Order ' + orderNo + ' was not found.');
    const lines = readLineRows_().filter(l => l.orderNo === orderNo);
    const line = lines[index];
    if (!line || line.key !== String(key)) throw new Error('The items on ' + orderNo + ' have changed. Refresh and try again.');
    // Keep each endpoint's details (rotary dial, amp draw, ...) by name; only names and serials come from this call.
    const at = {}, info = {};
    line.endpoints.forEach((e, j) => at[e.toLowerCase()] = j);
    EP_INFO.forEach(([k]) => info[k] = (endpoints || []).map(e => { const j = at[String(e).trim().toLowerCase()]; return j == null ? '' : line[k][j]; }));
    const pair = pairEndpoints_(endpoints, serials, false, info);
    checkEndpointsAllowed_(itemsByKey_()[line.key], pair.endpoints, line.endpoints);
    checkEndpointsUnique_([{ name: line.name, endpoints: pair.endpoints, had: line.endpoints }], endpointUses_('', line.row));
    linesSheet_().getRange(line.row, LINE_HEADERS.indexOf('Endpoints') + 1, 1, 2 + EP_INFO.length).setValues([lineEpCells_(pair)]);
    line.endpoints = pair.endpoints;
    const sheet = ordersSheet_();
    const H = h => ORDER_HEADERS.indexOf(h) + 1;
    sheet.getRange(order.row, H('Endpoints')).setValue(asText_(unionTags_(lines.map(l => l.endpoints)).join(', ')));
    sheet.getRange(order.row, H('Updated'), 1, 2).setValues([[new Date(), currentUser_()]]);
    return getOrders().find(o => o.orderNo === orderNo);
  });
}

/** Replaces an order's tags. Allowed in any status, like the address. */
function setOrderTags(orderNo, tags, field) {
  requireRole_('Editor');
  requireOrderOwner_(orderNo);
  const header = tagColumn_(field);
  const list = cleanTags_(tags, false, header.toLowerCase());
  return withLock_(() => {
    const order = readOrderRows_().find(o => o.orderNo === orderNo);
    if (!order) throw new Error('Order ' + orderNo + ' was not found.');
    const sheet = ordersSheet_();
    const H = h => ORDER_HEADERS.indexOf(h) + 1;
    sheet.getRange(order.row, H(header)).setValue(asText_(list.join(', ')));
    sheet.getRange(order.row, H('Updated'), 1, 2).setValues([[new Date(), currentUser_()]]);
    return getOrders().find(o => o.orderNo === orderNo);
  });
}

// ----- Tags on items -----
// Kept on their own "Item Tags" tab (one row per item, matched by Item Key) so the master sheet's columns stay as they are.
const ITEM_TAGS_SHEET = 'Item Tags';
const ITEM_TAG_HEADERS = ['Item Key', 'Item Name', 'Part Number', 'Assigned Job', 'Tags', 'Updated', 'Updated By'];
const itemTagsSheet_ = () => tabWithHeaders_(ITEM_TAGS_SHEET, ITEM_TAG_HEADERS);

function readItemTags_() {
  const sheet = itemTagsSheet_(), last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, ITEM_TAG_HEADERS.length).getValues()
    .map((r, i) => ({ row: i + 2, key: String(r[0]), tags: cleanTags_(r[4], true) })).filter(t => t.key);
}

/** Replaces an item's tags. Returns the saved list. */
function setItemTags(id, key, tags, field) {
  requireRole_('Editor');
  const header = tagColumn_(field);
  const list = cleanTags_(tags, false, header.toLowerCase());
  return withLock_(() => {
    const master = master_(), cols = columns_(master);
    const item = readItem_(master, cols, locate_(master, cols, id, key));
    const sheet = itemTagsSheet_();
    const rows = readItemTags_().filter(t => t.key === item.key).map(t => t.row);
    const C = h => ITEM_TAG_HEADERS.indexOf(h);
    const values = rows.length ? sheet.getRange(rows[0], 1, 1, ITEM_TAG_HEADERS.length).getValues()[0] : new Array(ITEM_TAG_HEADERS.length).fill('');
    [item.key, item.name, item.partNumber, item.job].forEach((v, i) => values[i] = asText_(v));
    values[C(header)] = asText_(list.join(', '));
    values[C('Updated')] = new Date(); values[C('Updated By')] = currentUser_();
    if (rows.length) sheet.getRange(rows[0], 1, 1, values.length).setValues([values]);
    else if (list.length) sheet.appendRow(values);
    rows.slice(1).sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));   // stray duplicate rows for the same item
    return list;
  });
}

/** Keeps an item's tags with it when it's renamed in the app. */
// ---------- Minimum stock ----------
// Per item: whether to watch its stock and the minimum to keep. Kept in its own tab (one row per item, by key), so the
// master sheet's columns don't change. Supplybot warns about green-labelled items being watched that are at or below it.
const MIN_STOCK_SHEET = 'Min Stock';
const MIN_STOCK_HEADERS = ['Item Key', 'Item Name', 'Part Number', 'Assigned Job', 'Track', 'Minimum', 'Updated', 'Updated By'];
const minStockSheet_ = () => tabWithHeaders_(MIN_STOCK_SHEET, MIN_STOCK_HEADERS);
function readMinStock_() {
  const sheet = minStockSheet_(), last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, MIN_STOCK_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2, key: String(r[0]), track: r[4] === true || /^(true|yes|y|1|x)$/i.test(String(r[4]).trim()),
    min: r[5] === '' || r[5] == null || isNaN(Number(r[5])) ? null : Number(r[5]) })).filter(m => m.key);
}
/** Turns watching an item's minimum stock on or off and sets the minimum (0–100000). Returns { minTrack, minQty }. */
function setMinStock(id, key, track, min) {
  requireRole_('Editor');
  track = !!track;
  const qty = min === '' || min == null ? null : Number(min);
  if (qty != null && (!(qty >= 0) || Math.floor(qty) !== qty || qty > 100000)) throw new Error('The minimum must be a whole number of 0 or more.');
  if (track && qty == null) throw new Error('Enter the minimum to keep in stock.');
  return withLock_(() => {
    const ms = master_(), cols = columns_(ms), item = readItem_(ms, cols, locate_(ms, cols, id, key));
    const sheet = minStockSheet_(), rows = readMinStock_().filter(m => m.key === item.key);
    const values = [asText_(item.key), asText_(item.name), asText_(item.partNumber), asText_(item.job), track, qty == null ? '' : qty, new Date(), currentUser_()];
    if (rows.length) sheet.getRange(rows[0].row, 1, 1, values.length).setValues([values]);
    else sheet.appendRow(values);
    return { minTrack: track, minQty: qty };
  });
}
function moveMinStock_(oldKey, item) {
  const sheet = minStockSheet_();
  readMinStock_().filter(m => m.key === oldKey).forEach(m =>
    sheet.getRange(m.row, 1, 1, 4).setValues([[asText_(item.key), asText_(item.name), asText_(item.partNumber), asText_(item.job)]]));
}

function moveItemTags_(oldKey, item) {
  const sheet = itemTagsSheet_();
  readItemTags_().filter(t => t.key === oldKey).forEach(t =>
    sheet.getRange(t.row, 1, 1, 4).setValues([[asText_(item.key), asText_(item.name), asText_(item.partNumber), asText_(item.job)]]));
}

function cleanUrl_(value) {
  let v = String(value || '').trim();
  if (!v) return '';
  if (!/^[a-z][a-z0-9+.-]*:/i.test(v)) v = 'https://' + v;
  if (!/^https?:\/\/[^\s]+$/i.test(v)) throw new Error('Links must be web addresses (https://…).');
  return v;
}

/** Sets or clears an order's Tracking Link. Allowed in any status, since tracking usually arrives after shipping. */
// ----- Links on an order (any number; each with an optional label) -----
// The older single "Tracking Link" column still works and is shown in the app as the first link, labelled "Tracking".

function readOrderLinks_() {
  const sheet = orderLinksSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const linkOrder = k => k.position > 0 ? k.position : 1e6 + k.row;
  return sheet.getRange(2, 1, last - 1, ORDER_LINK_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2, orderNo: String(r[2]), label: String(r[3]), url: String(r[4]), id: String(r[5]), position: Number(r[6]) || 0,
  })).filter(k => k.id && k.url).sort((a, b) => linkOrder(a) - linkOrder(b));
}

/**
 * Saves the order an order's links are shown in. ids = every link id in the new order; 'tracking' stands for a link
 * still in the older Tracking Link column, which is moved into the Order Links tab so it can be arranged like the rest.
 */
function reorderOrderLinks(orderNo, ids) {
  requireRole_('Editor');
  requireOrderOwner_(orderNo);
  ids = (ids || []).map(String);
  return withLock_(() => {
    const order = readOrderRows_().find(o => o.orderNo === orderNo);
    if (!order) throw new Error('Order ' + orderNo + ' was not found.');
    const sheet = orderLinksSheet_();
    if (ids.includes('tracking')) {
      if (!order.trackingUrl) throw new Error('The links changed. Refresh and try again.');
      const id = Utilities.getUuid();
      sheet.appendRow([new Date(), currentUser_(), orderNo, 'Tracking', order.trackingUrl, id, '']);
      ordersSheet_().getRange(order.row, ORDER_HEADERS.indexOf('Tracking Link') + 1).setValue('');
      ids = ids.map(x => x === 'tracking' ? id : x);
    }
    const mine = readOrderLinks_().filter(k => k.orderNo === orderNo);
    if (mine.length !== ids.length || !mine.every(k => ids.includes(k.id))) throw new Error('The links changed. Refresh and try again.');
    mine.forEach(k => sheet.getRange(k.row, ORDER_LINK_HEADERS.indexOf('Position') + 1).setValue(ids.indexOf(k.id) + 1));
    return getOrders().find(o => o.orderNo === orderNo);
  });
}

const cleanLabel_ = v => String(v || '').trim().slice(0, 80);

/** Adds a link to an order (any status). Returns the updated order. */
function addOrderLink(orderNo, label, url) {
  requireRole_('Editor');
  requireOrderOwner_(orderNo);
  const clean = cleanUrl_(url);
  if (!clean) throw new Error('Enter a web address.');
  return withLock_(() => {
    if (!readOrderRows_().some(o => o.orderNo === orderNo)) throw new Error('Order ' + orderNo + ' was not found.');
    const mine = readOrderLinks_().filter(k => k.orderNo === orderNo);
    if (mine.length >= MAX_ORDER_LINKS) throw new Error('An order can have up to ' + MAX_ORDER_LINKS + ' links.');
    const arranged = mine.some(k => k.position > 0);   // once arranged, new links go at the end of that order
    orderLinksSheet_().appendRow([new Date(), currentUser_(), orderNo, asText_(cleanLabel_(label)), clean, Utilities.getUuid(), arranged ? Math.max(...mine.map(k => k.position)) + 1 : '']);
    return getOrders().find(o => o.orderNo === orderNo);
  });
}

/** Changes a link's label and/or address. */
function updateOrderLink(linkId, label, url) {
  requireRole_('Editor');
  const clean = cleanUrl_(url);
  if (!clean) throw new Error('Enter a web address.');
  return withLock_(() => {
    const k = readOrderLinks_().find(x => x.id === linkId);
    if (!k) throw new Error('That link was removed. Refresh and try again.');
    requireOrderOwner_(k.orderNo);
    orderLinksSheet_().getRange(k.row, 4, 1, 2).setValues([[asText_(cleanLabel_(label)), clean]]);
    return getOrders().find(o => o.orderNo === k.orderNo);
  });
}

function deleteOrderLink(linkId) {
  requireRole_('Editor');
  return withLock_(() => {
    const k = readOrderLinks_().find(x => x.id === linkId);
    if (!k) throw new Error('That link was already removed.');
    requireOrderOwner_(k.orderNo);
    orderLinksSheet_().deleteRow(k.row);
    return getOrders().find(o => o.orderNo === k.orderNo);
  });
}

// ----- Additional costs on an order (any number: crane rental, permit, install labor…), in the Order Costs tab -----
// They show above the shipping cost and count in the order's total, like shipping. Allowed in any status.

function readOrderCosts_() {
  const sheet = orderCostsSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, ORDER_COST_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2, timestamp: fmtTime_(r[0]), user: String(r[1]), orderNo: String(r[2]), label: String(r[3]), amount: Number(r[4]) || 0, id: String(r[5]),
  })).filter(c => c.id && c.orderNo);
}

/** A description (required) and an amount for an additional cost; refuses anything else. */
function cleanOrderCost_(label, amount) {
  label = String(label || '').replace(/\s+/g, ' ').trim();
  if (!label) throw new Error('Enter what the cost is for.');
  if (label.length > 100) throw new Error('Keep the description under 100 characters.');
  const v = String(amount ?? '').trim(), n = Number(v.replace(/[$,\s]/g, ''));
  if (!v || !isFinite(n) || n < 0) throw new Error('The amount must be like 250 or $1,250.50.');
  return { label, amount: roundMoney_(n) };
}

function touchOrder_(order) {
  ordersSheet_().getRange(order.row, ORDER_HEADERS.indexOf('Updated') + 1, 1, 2).setValues([[new Date(), currentUser_()]]);
}

/** Adds an additional cost to an order. Returns the updated order. */
function addOrderCost(orderNo, label, amount) {
  requireRole_('Editor');
  requireOrderOwner_(orderNo);
  const c = cleanOrderCost_(label, amount);
  return withLock_(() => {
    const order = readOrderRows_().find(o => o.orderNo === orderNo);
    if (!order) throw new Error('Order ' + orderNo + ' was not found.');
    if (readOrderCosts_().filter(x => x.orderNo === orderNo).length >= MAX_ORDER_COSTS) throw new Error('An order can have up to ' + MAX_ORDER_COSTS + ' additional costs.');
    const sheet = orderCostsSheet_();
    sheet.appendRow([new Date(), currentUser_(), orderNo, asText_(c.label), c.amount, Utilities.getUuid()]);
    sheet.getRange(sheet.getLastRow(), 5).setNumberFormat('$#,##0.00');
    touchOrder_(order);
    return getOrders().find(o => o.orderNo === orderNo);
  });
}

/** Changes an additional cost's description and amount. */
function updateOrderCost(costId, label, amount) {
  requireRole_('Editor');
  const c = cleanOrderCost_(label, amount);
  return withLock_(() => {
    const x = readOrderCosts_().find(k => k.id === costId);
    if (!x) throw new Error('That cost was removed. Refresh and try again.');
    requireOrderOwner_(x.orderNo);
    orderCostsSheet_().getRange(x.row, 4, 1, 2).setValues([[asText_(c.label), c.amount]]);
    const order = readOrderRows_().find(o => o.orderNo === x.orderNo);
    if (order) touchOrder_(order);
    return getOrders().find(o => o.orderNo === x.orderNo);
  });
}

function deleteOrderCost(costId) {
  requireRole_('Editor');
  return withLock_(() => {
    const x = readOrderCosts_().find(k => k.id === costId);
    if (!x) throw new Error('That cost was already removed.');
    requireOrderOwner_(x.orderNo);
    orderCostsSheet_().deleteRow(x.row);
    const order = readOrderRows_().find(o => o.orderNo === x.orderNo);
    if (order) touchOrder_(order);
    return getOrders().find(o => o.orderNo === x.orderNo);
  });
}

function setOrderTracking(orderNo, url) {
  requireRole_('Editor');
  requireOrderOwner_(orderNo);
  const clean = cleanUrl_(url);
  return withLock_(() => {
    const order = readOrderRows_().find(o => o.orderNo === orderNo);
    if (!order) throw new Error('Order ' + orderNo + ' was not found.');
    const sheet = ordersSheet_();
    const H = h => ORDER_HEADERS.indexOf(h) + 1;
    sheet.getRange(order.row, H('Tracking Link')).setValue(clean);
    sheet.getRange(order.row, H('Updated'), 1, 2).setValues([[new Date(), currentUser_()]]);
    return getOrders().find(o => o.orderNo === orderNo);
  });
}

function readOrderDocs_() {
  const sheet = orderDocsSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, ORDER_DOC_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2, timestamp: fmtTime_(r[0]), user: String(r[1]), orderNo: String(r[2]), name: String(r[4]),
    sizeKB: Number(r[5]) || 0, fileId: String(r[6]), url: String(r[7]),
  })).filter(d => d.fileId);
}

/** PDFs attached to one order, newest first. */
function getOrderDocuments(orderNo) {
  requireRole_('Viewer');
  return readOrderDocs_().filter(d => d.orderNo === orderNo).reverse()
    .map(d => ({ fileId: d.fileId, name: d.name, sizeKB: d.sizeKB, user: d.user, timestamp: d.timestamp, url: d.url }));
}

/** Attaches a PDF (bill of lading, packing slip, invoice…) to an order. Allowed in any status. */
function uploadOrderDocument(orderNo, fileName, dataUrl) {
  requireRole_('Editor');
  requireOrderOwner_(orderNo);
  const pdf = parsePdf_(fileName, dataUrl);
  return withLock_(() => {
    const order = readOrderRows_().find(o => o.orderNo === orderNo);
    if (!order) throw new Error('Order ' + orderNo + ' was not found.');
    const project = [order.projectName, order.projectNo && '#' + order.projectNo].filter(String).join(' ');
    const file = savePdf_(ORDER_DOCS_FOLDER, pdf, 'Attached to order ' + orderNo + (project ? ' (' + project + ')' : '') + ' from Supplybot');
    orderDocsSheet_().appendRow([new Date(), currentUser_(), orderNo, asText_(project), asText_(pdf.name),
      Math.max(1, Math.round(pdf.bytes.length / 1024)), file.getId(), file.getUrl()]);
    return getOrderDocuments(orderNo);
  });
}

// ----- Supply Chain Requests (read-only view of the HubSpot ticket sync) -----
// The "New Inventory Requests" tab of the Supply Chain 2026 spreadsheet is kept up to date from HubSpot. Each Ticket name
// there is a HYPERLINK to the ticket in HubSpot; the app lists the tickets with that link, newest first.

const SC_SPREADSHEET_ID = '1AP1dTA9go1w71IJqKiexau15YwVtGPDokuAdmucFSvE';
const SC_SHEET = 'New Inventory Requests';
const SC_FIELDS = {
  name: 'Ticket name', created: 'Create date', createdBy: 'Created by user ID', model: 'Charger Model', chargers: 'Number of Chargers',
  loadSharing: '# of Chargers Load Sharing', pedestals: '# of Pedestals', pedestalType: 'Pedestal Type', status: 'Ticket status',
  shipDate: 'Requested Shipping Date', receiver: 'Receiving Company/Name', contact: 'Point of Contact Name', phone: 'Point of Contact Number',
  address: 'Shipping Address', pickupDate: 'Requested Pick Up Date', foreman: 'Foreman', deals: 'Ticket → Deals', subsidy: 'Subsidy',
  priority: 'Priority', owner: 'Ticket owner', lastActivity: 'Last activity date', ticketId: ['Ticket ID (sync key)', 'Ticket ID'],
  description: ['Ticket description', 'Description'],
};
const HUBSPOT_TICKET_URL = 'https://app.hubspot.com/contacts/44375004/record/0-5/';

/** All supply chain tickets: { synced, tickets: [{ name, url, status, ... }] }, newest first. */
function getSupplyChainRequests() {
  requireRole_('Viewer');
  requireScreen_('supply');
  let sheet;
  try { sheet = SpreadsheetApp.openById(SC_SPREADSHEET_ID).getSheetByName(SC_SHEET); }
  catch (e) { throw new Error('The app can’t open the Supply Chain 2026 spreadsheet. The app’s owner needs at least view access to it.'); }
  if (!sheet) throw new Error('The "' + SC_SHEET + '" tab wasn’t found in the Supply Chain 2026 spreadsheet.');
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const synced = headers.find(h => /last synced/i.test(h)) || '';
  if (lastRow < 2) return { synced, tickets: [] };
  const col = {};
  Object.keys(SC_FIELDS).forEach(k => col[k] = findCol_(headers, SC_FIELDS[k]));
  if (col.name === -1) throw new Error('The "Ticket name" column wasn’t found in "' + SC_SHEET + '".');
  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const shown = range.getDisplayValues(), raw = range.getValues();
  const formulas = sheet.getRange(2, col.name + 1, lastRow - 1, 1).getFormulas();
  const tickets = [];
  shown.forEach((row, i) => {
    const name = String(row[col.name] || '').trim();
    if (!name) return;
    const t = {};
    Object.keys(SC_FIELDS).forEach(k => { if (col[k] !== -1) t[k] = String(row[col[k]] || '').trim(); });
    // Link: from the cell's HYPERLINK formula, or built from the Ticket ID
    const m = /HYPERLINK\(\s*"(https:\/\/app\.hubspot\.com\/[^"]+)"/i.exec(formulas[i][0] || '');
    t.url = m ? m[1] : /^\d+$/.test(t.ticketId || '') ? HUBSPOT_TICKET_URL + t.ticketId : '';
    const c = col.created !== -1 ? raw[i][col.created] : '';
    t.createdAt = c instanceof Date ? c.getTime() : 0;   // for sorting (a number, never a Date, so it can be sent to the app)
    tickets.push(t);
  });
  tickets.sort((a, b) => b.createdAt - a.createdAt);
  return { synced, tickets };
}

// ----- Order as a PDF -----
// "Make PDF" on an order builds a one-page summary with the EnviroSpark logo and saves it to the order's PDFs.
// withCosts = false leaves out every price, for a packing list to hand to a driver or customer.
// withPhotos adds the order's photos at the end, two per row (capped so the PDF stays a sensible size).
const PDF_PHOTO_BUDGET_MB = 12;

/** " (S/N 123 · dial 3 · 32 A draw · 40 A breaker · load-shares with Port B)" for endpoint i of a line, or ''. */
function epDetails_(l, i) {
  const v = k => String(((l[k] || [])[i]) || '').trim();
  const unit = x => /^\d+(\.\d+)?$/.test(x) ? x + ' A' : x;
  const parts = [v('serials') && 'S/N ' + v('serials'), v('rotary') && 'dial ' + v('rotary'), v('amps') && unit(v('amps')) + ' draw',
    v('breaker') && unit(v('breaker')) + ' breaker', v('share') && 'load-shares with ' + v('share')].filter(Boolean);
  return parts.length ? ' (' + parts.join(' · ') + ')' : '';
}

/** Pounds from the sheet's "Approx. Weight" text ("9 lbs", "40lbs", ".25 lbs", "1 lb", "0.175", "2 kg"), or null. Same rules as the app. */
function weightLbs_(v) {
  const m = /(\d*\.?\d+)\s*(kgs?|kilo\w*|oz|ounces?|lbs?|pounds?|#)?/i.exec(String(v == null ? '' : v).replace(/,/g, ''));
  const n = m ? parseFloat(m[1]) : NaN;
  if (!(n > 0)) return null;
  const u = (m[2] || 'lb').toLowerCase();
  return u.charAt(0) === 'k' ? n * 2.20462 : u.charAt(0) === 'o' ? n / 16 : n;
}
const fmtLbs_ = n => (n >= 100 ? Math.round(n) : Math.round(n * 100) / 100).toLocaleString('en-US') + ' lbs';

const PDF_LOGO = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAA+gAAAGWCAMAAADcyllOAAAAnFBMVEUAAABOTUpOTUpOTUpOTUpOTUpOTUpOTUpOTUpOTUpOTUpOTUpOTUpEREREREREREROTUpEREREREREREREREREREREREREREREREREREREREREREREREREREROTUpOTUpOTUqUumWUumWUumWUumWUumWUumWUumWUumWUumWUumWUumWUumWUumWUumWUumWUumVERESUumVOTUqi2qJDAAAAMXRSTlMAMACAryDvv1AQ34/PQDAAcIC/7xAgcN9gUI+vz59AYJ8QQHCAACBgj79Qn9/PMO+vypZo3QAAKIZJREFUeF7swDEBAAAAwiD7p7bETlgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADg7NexEcIwFARRJBkjgyX4UInn+u8NJSYgZxjf7Othg035t0753wCUyT10AOe5uIcO4KJqHjqAKpk/OoDlKmXv0AHcpNU7dABJ0uQdOoBVUrEJHUD76PfdQ0M9augAemsR8Rxe25fYO19mDelgob+5t7vduEEgDMNR2jRtWpgygPEPpNfwyfd/b00URwiyWTmWf2b9nO6J5eX1wQBO7cbRBVqVNMngVcHXv1YcSeJVhUp69sM7tTt9zBtrDDOuCPTu+/jqfmJMu82/f0cr0J3pmQN2pegCxkeW2TSKDmVQMFTQqBiSZEApXH/pij6lcIjIyXQ7Fa+anr+yep/HF3nobgHAct9qYaGrZojIxIWexdQ6oaFTRMmSIB6V9qZCz9go2pLr+oh5Ek1+PY2vHqfOHTLuOycjdN0wMrmhZ7F1IkNvP7YkR4+SdbcWemaHrVbAzMjrV/htfJGH7gql2KujQ9d9QCY99CwpgaE7ixKTHBalRDcYembXXwGuHfAlDU0exjf/ptCbS8/busNCd01AJjz0GitxoVOPiiYpWlT8jYaexZZW1A34IqZ3j+Ob5yn0HpfY1B0Suk8WuN3QAfbSQveoJJIiosR086EDoaN1+H5BC5om/8bJ/YTxiWD83qH7hJr80GtGWOjEKFlHMmhUWvGh7/exVwMW6Ist9FdPeej+uaT2DN0ZixOEjqhlhd6h0pAMCSVL4kPf62OvGEsER5Pf4yQP3a9itVvoXQBOETrQiAqdAkqBRHComNOEjugPyBzoii30euh+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';

function orderPdfHtml_(order, lines, links, withCosts, photos) {
  photos = photos || { list: [], skipped: 0 };
  const e = escHtml_;
  const day = s => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || ''); return m ? Utilities.formatDate(new Date(+m[1], m[2] - 1, +m[3]), Session.getScriptTimeZone(), 'MMM d, yyyy') : String(s || ''); };
  const stamp = iso => iso ? Utilities.formatDate(new Date(iso), Session.getScriptTimeZone(), "MMM d, yyyy 'at' h:mm a") : '';
  const usd = n => '$' + (Math.round(Number(n) * 100) / 100).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const facts = [
    ['Order Date', day(order.orderDate)], ['Expected Pickup / Ship', day(order.shipDate)], ['Delivery Date', day(order.deliveryDate)], ['Status', order.status], ['Project Name', order.projectName], ['Project #', order.projectNo],
    ['Activity Type', order.activityType], ['PM', order.pm], ['From', order.fromLocation], ['To', order.toLocation],
    ['Point of Contact', order.contact], ['Freight Service', order.freight], ['Company / Vendor', order.vendor],
    ['Shipped', order.status === 'Shipped' ? stamp(order.shipped) : ''],
  ].filter(f => String(f[1] || '').trim());
  const half = Math.ceil(facts.length / 2);
  const factRows = Array.from({ length: half }, (_, i) => [facts[i], facts[i + half]]).map(([a, b]) =>
    '<tr>' + [a, b].map(f => f ? '<td class="k">' + e(f[0]) + '</td><td class="v">' + e(f[1]) + '</td>' : '<td></td><td></td>').join('') + '</tr>').join('');
  const units = lines.reduce((n, l) => n + l.qty, 0);
  const cost = lines.reduce((n, l) => n + l.ourCost, 0);
  const job = j => j && !/^available$/i.test(j) ? j : '';
  const lbsTotal = lines.reduce((n, l) => n + (l.lbs || 0), 0), noWeight = lines.filter(l => l.lbs == null).length;
  const itemRows = lines.map((l, i) => '<tr' + (i % 2 ? ' class="alt"' : '') + '><td>' + e(l.name) + (job(l.job) ? '<div class="sub">' + e(job(l.job)) + '</div>' : '') +
    ((l.endpoints || []).length ? '<div class="sub">Endpoints: ' + e(l.endpoints.map((p, i) => p + epDetails_(l, i)).join('; ')) + '</div>' : '') +
    (l.dims ? '<div class="sub">Box: ' + e(l.dims) + '</div>' : '') + '</td>' +
    '<td>' + e(l.partNumber) + '</td><td class="n">' + l.qty + '</td>' +
    '<td class="n">' + (l.lbs == null ? '—' : fmtLbs_(l.lbs) + (l.qty > 1 ? '<div class="sub">' + fmtLbs_(l.eachLbs) + ' ea</div>' : '')) + '</td>' +
    (withCosts ? '<td class="n">' + (l.cost == null ? '' : usd(l.cost)) + '</td><td class="n">' + usd(l.ourCost) + '</td>' : '') + '</tr>').join('');
  const totals = withCosts
    ? '<tr><td class="k">Our cost</td><td class="n">' + usd(cost) + '</td></tr>' +
      (order.extraCosts || []).map(c => '<tr><td class="k">' + e(c.label) + '</td><td class="n">' + usd(c.amount) + '</td></tr>').join('') +
      (order.shippingCost != null ? '<tr><td class="k">Shipping</td><td class="n">' + usd(order.shippingCost) + '</td></tr>' : '') +
      (order.shippingCost != null || (order.extraCosts || []).length
        ? '<tr class="grand"><td class="k">Total</td><td class="n">' + usd(cost + (order.shippingCost || 0) + (order.extraCost || 0)) + '</td></tr>' : '')
    : '';
  const linkList = (order.trackingUrl ? [{ label: 'Tracking', url: order.trackingUrl }] : []).concat(links);   // same order as in the app
  return '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,Helvetica,sans-serif;color:#1f2937;font-size:11pt;margin:0}' +
    '.head{width:100%;border-bottom:3px solid #8fb85f;margin-bottom:14px}.head td{vertical-align:bottom;padding-bottom:10px}' +
    '.logo{width:150px}.title{text-align:right}.title h1{margin:0;font-size:20pt;color:#374151}.title .st{font-size:11pt;color:#6b7280;margin-top:2px}' +
    '.facts{width:100%;border-collapse:collapse;margin-bottom:16px}.facts td{padding:4px 6px;vertical-align:top}.facts .k{color:#6b7280;width:17%;white-space:nowrap}.facts .v{width:33%;font-weight:bold}' +
    'h2{font-size:12pt;color:#374151;margin:16px 0 6px;text-transform:uppercase;letter-spacing:1px}' +
    '.items{width:100%;border-collapse:collapse}.items th{background:#374151;color:#fff;text-align:left;padding:7px 8px;font-size:10pt}' +
    '.items td{padding:7px 8px;border-bottom:1px solid #e5e7eb;vertical-align:top}.items tr.alt td{background:#f6f8f3}.items .n{text-align:right;white-space:nowrap}' +
    '.items .sub{color:#6b7280;font-size:9pt}.items tfoot td{font-weight:bold;border-top:2px solid #374151;border-bottom:0}' +
    '.totals{margin-left:auto;margin-top:10px;border-collapse:collapse;min-width:45%}.totals td{padding:4px 8px}.totals .k{color:#6b7280}.totals .n{text-align:right;font-weight:bold}' +
    '.totals tr.grand td{border-top:2px solid #374151;font-size:12pt;color:#111827}' +
    '.notes{white-space:pre-wrap;background:#f6f8f3;border-left:3px solid #8fb85f;padding:8px 10px}.links a{color:#2f6f3e}' +
    '.photos{width:100%;border-collapse:separate;border-spacing:0 10px}.photos td{width:50%;vertical-align:top;padding:0 5px;page-break-inside:avoid}' +
    '.photos td{text-align:center}.photos img{max-width:100%;max-height:250px;width:auto;height:auto;border-radius:6px;border:1px solid #e5e7eb}.photos .cap{color:#6b7280;font-size:8.5pt;margin-top:3px}' +
    '.shipto{border:1px solid #e5e7eb;border-left:3px solid #8fb85f;padding:8px 12px;margin:-6px 0 14px}.shipto .k{color:#6b7280;font-size:9pt;text-transform:uppercase;letter-spacing:1px}.shipto .v{font-weight:bold;margin-top:2px}' +
    '.foot{margin-top:24px;color:#9ca3af;font-size:8.5pt;border-top:1px solid #e5e7eb;padding-top:6px}' +
    '</style></head><body>' +
    '<table class="head"><tr><td><img class="logo" src="' + PDF_LOGO + '"></td>' +
    '<td class="title"><h1>' + (withCosts ? 'Order ' : 'Packing List ') + e(order.orderNo) + '</h1><div class="st">' +
    e([order.projectName, order.projectNo && '#' + order.projectNo].filter(String).join(' · ')) + '</div></td></tr></table>' +
    '<table class="facts">' + factRows + '</table>' +
    (order.shipTo ? '<div class="shipto"><div class="k">Ship To</div><div class="v">' + e(order.shipTo).replace(/\n/g, '<br>') + '</div></div>' : '') +
    '<h2>Items</h2><table class="items"><thead><tr><th>Item</th><th>Part #</th><th class="n">Qty</th><th class="n">Weight</th>' + (withCosts ? '<th class="n">Cost Price</th><th class="n">Our Cost</th>' : '') + '</tr></thead>' +
    '<tbody>' + (itemRows || '<tr><td colspan="6">No items on this order.</td></tr>') + '</tbody>' +
    '<tfoot><tr><td colspan="2">' + lines.length + ' item' + (lines.length === 1 ? '' : 's') + '</td><td class="n">' + units + '</td>' +
    '<td class="n">' + (noWeight === lines.length ? '—' : fmtLbs_(lbsTotal)) + '</td>' + (withCosts ? '<td></td><td class="n">' + usd(cost) + '</td>' : '') + '</tr></tfoot></table>' +
    (totals ? '<table class="totals">' + totals + '</table>' : '') +
    (noWeight && noWeight < lines.length ? '<div class="sub" style="color:#6b7280;font-size:9pt;margin-top:6px">Gross weight leaves out ' + noWeight + ' item' + (noWeight === 1 ? '' : 's') + ' with no weight in the inventory sheet.</div>' : '') +
    (order.notes ? '<h2>Notes</h2><div class="notes">' + e(order.notes) + '</div>' : '') +
    (linkList.length ? '<h2>Links</h2><div class="links">' + linkList.map(k => e(k.label || 'Link') + ': <a href="' + e(k.url) + '">' + e(k.url) + '</a>').join('<br>') + '</div>' : '') +
    (photos.list.length ? '<h2 style="page-break-before:' + (photos.list.length > 2 ? 'always' : 'auto') + '">Photos (' + (photos.list.length + photos.skipped) + ')</h2><table class="photos">' +
      Array.from({ length: Math.ceil(photos.list.length / 2) }, (_, i) => photos.list.slice(i * 2, i * 2 + 2)).map(pair => '<tr>' +
        pair.concat(pair.length < 2 ? [null] : []).map(ph => ph ? '<td><img src="' + ph.src + '"><div class="cap">' + e(stamp(ph.timestamp)) + (ph.user ? ' · ' + e(ph.user.split('@')[0]) : '') + '</div></td>' : '<td></td>').join('') +
        '</tr>').join('') + '</table>' +
      (photos.skipped ? '<div class="cap" style="color:#6b7280">' + photos.skipped + ' more photo' + (photos.skipped === 1 ? '' : 's') + ' in the app (left out to keep this PDF a manageable size).</div>' : '') : '') +
    '<div class="foot">Made with Supplybot on ' + e(stamp(new Date().toISOString())) + ' by ' + e(currentUser_()) + '</div>' +
    '</body></html>';
}

// ---------- Catalog PDF (Reports → 📒 Catalog) ----------
// The app sends the catalog exactly as previewed: the picked categories with their items, and what to show. The server
// adds the photos (only ones listed on items) and lays it out like the order PDFs, then saves it to the "Catalogs" folder.
const CATALOG_FOLDER = 'Catalogs';
const CATALOG_MAX_ITEMS = 800, CATALOG_PHOTO_PX = 220, CATALOG_PHOTO_BUDGET_MB = 18;
// Every catalog made is listed in the "Catalog PDFs" tab with its setup, so anyone can reopen the PDF or load it again.
const CATALOGS_SHEET = 'Catalog PDFs';
const CATALOG_HEADERS = ['Created', 'Created By', 'Title', 'Items', 'Qty', 'Total', 'PDF Name', 'File ID', 'PDF URL', 'Settings', 'ID'];
function catalogsSheet_() { return tabWithHeaders_(CATALOGS_SHEET, CATALOG_HEADERS); }
function readCatalogs_() {
  const sh = catalogsSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, CATALOG_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2, created: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), by: String(r[1]), title: String(r[2]),
    items: Number(r[3]) || 0, qty: Number(r[4]) || 0, total: r[5] === '' ? null : Number(r[5]), name: String(r[6]), fileId: String(r[7]),
    url: String(r[8]), settings: String(r[9]), id: String(r[10]),
  })).filter(c => c.id);
}
/** The saved catalogs, newest first (up to 100), for Reports → 📒 Catalog. */
function getCatalogs() {
  requireRole_('Viewer');
  requireScreen_('reports');
  return readCatalogs_().reverse().slice(0, 100).map(({ row, ...c }) => c);
}
/** Takes a catalog off the list (the PDF stays in Drive). Your own, or any for admins. */
function deleteCatalog(id) {
  requireRole_('Viewer');
  requireScreen_('reports');
  return withLock_(() => {
    const c = readCatalogs_().find(x => x.id === String(id));
    if (!c) return getCatalogs();
    if (c.by.toLowerCase() !== currentUser_().toLowerCase() && myRole_() !== 'Admin') throw new Error('Only the person who made it (or an admin) can remove it.');
    catalogsSheet_().deleteRow(c.row);
    return getCatalogs();
  });
}

/**
 * cat = { title, groups: [{ name, items: [{ name, partNumber, manufacturer, specs, dims, weight, qty, cost, photoId }] }], showQty, showCost, showPhotos,
 *   sales, headline, validTo ('yyyy-MM-dd'), contact }. sales: written for customers (catalogPdfSalesHtml_), not as an inventory list.
 */
function createCatalogPdf(cat) {
  requireRole_('Viewer');
  requireScreen_('reports');
  const { html, title, groups, count, units, value, showCost, tz } = buildCatalog_(cat);
  const name = (title + ' - ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')).replace(/[\\/:*?"<>|]+/g, ' ').trim().slice(0, 110) + '.pdf';
  const bytes = Utilities.newBlob(html, 'text/html', 'catalog.html').getAs('application/pdf').getBytes();
  const file = savePdf_(CATALOG_FOLDER, { bytes, name }, 'Catalog of ' + count + ' items (' + groups.map(g => g.name).join(', ') + ') made in Supplybot');
  // remember it (with the app's setup, so it can be loaded again)
  const settings = String(cat.settings || '').slice(0, 45000);
  withLock_(() => catalogsSheet_().appendRow([new Date(), currentUser_(), asText_(title), count, units, showCost ? Math.round(value * 100) / 100 : '',
    asText_(name), file.getId(), file.getUrl(), settings, Utilities.getUuid()]));
  return { url: file.getUrl(), name, count, catalogs: getCatalogs() };
}

/** The same catalog as a web page, to show on screen (🖥 Show on screen): exactly what the PDF will look like. */
function getCatalogPreview(cat) {
  requireRole_('Viewer');
  requireScreen_('reports');
  return buildCatalog_(cat).html;
}

/** Checks the catalog the app sent, fetches the photos and lays it out in the chosen design. Shared by the PDF and the screen. */
function buildCatalog_(cat) {
  const groups = (cat && Array.isArray(cat.groups) ? cat.groups : []).map(g => ({
    name: String(g.name || '(no category)').slice(0, 80),
    items: (Array.isArray(g.items) ? g.items : []).map(i => ({
      name: String(i.name || '').slice(0, 200), partNumber: String(i.partNumber || '').slice(0, 80), manufacturer: String(i.manufacturer || '').slice(0, 80),
      specs: String(i.specs || '').slice(0, 120), dims: String(i.dims || '').slice(0, 80), weight: String(i.weight || '').slice(0, 40),
      qty: Number(i.qty) || 0, cost: i.cost == null || i.cost === '' ? null : Number(i.cost), photoId: String(i.photoId || '').slice(0, 100), fav: !!i.fav,
    })).filter(i => i.name),
  })).filter(g => g.items.length);
  const count = groups.reduce((n, g) => n + g.items.length, 0);
  if (!count) throw new Error('Pick at least one category with items in it.');
  if (count > CATALOG_MAX_ITEMS) throw new Error('That’s ' + count + ' items; a catalog can hold up to ' + CATALOG_MAX_ITEMS + '. Pick fewer categories.');
  const showQty = !!cat.showQty, showCost = !!cat.showCost, showPhotos = !!cat.showPhotos;
  const atList = cat.priceLabel === 'list';   // the app adjusted Cost Price by a percent: show it as the price, not "at cost"
  // Photos: small copies, fetched in batches, within a size budget so the PDF stays manageable
  const photos = {};
  let skipped = 0;
  if (showPhotos) {
    const ids = [...new Set(groups.flatMap(g => g.items.map(i => i.photoId)).filter(Boolean))];
    const allowed = thumbAllowed_(ids);
    const ok = ids.filter(id => allowed.has(id));
    const px = ['gallery', 'showroom', 'brochure', 'night', 'blueprint', 'neon', 'datasheet', 'editorial', 'luxe', 'gauge'].includes(cat.design) ? 420 : CATALOG_PHOTO_PX;   // the big-photo designs get sharper copies
    for (let k = 0; k < ok.length; k += 20) driveThumbs_(ok.slice(k, k + 20), px, photos);
    let budget = CATALOG_PHOTO_BUDGET_MB * 1024 * 1024;
    Object.keys(photos).forEach(id => { const n = photos[id].length; if (n > budget) { delete photos[id]; skipped++; } else budget -= n; });
  }
  const e = escHtml_, tz = Session.getScriptTimeZone();
  const usd = n => '$' + (Math.round(Number(n) * 100) / 100).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const units = groups.reduce((n, g) => n + g.items.reduce((m, i) => m + i.qty, 0), 0);
  const value = groups.reduce((n, g) => n + g.items.reduce((m, i) => m + (i.cost == null ? 0 : i.cost * i.qty), 0), 0);
  const card = i => '<td class="card"><table class="in"><tr>' +
    (showPhotos ? '<td class="ph">' + (photos[i.photoId] ? '<img src="' + photos[i.photoId] + '">' : '<div class="nophoto">No photo</div>') + '</td>' : '') +
    '<td class="tx"><div class="nm">' + e(i.name) + '</div>' +
    [i.partNumber && 'Part # ' + i.partNumber, i.manufacturer].filter(Boolean).map(t => '<div class="sub">' + e(t) + '</div>').join('') +
    (i.specs ? '<div class="sub">' + e(i.specs) + '</div>' : '') +
    ([i.dims && 'Box ' + i.dims, i.weight && i.weight].filter(Boolean).length ? '<div class="sub">' + e([i.dims && 'Box ' + i.dims, i.weight].filter(Boolean).join(' · ')) + '</div>' : '') +
    ((showQty || (showCost && i.cost != null)) ? '<div class="ft">' + [showQty ? '<span>Qty ' + i.qty + '</span>' : '', showCost && i.cost != null ? '<b>' + usd(i.cost) + ' ea</b>' : ''].filter(Boolean).join(' · ') + '</div>' : '') +
    '</td></tr></table></td>';
  const sections = groups.map((g, gi) => '<h2' + (gi ? '' : ' class="first"') + '>' + e(g.name) + ' <span>' + g.items.length + ' item' + (g.items.length === 1 ? '' : 's') +
      (showQty ? ' · qty ' + g.items.reduce((n, i) => n + i.qty, 0) : '') + '</span></h2><table class="grid">' +
    Array.from({ length: Math.ceil(g.items.length / 2) }, (_, r) => '<tr>' + card(g.items[r * 2]) + (g.items[r * 2 + 1] ? card(g.items[r * 2 + 1]) : '<td class="card empty"></td>') + '</tr>').join('') +
    '</table>').join('');
  const title = String(cat.title || 'Catalog').slice(0, 80);
  const design = CATALOG_DESIGNS.includes(cat.design) ? cat.design : 'standard';
  const html = design !== 'standard' ? catalogDesignHtml_(design, title, groups, photos, { showQty, showCost, showPhotos, skipped, atList, sales: !!cat.sales,
      headline: String(cat.headline || '').slice(0, 90), validTo: String(cat.validTo || ''), contact: String(cat.contact || '').slice(0, 120) || currentUser_(),
      preparedFor: String(cat.preparedFor || '').slice(0, 160), units, value })
    : cat.sales ? catalogSalesHtml_(title, groups, photos, { showQty, showCost, showPhotos, skipped,
      headline: String(cat.headline || '').slice(0, 90), validTo: String(cat.validTo || ''), contact: String(cat.contact || '').slice(0, 120) || currentUser_(),
      preparedFor: String(cat.preparedFor || '').slice(0, 160) }) : '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,Helvetica,sans-serif;color:#1f2937;font-size:10pt;margin:0}' +
    '.head{width:100%;border-bottom:3px solid #8fb85f;margin-bottom:10px}.head td{vertical-align:bottom;padding-bottom:10px}.logo{width:150px}' +
    '.title{text-align:right}.title h1{margin:0;font-size:22pt;color:#374151}.title .st{font-size:10pt;color:#6b7280;margin-top:2px}' +
    '.sum{width:100%;border-collapse:collapse;margin:0 0 6px}.sum td{padding:6px 8px;background:#f6f8f3;border-left:3px solid #8fb85f}.sum b{font-size:13pt;color:#374151}.sum span{color:#6b7280;font-size:9pt}' +
    'h2{font-size:12.5pt;color:#374151;margin:16px 0 6px;padding-bottom:4px;border-bottom:1px solid #e5e7eb;text-transform:uppercase;letter-spacing:1px;page-break-after:avoid}' +
    'h2 span{font-size:9pt;color:#6b7280;text-transform:none;letter-spacing:0;font-weight:normal}' +
    '.grid{width:100%;border-collapse:separate;border-spacing:6px}.card{width:50%;vertical-align:top;border:1px solid #e5e7eb;border-radius:6px;padding:6px;page-break-inside:avoid}' +
    '.card.empty{border:0}.in{width:100%;border-collapse:collapse}.in td{vertical-align:top}.ph{width:92px;padding-right:8px}' +
    '.ph img{width:88px;height:88px;object-fit:contain;border-radius:4px}.nophoto{width:88px;height:88px;background:#f3f4f6;color:#9ca3af;font-size:8pt;text-align:center;line-height:88px;border-radius:4px}' +
    '.nm{font-weight:bold;font-size:10.5pt;color:#111827;margin-bottom:2px}.sub{color:#6b7280;font-size:8.5pt}.ft{margin-top:5px;font-size:9pt;color:#374151}.ft b{color:#2f6f3e}' +
    '.foot{margin-top:18px;color:#9ca3af;font-size:8pt;border-top:1px solid #e5e7eb;padding-top:6px}' +
    '</style></head><body>' +
    '<table class="head"><tr><td><img class="logo" src="' + PDF_LOGO + '"></td><td class="title"><h1>' + e(title) + '</h1><div class="st">' +
    e(Utilities.formatDate(new Date(), tz, 'MMMM d, yyyy')) + '</div></td></tr></table>' +
    '<table class="sum"><tr><td><b>' + count + '</b> <span>item' + (count === 1 ? '' : 's') + ' in ' + groups.length + ' categor' + (groups.length === 1 ? 'y' : 'ies') + ': ' +
    e(groups.map(g => g.name).join(', ')) + '</span>' + (showQty ? ' &nbsp; <b>' + units + '</b> <span>qty</span>' : '') + (showCost ? ' &nbsp; <b>' + usd(value) + '</b> <span>' + (atList ? 'total value' : 'value at cost') + '</span>' : '') + '</td></tr></table>' +
    sections +
    (skipped ? '<div class="sub" style="margin-top:8px">' + skipped + ' photo' + (skipped === 1 ? ' was' : 's were') + ' left out to keep this PDF a manageable size.</div>' : '') +
    '<div class="foot">Made with Supplybot on ' + e(Utilities.formatDate(new Date(), tz, "MMM d, yyyy 'at' h:mm a")) + ' by ' + e(currentUser_()) + '</div></body></html>';
  return { html, title, groups, count, units, value, showCost, tz };
}

/**
 * The Sales catalog PDF: the same items, written to sell. A SALE ribbon and headline, "products ready to ship" instead of
 * inventory counts, a badge on each product (In stock · ready to ship / Only N left!), its price (or "Call for pricing"),
 * and a "Ready to order?" box with who to contact and how long the prices hold. No cost or totals.
 */
function catalogSalesHtml_(title, groups, photos, o) {
  const e = escHtml_, tz = Session.getScriptTimeZone();
  const usd = n => '$' + (Math.round(Number(n) * 100) / 100).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const count = groups.reduce((n, g) => n + g.items.length, 0);
  const valid = /^\d{4}-\d{2}-\d{2}$/.test(o.validTo) ? Utilities.formatDate(new Date(o.validTo + 'T12:00:00'), tz, 'MMMM d, yyyy') : '';
  const badge = q => !o.showQty ? '<span class="bd ok">In stock</span>'
    : q <= 0 ? '<span class="bd out">Ask about availability</span>'
    : q <= 5 ? '<span class="bd low">Only ' + q + ' left!</span>' : '<span class="bd ok">' + q + ' in stock · ready to ship</span>';
  const card = i => '<td class="card"><table class="in"><tr>' +
    (o.showPhotos ? '<td class="ph">' + (photos[i.photoId] ? '<img src="' + photos[i.photoId] + '">' : '<div class="nophoto">Photo coming soon</div>') + '</td>' : '') +
    '<td class="tx">' + (i.fav ? '<span class="bd fav">★ Ordered before</span> ' : '') + badge(i.qty) + '<div class="nm">' + e(i.name) + '</div>' +
    (i.manufacturer ? '<div class="mf">by ' + e(i.manufacturer) + '</div>' : '') +
    (i.specs ? '<div class="sub">⚡ ' + e(i.specs) + '</div>' : '') +
    [i.partNumber && 'Part # ' + i.partNumber, i.dims && 'Ships in ' + i.dims].filter(Boolean).map(t => '<div class="sub">' + e(t) + '</div>').join('') +
    (o.showCost ? '<div class="pr">' + (i.cost != null ? '<b>' + usd(i.cost) + '</b><span> each</span>' : '<b class="call">Call for pricing</b>') + '</div>' : '') +
    '</td></tr></table></td>';
  const favs = groups.reduce((n, g) => n + g.items.filter(i => i.fav).length, 0);
  const sections = groups.map((g, gi) => '<h2' + (gi ? '' : ' class="first"') + '><span class="dot"></span>' + e(g.name) +
      ' <span class="n">' + g.items.length + ' product' + (g.items.length === 1 ? '' : 's') + ' available</span></h2><table class="grid">' +
    Array.from({ length: Math.ceil(g.items.length / 2) }, (_, r) => '<tr>' + card(g.items[r * 2]) + (g.items[r * 2 + 1] ? card(g.items[r * 2 + 1]) : '<td class="card empty"></td>') + '</tr>').join('') +
    '</table>').join('');
  return '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,Helvetica,sans-serif;color:#1f2937;font-size:10pt;margin:0}' +
    '.head{width:100%;border-collapse:collapse}.head td{vertical-align:middle;padding-bottom:8px}.logo{width:150px}' +
    '.rib{display:inline-block;background:#dc2626;color:#fff;font-weight:bold;font-size:11pt;letter-spacing:3px;padding:4px 14px;border-radius:3px}' +
    '.title{text-align:right}.title h1{margin:6px 0 0;font-size:24pt;color:#111827}' +
    '.hero{background:#14532d;color:#fff;border-radius:8px;padding:12px 16px;margin:4px 0 12px}' +
    '.hero .hl{font-size:16pt;font-weight:bold;color:#d9f99d}.hero .tg{font-size:10pt;margin-top:3px;color:#ecfccb}' +
    'h2{font-size:13pt;color:#14532d;margin:16px 0 6px;padding-bottom:4px;border-bottom:2px solid #8fb85f;page-break-after:avoid}' +
    'h2 .dot{display:inline-block;width:9px;height:9px;border-radius:50%;background:#8fb85f;margin-right:7px;vertical-align:1px}' +
    'h2 .n{font-size:9pt;color:#4b5563;font-weight:normal}' +
    '.grid{width:100%;border-collapse:separate;border-spacing:7px}.card{width:50%;vertical-align:top;border:1px solid #d9e7c8;border-radius:8px;padding:7px;page-break-inside:avoid;background:#fcfef9}' +
    '.card.empty{border:0;background:none}.in{width:100%;border-collapse:collapse}.in td{vertical-align:top}.ph{width:96px;padding-right:9px}' +
    '.ph img{width:92px;height:92px;object-fit:contain;border-radius:6px;background:#fff}.nophoto{width:92px;height:92px;background:#f3f4f6;color:#9ca3af;font-size:7.5pt;text-align:center;line-height:92px;border-radius:6px}' +
    '.bd{display:inline-block;font-size:7.5pt;font-weight:bold;padding:2px 7px;border-radius:9px;margin-bottom:4px}.bd.ok{background:#dcfce7;color:#166534}' +
    '.bd.low{background:#fee2e2;color:#b91c1c}.bd.out{background:#f3f4f6;color:#4b5563}.bd.fav{background:#fef3c7;color:#92400e}' +
    '.pf{font-size:10.5pt;color:#374151;margin-top:4px}.pf b{color:#111827}' +
    '.nm{font-weight:bold;font-size:11pt;color:#111827}.mf{font-size:8.5pt;color:#4b5563;font-style:italic;margin-bottom:2px}.sub{color:#6b7280;font-size:8.5pt}' +
    '.pr{margin-top:6px}.pr b{font-size:14pt;color:#15803d}.pr span{font-size:8.5pt;color:#6b7280}.pr .call{font-size:10pt;color:#b45309}' +
    '.cta{margin-top:16px;border:2px solid #8fb85f;border-radius:8px;padding:12px 16px;background:#f7fbf1;page-break-inside:avoid}' +
    '.cta .sbot{float:right;width:58px;height:58px;margin:0 0 6px 12px;border-radius:12px}' +
    '.cta .q{font-size:14pt;font-weight:bold;color:#14532d}.cta .w{margin-top:4px;font-size:10.5pt}.cta .f{margin-top:6px;font-size:8.5pt;color:#6b7280}' +
    '.foot{margin-top:12px;color:#9ca3af;font-size:7.5pt}' +
    '</style></head><body>' +
    '<table class="head"><tr><td><img class="logo" src="' + PDF_LOGO + '"></td><td class="title"><span class="rib">SALE</span><h1>' + e(title) + '</h1>' +
    (o.preparedFor ? '<div class="pf">Prepared for <b>' + e(o.preparedFor) + '</b></div>' : '') + '</td></tr></table>' +
    '<div class="hero"><div class="hl">' + e(o.headline || 'Quality equipment, in stock and ready to ship') + '</div><div class="tg">' +
    count + ' product' + (count === 1 ? '' : 's') + ' available now' + (favs ? ' · ★ ' + favs + ' you’ve ordered before' : '') + (valid ? ' · Prices good through ' + e(valid) : '') + ' · While supplies last</div></div>' +
    sections +
    (o.skipped ? '<div class="sub" style="margin-top:8px">More photos available on request.</div>' : '') +
    '<div class="cta"><img class="sbot" src="data:image/png;base64,' + FAVICON_BOT_PNG_BASE64 + '"><div class="q">Ready to order?</div><div class="w">Contact <b>' + e(o.contact) + '</b> to reserve yours, ask about volume pricing or get a quote.</div>' +
    '<div class="f">' + (valid ? 'Prices valid through ' + e(valid) + '. ' : '') + 'Quantities are limited and subject to prior sale. Prices exclude tax and shipping.</div></div>' +
    '<div class="foot">' + e(Utilities.formatDate(new Date(), tz, 'MMMM d, yyyy')) + ' · EnviroSpark</div></body></html>';
}

// ---------- Catalog designs (📒 Catalog → Design): Gallery, Showroom, Night, Price list, Brochure ----------
// Each works for both inventory and sales catalogs (o.sales switches the wording, the same as the Standard layouts).
// Tables only for layout: Google's PDF maker doesn't do flexbox or grid, and the screen view uses the same page.
const CATALOG_DESIGNS = ['standard', 'gallery', 'showroom', 'night', 'pricelist', 'brochure', 'blueprint', 'neon', 'datasheet', 'editorial', 'luxe', 'gauge'];
const CATALOG_DARK = ['night', 'blueprint', 'neon', 'luxe', 'gauge'];   // dark pages: the logo sits on a white tile
function catalogDesignHtml_(design, title, groups, photos, o) {
  const e = escHtml_, tz = Session.getScriptTimeZone();
  const usd = n => '$' + (Math.round(Number(n) * 100) / 100).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const count = groups.reduce((n, g) => n + g.items.length, 0);
  const favs = groups.reduce((n, g) => n + g.items.filter(i => i.fav).length, 0);
  const valid = /^\d{4}-\d{2}-\d{2}$/.test(o.validTo) ? Utilities.formatDate(new Date(o.validTo + 'T12:00:00'), tz, 'MMMM d, yyyy') : '';
  const today = Utilities.formatDate(new Date(), tz, 'MMMM d, yyyy');
  const T = {   // each design's colours
    gallery: { bg: '#ffffff', ink: '#111827', mute: '#6b7280', acc: '#5f8a2e', line: '#e5e7eb', band: '#f6f8f3' },
    showroom: { bg: '#ffffff', ink: '#0f172a', mute: '#64748b', acc: '#15803d', line: '#e2e8f0', band: '#f1f5f9' },
    night: { bg: '#0b1220', ink: '#e5e7eb', mute: '#94a3b8', acc: '#a3e635', line: '#1f2a44', band: '#111a2e' },
    pricelist: { bg: '#ffffff', ink: '#1f2937', mute: '#6b7280', acc: '#2f6f3e', line: '#e5e7eb', band: '#f3f4f6' },
    brochure: { bg: '#fbfdf7', ink: '#1c2a13', mute: '#5b6b4e', acc: '#4d7c0f', line: '#dfe9cf', band: '#4d7c0f' },
    blueprint: { bg: '#0d3b7a', ink: '#eaf2ff', mute: '#a9c3ea', acc: '#ffffff', line: '#6f97d6', band: '#0a2f63', hero: '#0a2f63', cta: '#0a2f63' },
    neon: { bg: '#05060a', ink: '#f5f7ff', mute: '#8d95ab', acc: '#22d3ee', acc2: '#f472b6', line: '#1d2233', band: '#0d1020', hero: '#0b0f1c', cta: '#0d1020' },
    datasheet: { bg: '#ffffff', ink: '#111111', mute: '#5f6368', acc: '#111111', line: '#d0d4da', band: '#f2f3f5', hero: '#111111' },
    editorial: { bg: '#fbf8f3', ink: '#1a1a1a', mute: '#6f6a62', acc: '#a4372a', line: '#e3dccf', band: '#f1ebe0', hero: '#1a1a1a' },
    luxe: { bg: '#0c0a07', ink: '#f5ecd7', mute: '#b9a984', acc: '#d4af37', line: '#3a3020', band: '#17130b', hero: '#17130b', cta: '#17130b' },
    gauge: { bg: '#0f172a', ink: '#e2e8f0', mute: '#94a3b8', acc: '#84cc16', line: '#23304d', band: '#111c33', hero: '#111c33', cta: '#111c33' },
  }[design];
  const maxQty = Math.max(1, ...groups.flatMap(g => g.items.map(i => Number(i.qty) || 0)));
  const specList = i => String(i.specs || '').split(/\s*·\s*/).filter(Boolean);
  const gaugeBar = i => {   // EV Dashboard: how full the shelf is, as a bar (relative to the catalog's best-stocked item)
    const q = Math.max(0, Number(i.qty) || 0), pct = Math.max(q ? 6 : 0, Math.round(q / maxQty * 100));
    const col = q <= 0 ? '#ef4444' : q <= 5 ? '#f59e0b' : '#84cc16';
    return '<table class="gauge"><tr><td class="gl">STOCK</td><td class="gt"><table class="gb"><tr><td style="width:' + pct + '%;background:' + col + '"></td><td></td></tr></table></td><td class="gq">' + q + '</td></tr></table>';
  };
  const price = i => !o.showCost ? '' : i.cost != null ? '<span class="pr">' + usd(i.cost) + '</span><span class="ea"> ' + (o.sales ? 'each' : 'ea') + '</span>'
    : o.sales ? '<span class="call">Call for pricing</span>' : '';
  const badge = i => (i.fav ? '<span class="bd fav">★ Ordered before</span> ' : '') + (!o.showQty ? (o.sales ? '<span class="bd ok">In stock</span>' : '')
    : o.sales ? (i.qty <= 0 ? '<span class="bd out">Ask about availability</span>' : i.qty <= 5 ? '<span class="bd low">Only ' + i.qty + ' left!</span>' : '<span class="bd ok">' + i.qty + ' in stock</span>')
    : '<span class="bd qty">Qty ' + i.qty + '</span>');
  const img = (i, cls) => !o.showPhotos ? '' : photos[i.photoId] ? '<img class="' + cls + '" src="' + photos[i.photoId] + '">'
    : '<div class="' + cls + ' none">' + (o.sales ? 'Photo coming soon' : 'No photo') + '</div>';
  const sub = i => [i.manufacturer && (o.sales ? 'by ' : '') + i.manufacturer, i.specs, i.partNumber && 'Part # ' + i.partNumber].filter(Boolean).map(t => '<div class="sub">' + e(t) + '</div>').join('');
  const rows = (list, n, cell) => Array.from({ length: Math.ceil(list.length / n) }, (_, r) => '<tr>' +
    Array.from({ length: n }, (_, k) => list[r * n + k] ? cell(list[r * n + k]) : '<td class="cell empty"></td>').join('') + '</tr>').join('');
  const groupTitle = g => e(g.name) + ' <span class="gn">' + g.items.length + (o.sales ? ' product' + (g.items.length === 1 ? '' : 's') + ' available' : ' item' + (g.items.length === 1 ? '' : 's')) + '</span>';
  let body = '';
  if (design === 'gallery') body = groups.map(g => '<h2>' + groupTitle(g) + '</h2><table class="grid">' + rows(g.items, 3, i => '<td class="cell"><div class="card">' +
      img(i, 'ph') + '<div class="tx">' + badge(i) + '<div class="nm">' + e(i.name) + '</div>' + sub(i) + '<div class="pp">' + price(i) + '</div></div></div></td>') + '</table>').join('');
  if (design === 'showroom') body = groups.map(g => '<h2>' + groupTitle(g) + '</h2><table class="list">' + g.items.map(i => '<tr>' + (o.showPhotos ? '<td class="phc">' + img(i, 'ph') + '</td>' : '') +
      '<td class="tx">' + badge(i) + '<div class="nm">' + e(i.name) + '</div>' + sub(i) + (i.dims ? '<div class="sub">' + (o.sales ? 'Ships in ' : 'Box ') + e(i.dims) + '</div>' : '') +
      '</td><td class="prc">' + price(i) + '</td></tr>').join('') + '</table>').join('');
  if (design === 'night') body = groups.map(g => '<h2><span class="glow">◆</span> ' + groupTitle(g) + '</h2><table class="grid">' + rows(g.items, 2, i => '<td class="cell"><table class="in"><tr>' +
      (o.showPhotos ? '<td class="phc">' + img(i, 'ph') + '</td>' : '') + '<td class="tx">' + badge(i) + '<div class="nm">' + e(i.name) + '</div>' + sub(i) + '<div class="pp">' + price(i) + '</div></td></tr></table></td>') + '</table>').join('');
  if (design === 'pricelist') body = '<table class="pl"><tr class="hd">' + (o.showPhotos ? '<th></th>' : '') + '<th>Item</th><th>Specs</th>' + (o.showQty ? '<th class="r">' + (o.sales ? 'Stock' : 'Qty') + '</th>' : '') +
      (o.showCost ? '<th class="r">Price</th>' : '') + '</tr>' + groups.map(g => '<tr class="gr"><td colspan="6">' + groupTitle(g) + '</td></tr>' + g.items.map((i, k) => '<tr class="' + (k % 2 ? 'alt' : '') + '">' +
      (o.showPhotos ? '<td class="phc">' + img(i, 'ph') + '</td>' : '') + '<td><div class="nm">' + (i.fav ? '★ ' : '') + e(i.name) + '</div><div class="sub">' + e([i.partNumber && 'Part # ' + i.partNumber, i.manufacturer].filter(Boolean).join(' · ')) + '</div></td>' +
      '<td class="sub">' + e(i.specs || '') + '</td>' + (o.showQty ? '<td class="r">' + (o.sales && i.qty <= 5 && i.qty > 0 ? '<b class="low">' + i.qty + '</b>' : i.qty) + '</td>' : '') +
      (o.showCost ? '<td class="r">' + (i.cost != null ? '<b>' + usd(i.cost) + '</b>' : o.sales ? '<span class="call">Call</span>' : '') + '</td>' : '') + '</tr>').join('')).join('') + '</table>';
  if (design === 'brochure') body = groups.map(g => '<div class="band">' + groupTitle(g) + '</div><table class="grid">' + rows(g.items, 2, i => '<td class="cell"><div class="card">' +
      img(i, 'ph') + '<div class="tx">' + badge(i) + '<div class="nm">' + e(i.name) + '</div>' + sub(i) + '<div class="pp">' + price(i) + '</div></div></div></td>') + '</table>').join('');
  if (design === 'blueprint') body = groups.map((g, gi) => '<h2>SECTION ' + String.fromCharCode(65 + gi % 26) + ' &mdash; ' + groupTitle(g) + '</h2><table class="grid">' + rows(g.items, 2, i => {
      const n = g.items.indexOf(i) + 1;
      return '<td class="cell"><div class="card"><table class="tb"><tr><td class="tl">ITEM ' + (n < 10 ? '0' : '') + n + '</td><td class="tr">DWG ' + e(i.partNumber || '—') + '</td></tr></table>' +
        img(i, 'ph') + '<div class="tx">' + badge(i) + '<div class="nm">' + e(i.name) + '</div>' + (i.manufacturer ? '<div class="sub">MFR: ' + e(i.manufacturer) + '</div>' : '') +
        specList(i).map(t => '<div class="note">&#9656; ' + e(t) + '</div>').join('') + (i.dims ? '<div class="note">&#9656; ' + e(i.dims) + '</div>' : '') +
        '<div class="pp">' + price(i) + '</div></div></div></td>'; }) + '</table>').join('');
  if (design === 'neon') body = groups.map(g => '<h2><span class="sl">&#47;&#47;</span> ' + groupTitle(g) + '</h2><table class="grid">' + rows(g.items, 3, i => '<td class="cell"><div class="card">' +
      img(i, 'ph') + '<div class="tx">' + badge(i) + '<div class="nm">' + e(i.name) + '</div>' + (i.partNumber ? '<div class="sub">' + e(i.partNumber) + '</div>' : '') +
      '<div class="chips">' + specList(i).map(t => '<span class="chip">' + e(t) + '</span>').join('') + '</div><div class="pp">' + price(i) + '</div></div></div></td>') + '</table>').join('');
  if (design === 'datasheet') body = groups.map(g => '<div class="dsh">' + groupTitle(g) + '</div>' + g.items.map(i => '<table class="ds"><tr>' +
      (o.showPhotos ? '<td class="phc">' + img(i, 'ph') + '</td>' : '') + '<td class="mid">' + (i.fav ? '<span class="bd fav">★ Ordered before</span>' : '') +
      '<div class="nm">' + e(i.name) + '</div>' + (i.partNumber ? '<div class="pn">' + e(i.partNumber) + '</div>' : '') +
      '<table class="kv">' + [i.manufacturer && ['Manufacturer', i.manufacturer]].concat(specList(i).map((t, k) => ['Spec ' + (k + 1), t])).concat([i.dims && ['Box', i.dims], i.weight && ['Weight', i.weight]])
        .filter(Boolean).map(([k, v]) => '<tr><td class="k">' + e(k) + '</td><td class="v">' + e(v) + '</td></tr>').join('') + '</table></td>' +
      '<td class="side">' + (o.showQty ? '<div class="sk">' + (o.sales ? (i.qty <= 0 ? 'Ask' : i.qty <= 5 ? 'Only ' + i.qty : i.qty) : i.qty) + '</div><div class="sl2">' + (o.sales ? 'in stock' : 'qty') + '</div>' : '') +
      '<div class="pp">' + price(i) + '</div></td></tr></table>').join('')).join('');
  if (design === 'editorial') body = groups.map((g, gi) => { const [lead, ...rest] = g.items;
      return '<div class="sec"><span class="no">' + (gi < 9 ? '0' : '') + (gi + 1) + '</span> ' + groupTitle(g) + '</div>' +
        '<table class="feat"><tr>' + (o.showPhotos ? '<td class="fph">' + img(lead, 'ph big') + '</td>' : '') + '<td class="ftx">' + badge(lead) + '<div class="kick">Featured</div><div class="nm big">' + e(lead.name) + '</div>' + sub(lead) + '<div class="pp">' + price(lead) + '</div></td></tr></table>' +
        (rest.length ? '<table class="grid">' + rows(rest, 3, i => '<td class="cell">' + img(i, 'ph') + '<div class="tx">' + badge(i) + '<div class="nm">' + e(i.name) + '</div>' + sub(i) + '<div class="pp">' + price(i) + '</div></div></td>') + '</table>' : ''); }).join('');
  if (design === 'luxe') body = groups.map(g => '<h2><span class="orn">&#10022;</span> ' + groupTitle(g) + ' <span class="orn">&#10022;</span></h2><table class="grid">' + rows(g.items, 3, i => '<td class="cell"><div class="card"><div class="inner">' +
      img(i, 'ph') + '<div class="tx">' + badge(i) + '<div class="nm">' + e(i.name) + '</div>' + sub(i) + '<div class="rule"></div><div class="pp">' + price(i) + '</div></div></div></div></td>') + '</table>').join('');
  if (design === 'gauge') body = groups.map(g => '<h2><span class="dot">&#9679;</span> ' + groupTitle(g) + '</h2><table class="grid">' + rows(g.items, 2, i => '<td class="cell"><table class="in"><tr>' +
      (o.showPhotos ? '<td class="phc">' + img(i, 'ph') + '</td>' : '') + '<td class="tx">' + (i.fav ? '<span class="bd fav">★ Ordered before</span>' : '') + '<div class="nm">' + e(i.name) + '</div>' +
      (i.partNumber ? '<div class="sub">' + e(i.partNumber) + '</div>' : '') + '<div class="chips">' + specList(i).map(t => '<span class="chip">&#9889; ' + e(t) + '</span>').join('') + '</div>' +
      (o.showQty ? gaugeBar(i) : '') + '<div class="pp">' + price(i) + '</div></td></tr></table></td>') + '</table>').join('');
  const intro = o.sales
    ? '<div class="hero"><div class="hl">' + e(o.headline || 'Quality equipment, in stock and ready to ship') + '</div><div class="tg">' + count + ' product' + (count === 1 ? '' : 's') + ' available now' +
      (favs ? ' · ★ ' + favs + ' you’ve ordered before' : '') + (valid ? ' · Prices good through ' + e(valid) : '') + ' · While supplies last</div></div>'
    : '<div class="hero inv"><div class="tg">' + count + ' item' + (count === 1 ? '' : 's') + ' in ' + groups.length + ' categor' + (groups.length === 1 ? 'y' : 'ies') +
      (o.showQty ? ' · ' + o.units + ' qty' : '') + (o.showCost ? ' · ' + usd(o.value) + (o.atList ? ' total value' : ' at cost') : '') + '</div></div>';
  const cta = o.sales ? '<div class="cta"><img class="sbot" src="data:image/png;base64,' + FAVICON_BOT_PNG_BASE64 + '"><div class="q">Ready to order?</div><div class="w">Contact <b>' + e(o.contact) +
    '</b> to reserve yours, ask about volume pricing or get a quote.</div><div class="f">' + (valid ? 'Prices valid through ' + e(valid) + '. ' : '') +
    'Quantities are limited and subject to prior sale. Prices exclude tax and shipping.</div></div>' : '';
  const heroBg = T.hero || (design === 'night' ? '#14532d' : design === 'showroom' ? '#0f172a' : T.acc);
  const dark = CATALOG_DARK.indexOf(design) !== -1;
  const css = 'body{font-family:Arial,Helvetica,sans-serif;color:' + T.ink + ';background:' + T.bg + ';font-size:10pt;margin:0}' +
    '.head{width:100%;border-collapse:collapse;margin-bottom:8px}.head td{vertical-align:middle}.logo{width:150px}.title{text-align:right}.title h1{margin:4px 0 0;font-size:22pt;color:' + T.ink + '}' +
    '.date{font-size:9pt;color:' + T.mute + '}.pf{font-size:10pt;margin-top:3px}.rib{display:inline-block;background:#dc2626;color:#fff;font-weight:bold;font-size:10pt;letter-spacing:3px;padding:3px 12px;border-radius:3px}' +
    '.hero{border-radius:8px;padding:11px 15px;margin:4px 0 10px;background:' + heroBg + ';color:#fff}.hero .hl{font-size:15pt;font-weight:bold;color:#ecfccb}.hero .tg{font-size:9.5pt;margin-top:2px;color:#f0fdf4}' +
    '.hero.inv{background:' + T.band + ';border-left:4px solid ' + T.acc + '}.hero.inv .tg{color:' + T.ink + '}' +
    'h2{font-size:12.5pt;color:' + T.acc + ';margin:14px 0 6px;padding-bottom:4px;border-bottom:2px solid ' + T.line + ';page-break-after:avoid}.gn{font-size:8.5pt;color:' + T.mute + ';font-weight:normal}' +
    '.bd{display:inline-block;font-size:7pt;font-weight:bold;padding:2px 6px;border-radius:8px;margin:0 3px 3px 0}.bd.ok{background:#dcfce7;color:#166534}.bd.low{background:#fee2e2;color:#b91c1c}' +
    '.bd.out{background:#f3f4f6;color:#4b5563}.bd.fav{background:#fef3c7;color:#92400e}.bd.qty{background:' + T.band + ';color:' + T.ink + '}' +
    '.nm{font-weight:bold;font-size:10.5pt;color:' + T.ink + '}.sub{font-size:8.3pt;color:' + T.mute + '}.pr{font-size:13pt;font-weight:bold;color:' + T.acc + '}.ea{font-size:8pt;color:' + T.mute + '}.call{font-weight:bold;color:#b45309}' +
    '.grid{width:100%;border-collapse:separate;border-spacing:8px}.cell{vertical-align:top;page-break-inside:avoid}.cell.empty{border:0;background:none}.in{width:100%;border-collapse:collapse}.in td{vertical-align:top}' +
    '.cta{margin-top:14px;border:2px solid ' + T.acc + ';border-radius:8px;padding:11px 15px;background:' + (T.cta || (design === 'night' ? '#111a2e' : '#f7fbf1')) + ';page-break-inside:avoid}' +
    (dark ? '.logo{background:#fff;padding:6px 10px;border-radius:8px}' : '') +
    '.cta .sbot{float:right;width:54px;height:54px;margin-left:10px;border-radius:12px}.cta .q{font-size:13pt;font-weight:bold;color:' + T.acc + '}.cta .w{margin-top:3px}' +
    '.cta .f{margin-top:5px;font-size:8pt;color:' + T.mute + '}.foot{margin-top:12px;color:' + T.mute + ';font-size:7.5pt}' +
    ({
      gallery: '.grid{border-spacing:10px}.cell{width:33%}.card{border:1px solid ' + T.line + ';border-radius:10px;padding:8px;background:#fff}' +
        '.ph{width:100%;height:150px;object-fit:contain;display:block;border-radius:6px;background:#fafafa}.ph.none{line-height:150px;text-align:center;color:#9ca3af;font-size:8pt;background:#f3f4f6}' +
        '.tx{padding:7px 2px 2px;text-align:center}.pp{margin-top:5px}',
      showroom: '.list{width:100%;border-collapse:collapse}.list td{vertical-align:middle;padding:12px 8px;border-bottom:1px solid ' + T.line + ';page-break-inside:avoid}.phc{width:200px}' +
        '.ph{width:190px;height:150px;object-fit:contain;border-radius:8px;background:#f8fafc;display:block}.ph.none{line-height:150px;text-align:center;color:#94a3b8;font-size:8pt}' +
        '.nm{font-size:14pt}.prc{text-align:right;width:150px}.pr{font-size:20pt}h2{color:' + T.ink + ';text-transform:uppercase;letter-spacing:2px;font-size:11pt;border-bottom:3px solid ' + T.acc + '}',
      night: '.cell{width:50%;background:#111a2e;border:1px solid ' + T.line + ';border-radius:10px;padding:9px}.phc{width:104px;padding-right:10px}' +
        '.ph{width:100px;height:100px;object-fit:contain;border-radius:8px;background:#0f172a;display:block}.ph.none{line-height:100px;text-align:center;color:#64748b;font-size:7.5pt}' +
        '.logo{background:#fff;padding:6px 10px;border-radius:8px}.glow{color:' + T.acc + '}h2{color:#ecfccb;border-bottom-color:#1f2a44}.pp{margin-top:6px}.bd.qty{background:#1f2a44;color:#e5e7eb}.foot{color:#64748b}',
      pricelist: '.pl{width:100%;border-collapse:collapse;font-size:9.3pt}.pl th{text-align:left;font-size:8pt;text-transform:uppercase;letter-spacing:1px;color:' + T.mute + ';border-bottom:2px solid ' + T.ink + ';padding:5px 6px}' +
        '.pl td{padding:5px 6px;border-bottom:1px solid ' + T.line + ';vertical-align:middle}.pl .r{text-align:right;white-space:nowrap}.pl tr.alt td{background:#fafafa}' +
        '.pl tr.gr td{background:' + T.band + ';font-weight:bold;color:' + T.acc + ';font-size:10pt;padding-top:9px}.phc{width:40px}.ph{width:36px;height:36px;object-fit:contain;display:block}' +
        '.ph.none{width:36px;height:36px;font-size:0;background:#f3f4f6;border-radius:4px}.nm{font-size:9.5pt}b.low{color:#b91c1c}',
      brochure: '.band{background:' + T.band + ';color:#fff;font-weight:bold;font-size:13pt;padding:8px 14px;border-radius:8px;margin:14px 0 4px;page-break-after:avoid}.band .gn{color:#ecfccb}' +
        '.cell{width:50%}.card{border:1px solid ' + T.line + ';border-top:4px solid ' + T.acc + ';border-radius:10px;background:#fff;overflow:hidden}' +
        '.ph{width:100%;height:170px;object-fit:contain;display:block;background:#f4f8ec}.ph.none{line-height:170px;text-align:center;color:#8aa070;font-size:8pt}.tx{padding:8px 11px 10px}.pp{margin-top:6px}',
      blueprint: 'body{font-family:"Courier New",Courier,monospace}.title h1{letter-spacing:2px;text-transform:uppercase}' +
        'h2{color:#fff;border-bottom:1px dashed ' + T.line + ';letter-spacing:2px;font-size:11pt}.gn{color:' + T.mute + '}' +
        '.cell{width:50%}.card{border:1.5px dashed #cfe0ff;padding:8px;background:#0f458c}.tb{width:100%;border-collapse:collapse;margin-bottom:6px;font-size:7.5pt;color:' + T.mute + '}' +
        '.tb td{border:1px solid ' + T.line + ';padding:2px 5px}.tr{text-align:right}.ph{width:100%;height:150px;object-fit:contain;display:block;background:#0b3770;border:1px solid ' + T.line + '}' +
        '.ph.none{line-height:150px;text-align:center;color:' + T.mute + ';font-size:8pt}.tx{padding-top:7px}.note{font-size:8.3pt;color:#dbe7ff}.nm{color:#fff;text-transform:uppercase;letter-spacing:.5px}' +
        '.pr{color:#fff}.bd.qty{background:#0a2f63;color:#dbe7ff;border:1px solid ' + T.line + '}.pp{margin-top:6px;border-top:1px dashed ' + T.line + ';padding-top:5px}.foot{color:' + T.mute + '}',
      neon: 'h2{color:#fff;text-transform:uppercase;letter-spacing:3px;font-size:11pt;border-bottom:1px solid ' + T.line + '}.sl{color:' + T.acc2 + '}.gn{color:' + T.mute + '}' +
        '.cell{width:33%}.card{background:#0b0e18;border-top:3px solid ' + T.acc + ';border-bottom:3px solid ' + T.acc2 + ';border-left:1px solid ' + T.line + ';border-right:1px solid ' + T.line + ';border-radius:6px;padding:8px}' +
        '.ph{width:100%;height:140px;object-fit:contain;display:block;background:#05060a;border-radius:4px}.ph.none{line-height:140px;text-align:center;color:#4b5268;font-size:8pt}' +
        '.tx{padding-top:7px}.chips{margin-top:4px}.chip{display:inline-block;font-size:7pt;color:' + T.acc + ';border:1px solid ' + T.acc + ';border-radius:10px;padding:1px 6px;margin:0 3px 3px 0}' +
        '.pr{color:' + T.acc + ';font-size:15pt}.pp{margin-top:5px}.bd.qty{background:#13182a;color:' + T.acc + '}.hero .hl{color:' + T.acc + '}.cta .q{color:' + T.acc2 + '}',
      datasheet: '.dsh{background:#111;color:#fff;font-weight:bold;font-size:11pt;letter-spacing:1px;text-transform:uppercase;padding:7px 12px;margin:14px 0 6px;page-break-after:avoid}.dsh .gn{color:#bbb}' +
        '.ds{width:100%;border-collapse:collapse;border-bottom:2px solid #111;margin-bottom:8px;page-break-inside:avoid}.ds td{vertical-align:top;padding:8px 6px}.phc{width:150px}' +
        '.ph{width:142px;height:120px;object-fit:contain;display:block;border:1px solid ' + T.line + ';background:#fafafa}.ph.none{line-height:120px;text-align:center;color:#9aa0a6;font-size:8pt}' +
        '.nm{font-size:13pt}.pn{display:inline-block;font-family:"Courier New",monospace;font-size:9pt;background:#111;color:#fff;padding:2px 7px;margin:3px 0 5px}' +
        '.kv{border-collapse:collapse;font-size:8.5pt;width:100%}.kv td{border-top:1px solid ' + T.line + ';padding:3px 6px}.k{color:' + T.mute + ';width:110px;text-transform:uppercase;font-size:7.5pt;letter-spacing:.5px}' +
        '.side{width:120px;text-align:right}.sk{font-size:22pt;font-weight:bold;line-height:1}.sl2{font-size:7.5pt;color:' + T.mute + ';text-transform:uppercase;letter-spacing:1px}.pr{color:#111;font-size:14pt}.pp{margin-top:8px}',
      editorial: 'body{font-family:Georgia,"Times New Roman",serif}.title h1{font-style:italic}.sec{font-size:20pt;margin:18px 0 8px;border-bottom:1px solid ' + T.ink + ';padding-bottom:4px;page-break-after:avoid}' +
        '.sec .no{color:' + T.acc + ';font-weight:bold}.sec .gn{font-family:Arial,sans-serif}.feat{width:100%;border-collapse:collapse;margin-bottom:6px;page-break-inside:avoid}.feat td{vertical-align:middle}' +
        '.fph{width:52%;padding-right:16px}.ph.big{width:100%;height:230px}.kick{font-family:Arial,sans-serif;font-size:7.5pt;letter-spacing:3px;text-transform:uppercase;color:' + T.acc + ';margin:4px 0}' +
        '.nm.big{font-size:20pt;line-height:1.15;font-weight:normal}.ftx .pr{font-size:18pt}.cell{width:33%;border-top:1px solid ' + T.line + ';padding-top:6px}.grid{border-spacing:12px 6px}' +
        '.ph{width:100%;height:130px;object-fit:contain;display:block;background:' + T.band + '}.ph.none{line-height:130px;text-align:center;color:' + T.mute + ';font-size:8pt}.tx{padding-top:6px}.pr{color:' + T.acc + '}.pp{margin-top:4px}',
      luxe: 'body{font-family:Georgia,"Times New Roman",serif}.title h1{letter-spacing:4px;text-transform:uppercase;color:' + T.acc + '}h2{text-align:center;color:' + T.acc + ';letter-spacing:4px;text-transform:uppercase;font-size:11pt;border-bottom:0}' +
        '.orn{color:' + T.acc + '}.gn{color:' + T.mute + '}.cell{width:33%}.card{border:3px double ' + T.acc + ';padding:3px;background:#120f09}.inner{border:1px solid #4a3d22;padding:8px;text-align:center}' +
        '.ph{width:100%;height:140px;object-fit:contain;display:block;background:#0c0a07}.ph.none{line-height:140px;text-align:center;color:#6f6346;font-size:8pt}' +
        '.nm{text-transform:uppercase;letter-spacing:1.5px;font-size:9.5pt;color:#f5ecd7;margin-top:6px}.rule{width:40px;height:1px;background:' + T.acc + ';margin:6px auto}.pr{color:' + T.acc + '}' +
        '.bd.qty{background:#2a2213;color:' + T.acc + '}.hero .hl{color:' + T.acc + '}.cta .q{color:' + T.acc + '}',
      gauge: 'h2{color:#fff;font-size:11.5pt;border-bottom:1px solid ' + T.line + '}.dot{color:' + T.acc + '}.gn{color:' + T.mute + '}' +
        '.cell{width:50%;background:#111c33;border:1px solid ' + T.line + ';border-left:4px solid ' + T.acc + ';border-radius:10px;padding:9px}.phc{width:104px;padding-right:10px}' +
        '.ph{width:100px;height:100px;object-fit:contain;border-radius:8px;background:#0b1222;display:block}.ph.none{line-height:100px;text-align:center;color:#475569;font-size:7.5pt}' +
        '.chips{margin-top:4px}.chip{display:inline-block;font-size:7pt;color:#d9f99d;background:#1a2a12;border-radius:10px;padding:1px 6px;margin:0 3px 3px 0}' +
        '.gauge{width:100%;border-collapse:collapse;margin-top:5px;font-size:7pt;color:' + T.mute + '}.gl{width:38px;letter-spacing:1px}.gq{width:30px;text-align:right;color:#fff;font-weight:bold;font-size:8.5pt}' +
        '.gb{width:100%;border-collapse:collapse;height:8px;background:#0b1222;border:1px solid ' + T.line + '}.gb td{height:8px;padding:0}.pr{color:' + T.acc + '}.pp{margin-top:5px}.hero .hl{color:#d9f99d}',
    })[design];
  return '<html><head><meta charset="utf-8"><style>' + css + '</style></head><body>' +
    '<table class="head"><tr><td><img class="logo" src="' + PDF_LOGO + '"></td><td class="title">' + (o.sales ? '<span class="rib">SALE</span>' : '') + '<h1>' + e(title) + '</h1>' +
    (o.preparedFor ? '<div class="pf">Prepared for <b>' + e(o.preparedFor) + '</b></div>' : '<div class="date">' + e(today) + '</div>') + '</td></tr></table>' +
    intro + body + (o.skipped ? '<div class="sub" style="margin-top:8px">' + (o.sales ? 'More photos available on request.' : o.skipped + ' photos were left out to keep the PDF a manageable size.') + '</div>' : '') +
    cta + '<div class="foot">' + e(today) + ' · EnviroSpark' + (o.sales ? '' : ' · made with Supplybot') + '</div></body></html>';
}

/** Builds the order's PDF, attaches it to the order and returns the order's PDF list plus the new file. */
function createOrderPdf(orderNo, withCosts, withPhotos) {
  requireRole_('Editor');
  return withLock_(() => {
    const order = readOrderRows_().find(o => o.orderNo === orderNo);
    if (!order) throw new Error('Order ' + orderNo + ' was not found.');
    const byKey = itemsByKey_();
    const lines = readLineRows_().filter(l => l.orderNo === orderNo).map(l => {
      const it = byKey[l.key], each = it ? weightLbs_(it.weight) : null;
      return Object.assign(l, { lbs: each == null ? null : each * l.qty, eachLbs: each, dims: it ? String(it.dimensions || '').replace(/\s+/g, ' ').trim() : '' });
    });
    const links = readOrderLinks_().filter(k => k.orderNo === orderNo);
    order.extraCosts = readOrderCosts_().filter(c => c.orderNo === orderNo);
    order.extraCost = roundMoney_(order.extraCosts.reduce((n, c) => n + c.amount, 0));
    const photos = { list: [], skipped: 0 };
    if (withPhotos) {
      let budget = PDF_PHOTO_BUDGET_MB * 1024 * 1024;
      readOrderPhotos_().filter(p => p.orderNo === orderNo).forEach(p => {
        try {
          const blob = DriveApp.getFileById(p.fileId).getBlob();
          const size = blob.getBytes().length;
          if (size > budget) { photos.skipped++; return; }
          budget -= size;
          photos.list.push({ src: dataUrl_(blob), timestamp: p.timestamp, user: p.user });
        } catch (e) { /* photo deleted from Drive: leave it out */ }
      });
    }
    const html = orderPdfHtml_(order, lines, links, !!withCosts, photos);
    const project = [order.projectName, order.projectNo && '#' + order.projectNo].filter(String).join(' ');
    let name = (withCosts ? '' : 'Packing List ') + orderNo + (order.projectName ? ' - ' + order.projectName : '');
    name = name.replace(/[\\/:*?"<>|]+/g, ' ').trim().slice(0, 110) + '.pdf';
    const bytes = Utilities.newBlob(html, 'text/html', 'order.html').getAs('application/pdf').getBytes();
    const file = savePdf_(ORDER_DOCS_FOLDER, { bytes, name }, 'Made from order ' + orderNo + (project ? ' (' + project + ')' : '') + ' in Supplybot');
    orderDocsSheet_().appendRow([new Date(), currentUser_(), orderNo, asText_(project), asText_(name),
      Math.max(1, Math.round(bytes.length / 1024)), file.getId(), file.getUrl()]);
    log_('Order PDF made', { id: '', name: orderNo, partNumber: '', job: '', quantity: '' }, 0, name);
    return { docs: getOrderDocuments(orderNo), url: file.getUrl(), name };
  });
}

// ----- Photos attached to an order (packed pallet, damage, delivery…) -----

function readOrderPhotos_() {
  const sheet = orderPhotosSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const photoOrder = p => p.position > 0 ? p.position : 1e6 + p.row;
  return sheet.getRange(2, 1, last - 1, ORDER_PHOTO_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2, timestamp: fmtTime_(r[0]), user: String(r[1]), orderNo: String(r[2]), fileId: String(r[4]), position: Number(r[6]) || 0,
  })).filter(p => p.fileId).sort((a, b) => photoOrder(a) - photoOrder(b));
}

/** Saves the order an order's photos are shown in (gallery, viewer and PDF). fileIds = every photo, in the new order. */
function reorderOrderPhotos(orderNo, fileIds) {
  requireRole_('Editor');
  requireOrderOwner_(orderNo);
  fileIds = (fileIds || []).map(String);
  return withLock_(() => {
    const mine = readOrderPhotos_().filter(p => p.orderNo === orderNo);
    if (mine.length !== fileIds.length || !mine.every(p => fileIds.includes(p.fileId))) throw new Error('The photos changed. Refresh and try again.');
    const sheet = orderPhotosSheet_(), col = ORDER_PHOTO_HEADERS.indexOf('Position') + 1;
    mine.forEach(p => sheet.getRange(p.row, col).setValue(fileIds.indexOf(p.fileId) + 1));
    return getOrderPhotos(orderNo);
  });
}

/** Photos on one order, in their arranged order (oldest first until arranged): [{ fileId, user, timestamp }]. */
function getOrderPhotos(orderNo) {
  requireRole_('Viewer');
  return readOrderPhotos_().filter(p => p.orderNo === orderNo).map(p => ({ fileId: p.fileId, user: p.user, timestamp: p.timestamp }));
}

/** Saves a photo (JPEG data URL, already shrunk on the phone) to the "Order Photos" Drive folder and attaches it to the order. */
function uploadOrderPhoto(orderNo, dataUrl) {
  requireRole_('Editor');
  requireOrderOwner_(orderNo);
  const m = /^data:(image\/[\w+.-]+);base64,(.+)$/.exec(dataUrl || '');
  if (!m) throw new Error('Invalid image.');
  return withLock_(() => {
    const order = readOrderRows_().find(o => o.orderNo === orderNo);
    if (!order) throw new Error('Order ' + orderNo + ' was not found.');
    const mine = readOrderPhotos_().filter(p => p.orderNo === orderNo), count = mine.length;
    const arranged = mine.some(p => p.position > 0);   // once arranged, new photos go at the end of that order
    if (count >= MAX_ORDER_PHOTOS) throw new Error('An order can have up to ' + MAX_ORDER_PHOTOS + ' photos.');
    const it = DriveApp.getFoldersByName(ORDER_PHOTOS_FOLDER);
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder(ORDER_PHOTOS_FOLDER);
    const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], orderNo + ' photo ' + (count + 1) + '.jpg'));
    const project = [order.projectName, order.projectNo && '#' + order.projectNo].filter(String).join(' ');
    orderPhotosSheet_().appendRow([new Date(), currentUser_(), orderNo, asText_(project), file.getId(), file.getUrl(),
      arranged ? Math.max(...mine.map(p => p.position)) + 1 : '']);
    return getOrderPhotos(orderNo);
  });
}

/** Removes a photo from its order; the file goes to Drive's trash (restorable for 30 days). */
function deleteOrderPhoto(fileId) {
  requireRole_('Editor');
  return withLock_(() => {
    const p = readOrderPhotos_().find(x => x.fileId === fileId);
    if (!p) throw new Error('That photo was already removed.');
    requireOrderOwner_(p.orderNo);
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* already gone */ }
    orderPhotosSheet_().deleteRow(p.row);
    return getOrderPhotos(p.orderNo);
  });
}

/** Serves an order photo as a data URL; only files listed in the Order Photos tab (the app runs as its owner). */
function getOrderPhoto(fileId) {
  requireRole_('Viewer');
  if (!readOrderPhotos_().some(p => p.fileId === fileId)) throw new Error('Photo not found.');
  return dataUrl_(DriveApp.getFileById(fileId).getBlob());
}

/** Removes a PDF from its order; the file goes to Drive's trash (restorable for 30 days). */
function deleteOrderDocument(fileId) {
  requireRole_('Editor');
  return withLock_(() => {
    const sheet = orderDocsSheet_();
    const d = readOrderDocs_().find(x => x.fileId === fileId);
    if (!d) throw new Error('That document was already removed.');
    requireOrderOwner_(d.orderNo);
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* already gone */ }
    sheet.deleteRow(d.row);
    return getOrderDocuments(d.orderNo);
  });
}

function setOrderStatus_(order, status) {
  const sheet = ordersSheet_();
  const H = h => ORDER_HEADERS.indexOf(h) + 1;
  sheet.getRange(order.row, H('Status')).setValue(status);
  sheet.getRange(order.row, H(status), 1, 2).setValues([[new Date(), currentUser_()]]);
}

/** After an item's name, part # or job is edited, re-point its order lines at the new key. */
function moveOrderLines_(oldKey, item) {
  const sheet = linesSheet_();
  readLineRows_().filter(l => l.key === oldKey).forEach(l => sheet.getRange(l.row, 8).setValue(asText_(item.key)));
}

// ---------- Comment notifications ----------
// Each person's "last checked" time is kept in Script Properties (the app runs as its owner, so per-user
// properties would all belong to the owner). A comment is new for you if someone else posted it after that time.

const SEEN_PREFIX = 'commentsSeen:';
const NOTIFY_PAGE = 50;

/**
 * Every comment on every item, newest first, one page at a time: { unread, total, comments }.
 * unread counts only other people's comments since you last checked; each comment says isNew / isMine.
 */
/**
 * All comments, newest first, a page at a time. filter (My profile): 'mine' (or true) = only the current person's own
 * comments; 'mentions' = only comments that @mention them. A mention is stored in the text as @{email}.
 */
function getNotifications(offset, filter) {
  requireRole_('Viewer');
  offset = Math.max(0, Number(offset) || 0);
  const me = currentUser_();
  const seen = PropertiesService.getScriptProperties().getProperty(SEEN_PREFIX + me);
  // First time: count the last 7 days as new, so recent activity shows up.
  const since = seen ? new Date(seen) : new Date(Date.now() - 7 * 24 * 3600 * 1000);
  const sheet = commentsSheet_();
  const last = sheet.getLastRow();
  const itemRows = last < 2 ? [] : sheet.getRange(2, 1, last - 1, COMMENT_HEADERS.length).getValues()
    .filter(r => r[7])
    .map(r => {
      const mine = String(r[1]) === me;
      return {
        id: String(r[7]),
        timestamp: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
        user: String(r[1]), itemName: String(r[2]), job: String(r[4]), text: String(r[5]), key: String(r[6]),
        isMine: mine, mentionsMe: mentions_(r[5], me),
        isNew: !mine && r[0] instanceof Date && r[0] > since,
      };
    });
  const orderRows = readOrderComments_().map(c => ({
    id: c.id, timestamp: c.timestamp, user: c.user, itemName: c.orderNo + (c.project ? ' · ' + c.project : ''), job: '', text: c.text,
    orderNo: c.orderNo, kind: 'order', isMine: c.user === me, mentionsMe: mentions_(c.text, me), isNew: c.user !== me && !!c.at && c.at > since.getTime(),
  }));
  const all = itemRows.concat(orderRows, noteMentionRows_(me, since)).sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp)));
  const rows = filter === 'mentions' ? all.filter(c => c.mentionsMe) : filter ? all.filter(c => c.isMine) : all;
  const page = rows.slice(offset, offset + NOTIFY_PAGE);
  const rx = reactionsByComment_();
  page.forEach(c => c.reactions = rx[c.id] || []);
  return {
    unread: all.filter(c => c.isNew).length,
    mineTotal: all.filter(c => c.isMine).length, mentionTotal: all.filter(c => c.mentionsMe).length,
    total: rows.length,
    comments: page,
  };
}

/** True when a comment's text @mentions this email. */
function mentions_(text, email) {
  return !!email && String(text).toLowerCase().indexOf('@{' + String(email).toLowerCase() + '}') !== -1;
}

/** Marks every comment up to now as seen for the current person. */
function markCommentsSeen() {
  requireRole_('Viewer');
  PropertiesService.getScriptProperties().setProperty(SEEN_PREFIX + currentUser_(), new Date().toISOString());
  return true;
}

/** After an item's name, part # or job is edited, re-point its comments at the new key. */
function moveComments_(oldKey, item) {
  const sheet = commentsSheet_();
  readComments_(sheet).filter(c => c.key === oldKey).forEach(c => {
    sheet.getRange(c.row, 3, 1, 3).setValues([[asText_(item.name), asText_(item.partNumber), asText_(item.job)]]);
    sheet.getRange(c.row, 7).setValue(asText_(item.key));
  });
}

function getLog(limit) {
  requireRole_('Viewer');
  requireScreen_('log');
  const sheet = logSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const n = Math.min(limit || 100, lastRow - 1);
  return sheet.getRange(lastRow - n + 1, 1, n, LOG_HEADERS.length).getValues()
    .reverse()
    .map(r => ({
      timestamp: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      user: String(r[1]), action: String(r[2]), name: String(r[4]), job: String(r[6]),
      change: Number(r[7]) || 0, newQuantity: Number(r[8]) || 0, note: String(r[9]),
    }));
}

// ---------- Live stock while building an order ----------
// The order form asks every 10 seconds (getLiveStock) for the latest stock of the items on it, what other Draft orders
// claim, what other people changed, and who has one of those items open in Edit item (markEditing, pinged every 20 s).
// "Who is editing" lives in the script cache, not the spreadsheet: one entry per person that expires a minute after
// their last ping (so a phone that just closes cleans up by itself), plus a list of who has pinged lately.

const EDIT_TTL_S = 60;
const EDITORS_KEY = 'editing:people';
const cacheJson_ = (cache, key, fallback) => { try { return JSON.parse(cache.get(key)) ?? fallback; } catch (e) { return fallback; } };

/** Marks the signed-in person as editing an item (key); done = true when they close Edit item. */
function markEditing(key, done) {
  requireRole_('Editor');
  key = String(key || '');
  const cache = CacheService.getScriptCache(), me = currentUser_().toLowerCase(), mine = 'editing:' + me;
  if (done) {
    const cur = cacheJson_(cache, mine, null);
    if (cur && cur.key === key) cache.remove(mine);
    return true;
  }
  if (!key) return false;
  cache.put(mine, JSON.stringify({ key, at: Date.now() }), EDIT_TTL_S);
  const people = cacheJson_(cache, EDITORS_KEY, []);
  cache.put(EDITORS_KEY, JSON.stringify(people.includes(me) ? people : people.concat(me).slice(-200)), 6 * 3600);
  return true;
}

/**
 * Fresh numbers for the items on an order being built. keys: their item keys; since: the `now` from the previous call
 * (null the first time); orderNo: the order's own number when editing a draft (its own lines don't count as claims).
 * Returns { now, items: { key: { key, id, quantity } }, claims: { key: [{ orderNo, qty, by }] },
 *           editing: [{ key, user }], changes: [{ timestamp, user, action, change, name, newQty }] (newest first) }.
 */
function getLiveStock(keys, since, orderNo) {
  requireRole_('Viewer');
  const want = new Set((Array.isArray(keys) ? keys : []).slice(0, 300).map(String));
  const me = currentUser_().toLowerCase(), now = new Date();

  // Stock: one read of the master sheet
  const sheet = master_(), cols = columns_(sheet), last = lastItemRow_(sheet, cols), items = {};
  if (last >= 2 && want.size) {
    const all = readRows_(sheet, 2, last - 1);
    all.values.forEach((row, i) => {
      const it = rowToItem_(row, i + 2, cols, all.shown[i]);
      if (want.has(it.key) && !items[it.key]) items[it.key] = { key: it.key, id: it.id, quantity: it.quantity };
    });
  }

  // What other Draft orders already claim
  const drafts = {};
  readOrderRows_().forEach(o => { if (o.status === 'Draft' && o.orderNo !== String(orderNo || '')) drafts[o.orderNo] = o.createdBy; });
  const claims = {};
  readLineRows_().forEach(l => {
    if (want.has(l.key) && l.orderNo in drafts) (claims[l.key] = claims[l.key] || []).push({ orderNo: l.orderNo, qty: l.qty, by: drafts[l.orderNo] });
  });

  // Who else has one of these items open in Edit item
  const cache = CacheService.getScriptCache();
  const people = cacheJson_(cache, EDITORS_KEY, []).filter(p => p !== me);
  const marks = people.length ? cache.getAll(people.map(p => 'editing:' + p)) : {};
  const editing = [];
  people.forEach(p => {
    let m = null;
    try { m = JSON.parse(marks['editing:' + p] || 'null'); } catch (e) { /* ignore a bad entry */ }
    if (m && want.has(m.key)) editing.push({ key: m.key, user: p });
  });

  // What other people changed on these items since the last check (from the App Log)
  const changes = [], t0 = since ? new Date(since).getTime() : NaN;
  if (isFinite(t0) && want.size) {
    const log = logSheet_(), lastRow = log.getLastRow(), n = Math.min(300, lastRow - 1);
    if (n > 0) log.getRange(lastRow - n + 1, 1, n, LOG_HEADERS.length).getValues().reverse().forEach(r => {
      const at = r[0] instanceof Date ? r[0].getTime() : new Date(r[0]).getTime();
      if (!(at > t0) || String(r[1]).toLowerCase() === me) return;
      if (!want.has(keyOf_({ name: String(r[4]), partNumber: String(r[5]), job: String(r[6]) }))) return;
      changes.push({ timestamp: new Date(at).toISOString(), user: String(r[1]), action: String(r[2]), change: Number(r[7]) || 0,
        name: String(r[4]), newQty: Number(r[8]) || 0 });
    });
  }
  return { now: now.toISOString(), items, claims, editing, changes: changes.slice(0, 20) };
}

// ---------- Send PM a note (HubSpot) ----------
// On a shipped order, "📝 Send PM a note" builds a note about what shipped (getPmNote) and shows it in the app, where it
// can be edited. Nothing goes to HubSpot until "Post to HubSpot" (postPmNote) adds it as a note on the order's ticket.
// The ticket comes from the order's HubSpot ticket link (added by "Build order"), or "ID 123…" in the order's notes.
// Setup, once, by the app's owner: in HubSpot create a private app with the "tickets" scope, copy its access token, and
// in Apps Script add it under Project Settings → Script properties as HUBSPOT_TOKEN. The token never reaches the phones.
const HUBSPOT_API = 'https://api.hubapi.com';
const NOTE_TO_TICKET = 228;   // HubSpot's association type id: note → ticket
const MAX_NOTE = 20000;

const hubspotToken_ = () => String(PropertiesService.getScriptProperties().getProperty('HUBSPOT_TOKEN') || '').trim();

/** One HubSpot API call; throws a readable error for anything but success. */
function hubspot_(method, path, body) {
  const token = hubspotToken_();
  if (!token) throw new Error('HubSpot isn’t connected yet. The app’s owner needs to add HUBSPOT_TOKEN in Apps Script (see SETUP.md).');
  const opts = { method, headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true };
  if (body) Object.assign(opts, { contentType: 'application/json', payload: JSON.stringify(body) });
  const res = UrlFetchApp.fetch(HUBSPOT_API + path, opts);
  const code = res.getResponseCode();
  let data = null;
  try { data = JSON.parse(res.getContentText() || 'null'); } catch (e) { /* not JSON */ }
  if (code === 401 || code === 403) throw new Error('HubSpot refused the app’s token (' + code + '). Check HUBSPOT_TOKEN and that its private app has the tickets scope.'
    + (data && data.message ? ' HubSpot says: ' + data.message : ''));
  if (code === 404) throw new Error('HubSpot couldn’t find that ticket.');
  if (code >= 300) throw new Error('HubSpot error ' + code + (data && data.message ? ': ' + data.message : '.'));
  return data;
}

/** The HubSpot ticket ID for an order: from its ticket link, else "ID 123…" in its notes; '' when there's none. */
function ticketIdOf_(o) {
  for (const k of (o.links || [])) {
    const m = /\/record\/0-5\/(\d+)|\/ticket\/(\d+)/.exec(k.url || '');
    if (m) return m[1] || m[2];
  }
  const n = /\bID (\d{4,})\b/.exec(o.notes || '');
  return n ? n[1] : '';
}

/** "marlan.sandoval@…" -> "Marlan Sandoval". */
const displayName_ = email => String(email || '').split('@')[0].split(/[._-]+/).filter(Boolean)
  .map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');

/** The note's text: what shipped, where it's going, dates, tracking, every item with its endpoints and their details. */
function pmNoteText_(o) {
  const tz = Session.getScriptTimeZone();
  const day = v => { if (!v) return ''; const d = /^\d{4}-\d{2}-\d{2}$/.test(v) ? new Date(v + 'T12:00:00') : new Date(v); return isNaN(d) ? String(v) : Utilities.formatDate(d, tz, 'MMM d, yyyy'); };
  const job = j => j && !/^available$/i.test(j) ? j : '';
  const units = o.lines.reduce((n, l) => n + l.qty, 0);
  const out = [];
  out.push('📦 Order ' + o.orderNo + ' has shipped' + (o.projectName ? ' – ' + o.projectName : '') + (o.projectNo ? ' (#' + o.projectNo + ')' : ''));
  out.push('Shipped ' + day(o.shipped) + (o.shippedBy ? ' by ' + displayName_(o.shippedBy) : ''));
  if (o.shipDate) out.push('Pickup / ship date: ' + day(o.shipDate));
  if (o.deliveryDate) out.push('Expected delivery: ' + day(o.deliveryDate));
  if (o.toLocation) out.push('To: ' + o.toLocation);
  if (o.shipTo) out.push('Ship to: ' + o.shipTo.replace(/\n+/g, ', '));
  if (o.contact) out.push('Point of contact: ' + o.contact);
  if (o.freight) out.push('Freight: ' + o.freight);
  if (o.trackingUrl) out.push('Tracking: ' + o.trackingUrl);
  out.push('');
  out.push('Items shipped (' + units + ' unit' + (units === 1 ? '' : 's') + '):');
  o.lines.forEach(l => {
    out.push('• ' + l.qty + ' × ' + l.name + (l.partNumber ? ' (' + l.partNumber + ')' : '') + (job(l.job) ? ' · ' + job(l.job) : ''));
    if ((l.endpoints || []).length) out.push('   Endpoints: ' + l.endpoints.map((p, i) => p + epDetails_(l, i)).join('; '));
  });
  const links = (o.links || []).filter(k => !/hubspot\.com/i.test(k.url) && k.url !== o.trackingUrl);
  if (links.length) { out.push(''); out.push('Links:'); links.forEach(k => out.push('• ' + (k.label ? k.label + ': ' : '') + k.url)); }
  if (o.notes) { out.push(''); out.push('Order notes: ' + o.notes.replace(/\n+/g, ' / ')); }
  out.push('');
  out.push('Sent from Supplybot by ' + displayName_(currentUser_()) + '.');
  return out.join('\n');
}

/**
 * The note for a shipped order, for the app to show (and let you edit) before posting:
 * { orderNo, ticketId, ticketName, ticketUrl, connected, text, sent }. ticketName is filled in when HubSpot is connected.
 */
function getPmNote(orderNo) {
  requireRole_('Editor');
  const o = getOrders().find(x => x.orderNo === orderNo);
  if (!o) throw new Error('Order ' + orderNo + ' was not found.');
  if (o.status !== 'Shipped') throw new Error('Only shipped orders can send the PM a note (' + o.orderNo + ' is ' + o.status + ').');
  const ticketId = ticketIdOf_(o), connected = !!hubspotToken_();
  let ticketName = '';
  if (ticketId && connected) {
    try { ticketName = String((hubspot_('get', '/crm/v3/objects/tickets/' + ticketId + '?properties=subject').properties || {}).subject || ''); }
    catch (e) { ticketName = ''; }
  }
  return { orderNo: o.orderNo, ticketId, ticketName, ticketUrl: ticketId ? HUBSPOT_TICKET_URL + ticketId : '', connected, text: pmNoteText_(o), sent: o.pmNote };
}

/** Posts text as a note on HubSpot ticket ticketId, then records it on the order ("PM Note Sent") and in the App Log. */
function postPmNote(orderNo, ticketId, text) {
  requireRole_('Editor');
  ticketId = String(ticketId || '').trim();
  const m = /\/record\/0-5\/(\d+)|\/ticket\/(\d+)/.exec(ticketId);
  if (m) ticketId = m[1] || m[2];   // a pasted ticket link works too
  if (!/^\d{3,}$/.test(ticketId)) throw new Error('Enter the HubSpot ticket ID (the number at the end of the ticket’s link).');
  text = String(text || '').replace(/\r/g, '').trim();
  if (!text) throw new Error('The note is empty.');
  if (text.length > MAX_NOTE) throw new Error('The note is too long (' + MAX_NOTE + ' characters max).');
  const o = readOrderRows_().find(x => x.orderNo === orderNo);
  if (!o) throw new Error('Order ' + orderNo + ' was not found.');
  if (o.status !== 'Shipped') throw new Error('Only shipped orders can send the PM a note.');
  hubspot_('get', '/crm/v3/objects/tickets/' + ticketId + '?properties=subject');   // a mistyped ticket number stops here
  const html = text.split('\n').map(line => escHtml_(line)).join('<br>');
  const note = hubspot_('post', '/crm/v3/objects/notes', {
    properties: { hs_timestamp: new Date().toISOString(), hs_note_body: html },
    associations: [{ to: { id: ticketId }, types: [{ associationCategory: 'HUBSPOT_DEFINED', associationTypeId: NOTE_TO_TICKET }] }],
  });
  const me = currentUser_(), when = new Date();
  const sent = Utilities.formatDate(when, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a') + ' · ' + me + ' · ticket ' + ticketId;
  withLock_(() => {
    const row = readOrderRows_().find(x => x.orderNo === orderNo);
    if (row) ordersSheet_().getRange(row.row, ORDER_HEADERS.indexOf('PM Note Sent') + 1).setValue(asText_(sent));
    logSheet_().appendRow([when, me, 'PM note sent', '', orderNo, '', '', 0, '', 'HubSpot ticket ' + ticketId]);
  });
  return { noteId: note && note.id ? String(note.id) : '', ticketId, ticketUrl: HUBSPOT_TICKET_URL + ticketId, sent };
}

// ---------- Log entry in tracker (Supply Chain 2026 → Shipping) ----------
// On a shipped order, "🧾 Log entry in tracker" lists one row per item for the Shipping tab of the Supply Chain 2026
// spreadsheet (getTrackerDraft), to check and change in the app. "Add to tracker" (logOrderInTracker) inserts them at the
// top of the tab, under the header and its blank row, followed by a blank row, the way rows are entered by hand.
// Only these columns are filled; the tab's other columns are left alone. QB Roll up gets the tab's own lookup formula.
// The app's owner needs edit access to that tab (it's protected: "Supply Chain Only").
const TRACKER_SHEET_ID = 1299167115;   // the "Shipping" tab (by ID, so renaming it is safe)
// [field, column header]; network = the line's endpoints
const TRACKER_COLS = [['audit', 'Audit Period'], ['date', 'Date'], ['projectName', 'Project Name'], ['projectNo', 'Project #'],
  ['activity', 'Activity type'], ['service', 'Service/Monta Migration'], ['from', 'From Location'], ['to', 'To Location'],
  ['contact', 'Point of Contact'], ['job', 'Original Job # (gear only)'], ['freight', 'Freight Service / Company'],
  ['vendor', 'Company / Vendor'], ['pm', 'PM'], ['notes', 'Notes'], ['product', 'Product'], ['qb', 'QB Roll up'],
  ['qty', 'Shipped Qty'], ['network', 'Network Code']];
const TRACKER_LINE_FIELDS = ['job', 'notes', 'product', 'qty', 'network'];
const TRACKER_MAX_TEXT = 500;
// First guesses for the tracker's Product, from the app's item name (the choice made in the app is remembered and wins).
// null = no guess from this list (e.g. "60 kw Cord Management" is not the 60 kW charger).
const TRACKER_ALIASES = [[/cord management|cable management|cords?\b/, null], [/\baw ?32\b/, 'PhiHong/eDRV 32A'],
  [/\bax ?48\b/, 'AX48'], [/\bax ?80\b/, 'AX80'], [/\b180 ?kw\b/, 'Zerova 180kw Fast Charger'], [/\b120 ?kw\b/, 'Zerova 120kw Fast Charger'],
  [/\b60 ?kw\b/, 'Zerova 60kw Fast Charger'], [/inteset|interset/, 'Interset Cable Retractors'], [/slim/, 'ES Dual Pedestal Slim'],
  [/lindquist|linquist/, 'Envirospark Linquist Pedestal w Cable Management'], [/green motion/, 'Eaton 40A Green Motion Charger'],
  [/pmd ?10/, 'PMD-10 Pedestal'], [/autel ac single/, 'Autel AC Single'], [/autel ac elite/, 'Autel AC Elite'],
  [/wheel ?stops?/, 'Wheel stops'], [/safety bollard/, '3\' Safety Bollard'], [/sc3|swtch 32|switch 32/, 'SWTCH 32A'],
  [/connectivity box/, 'Tesla Connectivity Box'], [/electron wi ?fi/, 'Electron Wifi Gateway'], [/tesla (3|gen ?3) wall/, 'Tesla Gen 3'],
  [/rivian/, 'Rivian lvl2 Charger'], [/sema ?connect/, 'Sema Connect Pedestal'], [/generic ev/, 'EV Sign Generic'],
  [/envirospark vertical/, 'Envirospark Sign vertical'], [/lightning bolt.*arrow|sign.*arrow/, 'Envirospark Parking Sign w/ Arrows'],
  [/\bhcs ?30\b/, 'HCS-30'], [/\bhcs ?40r\b/, 'HCS-40R'], [/\bhcs ?40\b/, 'HCS-40'], [/\bhcs ?80\b/, 'HCS-80'],
  [/clipper.*holster/, 'Clipper Holster'], [/lectron/, 'Lectron J Holster']];

const trackerKey_ = s => String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();

function trackerSheet_() {
  let ss;
  try { ss = SpreadsheetApp.openById(SC_SPREADSHEET_ID); }
  catch (e) { throw new Error('The app can’t open the Supply Chain 2026 spreadsheet. The app’s owner needs edit access to it.'); }
  const sheet = ss.getSheets().find(s => s.getSheetId() === TRACKER_SHEET_ID);
  if (!sheet) throw new Error('The Shipping tab wasn’t found in the Supply Chain 2026 spreadsheet.');
  return sheet;
}

/** { field: column number } for the tracker tab; throws naming any column that's missing. */
function trackerCols_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const cols = {}, missing = [];
  TRACKER_COLS.forEach(([k, h]) => { const i = findCol_(headers, h); if (i === -1) missing.push(h); else cols[k] = i + 1; });
  if (missing.length) throw new Error('The Shipping tab is missing the column' + (missing.length > 1 ? 's ' : ' ') + missing.join(', ') + '.');
  return cols;
}

/** The dropdown choices of a tracker column (from the first of the top rows that has one), or [] when it has none. */
function trackerChoices_(sheet, col) {
  for (let r = 2; r <= 12; r++) {
    const v = sheet.getRange(r, col).getDataValidation();
    if (v && v.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      const seen = new Set();
      return (v.getCriteriaValues()[0] || []).map(x => String(x)).filter(x => x.trim() && !seen.has(x) && seen.add(x));
    }
  }
  return [];
}

const trackerLearned_ = () => { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('TRACKER_PRODUCTS') || '{}'); } catch (e) { return {}; } };

/** Best guess at the tracker Product for an app item name: what was picked before, else TRACKER_ALIASES, else a product named inside it. */
function trackerProduct_(name, products, learned) {
  const exact = p => products.find(x => x === p) || products.find(x => trackerKey_(x) === trackerKey_(p)) || '';
  const key = trackerKey_(name);
  if (learned[key] !== undefined && (learned[key] === '' || exact(learned[key]))) return learned[key] && exact(learned[key]);
  const direct = products.find(p => trackerKey_(p) === key);
  if (direct) return direct;
  for (const [re, p] of TRACKER_ALIASES) if (re.test(key)) { if (p === null) break; const hit = exact(p); if (hit) return hit; }
  return products.filter(p => trackerKey_(p).length > 3 && (' ' + key + ' ').includes(' ' + trackerKey_(p) + ' '))
    .sort((a, b) => b.length - a.length)[0] || '';
}

/** "October 2026 Counts" for 'yyyy-MM-dd'. */
function trackerAudit_(day) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(day || '');
  return m ? Utilities.formatDate(new Date(+m[1], +m[2] - 1, +m[3], 12), Session.getScriptTimeZone(), 'MMMM yyyy') + ' Counts' : '';
}

/**
 * The tracker rows for a shipped order, for the app to show (and let you change) before adding them:
 * { orderNo, logged, url, head: { audit, date, projectName, ... }, lines: [{ name, partNumber, job, notes, product, qty, network }],
 *   products, activities, services } — the last three are the tab's dropdown choices ([] = type anything).
 */
function getTrackerDraft(orderNo) {
  requireRole_('Editor');
  const o = getOrders().find(x => x.orderNo === orderNo);
  if (!o) throw new Error('Order ' + orderNo + ' was not found.');
  if (o.status !== 'Shipped') throw new Error('Only shipped orders can be logged in the tracker (' + o.orderNo + ' is ' + o.status + ').');
  const sheet = trackerSheet_(), cols = trackerCols_(sheet);
  const products = trackerChoices_(sheet, cols.product), activities = trackerChoices_(sheet, cols.activity), services = trackerChoices_(sheet, cols.service);
  const tz = Session.getScriptTimeZone();
  const shipped = o.shipped ? new Date(o.shipped) : new Date();
  const date = Utilities.formatDate(isNaN(shipped) ? new Date() : shipped, tz, 'yyyy-MM-dd');
  const pick = (list, v) => list.find(x => x.toLowerCase() === String(v || '').trim().toLowerCase()) || '';
  const activity = pick(activities, o.activityType) || (/^deliver/i.test(o.activityType || '') || !o.activityType ? pick(activities, 'Shipping') : '')
    || (activities.length ? '' : o.activityType || '');
  const learned = trackerLearned_();
  return {
    orderNo: o.orderNo, logged: o.tracker || '', url: 'https://docs.google.com/spreadsheets/d/' + SC_SPREADSHEET_ID + '/edit#gid=' + TRACKER_SHEET_ID,
    products, activities, services,
    head: { audit: trackerAudit_(date), date, projectName: o.projectName, projectNo: o.projectNo || ticketIdOf_(o), activity, service: '',
      from: o.fromLocation, to: o.toLocation, contact: o.contact, freight: o.freight, vendor: o.vendor, pm: o.pm },
    lines: o.lines.map(l => ({ name: l.name, partNumber: l.partNumber, qty: l.qty, job: /^available$/i.test(l.job || '') ? '' : l.job || '',
      notes: l.qty + ' ' + l.name, product: trackerProduct_(l.name, products, learned), network: (l.endpoints || []).join(', ') })),
  };
}

/**
 * Inserts the order's rows at the top of the tracker's Shipping tab (head = the fields every row shares, lines = one per
 * item, as getTrackerDraft gave them, after any changes), then records it on the order ("Tracker Logged") and in the App Log.
 */
function logOrderInTracker(orderNo, head, lines) {
  requireRole_('Editor');
  const o = readOrderRows_().find(x => x.orderNo === orderNo);
  if (!o) throw new Error('Order ' + orderNo + ' was not found.');
  if (o.status !== 'Shipped') throw new Error('Only shipped orders can be logged in the tracker.');
  return writeTrackerRows_(head, lines, (n, logged, when, me) => {
    const row = readOrderRows_().find(x => x.orderNo === orderNo);
    if (row) ordersSheet_().getRange(row.row, ORDER_HEADERS.indexOf('Tracker Logged') + 1).setValue(asText_(logged));
    logSheet_().appendRow([when, me, 'Logged in tracker', '', orderNo, '', '', 0, '', n + ' row' + (n === 1 ? '' : 's') + ' on Supply Chain 2026 → Shipping']);
  });
}

/**
 * The tracker rows for inventory just returned (Receive → Returning inventory), like getTrackerDraft does for a shipped
 * order: adjIds = the "Returned" adjustments that were saved together. Activity type Receiving, dated the day it came
 * back, From where it came back from, To where its order shipped from, and the order's project, PM and company when the
 * return names one of the app's orders. Same shape as getTrackerDraft, plus adjIds.
 * A purchase order received (Receive → Purchase order, "Received from vendor") works the same way: Activity type
 * Receiving, Project Name "PO <#>", From and Company / Vendor = the vendor, and the PO # in each row's notes.
 */
function getReturnTrackerDraft(adjIds) {
  requireRole_('Editor');
  const adjs = receiptAdjs_(adjIds), first = adjs[0];
  if (first.type === RECEIPT_TYPES.po) return poTrackerDraft_(adjs);
  const o = first.orderNo ? getOrders().find(x => x.orderNo === first.orderNo) : null;
  const sheet = trackerSheet_(), cols = trackerCols_(sheet);
  const products = trackerChoices_(sheet, cols.product), activities = trackerChoices_(sheet, cols.activity), services = trackerChoices_(sheet, cols.service);
  const pick = (list, v) => list.find(x => x.toLowerCase() === String(v || '').trim().toLowerCase()) || '';
  const date = Utilities.formatDate(new Date(first.timestamp), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const learned = trackerLearned_();
  return {
    adjIds: adjs.map(a => a.id), orderNo: first.orderNo || '', logged: (adjs.find(a => a.tracker) || {}).tracker || '',
    url: 'https://docs.google.com/spreadsheets/d/' + SC_SPREADSHEET_ID + '/edit#gid=' + TRACKER_SHEET_ID, products, activities, services,
    head: { audit: trackerAudit_(date), date, projectName: (o && o.projectName) || first.from, projectNo: o ? o.projectNo || ticketIdOf_(o) : '',
      activity: pick(activities, 'Receiving') || (activities.length ? '' : 'Receiving'), service: '',
      from: first.from || (o && o.toLocation) || '', to: (o && o.fromLocation) || '', contact: first.contact || (o && o.contact) || '',
      freight: '', vendor: (o && o.vendor) || '', pm: (o && o.pm) || '' },
    lines: adjs.map(a => ({ name: a.name, partNumber: a.partNumber, qty: a.change, job: /^available$/i.test(a.job || '') ? '' : a.job || '',
      notes: a.change + ' ' + a.name + (a.reason ? ' (' + a.reason + ')' : ''), product: trackerProduct_(a.name, products, learned),
      network: (a.endpoints || []).join(', ') })),
  };
}

/** Inserts a return's rows at the top of the tracker's Shipping tab (see getReturnTrackerDraft) and notes it in the App Log. */
function logReturnInTracker(adjIds, head, lines) {
  requireRole_('Editor');
  const adjs = receiptAdjs_(adjIds), po = adjs[0].type === RECEIPT_TYPES.po;
  const src = po ? [adjs[0].po && 'PO ' + adjs[0].po, adjs[0].vendor].filter(Boolean).join(' · ') : [adjs[0].orderNo, adjs[0].from].filter(Boolean).join(' · ');
  return writeTrackerRows_(head, lines, (n, logged, when, me) => {
    const sheet = adjSheet_(), col = ADJ_HEADERS.indexOf('Tracker Logged') + 1;
    adjs.forEach(a => sheet.getRange(a.row, col).setValue(asText_(logged)));
    logSheet_().appendRow([when, me, po ? 'PO receipt logged in tracker' : 'Return logged in tracker', '', src, '', '', 0, '', n + ' row' + (n === 1 ? '' : 's') + ' on Supply Chain 2026 → Shipping']);
  });
}

const RECEIPT_TYPES = { return: 'Returned', po: 'Received from vendor' };

/** The adjustments of one return or PO receipt (saved together), for the tracker. */
function receiptAdjs_(adjIds) {
  const ids = (adjIds || []).map(String), types = Object.values(RECEIPT_TYPES);
  const adjs = readAdjustments_().filter(a => ids.includes(a.id) && types.includes(a.type));
  if (!adjs.length) throw new Error('That receipt wasn’t found. It may have been undone.');
  return adjs.filter(a => a.type === adjs[0].type);
}

/** getReturnTrackerDraft for a received purchase order. */
function poTrackerDraft_(adjs) {
  const first = adjs[0], sheet = trackerSheet_(), cols = trackerCols_(sheet);
  const products = trackerChoices_(sheet, cols.product), activities = trackerChoices_(sheet, cols.activity), services = trackerChoices_(sheet, cols.service);
  const pick = (list, v) => list.find(x => x.toLowerCase() === String(v || '').trim().toLowerCase()) || '';
  const date = Utilities.formatDate(new Date(first.timestamp), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const learned = trackerLearned_(), po = first.po ? 'PO ' + first.po : '';
  return {
    adjIds: adjs.map(a => a.id), orderNo: '', po: first.po || '', vendor: first.vendor || '', logged: (adjs.find(a => a.tracker) || {}).tracker || '',
    url: 'https://docs.google.com/spreadsheets/d/' + SC_SPREADSHEET_ID + '/edit#gid=' + TRACKER_SHEET_ID, products, activities, services,
    head: { audit: trackerAudit_(date), date, projectName: po || first.vendor || '', projectNo: '',
      activity: pick(activities, 'Receiving') || (activities.length ? '' : 'Receiving'), service: '',
      from: first.vendor || '', to: '', contact: '', freight: '', vendor: first.vendor || '', pm: '' },
    lines: adjs.map(a => ({ name: a.name, partNumber: a.partNumber, qty: a.change, job: /^available$/i.test(a.job || '') ? '' : a.job || '',
      notes: a.change + ' ' + a.name + (po ? ' (' + po + ')' : ''), product: trackerProduct_(a.name, products, learned),
      network: (a.endpoints || []).join(', ') })),
  };
}

/**
 * Every return received (Receive → Returning inventory), newest first, for the Past returns list: the "Returned"
 * adjustments saved together (next to each other in Inventory Adjustments, same person, source, order, reason, contact
 * and note, within
 * 20 seconds) make one return. tracker = when it was logged in Supply Chain 2026 → Shipping ('' = not yet).
 */
function getReturns() { return getReceipts('return'); }

/** The same for kind 'po': every purchase order received (Receive → Purchase order), one per save, with vendor and PO #. */
// 📎 PDFs and links on a PO receipt or a return (the adjustments saved together): kept on each of its rows (Links / PDFs
// columns, as JSON), PDFs in the Drive folder "PO & Return Documents", viewable by link inside the company like order PDFs.
const RECEIPT_DOCS_FOLDER = 'PO & Return Documents';
function receiptRows_(adjIds) {
  const ids = (adjIds || []).map(String), rows = readAdjustments_().filter(a => ids.indexOf(a.id) >= 0);
  if (!rows.length) throw new Error('That receipt was not found (it may have been changed in the sheet).');
  return rows;
}
function writeReceiptCol_(rows, header, value) {
  const sheet = adjSheet_(), col = ADJ_HEADERS.indexOf(header) + 1;
  rows.forEach(a => sheet.getRange(a.row, col).setValue(value));
}
/**
 * 🏷 Tag (or fix) a past PO receipt or return: its PM, vendor / company and point of contact, on every row saved with it.
 * tags: { pm, vendor, contact } — only the ones given change; '' clears one. Recorded in Activity.
 */
function setReceiptTags(adjIds, tags) {
  requireRole_('Editor');
  requireScreen_('stock');
  tags = tags || {};
  const clean = cleanAdjInfo_(tags), cols = { pm: 'PM', vendor: 'Vendor', contact: 'Point of Contact' };
  return withLock_(() => {
    const rows = receiptRows_(adjIds), a = rows[0], sheet = adjSheet_(), changed = [];
    Object.keys(cols).forEach(k => {
      if (tags[k] === undefined || clean[k] === a[k]) return;
      const col = ADJ_HEADERS.indexOf(cols[k]) + 1; if (!col) return;
      rows.forEach(r => sheet.getRange(r.row, col).setValue(asText_(clean[k])));
      changed.push((k === 'pm' ? 'PM' : k === 'vendor' ? 'Vendor' : 'Contact') + ': ' + (clean[k] || '(none)'));
    });
    if (changed.length) logSheet_().appendRow([new Date(), currentUser_(), a.type === RECEIPT_TYPES.po ? 'PO receipt tagged' : 'Return tagged', '',
      a.type === RECEIPT_TYPES.po ? 'PO ' + (a.po || '') : 'Return' + (a.from ? ' from ' + a.from : '') + (a.orderNo ? ' (' + a.orderNo + ')' : ''), '', '', 0, '', changed.join(' · ')]);
    return { pm: clean.pm, vendor: clean.vendor, contact: clean.contact, changed: changed.length };
  });
}
function setReceiptLinks(adjIds, links) {
  requireRole_('Editor');
  requireScreen_('stock');
  const seen = {}, clean = cleanNoteLinks_(links).filter(l => { const k = String(l.url).toLowerCase().replace(/[/]+$/, ''); return seen[k] ? false : (seen[k] = true); });   // the same link only once
  return withLock_(() => { const rows = receiptRows_(adjIds); writeReceiptCol_(rows, 'Links', clean.length ? JSON.stringify(clean) : ''); return clean; });
}
function addReceiptPdf(adjIds, fileName, dataUrl) {
  requireRole_('Editor');
  requireScreen_('stock');
  const pdf = parsePdf_(fileName, dataUrl);
  return withLock_(() => {
    const rows = receiptRows_(adjIds), a = rows[0];
    const what = a.type === RECEIPT_TYPES.po ? 'PO ' + (a.po || '') + (a.vendor ? ' from ' + a.vendor : '') : 'Return' + (a.from ? ' from ' + a.from : '') + (a.orderNo ? ' (' + a.orderNo + ')' : '');
    const kb = Math.max(1, Math.round(pdf.bytes.length / 1024));
    if ((a.pdfs || []).some(f => f.kind !== 'photo' && String(f.name).toLowerCase() === String(pdf.name).toLowerCase() && Math.abs((f.kb || 0) - kb) <= 1)) throw new Error(pdf.name + ' is already attached to this ' + (a.type === RECEIPT_TYPES.po ? 'PO' : 'return') + '.');
    const file = savePdf_(RECEIPT_DOCS_FOLDER, pdf, 'Attached to ' + what + ' from Supplybot');
    const list = (a.pdfs || []).concat([{ id: file.getId(), name: pdf.name, url: file.getUrl(), kb: Math.max(1, Math.round(pdf.bytes.length / 1024)) }]).slice(-20);
    writeReceiptCol_(rows, 'PDFs', JSON.stringify(list));
    return list;
  });
}
/** 📷 A photo on a PO receipt or a return (shrunk on the phone first), kept with its PDFs; thumb: a tiny preview to show in the list. */
function addReceiptPhoto(adjIds, fileName, dataUrl, thumb) {
  requireRole_('Editor');
  requireScreen_('stock');
  const m = /^data:(image\/(jpeg|png|webp));base64,(.+)$/.exec(String(dataUrl || ''));
  if (!m) throw new Error('Photos can be JPEG, PNG or WebP.');
  const bytes = Utilities.base64Decode(m[3]);
  if (bytes.length > 6 * 1024 * 1024) throw new Error('That photo is over 6 MB.');
  thumb = /^data:image\/(jpeg|webp);base64,[A-Za-z0-9+\/=]+$/.test(String(thumb || '')) && String(thumb).length < 12000 ? String(thumb) : '';
  const name = String(fileName || 'photo').replace(/[\\/:*?"<>|]+/g, ' ').trim().slice(0, 120) || 'photo';
  return withLock_(() => {
    const rows = receiptRows_(adjIds), a = rows[0], kb = Math.max(1, Math.round(bytes.length / 1024));
    if ((a.pdfs || []).some(f => f.kind === 'photo' && String(f.name).toLowerCase() === name.toLowerCase() && Math.abs((f.kb || 0) - kb) <= 2)) throw new Error(name + ' is already on this ' + (a.type === RECEIPT_TYPES.po ? 'PO' : 'return') + '.');
    const it = DriveApp.getFoldersByName(RECEIPT_DOCS_FOLDER), folder = it.hasNext() ? it.next() : DriveApp.createFolder(RECEIPT_DOCS_FOLDER);
    const file = folder.createFile(Utilities.newBlob(bytes, m[1], name));
    file.setDescription('Photo on ' + (a.type === RECEIPT_TYPES.po ? 'PO ' + (a.po || '') : 'a return') + ' from Supplybot');
    try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* link sharing not allowed */ }
    const list = (a.pdfs || []).concat([{ id: file.getId(), name, url: file.getUrl(), kb: Math.max(1, Math.round(bytes.length / 1024)), kind: 'photo', thumb }]).slice(-30);
    if (JSON.stringify(list).length > 45000) throw new Error('This receipt has too many attachments. Remove some first.');
    writeReceiptCol_(rows, 'PDFs', JSON.stringify(list));
    return list;
  });
}
function removeReceiptPdf(adjIds, fileId) {
  requireRole_('Editor');
  requireScreen_('stock');
  return withLock_(() => {
    const rows = receiptRows_(adjIds), list = (rows[0].pdfs || []).filter(f => f.id !== fileId);
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* already gone */ }
    writeReceiptCol_(rows, 'PDFs', list.length ? JSON.stringify(list) : '');
    return list;
  });
}
function getReceipts(kind) {
  requireRole_('Viewer');
  requireScreen_('stock');
  const type = RECEIPT_TYPES[kind];
  if (!type) throw new Error('Unknown kind of receipt.');
  return receiptsOf_(type).reverse().slice(0, 2000);
}

/**
 * The returns saved with this order # (Receive → Returning inventory → Order #), oldest first, for the order's page:
 * what came back from it after it shipped. Anyone who can see the order can see these.
 */
function getOrderReturns(orderNo) {
  requireRole_('Viewer');
  const no = String(orderNo || '').trim().toLowerCase();
  if (!no) return [];
  return receiptsOf_(RECEIPT_TYPES.return, a => a.orderNo.trim().toLowerCase() === no)
    .map(r => ({ timestamp: r.timestamp, user: r.user, from: r.from, reason: r.reason, contact: r.contact, note: r.note, lines: r.lines }));
}

/** Receipts of one adjustment type, oldest first: the adjustments saved together make one (see getReceipts). */
function receiptsOf_(type, keep) {
  const out = [];
  let cur = null, prev = null;
  readAdjustments_().forEach(a => {
    if (a.type !== type || (keep && !keep(a))) { prev = null; return; }
    const sig = [a.user, a.from, a.orderNo, a.reason, a.contact, a.note, a.vendor, a.po, a.pm].join('\u0001');
    if (!(cur && prev && prev.sig === sig && prev.row === a.row - 1 && Math.abs(new Date(a.timestamp) - new Date(prev.timestamp)) < 20000)) {
      cur = { adjIds: [], timestamp: a.timestamp, user: a.user, from: a.from, orderNo: a.orderNo, reason: a.reason, contact: a.contact,
        note: a.note, vendor: a.vendor, po: a.po, pm: a.pm, links: a.links || [], pdfs: a.pdfs || [], tracker: '', lines: [] };
      out.push(cur);
    }
    cur.adjIds.push(a.id);
    if (a.tracker) cur.tracker = a.tracker;
    cur.lines.push({ name: a.name, partNumber: a.partNumber, key: a.key, qty: a.change, endpoints: a.endpoints || [] });
    prev = { sig, row: a.row, timestamp: a.timestamp };
  });
  return out;
}

/**
 * Checks the rows and inserts them at the top of the tracker's Shipping tab (under the header and its blank row, then a
 * blank row), then calls record(n, logged, when, me) inside the same lock to note where they came from. Returns { logged, rows }.
 */
function writeTrackerRows_(head, lines, record) {
  head = head || {};
  lines = (Array.isArray(lines) ? lines : []).filter(l => l && Number(l.qty) > 0);
  if (!lines.length) throw new Error('There are no items with a quantity to log.');
  if (lines.length > 200) throw new Error('That’s too many rows for one entry.');
  const text = (v, label) => { const s = String(v ?? '').replace(/\r/g, '').trim(); if (s.length > TRACKER_MAX_TEXT) throw new Error(label + ' is too long.'); return s; };
  const day = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(head.date || ''));
  if (!day) throw new Error('Pick the date.');
  const sheet = trackerSheet_(), cols = trackerCols_(sheet);
  const choice = (list, v, label) => { v = text(v, label); if (v && list.length && !list.includes(v)) throw new Error('“' + v + '” isn’t one of the tracker’s ' + label + ' choices.'); return v; };
  const products = trackerChoices_(sheet, cols.product);
  const shared = {
    audit: text(head.audit, 'Audit Period'), date: new Date(+day[1], +day[2] - 1, +day[3]),
    projectName: text(head.projectName, 'Project Name'), activity: choice(trackerChoices_(sheet, cols.activity), head.activity, 'Activity type'),
    service: choice(trackerChoices_(sheet, cols.service), head.service, 'Service/Monta Migration'),
    from: text(head.from, 'From Location'), to: text(head.to, 'To Location'), contact: text(head.contact, 'Point of Contact'),
    freight: text(head.freight, 'Freight Service / Company'), vendor: text(head.vendor, 'Company / Vendor'), pm: text(head.pm, 'PM'),
  };
  const pno = text(head.projectNo, 'Project #');
  shared.projectNo = /^\d{1,15}$/.test(pno) ? Number(pno) : pno;   // a number, like the rows typed in (the tab's lookups match on it)
  const rows = lines.map(l => ({ job: text(l.job, 'Original Job #'), notes: text(l.notes, 'Notes'), product: choice(products, l.product, 'Product'),
    qty: Math.round(Number(l.qty)), network: text(l.network, 'Network Code'), name: String(l.name || '') }));
  const n = rows.length, width = Math.max(...Object.values(cols));
  const me = currentUser_(), when = new Date();
  const logged = Utilities.formatDate(when, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a') + ' · ' + me + ' · ' + n + ' row' + (n === 1 ? '' : 's');
  withLock_(() => {
    const lastCol = Math.max(sheet.getLastColumn(), width);
    const row2Blank = sheet.getRange(2, 1, 1, lastCol).getDisplayValues()[0].every(v => !String(v).trim());
    const at = row2Blank ? 3 : 2;
    sheet.insertRowsBefore(at, n + 1);
    // the rows that were on top (now just below) give the new rows their look and dropdowns, and the QB Roll up formula
    let tpl = at + n + 1;
    while (tpl < at + n + 6 && sheet.getRange(tpl, 1, 1, lastCol).getDisplayValues()[0].every(v => !String(v).trim())) tpl++;
    sheet.getRange(tpl, 1, 1, lastCol).copyTo(sheet.getRange(at, 1, n, lastCol), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    const values = rows.map(r => {
      const out = new Array(width).fill('');
      TRACKER_COLS.forEach(([k]) => { if (k === 'qb') return; const v = TRACKER_LINE_FIELDS.includes(k) ? r[k] : shared[k];
        out[cols[k] - 1] = typeof v === 'string' ? asText_(v) : v; });
      return out;
    });
    sheet.getRange(at, 1, n, width).setValues(values);
    const qb = sheet.getRange(tpl, cols.qb).getFormulaR1C1();
    sheet.getRange(at, cols.qb, n, 1).setFormulasR1C1(rows.map(() => [qb || '=VLOOKUP(RC' + cols.product + ',Rollforward!R75C2:R128C3,2,FALSE)']));
    record(n, logged, when, me);
  });
  // remember each item's Product, so the next time the same item is logged it starts with it
  try {
    const learned = trackerLearned_();
    rows.forEach(r => { if (r.name) learned[trackerKey_(r.name)] = r.product; });
    const keys = Object.keys(learned);
    while (keys.length > 300) delete learned[keys.shift()];
    PropertiesService.getScriptProperties().setProperty('TRACKER_PRODUCTS', JSON.stringify(learned));
  } catch (e) { /* only a convenience */ }
  return { logged, rows: n };
}

// ---------- Keeping the app's data fresh ----------
/**
 * When the inventory spreadsheet last changed (any tab, whether changed in the app or right in Google Sheets), as a
 * number. Phones check this every 30 s and only reload items / orders when it moves, so lists stay current without ⟳.
 * It reads the file's details from Drive, not the sheet, so it's quick.
 */
function getDataStamp() {
  requireRole_('Viewer');
  return DriveApp.getFileById(SPREADSHEET_ID).getLastUpdated().getTime();
}

// ---------- Reports ----------
/**
 * Every stock adjustment (received from vendor, returned, added, removed, quick +/−) from one day to another, for the
 * Reports screen. from / to: 'yyyy-MM-dd', inclusive, in the script's time zone; blank = no limit on that side.
 * Oldest first, at most the latest 20,000.
 */
function getAdjustmentsBetween(from, to) {
  requireRole_('Viewer');
  requireScreen_('reports');
  const tz = Session.getScriptTimeZone(), day = iso => Utilities.formatDate(new Date(iso), tz, 'yyyy-MM-dd');
  from = String(from || ''); to = String(to || '');
  return readAdjustments_().filter(a => { const d = day(a.timestamp); return (!from || d >= from) && (!to || d <= to); })
    .slice(-20000).map(a => { delete a.row; return a; });
}
// ===================== 🤖 Ask Supplybot (AI help with the app) =====================
// Anyone with access can ask Supplybot how the app works. The question goes to Claude (Anthropic's API) with a guide to the
// app below. The API key lives in Script Properties (ANTHROPIC_API_KEY), set by an admin in the Ask Supplybot panel (🔑),
// never sent to the browser. Each person can ask up to SB_AI_DAILY questions a day.
const SB_AI_MODEL = 'claude-opus-5-5';
const SB_AI_DAILY = 80;
// For now only these people (and the app's owner) can use Ask Supplybot; everyone else doesn't see it and the server refuses.
const SB_AI_ALLOW = ['marlan.sandoval@envirosparkenergy.com'];
function sbAiAllowed_() { const me = String(currentUser_() || '').toLowerCase(); return !!me && (me === ownerEmail_() || SB_AI_ALLOW.indexOf(me) >= 0); }
function sbAiGuard_() { if (!sbAiAllowed_()) throw new Error('Ask Supplybot isn\'t available yet.'); }
const SB_AI_GUIDE = [
  'You are Supplybot, the friendly robot assistant inside the EnviroSpark "Supplybot" web app. EnviroSpark installs electrical and EV charging equipment; Supplybot is its inventory and orders app (a Google Apps Script web app on top of a Google Sheet). You help people use the app: how to find, do, or change things in it.',
  '',
  'How to answer:',
  '- Short, practical, step by step. Name the exact buttons, tabs and icons as they appear in the app (for example "tap 🚚 Orders", "⚙ by your pins", "My profile → 🎬 Opening slide-ups").',
  '- Plain words, phone-first (most people use it on a phone). Use short numbered steps or bullets; no long essays.',
  '- Only describe features listed in this guide. If you are not sure the app can do something, say so and suggest asking an admin. Never invent buttons or settings.',
  '- You cannot see the person\'s data or press buttons for them. If they ask about specific numbers (stock counts, a certain order), tell them where to look it up.',
  '- If a person lacks permission (View only), explain that an Admin can change their access.',
  '- Be warm and upbeat, a little playful, never pushy. One emoji now and then is fine.',
  '',
  '=== The app ===',
  'Bottom menu (tabs): 📦 Items (everything), ✅ In Stock, ▦ Scan (QR / barcode scanner), 🚚 Orders, 📊 Reports, 🕘 Activity. The top bar has the EnviroSpark logo, Supplybot (tap him for tips, a quote and "Ask Supplybot"), 🏷 All tags, your pins, 🔔 comments, ⟳ refresh, the hide-menus button (four corners) and your profile picture (My profile).',
  'Hide menus: the four-corners button hides the header and bottom menu for more room; a small floating row at the top right stands in (🔍 search, 🏷, pins, the menu emojis, 🔔, 💬, profile picture). Its › arrow shrinks it to just 🔍 and the picture; ‹ opens it again. Tap the logo at the top left to bring the menus back. A static hi-tech ▦ scan orb shows when menus are hidden or a panel is open.',
  'Search: the search box (or 🔍) finds items, orders, returns, POs received and contacts, and understands months ("October"). Every result has 📌 to pin it.',
  'Roles: Admin (everything, users, settings), Editor (make changes), View only (look, comment, make PDFs). An order created by someone else can be locked: only its creator can change it.',
  '',
  'Items: each item has a page with sections you can reorder (⇅ Arrange): 📋 Details (cost, part #, specs, weight), ▦ QR code (its own code to scan and print), 🔢 Quantity (− / + and Adjust stock), 📝 Draft orders, 🏷 Tags, 🔌 Endpoints, 🔁 Same part on other rows, 📦 Stock history, 📎 Documents, 💬 Comments, 📝 My notes. 🎛 changes layout and block style. Photos: add several, one is the display photo. QR codes: an item without one can generate a unique code (SB-XXXXXX); scanning it later opens the item. In the Items list totals box, "▦ QR labels for these N items" makes a printable PDF of labels (optional EnviroSpark logo). List views: 🖼 above the list changes how pictures look (list, banner, grid…). Low stock filter shows the items running low.',
  'Scan: ▦ Scan opens the camera to read QR codes and barcodes (or type a code / use a Bluetooth scanner, or take a photo of the barcode). "Scan to adjust" changes stock by scanning.',
  'Orders: ⚡ New Order builds an order from inventory. It stays a Draft until "🚚 Mark as shipped", which subtracts each ship quantity from stock. Cancelling a shipped order puts the stock back. Order page sections (reorder with ⇅ Arrange): 📋 Details, 📅 Dates (expected pickup/ship date, delivery date), 🏷 Tags, 📍 Ship to (address and map), 🔗 Links (tracking etc.; links can be pinned), 💲 Costs & shipping, 📦 Items (quantities, costs, totals, weights; endpoints per line fold like an accordion), 📷 Photos, 📎 Documents (attach PDFs, Make PDF), 💬 Comments, 📝 My notes. Every section folds like an accordion (tap its header); ⊟ Collapse all / ⊞ Expand all at the top; the app remembers which are closed. Shipped orders: "📝 Send PM a note" and "🧾 Log entry in tracker". Returns on shipped orders are recorded too.',
  'Projects: tap a project name (on an order card or page) to open its project card: all orders with that project summed, Excel download, notes; the PM and company names open their contact cards. Project names are the big highlighted title on order cards. My profile → 🖼 Order pictures → 🎞 can make project names glide across the order cards (Moving) or stay still.',
  'Receiving: 📥 Receive PO adds stock from a purchase order; ↩️ Returns adds back what came back from a job. Both can be searched and logged in the tracker. Adjustments lists stock changes made by hand.',
  'Reports (📊): stock and order reports, with saved Catalogs (📒) that can be pinned and noted.',
  'Supply Chain Requests: a list of requests for material; can be pinned as a shortcut.',
  'Contacts: people and companies (PM, vendor, etc.). A company card can show the whole company and who orders the most.',
  'Activity (🕘): the log of what changed and who changed it.',
  'Comments and mentions: comment on items and orders; type @ to mention someone (they get notified under 🔔). 💬 My comments shows what you wrote and what mentions you. You can drop your own custom emojis into comments (they show at normal emoji size).',
  'Tags: 🏷 tags on items and orders; 🏷 All tags lists them; a tag can be pinned to filter quickly.',
  '',
  'Pins: pin any search result (item, order, return, PO, contact), links, catalogs, notes, shortcuts (📌 on the Orders and Activity cards) and tags. Pins live in the top bar, a floating bubble, or free-floating (My profile → 📌 Pins). Tap a pin to open it, tap it again (while it lights up) to close it. Press and hold to unpin; on a phone a short hold shows its name in a comic bubble. Pins can wear one of your emojis, and pinned items can show the item\'s photo (tiny file, sharp). The ⚙ at the end of the pins opens Quick settings: 🙈 Hide all pins (one small 📌 brings them back), 📌 Pinned results (open, photo or emoji, reorder, unpin), shortcuts to profile sections, and ✎ Choose shortcuts.',
  '',
  'My notes (📝): private notes with links, PDFs, photos and videos (gallery and viewer), arrangeable. Tag a note to items, orders, contacts, projects, companies, catalogs, returns and POs; when you open the note, cards under the title take you straight to each one. A note can have a value ($, typed or ⚡ Smart: worked out live from its tally, tagged orders or items, or time × rate) and a 🧮 Tally (distances, quantities, values, start–end times) that adds up. A note\'s value can be added into an order\'s total. @mention someone in a note to share it read-only. Export a note to PDF (logo optional). Notes show up in search.',
  '',
  'My profile (tap your picture): 💬 My comments, 🎨 App theme (many themes, including Ocean, Sunset, Forest, Neon, Blueprint, Terminal, Paper, Candy, Hazard, Galaxy), 🧱 Block style, 🔤 Font, 📌 Pins, 🧭 Floating menu look, 🖼 Wallpaper, 😀 My emoji, 📝 My notes, 📋 Details & items format, ⏳ Loading logo (it can say what it is loading, above or below), 🖱 Photo hover, 🎬 Opening slide-ups (fourteen ways panels appear, plus ↔️ Page turns: Slide, Card stack, Card swipe, 📖 Page turn, Cross-fade, Off; ‹ Back turns the other way), ✕ Closing slide-ups, 🖼 Item pictures, 🖼 Order pictures. Profile picture can be a photo or an emoji. Settings are saved per person on that device.',
  '😀 Emoji builder: pick a shape (round, square, robot, star, heart, blob, cloud, bolt, book, car, box truck, jet, tank, navy ship, skull, space baby, armored helmet, gold droid, hero bot, dark lord, luchador), eyes (incl. glasses, alien, ninja with mask colour), mouth (incl. mustache, Santa beard, laugh, fangs, silly, zipper, whistle), extras (hard hat with colour and EnviroSpark logo, crown, wings, devil…), girl/boy look and hairstyles, colours and 14 animations. Save up to 8; use them on pins, as your profile picture, or in comments.',
  '',
  'Data and sync: the app checks the spreadsheet every 30 seconds and refreshes quietly when something changed; ⟳ refreshes now. The version is shown at the bottom right. If the app shows an error about files not matching, an admin should paste the latest Code.gs, Index, Styles and Script into Apps Script and deploy a new version.'
].join('\n');

function supplybotAiStatus() {
  const role = requireRole_('Viewer'); sbAiGuard_();
  return { ready: !!PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY'), admin: role === 'Admin', daily: SB_AI_DAILY };
}
function setSupplybotAiKey(key) {
  requireRole_('Admin'); sbAiGuard_();
  key = String(key || '').trim();
  const p = PropertiesService.getScriptProperties();
  if (!key) { p.deleteProperty('ANTHROPIC_API_KEY'); return { ready: false }; }
  if (!/^sk-ant-[A-Za-z0-9_\-]{20,}$/.test(key)) throw new Error('That doesn\'t look like an Anthropic API key (they start with sk-ant-).');
  p.setProperty('ANTHROPIC_API_KEY', key);
  return { ready: true };
}
/** messages: [{ role: 'user'|'assistant', text }] (the chat so far, newest last); ctx: { screen, version }. Returns { text }. */
function askSupplybot(messages, ctx) {
  const role = requireRole_('Viewer'); sbAiGuard_();
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('Ask Supplybot isn\'t set up yet. An admin can switch it on in the 🤖 Ask Supplybot panel (tap Supplybot in the top bar).');
  const me = String(currentUser_() || '').toLowerCase();
  const cache = CacheService.getScriptCache(), dayKey = 'sbai:' + me + ':' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const used = Number(cache.get(dayKey) || 0);
  if (used >= SB_AI_DAILY) throw new Error('You\'ve asked Supplybot ' + SB_AI_DAILY + ' questions today. He\'ll be ready again tomorrow. 🔋');
  const list = (Array.isArray(messages) ? messages : []).slice(-20)
    .map(m => ({ role: m && m.role === 'assistant' ? 'assistant' : 'user', content: String(m && m.text || '').slice(0, 4000) }))
    .filter(m => m.content);
  while (list.length && list[0].role !== 'user') list.shift();
  if (!list.length || list[list.length - 1].role !== 'user') throw new Error('Type a question first.');
  ctx = ctx || {};
  const about = 'About the person asking: role ' + (ROLE_NAMES[role] || role) + '. They are on the "' + String(ctx.screen || 'unknown').slice(0, 60) + '" screen. App version ' + String(ctx.version || '').slice(0, 20) + '.';
  const body = {
    model: SB_AI_MODEL,
    max_tokens: 8000,
    output_config: { effort: 'low' },   // quick, friendly help answers
    fallbacks: 'default',               // if a safety check declines, the API retries on a fallback model in the same call
    system: [{ type: 'text', text: SB_AI_GUIDE, cache_control: { type: 'ephemeral' } }, { type: 'text', text: about }],
    messages: list
  };
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(body),
    headers: { 'x-api-key': key, 'anthropic-version': '2026-10-21', 'anthropic-beta': 'server-side-fallback-2026-07-01' }
  });
  const code = res.getResponseCode();
  let data = {}; try { data = JSON.parse(res.getContentText()); } catch (e) { /* not JSON */ }
  if (code === 401 || code === 403) throw new Error('Supplybot\'s API key was refused. An admin can check it in the 🤖 Ask Supplybot panel (🔑 at the bottom).');
  if (code === 429 || code === 529) throw new Error('Supplybot is busy right now. Try again in a minute.');
  if (code >= 400) throw new Error('Supplybot couldn\'t answer (' + code + (data.error && data.error.message ? ': ' + data.error.message : '') + ').');
  cache.put(dayKey, String(used + 1), 21600);
  if (data.stop_reason === 'refusal') return { text: 'Sorry, I can\'t help with that one. Ask me anything about using Supplybot! 🤖' };
  const text = (data.content || []).filter(b => b.type === 'text').map(b => b.text).join('\n').trim();
  return { text: text || 'Hmm, I didn\'t come up with an answer. Try asking another way?', left: SB_AI_DAILY - used - 1 };
}

// ===================== 📍 Item GPS tags =====================
// Where an item is (a pole, a pedestal, a pallet on site): each tag keeps the coordinates, how accurate the phone's fix was,
// an optional note, and who stored it and when. Tags are never overwritten, so the history stays documented.
const ITEM_GPS_SHEET = 'Item GPS';
const ITEM_GPS_HEADERS = ['Time', 'User', 'Item Key', 'Item', 'Latitude', 'Longitude', 'Accuracy (m)', 'Note'];
function itemGpsSheet_() { const sh = tabWithHeaders_(ITEM_GPS_SHEET, ITEM_GPS_HEADERS); try { if (!sh.isSheetHidden()) sh.hideSheet(); } catch (e) { /* visible */ } return sh; }
function gpsRows_(key) {
  const sh = itemGpsSheet_(), last = sh.getLastRow(); if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, ITEM_GPS_HEADERS.length).getValues()
    .filter(r => String(r[2]) === String(key))
    .map(r => ({ time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), user: String(r[1]), lat: Number(r[4]), lng: Number(r[5]), acc: r[6] === '' ? null : Number(r[6]), note: String(r[7] || '') }))
    .filter(x => isFinite(x.lat) && isFinite(x.lng))
    .sort((a, b) => (a.time < b.time ? 1 : -1)).slice(0, 50);
}
/** The GPS tags of one item, newest first. */
function getItemGps(key) { requireRole_('Viewer'); return gpsRows_(String(key || '')); }
/** Stores a GPS tag for an item (who and when are recorded); returns the item's tags, newest first. */
function addItemGps(key, name, lat, lng, acc, note) {
  requireRole_('Editor');
  key = String(key || ''); lat = Number(lat); lng = Number(lng);
  if (!key) throw new Error('No item.');
  if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) throw new Error('Those coordinates don’t look right (latitude −90 to 90, longitude −180 to 180).');
  acc = acc == null || acc === '' || !isFinite(Number(acc)) ? '' : Math.round(Number(acc) * 10) / 10;
  return withLock_(() => {
    itemGpsSheet_().appendRow([new Date(), currentUser_(), key, String(name || '').slice(0, 200), Math.round(lat * 1e7) / 1e7, Math.round(lng * 1e7) / 1e7, acc, String(note || '').slice(0, 300)]);
    return gpsRows_(key);
  });
}
