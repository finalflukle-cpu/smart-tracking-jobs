/**
 * Smart Tracking Jobs — ระบบหลังบ้าน (Google Apps Script + Google Sheet)
 * --------------------------------------------------------------------
 * วิธีใช้ (ละเอียดใน docs/SETUP.md):
 *   1) สร้าง Google Sheet ใหม่ > ส่วนขยาย > Apps Script
 *   2) ลบโค้ดเดิมทั้งหมด แล้ววางไฟล์นี้แทน > กดบันทึก
 *   3) เลือกฟังก์ชัน setup แล้วกด "เรียกใช้" 1 ครั้ง (อนุญาตสิทธิ์)
 *   4) ทำให้ใช้งานได้ > การทำให้ใช้งานได้รายการใหม่ > เว็บแอป
 *      - ดำเนินการในฐานะ: ฉัน   - ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน
 *   5) คัดลอกลิงก์เว็บแอป (ลงท้าย /exec) ไปใส่ในแอป
 *
 * ห้ามแชร์ Google Sheet นี้ให้คนขับ — ข้อมูลทั้งหมดผ่านแอปเท่านั้น
 * PIN แอดมินเริ่มต้น: 000000 (เปลี่ยนได้ในหน้า ตั้งค่า ของแอป)
 */

var APP_VERSION = '1.3.0';
var DEFAULT_ADMIN_PIN = '000000';
var SESSION_DAYS = { admin: 7, driver: 30 };
var MAX_FAILS = 5;
var LOCK_MINUTES = 15;
var TZ_OFFSET_HOURS = 7; // เวลาประเทศไทย
var DEFAULT_RADIUS_M = 100; // รัศมีเช็คอินเริ่มต้น (เมตร) แก้ได้ในหน้า ตั้งค่า ของแอป
var MAX_PHOTOS = 5;          // รูปเช็คอินสูงสุดต่อจุด
var DAILY_PHOTOS = ['รถภายนอก', 'ป้ายทะเบียน', 'ภายในรถ']; // รูปตรวจความพร้อมประจำวัน (บังคับครบ)
var DAILY_ITEMS = ['name', 'phone', 'vehicle', 'gps'];      // รายการที่คนขับต้องติ๊กยืนยัน
var DEFAULT_GPS_MIN = 3;     // ส่งตำแหน่งทุกกี่นาที (ขณะเปิดแอป)
var LOCATION_KEEP_ROWS = 12000; // เก็บประวัติตำแหน่งย้อนหลังประมาณนี้ (เกินแล้วลบของเก่า)
var SETUP_KEY = 'setup_ok_v4';

var SCHEMA = {
  Settings: ['key', 'value'],
  Drivers: ['id', 'name', 'pin_hash', 'active', 'created_at', 'updated_at', 'photo_id', 'photo_url', 'phone'],
  Vehicles: ['id', 'plate', 'name', 'seats', 'note', 'photo_id', 'photo_url', 'active', 'created_at', 'updated_at'],
  Jobs: ['id', 'job_no', 'date', 'time', 'title', 'driver_id', 'plate', 'note', 'status',
         'route_polyline', 'distance_km', 'duration_min',
         'created_at', 'updated_at', 'started_at', 'finished_at', 'vehicle_id'],
  Stops: ['id', 'job_id', 'seq', 'name', 'address', 'lat', 'lng', 'map_url', 'note', 'planned_time',
          'status', 'arrived_at', 'arrive_lat', 'arrive_lng', 'photo_id', 'photo_url', 'checkin_note',
          'photo_urls', 'checkin_dist_m', 'checkin_acc_m'],
  Places: ['id', 'name', 'address', 'lat', 'lng', 'map_url', 'used_count', 'last_used'],
  Sessions: ['token', 'role', 'user_id', 'name', 'created_at', 'expires_at'],
  Log: ['time', 'who', 'action', 'detail'],
  DailyChecks: ['id', 'date', 'driver_id', 'driver_name', 'vehicle_id', 'plate', 'photo_urls', 'items', 'phone',
                'lat', 'lng', 'acc_m', 'created_at'],
  Locations: ['at', 'driver_id', 'name', 'lat', 'lng', 'acc_m', 'source', 'job_id']
};

var REQ = {}; // แคชภายในคำขอเดียว

/* ====================== จุดเข้าใช้งาน ====================== */

function doGet() {
  return json_({
    ok: true,
    app: 'Smart Tracking Jobs API',
    version: APP_VERSION,
    message: 'ระบบหลังบ้านทำงานปกติ — นำลิงก์นี้ไปใส่ในหน้า "เชื่อมต่อระบบ" ของแอป'
  });
}

function doPost(e) {
  REQ = {};
  var req;
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'รูปแบบคำขอไม่ถูกต้อง' });
  }
  try {
    ensureSetup_();
    var fn = ACTIONS[req.action];
    if (!fn) fail_('ไม่รู้จักคำสั่ง: ' + req.action);
    return json_({ ok: true, data: fn(req) });
  } catch (err) {
    if (err && err.isApp) return json_({ ok: false, error: err.message, code: err.code || '' });
    try { log_('system', 'error', String((err && err.stack) || err).slice(0, 2000)); } catch (e2) {}
    return json_({ ok: false, error: 'เกิดข้อผิดพลาดในระบบ: ' + ((err && err.message) || err) });
  }
}

/** เรียกใช้ครั้งแรกจากหน้า Apps Script เพื่อสร้างแผ่นงานและอนุญาตสิทธิ์ */
function setup() {
  REQ = {};
  CacheService.getScriptCache().remove(SETUP_KEY);
  ensureSetup_();
  DriveApp.getRootFolder();            // ขอสิทธิ์ Google Drive (เก็บรูปเช็คอิน)
  Maps.newGeocoder();                  // ขอสิทธิ์ค้นหาสถานที่
  Logger.log('ตั้งค่าเสร็จแล้ว PIN แอดมินเริ่มต้น = ' + DEFAULT_ADMIN_PIN);
}

/** กรณีลืม PIN แอดมิน: เรียกใช้ฟังก์ชันนี้จากหน้า Apps Script เพื่อรีเซ็ตเป็น 000000 */
function resetAdminPin() {
  REQ = {};
  ensureSetup_();
  setSetting_('admin_pin_hash', hash_('admin:' + DEFAULT_ADMIN_PIN));
  setSetting_('admin_pin_default', 'true');
  CacheService.getScriptCache().remove('fail:admin');
  Logger.log('รีเซ็ต PIN แอดมินเป็น ' + DEFAULT_ADMIN_PIN + ' แล้ว');
}

var ACTIONS = {
  ping: function () {
    var out = settingsOut_();
    out.app = 'Smart Tracking Jobs';
    out.version = APP_VERSION;
    return out;
  },
  publicDrivers: function () {
    return rows_('Drivers').filter(function (d) { return d.active === 'true'; })
      .map(function (d) { return { id: d.id, name: d.name, photo_url: d.photo_url }; })
      .sort(function (a, b) { return a.name.localeCompare(b.name, 'th'); });
  },
  login: login_,
  logout: logout_,
  me: me_,
  changeAdminPin: changeAdminPin_,
  saveSettings: saveSettings_,
  listDrivers: listDrivers_,
  saveDriver: saveDriver_,
  listVehicles: listVehicles_,
  saveVehicle: saveVehicle_,
  listJobs: listJobs_,
  getJob: getJob_,
  saveJob: saveJob_,
  cancelJob: cancelJob_,
  deleteJob: deleteJob_,
  checkIn: checkIn_,
  searchPlace: searchPlace_,
  resolveLink: resolveLink_,
  suggestPlace: suggestPlace_,
  placeDetails: placeDetails_,
  listPlaces: listPlaces_,
  myDay: myDay_,
  saveDailyCheck: saveDailyCheck_,
  reportLocation: reportLocation_,
  dayStatus: dayStatus_,
  clientLog: clientLog_
};

/* ====================== เข้าสู่ระบบ ====================== */

function login_(req) {
  var pin = String(req.pin || '');
  if (!/^\d{6}$/.test(pin)) fail_('PIN ต้องเป็นตัวเลข 6 หลัก');
  var company = getSetting_('company_name') || '';

  if (req.role === 'admin') {
    checkLock_('admin');
    if (hash_('admin:' + pin) !== getSetting_('admin_pin_hash')) badPin_('admin');
    clearFail_('admin');
    var admin = { id: 'admin', name: 'แอดมิน' };
    var token = withLock_(function () { cleanupSessions_(); var tk = createSession_('admin', admin); log_('admin', 'login', ''); return tk; });
    return { token: token, role: 'admin', user: admin, company: company,
             pinDefault: getSetting_('admin_pin_default') === 'true' };
  }

  if (req.role === 'driver') {
    var d = rows_('Drivers').filter(function (x) {
      return x.id === String(req.driverId || '') && x.active === 'true';
    })[0];
    if (!d) fail_('ไม่พบชื่อคนขับนี้ หรือถูกปิดการใช้งานแล้ว');
    checkLock_('d:' + d.id);
    if (hash_('driver:' + d.id + ':' + pin) !== d.pin_hash) badPin_('d:' + d.id);
    clearFail_('d:' + d.id);
    var user = { id: d.id, name: d.name, photo_url: d.photo_url || '' };
    var t = withLock_(function () { cleanupSessions_(); var tk = createSession_('driver', user); log_(d.name, 'login', ''); return tk; });
    return { token: t, role: 'driver', user: user, company: company };
  }
  fail_('ประเภทผู้ใช้ไม่ถูกต้อง');
}

function logout_(req) {
  var t = String(req.token || '');
  if (t) withLock_(function () { killSessions_(function (s) { return s.token === t; }); });
  return true;
}

function me_(req) {
  var s = auth_(req);
  var out = { role: s.role, user: { id: s.user_id, name: s.name }, company: getSetting_('company_name') || '' };
  if (s.role === 'admin') out.pinDefault = getSetting_('admin_pin_default') === 'true';
  return out;
}

function changeAdminPin_(req) {
  var s = auth_(req, 'admin');
  var oldPin = String(req.oldPin || ''), newPin = String(req.newPin || '');
  checkLock_('admin');
  if (hash_('admin:' + oldPin) !== getSetting_('admin_pin_hash')) badPin_('admin', 'PIN เดิมไม่ถูกต้อง');
  clearFail_('admin');
  if (!/^\d{6}$/.test(newPin)) fail_('PIN ใหม่ต้องเป็นตัวเลข 6 หลัก');
  if (newPin === DEFAULT_ADMIN_PIN) fail_('ห้ามใช้ 000000 เป็น PIN ใหม่');
  setSetting_('admin_pin_hash', hash_('admin:' + newPin));
  setSetting_('admin_pin_default', 'false');
  withLock_(function () {
    killSessions_(function (x) { return x.role === 'admin' && x.token !== s.token; });
    log_('admin', 'change_pin', '');
  });
  return true;
}

function saveSettings_(req) {
  auth_(req, 'admin');
  if (req.company_name !== undefined) {
    var name = clean_(req.company_name, 80);
    if (!name) fail_('กรุณาใส่ชื่อบริษัท');
    setSetting_('company_name', name);
  }
  if (req.checkin_radius_m !== undefined) {
    var r = Number(req.checkin_radius_m);
    if (!isFinite(r) || r < 0 || r > 5000 || Math.round(r) !== r) fail_('รัศมีเช็คอินต้องเป็นตัวเลข 0–5000 เมตร');
    setSetting_('checkin_radius_m', String(r));
  }
  if (req.daily_check !== undefined) setSetting_('daily_check', req.daily_check ? 'on' : 'off');
  if (req.gps_tracking !== undefined) setSetting_('gps_tracking', req.gps_tracking ? 'on' : 'off');
  if (req.gps_interval_min !== undefined) {
    var g = Number(req.gps_interval_min);
    if (!isFinite(g) || g < 1 || g > 30 || Math.round(g) !== g) fail_('ความถี่ส่งตำแหน่งต้องเป็น 1–30 นาที');
    setSetting_('gps_interval_min', String(g));
  }
  if (req.places_api_key !== undefined) {
    var key = clean_(req.places_api_key, 200);
    if (key) placesAutocomplete_('สยามพารากอน', key, ''); // ทดสอบก่อนบันทึก (ถ้าใช้ไม่ได้จะแจ้งเหตุผล)
    setSetting_('places_api_key', key);
  }
  return settingsOut_();
}

function settingsOut_() {
  return {
    company: getSetting_('company_name') || '',
    checkin_radius_m: radius_(),
    daily_check: dailyCheckOn_(),
    gps_tracking: gpsOn_(),
    gps_interval_min: gpsMin_(),
    places_api: !!getSetting_('places_api_key')
  };
}
function dailyCheckOn_() { return getSetting_('daily_check') !== 'off'; }
function gpsOn_() { return getSetting_('gps_tracking') !== 'off'; }
function gpsMin_() { var g = Number(getSetting_('gps_interval_min')); return g >= 1 && g <= 30 ? g : DEFAULT_GPS_MIN; }

/* ====================== คนขับ ====================== */

function listDrivers_(req) {
  auth_(req, 'admin');
  return rows_('Drivers').map(function (d) {
    return { id: d.id, name: d.name, phone: d.phone, photo_url: d.photo_url, active: d.active === 'true', created_at: d.created_at };
  }).sort(function (a, b) {
    if (a.active !== b.active) return a.active ? -1 : 1;
    return a.name.localeCompare(b.name, 'th');
  });
}

function saveDriver_(req) {
  auth_(req, 'admin');
  var d = req.driver || {};
  var name = clean_(d.name, 60).replace(/\s+/g, ' ');
  var pin = d.pin ? String(d.pin) : '';
  var active = !(d.active === false || d.active === 'false');
  var phone = clean_(d.phone, 20);
  var photo = parseImage_(d.photo);
  if (!name) fail_('กรุณาใส่ชื่อคนขับ');
  if (pin && !/^\d{6}$/.test(pin)) fail_('PIN ต้องเป็นตัวเลข 6 หลัก');

  return withLock_(function () {
    var all = rows_('Drivers');
    var dup = all.filter(function (x) {
      return x.active === 'true' && x.id !== d.id && x.name.toLowerCase() === name.toLowerCase();
    })[0];
    if (dup && active) fail_('มีคนขับชื่อ "' + name + '" อยู่แล้ว');

    if (d.id) {
      var cur = all.filter(function (x) { return x.id === d.id; })[0];
      if (!cur) fail_('ไม่พบคนขับ');
      var patch = { name: name, active: active ? 'true' : 'false', updated_at: now_() };
      if (d.phone !== undefined) patch.phone = phone;
      if (pin) patch.pin_hash = hash_('driver:' + cur.id + ':' + pin);
      if (photo) { var f1 = savePhoto_(photo.ext, photo.b64, 'คนขับ_' + name, 'w600'); patch.photo_id = f1.id; patch.photo_url = f1.url; }
      else if (d.removePhoto) { patch.photo_id = ''; patch.photo_url = ''; }
      update_('Drivers', cur, patch);
      if (pin || !active) killSessions_(function (s) { return s.role === 'driver' && s.user_id === cur.id; });
      if (pin) clearFail_('d:' + cur.id);
      log_('admin', 'update_driver', name + (pin ? ' (เปลี่ยน PIN)' : '') + (active ? '' : ' (ปิดใช้งาน)'));
      return { id: cur.id, name: name, active: active, photo_url: cur.photo_url };
    }

    if (!pin) fail_('กรุณาตั้ง PIN 6 หลักให้คนขับ');
    var id = 'D' + shortId_();
    var f2 = photo ? savePhoto_(photo.ext, photo.b64, 'คนขับ_' + name, 'w600') : null;
    insert_('Drivers', {
      id: id, name: name, pin_hash: hash_('driver:' + id + ':' + pin),
      active: 'true', created_at: now_(), updated_at: now_(),
      photo_id: f2 ? f2.id : '', photo_url: f2 ? f2.url : '', phone: phone
    });
    log_('admin', 'add_driver', name);
    return { id: id, name: name, active: true, photo_url: f2 ? f2.url : '' };
  });
}

/* ====================== รถ ====================== */

function listVehicles_(req) {
  auth_(req, 'admin');
  return rows_('Vehicles').map(vehicleOut_).sort(function (a, b) {
    if (a.active !== b.active) return a.active ? -1 : 1;
    return a.plate.localeCompare(b.plate, 'th');
  });
}

function vehicleOut_(v) {
  return { id: v.id, plate: v.plate, name: v.name, seats: num_(v.seats), note: v.note,
           photo_url: v.photo_url, active: v.active === 'true', created_at: v.created_at };
}

function saveVehicle_(req) {
  auth_(req, 'admin');
  var v = req.vehicle || {};
  var plate = clean_(v.plate, 20).replace(/\s+/g, ' ');
  var seats = v.seats === '' || v.seats === null || v.seats === undefined ? '' : Number(v.seats);
  var active = !(v.active === false || v.active === 'false');
  var photo = parseImage_(v.photo);
  if (!plate) fail_('กรุณาใส่ทะเบียนรถ');
  if (seats !== '' && (!isFinite(seats) || seats < 1 || seats > 99)) fail_('จำนวนที่นั่งต้องเป็นตัวเลข 1–99');

  return withLock_(function () {
    var all = rows_('Vehicles');
    var norm = function (x) { return String(x).replace(/\s+/g, '').toLowerCase(); };
    var dup = all.filter(function (x) { return x.active === 'true' && x.id !== v.id && norm(x.plate) === norm(plate); })[0];
    if (dup && active) fail_('มีรถทะเบียน "' + plate + '" อยู่แล้ว');
    var rec = { plate: plate, name: clean_(v.name, 60), seats: seats, note: clean_(v.note, 200),
                active: active ? 'true' : 'false', updated_at: now_() };
    if (photo) { var f = savePhoto_(photo.ext, photo.b64, 'รถ_' + plate, 'w1000'); rec.photo_id = f.id; rec.photo_url = f.url; }
    else if (v.removePhoto) { rec.photo_id = ''; rec.photo_url = ''; }
    if (v.id) {
      var cur = all.filter(function (x) { return x.id === v.id; })[0];
      if (!cur) fail_('ไม่พบรถคันนี้');
      update_('Vehicles', cur, rec);
      log_('admin', 'update_vehicle', plate + (active ? '' : ' (ปิดใช้งาน)'));
      return vehicleOut_(cur);
    }
    rec.id = 'V' + shortId_();
    rec.created_at = now_();
    insert_('Vehicles', rec);
    log_('admin', 'add_vehicle', plate);
    return vehicleOut_(rec);
  });
}

function parseImage_(dataUrl) {
  if (!dataUrl) return null;
  var m = String(dataUrl).match(/^data:image\/(jpeg|png|webp);base64,([A-Za-z0-9+\/=]+)$/);
  if (!m) fail_('ไฟล์รูปไม่ถูกต้อง');
  if (m[2].length > 8000000) fail_('รูปมีขนาดใหญ่เกินไป');
  return { ext: m[1], b64: m[2] };
}

/* ====================== Job ====================== */

function listJobs_(req) {
  var s = auth_(req);
  var today = today_();
  if (s.role === 'driver') {
    var from = addDays_(today, -1), to = addDays_(today, 30);
    return loadJobs_(function (j) {
      if (j.driver_id !== s.user_id || j.status === 'cancelled') return false;
      return (j.date >= from && j.date <= to) || (j.date < today && j.status !== 'done');
    });
  }
  var f = String(req.from || today), t = String(req.to || f);
  if (!isDate_(f) || !isDate_(t)) fail_('วันที่ไม่ถูกต้อง');
  if (t < f) { var tmp = f; f = t; t = tmp; }
  if (addDays_(f, 186) < t) fail_('ช่วงวันที่ยาวเกินไป (สูงสุด 6 เดือน)');
  return loadJobs_(function (j) { return j.date >= f && j.date <= t; });
}

function getJob_(req) {
  var s = auth_(req);
  var job = getJobById_(String(req.id || ''));
  if (!job) fail_('ไม่พบ Job นี้');
  if (s.role === 'driver' && job.driver_id !== s.user_id) fail_('Job นี้ไม่ใช่งานของคุณ', 'FORBIDDEN');
  return job;
}

function saveJob_(req) {
  auth_(req, 'admin');
  var inp = req.job || {};
  var stopsIn = req.stops || [];
  var date = String(inp.date || '');
  var time = String(inp.time || '');
  var driverId = String(inp.driver_id || '');
  var plate = clean_(inp.plate, 20);
  var vehicleId = String(inp.vehicle_id || '');
  if (!isDate_(date)) fail_('กรุณาเลือกวันที่');
  if (time && !/^\d{2}:\d{2}$/.test(time)) fail_('เวลาเริ่มไม่ถูกต้อง');
  if (!driverId) fail_('กรุณาเลือกคนขับ');
  if (!plate && !vehicleId) fail_('กรุณาเลือกรถ');
  if (!stopsIn.length) fail_('กรุณาเพิ่มสถานที่อย่างน้อย 1 จุด');
  if (stopsIn.length > 25) fail_('เพิ่มสถานที่ได้สูงสุด 25 จุดต่อ Job');

  var clean = stopsIn.map(function (s, i) {
    var name = clean_(s.name, 120);
    if (!name) fail_('สถานที่จุดที่ ' + (i + 1) + ' ยังไม่มีชื่อ');
    var lat = num_(s.lat), lng = num_(s.lng);
    if (lat === null || lng === null || Math.abs(lat) > 90 || Math.abs(lng) > 180) { lat = null; lng = null; }
    return {
      id: String(s.id || ''), name: name, address: clean_(s.address, 250),
      lat: lat, lng: lng, map_url: safeUrl_(s.map_url), note: clean_(s.note, 300),
      planned_time: /^\d{2}:\d{2}$/.test(String(s.planned_time || '')) ? String(s.planned_time) : ''
    };
  });

  return withLock_(function () {
    var drv = rows_('Drivers').filter(function (d) { return d.id === driverId; })[0];
    if (!drv) fail_('ไม่พบคนขับที่เลือก');
    var jobs = rows_('Jobs');
    var job = null;
    if (inp.id) {
      job = jobs.filter(function (j) { return j.id === String(inp.id); })[0];
      if (!job) fail_('ไม่พบ Job นี้');
      if (job.status === 'cancelled') fail_('Job นี้ถูกยกเลิกแล้ว แก้ไขไม่ได้');
    }
    if (drv.active !== 'true' && !(job && job.driver_id === drv.id)) fail_('คนขับคนนี้ถูกปิดการใช้งาน');
    if (vehicleId) {
      var veh = rows_('Vehicles').filter(function (x) { return x.id === vehicleId; })[0];
      if (!veh) fail_('ไม่พบรถที่เลือก');
      if (veh.active !== 'true' && !(job && job.vehicle_id === veh.id)) fail_('รถคันนี้ถูกปิดการใช้งาน');
      plate = veh.plate;
    }

    var oldStops = job ? rows_('Stops').filter(function (s) { return s.job_id === job.id; }) : [];
    var oldById = indexBy_(oldStops, 'id');
    var keep = {};
    clean.forEach(function (s) { if (s.id && oldById[s.id]) keep[s.id] = true; });
    oldStops.forEach(function (o) {
      if (o.status === 'arrived' && !keep[o.id]) fail_('ลบ "' + o.name + '" ไม่ได้ เพราะคนขับเช็คอินแล้ว');
    });
    if (job && job.driver_id !== drv.id && job.status !== 'pending') {
      fail_('เปลี่ยนคนขับไม่ได้ เพราะเริ่มงานไปแล้ว');
    }

    var jobId = job ? job.id : 'J' + shortId_();
    var jobNo = job ? job.job_no : nextJobNo_(jobs, date);
    var route = computeRoute_(clean, oldStops, job);

    deleteRows_('Stops', oldStops.map(function (o) { return o._row; }));
    var arrived = 0;
    clean.forEach(function (s, i) {
      var o = s.id ? oldById[s.id] : null;
      if (o && o.status === 'arrived') arrived++;
      insert_('Stops', {
        id: o ? o.id : 'S' + shortId_(), job_id: jobId, seq: i + 1,
        name: s.name, address: s.address, lat: s.lat === null ? '' : s.lat, lng: s.lng === null ? '' : s.lng,
        map_url: s.map_url, note: s.note, planned_time: s.planned_time,
        status: o ? o.status : 'pending', arrived_at: o ? o.arrived_at : '',
        arrive_lat: o ? o.arrive_lat : '', arrive_lng: o ? o.arrive_lng : '',
        photo_id: o ? o.photo_id : '', photo_url: o ? o.photo_url : '', checkin_note: o ? o.checkin_note : '',
        photo_urls: o ? o.photo_urls : '', checkin_dist_m: o ? o.checkin_dist_m : '', checkin_acc_m: o ? o.checkin_acc_m : ''
      });
    });

    var status = arrived === 0 ? 'pending' : (arrived === clean.length ? 'done' : 'in_progress');
    var rec = {
      id: jobId, job_no: jobNo, date: date, time: time, title: clean_(inp.title, 120),
      driver_id: drv.id, plate: plate, vehicle_id: vehicleId, note: clean_(inp.note, 500), status: status,
      route_polyline: route ? route.polyline : '', distance_km: route ? route.km : '',
      duration_min: route ? route.min : '',
      created_at: job ? job.created_at : now_(), updated_at: now_(),
      started_at: job ? job.started_at : '',
      finished_at: status === 'done' ? ((job && job.finished_at) || now_()) : ''
    };
    if (job) update_('Jobs', job, rec); else insert_('Jobs', rec);
    upsertPlaces_(clean);
    log_('admin', job ? 'update_job' : 'create_job', jobNo);
    return getJobById_(jobId);
  });
}

function cancelJob_(req) {
  auth_(req, 'admin');
  return withLock_(function () {
    var job = rows_('Jobs').filter(function (j) { return j.id === String(req.id || ''); })[0];
    if (!job) fail_('ไม่พบ Job นี้');
    if (job.status === 'done') fail_('Job นี้เสร็จแล้ว ยกเลิกไม่ได้');
    update_('Jobs', job, { status: 'cancelled', updated_at: now_() });
    log_('admin', 'cancel_job', job.job_no);
    return getJobById_(job.id);
  });
}

function deleteJob_(req) {
  auth_(req, 'admin');
  return withLock_(function () {
    var job = rows_('Jobs').filter(function (j) { return j.id === String(req.id || ''); })[0];
    if (!job) fail_('ไม่พบ Job นี้');
    var stops = rows_('Stops').filter(function (s) { return s.job_id === job.id; });
    if (stops.some(function (s) { return s.status === 'arrived'; })) {
      fail_('มีการเช็คอินแล้ว ลบไม่ได้ — ใช้ "ยกเลิก Job" แทน');
    }
    deleteRows_('Stops', stops.map(function (s) { return s._row; }));
    deleteRows_('Jobs', [job._row]);
    log_('admin', 'delete_job', job.job_no);
    return true;
  });
}

/* ====================== เช็คอิน ====================== */

function checkIn_(req) {
  var s = auth_(req, 'driver');
  var list = Array.isArray(req.photos) ? req.photos : (req.photo ? [req.photo] : []);
  if (!list.length) fail_('กรุณาถ่ายรูปเช็คอินอย่างน้อย 1 รูป');
  if (list.length > MAX_PHOTOS) fail_('แนบรูปได้สูงสุด ' + MAX_PHOTOS + ' รูปต่อจุด');
  var photos = list.map(function (p) {
    var m = String(p || '').match(/^data:image\/(jpeg|png|webp);base64,([A-Za-z0-9+\/=]+)$/);
    if (!m) fail_('รูปเช็คอินไม่ถูกต้อง กรุณาถ่ายใหม่');
    if (m[2].length > 10000000) fail_('รูปมีขนาดใหญ่เกินไป');
    return { ext: m[1], b64: m[2] };
  });
  var lat = num_(req.lat), lng = num_(req.lng), acc = num_(req.acc);
  if (dailyCheckOn_() && !findDailyCheck_(s.user_id, today_())) {
    fail_('กรุณาตรวจความพร้อมประจำวันก่อนเช็คอิน (เมนู "ตรวจความพร้อม" ด้านล่าง)', 'NEED_DAILY');
  }

  return withLock_(function () {
    var stop = rows_('Stops').filter(function (x) { return x.id === String(req.stopId || ''); })[0];
    if (!stop) fail_('ไม่พบสถานที่นี้ใน Job');
    var job = rows_('Jobs').filter(function (j) { return j.id === stop.job_id; })[0];
    if (!job || job.driver_id !== s.user_id) fail_('สถานที่นี้ไม่ใช่งานของคุณ', 'FORBIDDEN');
    if (job.status === 'cancelled') fail_('Job นี้ถูกยกเลิกแล้ว');
    if (stop.status === 'arrived') fail_('จุดนี้เช็คอินไปแล้ว');

    // ตรวจรัศมี: ต้องอยู่ใกล้จุดที่ปักหมุดไว้
    var radius = radius_();
    var sLat = num_(stop.lat), sLng = num_(stop.lng);
    var dist = null;
    if (sLat !== null && sLng !== null) {
      if (lat !== null && lng !== null) dist = distanceM_(sLat, sLng, lat, lng);
      if (radius > 0) {
        if (dist === null) fail_('ต้องเปิดตำแหน่ง (GPS) ก่อนเช็คอิน', 'NO_GPS');
        if (dist > radius) fail_('ตอนนี้อยู่ห่างจากจุดนัด ' + fmtDist_(dist) + ' ต้องอยู่ในรัศมี ' + radius + ' ม. จึงจะเช็คอินได้', 'OUT_OF_RANGE');
      }
    }

    var saved = photos.map(function (p, i) {
      return savePhoto_(p.ext, p.b64, job.job_no + '_จุด' + stop.seq + '_' + (i + 1));
    });
    var t = now_();
    update_('Stops', stop, {
      status: 'arrived', arrived_at: t,
      arrive_lat: lat === null ? '' : lat, arrive_lng: lng === null ? '' : lng,
      photo_id: saved[0].id, photo_url: saved[0].url, photo_urls: JSON.stringify(saved),
      checkin_note: clean_(req.note, 300),
      checkin_dist_m: dist === null ? '' : dist, checkin_acc_m: acc === null ? '' : Math.round(acc)
    });
    var all = rows_('Stops').filter(function (x) { return x.job_id === job.id; });
    var done = all.filter(function (x) { return x.status === 'arrived'; }).length;
    var patch = { updated_at: t, status: done >= all.length ? 'done' : 'in_progress' };
    if (!job.started_at) patch.started_at = t;
    if (patch.status === 'done') patch.finished_at = t;
    update_('Jobs', job, patch);
    if (lat !== null && lng !== null) addLocation_(s, lat, lng, acc, 'checkin', job.id);
    log_(s.name, 'check_in', job.job_no + ' จุด ' + stop.seq + ' ' + stop.name + ' · ' + saved.length + ' รูป' + (dist === null ? '' : ' · ห่าง ' + dist + ' ม.'));
    return getJobById_(job.id);
  });
}

function radius_() {
  var v = getSetting_('checkin_radius_m');
  return v === '' ? DEFAULT_RADIUS_M : Math.max(0, Number(v) || 0);
}

function distanceM_(lat1, lng1, lat2, lng2) {
  var R = 6371000, toR = Math.PI / 180;
  var dLat = (lat2 - lat1) * toR, dLng = (lng2 - lng1) * toR;
  var h = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
          Math.cos(lat1 * toR) * Math.cos(lat2 * toR) * Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return Math.round(2 * R * Math.asin(Math.sqrt(h)));
}

function fmtDist_(m) { return m < 1000 ? m + ' ม.' : (Math.round(m / 100) / 10) + ' กม.'; }

function savePhoto_(ext, b64, base, size) {
  var folder = photoFolder_();
  var bytes = Utilities.base64Decode(b64);
  var name = base + '_' + now_().replace(/[^0-9]/g, '') + '.' + (ext === 'jpeg' ? 'jpg' : ext);
  var file = folder.createFile(Utilities.newBlob(bytes, 'image/' + ext, name));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  var id = file.getId();
  return { id: id, url: 'https://drive.google.com/thumbnail?id=' + id + '&sz=' + (size || 'w1600') };
}

function photoFolder_() {
  var id = getSetting_('photo_folder_id');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) {}
  }
  var folder = DriveApp.createFolder('Smart Tracking รูปเช็คอิน — ' + ss_().getName());
  setSetting_('photo_folder_id', folder.getId());
  return folder;
}

/* ====================== แผนที่ ====================== */

function searchPlace_(req) {
  auth_(req, 'admin');
  var q = clean_(req.q, 150);
  if (q.length < 2) fail_('พิมพ์คำค้นหาอย่างน้อย 2 ตัวอักษร');
  var co = parseLatLng_(q);
  if (co) return [{ name: 'พิกัด ' + q, address: '', lat: co.lat, lng: co.lng, map_url: mapsUrl_(co.lat, co.lng) }];
  return geocodeList_(q, 6);
}

function resolveLink_(req) {
  auth_(req, 'admin');
  var url = String(req.url || '').trim();
  var co = parseLatLng_(url);
  if (co) return { name: '', address: '', lat: co.lat, lng: co.lng, map_url: mapsUrl_(co.lat, co.lng) };
  var mUrl = url.match(/https:\/\/[^\s]+/i); // รองรับข้อความที่แชร์มาพร้อมลิงก์
  if (!mUrl) fail_('กรุณาวางลิงก์จาก Google Maps (ขึ้นต้นด้วย https://)');
  url = mUrl[0];
  var found = followMapsLink_(url);
  if (found.lat !== null) return { name: found.name, address: '', lat: found.lat, lng: found.lng, map_url: url };
  // ลิงก์แชร์จาก iPhone บางแบบไม่มีพิกัด มีแค่ชื่อ/ที่อยู่ → ค้นพิกัดจากชื่อด้วย Google แทน
  var q = found.name || clean_(req.hint, 150);
  var g = q ? geocodeList_(q, 1)[0] : null;
  if (g) return { name: found.name || g.name, address: g.address, lat: g.lat, lng: g.lng, map_url: url, approx: true };
  return { name: found.name || '', address: '', lat: null, lng: null, map_url: url };
}

function followMapsLink_(url) {
  var cur = url, name = '';
  for (var i = 0; i < 6; i++) {
    if (!isGoogleHost_(cur)) fail_('รองรับเฉพาะลิงก์จาก Google Maps');
    var info = parseMapsUrl_(cur);
    if (info.name) name = info.name;
    if (info.lat !== null) return { name: name, lat: info.lat, lng: info.lng };
    var res = UrlFetchApp.fetch(cur, { followRedirects: false, muteHttpExceptions: true });
    var h = res.getHeaders() || {};
    var loc = h.Location || h.location;
    if (!loc) {
      var body = parseMapsUrl_(String(res.getContentText() || '').slice(0, 300000));
      return { name: name || body.name, lat: body.lat, lng: body.lng };
    }
    if (Array.isArray(loc)) loc = loc[0];
    cur = /^https?:\/\//i.test(loc) ? loc : cur.replace(/^(https?:\/\/[^\/]+).*$/i, '$1') + loc;
  }
  var last = parseMapsUrl_(cur);
  return { name: last.name || name, lat: last.lat, lng: last.lng };
}

/* ค้นหาด้วย Geocoder ของ Apps Script (ฟรี) + จำผลไว้ 6 ชม. เพื่อประหยัดโควตา */
function geocodeList_(q, max) {
  var cache = CacheService.getScriptCache();
  var key = 'geo:' + hex_(q);
  var hit = cache.get(key);
  var list;
  if (hit) {
    list = JSON.parse(hit);
  } else {
    var r = Maps.newGeocoder().setRegion('th').setLanguage('th').geocode(q);
    if (!r || r.status === 'ZERO_RESULTS') list = [];
    else if (r.status !== 'OK') fail_('ค้นหาไม่สำเร็จ (' + r.status + ') ลองวางลิงก์ Google Maps แทน');
    else {
      list = (r.results || []).slice(0, 6).map(function (x) {
        var loc = x.geometry.location;
        return { name: guessName_(x, q), address: x.formatted_address || '', lat: loc.lat, lng: loc.lng,
                 place_id: x.place_id || '', map_url: mapsUrl_(loc.lat, loc.lng, x.place_id) };
      });
    }
    try { cache.put(key, JSON.stringify(list), 21600); } catch (e) {}
  }
  return list.slice(0, max || 6);
}

/* คำแนะนำสถานที่ระหว่างพิมพ์: ใช้ Google Places ถ้าแอดมินใส่ API key ไว้ ไม่งั้นใช้ Geocoder ฟรี */
function suggestPlace_(req) {
  auth_(req, 'admin');
  var q = clean_(req.q, 150);
  if (q.length < 2) return { source: 'none', items: [] };
  var co = parseLatLng_(q);
  if (co) return { source: 'coords', items: [{ name: 'พิกัด ' + q, address: '', lat: co.lat, lng: co.lng, map_url: mapsUrl_(co.lat, co.lng) }] };
  var key = getSetting_('places_api_key');
  if (key) {
    try { return { source: 'places', items: placesAutocomplete_(q, key, clean_(req.session, 80)) }; }
    catch (e) { /* ใช้ Geocoder แทน */ }
  }
  try { return { source: 'geocoder', items: geocodeList_(q, 5) }; }
  catch (e2) { return { source: 'geocoder', items: [], error: e2.message }; }
}

function placesAutocomplete_(q, key, session) {
  var body = { input: q, languageCode: 'th', regionCode: 'th', includedRegionCodes: ['th'] };
  if (session) body.sessionToken = session;
  var res = UrlFetchApp.fetch('https://places.googleapis.com/v1/places:autocomplete', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'X-Goog-Api-Key': key }, payload: JSON.stringify(body)
  });
  var j = {};
  try { j = JSON.parse(res.getContentText() || '{}'); } catch (e) {}
  if (res.getResponseCode() !== 200) {
    fail_('Google Places ใช้งานไม่ได้: ' + ((j.error && j.error.message) || ('HTTP ' + res.getResponseCode())).slice(0, 200));
  }
  return (j.suggestions || []).filter(function (x) { return x.placePrediction; }).slice(0, 6).map(function (x) {
    var p = x.placePrediction, sf = p.structuredFormat || {};
    return {
      place_id: p.placeId, needs_details: true, lat: null, lng: null,
      name: (sf.mainText && sf.mainText.text) || (p.text && p.text.text) || '',
      address: (sf.secondaryText && sf.secondaryText.text) || ''
    };
  });
}

function placeDetails_(req) {
  auth_(req, 'admin');
  var id = String(req.place_id || '').replace(/[^A-Za-z0-9_\-]/g, '');
  if (!id) fail_('ไม่พบรหัสสถานที่');
  var key = getSetting_('places_api_key');
  if (!key) fail_('ยังไม่ได้ตั้งค่า Google Places API key');
  var session = clean_(req.session, 80);
  var res = UrlFetchApp.fetch('https://places.googleapis.com/v1/places/' + id + '?languageCode=th' +
    (session ? '&sessionToken=' + encodeURIComponent(session) : ''), {
    method: 'get', muteHttpExceptions: true,
    headers: { 'X-Goog-Api-Key': key, 'X-Goog-FieldMask': 'id,displayName,formattedAddress,location' }
  });
  var j = {};
  try { j = JSON.parse(res.getContentText() || '{}'); } catch (e) {}
  if (res.getResponseCode() !== 200 || !j.location) fail_('อ่านข้อมูลสถานที่ไม่ได้: ' + ((j.error && j.error.message) || ('HTTP ' + res.getResponseCode())).slice(0, 200));
  var lat = j.location.latitude, lng = j.location.longitude;
  return { name: (j.displayName && j.displayName.text) || '', address: j.formattedAddress || '', lat: lat, lng: lng,
           place_id: id, map_url: mapsUrl_(lat, lng, id) };
}

function listPlaces_(req) {
  auth_(req, 'admin');
  return rows_('Places').map(function (p) {
    return { id: p.id, name: p.name, address: p.address, lat: num_(p.lat), lng: num_(p.lng),
             map_url: p.map_url, used_count: Number(p.used_count) || 0, last_used: p.last_used };
  }).sort(function (a, b) { return b.used_count - a.used_count || (b.last_used > a.last_used ? 1 : -1); })
    .slice(0, 300);
}

function computeRoute_(stops, oldStops, job) {
  var pts = stops.filter(function (s) { return s.lat !== null && s.lng !== null; });
  if (pts.length < 2 || pts.length > 25) return null;
  var sig = pts.map(function (p) { return p.lat.toFixed(5) + ',' + p.lng.toFixed(5); }).join('|');
  if (job && job.route_polyline) {
    var oldSig = oldStops.slice().sort(function (a, b) { return Number(a.seq) - Number(b.seq); })
      .filter(function (s) { return num_(s.lat) !== null && num_(s.lng) !== null; })
      .map(function (s) { return Number(s.lat).toFixed(5) + ',' + Number(s.lng).toFixed(5); }).join('|');
    if (oldSig === sig) return { polyline: job.route_polyline, km: job.distance_km, min: job.duration_min };
  }
  try {
    var last = pts[pts.length - 1];
    var df = Maps.newDirectionFinder()
      .setOrigin(pts[0].lat, pts[0].lng)
      .setDestination(last.lat, last.lng)
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .setLanguage('th');
    pts.slice(1, -1).forEach(function (p) { df.addWaypoint(p.lat, p.lng); });
    var r = df.getDirections();
    if (!r || r.status !== 'OK' || !r.routes || !r.routes.length) return null;
    var rt = r.routes[0], meters = 0, secs = 0;
    rt.legs.forEach(function (l) { meters += l.distance.value; secs += l.duration.value; });
    return { polyline: rt.overview_polyline.points, km: Math.round(meters / 100) / 10, min: Math.round(secs / 60) };
  } catch (e) {
    return null;
  }
}

function parseLatLng_(s) {
  var m = String(s || '').match(/^\s*(-?\d{1,2}(?:\.\d+)?)\s*,\s*(-?\d{1,3}(?:\.\d+)?)\s*$/);
  if (!m) return null;
  var lat = Number(m[1]), lng = Number(m[2]);
  if (Math.abs(lat) > 90 || Math.abs(lng) > 180) return null;
  return { lat: lat, lng: lng };
}

function parseMapsUrl_(u) {
  var s = String(u || '');
  try { s = decodeURIComponent(s.replace(/\+/g, ' ')); } catch (e) {}
  var out = { name: '', lat: null, lng: null };
  var m = s.match(/!3d(-?\d+(?:\.\d+)?)!4d(-?\d+(?:\.\d+)?)/) ||
          s.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/) ||
          s.match(/[?&](?:q|query|ll|destination|daddr|center)=(-?\d+\.\d+)\s*,\s*(-?\d+\.\d+)/) ||
          s.match(/\/maps\/(?:search|place|dir)\/[^@]*?(-?\d{1,2}\.\d{3,}),\s*(-?\d{1,3}\.\d{3,})/);
  if (m) {
    var lat = Number(m[1]), lng = Number(m[2]);
    if (Math.abs(lat) <= 90 && Math.abs(lng) <= 180) { out.lat = lat; out.lng = lng; }
  }
  var n = s.match(/\/place\/([^\/@?&]+)/) || s.match(/[?&]q=([^&]+)/);
  if (n && !parseLatLng_(n[1])) out.name = n[1].trim().slice(0, 120);
  return out;
}

function isGoogleHost_(u) {
  var m = String(u).match(/^https?:\/\/([^\/?#:]+)/i);
  if (!m) return false;
  return /^(www\.|maps\.)?(google\.[a-z.]{2,7}|app\.goo\.gl|goo\.gl|g\.co|consent\.google\.com)$/i.test(m[1]) ||
         /^maps\.app\.goo\.gl$/i.test(m[1]);
}

function mapsUrl_(lat, lng, placeId) {
  return 'https://www.google.com/maps/search/?api=1&query=' + lat + ',' + lng +
    (placeId ? '&query_place_id=' + encodeURIComponent(placeId) : '');
}

function guessName_(result, q) {
  var types = result.types || [];
  var poi = types.some(function (t) { return /establishment|point_of_interest|premise|airport|park|university|school|hospital|shopping_mall/.test(t); });
  var c = result.address_components && result.address_components[0];
  if (poi && c && c.long_name && !/^\d/.test(c.long_name)) return c.long_name;
  return q;
}

function upsertPlaces_(stops) {
  var places = rows_('Places');
  stops.forEach(function (s) {
    var hit = places.filter(function (p) {
      if (p.name !== s.name) return false;
      if (s.lat !== null && num_(p.lat) !== null) {
        return Math.abs(num_(p.lat) - s.lat) < 0.0005 && Math.abs(num_(p.lng) - s.lng) < 0.0005;
      }
      return p.map_url === s.map_url;
    })[0];
    if (hit && hit._row < 0) return; // เพิ่งเพิ่มในรอบนี้
    if (hit) {
      update_('Places', hit, {
        used_count: (Number(hit.used_count) || 0) + 1, last_used: now_(),
        address: s.address || hit.address, map_url: s.map_url || hit.map_url
      });
    } else {
      var rec = { id: 'P' + shortId_(), name: s.name, address: s.address,
                  lat: s.lat === null ? '' : s.lat, lng: s.lng === null ? '' : s.lng,
                  map_url: s.map_url, used_count: 1, last_used: now_() };
      insert_('Places', rec);
      rec._row = -1;
      places.push(rec);
    }
  });
}

/* ====================== ตรวจความพร้อมประจำวัน ====================== */

function myDay_(req) {
  var s = auth_(req, 'driver');
  var today = today_();
  var drv = rows_('Drivers').filter(function (d) { return d.id === s.user_id; })[0] || {};
  var vehicles = rows_('Vehicles').filter(function (v) { return v.active === 'true'; });
  var mine = rows_('Jobs').filter(function (j) {
    return j.driver_id === s.user_id && j.date === today && j.status !== 'cancelled';
  }).sort(function (a, b) { return (a.time || '99:99') < (b.time || '99:99') ? -1 : 1; });
  var first = mine.filter(function (j) { return j.status !== 'done'; })[0] || mine[0] || null;
  var check = findDailyCheck_(s.user_id, today);
  var out = settingsOut_();
  out.date = today;
  out.driver = { id: s.user_id, name: drv.name || s.name, phone: drv.phone || '', photo_url: drv.photo_url || '' };
  out.check = check ? checkOut_(check, indexBy_(vehicles, 'id')) : null;
  out.vehicle_id = first ? first.vehicle_id : '';
  out.plate = first ? first.plate : '';
  out.jobs_today = mine.length;
  out.vehicles = vehicles.map(function (v) { return { id: v.id, plate: v.plate, name: v.name, seats: num_(v.seats), photo_url: v.photo_url }; })
    .sort(function (a, b) { return a.plate.localeCompare(b.plate, 'th'); });
  out.photo_labels = DAILY_PHOTOS;
  out.version = APP_VERSION;
  return out;
}

function saveDailyCheck_(req) {
  var s = auth_(req, 'driver');
  var list = Array.isArray(req.photos) ? req.photos : [];
  if (list.length !== DAILY_PHOTOS.length || list.some(function (p) { return !p; })) {
    fail_('กรุณาถ่ายรูปให้ครบ ' + DAILY_PHOTOS.length + ' รูป (' + DAILY_PHOTOS.join(', ') + ')');
  }
  var photos = list.map(function (p) {
    var m = String(p || '').match(/^data:image\/(jpeg|png|webp);base64,([A-Za-z0-9+\/=]+)$/);
    if (!m) fail_('รูปไม่ถูกต้อง กรุณาถ่ายใหม่');
    if (m[2].length > 10000000) fail_('รูปมีขนาดใหญ่เกินไป');
    return { ext: m[1], b64: m[2] };
  });
  var items = req.items || {};
  DAILY_ITEMS.forEach(function (k) { if (items[k] !== true) fail_('กรุณาติ๊กยืนยันให้ครบทุกข้อ'); });
  var vehicleId = String(req.vehicle_id || '');
  var phone = clean_(req.phone, 20);
  var lat = num_(req.lat), lng = num_(req.lng), acc = num_(req.acc);
  var today = today_();

  return withLock_(function () {
    if (findDailyCheck_(s.user_id, today)) fail_('วันนี้ตรวจความพร้อมไปแล้ว');
    var vehicles = rows_('Vehicles');
    var veh = vehicleId ? vehicles.filter(function (v) { return v.id === vehicleId; })[0] : null;
    if (vehicleId && !veh) fail_('ไม่พบรถที่เลือก');
    if (!veh && vehicles.some(function (v) { return v.active === 'true'; })) fail_('กรุณาเลือกรถที่ใช้วันนี้');
    var plate = veh ? veh.plate : clean_(req.plate, 20);
    var drv = rows_('Drivers').filter(function (d) { return d.id === s.user_id; })[0];
    if (drv && phone && phone !== drv.phone) update_('Drivers', drv, { phone: phone, updated_at: now_() });
    var saved = photos.map(function (p, i) {
      var f = savePhoto_(p.ext, p.b64, 'ตรวจรถ_' + today + '_' + s.name + '_' + DAILY_PHOTOS[i], 'w1200');
      return { label: DAILY_PHOTOS[i], id: f.id, url: f.url };
    });
    var rec = {
      id: 'C' + shortId_(), date: today, driver_id: s.user_id, driver_name: s.name,
      vehicle_id: veh ? veh.id : '', plate: plate, photo_urls: JSON.stringify(saved),
      items: DAILY_ITEMS.join(','), phone: phone || (drv ? drv.phone : ''),
      lat: lat === null ? '' : lat, lng: lng === null ? '' : lng, acc_m: acc === null ? '' : Math.round(acc),
      created_at: now_()
    };
    insert_('DailyChecks', rec);
    if (lat !== null && lng !== null) addLocation_(s, lat, lng, acc, 'daily_check', '');
    log_(s.name, 'daily_check', (plate || '-') + ' · ' + saved.length + ' รูป');
    return checkOut_(rec, indexBy_(vehicles, 'id'));
  });
}

function findDailyCheck_(driverId, date) {
  return rows_('DailyChecks').filter(function (c) { return c.driver_id === driverId && c.date === date; })[0] || null;
}

function checkOut_(c, vehicles) {
  var photos = [];
  try { photos = c.photo_urls ? JSON.parse(c.photo_urls) : []; } catch (e) { photos = []; }
  var v = vehicles && c.vehicle_id ? vehicles[c.vehicle_id] : null;
  return {
    id: c.id, date: c.date, driver_id: c.driver_id, driver_name: c.driver_name,
    vehicle_id: c.vehicle_id, plate: c.plate, vehicle_name: v ? v.name : '',
    photos: photos.map(function (p) { return { label: p.label, url: p.url }; }),
    phone: c.phone, lat: num_(c.lat), lng: num_(c.lng), acc_m: num_(c.acc_m), created_at: c.created_at
  };
}

/* ====================== ตำแหน่งคนขับ (ส่งเฉพาะตอนเปิดแอป) ====================== */

function reportLocation_(req) {
  var s = auth_(req, 'driver');
  var lat = num_(req.lat), lng = num_(req.lng), acc = num_(req.acc);
  if (lat === null || lng === null || Math.abs(lat) > 90 || Math.abs(lng) > 180) fail_('ตำแหน่งไม่ถูกต้อง');
  if (!gpsOn_()) return { saved: false, reason: 'off' };
  var cache = CacheService.getScriptCache();
  var k = 'locat:' + s.user_id;
  if (cache.get(k) && !req.force) return { saved: false, reason: 'too_soon', interval_min: gpsMin_() };
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) return { saved: false, reason: 'busy' };
  try {
    addLocation_(s, lat, lng, acc, req.source === 'manual' ? 'manual' : 'gps', clean_(req.job_id, 30));
  } finally { lock.releaseLock(); }
  cache.put(k, '1', 40);
  return { saved: true, at: now_(), interval_min: gpsMin_() };
}

function addLocation_(s, lat, lng, acc, source, jobId) {
  insert_('Locations', {
    at: now_(), driver_id: s.user_id, name: s.name, lat: Math.round(lat * 1e6) / 1e6, lng: Math.round(lng * 1e6) / 1e6,
    acc_m: acc === null || acc === undefined ? '' : Math.round(acc), source: source, job_id: jobId || ''
  });
  var sh = sheet_('Locations');
  if (sh.getLastRow() > LOCATION_KEEP_ROWS + 2000) sh.deleteRows(2, 2000);
}

/* ภาพรวมประจำวันสำหรับแอดมิน: ใครตรวจความพร้อมแล้ว + ตำแหน่งล่าสุดของรถแต่ละคัน */
function dayStatus_(req) {
  auth_(req, 'admin');
  var date = isDate_(req.date) ? String(req.date) : today_();
  var vehicles = indexBy_(rows_('Vehicles'), 'id');
  var checks = rows_('DailyChecks').filter(function (c) { return c.date === date; })
    .map(function (c) { return checkOut_(c, vehicles); });
  var locs = date === today_() ? rowsTail_('Locations', 4000) : rows_('Locations');
  locs = locs.filter(function (l) { return String(l.at).slice(0, 10) === date; });
  var latest = {}, trail = [];
  locs.forEach(function (l) {
    var p = { driver_id: l.driver_id, name: l.name, lat: num_(l.lat), lng: num_(l.lng), acc_m: num_(l.acc_m), at: l.at, source: l.source };
    if (!latest[l.driver_id] || latest[l.driver_id].at <= l.at) latest[l.driver_id] = p;
    if (req.trail && l.driver_id === String(req.trail)) trail.push([p.lat, p.lng, l.at]);
  });
  var drivers = rows_('Drivers').filter(function (d) { return d.active === 'true'; }).map(function (d) {
    return { id: d.id, name: d.name, phone: d.phone, photo_url: d.photo_url };
  });
  return {
    date: date, daily_check: dailyCheckOn_(), gps_tracking: gpsOn_(), gps_interval_min: gpsMin_(),
    checks: checks, drivers: drivers,
    locations: Object.keys(latest).map(function (k) { return latest[k]; }),
    trail: trail.slice(-600)
  };
}

/* บันทึกปัญหาที่เกิดบนมือถือ (เช่น เชื่อมต่อหลุด) เพื่อดูสาเหตุภายหลังในแผ่นงาน Log */
function clientLog_(req) {
  var cache = CacheService.getScriptCache();
  var n = Number(cache.get('clientlog') || 0);
  if (n >= 40) return false;
  cache.put('clientlog', String(n + 1), 600);
  var who = 'ไม่ทราบ';
  try { if (req.token) who = auth_(req).name; } catch (e) {}
  log_(who, 'client_error', (clean_(req.msg, 300) + ' · ' + clean_(req.where, 60) + ' · ' + clean_(req.ua, 160)).slice(0, 600));
  return true;
}

/* ====================== ข้อมูล Job ====================== */

function getJobById_(id) {
  var list = loadJobs_(function (j) { return j.id === id; });
  return list[0] || null;
}

function loadJobs_(filter) {
  var radius = radius_();
  var vehicles = indexBy_(rows_('Vehicles'), 'id');
  var drivers = indexBy_(rows_('Drivers'), 'id');
  var jobs = rows_('Jobs').filter(filter);
  var ids = {};
  jobs.forEach(function (j) { ids[j.id] = true; });
  var byJob = {};
  rows_('Stops').forEach(function (s) {
    if (!ids[s.job_id]) return;
    (byJob[s.job_id] = byJob[s.job_id] || []).push(s);
  });
  return jobs.map(function (j) {
    var stops = (byJob[j.id] || []).map(stopOut_).sort(function (a, b) { return a.seq - b.seq; });
    var out = {};
    SCHEMA.Jobs.forEach(function (c) { out[c] = j[c]; });
    out.distance_km = num_(j.distance_km);
    out.duration_min = num_(j.duration_min);
    out.driver_name = (drivers[j.driver_id] || {}).name || '(ไม่พบคนขับ)';
    out.driver_photo_url = (drivers[j.driver_id] || {}).photo_url || '';
    var veh = j.vehicle_id ? vehicles[j.vehicle_id] : null;
    out.vehicle_name = veh ? veh.name : '';
    out.vehicle_photo_url = veh ? veh.photo_url : '';
    out.vehicle_seats = veh ? num_(veh.seats) : null;
    out.stops = stops;
    out.total_count = stops.length;
    out.done_count = stops.filter(function (s) { return s.status === 'arrived'; }).length;
    out.checkin_radius_m = radius;
    return out;
  }).sort(function (a, b) {
    return (a.date + (a.time || '99:99') + a.job_no) < (b.date + (b.time || '99:99') + b.job_no) ? -1 : 1;
  });
}

function stopOut_(s) {
  var o = {};
  SCHEMA.Stops.forEach(function (c) { o[c] = s[c]; });
  o.seq = Number(s.seq) || 0;
  ['lat', 'lng', 'arrive_lat', 'arrive_lng', 'checkin_dist_m', 'checkin_acc_m'].forEach(function (k) { o[k] = num_(s[k]); });
  var photos = [];
  try { photos = s.photo_urls ? JSON.parse(s.photo_urls) : []; } catch (e) { photos = []; }
  if (!photos.length && s.photo_url) photos = [{ id: s.photo_id, url: s.photo_url }];
  o.photos = photos.map(function (p) { return p.url; });
  delete o.photo_urls;
  return o;
}

function nextJobNo_(jobs, date) {
  var prefix = 'JOB-' + date.slice(2).replace(/-/g, '') + '-';
  var max = 0;
  jobs.forEach(function (j) {
    if (j.job_no.indexOf(prefix) === 0) max = Math.max(max, Number(j.job_no.slice(prefix.length)) || 0);
  });
  var n = String(max + 1);
  return prefix + (n.length < 2 ? '0' + n : n);
}

/* ====================== สิทธิ์และเซสชัน ====================== */

function auth_(req, role) {
  var t = String(req.token || '');
  if (!t) fail_('กรุณาเข้าสู่ระบบ', 'AUTH');
  var cache = CacheService.getScriptCache();
  var raw = cache.get('sess:' + t);
  var s = raw ? JSON.parse(raw) : null;
  if (!s) {
    var row = rows_('Sessions').filter(function (x) { return x.token === t; })[0];
    if (!row) fail_('หมดเวลาการใช้งาน กรุณาเข้าสู่ระบบใหม่', 'AUTH');
    s = { token: row.token, role: row.role, user_id: row.user_id, name: row.name, expires_at: row.expires_at };
    cache.put('sess:' + t, JSON.stringify(s), 600);
  }
  if (Number(s.expires_at) < Date.now()) fail_('หมดเวลาการใช้งาน กรุณาเข้าสู่ระบบใหม่', 'AUTH');
  if (role && s.role !== role) fail_('ไม่มีสิทธิ์ใช้งานส่วนนี้', 'FORBIDDEN');
  return s;
}

function createSession_(role, user) {
  var token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  insert_('Sessions', {
    token: token, role: role, user_id: user.id, name: user.name, created_at: now_(),
    expires_at: String(Date.now() + SESSION_DAYS[role] * 86400000)
  });
  return token;
}

function killSessions_(pred) {
  var list = rows_('Sessions').filter(pred);
  var cache = CacheService.getScriptCache();
  list.forEach(function (s) { cache.remove('sess:' + s.token); });
  deleteRows_('Sessions', list.map(function (s) { return s._row; }));
}

function cleanupSessions_() {
  var now = Date.now();
  var old = rows_('Sessions').filter(function (s) { return Number(s.expires_at) < now; });
  if (old.length) deleteRows_('Sessions', old.map(function (s) { return s._row; }));
}

function checkLock_(key) {
  var n = Number(CacheService.getScriptCache().get('fail:' + key) || 0);
  if (n >= MAX_FAILS) fail_('ใส่ PIN ผิดเกิน ' + MAX_FAILS + ' ครั้ง กรุณารอ ' + LOCK_MINUTES + ' นาทีแล้วลองใหม่', 'LOCKED');
}

function badPin_(key, msg) {
  var cache = CacheService.getScriptCache();
  var n = Number(cache.get('fail:' + key) || 0) + 1;
  cache.put('fail:' + key, String(n), LOCK_MINUTES * 60);
  var left = MAX_FAILS - n;
  fail_((msg || 'PIN ไม่ถูกต้อง') + (left > 0 ? ' (ลองได้อีก ' + left + ' ครั้ง)' : ' — ระบบล็อก ' + LOCK_MINUTES + ' นาที'), 'BAD_PIN');
}

function clearFail_(key) { CacheService.getScriptCache().remove('fail:' + key); }

/* ====================== ตั้งค่าและตาราง ====================== */

function ensureSetup_() {
  var cache = CacheService.getScriptCache();
  if (cache.get(SETUP_KEY)) return;
  var ss = ss_();
  Object.keys(SCHEMA).forEach(function (name) {
    var cols = SCHEMA[name];
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    var head = sh.getRange(1, 1, 1, cols.length);
    var cur = head.getValues()[0].map(String);
    if (cur.join('|') !== cols.join('|')) {
      head.setNumberFormat('@').setValues([cols]);
      head.setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  if (!getSetting_('admin_pin_hash')) {
    setSetting_('salt', Utilities.getUuid());
    setSetting_('admin_pin_hash', hash_('admin:' + DEFAULT_ADMIN_PIN));
    setSetting_('admin_pin_default', 'true');
    if (!getSetting_('company_name')) setSetting_('company_name', 'บริษัทของฉัน');
  }
  if (getSetting_('checkin_radius_m') === '') setSetting_('checkin_radius_m', String(DEFAULT_RADIUS_M));
  cache.put(SETUP_KEY, '1', 21600);
}

function getSetting_(key) {
  if (!REQ.settings) {
    REQ.settings = {};
    rows_('Settings').forEach(function (r) { REQ.settings[r.key] = r; });
  }
  return REQ.settings[key] ? REQ.settings[key].value : '';
}

function setSetting_(key, value) {
  getSetting_(key);
  var row = REQ.settings[key];
  if (row) {
    update_('Settings', row, { value: value });
  } else {
    insert_('Settings', { key: key, value: value });
    REQ.settings = null;
  }
}

function ss_() {
  if (REQ.ss) return REQ.ss;
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    if (!id) fail_('ไม่พบ Google Sheet — ต้องสร้าง Apps Script จากเมนู ส่วนขยาย ของ Google Sheet');
    active = SpreadsheetApp.openById(id);
  }
  REQ.ss = active;
  return active;
}

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) fail_('ไม่พบแผ่นงาน ' + name);
  return sh;
}

function rows_(name) {
  var sh = sheet_(name), cols = SCHEMA[name], last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, cols.length).getValues().map(function (r, i) {
    var o = { _row: i + 2 };
    cols.forEach(function (c, j) { o[c] = r[j] === null || r[j] === undefined ? '' : String(r[j]); });
    return o;
  }).filter(function (o) { return o[cols[0]] !== ''; });
}

function rowsTail_(name, n) {
  var sh = sheet_(name), cols = SCHEMA[name], last = sh.getLastRow();
  if (last < 2) return [];
  var start = Math.max(2, last - n + 1);
  return sh.getRange(start, 1, last - start + 1, cols.length).getValues().map(function (r, i) {
    var o = { _row: start + i };
    cols.forEach(function (c, j) { o[c] = r[j] === null || r[j] === undefined ? '' : String(r[j]); });
    return o;
  }).filter(function (o) { return o[cols[0]] !== ''; });
}

function cell_(v) {
  var s = v === null || v === undefined ? '' : String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s; // กันสูตรใน Sheet
}

function insert_(name, obj) {
  var sh = sheet_(name), cols = SCHEMA[name];
  var row = cols.map(function (c) { return cell_(obj[c]); });
  sh.getRange(sh.getLastRow() + 1, 1, 1, cols.length).setNumberFormat('@').setValues([row]);
  if (name === 'Settings') REQ.settings = null;
}

function update_(name, rowObj, patch) {
  var cols = SCHEMA[name];
  Object.keys(patch).forEach(function (k) { rowObj[k] = patch[k] === null || patch[k] === undefined ? '' : String(patch[k]); });
  var row = cols.map(function (c) { return cell_(rowObj[c]); });
  sheet_(name).getRange(rowObj._row, 1, 1, cols.length).setNumberFormat('@').setValues([row]);
}

function deleteRows_(name, rowNums) {
  var sh = sheet_(name);
  rowNums.filter(function (r) { return r > 1; }).sort(function (a, b) { return b - a; })
    .forEach(function (r) { sh.deleteRow(r); });
}

function log_(who, action, detail) {
  var sh = ss_().getSheetByName('Log');
  if (!sh) return;
  insert_('Log', { time: now_(), who: who, action: action, detail: detail });
  if (sh.getLastRow() > 8000) sh.deleteRows(2, 2000);
}

/* ====================== ตัวช่วย ====================== */

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) fail_('ระบบกำลังยุ่ง กรุณาลองใหม่อีกครั้ง');
  try { return fn(); } finally { lock.releaseLock(); }
}

function hash_(s) {
  var salt = getSetting_('salt') || '';
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + s, Utilities.Charset.UTF_8);
  return bytes.map(function (b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function hex_(s) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8);
  return bytes.map(function (b) { var v = (b < 0 ? b + 256 : b).toString(16); return v.length === 1 ? '0' + v : v; }).join('').slice(0, 40);
}

function bkk_() {
  var t = new Date(Date.now() + TZ_OFFSET_HOURS * 3600000);
  function p(n) { return (n < 10 ? '0' : '') + n; }
  return {
    date: t.getUTCFullYear() + '-' + p(t.getUTCMonth() + 1) + '-' + p(t.getUTCDate()),
    time: p(t.getUTCHours()) + ':' + p(t.getUTCMinutes()) + ':' + p(t.getUTCSeconds())
  };
}
function now_() { var b = bkk_(); return b.date + ' ' + b.time; }
function today_() { return bkk_().date; }
function addDays_(date, n) {
  var d = new Date(date + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}
function isDate_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')); }
function num_(v) {
  if (v === '' || v === null || v === undefined) return null;
  var n = Number(v);
  return isFinite(n) ? n : null;
}
function clean_(v, max) { return String(v === null || v === undefined ? '' : v).trim().slice(0, max || 200); }
function safeUrl_(u) {
  var s = String(u || '').trim();
  return /^https:\/\/[^\s"'<>]+$/i.test(s) && s.length <= 1500 ? s : '';
}
function shortId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function indexBy_(list, key) { var o = {}; list.forEach(function (x) { o[x[key]] = x; }); return o; }
function fail_(msg, code) { var e = new Error(msg); e.isApp = true; e.code = code || ''; throw e; }
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
