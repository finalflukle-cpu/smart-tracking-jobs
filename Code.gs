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

var APP_VERSION = '1.0.0';
var DEFAULT_ADMIN_PIN = '000000';
var SESSION_DAYS = { admin: 7, driver: 30 };
var MAX_FAILS = 5;
var LOCK_MINUTES = 15;
var TZ_OFFSET_HOURS = 7; // เวลาประเทศไทย

var SCHEMA = {
  Settings: ['key', 'value'],
  Drivers: ['id', 'name', 'pin_hash', 'active', 'created_at', 'updated_at'],
  Jobs: ['id', 'job_no', 'date', 'time', 'title', 'driver_id', 'plate', 'note', 'status',
         'route_polyline', 'distance_km', 'duration_min',
         'created_at', 'updated_at', 'started_at', 'finished_at'],
  Stops: ['id', 'job_id', 'seq', 'name', 'address', 'lat', 'lng', 'map_url', 'note', 'planned_time',
          'status', 'arrived_at', 'arrive_lat', 'arrive_lng', 'photo_id', 'photo_url', 'checkin_note'],
  Places: ['id', 'name', 'address', 'lat', 'lng', 'map_url', 'used_count', 'last_used'],
  Sessions: ['token', 'role', 'user_id', 'name', 'created_at', 'expires_at'],
  Log: ['time', 'who', 'action', 'detail']
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
  CacheService.getScriptCache().remove('setup_ok_v1');
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
    return { app: 'Smart Tracking Jobs', version: APP_VERSION, company: getSetting_('company_name') || '' };
  },
  publicDrivers: function () {
    return rows_('Drivers').filter(function (d) { return d.active === 'true'; })
      .map(function (d) { return { id: d.id, name: d.name }; })
      .sort(function (a, b) { return a.name.localeCompare(b.name, 'th'); });
  },
  login: login_,
  logout: logout_,
  me: me_,
  changeAdminPin: changeAdminPin_,
  saveSettings: saveSettings_,
  listDrivers: listDrivers_,
  saveDriver: saveDriver_,
  listJobs: listJobs_,
  getJob: getJob_,
  saveJob: saveJob_,
  cancelJob: cancelJob_,
  deleteJob: deleteJob_,
  checkIn: checkIn_,
  searchPlace: searchPlace_,
  resolveLink: resolveLink_,
  listPlaces: listPlaces_
};

/* ====================== เข้าสู่ระบบ ====================== */

function login_(req) {
  var pin = String(req.pin || '');
  if (!/^\d{6}$/.test(pin)) fail_('PIN ต้องเป็นตัวเลข 6 หลัก');
  cleanupSessions_();
  var company = getSetting_('company_name') || '';

  if (req.role === 'admin') {
    checkLock_('admin');
    if (hash_('admin:' + pin) !== getSetting_('admin_pin_hash')) badPin_('admin');
    clearFail_('admin');
    var admin = { id: 'admin', name: 'แอดมิน' };
    var token = createSession_('admin', admin);
    log_('admin', 'login', '');
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
    var user = { id: d.id, name: d.name };
    var t = createSession_('driver', user);
    log_(d.name, 'login', '');
    return { token: t, role: 'driver', user: user, company: company };
  }
  fail_('ประเภทผู้ใช้ไม่ถูกต้อง');
}

function logout_(req) {
  var t = String(req.token || '');
  if (t) killSessions_(function (s) { return s.token === t; });
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
  killSessions_(function (x) { return x.role === 'admin' && x.token !== s.token; });
  log_('admin', 'change_pin', '');
  return true;
}

function saveSettings_(req) {
  auth_(req, 'admin');
  var name = clean_(req.company_name, 80);
  if (!name) fail_('กรุณาใส่ชื่อบริษัท');
  setSetting_('company_name', name);
  return { company: name };
}

/* ====================== คนขับ ====================== */

function listDrivers_(req) {
  auth_(req, 'admin');
  return rows_('Drivers').map(function (d) {
    return { id: d.id, name: d.name, active: d.active === 'true', created_at: d.created_at };
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
      if (pin) patch.pin_hash = hash_('driver:' + cur.id + ':' + pin);
      update_('Drivers', cur, patch);
      if (pin || !active) killSessions_(function (s) { return s.role === 'driver' && s.user_id === cur.id; });
      if (pin) clearFail_('d:' + cur.id);
      log_('admin', 'update_driver', name + (pin ? ' (เปลี่ยน PIN)' : '') + (active ? '' : ' (ปิดใช้งาน)'));
      return { id: cur.id, name: name, active: active };
    }

    if (!pin) fail_('กรุณาตั้ง PIN 6 หลักให้คนขับ');
    var id = 'D' + shortId_();
    insert_('Drivers', {
      id: id, name: name, pin_hash: hash_('driver:' + id + ':' + pin),
      active: 'true', created_at: now_(), updated_at: now_()
    });
    log_('admin', 'add_driver', name);
    return { id: id, name: name, active: true };
  });
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
  if (!isDate_(date)) fail_('กรุณาเลือกวันที่');
  if (time && !/^\d{2}:\d{2}$/.test(time)) fail_('เวลาเริ่มไม่ถูกต้อง');
  if (!driverId) fail_('กรุณาเลือกคนขับ');
  if (!plate) fail_('กรุณาใส่ทะเบียนรถ');
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
        photo_id: o ? o.photo_id : '', photo_url: o ? o.photo_url : '', checkin_note: o ? o.checkin_note : ''
      });
    });

    var status = arrived === 0 ? 'pending' : (arrived === clean.length ? 'done' : 'in_progress');
    var rec = {
      id: jobId, job_no: jobNo, date: date, time: time, title: clean_(inp.title, 120),
      driver_id: drv.id, plate: plate, note: clean_(inp.note, 500), status: status,
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
  var m = String(req.photo || '').match(/^data:image\/(jpeg|png|webp);base64,([A-Za-z0-9+\/=]+)$/);
  if (!m) fail_('กรุณาถ่ายรูปเช็คอิน');
  if (m[2].length > 10000000) fail_('รูปมีขนาดใหญ่เกินไป');

  return withLock_(function () {
    var stop = rows_('Stops').filter(function (x) { return x.id === String(req.stopId || ''); })[0];
    if (!stop) fail_('ไม่พบสถานที่นี้ใน Job');
    var job = rows_('Jobs').filter(function (j) { return j.id === stop.job_id; })[0];
    if (!job || job.driver_id !== s.user_id) fail_('สถานที่นี้ไม่ใช่งานของคุณ', 'FORBIDDEN');
    if (job.status === 'cancelled') fail_('Job นี้ถูกยกเลิกแล้ว');
    if (stop.status === 'arrived') fail_('จุดนี้เช็คอินไปแล้ว');

    var file = savePhoto_(m[1], m[2], job.job_no + '_จุด' + stop.seq);
    var lat = num_(req.lat), lng = num_(req.lng);
    var t = now_();
    update_('Stops', stop, {
      status: 'arrived', arrived_at: t,
      arrive_lat: lat === null ? '' : lat, arrive_lng: lng === null ? '' : lng,
      photo_id: file.id, photo_url: file.url, checkin_note: clean_(req.note, 300)
    });
    var all = rows_('Stops').filter(function (x) { return x.job_id === job.id; });
    var done = all.filter(function (x) { return x.status === 'arrived'; }).length;
    var patch = { updated_at: t, status: done >= all.length ? 'done' : 'in_progress' };
    if (!job.started_at) patch.started_at = t;
    if (patch.status === 'done') patch.finished_at = t;
    update_('Jobs', job, patch);
    log_(s.name, 'check_in', job.job_no + ' จุด ' + stop.seq + ' ' + stop.name);
    return getJobById_(job.id);
  });
}

function savePhoto_(ext, b64, base) {
  var folder = photoFolder_();
  var bytes = Utilities.base64Decode(b64);
  var name = base + '_' + now_().replace(/[^0-9]/g, '') + '.' + (ext === 'jpeg' ? 'jpg' : ext);
  var file = folder.createFile(Utilities.newBlob(bytes, 'image/' + ext, name));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  var id = file.getId();
  return { id: id, url: 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1600' };
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
  var r = Maps.newGeocoder().setRegion('th').setLanguage('th').geocode(q);
  if (!r || r.status === 'ZERO_RESULTS') return [];
  if (r.status !== 'OK') fail_('ค้นหาไม่สำเร็จ (' + r.status + ') ลองวางลิงก์ Google Maps แทน');
  return (r.results || []).slice(0, 6).map(function (x) {
    var loc = x.geometry.location;
    return {
      name: guessName_(x, q), address: x.formatted_address || '',
      lat: loc.lat, lng: loc.lng, map_url: mapsUrl_(loc.lat, loc.lng, x.place_id)
    };
  });
}

function resolveLink_(req) {
  auth_(req, 'admin');
  var url = String(req.url || '').trim();
  var co = parseLatLng_(url);
  if (co) return { name: '', lat: co.lat, lng: co.lng, map_url: mapsUrl_(co.lat, co.lng) };
  var mUrl = url.match(/https:\/\/[^\s]+/i); // รองรับข้อความที่แชร์มาพร้อมลิงก์
  if (!mUrl) fail_('กรุณาวางลิงก์จาก Google Maps (ขึ้นต้นด้วย https://)');
  url = mUrl[0];
  var cur = url;
  for (var i = 0; i < 6; i++) {
    if (!isGoogleHost_(cur)) fail_('รองรับเฉพาะลิงก์จาก Google Maps');
    var info = parseMapsUrl_(cur);
    if (info.lat !== null) return { name: info.name, lat: info.lat, lng: info.lng, map_url: url };
    var res = UrlFetchApp.fetch(cur, { followRedirects: false, muteHttpExceptions: true });
    var h = res.getHeaders() || {};
    var loc = h.Location || h.location;
    if (!loc) {
      var body = parseMapsUrl_(String(res.getContentText() || '').slice(0, 300000));
      if (body.lat !== null) return { name: info.name || body.name, lat: body.lat, lng: body.lng, map_url: url };
      break;
    }
    if (Array.isArray(loc)) loc = loc[0];
    cur = /^https?:\/\//i.test(loc) ? loc : cur.replace(/^(https?:\/\/[^\/]+).*$/i, '$1') + loc;
  }
  var last = parseMapsUrl_(cur);
  return { name: last.name || '', lat: last.lat, lng: last.lng, map_url: url };
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

/* ====================== ข้อมูล Job ====================== */

function getJobById_(id) {
  var list = loadJobs_(function (j) { return j.id === id; });
  return list[0] || null;
}

function loadJobs_(filter) {
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
    out.stops = stops;
    out.total_count = stops.length;
    out.done_count = stops.filter(function (s) { return s.status === 'arrived'; }).length;
    return out;
  }).sort(function (a, b) {
    return (a.date + (a.time || '99:99') + a.job_no) < (b.date + (b.time || '99:99') + b.job_no) ? -1 : 1;
  });
}

function stopOut_(s) {
  var o = {};
  SCHEMA.Stops.forEach(function (c) { o[c] = s[c]; });
  o.seq = Number(s.seq) || 0;
  ['lat', 'lng', 'arrive_lat', 'arrive_lng'].forEach(function (k) { o[k] = num_(s[k]); });
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
  if (cache.get('setup_ok_v1')) return;
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
  cache.put('setup_ok_v1', '1', 21600);
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
