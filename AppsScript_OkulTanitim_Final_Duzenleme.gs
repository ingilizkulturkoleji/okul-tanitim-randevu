/**
 * İngiliz Kültür Koleji - Okul Tanıtım Randevu Sistemi
 * Veli e-posta + İngiliz tren/peron bileti tarzı otomatik onay maili eklenmiş sürüm.
 * Web App dağıtımı: Execute as Me, Anyone access.
 */

var TZ = 'Europe/Istanbul';
var DEFAULT_EMAILS = ['info@ingilizkultur.com.tr', 'batikent@ingilizkultur.com.tr'];
var SHEETS = {
  BOOKINGS: 'Randevular',
  SETTINGS: 'Ayarlar',
  ROOMS: 'Birimler',
  SLOTS: 'Saatler',
  CLOSED: 'KapaliTarihler',
  ROOM_DATES: 'OdaTarihDurumu'
};

function doGet(e) {
  try {
    ensureSheets_();
    var action = (e.parameter.action || 'availability').trim();
    var response;

    if (action === 'adminData') {
      response = getAdminData_(e.parameter.pin);
    } else if (action === 'adminSave') {
      var savePayload = parsePayload_(e);
      response = withLock_(function(){ return saveAdmin_(savePayload); });
    } else if (action === 'book') {
      var bookPayload = parsePayload_(e);
      response = withLock_(function(){ return bookAppointment_(bookPayload); });
    } else {
      response = getAvailability_(e.parameter.date, e.parameter.room);
    }
    return output_(response, e.parameter.callback);
  } catch (err) {
    return output_({ result: 'error', message: err.toString() }, e && e.parameter && e.parameter.callback);
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch (err) { return json_({ result: 'error', message: 'Sunucu yoğun. Lütfen birkaç saniye sonra tekrar deneyin.' }); }

  try {
    ensureSheets_();
    var data = JSON.parse(e.postData.contents || '{}');
    if (data.action === 'adminSave') return json_(saveAdmin_(data));
    return json_(bookAppointment_(data));
  } catch (err) {
    return json_({ result: 'error', message: err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function parsePayload_(e) {
  if (e.parameter && e.parameter.payload) return JSON.parse(e.parameter.payload);
  return e.parameter || {};
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch (err) { return { result: 'error', message: 'Sunucu yoğun. Lütfen birkaç saniye sonra tekrar deneyin.' }; }
  try { return fn(); }
  finally { try { lock.releaseLock(); } catch (e2) {} }
}

function setupWorkbook() {
  ensureSheets_();
  return 'Tablo yapısı hazırlandı.';
}

function ensureSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var specs = {};
  specs[SHEETS.BOOKINGS] = ['KayitZamani','VeliAdiSoyadi','OgrenciAdiSoyadi','Kademe','Telefon','VeliEposta','GorusmeBirimi','RandevuTarihi','RandevuSaati','Durum','Kaynak','Kampanya','ReklamSeti','Reklam','Notlar'];
  specs[SHEETS.SETTINGS] = ['Anahtar','Deger','Aciklama'];
  specs[SHEETS.ROOMS] = ['Kod','GorunenAd','Aktif'];
  specs[SHEETS.SLOTS] = ['Saat','Aktif'];
  specs[SHEETS.CLOSED] = ['Tarih','Aciklama'];
  specs[SHEETS.ROOM_DATES] = ['Tarih','OdaKodu','Durum','Aciklama'];

  Object.keys(specs).forEach(function(name){
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,specs[name].length).setValues([specs[name]]);
    } else {
      var currentHeaders = sh.getRange(1,1,1,Math.max(sh.getLastColumn(), specs[name].length)).getValues()[0];
      if (name === SHEETS.BOOKINGS && currentHeaders.indexOf('VeliEposta') === -1) {
        sh.insertColumnAfter(5);
        sh.getRange(1,6).setValue('VeliEposta');
      }
    }
  });
  seedDefaults_();
}

function seedDefaults_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = ss.getSheetByName(SHEETS.SETTINGS);
  if (settings.getLastRow() < 2) {
    settings.getRange(2,1,10,3).setValues([
      ['admin_pin','1234','/admin giriş PIN. İlk kurulumdan sonra değiştirin.'],
      ['system_active','TRUE','TRUE/FALSE'],
      ['sunday_closed','TRUE','Pazar günleri kapalı mı?'],
      ['min_lead_hours','2','Randevuya minimum kaç saat kala alınabilir?'],
      ['admin_emails', DEFAULT_EMAILS.join(','), 'Bildirim e-postaları'],
      ['success_message','Randevu talebiniz alınmıştır. Randevu biletiniz e-posta adresinize gönderilmiştir.','Veliye gösterilecek başarı mesajı'],
      ['school_phone','444 9 507','İletişim telefonu'],
      ['parent_mail_enabled','TRUE','Veliye otomatik randevu bileti maili gönderilsin mi? TRUE/FALSE'],
      ['ticket_brand_note','İngiliz Kültür Koleji Tanıtım Peronu','Veliye giden bilet mailinde küçük marka notu'],
      ['mail_reply_to','batikent@ingilizkultur.com.tr','Veli mailinde cevap adresi olarak gösterilecek adres']
    ]);
  } else {
    ensureSettingRow_('parent_mail_enabled','TRUE','Veliye otomatik randevu bileti maili gönderilsin mi? TRUE/FALSE');
    ensureSettingRow_('ticket_brand_note','İngiliz Kültür Koleji Tanıtım Peronu','Veliye giden bilet mailinde küçük marka notu');
    ensureSettingRow_('mail_reply_to','batikent@ingilizkultur.com.tr','Veli mailinde cevap adresi olarak gösterilecek adres');
  }
  var rooms = ss.getSheetByName(SHEETS.ROOMS);
  var defaultRooms = [
    ['anaokulu','Anaokulu',true],
    ['ilkokul','İlkokul',true],
    ['ortaokul','Ortaokul',true]
  ];
  if (rooms.getLastRow() < 2) {
    rooms.getRange(2,1,defaultRooms.length,3).setValues(defaultRooms);
  } else {
    migrateDefaultRooms_(rooms, defaultRooms);
  }
  var slots = ss.getSheetByName(SHEETS.SLOTS);
  if (slots.getLastRow() < 2) slots.getRange(2,1,8,2).setValues([
    ['09:00',true],['10:00',true],['11:00',true],['12:00',true],['13:30',true],['14:30',true],['15:30',true],['16:30',true]
  ]);
}

function ensureSettingRow_(key, value, description) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
  var rows = getRows_(SHEETS.SETTINGS);
  var exists = rows.some(function(r){ return String(r[0]||'').trim() === key; });
  if (!exists) sh.appendRow([key, value, description]);
}

function migrateDefaultRooms_(sh, defaultRooms) {
  var rows = getRows_(SHEETS.ROOMS);
  var codes = rows.map(function(r){ return String(r[0] || '').trim().toLowerCase(); });
  var hasOldRoomNames = codes.indexOf('oxford') !== -1 || codes.indexOf('cambridge') !== -1;
  var missingSchoolLevels = codes.indexOf('anaokulu') === -1 || codes.indexOf('ilkokul') === -1 || codes.indexOf('ortaokul') === -1;
  if (hasOldRoomNames || missingSchoolLevels) {
    resetSheet_(sh, ['Kod','GorunenAd','Aktif'], defaultRooms);
  }
}

function getConfig_() {
  var rows = getRows_(SHEETS.SETTINGS);
  var obj = {};
  rows.forEach(function(r){ if (r[0]) obj[String(r[0]).trim()] = r[1]; });
  return {
    adminPin: String(obj.admin_pin || '1234'),
    systemActive: String(obj.system_active || 'TRUE').toUpperCase() !== 'FALSE',
    sundayClosed: String(obj.sunday_closed || 'TRUE').toUpperCase() !== 'FALSE',
    minLeadHours: Number(obj.min_lead_hours || 0),
    adminEmails: String(obj.admin_emails || DEFAULT_EMAILS.join(',')).split(',').map(function(x){return x.trim();}).filter(Boolean),
    successMessage: String(obj.success_message || 'Randevu talebiniz alınmıştır. Randevu biletiniz e-posta adresinize gönderilmiştir.'),
    parentMailEnabled: String(obj.parent_mail_enabled || 'TRUE').toUpperCase() !== 'FALSE',
    ticketBrandNote: String(obj.ticket_brand_note || 'İngiliz Kültür Koleji Tanıtım Peronu'),
    mailReplyTo: String(obj.mail_reply_to || 'batikent@ingilizkultur.com.tr')
  };
}

function getRooms_() {
  return getRows_(SHEETS.ROOMS).map(function(r){ return { code:String(r[0]||'').trim(), name:String(r[1]||'').trim(), active:String(r[2]).toUpperCase() !== 'FALSE' }; }).filter(function(x){ return x.code && x.name; });
}
function getRoomName_(code) {
  var rooms = getRooms_();
  for (var i=0; i<rooms.length; i++) if (rooms[i].code === code) return rooms[i].name;
  return code || '';
}
function getSlots_() {
  return getRows_(SHEETS.SLOTS).filter(function(r){ return String(r[1]).toUpperCase() !== 'FALSE'; }).map(function(r){ return formatTime_(r[0]); }).filter(Boolean);
}
function getClosedDates_() {
  return getRows_(SHEETS.CLOSED).map(function(r){ return { date: formatDate_(r[0]), reason: String(r[1]||'') }; }).filter(function(x){ return x.date; });
}
function getRoomDates_() {
  return getRows_(SHEETS.ROOM_DATES).map(function(r){ return { date: formatDate_(r[0]), room:String(r[1]||'').trim(), status:String(r[2]||'closed').trim().toLowerCase(), reason:String(r[3]||'') }; }).filter(function(x){ return x.date && x.room; });
}

function getAvailability_(date, room) {
  var cfg = getConfig_();
  var rooms = getRooms_().filter(function(r){ return r.active; });
  var slots = getSlots_();
  if (!date || !room) return { result:'success', settings: publicSettings_(cfg), rooms: rooms, slots: slots };
  if (!cfg.systemActive) return { result:'success', isClosed:true, message:'Randevu sistemi şu anda kapalıdır.', rooms:rooms, slots:slots, availableSlots:[], taken:[] };
  var roomExists = rooms.some(function(r){ return r.code === room; });
  if (!roomExists) return { result:'error', message:'Seçilen görüşme birimi aktif değil.' };
  var closed = closureReason_(date, room, cfg);
  var taken = getTaken_(date, room);
  var available = closed ? [] : slots.filter(function(t){ return taken.indexOf(room + '-' + t) === -1 && passesLeadTime_(date, t, cfg.minLeadHours); });
  return { result:'success', settings: publicSettings_(cfg), rooms: rooms, slots: slots, taken: taken, availableSlots: available, isClosed: !!closed, message: closed || '' };
}
function publicSettings_(cfg) { return { systemActive: cfg.systemActive, sundayClosed: cfg.sundayClosed, minLeadHours: cfg.minLeadHours }; }

function closureReason_(date, room, cfg) {
  var rd = getRoomDates_().filter(function(x){ return x.date === date && x.room === room; });
  var openOverride = rd.some(function(x){ return x.status === 'open'; });
  if (openOverride) return '';
  if (cfg.sundayClosed && isSunday_(date)) return 'Pazar günleri randevu sistemimiz kapalıdır.';
  var cd = getClosedDates_().filter(function(x){ return x.date === date; });
  if (cd.length) return cd[0].reason || 'Bu tarih randevuya kapalıdır.';
  var closedRoom = rd.filter(function(x){ return x.status === 'closed'; });
  if (closedRoom.length) return closedRoom[0].reason || 'Bu görüşme birimi seçilen tarihte kapalıdır.';
  return '';
}

function getTaken_(date, room) {
  var rows = getRows_(SHEETS.BOOKINGS), out = [];
  rows.forEach(function(r){
    var status = String(r[9] || 'Aktif').toLowerCase();
    if (status === 'iptal' || status === 'cancelled') return;
    var rowRoom = String(r[6]||'').trim();
    var rowDate = formatDate_(r[7]);
    var rowTime = formatTime_(r[8]);
    if (rowDate === date && rowRoom === room && rowTime) out.push(rowRoom + '-' + rowTime);
  });
  return out;
}

function bookAppointment_(data) {
  var cfg = getConfig_();
  if (!cfg.systemActive) return { result:'error', message:'Randevu sistemi şu anda kapalıdır. Lütfen okul ile iletişime geçiniz.' };
  var required = ['veliAdiSoyadi','ogrenciAdiSoyadi','ogrenciSinifi','telefon','veliEposta','secilenOda','randevuTarihi','randevuSaati'];
  required.forEach(function(k){ if (!data[k]) throw new Error('Eksik alan: ' + k); });
  data.telefon = String(data.telefon).replace(/[^0-9+]/g,'');
  data.veliEposta = String(data.veliEposta || '').trim().toLowerCase();
  if (data.telefon.length < 10) return { result:'error', message:'Lütfen geçerli bir telefon numarası yazınız.' };
  if (!isValidEmail_(data.veliEposta)) return { result:'error', message:'Lütfen geçerli bir e-posta adresi yazınız.' };
  var activeRooms = getRooms_().filter(function(r){ return r.active; }).map(function(r){ return r.code; });
  if (activeRooms.indexOf(String(data.secilenOda).trim()) === -1) return { result:'error', message:'Seçilen görüşme birimi aktif değil.' };
  var slots = getSlots_();
  if (slots.indexOf(formatTime_(data.randevuSaati)) === -1) return { result:'error', message:'Seçilen randevu saati aktif değil.' };
  data.randevuSaati = formatTime_(data.randevuSaati);
  data.randevuTarihi = formatDate_(data.randevuTarihi);
  var closed = closureReason_(data.randevuTarihi, data.secilenOda, cfg);
  if (closed) return { result:'error', message: closed };
  if (!passesLeadTime_(data.randevuTarihi, data.randevuSaati, cfg.minLeadHours)) return { result:'error', message:'Bu saat için randevu alınamaz. Lütfen daha ileri bir saat seçiniz.' };
  var taken = getTaken_(data.randevuTarihi, data.secilenOda);
  if (taken.indexOf(data.secilenOda + '-' + data.randevuSaati) !== -1) return { result:'error', message:'Bu randevu saniyeler önce başka bir veli tarafından alındı. Lütfen farklı bir saat seçiniz.' };

  data.gorusmeBirimiAdi = getRoomName_(data.secilenOda);
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.BOOKINGS);
  sh.appendRow([new Date(), data.veliAdiSoyadi, data.ogrenciAdiSoyadi, data.ogrenciSinifi, data.telefon, data.veliEposta, data.secilenOda, data.randevuTarihi, data.randevuSaati, 'Aktif', data.kaynak || '', data.kampanya || '', data.reklamSeti || '', data.reklam || '', '']);
  sendAdminNotification_(data, cfg.adminEmails);
  if (cfg.parentMailEnabled) sendParentTicketEmail_(data, cfg);
  return { result:'success', message: cfg.successMessage };
}

function sendAdminNotification_(data, emails) {
  var to = (emails && emails.length ? emails : DEFAULT_EMAILS).join(',');
  var subject = 'Yeni Okul Tanıtım Randevusu: ' + data.ogrenciAdiSoyadi;
  var body = [
    'Yeni okul tanıtım randevusu alındı.', '',
    'Veli: ' + data.veliAdiSoyadi,
    'Öğrenci: ' + data.ogrenciAdiSoyadi,
    'Kademe/Sınıf: ' + data.ogrenciSinifi,
    'Telefon: ' + data.telefon,
    'Veli E-posta: ' + data.veliEposta,
    'Görüşme Birimi: ' + data.gorusmeBirimiAdi + ' (' + data.secilenOda + ')',
    'Randevu: ' + displayDate_(data.randevuTarihi) + ' ' + data.randevuSaati, '',
    'Reklam Kaynağı: ' + (data.kaynak || '-'),
    'Kampanya: ' + (data.kampanya || '-'),
    'Reklam Seti: ' + (data.reklamSeti || '-'),
    'Reklam: ' + (data.reklam || '-')
  ].join('\n');
  MailApp.sendEmail(to, subject, body);
}

function sendParentTicketEmail_(data, cfg) {
  var subject = 'Randevu Biletiniz Hazır | İngiliz Kültür Koleji';
  var plain = [
    'Sayın ' + data.veliAdiSoyadi + ',', '',
    'Okul tanıtım randevunuz oluşturulmuştur.', '',
    'Randevu Tarihi: ' + displayDate_(data.randevuTarihi),
    'Randevu Saati: ' + data.randevuSaati,
    'Görüşme Birimi: ' + data.gorusmeBirimiAdi,
    'Öğrenci: ' + data.ogrenciAdiSoyadi,
    'Kademe/Sınıf: ' + data.ogrenciSinifi, '',
    'Sorularınız için: 444 9 507'
  ].join('\n');
  var html = buildTicketHtml_(data, cfg);
  MailApp.sendEmail({
    to: data.veliEposta,
    subject: subject,
    body: plain,
    htmlBody: html,
    name: 'İngiliz Kültür Koleji',
    replyTo: cfg.mailReplyTo || 'batikent@ingilizkultur.com.tr'
  });
}

function buildTicketHtml_(data, cfg) {
  var dateText = escapeHtml_(displayDate_(data.randevuTarihi));
  var timeText = escapeHtml_(data.randevuSaati);
  var roomText = escapeHtml_(data.gorusmeBirimiAdi || data.secilenOda);
  var parent = escapeHtml_(data.veliAdiSoyadi);
  var student = escapeHtml_(data.ogrenciAdiSoyadi);
  var grade = escapeHtml_(data.ogrenciSinifi);
  var ticketNo = escapeHtml_(makeTicketNo_(data));
  var note = escapeHtml_(cfg.ticketBrandNote || 'İngiliz Kültür Koleji Tanıtım Peronu');
  return '<div style="margin:0;padding:0;background:#f4f6ff;font-family:Arial,Helvetica,sans-serif;color:#1f2937;">' +
    '<div style="max-width:680px;margin:0 auto;padding:28px 14px;">' +
      '<div style="background:#ffffff;border-radius:22px;overflow:hidden;border:1px solid #e5e7eb;box-shadow:0 18px 45px rgba(22,32,111,.14);">' +
        '<div style="background:#16206f;padding:24px;text-align:center;color:#fff;">' +
          '<img src="https://i.ibb.co/p54c0cN/ikk-LOGO-PNG.png" alt="İngiliz Kültür Koleji" style="max-width:190px;background:#fff;border-radius:14px;padding:10px;margin-bottom:16px;">' +
          '<div style="font-size:13px;letter-spacing:2px;text-transform:uppercase;opacity:.85;">Okul Tanıtım Randevusu</div>' +
          '<h1 style="margin:8px 0 0;font-size:28px;line-height:1.1;">Randevu Biletiniz Hazır</h1>' +
        '</div>' +
        '<div style="padding:22px;">' +
          '<p style="font-size:16px;line-height:1.55;margin:0 0 18px;">Sayın <strong>' + parent + '</strong>, okul tanıtım randevunuz başarıyla oluşturulmuştur. Görüşmeye gelirken aşağıdaki randevu biletindeki bilgileri kontrol etmenizi rica ederiz.</p>' +
          '<div style="border:2px dashed #c7d2fe;border-radius:20px;overflow:hidden;background:#fff;">' +
            '<div style="display:flex;flex-wrap:wrap;">' +
              '<div style="flex:2;min-width:270px;padding:22px;background:linear-gradient(135deg,#ffffff,#f8faff);">' +
                '<div style="color:#d71920;font-size:12px;font-weight:bold;letter-spacing:2px;text-transform:uppercase;">' + note + '</div>' +
                '<div style="margin-top:14px;font-size:34px;font-weight:800;color:#16206f;line-height:1;">' + timeText + '</div>' +
                '<div style="margin-top:8px;font-size:18px;font-weight:bold;color:#111827;">' + dateText + '</div>' +
                '<div style="margin-top:18px;padding:14px;border-radius:14px;background:#eef2ff;color:#16206f;font-weight:bold;">Peron / Oda: ' + roomText + '</div>' +
              '</div>' +
              '<div style="flex:1;min-width:220px;padding:22px;background:#16206f;color:#fff;">' +
                '<div style="font-size:12px;opacity:.8;text-transform:uppercase;letter-spacing:1.5px;">Bilet No</div>' +
                '<div style="font-size:22px;font-weight:800;margin:5px 0 18px;">' + ticketNo + '</div>' +
                '<div style="font-size:12px;opacity:.8;text-transform:uppercase;letter-spacing:1.5px;">Öğrenci</div>' +
                '<div style="font-size:16px;font-weight:bold;margin:5px 0 14px;">' + student + '</div>' +
                '<div style="font-size:12px;opacity:.8;text-transform:uppercase;letter-spacing:1.5px;">Kademe</div>' +
                '<div style="font-size:16px;font-weight:bold;margin-top:5px;">' + grade + '</div>' +
              '</div>' +
            '</div>' +
          '</div>' +
          '<div style="margin-top:18px;background:#fff7ed;border:1px solid #fed7aa;border-radius:16px;padding:14px;color:#9a3412;font-size:14px;line-height:1.45;">Randevu saatinde okulumuza gelmeniz yeterlidir. Değişiklik yapmak isterseniz 444 9 507 numarasından bize ulaşabilirsiniz.</div>' +
          '<div style="margin-top:22px;font-size:13px;color:#6b7280;text-align:center;">İngiliz Kültür Koleji • 444 9 507</div>' +
        '</div>' +
      '</div>' +
    '</div>' +
  '</div>';
}

function getAdminData_(pin) {
  var cfg = getConfig_();
  if (String(pin || '') !== cfg.adminPin) return { result:'error', message:'Admin PIN hatalı.' };
  return { result:'success', settings: publicSettingsFull_(cfg), rooms: getRooms_(), slots: getSlots_(), closedDates: getClosedDates_(), roomDates: getRoomDates_() };
}
function saveAdmin_(data) {
  var cfg = getConfig_();
  if (String(data.pin || '') !== cfg.adminPin) return { result:'error', message:'Admin PIN hatalı.' };
  writeSettings_(data.settings || {}, cfg.adminPin);
  writeRooms_(data.rooms || []);
  writeSlots_(data.slots || []);
  writeClosed_(data.closedDates || []);
  writeRoomDates_(data.roomDates || []);
  return { result:'success' };
}
function publicSettingsFull_(cfg) { return { systemActive:cfg.systemActive, sundayClosed:cfg.sundayClosed, minLeadHours:cfg.minLeadHours, adminEmails:cfg.adminEmails, successMessage:cfg.successMessage }; }
function writeSettings_(s, pin) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
  var rows = [
    ['admin_pin', pin, '/admin giriş PIN. Değiştirmek için burada düzenleyin.'],
    ['system_active', s.systemActive ? 'TRUE' : 'FALSE', 'TRUE/FALSE'],
    ['sunday_closed', s.sundayClosed ? 'TRUE' : 'FALSE', 'Pazar günleri kapalı mı?'],
    ['min_lead_hours', Number(s.minLeadHours || 0), 'Randevuya minimum kaç saat kala alınabilir?'],
    ['admin_emails', (s.adminEmails || DEFAULT_EMAILS).join(','), 'Bildirim e-postaları'],
    ['success_message', s.successMessage || 'Randevu talebiniz alınmıştır. Randevu biletiniz e-posta adresinize gönderilmiştir.', 'Veliye gösterilecek başarı mesajı'],
    ['school_phone','444 9 507','İletişim telefonu'],
    ['parent_mail_enabled','TRUE','Veliye otomatik randevu bileti maili gönderilsin mi? TRUE/FALSE'],
    ['ticket_brand_note','İngiliz Kültür Koleji Tanıtım Peronu','Veliye giden bilet mailinde küçük marka notu'],
    ['mail_reply_to','batikent@ingilizkultur.com.tr','Veli mailinde cevap adresi olarak gösterilecek adres']
  ];
  resetSheet_(sh, ['Anahtar','Deger','Aciklama'], rows);
}
function writeRooms_(rooms) { resetSheet_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ROOMS), ['Kod','GorunenAd','Aktif'], rooms.map(function(r){ return [r.code,r.name,!!r.active]; })); }
function writeSlots_(slots) { resetSheet_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SLOTS), ['Saat','Aktif'], slots.map(function(t){ return [t,true]; })); }
function writeClosed_(items) { resetSheet_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CLOSED), ['Tarih','Aciklama'], items.map(function(x){ return [x.date,x.reason||'']; })); }
function writeRoomDates_(items) { resetSheet_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ROOM_DATES), ['Tarih','OdaKodu','Durum','Aciklama'], items.map(function(x){ return [x.date,x.room,x.status||'closed',x.reason||'']; })); }
function resetSheet_(sh, headers, rows) { sh.clearContents(); sh.getRange(1,1,1,headers.length).setValues([headers]); if (rows.length) sh.getRange(2,1,rows.length,headers.length).setValues(rows); }

function getRows_(sheetName) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}
function formatDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  var s = String(v).trim();
  // Native HTML date value: yyyy-MM-dd or yyyy-MM-ddTHH:mm:ss
  var mIso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (mIso) return mIso[1] + '-' + mIso[2] + '-' + mIso[3];
  // Google Sheets Turkey display value can become dd.MM.yyyy or dd/MM/yyyy
  var mTr = s.match(/^(\d{1,2})[\.\/](\d{1,2})[\.\/](\d{4})$/);
  if (mTr) return mTr[3] + '-' + ('0' + mTr[2]).slice(-2) + '-' + ('0' + mTr[1]).slice(-2);
  // Last fallback: try Date parsing safely
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  return s.substring(0,10);
}
function formatTime_(v) { if (!v) return ''; if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, TZ, 'HH:mm'); var s=String(v).trim(); return s.length >= 5 ? s.substring(0,5) : s; }
function displayDate_(yyyyMmDd) { var d = new Date(yyyyMmDd + 'T12:00:00'); var days=['Pazar','Pazartesi','Salı','Çarşamba','Perşembe','Cuma','Cumartesi']; return Utilities.formatDate(d, TZ, 'dd.MM.yyyy') + ' ' + days[d.getDay()]; }
function isSunday_(dateStr) { return new Date(dateStr + 'T12:00:00').getDay() === 0; }
function passesLeadTime_(dateStr, timeStr, hours) { var target = new Date(dateStr + 'T' + timeStr + ':00'); var min = new Date(new Date().getTime() + Number(hours || 0) * 3600000); return target.getTime() >= min.getTime(); }
function isValidEmail_(email) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '')); }
function makeTicketNo_(data) { return 'IKK-' + String(data.randevuTarihi || '').replace(/-/g,'') + '-' + String(data.randevuSaati || '').replace(':','') + '-' + String(data.secilenOda || '').toUpperCase().substring(0,3); }
function escapeHtml_(s) { return String(s || '').replace(/[&<>"']/g, function(m){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]; }); }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function output_(obj, callback) {
  var text = JSON.stringify(obj);
  if (callback) {
    var safeCallback = String(callback).replace(/[^a-zA-Z0-9_.$]/g, '');
    return ContentService.createTextOutput(safeCallback + '(' + text + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(obj);
}
