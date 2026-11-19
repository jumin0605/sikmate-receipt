/* 같은 브라우저에서 로그인한 사람이 바뀌면, 앞사람이 남긴 이 브라우저 안의 임시 기록
   (식단표 사본, 사진 목록, 화면 설정 등)을 지운다.
   각자의 기록은 서버에 따로 저장되어 있으므로 지워도 사라지지 않는다.
   모든 화면의 <head> 맨 앞에서 먼저 실행되어야 한다. */
(function () {
  function cookie(name) {
    var m = document.cookie.match(new RegExp('(?:^|; )' + name + '=([^;]*)'));
    try { return m ? decodeURIComponent(m[1]) : ''; } catch (e) { return ''; }
  }
  var uid = cookie('sm_uid');
  if (!uid) return;
  var prev = '';
  try { prev = localStorage.getItem('smUid') || ''; } catch (e) { return; }
  if (prev === uid) return;
  if (prev) {                       // 다른 사람으로 바뀜: 앞사람의 브라우저 기록을 지운다
    try { localStorage.clear(); } catch (e) {}
    try {
      if (window.indexedDB && indexedDB.databases) {
        indexedDB.databases().then(function (list) {
          (list || []).forEach(function (d) { if (d && d.name) indexedDB.deleteDatabase(d.name); });
        });
      } else if (window.indexedDB) {
        indexedDB.deleteDatabase('sikmate-shared');
      }
    } catch (e) {}
  }
  try { localStorage.setItem('smUid', uid); } catch (e) {}
})();
