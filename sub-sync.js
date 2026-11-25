/* ════════════════════════════════════════════════════════════════
   學科紀錄雲端同步（學習大冒險 math-map.html 和 數學測驗 subjects.html 共用）

   走 LINE 點數機器人的「同步」工作表（跟英文口說 App 同一張表、同一個網址和密碼），
   用學科自己的分類名稱，英文 App 看到不認得的分類會略過、不會刪：
     subMap   id = 名字|關卡uid   每關的星星、最高紀錄、AI 講評、考試歷史
     subChest id = 名字|關卡uid   綠色寶箱打開的日期
     subWrong id = 名字|單元id    錯題本（下次優先出的題目）
     subStats id = 名字           學習狀況統計（各主題最近分數、常犯的錯、每天題數）
     subSet   id = starRules / rules / grades   家長設定（星星獎勵、及格標準、年級）

   什麼時候同步：打開頁面先抓雲端的、再傳這台改過的；過關、拿到星星／點數、錯題本、
   統計、寶箱、家長設定有變動時，3 秒後自動同步；網路斷掉就先留在這台，連上再補。

   兩台都改過同一筆時：
     subMap   合併：最高星數、已給獎勵取大的；寶箱有開就算開；歷史合在一起；講評用比較新的
     subStats 合併：每個主題取練習題數多的那份；每天題數取多的；錯誤次數取大的
     其他     這台還沒傳上去的改動優先，不然用雲端的
   ════════════════════════════════════════════════════════════════ */
(function(){
  "use strict";
  var store = {
    get: function(k, d){ try{ var v = localStorage.getItem(k); return v === null ? d : JSON.parse(v); }catch(e){ return d; } },
    set: function(k, v){ try{ localStorage.setItem(k, JSON.stringify(v)); return true; }catch(e){ return false; } }
  };
  var KINDS = ["subMap", "subChest", "subWrong", "subStats", "subSet"];
  var SET_KEYS = {starRules: "subStarRules", rules: "subRules", grades: "subGrades"};
  function hsh(s){ var h = 5381; for (var i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0; return String(h >>> 0) + ":" + s.length; }
  // 📱 新裝置一鍵設定：跟英文 App 同一種設定碼（網址後面帶 #setup=…），學習大冒險和學科測驗也認得
  (function(){
    var m = location.hash.match(/^#setup=(.+)$/); if (!m) return;
    try{
      var o = JSON.parse(decodeURIComponent(escape(atob(decodeURIComponent(m[1])))));
      if (o && o.app === "kousuo" && confirm("要用這組設定嗎？\n（Gemini 金鑰、機器人網址和連線密碼會存到這台）")){
        ["gkey", "gkey2", "azKey", "azRegion", "ptCloudUrl", "ptCloudKey"].forEach(function(k){ if (o[k]) store.set(k, o[k]); });
        if (o.ptPin) store.set("ptPin", o.ptPin);
      }
    }catch(e){ alert("設定碼讀不出來：" + e.message); }
    try{ history.replaceState(null, "", location.pathname + location.search); }catch(_){ location.hash = ""; }
  })();
  function cloud(){ return {url: store.get("ptCloudUrl", ""), key: store.get("ptCloudKey", "")}; }
  function on(){ var c = cloud(); return !!(c.url && c.key); }

  // ── 把這台的資料整理成一筆一筆 ──
  function collect(){
    var out = {};
    function put(k, id, d){ if (d != null) out[k + "|" + id] = {k: k, id: id, d: d}; }
    var m = store.get("subMap", {});
    Object.keys(m).forEach(function(nm){ Object.keys(m[nm] || {}).forEach(function(uid){ var r = m[nm][uid]; if (r && r.hist) put("subMap", nm + "|" + uid, r); }); });
    var c = store.get("subMapChest", {});
    Object.keys(c).forEach(function(nm){ Object.keys(c[nm] || {}).forEach(function(uid){ put("subChest", nm + "|" + uid, c[nm][uid]); }); });
    var w = store.get("subWrong", {});
    Object.keys(w).forEach(function(key){ put("subWrong", key, w[key]); });
    var s = store.get("mathStats", {});
    Object.keys(s).forEach(function(nm){ put("subStats", nm, s[nm]); });
    Object.keys(SET_KEYS).forEach(function(id){ var v = store.get(SET_KEYS[id], null); if (v != null) put("subSet", id, v); });
    return out;
  }
  // ── 把一筆寫回這台 ──
  function apply(k, id, d){
    var p = id.split("|");
    if (k === "subMap"){ var m = store.get("subMap", {}); (m[p[0]] = m[p[0]] || {})[p.slice(1).join("|")] = d; store.set("subMap", m); }
    else if (k === "subChest"){ var c = store.get("subMapChest", {}); (c[p[0]] = c[p[0]] || {})[p.slice(1).join("|")] = d; store.set("subMapChest", c); }
    else if (k === "subWrong"){ var w = store.get("subWrong", {}); w[id] = d; store.set("subWrong", w); }
    else if (k === "subStats"){ var s = store.get("mathStats", {}); s[id] = d; store.set("mathStats", s); }
    else if (k === "subSet" && SET_KEYS[id]) store.set(SET_KEYS[id], d);
  }
  // ── 兩邊都改過時怎麼合併 ──
  function mergeMap(a, b){
    if (!a || !a.hist) return b; if (!b || !b.hist) return a;
    var seen = {}, hist = a.hist.concat(b.hist).filter(function(x){ var s = JSON.stringify(x); if (seen[s]) return false; seen[s] = 1; return true; });
    hist.sort(function(x, y){ return (x.t || 0) - (y.t || 0) || String(x.d).localeCompare(String(y.d)); });
    var la = a.last || {}, lb = b.last || {}, last = (lb.t || 0) > (la.t || 0) || (!(la.t || lb.t) && String(lb.d) > String(la.d)) ? lb : la;
    var best = Math.max(a.best || 0, b.best || 0);
    return {best: best, paid: Math.max(a.paid || 0, b.paid || 0), pass: best >= 4, chest: !!(a.chest || b.chest), hist: hist.slice(-20), last: last};
  }
  function mergeStats(a, b){
    if (!a) return b; if (!b) return a;
    var o = {t: {}, err: {}, day: {}};
    [a, b].forEach(function(x){
      Object.keys(x.t || {}).forEach(function(k){ var v = x.t[k]; if (!o.t[k] || (v.n || 0) > (o.t[k].n || 0)) o.t[k] = v; });
      Object.keys(x.err || {}).forEach(function(k){ o.err[k] = Math.max(o.err[k] || 0, x.err[k] || 0); });
      Object.keys(x.day || {}).forEach(function(k){ var v = x.day[k]; if (!o.day[k] || (v.n || 0) > (o.day[k].n || 0)) o.day[k] = v; });
    });
    return o;
  }
  function merge(k, local, remote, localDirty){
    if (k === "subMap") return mergeMap(local, remote);
    if (k === "subStats") return mergeStats(local, remote);
    if (k === "subChest") return local || remote;
    return localDirty ? local : remote;
  }

  var hash = store.get("subSyncHash", {}), since = store.get("subSyncSince", 0), busy = false, again = false, timer = null;
  var state = {last: store.get("subSyncLast", 0), err: "", busy: false};
  function note(){ try{ window.dispatchEvent(new CustomEvent("subsync-status", {detail: state})); }catch(e){} }

  function run(){
    if (!on()){ state.err = "off"; note(); return Promise.resolve(false); }
    if (busy){ again = true; return Promise.resolve(false); }
    busy = true; state.busy = true; clearTimeout(timer); note();
    var c = cloud(), changed = 0;
    return fetch(c.url + (c.url.indexOf("?") >= 0 ? "&" : "?") + "key=" + encodeURIComponent(c.key) + "&sync=1&since=" + since + "&_=" + Date.now())
      .then(function(r){ return r.json(); })
      .then(function(j){
        if (!j || !j.ok) throw new Error((j && j.error) || "同步失敗");
        // ① 收雲端的變動（只看學科的分類；順便記下雲端的小朋友名單，給沒開過英文 App 的這台用）
        var cur = collect(), readers = store.get("subReadersCloud", {}), eng = store.get("subEngCloud", {lesson: {}, vocab: {}});
        (j.items || []).forEach(function(it){
          if (it.k === "reader" && it.d && it.d.name){ if (it.del) delete readers[it.id]; else readers[it.id] = it.d; return; }
          if (it.k === "lesson" || it.k === "vocab"){   // 英文口說練習的課文、單字卡：只留地圖要用的欄位
            var box = eng[it.k] = eng[it.k] || {};
            if (it.del || !it.d){ delete box[it.id]; return; }
            var d = it.d, txt = function(x){ return String(typeof x === "string" ? x : (x && (x.en || x.w || x.word)) || "").slice(0, 80); };
            box[it.id] = {id: d.id, title: d.title, at: d.at, owner: d.owner, scores: d.scores || {}, items: it.k === "lesson" ? (d.items || []).map(function(x){ return {en: txt(x)}; }) : undefined, words: it.k === "vocab" ? (d.words || []).map(function(x){ return {en: txt(x)}; }) : undefined};
            return;
          }
          if (KINDS.indexOf(it.k) < 0 || it.del || it.d == null) return;
          var key = it.k + "|" + it.id, loc = cur[key], rh = hsh(JSON.stringify(it.d));
          if (!loc){ apply(it.k, it.id, it.d); changed++; }
          else {
            var lh = hsh(JSON.stringify(loc.d));
            if (lh !== rh){
              var m = merge(it.k, loc.d, it.d, lh !== hash[key]);
              if (hsh(JSON.stringify(m)) !== lh){ apply(it.k, it.id, m); changed++; }
            }
          }
          hash[key] = rh;
        });
        store.set("subReadersCloud", readers); store.set("subEngCloud", eng);
        since = Math.max(0, (+j.now || Date.now()) - 60000);
        // ② 傳這台改過的
        var now = Date.now(), push = [];
        cur = collect();
        Object.keys(cur).forEach(function(key){ var it = cur[key], h = hsh(JSON.stringify(it.d)); if (h !== hash[key]) push.push({k: it.k, id: it.id, u: now, d: it.d, h: h}); });
        var p = Promise.resolve();
        for (var i = 0; i < push.length; i += 40){
          (function(batch){
            p = p.then(function(){
              return fetch(c.url, {method: "POST", body: JSON.stringify({key: c.key, action: "sync", items: batch.map(function(x){ return {k: x.k, id: x.id, u: x.u, d: x.d}; })})})
                .then(function(r){ return r.json(); })
                .then(function(j2){ if (!j2 || !j2.ok) throw new Error((j2 && j2.error) || "上傳失敗"); batch.forEach(function(x){ hash[x.k + "|" + x.id] = x.h; }); });
            });
          })(push.slice(i, i + 40));
        }
        return p.then(function(){ return push.length; });
      })
      .then(function(pushed){
        store.set("subSyncHash", hash); store.set("subSyncSince", since);
        state.last = Date.now(); state.err = ""; state.got = changed; state.sent = pushed; store.set("subSyncLast", state.last);
        if (changed) try{ window.dispatchEvent(new CustomEvent("subsync", {detail: {changed: changed}})); }catch(e){}
        return true;
      })
      .catch(function(e){ state.err = e.message || "同步失敗"; store.set("subSyncHash", hash); return false; })
      .then(function(ok){ busy = false; state.busy = false; note(); if (again){ again = false; timer = setTimeout(run, 1000); } return ok; });
  }
  function soon(ms){ if (!on()) return; clearTimeout(timer); timer = setTimeout(run, ms == null ? 3000 : ms); }
  addEventListener("online", function(){ soon(500); });
  // 沒開過英文 App 的這台：用雲端同步下來的名單
  function readers(){
    var loc = store.get("readers", []);
    if (loc && loc.length) return loc;
    var c = store.get("subReadersCloud", {});
    return Object.keys(c).map(function(id){ return c[id]; }).filter(function(r){ return r && r.id && r.name; });
  }
  window.SubSync = {run: run, soon: soon, on: on, state: state, readers: readers, _merge: merge};
})();
