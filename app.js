/* =========================================================================
   Kalkulator HPP Kaos — app.js
   Alur:
     TAHAP 1  HPP Bahan  = Biaya Kain/kaos + Biaya RIB/kaos
                Biaya bahan = (harga per Kg) x (gram per kaos / 1000)
     TAHAP 2  HPP Akhir/pcs = HPP Bahan + Ongkos Jahit + Ongkos Sablon
                             + Biaya Lain + (Transport batch / Jumlah pcs)
                Harga Jual/pcs = HPP Akhir x (1 + Margin%)
   Data disimpan otomatis di localStorage.
   ========================================================================= */

(function () {
  "use strict";

  var STORAGE_KEY = "hpp-kaos-v1";

  /* ---------- util ---------- */
  function $(id) { return document.getElementById(id); }

  function toNum(v) {
    var n = parseFloat(v);
    return isFinite(n) && n > 0 ? n : 0;
  }

  var rupiah = new Intl.NumberFormat("id-ID", {
    style: "currency", currency: "IDR", maximumFractionDigits: 0
  });
  function formatRp(n) {
    if (!isFinite(n)) n = 0;
    return rupiah.format(Math.round(n));
  }

  /* ---------- TAHAP 1 : HPP Bahan ---------- */
  function biayaPerKaos(hargaPerKg, gramPerKaos) {
    return hargaPerKg * (gramPerKaos / 1000);
  }

  function hitungHppBahan() {
    var biayaKain = biayaPerKaos(toNum($("hargaKain").value), toNum($("gramKain").value));
    var biayaRib = biayaPerKaos(toNum($("hargaRib").value), toNum($("gramRib").value));
    var hppBahan = biayaKain + biayaRib;

    $("outBiayaKain").textContent = formatRp(biayaKain);
    $("outBiayaRib").textContent = formatRp(biayaRib);
    $("outHppBahan").textContent = formatRp(hppBahan);
    return hppBahan;
  }

  /* ---------- TAHAP 2 : baris design ---------- */
  var tbody = $("designRows");

  function createRow(data) {
    data = data || {};
    var tr = document.createElement("tr");
    tr.innerHTML =
      '<td class="rownum"></td>' +
      '<td class="text"><input type="text"   class="f-nama"      placeholder="Nama design"></td>' +
      '<td><input type="number" class="f-jumlah"   min="0" step="1"   placeholder="0"></td>' +
      '<td><input type="number" class="f-jahit"    min="0" step="500" placeholder="0"></td>' +
      '<td><input type="number" class="f-sablon"   min="0" step="500" placeholder="0"></td>' +
      '<td><input type="number" class="f-lain"     min="0" step="500" placeholder="0"></td>' +
      '<td><input type="number" class="f-transport" min="0" step="1000" placeholder="0"></td>' +
      '<td class="calc o-hppbahan">Rp0</td>' +
      '<td class="calc o-hppakhir">Rp0</td>' +
      '<td class="calc o-jual">Rp0</td>' +
      '<td class="calc o-modal">Rp0</td>' +
      '<td class="calc o-totaljual">Rp0</td>' +
      '<td class="calc o-laba">Rp0</td>' +
      '<td><button class="row-del" title="Hapus baris">&times;</button></td>';

    tr.querySelector(".f-nama").value = data.nama || "";
    tr.querySelector(".f-jumlah").value = data.jumlah || "";
    tr.querySelector(".f-jahit").value = data.jahit || "";
    tr.querySelector(".f-sablon").value = data.sablon || "";
    tr.querySelector(".f-lain").value = data.lain || "";
    tr.querySelector(".f-transport").value = data.transport || "";

    tr.addEventListener("input", recalcAll);
    tr.querySelector(".row-del").addEventListener("click", function () {
      tr.parentNode.removeChild(tr);
      recalcAll();
    });

    tbody.appendChild(tr);
    return tr;
  }

  function recalcRow(tr, hppBahan) {
    var jumlah = toNum(tr.querySelector(".f-jumlah").value);
    var jahit = toNum(tr.querySelector(".f-jahit").value);
    var sablon = toNum(tr.querySelector(".f-sablon").value);
    var lain = toNum(tr.querySelector(".f-lain").value);
    var transport = toNum(tr.querySelector(".f-transport").value);
    var margin = toNum($("margin").value) / 100;

    var transportPerPcs = jumlah > 0 ? transport / jumlah : 0;
    var hppAkhir = hppBahan + jahit + sablon + lain + transportPerPcs;
    var hargaJual = hppAkhir * (1 + margin);

    var totalModal = hppAkhir * jumlah;
    var totalJual = hargaJual * jumlah;
    var laba = totalJual - totalModal;

    tr.querySelector(".o-hppbahan").textContent = formatRp(hppBahan);
    tr.querySelector(".o-hppakhir").textContent = formatRp(hppAkhir);
    tr.querySelector(".o-jual").textContent = formatRp(hargaJual);
    tr.querySelector(".o-modal").textContent = formatRp(totalModal);
    tr.querySelector(".o-totaljual").textContent = formatRp(totalJual);
    tr.querySelector(".o-laba").textContent = formatRp(laba);

    return { jumlah: jumlah, totalModal: totalModal, totalJual: totalJual, laba: laba };
  }

  /* ---------- recalc keseluruhan ---------- */
  function recalcAll() {
    var hppBahan = hitungHppBahan();
    var rows = tbody.querySelectorAll("tr");
    var sumQty = 0, sumModal = 0, sumJual = 0, sumLaba = 0;

    rows.forEach(function (tr, i) {
      tr.querySelector(".rownum").textContent = i + 1;
      var r = recalcRow(tr, hppBahan);
      sumQty += r.jumlah;
      sumModal += r.totalModal;
      sumJual += r.totalJual;
      sumLaba += r.laba;
    });

    $("sumQty").textContent = sumQty.toLocaleString("id-ID");
    $("sumModal").textContent = formatRp(sumModal);
    $("sumJual").textContent = formatRp(sumJual);
    $("sumLaba").textContent = formatRp(sumLaba);

    $("statQty").textContent = sumQty.toLocaleString("id-ID") + " pcs";
    $("statModal").textContent = formatRp(sumModal);
    $("statJual").textContent = formatRp(sumJual);
    $("statLaba").textContent = formatRp(sumLaba);

    save();
  }

  /* ---------- simpan / muat (localStorage) ---------- */
  function collectState() {
    var rows = [];
    tbody.querySelectorAll("tr").forEach(function (tr) {
      rows.push({
        nama: tr.querySelector(".f-nama").value,
        jumlah: tr.querySelector(".f-jumlah").value,
        jahit: tr.querySelector(".f-jahit").value,
        sablon: tr.querySelector(".f-sablon").value,
        lain: tr.querySelector(".f-lain").value,
        transport: tr.querySelector(".f-transport").value
      });
    });
    return {
      hargaKain: $("hargaKain").value,
      gramKain: $("gramKain").value,
      hargaRib: $("hargaRib").value,
      gramRib: $("gramRib").value,
      margin: $("margin").value,
      rows: rows
    };
  }

  function save() {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(collectState())); } catch (e) {}
  }

  function load() {
    var raw;
    try { raw = localStorage.getItem(STORAGE_KEY); } catch (e) { raw = null; }
    if (!raw) return false;
    var s;
    try { s = JSON.parse(raw); } catch (e) { return false; }

    $("hargaKain").value = s.hargaKain || "";
    $("gramKain").value = s.gramKain || "";
    $("hargaRib").value = s.hargaRib || "";
    $("gramRib").value = s.gramRib || "";
    $("margin").value = (s.margin !== undefined && s.margin !== "") ? s.margin : 40;

    tbody.innerHTML = "";
    (s.rows && s.rows.length ? s.rows : [{}]).forEach(createRow);
    return true;
  }

  /* ---------- contoh data ---------- */
  function fillSample() {
    $("hargaKain").value = 95000;
    $("gramKain").value = 250;   // body kaos ~250 gram
    $("hargaRib").value = 110000;
    $("gramRib").value = 30;     // kerah/manset ~30 gram
    $("margin").value = 40;

    tbody.innerHTML = "";
    createRow({ nama: "Logo Depan 1 Warna", jumlah: 50, jahit: 8000, sablon: 7000, lain: 3000, transport: 25000 });
    createRow({ nama: "Full Print Belakang", jumlah: 30, jahit: 8000, sablon: 12000, lain: 3000, transport: 25000 });
    createRow({ nama: "Sablon 3 Warna", jumlah: 40, jahit: 8000, sablon: 15000, lain: 3500, transport: 30000 });
    recalcAll();
  }

  function resetAll() {
    if (!confirm("Hapus semua data dan mulai dari kosong?")) return;
    try { localStorage.removeItem(STORAGE_KEY); } catch (e) {}
    ["hargaKain", "gramKain", "hargaRib", "gramRib"].forEach(function (id) { $(id).value = ""; });
    $("margin").value = 40;
    tbody.innerHTML = "";
    createRow({});
    recalcAll();
  }

  /* ---------- init ---------- */
  function init() {
    $("btnAdd").addEventListener("click", function () { createRow({}); recalcAll(); });
    $("btnSample").addEventListener("click", fillSample);
    $("btnReset").addEventListener("click", resetAll);

    ["hargaKain", "gramKain", "hargaRib", "gramRib", "margin"].forEach(function (id) {
      $(id).addEventListener("input", recalcAll);
    });

    if (!load()) {
      createRow({});       // mulai dengan satu baris kosong
    }
    recalcAll();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
