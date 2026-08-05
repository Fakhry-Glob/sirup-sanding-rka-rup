// ==UserScript==
// @name         SiRUP RKA & RUP Exporter & Sander
// @namespace    http://tampermonkey.net/
// @version      2.0
// @description  Crawl RKA dan RUP dari SiRUP, lalu ekspor jadi laporan sanding Excel (dashboard, ringkasan, sanding berjenjang, detail per program). Tahun anggaran & satker terdeteksi otomatis.
// @author       Fakhry-Glob
// @match        https://sirup.inaproc.id/sirup/*
// @require      https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js
// @updateURL    https://raw.githubusercontent.com/Fakhry-Glob/sirup-sanding-rka-rup/master/sirup_exporter.user.js
// @downloadURL  https://raw.githubusercontent.com/Fakhry-Glob/sirup-sanding-rka-rup/master/sirup_exporter.user.js
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // ═════════════════════════════════════════════════════════ KONFIGURASI ══
    const APP_TITLE = 'Sanding RKA & RUP';
    const APP_VERSION = '2.0';

    // Konteks runtime: diisi otomatis oleh detectContext(), bisa dikoreksi
    // pengguna lewat panel pra-ekspor sebelum crawling dimulai.
    const ctx = {
        tahun: new Date().getFullYear(),
        satkerId: '',
        satkerName: ''
    };

    let abortRequested = false;
    function throwIfAborted() {
        if (abortRequested) throw new Error('Proses dibatalkan oleh pengguna.');
    }

    // ═══════════════════════════════════════════════════════════════ STYLE ══
    const CSS = `
    .srx, .srx *, .srx *::before, .srx *::after { box-sizing: border-box; }
    .srx { font-family: "Segoe UI", system-ui, -apple-system, sans-serif; color: #0f172a; }

    .srx-fab {
        position: fixed; right: 24px; bottom: 24px; z-index: 99990;
        display: inline-flex; align-items: center; gap: 9px;
        padding: 12px 20px; border: 0; border-radius: 999px; cursor: pointer;
        font: 600 14px/1 "Segoe UI", system-ui, sans-serif; color: #fff;
        background: linear-gradient(135deg, #1F497D 0%, #2E6DA4 100%);
        box-shadow: 0 6px 18px rgba(31,73,125,.35);
        transition: transform .15s ease, box-shadow .15s ease;
    }
    .srx-fab:hover { transform: translateY(-2px); box-shadow: 0 10px 26px rgba(31,73,125,.45); }
    .srx-fab:active { transform: translateY(0); }
    .srx-fab svg { width: 17px; height: 17px; flex: none; }

    .srx-overlay {
        position: fixed; inset: 0; z-index: 99991;
        background: rgba(15,23,42,.55); backdrop-filter: blur(2px);
        display: flex; align-items: center; justify-content: center; padding: 20px;
        animation: srx-fade .18s ease;
    }
    @keyframes srx-fade { from { opacity: 0 } to { opacity: 1 } }

    .srx-dialog {
        width: 640px; max-width: 100%; max-height: 90vh; overflow: hidden;
        display: flex; flex-direction: column;
        background: #fff; border-radius: 14px;
        box-shadow: 0 24px 60px rgba(2,6,23,.35);
        animation: srx-pop .2s cubic-bezier(.2,.8,.3,1);
    }
    @keyframes srx-pop { from { opacity: 0; transform: translateY(12px) scale(.98) } to { opacity: 1; transform: none } }

    .srx-head {
        padding: 18px 22px; color: #fff;
        background: linear-gradient(135deg, #1F497D 0%, #2E6DA4 100%);
        display: flex; align-items: flex-start; gap: 12px;
    }
    .srx-head h3 { margin: 0; font-size: 16px; font-weight: 700; letter-spacing: .2px; }
    .srx-head p  { margin: 3px 0 0; font-size: 12px; opacity: .82; }
    .srx-ver {
        margin-left: auto; flex: none; font-size: 11px; font-weight: 600;
        padding: 3px 9px; border-radius: 999px; background: rgba(255,255,255,.18);
    }

    .srx-body { padding: 18px 22px; overflow-y: auto; }
    .srx-foot {
        padding: 13px 22px; border-top: 1px solid #e2e8f0; background: #f8fafc;
        display: flex; align-items: center; gap: 10px;
    }
    .srx-foot .srx-spacer { margin-left: auto; }

    .srx-btn {
        padding: 9px 18px; border-radius: 8px; border: 1px solid transparent;
        font: 600 13px/1.2 "Segoe UI", system-ui, sans-serif; cursor: pointer;
        transition: background .15s ease, border-color .15s ease, opacity .15s ease;
    }
    .srx-btn:disabled { opacity: .5; cursor: not-allowed; }
    .srx-btn-primary { background: #1F497D; color: #fff; }
    .srx-btn-primary:hover:not(:disabled) { background: #17395f; }
    .srx-btn-ghost { background: #fff; color: #475569; border-color: #cbd5e1; }
    .srx-btn-ghost:hover:not(:disabled) { background: #f1f5f9; }
    .srx-btn-danger { background: #fff; color: #b91c1c; border-color: #fca5a5; }
    .srx-btn-danger:hover:not(:disabled) { background: #fef2f2; }

    .srx-field { margin-bottom: 15px; }
    .srx-field:last-child { margin-bottom: 0; }
    .srx-label { display: block; font-size: 12px; font-weight: 600; color: #334155; margin-bottom: 5px; }
    .srx-hint  { font-size: 11px; color: #64748b; margin-top: 5px; line-height: 1.45; }
    .srx-input, .srx-select {
        width: 100%; padding: 9px 11px; border: 1px solid #cbd5e1; border-radius: 8px;
        font: 400 13px/1.3 "Segoe UI", system-ui, sans-serif; color: #0f172a; background: #fff;
    }
    .srx-input:focus, .srx-select:focus { outline: 2px solid #93c5fd; outline-offset: -1px; border-color: #2E6DA4; }
    .srx-grid { display: grid; grid-template-columns: 160px 1fr; gap: 14px; }

    .srx-note {
        display: flex; gap: 9px; padding: 10px 12px; border-radius: 8px;
        background: #eff6ff; border: 1px solid #bfdbfe; color: #1e40af;
        font-size: 12px; line-height: 1.5;
    }

    .srx-steps { display: flex; margin: 0 0 16px; padding: 0; list-style: none; }
    .srx-steps li {
        flex: 1; position: relative; text-align: center;
        font-size: 11px; font-weight: 600; color: #94a3b8; padding-top: 24px;
    }
    .srx-steps li::before {
        content: ''; position: absolute; top: 5px; left: 50%; transform: translateX(-50%);
        width: 12px; height: 12px; border-radius: 50%;
        background: #fff; border: 2px solid #cbd5e1; z-index: 1;
    }
    .srx-steps li::after {
        content: ''; position: absolute; top: 10px; left: 50%; width: 100%; height: 2px; background: #e2e8f0;
    }
    .srx-steps li:last-child::after { display: none; }
    .srx-steps li.done { color: #1F497D; }
    .srx-steps li.done::before { background: #1F497D; border-color: #1F497D; }
    .srx-steps li.done::after  { background: #1F497D; }
    .srx-steps li.active { color: #1F497D; }
    .srx-steps li.active::before {
        background: #2E6DA4; border-color: #2E6DA4;
        box-shadow: 0 0 0 4px rgba(46,109,164,.2); animation: srx-pulse 1.4s ease-in-out infinite;
    }
    @keyframes srx-pulse { 50% { box-shadow: 0 0 0 7px rgba(46,109,164,.08) } }

    .srx-progress-row { display: flex; align-items: center; gap: 12px; margin-bottom: 12px; }
    .srx-track { flex: 1; height: 9px; border-radius: 999px; background: #e2e8f0; overflow: hidden; }
    .srx-bar {
        height: 100%; width: 0%; border-radius: 999px;
        background: linear-gradient(90deg, #1F497D, #2E6DA4);
        transition: width .35s ease;
    }
    .srx-bar.running {
        background-image: linear-gradient(90deg, #1F497D, #2E6DA4),
            repeating-linear-gradient(45deg, rgba(255,255,255,.18) 0 8px, transparent 8px 16px);
        background-blend-mode: overlay;
        animation: srx-stripe 1s linear infinite;
    }
    @keyframes srx-stripe { to { background-position: 32px 0, 0 0 } }
    .srx-bar.ok   { background: linear-gradient(90deg, #15803d, #22c55e); }
    .srx-bar.fail { background: linear-gradient(90deg, #b91c1c, #ef4444); }
    .srx-pct   { font: 700 13px/1 "Segoe UI", monospace; color: #1F497D; min-width: 42px; text-align: right; }
    .srx-timer { font: 400 12px/1 "Segoe UI", monospace; color: #64748b; min-width: 46px; text-align: right; }

    .srx-log {
        height: 224px; overflow-y: auto; padding: 6px;
        border: 1px solid #e2e8f0; border-radius: 8px; background: #f8fafc;
    }
    .srx-log p {
        margin: 0; padding: 4px 9px; font-size: 12px; line-height: 1.45; color: #334155;
        border-left: 3px solid transparent; border-radius: 3px; word-break: break-word;
    }
    .srx-log p.warn    { color: #92400e; background: #fffbeb; border-left-color: #f59e0b; }
    .srx-log p.error   { color: #991b1b; background: #fef2f2; border-left-color: #ef4444; font-weight: 600; }
    .srx-log p.success { color: #166534; background: #f0fdf4; border-left-color: #22c55e; font-weight: 600; }
    .srx-log p.step    { color: #1e3a8a; background: #eff6ff; border-left-color: #2E6DA4; font-weight: 600; margin-top: 3px; }

    .srx-summary { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin-bottom: 14px; }
    .srx-card { padding: 11px 13px; border: 1px solid #e2e8f0; border-radius: 9px; background: #f8fafc; }
    .srx-card span { display: block; font-size: 10.5px; font-weight: 600; color: #64748b; text-transform: uppercase; letter-spacing: .4px; }
    .srx-card strong { display: block; margin-top: 4px; font-size: 17px; font-weight: 700; color: #1F497D; }
    `;

    function injectStyles() {
        if (document.getElementById('srx-style')) return;
        const el = document.createElement('style');
        el.id = 'srx-style';
        el.textContent = CSS;
        (document.head || document.documentElement).appendChild(el);
    }

    // ══════════════════════════════════════════════════════ DETEKSI KONTEKS ══
    function findSatkerSelect() {
        return document.querySelector('#idSatker, select[name="idSatker"], #satker, #id_satker');
    }

    function findTahunSelect() {
        return document.querySelector('#tahunAnggaran, select[name="tahunAnggaran"], #tahun, select[name="tahun"]');
    }

    function detectTahun() {
        const sel = findTahunSelect();
        if (sel && /^\d{4}$/.test(String(sel.value || '').trim())) return parseInt(sel.value, 10);

        const fromUrl = new URLSearchParams(location.search).get('tahun');
        if (fromUrl && /^\d{4}$/.test(fromUrl)) return parseInt(fromUrl, 10);

        const m = (document.body.innerText || '').match(/Tahun\s*Anggaran\s*[:\-]?\s*(20\d{2})/i);
        if (m) return parseInt(m[1], 10);

        return new Date().getFullYear();
    }

    function detectSatker() {
        let id = '', name = '';

        const sel = findSatkerSelect();
        if (sel) {
            id = String(sel.value || '').trim();
            const opt = sel.options ? sel.options[sel.selectedIndex] : null;
            if (opt && opt.text) name = opt.text.trim();
        }
        if (!id) {
            const p = new URLSearchParams(location.search);
            id = p.get('idSatker') || p.get('satker') || p.get('id_satker') || '';
        }
        if (!name) {
            const el = document.querySelector('#namaSatker, .satker-name, .nama-satker');
            if (el) name = el.innerText.trim();
        }
        if (!name) {
            const m = (document.body.innerText || '').match(/Satuan\s*Kerja\s*[:\-]\s*([^\n\r]{4,120})/i);
            if (m) name = m[1].trim();
        }
        return { id, name };
    }

    function detectContext() {
        ctx.tahun = detectTahun();
        const s = detectSatker();
        ctx.satkerId = s.id;
        ctx.satkerName = s.name;
    }

    function tahunOptions() {
        const opts = new Set();
        const sel = findTahunSelect();
        if (sel) {
            Array.from(sel.options || []).forEach(o => {
                const v = String(o.value || '').trim();
                if (/^\d{4}$/.test(v)) opts.add(parseInt(v, 10));
            });
        }
        const now = new Date().getFullYear();
        [now - 1, now, now + 1].forEach(y => opts.add(y));
        opts.add(ctx.tahun);
        return Array.from(opts).sort((a, b) => b - a);
    }

    // ═══════════════════════════════════════════════════════════════ TOMBOL ══
    function injectButton() {
        if (document.getElementById('srx-fab')) return;
        injectStyles();

        const btn = document.createElement('button');
        btn.id = 'srx-fab';
        btn.type = 'button';
        btn.className = 'srx srx-fab';
        btn.title = 'Ekspor sanding RKA & RUP ke Excel';
        btn.innerHTML =
            '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" ' +
            'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
            '<path d="M3 3v18h18"/><rect x="7" y="11" width="3" height="7"/>' +
            '<rect x="12" y="7" width="3" height="11"/><rect x="17" y="13" width="3" height="5"/></svg>' +
            '<span>Ekspor Sanding RKA &amp; RUP</span>';

        btn.addEventListener('click', openPanel);
        document.body.appendChild(btn);
    }

    // ══════════════════════════════════════════════════ PANEL PRA-EKSPOR ══
    function closeOverlay() {
        const el = document.getElementById('srx-overlay');
        if (el) el.remove();
        document.removeEventListener('keydown', onEsc);
    }

    function onEsc(e) {
        if (e.key === 'Escape') {
            const dlg = document.getElementById('srx-overlay');
            if (dlg && dlg.dataset.closable === '1') closeOverlay();
        }
    }

    function buildOverlay(headTitle, headSub, bodyHtml, closable) {
        closeOverlay();
        const overlay = document.createElement('div');
        overlay.id = 'srx-overlay';
        overlay.className = 'srx srx-overlay';
        overlay.dataset.closable = closable ? '1' : '0';
        overlay.innerHTML =
            '<div class="srx-dialog" role="dialog" aria-modal="true">' +
                '<div class="srx-head">' +
                    '<div><h3>' + headTitle + '</h3><p>' + headSub + '</p></div>' +
                    '<span class="srx-ver">v' + APP_VERSION + '</span>' +
                '</div>' +
                '<div class="srx-body">' + bodyHtml + '</div>' +
                '<div class="srx-foot"></div>' +
            '</div>';
        document.body.appendChild(overlay);
        document.addEventListener('keydown', onEsc);
        return overlay;
    }

    function openPanel() {
        detectContext();
        injectStyles();

        const years = tahunOptions()
            .map(y => '<option value="' + y + '"' + (y === ctx.tahun ? ' selected' : '') + '>' + y + '</option>')
            .join('');

        const overlay = buildOverlay(
            'Ekspor ' + APP_TITLE,
            'Periksa dulu tahun anggaran & satker, lalu jalankan penarikan data.',
            '<div class="srx-field srx-grid">' +
                '<label class="srx-label" for="srx-tahun">Tahun Anggaran</label>' +
                '<div><select class="srx-select" id="srx-tahun">' + years + '</select>' +
                '<div class="srx-hint">Dipakai untuk menarik daftar paket RUP. Terdeteksi otomatis dari halaman ini.</div></div>' +
            '</div>' +
            '<div class="srx-field srx-grid">' +
                '<label class="srx-label" for="srx-satker">Nama Satker</label>' +
                '<div><input class="srx-input" id="srx-satker" type="text" placeholder="Nama satuan kerja" value="' +
                    String(ctx.satkerName || '').replace(/"/g, '&quot;') + '">' +
                '<div class="srx-hint">Muncul di kop setiap sheet laporan. ID Satker aktif: <b>' +
                    (ctx.satkerId || 'bawaan sesi login') + '</b></div></div>' +
            '</div>' +
            '<div class="srx-note">' +
                '<span>&#9432;</span>' +
                '<span>Penarikan data membaca seluruh Program &rarr; Kegiatan &rarr; KRO &rarr; RO &rarr; Komponen ' +
                'beserta detail tiap paket RUP. Untuk satker besar prosesnya bisa beberapa menit — biarkan tab ini terbuka.</span>' +
            '</div>',
            true
        );

        const foot = overlay.querySelector('.srx-foot');
        foot.innerHTML =
            '<span class="srx-spacer"></span>' +
            '<button type="button" class="srx-btn srx-btn-ghost" id="srx-cancel">Batal</button>' +
            '<button type="button" class="srx-btn srx-btn-primary" id="srx-start">Mulai Ekspor</button>';

        overlay.querySelector('#srx-cancel').addEventListener('click', closeOverlay);
        overlay.querySelector('#srx-start').addEventListener('click', () => {
            ctx.tahun = parseInt(overlay.querySelector('#srx-tahun').value, 10) || new Date().getFullYear();
            ctx.satkerName = overlay.querySelector('#srx-satker').value.trim();
            startExport();
        });
    }

    // ═══════════════════════════════════════════════════════ PANEL PROGRES ══
    const STEPS = ['Persiapan', 'Tarik RKA', 'Tarik RUP', 'Menyanding', 'Susun Excel'];
    let ui = null;
    let startedAt = 0;
    let timerId = null;

    function stepForProgress(p) {
        if (p == null) return null;
        if (p < 10) return 0;
        if (p < 50) return 1;
        if (p < 85) return 2;
        if (p < 90) return 3;
        return 4;
    }

    function classifyMessage(msg) {
        const m = String(msg);
        if (/^ERROR|gagal|error/i.test(m)) return 'error';
        if (/\[Swakelola|peringatan|tidak ditemukan|dibatalkan/i.test(m)) return 'warn';
        if (/berhasil|selesai!/i.test(m)) return 'success';
        if (/^\[RKA\]|^\[RUP\]|^Memproses|^Membangun|^Mengambil/i.test(m)) return 'step';
        return 'info';
    }

    function fmtElapsed(ms) {
        const s = Math.floor(ms / 1000);
        return String(Math.floor(s / 60)).padStart(2, '0') + ':' + String(s % 60).padStart(2, '0');
    }

    function showLogger() {
        abortRequested = false;
        injectStyles();

        const overlay = buildOverlay(
            'Memproses ' + APP_TITLE,
            'Tahun ' + ctx.tahun + (ctx.satkerName ? ' • ' + ctx.satkerName : ''),
            '<ul class="srx-steps">' + STEPS.map(s => '<li>' + s + '</li>').join('') + '</ul>' +
            '<div class="srx-progress-row">' +
                '<div class="srx-track"><div class="srx-bar running"></div></div>' +
                '<div class="srx-pct">0%</div><div class="srx-timer">00:00</div>' +
            '</div>' +
            '<div class="srx-log"></div>',
            false
        );

        const foot = overlay.querySelector('.srx-foot');
        foot.innerHTML =
            '<span class="srx-spacer"></span>' +
            '<button type="button" class="srx-btn srx-btn-danger" id="srx-abort">Hentikan</button>';
        foot.querySelector('#srx-abort').addEventListener('click', (e) => {
            abortRequested = true;
            e.target.disabled = true;
            log('Permintaan berhenti diterima, menunggu proses berjalan selesai...', null, 'warn');
        });

        ui = {
            overlay,
            bar: overlay.querySelector('.srx-bar'),
            pct: overlay.querySelector('.srx-pct'),
            timer: overlay.querySelector('.srx-timer'),
            logBox: overlay.querySelector('.srx-log'),
            steps: Array.from(overlay.querySelectorAll('.srx-steps li')),
            foot
        };

        startedAt = Date.now();
        clearInterval(timerId);
        timerId = setInterval(() => {
            if (ui) ui.timer.textContent = fmtElapsed(Date.now() - startedAt);
        }, 1000);

        setStep(0);
    }

    function setStep(idx) {
        if (!ui) return;
        ui.steps.forEach((li, i) => {
            li.classList.toggle('done', i < idx);
            li.classList.toggle('active', i === idx);
        });
    }

    function log(message, progress = null, type = null) {
        if (!ui) return;
        const p = document.createElement('p');
        p.className = type || classifyMessage(message);
        p.textContent = message;
        ui.logBox.appendChild(p);
        ui.logBox.scrollTop = ui.logBox.scrollHeight;

        if (progress !== null) {
            const pct = Math.max(0, Math.min(100, Math.round(progress)));
            ui.bar.style.width = pct + '%';
            ui.pct.textContent = pct + '%';
            const s = stepForProgress(progress);
            if (s !== null) setStep(s);
        }
    }

    function finishSuccess(cards) {
        if (!ui) return;
        clearInterval(timerId);
        ui.bar.classList.remove('running');
        ui.bar.classList.add('ok');
        ui.bar.style.width = '100%';
        ui.pct.textContent = '100%';
        ui.steps.forEach(li => { li.classList.remove('active'); li.classList.add('done'); });

        if (cards && cards.length) {
            const box = document.createElement('div');
            box.className = 'srx-summary';
            box.innerHTML = cards
                .map(c => '<div class="srx-card"><span>' + c.label + '</span><strong>' + c.value + '</strong></div>')
                .join('');
            ui.logBox.parentNode.insertBefore(box, ui.logBox);
        }

        ui.foot.innerHTML =
            '<span class="srx-spacer"></span>' +
            '<button type="button" class="srx-btn srx-btn-primary" id="srx-done">Selesai</button>';
        ui.foot.querySelector('#srx-done').addEventListener('click', hideLogger);
        ui.overlay.dataset.closable = '1';
    }

    function finishError(err) {
        if (!ui) return;
        clearInterval(timerId);
        ui.bar.classList.remove('running');
        ui.bar.classList.add('fail');
        ui.steps.forEach(li => li.classList.remove('active'));

        ui.foot.innerHTML =
            '<span class="srx-spacer"></span>' +
            '<button type="button" class="srx-btn srx-btn-ghost" id="srx-close">Tutup</button>' +
            '<button type="button" class="srx-btn srx-btn-primary" id="srx-retry">Coba Lagi</button>';
        ui.foot.querySelector('#srx-close').addEventListener('click', hideLogger);
        ui.foot.querySelector('#srx-retry').addEventListener('click', () => { hideLogger(); openPanel(); });
        ui.overlay.dataset.closable = '1';
    }

    function hideLogger() {
        clearInterval(timerId);
        closeOverlay();
        ui = null;
    }

    // Helper Functions
    async function postForm(url, body) {
        const res = await fetch(url, {
            method: "POST",
            headers: {
                "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                "X-Requested-With": "XMLHttpRequest"
            },
            body: body
        });
        return res.json();
    }

    async function getText(url) {
        const res = await fetch(url);
        return res.text();
    }

    // Helper for batched concurrent execution
    async function mapConcurrent(items, concurrency, fn) {
        const results = [];
        const copy = [...items];
        let index = 0;
        
        async function worker() {
            while (index < copy.length) {
                const currentIdx = index++;
                const item = copy[currentIdx];
                try {
                    results[currentIdx] = await fn(item, currentIdx, items.length);
                } catch (err) {
                    results[currentIdx] = null;
                    console.error(`Error processing item at index ${currentIdx}:`, err);
                }
            }
        }
        
        const workers = Array(Math.min(concurrency, items.length))
            .fill(null)
            .map(() => worker());
            
        await Promise.all(workers);
        return results;
    }

    // Helper to parse checkbox values from Datatables cell (boolean, string, or HTML)
    function parseCheckboxValue(val) {
        if (val === true || val === "true") return true;
        if (val === false || val === "false") return false;
        if (typeof val === 'string') {
            const lower = val.toLowerCase();
            if (lower.includes('checkbox') || lower.includes('input')) {
                return lower.includes('checked');
            }
            if (lower.includes('glyphicon-ok')) return true;
            if (lower.includes('glyphicon-remove')) return false;
        }
        return false;
    }

    // Robust MAK parser
    function parseMak(mak) {
        if (!mak) return null;
        mak = mak.toString().replace(/\s+/g, "").trim();
        const parts = mak.split(".");
        // 9 parts (Tahun.Satker.Prog.Keg.KRO.RO.Komp.Subkomp.Akun)
        if (parts.length >= 9 && /^\d+$/.test(parts[0]) && /^\d+$/.test(parts[1])) {
            return {
                prog: parts[2],
                keg: parts[3],
                out: parts[4],
                ro: parts[5],
                komp: parts[6],
                subkomp: parts[7],
                akun: parts[8],
                comp_key: `${parts[2]}.${parts[3]}.${parts[4]}.${parts[5]}.${parts[6]}`,
                key: `${parts[2]}.${parts[3]}.${parts[4]}.${parts[5]}.${parts[6]}.${parts[7]}.${parts[8]}`
            };
        }
        // 7 parts (Prog.Keg.KRO.RO.Komp.Subkomp.Akun)
        if (parts.length >= 7) {
            return {
                prog: parts[0],
                keg: parts[1],
                out: parts[2],
                ro: parts[3],
                komp: parts[4],
                subkomp: parts[5],
                akun: parts[6],
                comp_key: `${parts[0]}.${parts[1]}.${parts[2]}.${parts[3]}.${parts[4]}`,
                key: `${parts[0]}.${parts[1]}.${parts[2]}.${parts[3]}.${parts[4]}.${parts[5]}.${parts[6]}`
            };
        }
        return null;
    }

    function parseRkaTable(html) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');
        const trs = Array.from(doc.querySelectorAll('tr'));
        if (trs.length === 0) return [];
        
        const rows = [];
        for (let i = 1; i < trs.length; i++) {
            const tr = trs[i];
            if (!tr.className) continue;
            
            const tds = Array.from(tr.querySelectorAll('td'));
            if (tds.length < 5) continue;
            
            let level = -1;
            if (tr.classList.contains('komponenTr')) level = 0;
            else if (tr.classList.contains('subKomponenTr')) level = 1;
            else if (tr.classList.contains('akunTr')) level = 2;
            else if (tr.classList.contains('detilTr')) level = 3;
            
            const code = tds[0].innerText.trim();
            const desc = tds[1].innerText.trim();
            const descPrev = tds[2].innerText.trim();
            
            const paguText = tds[3].innerText.trim().replace(/\./g, '');
            const pagu = paguText === '-' ? '-' : parseInt(paguText) || 0;
            
            const paguPrevText = tds[4].innerText.trim().replace(/\./g, '');
            const paguPrev = paguPrevText === '-' ? '-' : parseInt(paguPrevText) || 0;
            
            const p_ch = tds[5]?.querySelector('input')?.checked || false;
            const s_ch = tds[6]?.querySelector('input')?.checked || false;
            const my_ch = tds[7]?.querySelector('input')?.checked || false;
            const np_ch = tds[8]?.querySelector('input')?.checked || false;
            const gj_ch = tds[9]?.querySelector('input')?.checked || false;
            
            rows.push({
                level, code, desc, descPrev, pagu, paguPrev,
                p_ch, s_ch, my_ch, np_ch, gj_ch
            });
        }
        return rows;
    }

    // MAIN EXPORT CONTROLLER
    async function startExport() {
        showLogger();
        log('Memulai proses ekspor sanding RKA & RUP...', 5);
        
        try {
            // Satker & tahun anggaran sudah dikonfirmasi lewat panel pra-ekspor
            const activeSatkerId = ctx.satkerId;

            log(`Tahun anggaran: ${ctx.tahun}`, 6);
            if (activeSatkerId) {
                log(`Satker aktif: ${ctx.satkerName || '(nama tidak terdeteksi)'} [ID ${activeSatkerId}]`, 8);
            } else {
                log('Menggunakan Satker bawaan sesi login (ID tidak terdeteksi di halaman).', 8, 'warn');
            }

            // 1. Crawl RKA
            log('Mengambil daftar Program RKA dari server...', 10);
            
            let programSelect = document.querySelector('#idProgram');
            let programs = [];
            if (programSelect) {
                programs = Array.from(programSelect.options)
                    .filter(o => o.value !== "")
                    .map(o => ({ id: o.value, text: o.text }));
            } else {
                let rkaUrl = '/sirup/rkactr/index2';
                if (activeSatkerId) {
                    rkaUrl += `?idSatker=${activeSatkerId}&satker=${activeSatkerId}&id_satker=${activeSatkerId}`;
                }
                log(`Memuat list program via background fetch dari ${rkaUrl}...`, 12);
                const rkaHtml = await getText(rkaUrl);
                const parser = new DOMParser();
                const rkaDoc = parser.parseFromString(rkaHtml, 'text/html');
                const select = rkaDoc.querySelector('#idProgram');
                if (!select) throw new Error('Gagal memuat list Program RKA. Pastikan Anda sudah login.');
                programs = Array.from(select.options)
                    .filter(o => o.value !== "")
                    .map(o => ({ id: o.value, text: o.text }));
            }
            
            const rkaData = [];
            let totalSteps = programs.length;
            let currentStep = 0;
            
            for (const prog of programs) {
                throwIfAborted();
                currentStep++;
                log(`[RKA] Memproses Program: ${prog.text}...`, 10 + (currentStep / totalSteps * 30));
                
                const kegiatans = await postForm('/sirup/selfservice/daftarkegiatanbyprogram', `idProgram=${prog.id}`);
                const progData = { id: prog.id, text: prog.text, kegiatans: [] };
                
                for (const keg of kegiatans) {
                    const outputs = await postForm('/sirup/selfservice/daftaroutputbykegiatan', `idKegiatan=${keg.id}`);
                    const kegData = { id: keg.id, name: keg.nama, code: keg.kode_kegiatans || keg.kode_kegiatand, outputs: [] };
                    
                    for (const out of outputs) {
                        const outData = { id: out.id, name: out.nama, code: out.kode_output_string, suboutputs: [] };
                        
                        // Crawl Sub-outputs (RO / Rincian Output)
                        const suboutputs = await postForm('/sirup/selfservice/daftarsuboutputbyoutput', `idOutput=${out.id}`);
                        
                        for (const ro of suboutputs) {
                            const ro_code = ro.kode_suboutput_string || ro.kode_suboutput || ro.kode_sub_output_string || ro.kode || ro.kode_string || "";
                            const roData = { id: ro.id, name: ro.nama, code: ro_code, komponens: [] };
                            
                            // Crawl Komponens under RO
                            const komponens = await postForm('/sirup/selfservice/daftarkomponenbysuboutput', `idSubOutput=${ro.id}`);
                            
                            // Fetch all component tables concurrently under this RO
                            await Promise.all(komponens.map(async (komp) => {
                                const tableHtml = await getText(`/sirup/rkactr/rkakontentable2018?idKomponen=${komp.id}`);
                                const tableRows = parseRkaTable(tableHtml);
                                
                                roData.komponens.push({
                                    id: komp.id,
                                    name: komp.nama,
                                    code: komp.kode_komponen_string,
                                    pagu: komp.pagu,
                                    rows: tableRows
                                });
                            }));
                            outData.suboutputs.push(roData);
                        }
                        kegData.outputs.push(outData);
                    }
                    progData.kegiatans.push(kegData);
                }
                rkaData.push(progData);
            }
            
            log('Pengambilan data RKA selesai!', 50);

            // 2. Crawl RUP Penyedia via Direct XHR
            log('Mengambil data RUP Paket Penyedia dari server...', 55);
            
            let rupBody = 'draw=1&start=0&length=1000';
            if (activeSatkerId) {
                rupBody += `&idSatker=${activeSatkerId}&satker=${activeSatkerId}&id_satker=${activeSatkerId}`;
            }
            const rupRes = await postForm(`/sirup/datatablectr/dataruppenyedia2018?tahun=${ctx.tahun}`, rupBody);
            const rupRows = rupRes.aaData || [];

            // Paket swakelola belum ikut disandingkan — hitung saja sebagai catatan
            try {
                const swaRes = await postForm(`/sirup/datatablectr/datarupswakelola2018?tahun=${ctx.tahun}`, rupBody);
                const swaRows = swaRes.aaData || [];
                if (swaRows.length > 0) {
                    log(`Catatan: ada ${swaRows.length} paket swakelola di RUP. Laporan ini hanya menyanding paket penyedia.`, 57, 'warn');
                }
            } catch (swaErr) {
                console.error("Swakelola fetch error:", swaErr);
                log('Catatan: daftar paket swakelola tidak bisa diambil (diabaikan).', 57, 'warn');
            }

            const rupPackets = [];
            for (const row of rupRows) {
                rupPackets.push({
                    id: row[0],
                    keg_name: row[1],
                    name: row[2],
                    pagu: parseInt(row[3].toString().replace(/\./g, '')) || 0,
                    waktu: row[4],
                    sumber_dana: row[5],
                    aktif: parseCheckboxValue(row[6]), 
                    fd: parseCheckboxValue(row[7]), 
                    umumkan: parseCheckboxValue(row[8]), 
                    mak: row[13] || ""
                });
            }
            
            log(`Ditemukan ${rupPackets.length} paket. Mengambil detail MAK sub-paket secara concurrent (kecepatan tinggi)...`, 65);
            
            const rupDetails = {};
            let fetchedCount = 0;
            
            await mapConcurrent(rupPackets, 10, async (p) => {
                throwIfAborted();
                const detailHtml = await getText(`/sirup/penyedia/${p.id}`);
                const parser = new DOMParser();
                const detailDoc = parser.parseFromString(detailHtml, 'text/html');
                
                const tables = Array.from(detailDoc.querySelectorAll('table'));
                let fundingTable = null;
                for (const t of tables) {
                    const headerText = t.innerText.toLowerCase();
                    if (headerText.includes('sumber dana') && headerText.includes('mak') && headerText.includes('pagu')) {
                        fundingTable = t;
                        break;
                    }
                }
                
                const items = [];
                if (fundingTable) {
                    const trs = Array.from(fundingTable.querySelectorAll('tr'));
                    for (let i = 1; i < trs.length; i++) {
                        const tr = trs[i];
                        const tds = Array.from(tr.querySelectorAll('td'));
                        if (tds.length < 6) continue;
                        
                        const noText = tds[0].innerText.trim();
                        if (!/^\d+\.?$/.test(noText)) continue;
                        
                        const mak = tds[4].innerText.trim();
                        const paguText = tds[5].innerText.trim().replace(/Rp\./g, '').replace(/\./g, '').trim();
                        const pagu = parseInt(paguText) || 0;
                        
                        items.push({ mak, pagu });
                    }
                }
                rupDetails[p.id] = items;
                
                fetchedCount++;
                if (fetchedCount % 5 === 0 || fetchedCount === rupPackets.length) {
                    log(`[RUP] Mengambil detail sub-paket: ${fetchedCount}/${rupPackets.length} selesai...`, 65 + (fetchedCount / rupPackets.length * 20));
                }
            });
            
            // mapConcurrent menelan error per item, jadi status batal diperiksa lagi
            // di sini supaya laporan setengah jadi tidak ikut dibangun.
            throwIfAborted();

            log('Pemuatan data RKA & RUP selesai! Memproses penyandingan...', 85);
            
            // 3. Process data & Match
            const rup_all_lines = [];
            
            for (const p of rupPackets) {
                const items = rupDetails[p.id] || [];
                const is_terumumkan = p.aktif && p.fd && p.umumkan;
                for (const item of items) {
                    const mak = item.mak;
                    const pagu = item.pagu;
                    
                    const parsed = parseMak(mak);
                    if (parsed) {
                        rup_all_lines.push({
                            packet_id: p.id,
                            packet_name: p.name,
                            mak: mak,
                            pagu: pagu,
                            prog: parsed.prog,
                            keg: parsed.keg,
                            out: parsed.out,
                            ro: parsed.ro,
                            komp: parsed.komp,
                            subkomp: parsed.subkomp,
                            akun: parsed.akun,
                            comp_key: parsed.comp_key,
                            key: parsed.key,
                            is_terumumkan: is_terumumkan
                        });
                    }
                }
            }
            
            log('Membangun file Excel...', 90);

            // 4. Generate Excel using ExcelJS
            const summary = await buildExcel(rkaData, rupPackets, rup_all_lines);

            log(`Berhasil! File "${summary.fileName}" sudah diunduh.`, 100, 'success');
            finishSuccess([
                { label: 'Paket RUP', value: summary.packets.toLocaleString('id-ID') },
                { label: 'Capaian Umumkan', value: (summary.pct * 100).toFixed(1) + '%' },
                { label: 'Tanpa Sandingan', value: summary.unmatched.toLocaleString('id-ID') }
            ]);

        } catch (e) {
            console.error(e);
            log('ERROR: ' + e.message, null, 'error');
            finishError(e);
        }
    }

    // EXCEL BUILDER FUNCTION (ExcelJS)
    async function buildExcel(rkaData, rupPackets, rup_all_lines) {
        const wb = new ExcelJS.Workbook();

        wb.creator = 'SiRUP Sanding RKA & RUP v' + APP_VERSION;
        wb.lastModifiedBy = wb.creator;
        wb.created = new Date();
        wb.modified = new Date();
        wb.title = `Sanding RKA & RUP TA ${ctx.tahun}`;
        wb.company = ctx.satkerName || '';

        // Palet warna. Semua nilai WAJIB 8 digit ARGB (FF + RRGGBB) — nilai
        // 6 digit ditulis apa adanya ke XML dan dibaca salah oleh Excel.
        const DARK_BLUE = "FF1F497D";        // header utama
        const MID_BLUE = "FF2E6DA4";         // aksen
        const LIGHT_BLUE_LVL0 = "FFB8CCE4";  // subtotal program
        const LIGHT_BLUE_LVL1 = "FFDCE6F1";  // subtotal kegiatan
        const LIGHT_GRAY_LVL2 = "FFF2F2F2";  // subtotal KRO
        const LIGHT_GRAY_LVL3 = "FFF9F9F9";  // subtotal RO
        const LIGHT_GREEN = "FFE2EFDA";
        const LIGHT_RED = "FFFCE4D6";
        const WHITE_COLOR = "FFFFFFFF";
        const LIGHT_ORANGE = "FFFFF2CC";
        const ZEBRA = "FFF7F9FC";            // baris selang-seling
        const INK = "FF0F172A";
        const INK_SOFT = "FF595959";

        const FONT = "Segoe UI";
        const FMT_RP = '#,##0;[Red]-#,##0;"-"';          // dalam tabel (header sudah menyebut Rp)
        const FMT_RP_FULL = '"Rp"\\ #,##0;[Red]-"Rp"\\ #,##0';  // untuk kartu KPI
        const FMT_PCT = '0.0%';

        // Borders template
        const border_thin = {
            top: { style: 'thin', color: { argb: 'FFD9DEE6' } },
            left: { style: 'thin', color: { argb: 'FFD9DEE6' } },
            bottom: { style: 'thin', color: { argb: 'FFD9DEE6' } },
            right: { style: 'thin', color: { argb: 'FFD9DEE6' } }
        };

        const solid = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
        const rp = (n) => 'Rp ' + (Number(n) || 0).toLocaleString('id-ID');
        const stamp = new Date().toLocaleString('id-ID', { dateStyle: 'long', timeStyle: 'short' });

        // Kop seragam untuk semua sheet: judul + baris konteks satker/tahun.
        function writeTitle(ws, title, lastCol, subtitle) {
            const span = `A1:${lastCol}1`;
            ws.mergeCells(span);
            const c = ws.getCell('A1');
            c.value = title;
            c.font = { name: FONT, size: 15, bold: true, color: { argb: WHITE_COLOR } };
            c.fill = solid(DARK_BLUE);
            c.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
            ws.getRow(1).height = 30;

            ws.mergeCells(`A2:${lastCol}2`);
            const s = ws.getCell('A2');
            s.value = subtitle || `${ctx.satkerName || 'Satker'}  •  Tahun Anggaran ${ctx.tahun}  •  Diekspor ${stamp}`;
            s.font = { name: FONT, size: 9, italic: true, color: { argb: INK_SOFT } };
            s.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
            ws.getRow(2).height = 16;
        }

        function writeHeader(ws, rowNum, headers, widths) {
            const row = ws.getRow(rowNum);
            row.height = 32;
            for (let c = 1; c <= headers.length; c++) {
                const cell = ws.getCell(rowNum, c);
                cell.value = headers[c - 1];
                cell.font = { name: FONT, size: 9.5, bold: true, color: { argb: WHITE_COLOR } };
                cell.fill = solid(DARK_BLUE);
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                cell.border = border_thin;
            }
            if (widths) widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
        }

        // Freeze pane + autofilter + setelan cetak, dipanggil setelah sheet penuh.
        function finishSheet(ws, opts) {
            const o = opts || {};
            ws.views = [{
                state: 'frozen',
                xSplit: o.freezeCols || 0,
                ySplit: o.headerRow || 0,
                showGridLines: false,
                activeCell: 'A' + ((o.headerRow || 0) + 1)
            }];
            if (o.headerRow && o.lastCol && o.lastRow && o.lastRow > o.headerRow) {
                ws.autoFilter = {
                    from: { row: o.headerRow, column: o.filterFromCol || 1 },
                    to: { row: o.lastRow, column: o.lastCol }
                };
            }
            if (o.tabColor) ws.properties.tabColor = { argb: o.tabColor };

            ws.pageSetup = {
                orientation: o.orientation || 'landscape',
                paperSize: 9,                       // A4
                fitToPage: true,
                fitToWidth: 1,
                fitToHeight: 0,
                horizontalCentered: true,
                margins: { left: 0.3, right: 0.3, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
                printTitlesRow: o.headerRow ? `${o.headerRow}:${o.headerRow}` : undefined
            };
            ws.headerFooter = {
                oddFooter: '&L&9' + (ctx.satkerName || '') + ' — TA ' + ctx.tahun + '&R&9Hal &P dari &N'
            };
        }

        // Selang-seling baris data supaya tabel panjang tetap terbaca.
        function zebra(ws, fromRow, toRow, lastCol) {
            for (let r = fromRow; r <= toRow; r++) {
                if ((r - fromRow) % 2 === 0) continue;
                for (let c = 1; c <= lastCol; c++) {
                    const cell = ws.getCell(r, c);
                    if (!cell.fill || cell.fill.type !== 'pattern') cell.fill = solid(ZEBRA);
                }
            }
        }

        // --- 0. PRE-CALCULATE ALL TOTALS & STATS ---
        let total_satker_pagu = 0;
        let total_non_pengadaan = 0;
        let total_target_pengadaan = 0;
        let total_rup_pagu = 0;
        let tot_pct = 0;
        
        // Sum total RKA pagus and non-pengadaan
        for (const prog of rkaData) {
            const prog_code = prog.text.split("]")[0].replace("[", "").trim();
            for (const keg of prog.kegiatans || []) {
                for (const out of keg.outputs || []) {
                    const suboutputs = out.suboutputs || [{ id: 0, name: out.name, code: "000", komponens: out.komponens || [] }];
                    for (const ro of suboutputs) {
                        for (const komp of ro.komponens || []) {
                            total_satker_pagu += komp.pagu || 0;
                            
                            let komp_np = false, komp_gj = false;
                            let subkomp_np = false, subkomp_gj = false;
                            let akun_np = false, akun_gj = false;
                            
                            for (const r of komp.rows || []) {
                                if (r.level === 0) {
                                    komp_np = r.np_ch;
                                    komp_gj = r.gj_ch;
                                } else if (r.level === 1) {
                                    subkomp_np = r.np_ch;
                                    subkomp_gj = r.gj_ch;
                                } else if (r.level === 2) {
                                    akun_np = r.np_ch;
                                    akun_gj = r.gj_ch;
                                } else if (r.level === 3) {
                                    const is_np = r.np_ch || akun_np || subkomp_np || komp_np;
                                    const is_gj = r.gj_ch || akun_gj || subkomp_gj || komp_gj;
                                    if (is_np || is_gj) {
                                        total_non_pengadaan += typeof r.pagu === 'number' ? r.pagu : 0;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        total_target_pengadaan = total_satker_pagu - total_non_pengadaan;
        
        // Sum RUP totals (Only fully announced ones) + kumpulkan statistik per
        // komponen untuk tabel "selisih terbesar" di Dashboard.
        const comp_stats = [];
        for (const prog of rkaData) {
            const prog_code = prog.text.split("]")[0].replace("[", "").trim();
            for (const keg of prog.kegiatans || []) {
                for (const out of keg.outputs || []) {
                    const suboutputs = out.suboutputs || [{ id: 0, name: out.name, code: "000", komponens: out.komponens || [] }];
                    for (const ro of suboutputs) {
                        for (const komp of ro.komponens || []) {
                            const comp_key = `${prog_code}.${keg.code}.${out.code}.${ro.code}.${komp.code}`;
                            const comp_rup = rup_all_lines.filter(l => l.comp_key === comp_key && l.is_terumumkan);
                            const rup_sum = comp_rup.reduce((sum, l) => sum + l.pagu, 0);
                            total_rup_pagu += rup_sum;

                            let komp_np = false, komp_gj = false;
                            let subkomp_np = false, subkomp_gj = false;
                            let akun_np = false, akun_gj = false;
                            let np_gaji_sum = 0;

                            for (const r of komp.rows || []) {
                                if (r.level === 0) {
                                    komp_np = r.np_ch; komp_gj = r.gj_ch;
                                } else if (r.level === 1) {
                                    subkomp_np = r.np_ch; subkomp_gj = r.gj_ch;
                                } else if (r.level === 2) {
                                    akun_np = r.np_ch; akun_gj = r.gj_ch;
                                } else if (r.level === 3) {
                                    const is_np = r.np_ch || akun_np || subkomp_np || komp_np;
                                    const is_gj = r.gj_ch || akun_gj || subkomp_gj || komp_gj;
                                    if (is_np || is_gj) np_gaji_sum += (typeof r.pagu === 'number' ? r.pagu : 0);
                                }
                            }

                            const target = (komp.pagu || 0) - np_gaji_sum;
                            comp_stats.push({
                                prog_code,
                                keg_name: keg.name,
                                ro_name: ro.name,
                                komp_name: komp.name,
                                comp_key,
                                target,
                                rup: rup_sum,
                                gap: target - rup_sum
                            });
                        }
                    }
                }
            }
        }
        const top_gaps = comp_stats.filter(c => c.gap > 0).sort((a, b) => b.gap - a.gap).slice(0, 10);

        // Gather all RKA detailed Akun keys (7-parts)
        const rka_detailed_keys = new Set();
        for (const prog of rkaData) {
            const prog_code = prog.text.split("]")[0].replace("[", "").trim();
            for (const keg of prog.kegiatans || []) {
                for (const out of keg.outputs || []) {
                    const suboutputs = out.suboutputs || [{ id: 0, name: out.name, code: "000", komponens: out.komponens || [] }];
                    for (const ro of suboutputs) {
                        for (const komp of ro.komponens || []) {
                            const comp_key = `${prog_code}.${keg.code}.${out.code}.${ro.code}.${komp.code}`;
                            let subkomp_code = "";
                            for (const r of komp.rows || []) {
                                if (r.level === 1) {
                                    subkomp_code = r.code;
                                } else if (r.level === 2) {
                                    rka_detailed_keys.add(`${comp_key}.${subkomp_code}.${r.code}`);
                                }
                            }
                        }
                    }
                }
            }
        }

        const unmatched_rup_lines = [];
        // Check each line in RUP against 7-part RKA key
        for (const line of rup_all_lines) {
            if (!rka_detailed_keys.has(line.key)) {
                unmatched_rup_lines.push({
                    packet_id: line.packet_id,
                    packet_name: line.packet_name,
                    mak: line.mak,
                    pagu: line.pagu,
                    reason: "Mata Anggaran (MAK) tidak ditemukan di RKA (Anggaran dihapus atau salah input)"
                });
            }
        }
        for (const p of rupPackets) {
            const p_lines = rup_all_lines.filter(l => l.packet_id === p.id);
            if (p_lines.length === 0) {
                const parsed = parseMak(p.mak);
                if (parsed && !rka_detailed_keys.has(parsed.key)) {
                    unmatched_rup_lines.push({
                        packet_id: p.id,
                        packet_name: p.name,
                        mak: p.mak || "(Kosong)",
                        pagu: p.pagu,
                        reason: "Mata Anggaran (MAK) tidak ditemukan di RKA (Anggaran dihapus atau salah input)"
                    });
                } else if (!parsed) {
                    unmatched_rup_lines.push({
                        packet_id: p.id,
                        packet_name: p.name,
                        mak: p.mak || "(Kosong)",
                        pagu: p.pagu,
                        reason: "Format Kode MAK di RUP tidak valid atau kosong"
                    });
                }
            }
        }
        // Deduplicate
        const seen_unmatched = new Set();
        const unique_unmatched = [];
        for (const item of unmatched_rup_lines) {
            const key = `${item.packet_id}_${item.mak}`;
            if (!seen_unmatched.has(key)) {
                seen_unmatched.add(key);
                unique_unmatched.push(item);
            }
        }
        const unique_unmatched_count = unique_unmatched.length;

        // ----------------- SHEET 0: DASHBOARD EVALUASI -----------------
        // Grid dashboard: 8 kolom (A..H) dibagi jadi 4 slot kartu @2 kolom,
        // supaya kartu KPI, tabel statistik, dan tabel selisih tetap sejajar.
        const ws_dash = wb.addWorksheet("Dashboard Evaluasi");
        [22, 16, 22, 16, 22, 16, 22, 16].forEach((w, i) => { ws_dash.getColumn(i + 1).width = w; });

        writeTitle(ws_dash, "DASHBOARD MONITORING & EVALUASI INTEGRASI RKA–RUP", "H");

        tot_pct = total_target_pengadaan > 0 ? (total_rup_pagu / total_target_pengadaan) : 0;

        // --- Kartu KPI (baris 4-6) ---
        const pct_color = tot_pct >= 0.999 ? 'FF15803D' : (tot_pct >= 0.75 ? 'FFB45309' : 'FFB91C1C');
        const cards = [
            { col: 1, label: "TOTAL PAGU SATKER (A)", value: total_satker_pagu, fmt: FMT_RP_FULL, head: DARK_BLUE, ink: DARK_BLUE },
            { col: 3, label: "NON-PENGADAAN (B)", value: total_non_pengadaan, fmt: FMT_RP_FULL, head: "FF7F7F7F", ink: "FF595959" },
            { col: 5, label: "TARGET PENGADAAN (C = A − B)", value: total_target_pengadaan, fmt: FMT_RP_FULL, head: "FF203764", ink: "FF203764" },
            { col: 7, label: "RUP TERUMUMKAN (D)", value: total_rup_pagu, fmt: FMT_RP_FULL, head: "FF375623", ink: "FF375623" }
        ];

        for (const card of cards) {
            const c1 = card.col, c2 = card.col + 1;
            ws_dash.mergeCells(4, c1, 4, c2);
            const head = ws_dash.getCell(4, c1);
            head.value = card.label;
            head.font = { name: FONT, size: 8.5, bold: true, color: { argb: WHITE_COLOR } };
            head.fill = solid(card.head);
            head.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

            ws_dash.mergeCells(5, c1, 6, c2);
            const val = ws_dash.getCell(5, c1);
            val.value = card.value;
            val.numFmt = card.fmt;
            val.font = { name: FONT, size: 14, bold: true, color: { argb: card.ink } };
            val.alignment = { horizontal: 'center', vertical: 'middle' };
            val.fill = solid(WHITE_COLOR);
            for (let r = 4; r <= 6; r++) {
                for (let c = c1; c <= c2; c++) ws_dash.getCell(r, c).border = border_thin;
            }
        }
        ws_dash.getRow(4).height = 26;
        ws_dash.getRow(6).height = 20;

        // --- Baris capaian (baris 8) ---
        ws_dash.mergeCells("A8:E8");
        const cap_label = ws_dash.getCell('A8');
        cap_label.value = "CAPAIAN PENGUMUMAN RUP TERHADAP TARGET PENGADAAN (D ÷ C)";
        cap_label.font = { name: FONT, size: 10, bold: true, color: { argb: WHITE_COLOR } };
        cap_label.fill = solid(MID_BLUE);
        cap_label.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

        ws_dash.mergeCells("F8:H8");
        const cap_val = ws_dash.getCell('F8');
        cap_val.value = tot_pct;
        cap_val.numFmt = FMT_PCT;
        cap_val.font = { name: FONT, size: 14, bold: true, color: { argb: pct_color } };
        cap_val.fill = solid(tot_pct >= 0.999 ? LIGHT_GREEN : LIGHT_ORANGE);
        cap_val.alignment = { horizontal: 'center', vertical: 'middle' };
        for (let c = 1; c <= 8; c++) ws_dash.getCell(8, c).border = border_thin;
        ws_dash.getRow(8).height = 26;

        // --- Tabel statistik (header baris 10) ---
        ws_dash.getCell('A10').value = "STATISTIK EVALUASI PAKET RKA–RUP";
        ws_dash.getCell('A10').font = { name: FONT, size: 11, bold: true, color: { argb: DARK_BLUE } };

        const stat_head_row = 11;
        [['A', 'C', "Indikator Evaluasi Anggaran & RUP"],
         ['D', 'E', "Nilai / Jumlah"],
         ['F', 'H', "Satuan"]].forEach(([a, b, label]) => {
            ws_dash.mergeCells(`${a}${stat_head_row}:${b}${stat_head_row}`);
            const cell = ws_dash.getCell(`${a}${stat_head_row}`);
            cell.value = label;
            cell.font = { name: FONT, size: 9.5, bold: true, color: { argb: WHITE_COLOR } };
            cell.fill = solid(DARK_BLUE);
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
        });
        for (let c = 1; c <= 8; c++) ws_dash.getCell(stat_head_row, c).border = border_thin;
        ws_dash.getRow(stat_head_row).height = 22;

        const pkt_sah = rupPackets.filter(p => p.aktif && p.fd && p.umumkan).length;
        const pkt_draft = rupPackets.length - pkt_sah;
        const stats_data = [
            ["Total belanja non-pengadaan (NP / Gaji satker) — komponen B", total_non_pengadaan, "Rupiah", false],
            ["Selisih pengadaan yang belum diumumkan (C − D)", total_target_pengadaan - total_rup_pagu, "Rupiah", (total_target_pengadaan - total_rup_pagu) > 0],
            ["Jumlah paket RUP terdaftar (penyedia)", rupPackets.length, "Paket", false],
            ["Jumlah paket RUP sah (terumumkan KPA)", pkt_sah, "Paket", false],
            ["Jumlah paket RUP belum terumumkan (draft / dibatalkan)", pkt_draft, "Paket", pkt_draft > 0],
            ["Jumlah paket RUP tanpa sandingan RKA (potensi salah input MAK)", unique_unmatched_count, "Paket", unique_unmatched_count > 0]
        ];

        stats_data.forEach((row, i) => {
            const r = stat_head_row + 1 + i;
            ws_dash.mergeCells(`A${r}:C${r}`);
            ws_dash.mergeCells(`D${r}:E${r}`);
            ws_dash.mergeCells(`F${r}:H${r}`);

            const label = ws_dash.getCell(`A${r}`);
            label.value = row[0];
            label.alignment = { horizontal: 'left', vertical: 'middle', indent: 1, wrapText: true };

            const val = ws_dash.getCell(`D${r}`);
            val.value = row[1];
            val.numFmt = row[2] === "Rupiah" ? FMT_RP : '#,##0';
            val.alignment = { horizontal: 'right', vertical: 'middle', indent: 1 };

            const unit = ws_dash.getCell(`F${r}`);
            unit.value = row[2];
            unit.alignment = { horizontal: 'center', vertical: 'middle' };

            for (let c = 1; c <= 8; c++) {
                const cell = ws_dash.getCell(r, c);
                cell.border = border_thin;
                cell.font = { name: FONT, size: 9.5, color: { argb: INK } };
                if (i % 2 === 1) cell.fill = solid(ZEBRA);
            }
            if (row[3]) {
                val.font = { name: FONT, size: 9.5, bold: true, color: { argb: 'FFC00000' } };
                val.fill = solid(LIGHT_ORANGE);
            }
            ws_dash.getRow(r).height = 18;
        });

        // --- Tabel 10 selisih terbesar ---
        let dash_row = stat_head_row + stats_data.length + 3;
        ws_dash.getCell(`A${dash_row}`).value = "10 KOMPONEN DENGAN SELISIH PENGADAAN TERBESAR";
        ws_dash.getCell(`A${dash_row}`).font = { name: FONT, size: 11, bold: true, color: { argb: DARK_BLUE } };
        dash_row++;

        const gap_head_row = dash_row;
        [['A', 'C', "Komponen"],
         ['D', 'E', "Target Pengadaan"],
         ['F', 'G', "RUP Terumumkan"],
         ['H', 'H', "Selisih"]].forEach(([a, b, label]) => {
            if (a !== b) ws_dash.mergeCells(`${a}${gap_head_row}:${b}${gap_head_row}`);
            const cell = ws_dash.getCell(`${a}${gap_head_row}`);
            cell.value = label;
            cell.font = { name: FONT, size: 9.5, bold: true, color: { argb: WHITE_COLOR } };
            cell.fill = solid(DARK_BLUE);
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
        });
        for (let c = 1; c <= 8; c++) ws_dash.getCell(gap_head_row, c).border = border_thin;
        ws_dash.getRow(gap_head_row).height = 20;

        dash_row = gap_head_row + 1;
        if (top_gaps.length === 0) {
            ws_dash.mergeCells(`A${dash_row}:H${dash_row}`);
            const cell = ws_dash.getCell(`A${dash_row}`);
            cell.value = "Tidak ada selisih — seluruh target pengadaan sudah terumumkan di RUP.";
            cell.font = { name: FONT, size: 9.5, italic: true, color: { argb: 'FF166534' } };
            cell.fill = solid(LIGHT_GREEN);
            cell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
            for (let c = 1; c <= 8; c++) ws_dash.getCell(dash_row, c).border = border_thin;
            dash_row++;
        } else {
            top_gaps.forEach((g, i) => {
                const r = dash_row + i;
                ws_dash.mergeCells(`A${r}:C${r}`);
                ws_dash.mergeCells(`D${r}:E${r}`);
                ws_dash.mergeCells(`F${r}:G${r}`);

                const name = ws_dash.getCell(`A${r}`);
                name.value = `[${g.prog_code}] ${g.komp_name}`;
                name.alignment = { horizontal: 'left', vertical: 'middle', indent: 1, wrapText: true };
                name.note = `Kegiatan: ${g.keg_name}\nRO: ${g.ro_name}\nMAK: ${g.comp_key}`;

                ws_dash.getCell(`D${r}`).value = g.target;
                ws_dash.getCell(`D${r}`).numFmt = FMT_RP;
                ws_dash.getCell(`F${r}`).value = g.rup;
                ws_dash.getCell(`F${r}`).numFmt = FMT_RP;
                ws_dash.getCell(`H${r}`).value = g.gap;
                ws_dash.getCell(`H${r}`).numFmt = FMT_RP;

                for (let c = 1; c <= 8; c++) {
                    const cell = ws_dash.getCell(r, c);
                    cell.border = border_thin;
                    cell.font = { name: FONT, size: 9.5, color: { argb: INK } };
                    if (c >= 4) cell.alignment = { horizontal: 'right', vertical: 'middle', indent: 1 };
                    if (i % 2 === 1) cell.fill = solid(ZEBRA);
                }
                ws_dash.getCell(`H${r}`).font = { name: FONT, size: 9.5, bold: true, color: { argb: 'FFC00000' } };
                ws_dash.getRow(r).height = 18;
            });
            dash_row += top_gaps.length;
        }

        // --- Legenda warna status ---
        dash_row += 1;
        ws_dash.getCell(`A${dash_row}`).value = "KETERANGAN WARNA STATUS (sheet Sanding & Detail)";
        ws_dash.getCell(`A${dash_row}`).font = { name: FONT, size: 10, bold: true, color: { argb: DARK_BLUE } };
        dash_row++;

        [[1, LIGHT_GREEN, "Sesuai — RUP sudah menutup target"],
         [3, LIGHT_BLUE_LVL1, "Parsial — baru sebagian diumumkan"],
         [5, LIGHT_RED, "Kelebihan / belum diumumkan"],
         [7, LIGHT_ORANGE, "Paket masih draft atau dibatalkan"]].forEach(([col, color, text]) => {
            const sw = ws_dash.getCell(dash_row, col);
            sw.fill = solid(color);
            sw.border = border_thin;
            const lbl = ws_dash.getCell(dash_row, col + 1);
            lbl.value = text;
            lbl.font = { name: FONT, size: 8.5, color: { argb: INK_SOFT } };
            lbl.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
        });
        ws_dash.getRow(dash_row).height = 26;

        finishSheet(ws_dash, { headerRow: 0, tabColor: DARK_BLUE, orientation: 'portrait' });

        // ----------------- SHEET 1: RINGKASAN PAGU -----------------
        const ws_summary = wb.addWorksheet("Ringkasan Pagu");

        writeTitle(ws_summary, `RINGKASAN PAGU ANGGARAN SATKER TA ${ctx.tahun}`, "L");

        ws_summary.mergeCells("A4:C4");
        ws_summary.getCell('A4').value = "TOTAL PAGU SATKER";
        ws_summary.getCell('A4').font = { name: FONT, size: 10, bold: true, color: { argb: WHITE_COLOR } };
        ws_summary.getCell('A4').fill = solid(DARK_BLUE);
        ws_summary.getCell('A4').alignment = { horizontal: 'center', vertical: 'middle' };

        total_satker_pagu = 0;

        const headers_summary = [
            "No", "Kode Program", "Nama Program", "Kode Kegiatan", "Nama Kegiatan",
            "Kode KRO", "Nama KRO", "Kode RO", "Nama RO", "Kode Komponen", "Nama Komponen", "Pagu Komponen"
        ];

        writeHeader(ws_summary, 7, headers_summary,
            [6, 13, 28, 13, 28, 13, 28, 13, 28, 13, 34, 18]);

        let row_idx = 8;
        let num = 1;
        
        for (const prog of rkaData) {
            const prog_code = prog.text.split("]")[0].replace("[", "").trim();
            const prog_name = prog.text.split("]")[1].trim();
            
            for (const keg of prog.kegiatans || []) {
                for (const out of keg.outputs || []) {
                    const suboutputs = out.suboutputs || [{ id: 0, name: out.name, code: "000", komponens: out.komponens || [] }];
                    
                    for (const ro of suboutputs) {
                        for (const komp of ro.komponens || []) {
                            ws_summary.getCell(row_idx, 1).value = num;
                            ws_summary.getCell(row_idx, 2).value = prog_code;
                            ws_summary.getCell(row_idx, 3).value = prog_name;
                            ws_summary.getCell(row_idx, 4).value = keg.code;
                            ws_summary.getCell(row_idx, 5).value = keg.name;
                            ws_summary.getCell(row_idx, 6).value = out.code;
                            ws_summary.getCell(row_idx, 7).value = out.name;
                            ws_summary.getCell(row_idx, 8).value = ro.code;
                            ws_summary.getCell(row_idx, 9).value = ro.name;
                            ws_summary.getCell(row_idx, 10).value = komp.code;
                            ws_summary.getCell(row_idx, 11).value = komp.name;
                            ws_summary.getCell(row_idx, 12).value = komp.pagu;
                            ws_summary.getCell(row_idx, 12).numFmt = FMT_RP;

                            for (let col_c = 1; col_c <= 12; col_c++) {
                                const cell = ws_summary.getCell(row_idx, col_c);
                                cell.border = border_thin;
                                cell.font = { name: FONT, size: 9, color: { argb: INK } };
                                if (col_c === 1 || col_c === 2 || col_c === 4 || col_c === 6 || col_c === 8 || col_c === 10) {
                                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                                } else if (col_c === 12) {
                                    cell.alignment = { horizontal: 'right', vertical: 'middle' };
                                } else {
                                    cell.alignment = { horizontal: 'left', vertical: 'middle' };
                                }
                            }
                            
                            total_satker_pagu += komp.pagu;
                            row_idx++;
                            num++;
                        }
                    }
                }
            }
        }

        zebra(ws_summary, 8, row_idx - 1, 12);

        // Baris TOTAL
        for (let c = 1; c <= 12; c++) {
            const cell = ws_summary.getCell(row_idx, c);
            cell.fill = solid(LIGHT_BLUE_LVL0);
            cell.border = border_thin;
            cell.font = { name: FONT, size: 10, bold: true, color: { argb: DARK_BLUE } };
        }
        ws_summary.getCell(row_idx, 11).value = "TOTAL PAGU SATKER";
        ws_summary.getCell(row_idx, 11).alignment = { horizontal: 'right', vertical: 'middle' };
        ws_summary.getCell(row_idx, 12).value = total_satker_pagu;
        ws_summary.getCell(row_idx, 12).numFmt = FMT_RP;
        ws_summary.getCell(row_idx, 12).alignment = { horizontal: 'right', vertical: 'middle' };
        ws_summary.getRow(row_idx).height = 20;

        // Kartu total di kop sheet
        ws_summary.mergeCells("A5:C6");
        const sum_kpi = ws_summary.getCell('A5');
        sum_kpi.value = total_satker_pagu;
        sum_kpi.font = { name: FONT, size: 18, bold: true, color: { argb: DARK_BLUE } };
        sum_kpi.numFmt = FMT_RP_FULL;
        sum_kpi.alignment = { horizontal: 'center', vertical: 'middle' };
        for (let r = 4; r <= 6; r++) {
            for (let c = 1; c <= 3; c++) ws_summary.getCell(r, c).border = border_thin;
        }

        finishSheet(ws_summary, { headerRow: 7, lastRow: row_idx - 1, lastCol: 12, tabColor: MID_BLUE });


        // ----------------- SHEET 2: SANDING RKA & RUP -----------------
        const ws_sanding = wb.addWorksheet("Sanding RKA & RUP");

        writeTitle(ws_sanding, "PENYANDINGAN PAGU RKA DENGAN RUP TERUMUMKAN", "N");

        const headers_sanding = [
            "No", "Program", "Kegiatan", "KRO (Output)", "RO (Sub-Output)", "Komponen",
            "Key Otorisasi (MAK)", "Pagu RKA (A)", "Non-Pengadaan (B)",
            "Target Pengadaan (C = A − B)", "RUP Terumumkan (D)",
            "Selisih (C − D)", "Capaian", "Status Evaluasi"
        ];

        writeHeader(ws_sanding, 4, headers_sanding,
            [6, 10, 26, 26, 26, 28, 24, 16, 16, 18, 18, 18, 11, 24]);

        function writeSandingRow(rowNum, numVal, prog_code, keg_name, kro_name, ro_name, comp_name, comp_key,
                                  pagu_rka, np_gaji_sum, target_pengadaan, rup_sum,
                                  bg_fill, font_style) {
            const selisih_pengadaan = target_pengadaan - rup_sum;
            
            let pct = 0;
            let status_text = "";
            let status_fill = WHITE_COLOR;
            let status_ink = INK;

            if (target_pengadaan === 0) {
                if (rup_sum === 0) {
                    pct = 1.0;
                    status_text = "Sesuai (Non-Pengadaan)";
                    status_fill = LIGHT_GREEN;
                    status_ink = 'FF166534';
                } else {
                    // Tidak ada target tapi ada paket diumumkan: persentase tak
                    // terdefinisi, jadi sel capaian dikosongkan.
                    pct = null;
                    status_text = "Kelebihan Umumkan";
                    status_fill = LIGHT_RED;
                    status_ink = 'FF991B1B';
                }
            } else {
                pct = rup_sum / target_pengadaan;
                if (pct === 1.0) {
                    status_text = "Sesuai (100%)";
                    status_fill = LIGHT_GREEN;
                    status_ink = 'FF166534';
                } else if (pct > 0 && pct < 1.0) {
                    status_text = `Parsial (${(pct*100).toFixed(1)}%)`;
                    status_fill = LIGHT_BLUE_LVL1;
                    status_ink = 'FF1E40AF';
                } else if (pct > 1.0) {
                    status_text = `Kelebihan (${(pct*100).toFixed(1)}%)`;
                    status_fill = LIGHT_RED;
                    status_ink = 'FF991B1B';
                } else {
                    status_text = "Belum Diumumkan";
                    status_fill = LIGHT_RED;
                    status_ink = 'FF991B1B';
                }
            }

            ws_sanding.getCell(rowNum, 1).value = numVal;
            ws_sanding.getCell(rowNum, 2).value = prog_code;
            ws_sanding.getCell(rowNum, 3).value = keg_name;
            ws_sanding.getCell(rowNum, 4).value = kro_name;
            ws_sanding.getCell(rowNum, 5).value = ro_name;
            ws_sanding.getCell(rowNum, 6).value = comp_name;
            ws_sanding.getCell(rowNum, 7).value = comp_key;
            
            ws_sanding.getCell(rowNum, 8).value = pagu_rka;
            ws_sanding.getCell(rowNum, 8).numFmt = FMT_RP;

            ws_sanding.getCell(rowNum, 9).value = np_gaji_sum;
            ws_sanding.getCell(rowNum, 9).numFmt = FMT_RP;

            ws_sanding.getCell(rowNum, 10).value = target_pengadaan;
            ws_sanding.getCell(rowNum, 10).numFmt = FMT_RP;

            ws_sanding.getCell(rowNum, 11).value = rup_sum;
            ws_sanding.getCell(rowNum, 11).numFmt = FMT_RP;

            ws_sanding.getCell(rowNum, 12).value = selisih_pengadaan;
            ws_sanding.getCell(rowNum, 12).numFmt = FMT_RP;

            ws_sanding.getCell(rowNum, 13).value = pct;
            ws_sanding.getCell(rowNum, 13).numFmt = FMT_PCT;

            for (let col_c = 1; col_c <= 14; col_c++) {
                const cell = ws_sanding.getCell(rowNum, col_c);
                cell.border = border_thin;
                cell.font = font_style;
                // Kolom 14 memakai warna status, jangan ditimpa warna subtotal.
                if (bg_fill && col_c !== 14) {
                    cell.fill = solid(bg_fill);
                }

                if (col_c === 1 || col_c === 2 || col_c === 7 || col_c === 14) {
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                } else if (col_c >= 8 && col_c <= 13) {
                    cell.alignment = { horizontal: 'right', vertical: 'middle', indent: 1 };
                } else {
                    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
                }
            }

            // Kolom status: warna & teks sendiri, ditulis terakhir supaya menang.
            const status_cell = ws_sanding.getCell(rowNum, 14);
            status_cell.value = status_text;
            status_cell.fill = solid(status_fill);
            status_cell.font = { name: FONT, size: 9, bold: true, color: { argb: status_ink } };

            // Selisih yang masih menganga ditandai merah.
            if (selisih_pengadaan > 0) {
                ws_sanding.getCell(rowNum, 12).font = {
                    name: FONT, size: font_style.size || 9, bold: font_style.bold || false, color: { argb: 'FFC00000' }
                };
            }
        }

        row_idx = 5;
        num = 1;
        total_rup_pagu = 0;
        total_non_pengadaan = 0;
        total_target_pengadaan = 0;

        const font_sub = { name: FONT, size: 9, bold: true, color: { argb: DARK_BLUE } };
        const font_det = { name: FONT, size: 9, color: { argb: INK } };

        for (const prog of rkaData) {
            const prog_code = prog.text.split("]")[0].replace("[", "").trim();
            
            let prog_rka = 0;
            let prog_np = 0;
            let prog_target = 0;
            let prog_rup = 0;
            
            for (const keg of prog.kegiatans || []) {
                let keg_rka = 0;
                let keg_np = 0;
                let keg_target = 0;
                let keg_rup = 0;
                
                for (const out of keg.outputs || []) {
                    let out_rka = 0;
                    let out_np = 0;
                    let out_target = 0;
                    let out_rup = 0;
                    
                    const suboutputs = out.suboutputs || [{ id: 0, name: out.name, code: "000", komponens: out.komponens || [] }];
                    
                    for (const ro of suboutputs) {
                        let ro_rka = 0;
                        let ro_np = 0;
                        let ro_target = 0;
                        let ro_rup = 0;
                        
                        for (const komp of ro.komponens || []) {
                            const comp_key = `${prog_code}.${keg.code}.${out.code}.${ro.code}.${komp.code}`;
                            
                            const rup_sum = rup_all_lines
                                .filter(line => line.comp_key === comp_key && line.is_terumumkan)
                                .reduce((sum, line) => sum + line.pagu, 0);
                            
                            // Hierarchical propagation of NP/Gaji flags
                            let komp_np = false, komp_gj = false;
                            let subkomp_np = false, subkomp_gj = false;
                            let akun_np = false, akun_gj = false;
                            
                            let np_gaji_sum = 0;
                            for (const r of komp.rows || []) {
                                if (r.level === 0) {
                                    komp_np = r.np_ch;
                                    komp_gj = r.gj_ch;
                                } else if (r.level === 1) {
                                    subkomp_np = r.np_ch;
                                    subkomp_gj = r.gj_ch;
                                } else if (r.level === 2) {
                                    akun_np = r.np_ch;
                                    akun_gj = r.gj_ch;
                                } else if (r.level === 3) {
                                    const is_np = r.np_ch || akun_np || subkomp_np || komp_np;
                                    const is_gj = r.gj_ch || akun_gj || subkomp_gj || komp_gj;
                                    if (is_np || is_gj) {
                                        np_gaji_sum += (typeof r.pagu === 'number' ? r.pagu : 0);
                                    }
                                }
                            }
                            
                            const target_pengadaan = komp.pagu - np_gaji_sum;
                            
                            // Write detail Komponen row
                            writeSandingRow(
                                row_idx, num, prog_code, keg.name, out.name, ro.name, komp.name, comp_key,
                                komp.pagu, np_gaji_sum, target_pengadaan, rup_sum,
                                null, font_det
                            );
                            
                            ro_rka += komp.pagu;
                            ro_np += np_gaji_sum;
                            ro_target += target_pengadaan;
                            ro_rup += rup_sum;
                            
                            row_idx++;
                            num++;
                        }
                        
                        // Write Subtotal RO (Level 3 - Softest Gray)
                        writeSandingRow(
                            row_idx, "", prog_code, "", "", "", `SUBTOTAL RO: ${ro.name}`, `${prog_code}.${keg.code}.${out.code}.${ro.code}`,
                            ro_rka, ro_np, ro_target, ro_rup,
                            LIGHT_GRAY_LVL3, font_sub
                        );
                        
                        out_rka += ro_rka;
                        out_np += ro_np;
                        out_target += ro_target;
                        out_rup += ro_rup;
                        
                        row_idx++;
                    }
                    
                    // Write Subtotal KRO (Level 2 - Light Gray F2F2F2)
                    writeSandingRow(
                        row_idx, "", prog_code, "", "", "", `SUBTOTAL KRO: ${out.name}`, `${prog_code}.${keg.code}.${out.code}`,
                        out_rka, out_np, out_target, out_rup,
                        LIGHT_GRAY_LVL2, font_sub
                    );
                    
                    keg_rka += out_rka;
                    keg_np += out_np;
                    keg_target += out_target;
                    keg_rup += out_rup;
                    
                    row_idx++;
                }
                
                // Write Subtotal Kegiatan (Level 1 - Light Blue DCE6F1)
                writeSandingRow(
                    row_idx, "", prog_code, "", "", "", `SUBTOTAL KEGIATAN: ${keg.name}`, `${prog_code}.${keg.code}`,
                    keg_rka, keg_np, keg_target, keg_rup,
                    LIGHT_BLUE_LVL1, font_sub
                );
                
                prog_rka += keg_rka;
                prog_np += keg_np;
                prog_target += keg_target;
                prog_rup += keg_rup;
                
                row_idx++;
            }
            
            // Write Subtotal Program (Level 0 - Light Blue B8CCE4)
            writeSandingRow(
                row_idx, "", prog_code, "", "", "", `SUBTOTAL PROGRAM: ${prog.text.split(']')[1].trim()}`, `${prog_code}`,
                prog_rka, prog_np, prog_target, prog_rup,
                LIGHT_BLUE_LVL0, font_sub
            );
            
            total_non_pengadaan += prog_np;
            total_target_pengadaan += prog_target;
            total_rup_pagu += prog_rup;
            
            row_idx++;
        }

        // Baris TOTAL
        tot_pct = total_target_pengadaan > 0 ? (total_rup_pagu / total_target_pengadaan) : 0;
        const sanding_last_row = row_idx;

        for (let c = 1; c <= 14; c++) {
            const cell = ws_sanding.getCell(row_idx, c);
            cell.fill = solid(DARK_BLUE);
            cell.border = border_thin;
            cell.font = { name: FONT, size: 10, bold: true, color: { argb: WHITE_COLOR } };
            cell.alignment = { horizontal: 'right', vertical: 'middle', indent: 1 };
        }
        ws_sanding.getCell(row_idx, 7).value = "TOTAL SATKER";
        ws_sanding.getCell(row_idx, 8).value = total_satker_pagu;
        ws_sanding.getCell(row_idx, 8).numFmt = FMT_RP;
        ws_sanding.getCell(row_idx, 9).value = total_non_pengadaan;
        ws_sanding.getCell(row_idx, 9).numFmt = FMT_RP;
        ws_sanding.getCell(row_idx, 10).value = total_target_pengadaan;
        ws_sanding.getCell(row_idx, 10).numFmt = FMT_RP;
        ws_sanding.getCell(row_idx, 11).value = total_rup_pagu;
        ws_sanding.getCell(row_idx, 11).numFmt = FMT_RP;
        ws_sanding.getCell(row_idx, 12).value = total_target_pengadaan - total_rup_pagu;
        ws_sanding.getCell(row_idx, 12).numFmt = FMT_RP;
        ws_sanding.getCell(row_idx, 13).value = tot_pct;
        ws_sanding.getCell(row_idx, 13).numFmt = FMT_PCT;
        ws_sanding.getRow(row_idx).height = 22;

        ws_sanding.getCell('A3').value =
            `Pagu RKA ${rp(total_satker_pagu)}  •  Target pengadaan ${rp(total_target_pengadaan)}  •  ` +
            `RUP terumumkan ${rp(total_rup_pagu)} (${(tot_pct * 100).toFixed(1)}%)  •  ` +
            `Sisa belum diumumkan ${rp(total_target_pengadaan - total_rup_pagu)}`;
        ws_sanding.getCell('A3').font = { name: FONT, size: 9.5, bold: true, color: { argb: DARK_BLUE } };
        ws_sanding.getCell('A3').alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
        ws_sanding.getCell('A3').fill = solid(LIGHT_BLUE_LVL1);
        ws_sanding.mergeCells("A3:N3");
        ws_sanding.getRow(3).height = 20;

        finishSheet(ws_sanding, {
            headerRow: 4, lastRow: sanding_last_row - 1, lastCol: 14, tabColor: DARK_BLUE
        });


        // ----------------- SHEET 3: DAFTAR PAKET RUP -----------------
        const ws_rup = wb.addWorksheet("Daftar Paket RUP");

        writeTitle(ws_rup, "DAFTAR PAKET PENYEDIA DI RUP", "M");

        const headers_rup = [
            "No", "ID Paket", "Nama Kegiatan (RUP)", "Nama Paket", "Pagu RUP",
            "Waktu Pemilihan", "Sumber Dana", "A", "FD", "U", "Status Paket",
            "Kode Otorisasi (MAK)", "Komponen RKA Tersanding"
        ];

        writeHeader(ws_rup, 4, headers_rup,
            [6, 12, 28, 40, 16, 16, 13, 6, 6, 6, 20, 30, 26]);
        ws_rup.getCell(4, 8).note = "A = Draft PPK";
        ws_rup.getCell(4, 9).note = "FD = Final Draft PPK";
        ws_rup.getCell(4, 10).note = "U = Sudah diumumkan KPA";

        row_idx = 5;
        for (let idx = 0; idx < rupPackets.length; idx++) {
            const p = rupPackets[idx];
            const sah = p.aktif && p.fd && p.umumkan;

            ws_rup.getCell(row_idx, 1).value = idx + 1;
            ws_rup.getCell(row_idx, 2).value = p.id;
            ws_rup.getCell(row_idx, 3).value = p.keg_name;
            ws_rup.getCell(row_idx, 4).value = p.name;

            ws_rup.getCell(row_idx, 5).value = p.pagu;
            ws_rup.getCell(row_idx, 5).numFmt = FMT_RP;

            ws_rup.getCell(row_idx, 6).value = p.waktu;
            ws_rup.getCell(row_idx, 7).value = p.sumber_dana;
            ws_rup.getCell(row_idx, 8).value = p.aktif ? "✓" : "–";
            ws_rup.getCell(row_idx, 9).value = p.fd ? "✓" : "–";
            ws_rup.getCell(row_idx, 10).value = p.umumkan ? "✓" : "–";
            ws_rup.getCell(row_idx, 11).value = sah ? "Terumumkan" : "Draft / Batal";
            ws_rup.getCell(row_idx, 12).value = p.mak;

            const p_comp_keys = Array.from(new Set(rup_all_lines.filter(l => l.packet_id === p.id).map(l => l.comp_key)));
            ws_rup.getCell(row_idx, 13).value = p_comp_keys.join(", ");

            for (let col_c = 1; col_c <= 13; col_c++) {
                const cell = ws_rup.getCell(row_idx, col_c);
                cell.border = border_thin;
                cell.font = { name: FONT, size: 9, color: { argb: sah ? INK : INK_SOFT } };
                if (col_c === 5) {
                    cell.alignment = { horizontal: 'right', vertical: 'middle', indent: 1 };
                } else if (col_c === 3 || col_c === 4) {
                    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
                } else {
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                }
            }

            const st = ws_rup.getCell(row_idx, 11);
            st.fill = solid(sah ? LIGHT_GREEN : LIGHT_ORANGE);
            st.font = { name: FONT, size: 9, bold: true, color: { argb: sah ? 'FF166534' : 'FFB45309' } };

            row_idx++;
        }

        zebra(ws_rup, 5, row_idx - 1, 13);

        const rup_last_row = row_idx;
        for (let c = 1; c <= 13; c++) {
            const cell = ws_rup.getCell(row_idx, c);
            cell.fill = solid(LIGHT_BLUE_LVL0);
            cell.border = border_thin;
            cell.font = { name: FONT, size: 10, bold: true, color: { argb: DARK_BLUE } };
            cell.alignment = { horizontal: 'right', vertical: 'middle', indent: 1 };
        }
        ws_rup.getCell(row_idx, 4).value = "TOTAL PAGU RUP TERUMUMKAN";
        ws_rup.getCell(row_idx, 5).value = total_rup_pagu;
        ws_rup.getCell(row_idx, 5).numFmt = FMT_RP;
        ws_rup.getRow(row_idx).height = 20;

        finishSheet(ws_rup, { headerRow: 4, lastRow: rup_last_row - 1, lastCol: 13, tabColor: 'FF375623' });


        // ----------------- SHEET 3A: PAKET RUP TANPA SANDINGAN -----------------
        const ws_no_sanding = wb.addWorksheet("Paket RUP Tanpa Sandingan");

        writeTitle(ws_no_sanding, "PAKET RUP TANPA SANDINGAN DI RKA", "F");

        const headers_no_sanding = [
            "No", "ID Paket RUP", "Nama Paket RUP", "Kode MAK di RUP", "Pagu Paket", "Keterangan / Alasan"
        ];

        writeHeader(ws_no_sanding, 4, headers_no_sanding, [6, 15, 40, 30, 18, 48]);


        let no_sanding_row = 5;
        for (let idx = 0; idx < unique_unmatched.length; idx++) {
            const item = unique_unmatched[idx];
            ws_no_sanding.getCell(no_sanding_row, 1).value = idx + 1;
            ws_no_sanding.getCell(no_sanding_row, 2).value = item.packet_id;
            ws_no_sanding.getCell(no_sanding_row, 3).value = item.packet_name;
            ws_no_sanding.getCell(no_sanding_row, 4).value = item.mak;
            
            ws_no_sanding.getCell(no_sanding_row, 5).value = item.pagu;
            ws_no_sanding.getCell(no_sanding_row, 5).numFmt = FMT_RP;

            ws_no_sanding.getCell(no_sanding_row, 6).value = item.reason;

            for (let col_c = 1; col_c <= 6; col_c++) {
                const cell = ws_no_sanding.getCell(no_sanding_row, col_c);
                cell.border = border_thin;
                cell.font = { name: FONT, size: 9, color: { argb: INK } };
                cell.fill = solid(LIGHT_RED);

                if (col_c === 1 || col_c === 2 || col_c === 4) {
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                } else if (col_c === 5) {
                    cell.alignment = { horizontal: 'right', vertical: 'middle', indent: 1 };
                } else {
                    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
                }
            }
            no_sanding_row++;
        }

        if (unique_unmatched.length === 0) {
            ws_no_sanding.mergeCells("A5:F5");
            const ok = ws_no_sanding.getCell('A5');
            ok.value = "Bagus — semua paket RUP punya sandingan MAK yang valid di RKA.";
            ok.font = { name: FONT, size: 10, bold: true, color: { argb: 'FF166534' } };
            ok.fill = solid(LIGHT_GREEN);
            ok.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
            ws_no_sanding.getRow(5).height = 22;
            for (let c = 1; c <= 6; c++) ws_no_sanding.getCell(5, c).border = border_thin;
            no_sanding_row = 6;
        }

        finishSheet(ws_no_sanding, {
            headerRow: 4, lastRow: no_sanding_row - 1, lastCol: 6,
            tabColor: unique_unmatched.length ? 'FFC00000' : 'FF375623'
        });


        // ----------------- SHEET 4 & 5: DETAIL PER PROGRAM -----------------
        for (const prog of rkaData) {
            const prog_code = prog.text.split("]")[0].replace("[", "").trim();
            const sheet_name = `Detail - ${prog_code}`;
            const ws = wb.addWorksheet(sheet_name);

            writeTitle(ws, `DETAIL RKA & SANDINGAN RUP — PROGRAM ${prog_code}`, "S");

            ws.mergeCells("A3:S3");
            ws.getCell('A3').value = prog.text;
            ws.getCell('A3').font = { name: FONT, size: 10, bold: true, color: { argb: DARK_BLUE } };
            ws.getCell('A3').fill = solid(LIGHT_BLUE_LVL1);
            ws.getCell('A3').alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
            ws.getRow(3).height = 20;

            const headers_det = [
                "Kegiatan", "KRO (Output)", "RO (Sub-Output)", "Komponen", "Kode",
                "Uraian", "Uraian Sebelum Revisi", "Pagu RKA", "Pagu Sebelum Revisi",
                "P", "S", "MY", "NP", "Gaji",
                "ID Paket RUP", "Nama Paket RUP", "Pagu RUP Terumumkan", "Selisih Pengadaan", "Rencana Pemilihan"
            ];

            writeHeader(ws, 5, headers_det,
                [26, 26, 26, 26, 16, 46, 46, 16, 16, 5, 5, 5, 5, 5, 16, 32, 20, 20, 18]);
            ws.getCell(5, 10).note = "P = Pengadaan";
            ws.getCell(5, 11).note = "S = Swakelola";
            ws.getCell(5, 12).note = "MY = Multiyears";
            ws.getCell(5, 13).note = "NP = Non-Pengadaan";

            let curr_row = 6;
            const matched_rup_indices = new Set();
            
            for (const keg of prog.kegiatans || []) {
                const keg_text = `[${keg.code}] ${keg.name}`;
                for (const out of keg.outputs || []) {
                    const kro_text = `[${out.code}] ${out.name}`;
                    
                    const suboutputs = out.suboutputs || [{ id: 0, name: out.name, code: "000", komponens: out.komponens || [] }];
                    
                    for (const ro of suboutputs) {
                        const ro_text = `[${ro.code}] ${ro.name}`;
                        
                        for (const komp of ro.komponens || []) {
                            const komp_text = `[${komp.code}] ${komp.name}`;
                            
                            const comp_key = `${prog_code}.${keg.code}.${out.code}.${ro.code}.${komp.code}`;
                            const comp_rup_lines = rup_all_lines.filter(l => l.comp_key === comp_key);
                            const komp_rows = komp.rows || [];
                            
                            const start_row = curr_row;
                            
                            // --- PRE-CALCULATE target_pagu FOR THIS COMPONENT WITH PROPAGATION ---
                            const rows_target = {};
                            const rows_is_np = {};
                            const rows_is_gj = {};
                            
                            let komp_np = false, komp_gj = false;
                            let subkomp_np = false, subkomp_gj = false;
                            let akun_np = false, akun_gj = false;
                            
                            for (let r_idx = 0; r_idx < komp_rows.length; r_idx++) {
                                const r = komp_rows[r_idx];
                                const pagu = typeof r.pagu === 'number' ? r.pagu : 0;
                                
                                if (r.level === 0) {
                                    komp_np = r.np_ch;
                                    komp_gj = r.gj_ch;
                                } else if (r.level === 1) {
                                    subkomp_np = r.np_ch;
                                    subkomp_gj = r.gj_ch;
                                } else if (r.level === 2) {
                                    akun_np = r.np_ch;
                                    akun_gj = r.gj_ch;
                                } else if (r.level === 3) {
                                    const is_np = r.np_ch || akun_np || subkomp_np || komp_np;
                                    const is_gj = r.gj_ch || akun_gj || subkomp_gj || komp_gj;
                                    rows_is_np[r_idx] = is_np;
                                    rows_is_gj[r_idx] = is_gj;
                                    rows_target[r_idx] = (is_np || is_gj) ? 0 : pagu;
                                }
                            }
                            
                            // Calculate level 2 (Akun) target pagus
                            let curr_akun_idx = null;
                            for (let r_idx = 0; r_idx < komp_rows.length; r_idx++) {
                                const r = komp_rows[r_idx];
                                if (r.level === 2) {
                                    curr_akun_idx = r_idx;
                                    rows_target[r_idx] = 0;
                                } else if (r.level === 3 && curr_akun_idx !== null) {
                                    rows_target[curr_akun_idx] += rows_target[r_idx];
                                }
                            }
                            
                            // Calculate level 1 (Sub-komponen) target pagus
                            let curr_sub_idx = null;
                            for (let r_idx = 0; r_idx < komp_rows.length; r_idx++) {
                                const r = komp_rows[r_idx];
                                if (r.level === 1) {
                                    curr_sub_idx = r_idx;
                                    rows_target[r_idx] = 0;
                                } else if (r.level === 3 && curr_sub_idx !== null) {
                                    rows_target[curr_sub_idx] += rows_target[r_idx];
                                }
                            }
                            
                            // Calculate level 0 (Komponen) target pagus
                            let curr_komp_idx = null;
                            for (let r_idx = 0; r_idx < komp_rows.length; r_idx++) {
                                const r = komp_rows[r_idx];
                                if (r.level === 0) {
                                    curr_komp_idx = r_idx;
                                    rows_target[r_idx] = 0;
                                } else if (r.level === 3 && curr_komp_idx !== null) {
                                    rows_target[curr_komp_idx] += rows_target[r_idx];
                                }
                            }
                            
                            // --- DETAILED MATCHING FOR LEVEL 3 DETAIL ROWS (PROPORTIONAL ALLOCATION) ---
                            const detail_row_objects = [];
                            let curr_subkomp = "";
                            let curr_akun = "";
                            for (let r_idx = 0; r_idx < komp_rows.length; r_idx++) {
                                const r = komp_rows[r_idx];
                                if (r.level === 1) {
                                    curr_subkomp = r.code;
                                } else if (r.level === 2) {
                                    curr_akun = r.code;
                                } else if (r.level === 3) {
                                    const is_np = rows_is_np[r_idx] || false;
                                    const is_gj = rows_is_gj[r_idx] || false;
                                    if (!is_np && !is_gj) {
                                        detail_row_objects.push({
                                            r_idx: r_idx,
                                            pagu: typeof r.pagu === 'number' ? r.pagu : 0,
                                            key: `${prog_code}.${keg.code}.${out.code}.${ro.code}.${komp.code}.${curr_subkomp}.${curr_akun}`,
                                            allocated_pagu: 0,
                                            matched_ids: [],
                                            matched_names: [],
                                            is_draft_match: false
                                        });
                                    }
                                }
                            }
                            
                            // Group RUP lines by their Akun key and separate into Terumumkan and Draft pools
                            const rup_pools_terumumkan = [];
                            const rup_pools_draft = [];
                            for (const line of comp_rup_lines) {
                                const item = {
                                    packet_id: line.packet_id,
                                    packet_name: line.packet_name,
                                    pagu: line.pagu,
                                    remaining_pagu: line.pagu,
                                    key: line.key
                                };
                                if (line.is_terumumkan) {
                                    rup_pools_terumumkan.push(item);
                                } else {
                                    rup_pools_draft.push(item);
                                }
                            }
                            
                            // Perform Proportional Allocation for each unique Akun key
                            const unique_keys = new Set(detail_row_objects.map(d => d.key));
                            for (const k of unique_keys) {
                                const k_rka_rows = detail_row_objects.filter(d => d.key === k);
                                const k_rup_pools_ter = rup_pools_terumumkan.filter(p => p.key === k);
                                const k_rup_pools_drf = rup_pools_draft.filter(p => p.key === k);
                                
                                // 1. First Pass: Allocate Terumumkan RUP
                                for (const d_obj of k_rka_rows) {
                                    let needed = d_obj.pagu;
                                    for (const pool of k_rup_pools_ter) {
                                        if (pool.remaining_pagu > 0 && needed > 0) {
                                            const amount = Math.min(needed, pool.remaining_pagu);
                                            pool.remaining_pagu -= amount;
                                            d_obj.allocated_pagu += amount;
                                            needed -= amount;
                                            
                                            if (!d_obj.matched_ids.includes(pool.packet_id)) {
                                                d_obj.matched_ids.push(pool.packet_id);
                                            }
                                            if (!d_obj.matched_names.includes(pool.packet_name)) {
                                                d_obj.matched_names.push(pool.packet_name);
                                            }
                                        }
                                    }
                                }
                                
                                // Leftover Terumumkan RUP goes to the first RKA row
                                const leftover = k_rup_pools_ter.reduce((sum, p) => sum + p.remaining_pagu, 0);
                                if (leftover > 0 && k_rka_rows.length > 0) {
                                    const first_row = k_rka_rows[0];
                                    first_row.allocated_pagu += leftover;
                                    for (const pool of k_rup_pools_ter) {
                                        if (pool.remaining_pagu > 0) {
                                            if (!first_row.matched_ids.includes(pool.packet_id)) {
                                                first_row.matched_ids.push(pool.packet_id);
                                            }
                                            if (!first_row.matched_names.includes(pool.packet_name)) {
                                                first_row.matched_names.push(pool.packet_name);
                                            }
                                            pool.remaining_pagu = 0;
                                        }
                                    }
                                }
                                
                                // 2. Second Pass: Link remaining unmatched RKA rows to Draft RUPs (warnings)
                                for (const d_obj of k_rka_rows) {
                                    if (d_obj.allocated_pagu < d_obj.pagu) {
                                        for (const pool of k_rup_pools_drf) {
                                            if (pool.remaining_pagu > 0) {
                                                d_obj.is_draft_match = true;
                                                const drf_id = `[DRAFT/BATAL] ${pool.packet_id}`;
                                                const drf_name = `[Draft/Batal] ${pool.packet_name}`;
                                                
                                                if (!d_obj.matched_ids.includes(drf_id)) {
                                                    d_obj.matched_ids.push(drf_id);
                                                }
                                                if (!d_obj.matched_names.includes(drf_name)) {
                                                    d_obj.matched_names.push(drf_name);
                                                }
                                                pool.remaining_pagu = 0; // mark as matched
                                            }
                                        }
                                    }
                                }
                            }

                            // --- ITERATE AND WRITE ROWS ---
                            let subkomp_code = "";
                            let akun_code = "";
                            
                            for (let row_idx_in_comp = 0; row_idx_in_comp < komp_rows.length; row_idx_in_comp++) {
                                const row_data = komp_rows[row_idx_in_comp];
                                const level = row_data.level;
                                const code = row_data.code;
                                const desc = row_data.desc;
                                const desc_prev = row_data.descPrev;
                                const pagu = row_data.pagu;
                                const pagu_prev = row_data.paguPrev;
                                
                                const p_ch = row_data.p_ch;
                                const s_ch = row_data.s_ch;
                                const my_ch = row_data.my_ch;
                                const np_ch = row_data.np_ch;
                                const gj_ch = row_data.gj_ch;
                                
                                if (level === 1) subkomp_code = code;
                                else if (level === 2) akun_code = code;
                                
                                ws.getCell(curr_row, 5).value = code;
                                ws.getCell(curr_row, 6).value = desc;
                                ws.getCell(curr_row, 7).value = desc_prev;
                                
                                if (typeof pagu === 'number') {
                                    ws.getCell(curr_row, 8).value = pagu;
                                    ws.getCell(curr_row, 8).numFmt ='#,##0';
                                } else {
                                    ws.getCell(curr_row, 8).value = pagu;
                                }
                                
                                if (typeof pagu_prev === 'number') {
                                    ws.getCell(curr_row, 9).value = pagu_prev;
                                    ws.getCell(curr_row, 9).numFmt ='#,##0';
                                } else {
                                    ws.getCell(curr_row, 9).value = pagu_prev;
                                }
                                
                                ws.getCell(curr_row, 10).value = p_ch ? "✓" : "";
                                ws.getCell(curr_row, 11).value = s_ch ? "✓" : "";
                                ws.getCell(curr_row, 12).value = my_ch ? "✓" : "";
                                ws.getCell(curr_row, 13).value = np_ch ? "✓" : "";
                                ws.getCell(curr_row, 14).value = gj_ch ? "✓" : "";
                                
                                let matched_pkt_id = "";
                                let matched_pkt_name = "";
                                let rup_pagu_val = 0;
                                let is_np_gaji = false;
                                let is_draft = false;
                                let matched_pkt_waktu = "";
                                
                                if (level === 0) {
                                    rup_pagu_val = comp_rup_lines.filter(l => l.is_terumumkan).reduce((sum, l) => sum + l.pagu, 0);
                                } else if (level === 1) {
                                    rup_pagu_val = comp_rup_lines.filter(l => l.subkomp === subkomp_code && l.is_terumumkan).reduce((sum, l) => sum + l.pagu, 0);
                                } else if (level === 2) {
                                    const akun_key = `${prog_code}.${keg.code}.${out.code}.${ro.code}.${komp.code}.${subkomp_code}.${code}`;
                                    rup_pagu_val = comp_rup_lines.filter(l => l.key === akun_key && l.is_terumumkan).reduce((sum, l) => sum + l.pagu, 0);
                                } else if (level === 3) {
                                    const is_np = rows_is_np[row_idx_in_comp] || false;
                                    const is_gj = rows_is_gj[row_idx_in_comp] || false;
                                    is_np_gaji = is_np || is_gj;
                                    
                                    if (is_np) {
                                        matched_pkt_id = "NP";
                                        matched_pkt_name = "Non-Pengadaan (Honor/Perdin/Uang Makan/PJLP/PPPK Paruh Waktu)";
                                    } else if (is_gj) {
                                        matched_pkt_id = "Gaji";
                                        matched_pkt_name = "Gaji & Tunjangan Pegawai";
                                    } else {
                                        const d_obj = detail_row_objects.find(d => d.r_idx === row_idx_in_comp);
                                        if (d_obj && d_obj.matched_ids && d_obj.matched_ids.length > 0) {
                                            matched_pkt_id = d_obj.matched_ids.join(", ");
                                            matched_pkt_name = d_obj.matched_names.join(", ");
                                            rup_pagu_val = d_obj.allocated_pagu;
                                            is_draft = d_obj.is_draft_match;
                                            
                                            const matched_times = [];
                                            for (const id of d_obj.matched_ids) {
                                                const clean_id = id.toString().replace(/\[DRAFT\/BATAL\] /g, "").trim();
                                                const pkt = rupPackets.find(p => p.id == clean_id);
                                                if (pkt && pkt.waktu && !matched_times.includes(pkt.waktu)) {
                                                    matched_times.push(pkt.waktu);
                                                }
                                            }
                                            matched_pkt_waktu = matched_times.join(", ");
                                        }
                                    }
                                }
                                
                                ws.getCell(curr_row, 15).value = matched_pkt_id;
                                ws.getCell(curr_row, 16).value = matched_pkt_name;
                                ws.getCell(curr_row, 19).value = matched_pkt_waktu;
                                ws.getCell(curr_row, 19).alignment = { horizontal: 'center', vertical: 'middle' };
                                
                                const target_pagu = rows_target[row_idx_in_comp] || 0;
                                
                                if (is_np_gaji) {
                                    ws.getCell(curr_row, 17).value = 0;
                                    ws.getCell(curr_row, 17).numFmt ='#,##0';
                                    ws.getCell(curr_row, 17).alignment = { horizontal: 'right', vertical: 'middle' };
                                    ws.getCell(curr_row, 18).value = 0;
                                    ws.getCell(curr_row, 18).numFmt ='#,##0';
                                    ws.getCell(curr_row, 18).alignment = { horizontal: 'right', vertical: 'middle' };
                                } else {
                                    ws.getCell(curr_row, 17).value = rup_pagu_val;
                                    ws.getCell(curr_row, 17).numFmt ='#,##0';
                                    ws.getCell(curr_row, 17).alignment = { horizontal: 'right', vertical: 'middle' };
                                    
                                    ws.getCell(curr_row, 18).value = target_pagu - rup_pagu_val;
                                    ws.getCell(curr_row, 18).numFmt ='#,##0';
                                    ws.getCell(curr_row, 18).alignment = { horizontal: 'right', vertical: 'middle' };
                                }
                                
                                // Styles
                                let row_fill = WHITE_COLOR;
                                let row_font_bold = false;
                                let row_font_italic = false;
                                
                                if (level === 0) {
                                    row_fill = LIGHT_BLUE_LVL0;
                                    row_font_bold = true;
                                } else if (level === 1) {
                                    row_fill = LIGHT_BLUE_LVL1;
                                    row_font_bold = true;
                                } else if (level === 2) {
                                    row_fill = LIGHT_GRAY_LVL2;
                                } else {
                                    row_font_italic = true;
                                }
                                
                                for (let col_c = 5; col_c <= 19; col_c++) {
                                    const cell = ws.getCell(curr_row, col_c);
                                    
                                    let cell_fill = row_fill;
                                    let cell_italic = row_font_italic || (level === 3 && matched_pkt_id !== "");
                                    let font_color = 'FF000000';
                                    
                                    if (level === 3) {
                                        if (is_draft && (col_c === 15 || col_c === 16 || col_c === 17 || col_c === 19)) {
                                            cell_fill = 'FFFFF2CC'; // light orange background
                                            font_color = 'FFC95D00'; // dark orange text
                                            cell_italic = true;
                                        } else if (matched_pkt_id === "NP" || matched_pkt_id === "Gaji") {
                                            font_color = 'FF7F7F7F'; // soft gray
                                        } else {
                                            font_color = matched_pkt_id ? 'FF2E7D32' : 'FF595959';
                                        }
                                    }
                                    
                                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cell_fill } };
                                    cell.border = border_thin;
                                    
                                    cell.font = {
                                        name: "Segoe UI",
                                        size: level === 3 ? 9 : 10,
                                        bold: row_font_bold,
                                        italic: cell_italic,
                                        color: { argb: font_color }
                                    };
                                    
                                    if (col_c === 5 || (col_c >= 10 && col_c <= 15)) {
                                        cell.alignment = { horizontal: 'center', vertical: 'middle' };
                                    } else if (col_c === 6 || col_c === 7 || col_c === 16) {
                                        cell.alignment = { horizontal: 'left', vertical: 'middle', indent: level * 2, wrapText: true };
                                    } else if (col_c === 17 || col_c === 18) {
                                        // already set
                                    } else {
                                        cell.alignment = { horizontal: 'right', vertical: 'middle' };
                                    }
                                }
                                curr_row++;
                            }
                            
                            const end_row = curr_row - 1;
                            if (end_row >= start_row) {
                                ws.getCell(start_row, 1).value = keg_text;
                                ws.getCell(start_row, 2).value = kro_text;
                                ws.getCell(start_row, 3).value = ro_text;
                                ws.getCell(start_row, 4).value = komp_text;
                                
                                for (let r = start_row; r <= end_row; r++) {
                                    for (let c = 1; c <= 4; c++) {
                                        const cell = ws.getCell(r, c);
                                        cell.font = { name: FONT, size: 9, color: { argb: 'FF333333' } };
                                        cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
                                        cell.border = border_thin;
                                        cell.fill = solid('FFFAFBFD');
                                    }
                                }
                                
                                if (end_row > start_row) {
                                    ws.mergeCells(start_row, 1, end_row, 1);
                                    ws.mergeCells(start_row, 2, end_row, 2);
                                    ws.mergeCells(start_row, 3, end_row, 3);
                                    ws.mergeCells(start_row, 4, end_row, 4);
                                }
                            }
                            
                            curr_row += 2;
                        }
                    }
                }
            }
            
            // Kolom A–D (Kegiatan/KRO/RO/Komponen) ikut dibekukan supaya konteks
            // baris tidak hilang saat menggeser ke kanan. Autofilter mulai kolom E
            // karena A–D berisi sel gabungan.
            finishSheet(ws, {
                headerRow: 5, freezeCols: 4, filterFromCol: 5,
                lastRow: curr_row - 1, lastCol: 19, tabColor: MID_BLUE
            });
        }

        // Save file
        const safeSatker = (ctx.satkerName || 'Satker')
            .replace(/[\\/:*?"<>|]/g, '')
            .replace(/\s+/g, '_')
            .slice(0, 60);
        const d = new Date();
        const dateTag = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
        const fileName = `Sanding_RKA_RUP_${safeSatker}_TA${ctx.tahun}_${dateTag}.xlsx`;

        const buffer = await wb.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        saveAs(blob, fileName);

        return {
            fileName,
            packets: rupPackets.length,
            pct: tot_pct,
            unmatched: unique_unmatched_count
        };
    }

    // ═════════════════════════════════════════════════════════════════ INIT ══
    // SiRUP memuat sebagian halaman lewat AJAX, jadi tombol dipasang ulang
    // kalau body sempat di-render ulang.
    function init() {
        injectStyles();
        injectButton();

        const obs = new MutationObserver(() => {
            if (!document.getElementById('srx-fab')) injectButton();
        });
        obs.observe(document.body, { childList: true, subtree: false });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => setTimeout(init, 800));
    } else {
        setTimeout(init, 800);
    }
})();
