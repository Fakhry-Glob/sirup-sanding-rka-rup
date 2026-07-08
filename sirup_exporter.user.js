// ==UserScript==
// @name         SiRUP RKA & RUP Exporter & Sander
// @namespace    http://tampermonkey.net/
// @version      1.9
// @description  Crawl RKA and RUP details from SiRUP and export to a beautifully formatted Excel comparison report with unique RO matching, hierarchical subtotals, inherited NP/Gaji flag propagation, and high-performance concurrency.
// @author       Antigravity
// @match        https://sirup.inaproc.id/sirup/*
// @require      https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js
// @updateURL    https://raw.githubusercontent.com/Fakhry-Glob/sirup-sanding-rka-rup/master/sirup_exporter.user.js
// @downloadURL  https://raw.githubusercontent.com/Fakhry-Glob/sirup-sanding-rka-rup/master/sirup_exporter.user.js
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // Inject export button into the page
    function injectButton() {
        if (document.getElementById('btn-export-sanding')) return;

        const btn = document.createElement('button');
        btn.id = 'btn-export-sanding';
        btn.className = 'btn btn-primary';
        btn.style.position = 'fixed';
        btn.style.bottom = '20px';
        btn.style.right = '20px';
        btn.style.zIndex = '9999';
        btn.style.boxShadow = '0 4px 10px rgba(0,0,0,0.2)';
        btn.style.borderRadius = '50px';
        btn.style.padding = '12px 24px';
        btn.style.fontSize = '14px';
        btn.style.fontWeight = 'bold';
        btn.style.border = 'none';
        btn.style.backgroundColor = '#1F497D';
        btn.style.color = '#ffffff';
        btn.style.cursor = 'pointer';
        btn.innerHTML = '📊 Ekspor Sanding RKA & RUP';
        
        btn.addEventListener('click', startExport);
        document.body.appendChild(btn);
    }

    // Modal/Progress Logger UI
    let logModal, logContent, progressBar;
    function showLogger() {
        if (!logModal) {
            logModal = document.createElement('div');
            logModal.style.position = 'fixed';
            logModal.style.top = '50%';
            logModal.style.left = '50%';
            logModal.style.transform = 'translate(-50%, -50%)';
            logModal.style.width = '550px';
            logModal.style.maxHeight = '450px';
            logModal.style.backgroundColor = '#ffffff';
            logModal.style.border = '1px solid #ccc';
            logModal.style.boxShadow = '0 10px 25px rgba(0,0,0,0.3)';
            logModal.style.borderRadius = '8px';
            logModal.style.padding = '20px';
            logModal.style.zIndex = '10000';
            logModal.style.fontFamily = '"Segoe UI", Tahoma, Geneva, Verdana, sans-serif';

            const header = document.createElement('h4');
            header.style.margin = '0 0 15px 0';
            header.style.color = '#1F497D';
            header.innerText = 'Proses Sinkronisasi & Ekspor RKA-RUP';
            logModal.appendChild(header);

            const progressContainer = document.createElement('div');
            progressContainer.style.width = '100%';
            progressContainer.style.backgroundColor = '#f1f1f1';
            progressContainer.style.borderRadius = '4px';
            progressContainer.style.marginBottom = '15px';
            
            progressBar = document.createElement('div');
            progressBar.style.width = '0%';
            progressBar.style.height = '15px';
            progressBar.style.backgroundColor = '#1F497D';
            progressBar.style.borderRadius = '4px';
            progressBar.style.transition = 'width 0.3s ease';
            progressContainer.appendChild(progressBar);
            logModal.appendChild(progressContainer);

            logContent = document.createElement('div');
            logContent.style.height = '220px';
            logContent.style.overflowY = 'auto';
            logContent.style.fontSize = '12px';
            logContent.style.border = '1px solid #eee';
            logContent.style.padding = '10px';
            logContent.style.backgroundColor = '#fafafa';
            logContent.style.lineHeight = '1.5';
            logModal.appendChild(logContent);

            const backdrop = document.createElement('div');
            backdrop.id = 'export-backdrop';
            backdrop.style.position = 'fixed';
            backdrop.style.top = '0';
            backdrop.style.left = '0';
            backdrop.style.width = '100%';
            backdrop.style.height = '100%';
            backdrop.style.backgroundColor = 'rgba(0,0,0,0.5)';
            backdrop.style.zIndex = '9998';
            document.body.appendChild(backdrop);
            document.body.appendChild(logModal);
        } else {
            logModal.style.display = 'block';
            document.getElementById('export-backdrop').style.display = 'block';
            logContent.innerHTML = '';
            progressBar.style.width = '0%';
        }
    }

    function log(message, progress = null) {
        const p = document.createElement('p');
        p.style.margin = '3px 0';
        p.innerText = message;
        logContent.appendChild(p);
        logContent.scrollTop = logContent.scrollHeight;
        if (progress !== null) {
            progressBar.style.width = progress + '%';
        }
    }

    function hideLogger() {
        if (logModal) {
            logModal.style.display = 'none';
            document.getElementById('export-backdrop').style.display = 'none';
        }
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
            // Detect active Satker ID from page dropdown or URL
            let activeSatkerId = '';
            const satkerSelect = document.querySelector('#idSatker') || document.querySelector('select[name="idSatker"]') || document.querySelector('#satker') || document.querySelector('#id_satker');
            if (satkerSelect) {
                activeSatkerId = satkerSelect.value;
            }
            const urlParams = new URLSearchParams(window.location.search);
            if (!activeSatkerId) {
                activeSatkerId = urlParams.get('idSatker') || urlParams.get('satker') || urlParams.get('id_satker') || '';
            }
            
            if (activeSatkerId) {
                log(`Mendeteksi Satker Aktif ID: ${activeSatkerId}`, 8);
            } else {
                log(`Menggunakan Satker bawaan (Default)...`, 8);
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
            const rupRes = await postForm('/sirup/datatablectr/dataruppenyedia2018?tahun=2026', rupBody);
            const rupRows = rupRes.aaData || [];
            
            // Diagnostic fetch for Swakelola
            try {
                const swaRes = await postForm('/sirup/datatablectr/datarupswakelola2018?tahun=2026', rupBody);
                const swaRows = swaRes.aaData || [];
                if (swaRows.length > 0) {
                    console.log("Swakelola Row 0:", swaRows[0]);
                    log(`[Swakelola Diagnostic] Ditemukan ${swaRows.length} paket. Row 0: ` + JSON.stringify(swaRows[0]).substring(0, 150), 57);
                } else {
                    log("[Swakelola Diagnostic] Tidak ada paket swakelola ditemukan.", 57);
                }
            } catch (swaErr) {
                console.error("Swakelola fetch error:", swaErr);
                log("[Swakelola Diagnostic] Gagal mengambil: " + swaErr.message, 57);
            }
            
            const rupPackets = [];
            for (const row of rupRows) {
                if (row[0] == "63181122" || row[0] == 63181122) {
                    log(`[DEBUG 63181122] A=${row[6]} FD=${row[7]} U=${row[8]} raw=${JSON.stringify(row)}`, 58);
                }
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
            await buildExcel(rkaData, rupPackets, rup_all_lines);
            
            log('Excel berhasil di-generate dan diunduh!', 100);
            setTimeout(hideLogger, 2000);
            
        } catch (e) {
            log('ERROR: ' + e.message, null);
            console.error(e);
            const closeBtn = document.createElement('button');
            closeBtn.className = 'btn btn-danger';
            closeBtn.innerText = 'Tutup';
            closeBtn.style.marginTop = '10px';
            closeBtn.addEventListener('click', hideLogger);
            logContent.appendChild(closeBtn);
        }
    }

    // EXCEL BUILDER FUNCTION (ExcelJS)
    async function buildExcel(rkaData, rupPackets, rup_all_lines) {
        const wb = new ExcelJS.Workbook();
        
        // Colors
        const DARK_BLUE = "1F497D";
        const LIGHT_BLUE_LVL0 = "B8CCE4";
        const LIGHT_BLUE_LVL1 = "DCE6F1";
        const LIGHT_GRAY_LVL2 = "F2F2F2";
        const LIGHT_GRAY_LVL3 = "F9F9F9";
        const LIGHT_GREEN = "E2EFDA";
        const LIGHT_RED = "FCE4D6";
        const WHITE_COLOR = "FFFFFF";
        const LIGHT_ORANGE = "FFF2CC";
        
        // Borders template
        const border_thin = {
            top: { style: 'thin', color: { argb: 'FFD3D3D3' } },
            left: { style: 'thin', color: { argb: 'FFD3D3D3' } },
            bottom: { style: 'thin', color: { argb: 'FFD3D3D3' } },
            right: { style: 'thin', color: { argb: 'FFD3D3D3' } }
        };

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
        
        // Sum RUP totals (Only fully announced ones)
        for (const prog of rkaData) {
            const prog_code = prog.text.split("]")[0].replace("[", "").trim();
            for (const keg of prog.kegiatans || []) {
                for (const out of keg.outputs || []) {
                    const suboutputs = out.suboutputs || [{ id: 0, name: out.name, code: "000", komponens: out.komponens || [] }];
                    for (const ro of suboutputs) {
                        for (const komp of ro.komponens || []) {
                            const comp_key = `${prog_code}.${keg.code}.${out.code}.${ro.code}.${komp.code}`;
                            const comp_rup = rup_all_lines.filter(l => l.comp_key === comp_key && l.is_terumumkan);
                            total_rup_pagu += comp_rup.reduce((sum, l) => sum + l.pagu, 0);
                        }
                    }
                }
            }
        }

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
        const ws_dash = wb.addWorksheet("Dashboard Evaluasi");
        ws_dash.views = [{ showGridLines: true }];
        
        // Title
        ws_dash.getCell('A1').value = "DASHBOARD MONITORING & EVALUASI INTEGRASI RKA-RUP";
        ws_dash.getCell('A1').font = { name: "Segoe UI", size: 16, bold: true, color: { argb: 'FF1F497D' } };
        ws_dash.mergeCells("A1:G1");
        
        ws_dash.getCell('A2').value = `Tanggal Ekspor: ${new Date().toLocaleDateString('id-ID')} | Satker: Sekretariat Badan Penyuluhan dan Pengembangan SDM Kelautan dan Perikanan`;
        ws_dash.getCell('A2').font = { name: "Segoe UI", size: 10, italic: true, color: { argb: 'FF595959' } };
        ws_dash.mergeCells("A2:G2");
        
        // Card 1: TOTAL PAGU SATKER (A4:B6)
        ws_dash.mergeCells("A4:B4");
        ws_dash.getCell('A4').value = "TOTAL PAGU SATKER (A)";
        ws_dash.getCell('A4').font = { name: "Segoe UI", size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
        ws_dash.getCell('A4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F497D' } };
        ws_dash.getCell('A4').alignment = { horizontal: 'center', vertical: 'middle' };
        
        ws_dash.mergeCells("A5:B6");
        ws_dash.getCell('A5').value = total_satker_pagu;
        ws_dash.getCell('A5').font = { name: "Segoe UI", size: 14, bold: true, color: { argb: 'FF1F497D' } };
        ws_dash.getCell('A5').numFormat = 'Rp #,##0';
        ws_dash.getCell('A5').alignment = { horizontal: 'center', vertical: 'middle' };
        ws_dash.getCell('A5').border = border_thin;
        
        // Card 2: TARGET PENGADAAN (C4:D6)
        ws_dash.mergeCells("C4:D4");
        ws_dash.getCell('C4').value = "TARGET PENGADAAN (C = A - B)";
        ws_dash.getCell('C4').font = { name: "Segoe UI", size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
        ws_dash.getCell('C4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF203764' } };
        ws_dash.getCell('C4').alignment = { horizontal: 'center', vertical: 'middle' };
        
        ws_dash.mergeCells("C5:D6");
        ws_dash.getCell('C5').value = total_target_pengadaan;
        ws_dash.getCell('C5').font = { name: "Segoe UI", size: 14, bold: true, color: { argb: 'FF203764' } };
        ws_dash.getCell('C5').numFormat = 'Rp #,##0';
        ws_dash.getCell('C5').alignment = { horizontal: 'center', vertical: 'middle' };
        ws_dash.getCell('C5').border = border_thin;

        // Card 3: RUP TERUMUMKAN (E4:F6)
        ws_dash.mergeCells("E4:F4");
        ws_dash.getCell('E4').value = "RUP TERUMUMKAN (D)";
        ws_dash.getCell('E4').font = { name: "Segoe UI", size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
        ws_dash.getCell('E4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF375623' } };
        ws_dash.getCell('E4').alignment = { horizontal: 'center', vertical: 'middle' };
        
        ws_dash.mergeCells("E5:F6");
        ws_dash.getCell('E5').value = total_rup_pagu;
        ws_dash.getCell('E5').font = { name: "Segoe UI", size: 14, bold: true, color: { argb: 'FF375623' } };
        ws_dash.getCell('E5').numFormat = 'Rp #,##0';
        ws_dash.getCell('E5').alignment = { horizontal: 'center', vertical: 'middle' };
        ws_dash.getCell('E5').border = border_thin;

        // Card 4: PERSENTASE CAPAIAN (G4:H6)
        ws_dash.mergeCells("G4:H4");
        ws_dash.getCell('G4').value = "PERSENTASE CAPAIAN (%)";
        ws_dash.getCell('G4').font = { name: "Segoe UI", size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
        ws_dash.getCell('G4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7030A0' } };
        ws_dash.getCell('G4').alignment = { horizontal: 'center', vertical: 'middle' };
        
        tot_pct = total_target_pengadaan > 0 ? (total_rup_pagu / total_target_pengadaan) : 0;
        ws_dash.mergeCells("G5:H6");
        ws_dash.getCell('G5').value = tot_pct;
        ws_dash.getCell('G5').font = { name: "Segoe UI", size: 16, bold: true, color: { argb: tot_pct === 1 ? 'FF2E7D32' : 'FFC00000' } };
        ws_dash.getCell('G5').numFormat = '0.0%';
        ws_dash.getCell('G5').alignment = { horizontal: 'center', vertical: 'middle' };
        ws_dash.getCell('G5').border = border_thin;

        // Section Title: Ringkasan Statistik
        ws_dash.getCell('A8').value = "STATISTIK EVALUASI PAKET RKA-RUP";
        ws_dash.getCell('A8').font = { name: "Segoe UI", size: 12, bold: true, color: { argb: 'FF1F497D' } };
        
        // Table Headers (A10:D10)
        const stat_headers = ["No", "Indikator Evaluasi Anggaran & RUP", "Nilai / Jumlah", "Satuan"];
        for (let col = 1; col <= 4; col++) {
            const cell = ws_dash.getCell(10, col);
            cell.value = stat_headers[col-1];
            cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_BLUE } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = border_thin;
        }

        // Stats Rows
        const stats_data = [
            ["1", "Total Belanja Non-Pengadaan (NP / Gaji Satker) (B)", total_non_pengadaan, "Rupiah (Rp)"],
            ["2", "Jumlah Paket RUP Terdaftar (Penyedia)", rupPackets.length, "Paket"],
            ["3", "Jumlah Paket RUP Sah (Terumumkan KPA)", rupPackets.filter(p => p.aktif && p.fd && p.umumkan).length, "Paket"],
            ["4", "Jumlah Paket RUP Belum Terumumkan (Draft / Dibatalkan)", rupPackets.filter(p => !(p.aktif && p.fd && p.umumkan)).length, "Paket"],
            ["5", "Jumlah Paket RUP Kehilangan Sandingan (Potensi Salah Input MAK)", unique_unmatched_count, "Paket"]
        ];

        for (let i = 0; i < stats_data.length; i++) {
            const r_idx = 11 + i;
            const row = stats_data[i];
            
            ws_dash.getCell(r_idx, 1).value = row[0];
            ws_dash.getCell(r_idx, 1).alignment = { horizontal: 'center' };
            
            ws_dash.getCell(r_idx, 2).value = row[1];
            ws_dash.getCell(r_idx, 2).alignment = { horizontal: 'left' };
            
            const val_cell = ws_dash.getCell(r_idx, 3);
            val_cell.value = row[2];
            if (row[3] === "Rupiah (Rp)") {
                val_cell.numFormat = '#,##0';
                val_cell.alignment = { horizontal: 'right' };
            } else {
                val_cell.numFormat = '#,##0';
                val_cell.alignment = { horizontal: 'center' };
            }
            
            ws_dash.getCell(r_idx, 4).value = row[3] === "Rupiah (Rp)" ? "Rupiah (Rp)" : "Paket";
            ws_dash.getCell(r_idx, 4).alignment = { horizontal: 'center' };
            
            for (let c = 1; c <= 4; c++) {
                const cell = ws_dash.getCell(r_idx, c);
                cell.font = { name: "Segoe UI", size: 9 };
                cell.border = border_thin;
                if (c === 3 && (i === 3 || i === 4) && row[2] > 0) {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
                    cell.font = { name: "Segoe UI", size: 9, bold: true, color: { argb: 'FFC00000' } };
                }
            }
        }
        
        ws_dash.getColumn(1).width = 5;
        ws_dash.getColumn(2).width = 50;
        ws_dash.getColumn(3).width = 25;
        ws_dash.getColumn(4).width = 15;
        ws_dash.getColumn(5).width = 18;
        ws_dash.getColumn(6).width = 18;
        ws_dash.getColumn(7).width = 18;
        ws_dash.getColumn(8).width = 18;

        // ----------------- SHEET 1: RINGKASAN PAGU -----------------
        const ws_summary = wb.addWorksheet("Ringkasan Pagu");
        ws_summary.views = [{ showGridLines: true }];
        
        ws_summary.getCell('A1').value = "RINGKASAN PAGU ANGGARAN SATKER TAHUN 2026";
        ws_summary.getCell('A1').font = { name: "Segoe UI", size: 16, bold: true, color: { argb: 'FF1F497D' } };
        ws_summary.mergeCells("A1:L1");

        ws_summary.getCell('A3').value = "TOTAL PAGU SATKER";
        ws_summary.getCell('A3').font = { name: "Segoe UI", size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        ws_summary.getCell('A3').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_BLUE } };
        ws_summary.getCell('A3').alignment = { horizontal: 'center', vertical: 'middle' };
        
        total_satker_pagu = 0;
        
        const headers_summary = [
            "No", "Kode Program", "Nama Program", "Kode Kegiatan", "Nama Kegiatan",
            "Kode KRO", "Nama KRO", "Kode RO", "Nama RO", "Kode Komponen", "Nama Komponen", "Pagu Komponen"
        ];
        
        for (let c = 1; c <= headers_summary.length; c++) {
            const cell = ws_summary.getCell(6, c);
            cell.value = headers_summary[c-1];
            cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_BLUE } };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            cell.border = border_thin;
        }

        let row_idx = 7;
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
                            ws_summary.getCell(row_idx, 12).numFormat = '#,##0';
                            
                            for (let col_c = 1; col_c <= 12; col_c++) {
                                const cell = ws_summary.getCell(row_idx, col_c);
                                cell.border = border_thin;
                                cell.font = { name: "Segoe UI", size: 9 };
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

        ws_summary.getCell(row_idx, 11).value = "TOTAL";
        ws_summary.getCell(row_idx, 11).font = { name: "Segoe UI", size: 10, bold: true };
        ws_summary.getCell(row_idx, 11).alignment = { horizontal: 'right', vertical: 'middle' };
        ws_summary.getCell(row_idx, 11).border = border_thin;

        ws_summary.getCell(row_idx, 12).value = total_satker_pagu;
        ws_summary.getCell(row_idx, 12).font = { name: "Segoe UI", size: 10, bold: true };
        ws_summary.getCell(row_idx, 12).numFormat = '#,##0';
        ws_summary.getCell(row_idx, 12).alignment = { horizontal: 'right', vertical: 'middle' };
        ws_summary.getCell(row_idx, 12).border = border_thin;

        ws_summary.getCell('A4').value = total_satker_pagu;
        ws_summary.getCell('A4').font = { name: "Segoe UI", size: 18, bold: true, color: { argb: 'FF1F497D' } };
        ws_summary.getCell('A4').numFormat = 'Rp #,##0';
        ws_summary.getCell('A4').alignment = { horizontal: 'center', vertical: 'middle' };
        ws_summary.mergeCells("A3:C3");
        ws_summary.mergeCells("A4:C4");

        ws_summary.getColumn(1).width = 5;
        ws_summary.getColumn(2).width = 12;
        ws_summary.getColumn(3).width = 25;
        ws_summary.getColumn(4).width = 12;
        ws_summary.getColumn(5).width = 25;
        ws_summary.getColumn(6).width = 12;
        ws_summary.getColumn(7).width = 25;
        ws_summary.getColumn(8).width = 12;
        ws_summary.getColumn(9).width = 25;
        ws_summary.getColumn(10).width = 12;
        ws_summary.getColumn(11).width = 30;
        ws_summary.getColumn(12).width = 18;


        // ----------------- SHEET 2: SANDING RKA & RUP -----------------
        const ws_sanding = wb.addWorksheet("Sanding RKA & RUP");
        ws_sanding.views = [{ showGridLines: true }];
        
        ws_sanding.getCell('A1').value = "PENYANDINGAN PAGU RKA DENGAN REALISASI RUP (TERUMUMKAN)";
        ws_sanding.getCell('A1').font = { name: "Segoe UI", size: 16, bold: true, color: { argb: 'FF1F497D' } };
        ws_sanding.mergeCells("A1:N1");

        const headers_sanding = [
            "No", "Program", "Kegiatan", "KRO (Output)", "RO (Sub-Output)", "Komponen", 
            "Key Otorisasi (MAK)", "Pagu RKA (A)", "Pagu Non-Pengadaan (B)", 
            "Target Pengadaan (C = A - B)", "Pagu RUP Terumumkan (D)", 
            "Selisih Pengadaan (C - D)", "Persentase (%)", "Status Evaluasi"
        ];

        for (let c = 1; c <= headers_sanding.length; c++) {
            const cell = ws_sanding.getCell(4, c);
            cell.value = headers_sanding[c-1];
            cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_BLUE } };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            cell.border = border_thin;
        }

        function writeSandingRow(rowNum, numVal, prog_code, keg_name, kro_name, ro_name, comp_name, comp_key,
                                  pagu_rka, np_gaji_sum, target_pengadaan, rup_sum,
                                  bg_fill, font_style) {
            const selisih_pengadaan = target_pengadaan - rup_sum;
            
            let pct = 0;
            let status_text = "";
            let status_fill = WHITE_COLOR;
            
            if (target_pengadaan === 0) {
                if (rup_sum === 0) {
                    pct = 1.0;
                    status_text = "Sesuai (Non-Pengadaan)";
                    status_fill = LIGHT_GREEN;
                } else {
                    pct = 9.99;
                    status_text = "Kelebihan Umumkan";
                    status_fill = LIGHT_RED;
                }
            } else {
                pct = rup_sum / target_pengadaan;
                if (pct === 1.0) {
                    status_text = "Sesuai (100%)";
                    status_fill = LIGHT_GREEN;
                } else if (pct > 0 && pct < 1.0) {
                    status_text = `Parsial (${(pct*100).toFixed(1)}%)`;
                    status_fill = LIGHT_BLUE_LVL1;
                } else if (pct > 1.0) {
                    status_text = `Kelebihan (${(pct*100).toFixed(1)}%)`;
                    status_fill = LIGHT_RED;
                } else {
                    status_text = "Belum Diumumkan";
                    status_fill = LIGHT_RED;
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
            ws_sanding.getCell(rowNum, 8).numFormat = '#,##0';
            
            ws_sanding.getCell(rowNum, 9).value = np_gaji_sum;
            ws_sanding.getCell(rowNum, 9).numFormat = '#,##0';
            
            ws_sanding.getCell(rowNum, 10).value = target_pengadaan;
            ws_sanding.getCell(rowNum, 10).numFormat = '#,##0';
            
            ws_sanding.getCell(rowNum, 11).value = rup_sum;
            ws_sanding.getCell(rowNum, 11).numFormat = '#,##0';
            
            ws_sanding.getCell(rowNum, 12).value = selisih_pengadaan;
            ws_sanding.getCell(rowNum, 12).numFormat = '#,##0';
            
            ws_sanding.getCell(rowNum, 13).value = pct;
            ws_sanding.getCell(rowNum, 13).numFormat = '0.0%';
            
            const status_cell = ws_sanding.getCell(rowNum, 14);
            status_cell.value = status_text;
            status_cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: status_fill } };
            
            for (let col_c = 1; col_c <= 14; col_c++) {
                const cell = ws_sanding.getCell(rowNum, col_c);
                cell.border = border_thin;
                cell.font = font_style;
                if (bg_fill) {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg_fill } };
                }
                
                if (col_c === 1 || col_c === 2 || col_c === 7 || col_c === 14) {
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                } else if (col_c >= 8 && col_c <= 13) {
                    cell.alignment = { horizontal: 'right', vertical: 'middle' };
                } else {
                    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
                }
            }
        }

        row_idx = 5;
        num = 1;
        total_rup_pagu = 0;
        total_non_pengadaan = 0;
        total_target_pengadaan = 0;
        
        const font_sub = { name: "Segoe UI", size: 9, bold: true };
        const font_det = { name: "Segoe UI", size: 9 };

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

        // Total Row Sanding
        ws_sanding.getCell(row_idx, 7).value = "TOTAL";
        ws_sanding.getCell(row_idx, 7).font = { name: "Segoe UI", size: 10, bold: true };
        ws_sanding.getCell(row_idx, 7).alignment = { horizontal: 'right', vertical: 'middle' };
        ws_sanding.getCell(row_idx, 7).border = border_thin;

        ws_sanding.getCell(row_idx, 8).value = total_satker_pagu;
        ws_sanding.getCell(row_idx, 8).font = { name: "Segoe UI", size: 10, bold: true };
        ws_sanding.getCell(row_idx, 8).numFormat = '#,##0';
        ws_sanding.getCell(row_idx, 8).border = border_thin;

        ws_sanding.getCell(row_idx, 9).value = total_non_pengadaan;
        ws_sanding.getCell(row_idx, 9).font = { name: "Segoe UI", size: 10, bold: true };
        ws_sanding.getCell(row_idx, 9).numFormat = '#,##0';
        ws_sanding.getCell(row_idx, 9).border = border_thin;

        ws_sanding.getCell(row_idx, 10).value = total_target_pengadaan;
        ws_sanding.getCell(row_idx, 10).font = { name: "Segoe UI", size: 10, bold: true };
        ws_sanding.getCell(row_idx, 10).numFormat = '#,##0';
        ws_sanding.getCell(row_idx, 10).border = border_thin;

        ws_sanding.getCell(row_idx, 11).value = total_rup_pagu;
        ws_sanding.getCell(row_idx, 11).font = { name: "Segoe UI", size: 10, bold: true };
        ws_sanding.getCell(row_idx, 11).numFormat = '#,##0';
        ws_sanding.getCell(row_idx, 11).border = border_thin;

        ws_sanding.getCell(row_idx, 12).value = total_target_pengadaan - total_rup_pagu;
        ws_sanding.getCell(row_idx, 12).font = { name: "Segoe UI", size: 10, bold: true };
        ws_sanding.getCell(row_idx, 12).numFormat = '#,##0';
        ws_sanding.getCell(row_idx, 12).border = border_thin;

        tot_pct = total_target_pengadaan > 0 ? (total_rup_pagu / total_target_pengadaan) : 0;
        ws_sanding.getCell(row_idx, 13).value = tot_pct;
        ws_sanding.getCell(row_idx, 13).font = { name: "Segoe UI", size: 10, bold: true };
        ws_sanding.getCell(row_idx, 13).numFormat = '0.0%';
        ws_sanding.getCell(row_idx, 13).border = border_thin;

        ws_sanding.getCell('A2').value = `Total Pagu RKA: Rp ${total_satker_pagu.toLocaleString('id')} | Target Pengadaan: Rp ${total_target_pengadaan.toLocaleString('id')} | Terumumkan RUP: Rp ${total_rup_pagu.toLocaleString('id')} (${(tot_pct*100).toFixed(1)}%)`;
        ws_sanding.getCell('A2').font = { name: "Segoe UI", size: 10, italic: true };
        ws_sanding.mergeCells("A2:N2");

        ws_sanding.getColumn(1).width = 6;
        ws_sanding.getColumn(2).width = 10;
        ws_sanding.getColumn(3).width = 25;
        ws_sanding.getColumn(4).width = 25;
        ws_sanding.getColumn(5).width = 25;
        ws_sanding.getColumn(6).width = 25;
        ws_sanding.getColumn(7).width = 22;
        ws_sanding.getColumn(8).width = 15;
        ws_sanding.getColumn(9).width = 22;
        ws_sanding.getColumn(10).width = 22;
        ws_sanding.getColumn(11).width = 22;
        ws_sanding.getColumn(12).width = 22;
        ws_sanding.getColumn(13).width = 12;
        ws_sanding.getColumn(14).width = 25;


        // ----------------- SHEET 3: DAFTAR PAKET RUP -----------------
        const ws_rup = wb.addWorksheet("Daftar Paket RUP");
        ws_rup.views = [{ showGridLines: true }];
        
        ws_rup.getCell('A1').value = "DAFTAR PAKET PENYEDIA EXISTING YANG TERUMUMKAN (RUP)";
        ws_rup.getCell('A1').font = { name: "Segoe UI", size: 16, bold: true, color: { argb: 'FF1F497D' } };
        ws_rup.mergeCells("A1:L1");

        const headers_rup = [
            "No", "ID Paket", "Nama Kegiatan (RUP)", "Nama Paket", "Pagu RUP (Rp)",
            "Waktu Pemilihan", "Sumber Dana", "A (Draft PPK)", "FD (Final Draft PPK)",
            "U (Terumumkan KPA)", "Kode Otorisasi (MAK)", "Mata Anggaran Mapped"
        ];

        for (let c = 1; c <= headers_rup.length; c++) {
            const cell = ws_rup.getCell(4, c);
            cell.value = headers_rup[c-1];
            cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_BLUE } };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            cell.border = border_thin;
        }

        row_idx = 5;
        for (let idx = 0; idx < rupPackets.length; idx++) {
            const p = rupPackets[idx];
            ws_rup.getCell(row_idx, 1).value = idx + 1;
            ws_rup.getCell(row_idx, 2).value = p.id;
            ws_rup.getCell(row_idx, 3).value = p.keg_name;
            ws_rup.getCell(row_idx, 4).value = p.name;
            
            ws_rup.getCell(row_idx, 5).value = p.pagu;
            ws_rup.getCell(row_idx, 5).numFormat = '#,##0';
            
            ws_rup.getCell(row_idx, 6).value = p.waktu;
            ws_rup.getCell(row_idx, 7).value = p.sumber_dana;
            ws_rup.getCell(row_idx, 8).value = p.aktif ? "✓" : "";
            ws_rup.getCell(row_idx, 9).value = p.fd ? "✓" : "";
            ws_rup.getCell(row_idx, 10).value = p.umumkan ? "✓" : "";
            ws_rup.getCell(row_idx, 11).value = p.mak;
            
            const p_comp_keys = Array.from(new Set(rup_all_lines.filter(l => l.packet_id === p.id).map(l => l.comp_key)));
            ws_rup.getCell(row_idx, 12).value = p_comp_keys.join(", ");

            for (let col_c = 1; col_c <= 12; col_c++) {
                const cell = ws_rup.getCell(row_idx, col_c);
                cell.border = border_thin;
                cell.font = { name: "Segoe UI", size: 9 };
                if (col_c === 1 || col_c === 2 || col_c === 6 || col_c === 7 || (col_c >= 8 && col_c <= 10) || col_c === 11 || col_c === 12) {
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                } else if (col_c === 5) {
                    cell.alignment = { horizontal: 'right', vertical: 'middle' };
                } else {
                    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
                }
            }
            row_idx++;
        }

        ws_rup.getCell(row_idx, 4).value = "TOTAL";
        ws_rup.getCell(row_idx, 4).font = { name: "Segoe UI", size: 10, bold: true };
        ws_rup.getCell(row_idx, 4).alignment = { horizontal: 'right', vertical: 'middle' };
        ws_rup.getCell(row_idx, 4).border = border_thin;

        ws_rup.getCell(row_idx, 5).value = total_rup_pagu;
        ws_rup.getCell(row_idx, 5).font = { name: "Segoe UI", size: 10, bold: true };
        ws_rup.getCell(row_idx, 5).numFormat = '#,##0';
        ws_rup.getCell(row_idx, 5).border = border_thin;

        ws_rup.getColumn(1).width = 5;
        ws_rup.getColumn(2).width = 12;
        ws_rup.getColumn(3).width = 25;
        ws_rup.getColumn(4).width = 35;
        ws_rup.getColumn(5).width = 15;
        ws_rup.getColumn(6).width = 15;
        ws_rup.getColumn(7).width = 12;
        ws_rup.getColumn(8).width = 15;
        ws_rup.getColumn(9).width = 18;
        ws_rup.getColumn(10).width = 18;
        ws_rup.getColumn(11).width = 30;
        ws_rup.getColumn(12).width = 20;


        // ----------------- SHEET 3A: PAKET RUP TANPA SANDINGAN -----------------
        const ws_no_sanding = wb.addWorksheet("Paket RUP Tanpa Sandingan");
        ws_no_sanding.views = [{ showGridLines: true }];
        
        ws_no_sanding.getCell('A1').value = "DAFTAR PAKET RUP YANG TIDAK MEMILIKI SANDINGAN DI RKA";
        ws_no_sanding.getCell('A1').font = { name: "Segoe UI", size: 16, bold: true, color: { argb: 'FF1F497D' } };
        ws_no_sanding.mergeCells("A1:F1");

        const headers_no_sanding = [
            "No", "ID Paket RUP", "Nama Paket RUP", "Kode MAK di RUP", "Pagu Paket (Rp)", "Keterangan / Alasan"
        ];

        for (let c = 1; c <= headers_no_sanding.length; c++) {
            const cell = ws_no_sanding.getCell(4, c);
            cell.value = headers_no_sanding[c-1];
            cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_BLUE } };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            cell.border = border_thin;
        }





        let no_sanding_row = 5;
        for (let idx = 0; idx < unique_unmatched.length; idx++) {
            const item = unique_unmatched[idx];
            ws_no_sanding.getCell(no_sanding_row, 1).value = idx + 1;
            ws_no_sanding.getCell(no_sanding_row, 2).value = item.packet_id;
            ws_no_sanding.getCell(no_sanding_row, 3).value = item.packet_name;
            ws_no_sanding.getCell(no_sanding_row, 4).value = item.mak;
            
            ws_no_sanding.getCell(no_sanding_row, 5).value = item.pagu;
            ws_no_sanding.getCell(no_sanding_row, 5).numFormat = '#,##0';
            
            ws_no_sanding.getCell(no_sanding_row, 6).value = item.reason;

            for (let col_c = 1; col_c <= 6; col_c++) {
                const cell = ws_no_sanding.getCell(no_sanding_row, col_c);
                cell.border = border_thin;
                cell.font = { name: "Segoe UI", size: 9 };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT_RED } };
                
                if (col_c === 1 || col_c === 2 || col_c === 4) {
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                } else if (col_c === 5) {
                    cell.alignment = { horizontal: 'right', vertical: 'middle' };
                } else {
                    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
                }
            }
            no_sanding_row++;
        }

        ws_no_sanding.getColumn(1).width = 5;
        ws_no_sanding.getColumn(2).width = 15;
        ws_no_sanding.getColumn(3).width = 35;
        ws_no_sanding.getColumn(4).width = 30;
        ws_no_sanding.getColumn(5).width = 18;
        ws_no_sanding.getColumn(6).width = 45;


        // ----------------- SHEET 4 & 5: DETAIL PER PROGRAM -----------------
        for (const prog of rkaData) {
            const prog_code = prog.text.split("]")[0].replace("[", "").trim();
            const sheet_name = `Detail - ${prog_code}`;
            const ws = wb.addWorksheet(sheet_name);
            ws.views = [{ showGridLines: true }];
            
            ws.getCell('A1').value = `DETAIL RENCANA KERJA ANGGARAN (RKA) - PROGRAM ${prog_code}`;
            ws.getCell('A1').font = { name: "Segoe UI", size: 16, bold: true, color: { argb: 'FF1F497D' } };
            ws.mergeCells("A1:S1");
            
            ws.getCell('A3').value = "Program:";
            ws.getCell('A3').font = { name: "Segoe UI", size: 10, bold: true };
            ws.getCell('B3').value = prog.text;
            ws.getCell('B3').font = { name: "Segoe UI", size: 10 };
            ws.mergeCells("B3:S3");
            
            const headers_det = [
                "Kegiatan", "KRO (Output)", "RO (Sub-Output)", "Komponen", "Kode (P/K/O/SO/K/SK/A/D)",
                "Uraian", "Uraian Sebelum Revisi", "Pagu RKA", "Pagu Sebelum Revisi",
                "P", "S", "Multiyears", "NP", "Gaji",
                "ID Paket RUP", "Nama Paket RUP", "Pagu RUP (Announced)", "Selisih Pengadaan", "Rencana Pemilihan"
            ];
            
            for (let c = 1; c <= headers_det.length; c++) {
                const cell = ws.getCell(5, c);
                cell.value = headers_det[c-1];
                cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_BLUE } };
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                cell.border = border_thin;
            }
            
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
                            
                            const start_row = curr_row;
                            
                            // --- PRE-CALCULATE target_pagu FOR THIS COMPONENT WITH PROPAGATION ---
                            const rows_target = {};
                            const rows_is_np = {};
                            const rows_is_gj = {};
                            
                            let komp_np = false, komp_gj = false;
                            let subkomp_np = false, subkomp_gj = false;
                            let akun_np = false, akun_gj = false;
                            
                            for (let r_idx = 0; r_idx < komp.rows.length; r_idx++) {
                                const r = komp.rows[r_idx];
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
                            for (let r_idx = 0; r_idx < komp.rows.length; r_idx++) {
                                const r = komp.rows[r_idx];
                                if (r.level === 2) {
                                    curr_akun_idx = r_idx;
                                    rows_target[r_idx] = 0;
                                } else if (r.level === 3 && curr_akun_idx !== null) {
                                    rows_target[curr_akun_idx] += rows_target[r_idx];
                                }
                            }
                            
                            // Calculate level 1 (Sub-komponen) target pagus
                            let curr_sub_idx = null;
                            for (let r_idx = 0; r_idx < komp.rows.length; r_idx++) {
                                const r = komp.rows[r_idx];
                                if (r.level === 1) {
                                    curr_sub_idx = r_idx;
                                    rows_target[r_idx] = 0;
                                } else if (r.level === 3 && curr_sub_idx !== null) {
                                    rows_target[curr_sub_idx] += rows_target[r_idx];
                                }
                            }
                            
                            // Calculate level 0 (Komponen) target pagus
                            let curr_komp_idx = null;
                            for (let r_idx = 0; r_idx < komp.rows.length; r_idx++) {
                                const r = komp.rows[r_idx];
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
                            for (let r_idx = 0; r_idx < komp.rows.length; r_idx++) {
                                const r = komp.rows[r_idx];
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
                            
                            for (let row_idx_in_comp = 0; row_idx_in_comp < komp.rows.length; row_idx_in_comp++) {
                                const row_data = komp.rows[row_idx_in_comp];
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
                                    ws.getCell(curr_row, 8).numFormat = '#,##0';
                                } else {
                                    ws.getCell(curr_row, 8).value = pagu;
                                }
                                
                                if (typeof pagu_prev === 'number') {
                                    ws.getCell(curr_row, 9).value = pagu_prev;
                                    ws.getCell(curr_row, 9).numFormat = '#,##0';
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
                                    ws.getCell(curr_row, 17).numFormat = '#,##0';
                                    ws.getCell(curr_row, 17).alignment = { horizontal: 'right', vertical: 'middle' };
                                    ws.getCell(curr_row, 18).value = 0;
                                    ws.getCell(curr_row, 18).numFormat = '#,##0';
                                    ws.getCell(curr_row, 18).alignment = { horizontal: 'right', vertical: 'middle' };
                                } else {
                                    ws.getCell(curr_row, 17).value = rup_pagu_val;
                                    ws.getCell(curr_row, 17).numFormat = '#,##0';
                                    ws.getCell(curr_row, 17).alignment = { horizontal: 'right', vertical: 'middle' };
                                    
                                    ws.getCell(curr_row, 18).value = target_pagu - rup_pagu_val;
                                    ws.getCell(curr_row, 18).numFormat = '#,##0';
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
                                        cell.font = { name: "Segoe UI", size: 9, color: { argb: 'FF333333' } };
                                        cell.alignment = { horizontal: 'left', vertical: 'center', wrapText: true };
                                        cell.border = border_thin;
                                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FAFBFD' } };
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
            
            ws.autoFilter = 'E5:S' + (curr_row - 1);
            
            ws.views = [
                { state: 'frozen', xSplit: 0, ySplit: 5, activeCell: 'A6' }
            ];
            
            ws.getColumn(1).width = 25;
            ws.getColumn(2).width = 25;
            ws.getColumn(3).width = 25;
            ws.getColumn(4).width = 25;
            ws.getColumn(5).width = 15;
            ws.getColumn(6).width = 45;
            ws.getColumn(7).width = 45;
            ws.getColumn(8).width = 15;
            ws.getColumn(9).width = 15;
            for (let c = 10; c <= 14; c++) ws.getColumn(c).width = 8;
            ws.getColumn(15).width = 15;
            ws.getColumn(16).width = 30;
            ws.getColumn(17).width = 20;
            ws.getColumn(18).width = 20;
            ws.getColumn(19).width = 18;
        }

        // Save file
        const buffer = await wb.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        saveAs(blob, 'rekap_rka_rup_sanding.xlsx');
    }

    // Run injector
    setTimeout(injectButton, 2000);
})();
