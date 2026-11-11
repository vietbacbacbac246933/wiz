'use strict';
/**
 * Test luong: .pdf -> /api/extract -> /api/export -> 1 file HTML (giong file Word).
 *
 * PDF duoc tao ngay trong test bang tay (khong can thu vien): 1 trang, font
 * Helvetica, noi dung khong nen. Chay:
 *   node tests/test-import-pdf.js
 */

const path = require('path');
const { spawn } = require('child_process');
const ROOT = path.resolve(__dirname, '..');

const PORT = 3999;
const BASE = process.env.TEST_BASE || ('http://127.0.0.1:' + PORT);

let passed = 0, failed = 0, child = null;
function check(name, cond, detail) {
    if (cond) { passed++; console.log('  PASS  ' + name); }
    else { failed++; console.log('  FAIL  ' + name + (detail ? '  -> ' + detail : '')); }
}

// PDF toi gian: 1 trang, moi phan tu cua `lines` la 1 dong van ban.
function buildPdf(lines) {
    const content = 'BT /F1 12 Tf 14 TL 50 780 Td\n' +
        lines.map(function (l) { return '(' + l.replace(/([\\()])/g, '\\$1') + ') Tj T*'; }).join('\n') +
        '\nET';
    const objects = [
        '<< /Type /Catalog /Pages 2 0 R >>',
        '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
        '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
        '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>',
        '<< /Length ' + Buffer.byteLength(content, 'latin1') + ' >>\nstream\n' + content + '\nendstream'
    ];
    let out = '%PDF-1.4\n';
    const offsets = [];
    objects.forEach(function (body, i) {
        offsets.push(Buffer.byteLength(out, 'latin1'));
        out += (i + 1) + ' 0 obj\n' + body + '\nendobj\n';
    });
    const xref = Buffer.byteLength(out, 'latin1');
    out += 'xref\n0 ' + (objects.length + 1) + '\n0000000000 65535 f \n';
    offsets.forEach(function (o) { out += String(o).padStart(10, '0') + ' 00000 n \n'; });
    out += 'trailer\n<< /Size ' + (objects.length + 1) + ' /Root 1 0 R >>\nstartxref\n' + xref + '\n%%EOF\n';
    return Buffer.from(out, 'latin1');
}

const PDF_TYPE = 'application/pdf';

async function ensureServer() {
    try {
        const r = await fetch(BASE + '/api/health');
        if (r.ok) return;
    } catch (e) { /* chua chay */ }
    child = spawn(process.execPath, ['server.js'], {
        cwd: ROOT,
        env: Object.assign({}, process.env, { PORT: String(PORT) }),
        stdio: 'ignore'
    });
    for (let i = 0; i < 60; i++) {
        await new Promise(r => setTimeout(r, 200));
        try {
            const r = await fetch(BASE + '/api/health');
            if (r.ok) return;
        } catch (e) { /* dang khoi dong */ }
    }
    throw new Error('Server khong khoi dong duoc o ' + BASE);
}

async function upload(name, buffer, type) {
    const fd = new FormData();
    fd.append('file', new Blob([buffer], { type: type || PDF_TYPE }), name);
    const res = await fetch(BASE + '/api/extract', { method: 'POST', body: fd });
    let json = null;
    try { json = await res.json(); } catch (e) { /* ignore */ }
    return { status: res.status, json };
}

(async function main() {
    await ensureServer();

    console.log('\n=== 1. PDF co lop van ban -> trich xuat cau hoi ===');
    const pdf = buildPdf([
        'DE KIEM TRA MON TOAN - 45 PHUT',
        'Cau 1. 2 + 2 bang may?',
        'A. 3',
        'B. 4 *',
        'C. 5',
        'D. 6',
        'Cau 2. Thu do cua Viet Nam la thanh pho nao?',
        'A. TP. Ho Chi Minh',
        'B. Da Nang',
        'C. Ha Noi *',
        'D. Can Tho',
        'Cau 3. 5 x 3 bang bao nhieu?',
        'A. 15',
        'B. 20',
        'C. 25',
        'D. 30'
    ]);
    const res = await upload('de-thi.pdf', pdf);
    check('extract thanh cong', res.status === 200 && res.json.success === true, JSON.stringify(res.json).slice(0, 200));
    check('nhan 3 cau', res.json && res.json.count === 3, 'count=' + (res.json && res.json.count));
    check('co sessionId', !!(res.json && res.json.sessionId));
    check('source.kind = pdf', res.json && res.json.source && res.json.source.kind === 'pdf', JSON.stringify(res.json && res.json.source));
    check('source.name = ten file', res.json && res.json.source && res.json.source.name === 'de-thi.pdf');

    const qs = (res.json && res.json.questions) || [];
    check('cau 1 nhan dap an B', qs[0] && qs[0].correct === 1, 'correct=' + (qs[0] && qs[0].correct));
    check('cau 1 co 4 phuong an', qs[0] && qs[0].options.length === 4, JSON.stringify(qs[0] && qs[0].options));
    check('cau 2 nhan dap an C', qs[1] && qs[1].correct === 2, 'correct=' + (qs[1] && qs[1].correct));
    check('cau 3 thieu dau dap an -> correct=null + needsReview',
        qs[2] && qs[2].correct === null && qs[2].needsReview === true, JSON.stringify(qs[2] && { c: qs[2].correct, n: qs[2].needsReview }));
    check('co canh bao mat dinh dang PDF',
        (res.json.sourceWarnings || []).join(' ').indexOf('PDF') === 0, JSON.stringify(res.json && res.json.sourceWarnings));

    console.log('\n=== 2. Xuat file HTML tu PDF (giong file Word) ===');
    // Cau 3 thieu dap an nen giao vien phai chon truoc khi xuat (giong file Word).
    const exportQuestions = qs.map(function (q, i) {
        return i === 2 ? Object.assign({}, q, { correct: 0, needsReview: false, reviewReasons: [] }) : q;
    });
    const ex = await fetch(BASE + '/api/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            sessionId: res.json.sessionId,
            title: 'De kiem tra tu PDF',
            questions: exportQuestions,
            settings: { mode: 'exam', minutes: 45 }
        })
    });
    const html = await ex.text();
    check('export tra 200', ex.status === 200, 'status=' + ex.status);
    check('la file HTML tu chay', /^<!doctype html>/i.test(html) && html.indexOf('id="quiz-data"') > 0,
        html.slice(0, 120));
    check('co du lieu cau hoi trong file xuat', html.indexOf('2 + 2 bang may') > 0);

    console.log('\n=== 3. PDF khong co chu (ban scan) -> bao loi ro rang ===');
    const scan = await upload('scan.pdf', buildPdf([]));
    check('pdf rong -> 400', scan.status === 400, 'status=' + scan.status);
    check('thong bao nhac toi ban scan/khong co noi dung chu',
        /scan|không có nội dung chữ|không đọc/i.test((scan.json && scan.json.error) || ''), JSON.stringify(scan.json));

    console.log('\n=== 4. File gia mao / sai dinh dang ===');
    const fake = await upload('gia.pdf', Buffer.from('day khong phai pdf'));
    check('.pdf gia bi tu choi', fake.status === 400 && /PDF/i.test(fake.json.error || ''), JSON.stringify(fake.json));
    const docxAsPdf = await upload('de.docx', Buffer.from([0x50, 0x4B, 0x03, 0x04, 0x41]), PDF_TYPE);
    check('noi dung .docx doi ten .pdf bi tu choi', docxAsPdf.status === 400, 'status=' + docxAsPdf.status);
    const big = await upload('qua-lon.pdf', Buffer.alloc(16 * 1024 * 1024, 0x41));
    check('pdf 16MB bi chan', big.status === 400 && /quá lớn/i.test(big.json.error || ''), JSON.stringify(big.json));

    console.log('\n' + passed + ' pass, ' + failed + ' fail');
    if (child) child.kill();
    process.exit(failed ? 1 : 0);
})().catch(function (err) {
    console.error('LOI TEST:', err);
    if (child) child.kill();
    process.exit(1);
});
