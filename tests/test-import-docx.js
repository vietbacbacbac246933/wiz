'use strict';
/**
 * Test luong chinh: .docx -> /api/extract -> /api/export -> 1 file HTML doc lap.
 *
 * Tu khoi dong server o cong 3999 neu chua co. Chay:
 *   node tests/test-import-docx.js
 */

const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const ROOT = path.resolve(__dirname, '..');
const { Document, Packer, Paragraph, TextRun, ImageRun } = require(path.join(ROOT, 'node_modules', 'docx'));

const PORT = 3999;
const BASE = process.env.TEST_BASE || ('http://127.0.0.1:' + PORT);
const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

let passed = 0, failed = 0, child = null;
function check(name, cond, detail) {
    if (cond) { passed++; console.log('  PASS  ' + name); }
    else { failed++; console.log('  FAIL  ' + name + (detail ? '  -> ' + detail : '')); }
}

const PNG_1x1 = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=',
    'base64'
);

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

function docxBuffer(children) {
    return Packer.toBuffer(new Document({ sections: [{ children }] }));
}

async function upload(name, buffer) {
    const fd = new FormData();
    fd.append('file', new Blob([buffer], { type: DOCX_TYPE }), name);
    const res = await fetch(BASE + '/api/extract', { method: 'POST', body: fd });
    let json = null;
    try { json = await res.json(); } catch (e) { /* ignore */ }
    return { status: res.status, json };
}

async function exportHtml(payload) {
    const res = await fetch(BASE + '/api/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
    });
    const text = await res.text();
    return { status: res.status, headers: res.headers, text };
}

(async function main() {
    await ensureServer();

    console.log('\n=== 1. Trang chu + health ===');
    const home = await fetch(BASE + '/');
    const homeBody = await home.text();
    check('/ tra ve trang giao dien', home.ok && /api\/extract/.test(homeBody), 'status=' + home.status);
    const health = await (await fetch(BASE + '/api/health')).json();
    check('/api/health -> ok', health.ok === true);

    console.log('\n=== 2. File mau that: data/de-toan-mau.docx ===');
    const mau = await upload('de-toan-mau.docx', fs.readFileSync(path.join(ROOT, 'data', 'de-toan-mau.docx')));
    check('extract thanh cong', mau.json && mau.json.success === true, JSON.stringify(mau.json).slice(0, 200));
    check('dung 10 cau', mau.json && mau.json.count === 10, 'count=' + (mau.json && mau.json.count));
    check('co sessionId', !!(mau.json && mau.json.sessionId));

    const mq = (mau.json && mau.json.questions) || [];
    check('cau 1 nhan dap an B', mq[0] && mq[0].correct === 1, 'correct=' + (mq[0] && mq[0].correct));
    check('cau nhieu dap an co correctList [0,1,3]',
        mq[6] && JSON.stringify(mq[6].correctList) === JSON.stringify([0, 1, 3]),
        JSON.stringify(mq[6] && mq[6].correctList));
    check('cau dung/sai co correctTF', mq[8] && JSON.stringify(mq[8].correctTF) === JSON.stringify([true, false, true, false]),
        JSON.stringify(mq[8] && mq[8].correctTF));
    check('cong thuc da giu LaTeX', /\\\(/.test(mq[0] && mq[0].question), JSON.stringify(mq[0] && mq[0].question).slice(0, 120));
    check('khong co cau nao bi danh dau can kiem tra', mau.json.needsReview === 0, 'needsReview=' + (mau.json && mau.json.needsReview));

    console.log('\n=== 3. Cau thieu dap an -> correct=null + needsReview ===');
    const missing = await docxBuffer([
        new Paragraph('Câu 1: Thủ đô của Việt Nam là thành phố nào?'),
        new Paragraph('A. TP. Hồ Chí Minh'),
        new Paragraph('B. Hà Nội *'),
        new Paragraph('C. Đà Nẵng'),
        new Paragraph('D. Cần Thơ'),
        new Paragraph('Câu 2: Kết quả của 2 + 2 bằng bao nhiêu?'),
        new Paragraph('A. 3'),
        new Paragraph('B. 4'),
        new Paragraph('C. 5'),
        new Paragraph('D. 6')
    ]);
    const miss = await upload('thieu-dap-an.docx', missing);
    const missQ = (miss.json && miss.json.questions) || [];
    check('nhan 2 cau', miss.json && miss.json.count === 2, 'count=' + (miss.json && miss.json.count));
    check('cau thieu dap an: correct = null', missQ[1] && missQ[1].correct === null, JSON.stringify(missQ[1] && missQ[1].correct));
    check('cau thieu dap an: needsReview = true', missQ[1] && missQ[1].needsReview === true);
    check('ly do co nhac den dap an',
        missQ[1] && (missQ[1].reviewReasons || []).join(' ').toLowerCase().indexOf('đáp án') >= 0,
        JSON.stringify(missQ[1] && missQ[1].reviewReasons));

    console.log('\n=== 4. File sai dinh dang / qua lon / hong ===');
    const txt = await upload('de-thi.txt', Buffer.from('hello'));
    check('.txt bi tu choi (400)', txt.status === 400, 'status=' + txt.status);
    const fake = await upload('gia.docx', Buffer.from('day khong phai file word'));
    check('.docx gia bi tu choi', fake.status === 400 && /docx/i.test(fake.json.error || ''), JSON.stringify(fake.json));
    const corrupt = await upload('hong.docx', Buffer.concat([Buffer.from([0x50, 0x4B, 0x03, 0x04]), Buffer.alloc(300, 0x41)]));
    check('docx hong -> 400, khong sap server', corrupt.status === 400 && corrupt.json.success === false, JSON.stringify(corrupt.json).slice(0, 160));
    const big = await upload('qua-lon.docx', Buffer.alloc(16 * 1024 * 1024, 0x41));
    check('file 16MB bi chan', big.status === 400 && /quá lớn/i.test(big.json.error || ''), JSON.stringify(big.json));

    console.log('\n=== 5. Xuat HTML tu file mau ===');
    const ex = await exportHtml({
        sessionId: mau.json.sessionId,
        title: 'Đề kiểm tra Toán',
        subject: 'Toán',
        questions: mau.json.questions,
        settings: { mode: 'exam', minutes: 45 }
    });
    check('export tra 200', ex.status === 200, 'status=' + ex.status);
    check('Content-Type text/html', /text\/html/.test(ex.headers.get('content-type') || ''), ex.headers.get('content-type'));
    check('Content-Disposition .html', /attachment;.*\.html/.test(ex.headers.get('content-disposition') || ''),
        ex.headers.get('content-disposition'));
    check('la file HTML day du', /^<!doctype html>/i.test(ex.text) && ex.text.indexOf('</html>') > 0);
    check('co du lieu cau hoi', ex.text.indexOf('id="quiz-data"') > 0 && ex.text.indexOf('Gia tri cua') > 0);
    check('cong thuc da thanh MathML', ex.text.indexOf('u003cmath') > 0);
    check('khong con token anh', ex.text.indexOf('[[IMG:') < 0);
    check('khong phu thuoc mang ngoai', !/<script[^>]+src=/i.test(ex.text) && !/<link[^>]+href=/i.test(ex.text));
    check('khong co manifest/service worker', ex.text.indexOf('manifest.webmanifest') < 0 && ex.text.indexOf('sw.js') < 0);
    check('khong goi /api/', !/(?:src|href|fetch\()\s*[="'(]*\/api\//i.test(ex.text));
    check('khong co http://localhost', ex.text.indexOf('http://localhost') < 0);
    check('co lich su lam bai (localStorage)', /localStorage/.test(ex.text) && /Lịch sử làm bài/.test(ex.text));
    check('co cham diem trong file', /scoreQuestion|function score/.test(ex.text));

    console.log('\n=== 6. Anh nhung -> data URI trong file xuat ===');
    const imgDoc = await docxBuffer([
        new Paragraph({
            children: [
                new TextRun('Câu 1: Hình minh họa: '),
                new ImageRun({ type: 'png', data: PNG_1x1, transformation: { width: 10, height: 10 } })
            ]
        }),
        new Paragraph('A. Một *'),
        new Paragraph('B. Hai')
    ]);
    const img = await upload('co-anh.docx', imgDoc);
    const imgQ = (img.json && img.json.questions) || [];
    check('extract anh: co token [[IMG:n]]', imgQ[0] && /\[\[IMG:\d+\]\]/.test(imgQ[0].question), JSON.stringify(imgQ[0] && imgQ[0].question));
    check('extract anh: imagePool co du lieu', !!(img.json && img.json.imagePool && Object.keys(img.json.imagePool).length));
    const imgEx = await exportHtml({
        sessionId: img.json.sessionId,
        title: 'Đề có ảnh',
        questions: img.json.questions,
        settings: { mode: 'practice' }
    });
    check('file xuat co anh data URI', imgEx.status === 200 && imgEx.text.indexOf('data:image/png;base64,') > 0);
    check('file xuat khong con token anh', imgEx.text.indexOf('[[IMG:') < 0);

    console.log('\n=== 7. Export tu choi du lieu sai / phien het han ===');
    const noSession = await exportHtml({ sessionId: 'khong-ton-tai', questions: mau.json.questions, title: 'x' });
    check('session khong ton tai -> 400', noSession.status === 400);
    const badQ = JSON.parse(JSON.stringify(mau.json.questions));
    badQ[1].correct = null;
    const bad = await exportHtml({ sessionId: mau.json.sessionId, title: 'x', questions: badQ });
    check('cau thieu dap an -> khong xuat (400)', bad.status === 400 && /chưa hợp lệ/i.test(JSON.parse(bad.text).error || ''),
        bad.text.slice(0, 160));
    check('file HTML khong duoc tao khi co cau loi', bad.text.indexOf('<!doctype html>') < 0);

    console.log('\n=== 8. Gioi han phien RAM ===');
    const importSessions = require(path.join(ROOT, 'lib', 'import-sessions'));
    importSessions.clearAll();
    for (let i = 0; i < 30; i++) importSessions.create({ pool: {} });
    check('so phien khong vuot tran 20', importSessions.count() <= 20, 'count=' + importSessions.count());
    importSessions.clearAll();

    console.log('\n----------------------------------------');
    console.log('RESULT: ' + passed + ' passed, ' + failed + ' failed');
    if (child) child.kill();
    process.exit(failed ? 1 : 0);
})().catch(function (err) {
    console.error('\nLoi khong mong doi: ' + (err && err.stack ? err.stack : err));
    if (child) child.kill();
    process.exit(2);
});
