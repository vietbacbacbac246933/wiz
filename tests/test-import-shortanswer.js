'use strict';
/**
 * Test import Word: nhận diện Phần trả lời ngắn, tách câu, lưu và chấm điểm.
 *
 * Chay: node tests/test-import-shortanswer.js
 */

const path = require('path');
const ROOT = path.resolve(__dirname, '..');
const { Document, Packer, Paragraph, TextRun } = require(path.join(ROOT, 'node_modules', 'docx'));
const JSZip = require(path.join(ROOT, 'node_modules', 'jszip'));
const { BASE, ensureServer, stopServer, uploadDocx } = require('./helpers');

let passed = 0, failed = 0;
function check(name, cond, detail) {
    if (cond) { passed++; console.log('  PASS  ' + name); }
    else { failed++; console.log('  FAIL  ' + name + (detail ? '  -> ' + detail : '')); }
}

const DNA_TEXT = 'Có 8 phân tử DNA tự nhân đôi một số lần bằng nhau đã tổng hợp được 112 mạch pôlinucleotide mới lấy nguyên liệu hoàn toàn từ môi trường nội bào. Số lần tự nhân đôi của mỗi phân tử DNA trên là bao nhiêu?';

async function buildDocx() {
    const doc = new Document({
        sections: [{
            children: [
                new Paragraph('PHẦN I. TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN'),
                new Paragraph('Câu 1. Thủ đô của Việt Nam là thành phố nào?'),
                new Paragraph('A. TP. Hồ Chí Minh'),
                new Paragraph('B. Hà Nội *'),
                new Paragraph('C. Đà Nẵng'),
                new Paragraph('D. Cần Thơ'),
                new Paragraph('Câu 2. Kết quả của 2 + 2 bằng bao nhiêu?'),
                new Paragraph('A. 3'),
                new Paragraph('B. 4 *'),
                new Paragraph('C. 5'),
                new Paragraph('D. 6'),

                new Paragraph('PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI'),
                new Paragraph('Câu 1. Xét các phát biểu sau về nước:'),
                new Paragraph('a) Nước có công thức hóa học là H2O. (Đ)'),
                new Paragraph('b) Nước không dẫn điện. (S)'),

                new Paragraph('PHẦN 3. TRẮC NGHIỆM TRẢ LỜI NGẮN'),
                new Paragraph('Câu 1. Tính 2 + 3.'),
                new Paragraph('Đáp án: 5'),
                new Paragraph('Câu 2. ' + DNA_TEXT),
                new Paragraph('Đáp án đúng: 3'),
                new Paragraph('Câu 3. Viết kết quả của 5/2 dưới dạng số thập phân.'),
                new Paragraph('Hết')
            ]
        }]
    });
    return Packer.toBuffer(doc);
}

async function preview(buffer) {
    return uploadDocx(buffer, 'de-tra-loi-ngan.docx');
}

async function exportHtml(body) {
    const res = await fetch(BASE + '/api/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
    });
    const buffer = Buffer.from(await res.arrayBuffer());
    let text = '';
    if (/application\/zip/.test(res.headers.get('content-type') || '')) {
        const zip = await JSZip.loadAsync(buffer);
        const name = Object.keys(zip.files).filter(function (n) { return /\.html$/i.test(n); })[0];
        if (name) text = await zip.files[name].async('string');
    } else {
        text = buffer.toString('utf8');
    }
    return { status: res.status, text: text };
}

async function buildNewForms() {
    function u(text) {
        return new Paragraph({ children: [new TextRun({ text: text, underline: {} })] });
    }
    const doc = new Document({
        sections: [{
            children: [
                new Paragraph('PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI'),
                new Paragraph('Câu 1. Xét các phát biểu sau:'),
                u('a) Mệnh đề được gạch chân.'),
                new Paragraph('b) Mệnh đề không gạch chân.'),
                u('c) Mệnh đề gạch chân thứ hai.'),
                new Paragraph('d) Mệnh đề còn lại.'),

                new Paragraph('PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN'),
                new Paragraph('Câu 1. Tính 2 + 3. Đáp án: 5'),
                new Paragraph('Câu 2. Tính 4 + 5.'),
                new Paragraph('9'),

                new Paragraph('PHẦN IV. (1 điểm) Học sinh trả lời các câu sau.'),
                new Paragraph('Câu 1. Tính 6 + 7.'),
                new Paragraph('13'),
                new Paragraph('Câu 2. Tính 3 + 4. Đáp án: 7')
            ]
        }]
    });
    return Packer.toBuffer(doc);
}

(async function main() {
    await ensureServer();

    console.log('\n=== 0. Gạch chân Đúng/Sai + đáp án ngắn viết cùng dòng/ở dòng riêng ===');
    const pNew = await preview(await buildNewForms());
    const qNew = (pNew.json && pNew.json.questions) || [];
    check('preview thành công', pNew.json && pNew.json.success === true, JSON.stringify(pNew.json).slice(0, 300));
    check('gạch chân = Đúng, còn lại = Sai',
        qNew[0] && JSON.stringify(qNew[0].correctTF) === JSON.stringify([true, false, true, false]),
        qNew[0] && JSON.stringify(qNew[0].correctTF));
    check('đáp án ngắn cùng dòng câu hỏi = 5', qNew[1] && qNew[1].correctAnswer === '5', qNew[1] && qNew[1].correctAnswer);
    check('không còn chữ "Đáp án" trong câu hỏi', qNew[1] && !/Đáp án/i.test(qNew[1].question), qNew[1] && qNew[1].question);
    check('đáp án ngắn ở dòng riêng = 9', qNew[2] && qNew[2].correctAnswer === '9', qNew[2] && qNew[2].correctAnswer);
    check('câu không ghi loại phần + số ở dòng riêng -> trả lời ngắn',
        qNew[3] && qNew[3].type === 'shortanswer' && qNew[3].correctAnswer === '13',
        qNew[3] && JSON.stringify({ t: qNew[3].type, a: qNew[3].correctAnswer }));
    check('câu không ghi loại phần + "Đáp án:" cùng dòng -> trả lời ngắn',
        qNew[4] && qNew[4].type === 'shortanswer' && qNew[4].correctAnswer === '7',
        qNew[4] && JSON.stringify({ t: qNew[4].type, a: qNew[4].correctAnswer }));
    check('các câu mới không cần kiểm tra', qNew.every(function (q) { return q.needsReview === false; }),
        JSON.stringify(qNew.map(function (q) { return q.reviewReasons; })));

    console.log('\n=== 1. Preview: nhận diện đúng Phần 3 ===');
    const docx = await buildDocx();
    const p = await preview(docx);
    check('preview thành công', p.json && p.json.success === true, JSON.stringify(p.json).slice(0, 300));
    const qs = (p.json && p.json.questions) || [];
    check('có 6 câu (2 một đáp án, 1 đúng/sai, 3 trả lời ngắn)', qs.length === 6, 'count=' + qs.length);
    check('câu 1 vẫn là một đáp án', qs[0] && qs[0].type === 'single', qs[0] && qs[0].type);
    check('câu 2 vẫn là một đáp án', qs[1] && qs[1].type === 'single', qs[1] && qs[1].type);
    check('câu 3 là đúng/sai', qs[2] && qs[2].type === 'truefalse', qs[2] && qs[2].type);
    check('câu 4 là trả lời ngắn', qs[3] && qs[3].type === 'shortanswer', qs[3] && qs[3].type);
    check('câu 5 là trả lời ngắn', qs[4] && qs[4].type === 'shortanswer', qs[4] && qs[4].type);
    check('câu 6 là trả lời ngắn', qs[5] && qs[5].type === 'shortanswer', qs[5] && qs[5].type);

    console.log('\n=== 2. Nội dung câu trả lời ngắn ===');
    check('câu trả lời ngắn đầu không có phương án A-D', qs[3] && qs[3].options.length === 0, JSON.stringify(qs[3] && qs[3].options));
    check('câu DNA giữ đủ nội dung', qs[4] && qs[4].question.indexOf('112 mạch pôlinucleotide') >= 0, qs[4] && qs[4].question);
    check('câu DNA không lẫn tiêu đề phần', qs[4] && !/PHẦN|TRẮC NGHIỆM TRẢ LỜI NGẮN/i.test(qs[4].question), qs[4] && qs[4].question);
    check('đáp án câu DNA trích được = 3', qs[4] && qs[4].correctAnswer === '3', qs[4] && qs[4].correctAnswer);
    check('câu 5/2 giữ nội dung phân số', qs[5] && /5\/2/.test(qs[5].question), qs[5] && qs[5].question);
    check('câu 5/2 không bị gộp vào câu DNA', qs[5] && /112 mạch/.test(qs[5].question) === false, qs[5] && qs[5].question);
    check('câu 5/2 chưa có đáp án -> cần kiểm tra', qs[5] && qs[5].correctAnswer === null && qs[5].needsReview === true, JSON.stringify(qs[5]));
    check('lý do nêu rõ thiếu đáp án', qs[5] && (qs[5].reviewReasons || []).join(' ').indexOf('đáp án') >= 0, JSON.stringify(qs[5] && qs[5].reviewReasons));
    check('"Hết" không lẫn vào câu cuối', qs[5] && !/Hết/i.test(qs[5].question), qs[5] && qs[5].question);

    console.log('\n=== 3. Xuất file: chặn câu trả lời ngắn thiếu đáp án ===');
    const payload = qs.map(q => ({
        include: true, type: q.type, question: q.question, options: q.options,
        correct: q.correct, correctList: q.correctList, correctTF: q.correctTF,
        correctAnswer: q.correctAnswer
    }));
    const bad = await exportHtml({ sessionId: p.json.sessionId, title: 'Đề trả lời ngắn', questions: payload });
    check('còn câu thiếu đáp án -> không xuất (400)', bad.status === 400, bad.text.slice(0, 200));

    payload[5].correctAnswer = '2.5';
    const ok = await exportHtml({ sessionId: p.json.sessionId, title: 'Đề trả lời ngắn', questions: payload });
    check('điền đáp án xong -> xuất được (200)', ok.status === 200, ok.text.slice(0, 200));
    check('file HTML có đủ 6 câu', ok.text.indexOf('"questions"') > 0 && ok.text.indexOf('112 mạch') > 0);
    check('đáp án trả lời ngắn 2.5 nằm trong file', ok.text.indexOf('2.5') > 0);

    console.log('\n----------------------------------------');
    console.log('Kết quả: ' + passed + ' PASS, ' + failed + ' FAIL');
    stopServer();
    process.exit(failed ? 1 : 0);
})().catch(e => {
    console.error(e);
    stopServer();
    process.exit(1);
});
