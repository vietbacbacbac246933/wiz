'use strict';
/**
 * Test import Word: đọc bảng đáp án ở cuối file và ghép vào đúng phần/câu/ý.
 *
 * Chay: node tests/test-answer-tables.js
 */

const path = require('path');
const ROOT = path.resolve(__dirname, '..');
const { Document, Packer, Paragraph, Table, TableRow, TableCell, VerticalMergeType } = require(path.join(ROOT, 'node_modules', 'docx'));
const JSZip = require(path.join(ROOT, 'node_modules', 'jszip'));
const { BASE, ensureServer, stopServer, uploadDocx } = require('./helpers');

let passed = 0, failed = 0;
function check(name, cond, detail) {
    if (cond) { passed++; console.log('  PASS  ' + name); }
    else { failed++; console.log('  FAIL  ' + name + (detail ? '  -> ' + detail : '')); }
}

function cell(text, opts) {
    opts = opts || {};
    return new TableCell({
        children: [new Paragraph(text == null ? '' : String(text))],
        columnSpan: opts.span,
        verticalMerge: opts.vMerge
    });
}
function table(rows) {
    return new Table({
        rows: rows.map(function (r) {
            return new TableRow({ children: r.map(function (c) { return cell(c.text, c); }) });
        })
    });
}
async function docx(children) {
    return Packer.toBuffer(new Document({ sections: [{ children: children }] }));
}

async function preview(buffer, name) {
    return uploadDocx(buffer, name || 'de.docx');
}

function tableSummary(json) {
    return (json && json.answerTables && json.answerTables.tables) || [];
}

function singleQuestion(n, text, correctAt) {
    const letters = ['A', 'B', 'C', 'D'];
    const out = [new Paragraph('Câu ' + n + '. ' + text)];
    letters.forEach(function (l, i) {
        out.push(new Paragraph(l + '. ' + (i === correctAt ? 'đúng *' : 'phương án ' + i)));
    });
    return out;
}

async function buildHorizontalSingle() {
    return docx([
        ...singleQuestion(1, 'Một cộng một bằng mấy?', 1),
        ...singleQuestion(2, 'Hai cộng hai bằng mấy?', 2),
        new Paragraph('ĐÁP ÁN TRẮC NGHIỆM'),
        table([
            ['Câu', '1', '2'].map(function (t) { return { text: t }; }),
            ['Đáp án', 'B', 'C'].map(function (t) { return { text: t }; })
        ])
    ]);
}

async function buildVerticalSingle() {
    return docx([
        new Paragraph('PHẦN I. TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN'),
        ...singleQuestion(1, 'Thủ đô Việt Nam?', 1),
        ...singleQuestion(2, 'Một cộng hai bằng mấy?', 2),
        new Paragraph('ĐÁP ÁN PHẦN I'),
        table([
            ['Câu', 'Đáp án'].map(function (t) { return { text: t }; }),
            ['1', 'B'].map(function (t) { return { text: t }; }),
            ['2', 'C'].map(function (t) { return { text: t }; })
        ])
    ]);
}

async function buildTrueFalse() {
    return docx([
        new Paragraph('PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI'),
        new Paragraph('Câu 1. Xét các phát biểu sau:'),
        new Paragraph('a) Nước có công thức H2O.'),
        new Paragraph('b) Nước không dẫn điện.'),
        new Paragraph('c) Nước đóng băng ở 0 độ C.'),
        new Paragraph('d) Nước có vị ngọt.'),
        new Paragraph('ĐÁP ÁN TRẮC NGHIỆM ĐÚNG SAI'),
        table([
            ['Câu', 'Lệnh hỏi', 'Đáp án'].map(function (t) { return { text: t }; }),
            ['1', 'a', 'Đ'].map(function (t) { return { text: t }; }),
            ['1', 'b', 'S'].map(function (t) { return { text: t }; }),
            ['1', 'c', 'Đ'].map(function (t) { return { text: t }; }),
            ['1', 'd', 'S'].map(function (t) { return { text: t }; })
        ])
    ]);
}

async function buildTrueFalseTwoGroups() {
    return docx([
        new Paragraph('PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI'),
        new Paragraph('Câu 1. Xét các phát biểu sau:'),
        new Paragraph('a) Mệnh đề 1a.'), new Paragraph('b) Mệnh đề 1b.'),
        new Paragraph('c) Mệnh đề 1c.'), new Paragraph('d) Mệnh đề 1d.'),
        new Paragraph('Câu 2. Xét các phát biểu sau:'),
        new Paragraph('a) Mệnh đề 2a.'), new Paragraph('b) Mệnh đề 2b.'),
        new Paragraph('c) Mệnh đề 2c.'), new Paragraph('d) Mệnh đề 2d.'),
        new Paragraph('ĐÁP ÁN PHẦN II'),
        table([
            ['Câu', 'Lệnh hỏi', 'Đáp án', 'Câu', 'Đáp án'].map(function (t) { return { text: t }; }),
            [{ text: '1', vMerge: VerticalMergeType.RESTART }, { text: 'a' }, { text: 'Đ' },
                { text: '2', vMerge: VerticalMergeType.RESTART }, { text: 'Đ' }],
            [{ text: '', vMerge: VerticalMergeType.CONTINUE }, { text: 'b' }, { text: 'Đ' },
                { text: '', vMerge: VerticalMergeType.CONTINUE }, { text: 'S' }],
            [{ text: '', vMerge: VerticalMergeType.CONTINUE }, { text: 'c' }, { text: 'S' },
                { text: '', vMerge: VerticalMergeType.CONTINUE }, { text: 'Đ' }],
            [{ text: '', vMerge: VerticalMergeType.CONTINUE }, { text: 'd' }, { text: 'S' },
                { text: '', vMerge: VerticalMergeType.CONTINUE }, { text: 'S' }]
        ])
    ]);
}

async function buildShortTwoPairs() {
    return docx([
        new Paragraph('PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN'),
        new Paragraph('Câu 1. Tính 2 + 3.'),
        new Paragraph('Câu 2. Tính 8 x 3.'),
        new Paragraph('ĐÁP ÁN TRẮC NGHIỆM TRẢ LỜI NGẮN'),
        table([
            [{ text: 'Câu', span: 2 }, { text: 'Đáp án', span: 2 }, { text: 'Câu', span: 2 }, { text: 'Đáp án', span: 2 }],
            [{ text: '1', span: 2 }, { text: '3', span: 2 }, { text: '2', span: 2 }, { text: '24', span: 2 }]
        ])
    ]);
}

async function buildMissingAnswer() {
    return docx([
        new Paragraph('PHẦN I. TRẮC NGHIỆM'),
        ...singleQuestion(1, 'Câu có đáp án?', 1),
        // Câu 2 không đánh dấu đáp án trong đề; bảng cũng thiếu câu 2.
        ...singleQuestion(2, 'Câu thiếu đáp án?', -1),
        new Paragraph('ĐÁP ÁN PHẦN I'),
        table([
            ['Câu', 'Đáp án'].map(function (t) { return { text: t }; }),
            ['1', 'B'].map(function (t) { return { text: t }; })
        ])
    ]);
}

async function buildEmptyCell() {
    return docx([
        new Paragraph('PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN'),
        new Paragraph('Câu 1. Tính 7 + 1.'),
        new Paragraph('ĐÁP ÁN PHẦN III'),
        table([
            ['Câu', 'Đáp án'].map(function (t) { return { text: t }; }),
            ['1', ''].map(function (t) { return { text: t }; })
        ])
    ]);
}

async function buildDecimalShort() {
    return docx([
        new Paragraph('PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN'),
        new Paragraph('Câu 1. Nhiệt độ là bao nhiêu?'),
        new Paragraph('Câu 2. Khối lượng là bao nhiêu?'),
        new Paragraph('ĐÁP ÁN PHẦN III'),
        table([
            ['Câu', 'Đáp án'].map(function (t) { return { text: t }; }),
            ['1', '-3,5'].map(function (t) { return { text: t }; }),
            ['2', '24.00'].map(function (t) { return { text: t }; })
        ])
    ]);
}

async function buildConflict() {
    return docx([
        new Paragraph('PHẦN I. TRẮC NGHIỆM'),
        ...singleQuestion(1, 'Câu có hai đáp án mâu thuẫn?', -1),
        new Paragraph('ĐÁP ÁN PHẦN I'),
        table([
            ['Câu', 'Đáp án', 'Câu', 'Đáp án'].map(function (t) { return { text: t }; }),
            ['1', 'A', '1', 'B'].map(function (t) { return { text: t }; })
        ])
    ]);
}

async function buildInvalidAnswer() {
    return docx([
        new Paragraph('PHẦN I. TRẮC NGHIỆM'),
        ...singleQuestion(1, 'Câu có đáp án không hợp lệ?', -1),
        new Paragraph('ĐÁP ÁN PHẦN I'),
        table([
            ['Câu', 'Đáp án'].map(function (t) { return { text: t }; }),
            ['1', 'Z'].map(function (t) { return { text: t }; })
        ])
    ]);
}

async function buildExtraAnswer() {
    return docx([
        new Paragraph('PHẦN I. TRẮC NGHIỆM'),
        ...singleQuestion(1, 'Chỉ có một câu?', -1),
        new Paragraph('ĐÁP ÁN PHẦN I'),
        table([
            ['Câu', 'Đáp án'].map(function (t) { return { text: t }; }),
            ['1', 'A'].map(function (t) { return { text: t }; }),
            ['2', 'B'].map(function (t) { return { text: t }; })
        ])
    ]);
}

async function buildUnrecognizedHeading() {
    return docx([
        new Paragraph('PHẦN I. TRẮC NGHIỆM'),
        ...singleQuestion(1, 'Câu chưa có đáp án?', -1),
        new Paragraph('HƯỚNG DẪN CHẤM'),
        table([
            ['Câu', 'Đáp án'].map(function (t) { return { text: t }; }),
            ['1', 'A'].map(function (t) { return { text: t }; })
        ])
    ]);
}

async function buildNoTable() {
    return docx([
        new Paragraph('Câu 1. Thủ đô Việt Nam?'),
        new Paragraph('A. TP. Hồ Chí Minh'),
        new Paragraph('B. Hà Nội *'),
        new Paragraph('C. Đà Nẵng'),
        new Paragraph('D. Cần Thơ')
    ]);
}

async function buildHorizontalParts() {
    return docx([
        new Paragraph('PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI'),
        new Paragraph('Câu 1. Xét các phát biểu sau:'),
        new Paragraph('a) Mệnh đề a.'), new Paragraph('b) Mệnh đề b.'),
        new Paragraph('c) Mệnh đề c.'), new Paragraph('d) Mệnh đề d.'),
        new Paragraph('ĐÁP ÁN TRẮC NGHIỆM ĐÚNG SAI'),
        table([
            ['Câu', 'a', 'b', 'c', 'd'].map(function (t) { return { text: t }; }),
            ['1', 'Đ', 'S', 'Đ', 'S'].map(function (t) { return { text: t }; })
        ])
    ]);
}

async function buildSideGroupTF() {
    return docx([
        new Paragraph('PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI'),
        new Paragraph('Câu 1. Xét các phát biểu sau:'),
        new Paragraph('a) Mệnh đề 1a.'), new Paragraph('b) Mệnh đề 1b.'),
        new Paragraph('c) Mệnh đề 1c.'), new Paragraph('d) Mệnh đề 1d.'),
        new Paragraph('Câu 2. Xét các phát biểu sau:'),
        new Paragraph('a) Mệnh đề 2a.'), new Paragraph('b) Mệnh đề 2b.'),
        new Paragraph('c) Mệnh đề 2c.'), new Paragraph('d) Mệnh đề 2d.'),
        new Paragraph('Câu 3. Xét các phát biểu sau:'),
        new Paragraph('a) Mệnh đề 3a.'), new Paragraph('b) Mệnh đề 3b.'),
        new Paragraph('c) Mệnh đề 3c.'),
        new Paragraph('Câu 4. Xét các phát biểu sau:'),
        new Paragraph('a) Mệnh đề 4a.'), new Paragraph('b) Mệnh đề 4b.'),
        new Paragraph('c) Mệnh đề 4c.'),
        new Paragraph('ĐÁP ÁN TRẮC NGHIỆM ĐÚNG SAI'),
        table([
            ['Câu', 'Lệnh hỏi', 'Đáp án (Đ/S)', 'Câu 2'].map(function (t) { return { text: t }; }),
            ['1', 'a', 'Đ', 'Đ'], ['1', 'b', 'Đ', 'Đ'], ['1', 'c', 'Đ', 'Đ'], ['1', 'd', 'S', 'S'],
            ['3', 'a', 'Đ', ''], ['3', 'b', 'Đ', ''], ['3', 'c', 'S', '']
        ].map(function (row) {
            return row.map(function (t) { return typeof t === 'string' ? { text: t } : t; });
        }))
    ]);
}

async function buildMergedTFAndShort() {
    return docx([
        new Paragraph('PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI'),
        new Paragraph('Câu 1. Xét các phát biểu sau:'),
        new Paragraph('a) Mệnh đề 1a.'), new Paragraph('b) Mệnh đề 1b.'),
        new Paragraph('c) Mệnh đề 1c.'), new Paragraph('d) Mệnh đề 1d.'),
        new Paragraph('Câu 2. Xét các phát biểu sau:'),
        new Paragraph('a) Mệnh đề 2a.'), new Paragraph('b) Mệnh đề 2b.'),
        new Paragraph('c) Mệnh đề 2c.'), new Paragraph('d) Mệnh đề 2d.'),
        new Paragraph('PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN'),
        new Paragraph('Câu 1. Tính 2 + 0.'),
        new Paragraph('Câu 2. Tính 2 + 2.'),
        new Paragraph('ĐÁP ÁN TRẮC NGHIỆM ĐÚNG SAI'),
        table([
            ['Câu', 'Lệnh hỏi', 'Đáp án', 'Câu', 'Đáp án'].map(function (t) { return { text: t }; }),
            [{ text: '1', vMerge: VerticalMergeType.RESTART }, { text: 'a' }, { text: 'Đ' },
                { text: '2', vMerge: VerticalMergeType.RESTART }, { text: 'S' }],
            [{ text: '', vMerge: VerticalMergeType.CONTINUE }, { text: 'b' }, { text: 'S' },
                { text: '', vMerge: VerticalMergeType.CONTINUE }, { text: 'Đ' }],
            [{ text: '', vMerge: VerticalMergeType.CONTINUE }, { text: 'c' }, { text: 'S' },
                { text: '', vMerge: VerticalMergeType.CONTINUE }, { text: 'Đ' }],
            [{ text: '', vMerge: VerticalMergeType.CONTINUE }, { text: 'd' }, { text: 'Đ' },
                { text: '', vMerge: VerticalMergeType.CONTINUE }, { text: 'Đ' }]
        ]),
        new Paragraph('ĐÁP ÁN TRẮC NGHIỆM TRẢ LỜI NGẮN'),
        table([
            ['Câu', 'Đáp án', 'Câu', 'Đáp án'].map(function (t) { return { text: t }; }),
            ['1', '2', '2', '4'].map(function (t) { return { text: t }; })
        ])
    ]);
}

async function buildMatrix() {
    const children = [new Paragraph('PHẦN I. TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN')];
    const letters = ['A', 'B', 'C', 'D'];
    for (let n = 1; n <= 12; n++) {
        children.push(new Paragraph('Câu ' + n + '. Câu hỏi số ' + n + '.'));
        letters.forEach(function (l, i) { children.push(new Paragraph(l + '. Phương án ' + i)); });
    }
    children.push(new Paragraph('PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI'));
    for (let n = 1; n <= 4; n++) {
        children.push(new Paragraph('Câu ' + n + '. Xét các phát biểu sau:'));
        ['a', 'b', 'c', 'd'].forEach(function (p) { children.push(new Paragraph(p + ') Mệnh đề ' + n + p + '.')); });
    }
    children.push(new Paragraph('PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN'));
    for (let n = 1; n <= 6; n++) children.push(new Paragraph('Câu ' + n + '. Tính biểu thức ' + n + '.'));

    const rows = [
        [{ text: 'Câu hỏi', vMerge: VerticalMergeType.RESTART }, { text: 'Mã đề thi', span: 8 }],
        [{ text: '', vMerge: VerticalMergeType.CONTINUE }].concat(['111', '112', '113', '114', '115', '116', '117', '118'].map(function (t) { return { text: t }; }))
    ];
    for (let n = 1; n <= 22; n++) {
        let value = n <= 12 ? letters[(n - 1) % 4]
            : n <= 16 ? ['ĐĐSS', 'SSĐĐ', 'ĐSĐS', 'SĐSĐ'][n - 13]
            : String(n - 7);
        rows.push([{ text: String(n) }, { text: value }, { text: 'B' }, { text: 'C' }, { text: 'D' },
            { text: 'A' }, { text: 'B' }, { text: 'C' }, { text: 'D' }]);
    }
    children.push(new Paragraph('ĐÁP ÁN'));
    children.push(table(rows));
    return docx(children);
}

async function buildFullFlow() {
    return docx([
        new Paragraph('PHẦN I. TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN'),
        ...singleQuestion(1, '1 + 1 bằng bao nhiêu?', 1),
        ...singleQuestion(2, '2 + 2 bằng bao nhiêu?', 2),
        new Paragraph('PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI'),
        new Paragraph('Câu 1. Xét các phát biểu sau:'),
        new Paragraph('a) Nước có công thức H2O.'),
        new Paragraph('b) Nước không dẫn điện.'),
        new Paragraph('c) Nước đóng băng ở 0 độ C.'),
        new Paragraph('d) Nước có vị ngọt.'),
        new Paragraph('PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN'),
        new Paragraph('Câu 1. Tính 2 + 3.'),
        new Paragraph('Câu 2. Tính 8 x 3.'),
        new Paragraph('ĐÁP ÁN PHẦN I'),
        table([
            ['Câu', '1', '2'].map(function (t) { return { text: t }; }),
            ['Đáp án', 'B', 'C'].map(function (t) { return { text: t }; })
        ]),
        new Paragraph('ĐÁP ÁN PHẦN II'),
        table([
            ['Câu', 'Lệnh hỏi', 'Đáp án'].map(function (t) { return { text: t }; }),
            ['1', 'a', 'Đ'].map(function (t) { return { text: t }; }),
            ['1', 'b', 'Đ'].map(function (t) { return { text: t }; }),
            ['1', 'c', 'S'].map(function (t) { return { text: t }; }),
            ['1', 'd', 'Đ'].map(function (t) { return { text: t }; })
        ]),
        new Paragraph('ĐÁP ÁN PHẦN III'),
        table([
            ['Câu', 'Đáp án', 'Câu', 'Đáp án'].map(function (t) { return { text: t }; }),
            ['1', '3', '2', '24'].map(function (t) { return { text: t }; })
        ])
    ]);
}

// ---------------------------------------------------------------------------

(async function main() {
    await ensureServer();

    console.log('\n=== 1. Bảng trắc nghiệm theo hàng ngang ===');
    const p1 = await preview(await buildHorizontalSingle());
    const q1 = (p1.json && p1.json.questions) || [];
    check('preview thành công', p1.json && p1.json.success === true, JSON.stringify(p1.json).slice(0, 300));
    check('có 2 câu, không có câu nào từ bảng', q1.length === 2, 'count=' + q1.length);
    check('câu 1 nhận B', q1[0] && q1[0].correct === 1, q1[0] && q1[0].correct);
    check('câu 2 nhận C', q1[1] && q1[1].correct === 2, q1[1] && q1[1].correct);
    check('không câu nào cần kiểm tra', q1.every(function (q) { return q.needsReview === false; }));
    check('báo cáo 1 bảng đáp án', tableSummary(p1.json).length === 1, JSON.stringify(tableSummary(p1.json)));
    check('không lẫn tiêu đề bảng vào câu hỏi', q1.every(function (q) { return !/ĐÁP ÁN|Đáp án/.test(q.question); }));

    console.log('\n=== 2. Bảng trắc nghiệm theo hàng dọc ===');
    const p2 = await preview(await buildVerticalSingle());
    const q2 = (p2.json && p2.json.questions) || [];
    check('có 2 câu', q2.length === 2, 'count=' + q2.length);
    check('câu 1 nhận B', q2[0] && q2[0].correct === 1, q2[0] && q2[0].correct);
    check('câu 2 nhận C', q2[1] && q2[1].correct === 2, q2[1] && q2[1].correct);
    check('nhận đúng phần I và số câu', q2[0] && q2[0].importSection === 'Phần I' && q2[0].importNumber === 1, JSON.stringify(q2[0] && { s: q2[0].importSection, n: q2[0].importNumber }));

    console.log('\n=== 3. Bảng Đúng/Sai có ý a, b, c, d ===');
    const p3 = await preview(await buildTrueFalse());
    const q3 = (p3.json && p3.json.questions) || [];
    check('có 1 câu đúng/sai', q3.length === 1 && q3[0].type === 'truefalse', q3[0] && q3[0].type);
    check('ghép đúng Đ/S từng ý', q3[0] && JSON.stringify(q3[0].correctTF) === JSON.stringify([true, false, true, false]), q3[0] && JSON.stringify(q3[0].correctTF));
    check('câu đúng/sai không cần kiểm tra', q3[0] && q3[0].needsReview === false, JSON.stringify(q3[0] && q3[0].reviewReasons));

    console.log('\n=== 4. Bảng Đúng/Sai hai cụm cột + ô gộp ===');
    const p4 = await preview(await buildTrueFalseTwoGroups());
    const q4 = (p4.json && p4.json.questions) || [];
    check('có 2 câu đúng/sai', q4.length === 2 && q4.every(function (q) { return q.type === 'truefalse'; }), 'count=' + q4.length);
    check('câu 1: a Đ, b Đ, c S, d S', q4[0] && JSON.stringify(q4[0].correctTF) === JSON.stringify([true, true, false, false]), q4[0] && JSON.stringify(q4[0].correctTF));
    check('câu 2: a Đ, b S, c Đ, d S', q4[1] && JSON.stringify(q4[1].correctTF) === JSON.stringify([true, false, true, false]), q4[1] && JSON.stringify(q4[1].correctTF));
    check('cả 2 câu không cần kiểm tra', q4.every(function (q) { return q.needsReview === false; }), JSON.stringify(q4.map(function (q) { return q.reviewReasons; })));

    console.log('\n=== 5. Bảng trả lời ngắn hai cặp cột ===');
    const p5 = await preview(await buildShortTwoPairs());
    const q5 = (p5.json && p5.json.questions) || [];
    check('có 2 câu trả lời ngắn', q5.length === 2 && q5.every(function (q) { return q.type === 'shortanswer'; }), 'count=' + q5.length);
    check('câu 1 = 3', q5[0] && q5[0].correctAnswer === '3', q5[0] && q5[0].correctAnswer);
    check('câu 2 = 24', q5[1] && q5[1].correctAnswer === '24', q5[1] && q5[1].correctAnswer);
    check('không câu nào cần kiểm tra', q5.every(function (q) { return q.needsReview === false; }));

    console.log('\n=== 6. Bảng Đúng/Sai hàng ngang (Câu | a | b | c | d) ===');
    const p6 = await preview(await buildHorizontalParts());
    const q6 = (p6.json && p6.json.questions) || [];
    check('ghép đúng Đ/S hàng ngang', q6[0] && JSON.stringify(q6[0].correctTF) === JSON.stringify([true, false, true, false]), q6[0] && JSON.stringify(q6[0].correctTF));

    console.log('\n=== 7. Một câu không có đáp án trong bảng ===');
    const p7 = await preview(await buildMissingAnswer());
    const q7 = (p7.json && p7.json.questions) || [];
    check('câu 1 có đáp án B', q7[0] && q7[0].correct === 1, q7[0] && q7[0].correct);
    check('câu 2 không bị tự gán đáp án', q7[1] && q7[1].correct === null, JSON.stringify(q7[1]));
    check('câu 2 được đánh dấu cần kiểm tra', q7[1] && q7[1].needsReview === true, JSON.stringify(q7[1] && q7[1].reviewReasons));

    console.log('\n=== 8. Ô đáp án trống ===');
    const p8 = await preview(await buildEmptyCell());
    const q8 = (p8.json && p8.json.questions) || [];
    check('câu trả lời ngắn không có đáp án', q8[0] && q8[0].correctAnswer === null, q8[0] && q8[0].correctAnswer);
    check('được đánh dấu cần kiểm tra', q8[0] && q8[0].needsReview === true, JSON.stringify(q8[0] && q8[0].reviewReasons));
    check('báo cáo ô trống', tableSummary(p8.json)[0] && tableSummary(p8.json)[0].warnings.join(' ').indexOf('trống') >= 0, JSON.stringify(tableSummary(p8.json)));

    console.log('\n=== 8b. Đáp án ngắn âm và thập phân ===');
    const p8b = await preview(await buildDecimalShort());
    const q8b = (p8b.json && p8b.json.questions) || [];
    check('giữ đúng số âm/thập phân (chuẩn hóa về dấu chấm)', q8b[0] && q8b[0].correctAnswer === '-3.5', q8b[0] && q8b[0].correctAnswer);
    check('giữ đúng 24.00 thành 24', q8b[1] && q8b[1].correctAnswer === '24', q8b[1] && q8b[1].correctAnswer);

    console.log('\n=== 8c. Hai đáp án mâu thuẫn cho cùng câu ===');
    const p8c = await preview(await buildConflict());
    const q8c = (p8c.json && p8c.json.questions) || [];
    check('câu bị đánh dấu cần kiểm tra', q8c[0] && q8c[0].needsReview === true, JSON.stringify(q8c[0] && q8c[0].reviewReasons));
    check('có cảnh báo mâu thuẫn', q8c[0] && q8c[0].answerWarnings && q8c[0].answerWarnings.join(' ').indexOf('nhiều giá trị') >= 0, JSON.stringify(q8c[0] && q8c[0].answerWarnings));

    console.log('\n=== 8d. Đáp án không hợp lệ (Z) ===');
    const p8d = await preview(await buildInvalidAnswer());
    const q8d = (p8d.json && p8d.json.questions) || [];
    check('không gán đáp án sai', q8d[0] && q8d[0].correct === null, q8d[0] && q8d[0].correct);
    check('đánh dấu cần kiểm tra', q8d[0] && q8d[0].needsReview === true, JSON.stringify(q8d[0] && q8d[0].reviewReasons));

    console.log('\n=== 8e. Đáp án thừa (không có câu tương ứng) ===');
    const p8e = await preview(await buildExtraAnswer());
    check('báo đáp án câu 2 không ghép được', tableSummary(p8e.json)[0] && tableSummary(p8e.json)[0].warnings.join(' ').indexOf('không ghép được') >= 0, JSON.stringify(tableSummary(p8e.json)));

    console.log('\n=== 8f. Bảng giống đáp án nhưng sai tiêu đề ===');
    const p8f = await preview(await buildUnrecognizedHeading());
    const q8f = (p8f.json && p8f.json.questions) || [];
    check('không tự ghép khi tiêu đề không phải ĐÁP ÁN', q8f[0] && q8f[0].correct === null, q8f[0] && q8f[0].correct);
    check('có cảnh báo tiêu đề', p8f.json && p8f.json.answerTables && p8f.json.answerTables.warnings.join(' ').indexOf('tiêu đề') >= 0, JSON.stringify(p8f.json && p8f.json.answerTables));

    console.log('\n=== 9. File không có bảng đáp án ===');
    const p9 = await preview(await buildNoTable());
    const q9 = (p9.json && p9.json.questions) || [];
    check('vẫn nhận câu hỏi bình thường', q9.length === 1 && q9[0].type === 'single', 'count=' + q9.length);
    check('vẫn nhận đáp án gạch chân/dấu *', q9[0] && q9[0].correct === 1, q9[0] && q9[0].correct);
    check('không có báo cáo bảng', !p9.json.answerTables || tableSummary(p9.json).length === 0, JSON.stringify(p9.json.answerTables));

    console.log('\n=== 10. Bảng đáp án nằm sau đề, không thành câu hỏi ===');
    const p10 = await preview(await buildFullFlow());
    const q10 = (p10.json && p10.json.questions) || [];
    check('đúng 5 câu hỏi (2 + 1 + 2)', q10.length === 5, 'count=' + q10.length);
    check('không có câu nào chứa tiêu đề bảng', q10.every(function (q) { return !/ĐÁP ÁN PHẦN/.test(q.question); }));
    check('đọc 3 bảng đáp án', tableSummary(p10.json).length === 3, JSON.stringify(tableSummary(p10.json)));
    check('các phần đều đánh số từ 1 và ghép đúng phần',
        q10[0] && q10[0].importSection === 'Phần I' && q10[0].importNumber === 1 &&
        q10[2] && q10[2].importSection === 'Phần II' && q10[2].importNumber === 1 &&
        q10[3] && q10[3].importSection === 'Phần III' && q10[3].importNumber === 1,
        JSON.stringify(q10.map(function (q) { return { s: q.importSection, n: q.importNumber, t: q.type }; })));
    check('câu 2 trả lời ngắn nhận 24 (không nhầm với phần I)',
        q10[4] && q10[4].correctAnswer === '24', q10[4] && q10[4].correctAnswer);
    check('không còn câu nào cần kiểm tra', q10.every(function (q) { return q.needsReview === false; }), JSON.stringify(q10.map(function (q) { return q.reviewReasons; })));

    console.log('\n=== 11. Xuất file ZIP từ đề đã ghép bảng đáp án ===');
    const exportRes = await fetch(BASE + '/api/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            sessionId: p10.json.sessionId,
            title: 'Đề đầy đủ',
            questions: q10,
            settings: { mode: 'exam', minutes: 30 }
        })
    });
    const zip = await JSZip.loadAsync(Buffer.from(await exportRes.arrayBuffer()));
    const htmlName = Object.keys(zip.files).filter(function (n) { return /\.html$/i.test(n); })[0];
    const exportText = htmlName ? await zip.files[htmlName].async('string') : '';
    check('xuất file thành công', exportRes.status === 200 && /application\/zip/.test(exportRes.headers.get('content-type') || ''),
        exportRes.status + ' ' + exportRes.headers.get('content-type'));
    check('zip có manifest + service worker + icon', ['manifest.webmanifest', 'sw.js', 'icon-192.png', 'icon-512.png']
        .every(function (n) { return !!zip.files[n]; }), JSON.stringify(Object.keys(zip.files)));
    check('là file HTML đầy đủ', /^<!doctype html>/i.test(exportText) && exportText.indexOf('</html>') > 0);
    check('có đủ 5 câu trong file', exportText.indexOf('"questions"') > 0 && exportText.indexOf('2 + 3') > 0);

    console.log('\n=== 12. Bảng Đúng/Sai có cột "Câu 2" bên phải ===');
    const p12 = await preview(await buildSideGroupTF(), 'cau-2.docx');
    const q12 = (p12.json && p12.json.questions) || [];
    check('có 4 câu đúng/sai', q12.length === 4 && q12.every(function (q) { return q.type === 'truefalse'; }), 'count=' + q12.length);
    check('câu 1: Đ,Đ,Đ,S', q12[0] && JSON.stringify(q12[0].correctTF) === JSON.stringify([true, true, true, false]), q12[0] && JSON.stringify(q12[0].correctTF));
    check('câu 2 lấy từ cột "Câu 2"', q12[1] && JSON.stringify(q12[1].correctTF) === JSON.stringify([true, true, true, false]), q12[1] && JSON.stringify(q12[1].correctTF));
    check('câu 3: Đ,Đ,S', q12[2] && JSON.stringify(q12[2].correctTF) === JSON.stringify([true, true, false]), q12[2] && JSON.stringify(q12[2].correctTF));
    check('câu 4 không bị bịa đáp án', q12[3] && q12[3].correctTF.every(function (v) { return v === null; }), q12[3] && JSON.stringify(q12[3].correctTF));
    check('câu 4 vẫn cần kiểm tra', q12[3] && q12[3].needsReview === true, q12[3] && JSON.stringify(q12[3].reviewReasons));

    console.log('\n=== 13. Ô câu gộp dọc + bảng trả lời ngắn riêng ===');
    const p13 = await preview(await buildMergedTFAndShort(), 'gop-doc.docx');
    const q13 = (p13.json && p13.json.questions) || [];
    check('có 2 đúng/sai + 2 trả lời ngắn', q13.length === 4 &&
        q13[0].type === 'truefalse' && q13[1].type === 'truefalse' &&
        q13[2].type === 'shortanswer' && q13[3].type === 'shortanswer',
        JSON.stringify(q13.map(function (q) { return q.type; })));
    check('câu 1: Đ,S,S,Đ', q13[0] && JSON.stringify(q13[0].correctTF) === JSON.stringify([true, false, false, true]), q13[0] && JSON.stringify(q13[0].correctTF));
    check('câu 2: S,Đ,Đ,Đ', q13[1] && JSON.stringify(q13[1].correctTF) === JSON.stringify([false, true, true, true]), q13[1] && JSON.stringify(q13[1].correctTF));
    check('trả lời ngắn 1 = 2, 2 = 4', q13[2] && q13[2].correctAnswer === '2' && q13[3] && q13[3].correctAnswer === '4',
        JSON.stringify([q13[2] && q13[2].correctAnswer, q13[3] && q13[3].correctAnswer]));
    check('không câu nào cần kiểm tra', q13.every(function (q) { return q.needsReview === false; }), JSON.stringify(q13.map(function (q) { return q.reviewReasons; })));

    console.log('\n=== 14. Bảng ma trận "Câu hỏi | Mã đề thi" (111..118) ===');
    const p14 = await preview(await buildMatrix(), 'ma-tran.docx');
    const q14 = (p14.json && p14.json.questions) || [];
    const m14 = tableSummary(p14.json)[0];
    check('có 22 câu', q14.length === 22, 'count=' + q14.length);
    check('nhận bảng ma trận', m14 && m14.type === 'matrix', JSON.stringify(m14));
    check('mã đề 111: câu 1 A, câu 2 B', q14[0] && q14[0].correct === 0 && q14[1] && q14[1].correct === 1,
        JSON.stringify([q14[0] && q14[0].correct, q14[1] && q14[1].correct]));
    check('mã đề 111: câu 13 Đ,Đ,S,S', q14[12] && JSON.stringify(q14[12].correctTF) === JSON.stringify([true, true, false, false]), q14[12] && JSON.stringify(q14[12].correctTF));
    check('mã đề 111: câu 17 = 10', q14[16] && q14[16].correctAnswer === '10', q14[16] && q14[16].correctAnswer);
    check('cảnh báo có nhiều mã đề và đã dùng 111', p14.json && p14.json.answerTables && p14.json.answerTables.warnings.join(' ').indexOf('8 mã đề') >= 0 && p14.json.answerTables.warnings.join(' ').indexOf('111') >= 0, JSON.stringify(p14.json && p14.json.answerTables && p14.json.answerTables.warnings));
    check('không câu nào cần kiểm tra', q14.every(function (q) { return q.needsReview === false; }), JSON.stringify(q14.map(function (q) { return q.reviewReasons; })));

    console.log('\n----------------------------------------');
    console.log('Kết quả: ' + passed + ' PASS, ' + failed + ' FAIL');
    stopServer();
    process.exit(failed ? 1 : 0);
})().catch(function (e) {
    console.error(e);
    process.exit(1);
});
