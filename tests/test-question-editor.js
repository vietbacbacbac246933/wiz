// Kiem thu quy tac diem, cham diem, sanitize HTML, kiem tra file tai len.
// Chay: node tests/test-question-editor.js
var assert = require('assert');
var qc = require('../lib/question-content');
var fileValidate = require('../lib/file-validate');

var norm = qc.normalizePoints;
var parse = qc.parsePointsInput;
var score = qc.scoreQuestion;
var sanitize = qc.sanitizeHtml;

// ---- Diem cau hoi ----
assert.strictEqual(norm(undefined), 1);
assert.strictEqual(norm(null), 1);
assert.strictEqual(norm(0), 1);
assert.strictEqual(norm(-3), 1);
assert.strictEqual(norm('2.345'), 2.35);
assert.strictEqual(norm(150), 100);
assert.strictEqual(parse('2.5'), 2.5);
assert.strictEqual(parse('2,5'), null); // API JSON dung dau cham
assert.strictEqual(parse('0'), null);
assert.strictEqual(parse('-1'), null);
assert.strictEqual(parse('abc'), null);
assert.strictEqual(parse('101'), null);
assert.strictEqual(parse(''), null);

// ---- Cham diem (0..1, trong so nhan sau) ----
assert.strictEqual(score({ type: 'single', correct: 1 }, 1), 1);
assert.strictEqual(score({ type: 'single', correct: 1 }, 0), 0);
assert.strictEqual(score({ type: 'multi', correctList: [0, 2] }, [2, 0]), 1);
assert.strictEqual(score({ type: 'multi', correctList: [0, 2] }, [0]), 0);
assert.strictEqual(score({ type: 'truefalse', correctTF: [true, false, true] }, [true, true, true]), 2 / 3);
assert.strictEqual(score({ type: 'truefalse', correctTF: [true, false] }, [true, false]), 1);

// ---- Thang điểm câu Đúng/Sai 4 ý ----
var tf4 = [true, false, true, false];
assert.strictEqual(score({ type: 'truefalse', correctTF: tf4 }, [true, false, true, false]), 1);
assert.strictEqual(score({ type: 'truefalse', correctTF: tf4 }, [true, false, true, true]), 0.5);
assert.strictEqual(score({ type: 'truefalse', correctTF: tf4 }, [true, false, false, true]), 0.25);
assert.strictEqual(score({ type: 'truefalse', correctTF: tf4 }, [true, true, false, true]), 0.1);
assert.strictEqual(score({ type: 'truefalse', correctTF: tf4 }, [false, true, false, true]), 0);
assert.strictEqual(score({ type: 'truefalse', correctTF: tf4, tfScoring: 'equal' }, [true, false, true, true]), 0.75);
assert.strictEqual(score({ type: 'truefalse', correctTF: tf4, tfScoring: 'equal' }, [true, false, false, true]), 0.5);
assert.strictEqual(score({ type: 'truefalse', correctTF: tf4, tfScoring: 'equal' }, [true, true, false, true]), 0.25);
assert.strictEqual(score({ type: 'shortanswer', correctAnswer: '2.5' }, '2,5'), 1);
assert.strictEqual(score({ type: 'shortanswer', correctAnswer: '2.5' }, '3'), 0);

// ---- Sanitize HTML: giu cong thuc, cho phep anh data URI, chan XSS ----
var latex = 'Phân số \\(\\frac{a+b}{c}\\) và hệ \\[\\begin{cases}x+y=3\\\\x-y=1\\end{cases}\\]';
var clean = sanitize(latex);
assert.ok(clean.indexOf('\\(\\frac{a+b}{c}\\)') >= 0);
assert.ok(clean.indexOf('\\begin{cases}') >= 0);

assert.ok(sanitize('<img src="data:image/png;base64,AAAA" alt="a">').indexOf('data:image/png') >= 0);
assert.strictEqual(sanitize('<img src="https://evil.example/x.png">').indexOf('evil.example'), -1);
assert.strictEqual(sanitize('<script>alert(1)</script><b onclick="x()">ok</b>'), '<b>ok</b>');
assert.strictEqual(sanitize('<div>dòng 1</div><div>dòng 2</div>'), '<div>dòng 1</div><div>dòng 2</div>');
assert.strictEqual(sanitize('<img src="data:image/png;base64,AAAA" onerror="alert(1)">').indexOf('onerror'), -1);

// ---- Kiem tra file anh ----
var png = Buffer.alloc(32);
[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A].forEach(function (b, i) { png[i] = b; });
var okPng = fileValidate.validateUpload({ originalname: 'a.png', mimetype: 'image/png', buffer: png });
assert.strictEqual(okPng.ok, true);
assert.strictEqual(okPng.imageType, 'image/png');

var fake = fileValidate.validateUpload({
    originalname: 'a.png', mimetype: 'image/png', buffer: Buffer.from('not an image')
});
assert.strictEqual(fake.ok, false);

var huge = Buffer.alloc(fileValidate.FILE_LIMITS.maxImageBytes + 1);
[0x89, 0x50, 0x4E, 0x47].forEach(function (b, i) { huge[i] = b; });
var tooBig = fileValidate.validateUpload({ originalname: 'a.png', mimetype: 'image/png', buffer: huge });
assert.strictEqual(tooBig.ok, false);

console.log('PASS: question-editor tests');
