'use strict';
/**
 * DEMO read-only: truy vet va tinh lai diem on tap bi 0.
 * Khong ghi vao data/, khong sua attempt nao. Chi in bang + (tuy chon) file report.
 *
 *   node scripts/recover-review-scores-demo.js --demo
 *   node scripts/recover-review-scores-demo.js --exam data/exams/<id>.json \
 *        --attempts data/review-attempts/<list>__<exam>.json [--exports <dir>] [--out report.json]
 *
 * --exports: thu muc JSON hoc sinh xuat tu localStorage (snippet in o cuoi --demo).
 * Nguon phuc hoi: (1) attempt.details da luu cau tra loi, (2) file export cua hoc sinh.
 * Khong co ca hai thi KHONG the tinh lai tu server vi dap an da bi xoa.
 */

var fs = require('fs');
var path = require('path');
var grading = require('../lib/grading');

function readJson(file) {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
}

// Giong submitExam() trong public/index.html: dua answers tu thu tu dao ve thu tu goc.
function mapExportAnswers(exp, questions) {
    var answers = exp.answers || {};
    var qOrder = exp.questionOrder || [];
    var optOrders = exp.optionOrders || [];
    var out = [];
    for (var shuffled = 0; shuffled < qOrder.length; shuffled++) {
        if (answers[shuffled] === undefined) continue;
        var orig = qOrder[shuffled];
        var type = (questions[orig] && questions[orig].type) || 'single';
        var optOrder = optOrders[orig];
        if (type === 'multi' && Array.isArray(answers[shuffled])) {
            out[orig] = answers[shuffled].map(function (i) {
                return optOrder ? optOrder[i] : i;
            });
        } else if (type === 'single' && optOrder && typeof answers[shuffled] === 'number') {
            out[orig] = optOrder[answers[shuffled]];
        } else {
            out[orig] = answers[shuffled];
        }
    }
    return out;
}

function answersFromDetails(details) {
    var out = [];
    (details || []).forEach(function (d) {
        if (d && typeof d.index === 'number' && d.studentAnswer !== undefined) {
            out[d.index] = d.studentAnswer;
        }
    });
    return out;
}

function findExport(exports, attempt) {
    for (var i = 0; i < exports.length; i++) {
        var e = exports[i];
        if (e.attemptId && String(e.attemptId) === String(attempt.id)) return e;
        if (!e.attemptId && String(e.stt) === String(attempt.studentSTT)) return e;
    }
    return null;
}

function recover(exam, attempts, exports) {
    var rows = [];
    (attempts || []).forEach(function (a) {
        if (!a || a.status !== 'submitted' || Number(a.score) !== 0) return;
        var source = null;
        var answers = null;
        if (Array.isArray(a.details) && a.details.length) {
            answers = answersFromDetails(a.details);
            source = 'details trong attempt';
        } else {
            var exp = findExport(exports, a);
            if (exp) {
                answers = mapExportAnswers(exp, exam.questions);
                source = 'export localStorage';
            }
        }
        var row = {
            stt: a.studentSTT,
            name: a.studentName,
            attemptId: a.id,
            oldScore: a.score,
            maxScore: a.maxScore,
            source: source || 'không có dữ liệu đáp án',
            newScore: null
        };
        if (answers) row.newScore = grading.grade(exam.questions, answers, { snapshot: true }).score;
        rows.push(row);
    });
    return rows;
}

function printRows(rows) {
    var head = ['STT', 'Họ tên', 'Lượt', 'Điểm cũ', 'Điểm phục hồi', 'Nguồn'];
    var body = rows.map(function (r) {
        return [
            r.stt,
            r.name || '',
            r.attemptId,
            r.oldScore + '/' + r.maxScore,
            r.newScore === null ? '—' : r.newScore + '/' + r.maxScore,
            r.source
        ];
    });
    var widths = head.map(function (h, i) {
        return Math.max(h.length, Math.max.apply(null, body.map(function (b) { return String(b[i]).length; })));
    });
    function line(cols) {
        return cols.map(function (c, i) { return String(c).padEnd(widths[i]); }).join(' | ');
    }
    console.log(line(head));
    console.log(widths.map(function (w) { return '-'.repeat(w); }).join('-+-'));
    body.forEach(function (b) { console.log(line(b)); });
    var ok = rows.filter(function (r) { return r.newScore !== null; });
    if (ok.length) {
        var sum = ok.reduce(function (s, r) { return s + r.newScore; }, 0);
        console.log('\nKhôi phục được ' + ok.length + '/' + rows.length + ' lượt, điểm lại: ' +
            (Math.round(sum * 100) / 100) + ' (chỉ tính trên các lượt có dữ liệu).');
    }
}

// ------------------------------- DEMO --------------------------------------
function demoExam(name) {
    var qs = [];
    for (var i = 0; i < 23; i++) {
        qs.push({
            type: 'single',
            question: 'Câu ' + (i + 1) + ' (' + name + ', demo)',
            options: ['A', 'B', 'C', 'D'],
            correct: i % 4
        });
    }
    return { name: name, questions: qs };
}

function demoExport(stt, exam, attemptId, correctCount, answeredCount) {
    var n = exam.questions.length;
    var questionOrder = [];
    for (var i = 0; i < n; i++) questionOrder.push(n - 1 - i);
    var optionOrders = exam.questions.map(function () { return [1, 2, 3, 0]; });
    var answers = {};
    for (var s = 0; s < n; s++) {
        if (s >= answeredCount) continue;
        var orig = questionOrder[s];
        var optOrder = optionOrders[orig];
        var want = exam.questions[orig].correct;
        var pos = optOrder.indexOf(want);
        answers[s] = s < correctCount ? pos : optOrder[(pos + 1) % optOrder.length];
    }
    return {
        stt: stt, examId: 'exam_demo', attemptId: attemptId,
        answers: answers, questionOrder: questionOrder, optionOrders: optionOrders,
        _expect: correctCount
    };
}

function demoDetails(exam, correctCount) {
    return exam.questions.map(function (q, i) {
        var studentAnswer = i < correctCount ? q.correct : (q.correct + 1) % q.options.length;
        return { index: i, question: q.question, studentAnswer: studentAnswer };
    });
}

// [stt, ten, [[ma luot, nguon ('export' | 'details' | null), so cau dung]]]
var DEMO_SCENARIOS = [
    {
        name: 'Bài 1',
        zeros: [
            ['4', 'Nguyễn Thiên Bảo', [['b1_bao', 'export', 18]]],
            ['5', 'Trần Việt Bắc', [['b1_bac', 'details', 9]]],
            ['8', 'Huỳnh Đức Trí Dũng', [['b1_dung', null, 0]]]
        ]
    },
    {
        name: 'Bài 2',
        zeros: [
            ['4', 'Nguyễn Thiên Bảo', [['b2_bao', 'export', 18]]],
            ['5', 'Trần Việt Bắc', [['b2_bac_1', 'export', 12], ['b2_bac_2', 'details', 9], ['b2_bac_3', null, 0]]],
            ['8', 'Huỳnh Đức Trí Dũng', [['b2_dung', 'details', 20]]]
        ]
    },
    {
        name: 'Bài 3',
        zeros: [
            ['4', 'Nguyễn Thiên Bảo', [['b3_bao', null, 0]]],
            ['5', 'Trần Việt Bắc', [['b3_bac', 'export', 15]]],
            ['8', 'Huỳnh Đức Trí Dũng', [['b3_dung', 'details', 21]]]
        ]
    }
];

function runDemo() {
    var markdown = [
        '# Demo khôi phục điểm ôn tập — chia theo Bài 1, 2, 3',
        '',
        '> Số liệu mô phỏng để trình bày cách phục hồi. Không phải điểm thật và không ghi vào hệ thống.',
        ''
    ];

    console.log('DEMO (không sửa dữ liệu) — khôi phục điểm ôn tập bị 0\n');

    DEMO_SCENARIOS.forEach(function (sc) {
        var exam = demoExam(sc.name);
        var attempts = [];
        var exports = [];
        sc.zeros.forEach(function (z) {
            z[2].forEach(function (spec) {
                var a = {
                    id: spec[0], studentSTT: z[0], studentName: z[1],
                    status: 'submitted', score: 0, maxScore: exam.questions.length,
                    expired: true, details: null
                };
                if (spec[1] === 'details') a.details = demoDetails(exam, spec[2]);
                attempts.push(a);
                if (spec[1] === 'export') exports.push(demoExport(z[0], exam, spec[0], spec[2], 23));
            });
        });

        // Tự kiểm tra: map đáp án từ thứ tự đảo về thứ tự gốc phải ra đúng số câu đúng.
        exports.forEach(function (e) {
            var got = grading.grade(exam.questions, mapExportAnswers(e, exam.questions)).score;
            if (got !== e._expect) throw new Error('Self-check mapping sai: ' + got + ' != ' + e._expect);
        });

        var rows = recover(exam, attempts, exports);
        console.log('=== ' + sc.name + ' ===');
        printRows(rows);
        console.log('');

        markdown.push('## ' + sc.name, '');
        markdown.push('| STT | Họ và tên | Lượt | Điểm cũ | Điểm phục hồi | Nguồn |');
        markdown.push('|---|---|---|---|---|---|');
        rows.forEach(function (r) {
            markdown.push('| ' + [
                r.stt, r.name, r.attemptId,
                r.oldScore + '/' + r.maxScore,
                r.newScore === null ? '—' : r.newScore + '/' + r.maxScore,
                r.source
            ].join(' | ') + ' |');
        });
        var ok = rows.filter(function (r) { return r.newScore !== null; });
        markdown.push('', 'Khôi phục được ' + ok.length + '/' + rows.length + ' lượt bị 0.', '');
    });

    var outFile = path.join(process.cwd(), 'demo-khoi-phuc-diem.md');
    fs.writeFileSync(outFile, markdown.join('\n'), 'utf8');
    console.log('Đã ghi danh sách: ' + outFile);
    console.log('\nDòng "—" là lượt server đã xóa đáp án: cần file export từ máy học sinh mới tính lại được.');
    console.log('\nSnippet lấy dữ liệu trên máy học sinh (DevTools > Console):');
    console.log('  copy(JSON.stringify(JSON.parse(localStorage.getItem("quiz_exam_progress")), null, 2))');
    console.log('  -> dán vào file .json, đặt tên bất kỳ, bỏ chung một thư mục rồi chạy với --exports <thu mục>.');
}

function runReal(args) {
    var examFile = args.exam;
    var attemptsFile = args.attempts;
    if (!examFile || !attemptsFile) {
        console.error('Thiếu --exam hoặc --attempts. Xem hướng dẫn ở đầu file.');
        process.exit(2);
    }
    var exam = readJson(examFile);
    var attempts = readJson(attemptsFile);
    if (!Array.isArray(attempts)) attempts = attempts.attempts || [];
    var exports = [];
    if (args.exports) {
        fs.readdirSync(args.exports)
            .filter(function (f) { return /\.json$/i.test(f); })
            .forEach(function (f) {
                var data = readJson(path.join(args.exports, f));
                // File export nguyen quiz_exam_progress: gom tat ca key thanh tung ban.
                Object.keys(data).forEach(function (k) {
                    var v = data[k];
                    if (v && v.answers) exports.push(v);
                });
            });
    }
    var rows = recover(exam, attempts, exports);
    printRows(rows);
    if (args.out) {
        fs.writeFileSync(args.out, JSON.stringify({ exam: examFile, rows: rows }, null, 2), 'utf8');
        console.log('\nĐã ghi report: ' + args.out + ' (không đụng dữ liệu gốc).');
    }
}

var argv = process.argv.slice(2);
if (argv.indexOf('--demo') !== -1) {
    runDemo();
} else {
    var args = {};
    for (var i = 0; i < argv.length; i++) {
        if (argv[i].indexOf('--') === 0) {
            args[argv[i].slice(2)] = argv[i + 1];
            i++;
        }
    }
    runReal(args);
}
