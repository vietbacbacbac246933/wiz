'use strict';
/**
 * server.js
 * ---------------------------------------------------------------------------
 * Cong cu noi bo toi gian: .docx -> trich xuat cau hoi -> tai file HTML tu chay.
 *
 *   GET  /               giao dien (public/index.html)
 *   POST /api/extract    multipart/form-data, field "file" (.docx, <= 15 MB)
 *   POST /api/export     JSON { sessionId, title, questions, settings } -> file .zip
 *   GET  /api/health     { ok: true }
 */

var express = require('express');
var path = require('path');
var multer = require('multer');

var fileValidate = require('./lib/file-validate');
var importSessions = require('./lib/import-sessions');
var docxImport = require('./lib/docx-import');
var questionParser = require('./lib/question-parser');
var questionNormalize = require('./lib/question-normalize');
var offlineQuiz = require('./lib/offline-quiz');

var PORT = parseInt(process.env.PORT, 10) || 3456;
var app = express();

// File Word chi giu trong RAM; khiem tra lai day du o /api/extract.
var docxUpload = multer({
    storage: multer.memoryStorage(),
    limits: {
        fileSize: fileValidate.FILE_LIMITS.maxDocxBytes,
        files: 1
    },
    fileFilter: function (req, file, cb) {
        if (fileValidate.detectKind(file.originalname, file.mimetype) !== 'docx') {
            var err = new Error('Chỉ nhận file Word (.docx).');
            err.code = 'UNSUPPORTED_TYPE';
            return cb(err);
        }
        cb(null, true);
    }
});

app.disable('x-powered-by');
app.use(express.json({ limit: '50mb' }));
// KaTeX chi de xem truoc cong thuc trong giao dien; file HTML xuat ra da co MathML.
app.use('/katex', express.static(path.join(__dirname, 'node_modules', 'katex', 'dist'), { maxAge: '7d' }));
app.use(express.static(path.join(__dirname, 'public'), { maxAge: '0' }));

app.get('/api/health', function (req, res) {
    res.json({ ok: true });
});

// Buoc 1+2: nhan file Word, phan tich va tra ve ban xem truoc.
app.post('/api/extract', docxUpload.single('file'), function (req, res) {
    var file = req.file;
    if (!file) {
        return res.status(400).json({ success: false, error: 'Chưa chọn file Word nào.' });
    }

    var check = fileValidate.validateUpload(file);
    if (!check.ok) {
        return res.status(400).json({ success: false, error: check.error });
    }
    if (check.kind !== 'docx') {
        return res.status(400).json({ success: false, error: 'Chỉ nhận file Word (.docx).' });
    }

    docxImport.mammothToHtml(file.buffer).then(function (result) {
        var parseOpts = { trackReview: true };
        var questions = questionParser.parseQuestionsFromHtml(result.value, parseOpts);
        var answerTableReport = parseOpts.answerTableReport || null;
        var completeness = questionParser.analyzeImportCompleteness(questions);

        if (completeness.warnings.length > 1) {
            console.warn('[IMPORT] ' + file.originalname + ': ' + completeness.warnings.join(' '));
        }

        // Anh nhung duoc tach thanh token [[IMG:n]]; du lieu anh giu trong phien RAM.
        var pool = questionNormalize.extractImagesToTokens(questions);
        var sessionId = importSessions.create({ pool: pool });
        var needsReview = questions.filter(function (q) { return q.needsReview; }).length;

        res.json({
            success: true,
            sessionId: sessionId,
            count: questions.length,
            needsReview: needsReview,
            warnings: completeness.warnings,
            answerTables: answerTableReport,
            completeness: completeness,
            imagePool: pool,
            source: {
                docx: file.originalname,
                images: []
            },
            questions: questions.map(function (q, idx) {
                var detected = !!q.answerDetected;
                return {
                    index: idx,
                    type: q.type || 'single',
                    question: q.question || '',
                    options: q.options || [],
                    // Chua xac dinh duoc dap an -> null (KHONG tu gan dap an A).
                    correct: (detected && typeof q.correct === 'number' && q.correct >= 0) ? q.correct : null,
                    correctList: (q.correctList && q.correctList.length) ? q.correctList : null,
                    correctTF: q.correctTF || null,
                    correctAnswer: q.type === 'shortanswer' ? (q.correctAnswer || null) : null,
                    points: questionNormalize.normalizePoints(q.points),
                    needsReview: !!q.needsReview,
                    reviewReasons: q.reviewReasons || [],
                    answerDetected: detected,
                    importSection: q.importMeta ? q.importMeta.sectionLabel : null,
                    importNumber: q.importMeta ? q.importMeta.number : null,
                    answerSource: q.answerSource || null,
                    answerTableNotes: q.answerTableNotes || null,
                    answerWarnings: q.answerWarnings || null
                };
            })
        });
    }).catch(function (err) {
        console.error('Lỗi đọc file:', err);
        res.status(400).json({
            success: false,
            error: 'Không thể đọc file: ' + (err && err.message ? err.message : 'lỗi không xác định')
        });
    });
});

// Buoc 3: validate lai toan bo cau hoi roi tra ve MOT file ZIP (HTML + PWA).
app.post('/api/export', function (req, res) {
    var body = req.body || {};
    var session = importSessions.get(body.sessionId);
    if (!session) {
        return res.status(400).json({ error: 'Phiên nhập đề đã hết hạn. Vui lòng chọn lại file Word.' });
    }

    var incoming = Array.isArray(body.questions) ? body.questions : [];
    if (!incoming.length) {
        return res.status(400).json({ error: 'Chưa chọn câu hỏi nào để xuất file.' });
    }

    var prepared = questionNormalize.prepareIncomingQuestions(session, incoming);
    if (prepared.skipped.length) {
        return res.status(400).json({
            error: 'Có ' + prepared.skipped.length + ' câu chưa hợp lệ, chưa thể xuất file.',
            skipped: prepared.skipped
        });
    }

    var title = String(body.title == null ? '' : body.title).replace(/[<>]/g, '').trim().slice(0, 100) || 'Bài kiểm tra';

    offlineQuiz.renderZip({
        title: title,
        questions: prepared.saved,
        settings: body.settings
    }).then(function (pkg) {
        res.setHeader('Content-Type', 'application/zip');
        res.setHeader('Content-Disposition', 'attachment; filename="' + pkg.fileName + '"');
        res.send(pkg.buffer);
    }).catch(function (err) {
        console.error('Lỗi xuất file ZIP:', err);
        res.status(400).json({ error: err && err.message ? err.message : 'Không tạo được file ZIP.' });
    });
});

// Loi upload (multer) -> JSON gon thay vi trang HTML mac dinh.
app.use(function (err, req, res, next) {
    if (err && (err.code === 'LIMIT_FILE_SIZE' || err.code === 'LIMIT_FILE_COUNT' ||
        err.code === 'LIMIT_UNEXPECTED_FILE' || err.code === 'UNSUPPORTED_TYPE')) {
        var msg = err.code === 'UNSUPPORTED_TYPE'
            ? err.message
            : 'File quá lớn hoặc không hợp lệ (tối đa 15 MB).';
        return res.status(400).json({ success: false, error: msg });
    }
    next(err);
});

app.listen(PORT, '0.0.0.0', function () {
    console.log('Server đang chạy: http://localhost:' + PORT);
});
