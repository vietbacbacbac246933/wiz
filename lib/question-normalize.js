'use strict';
/**
 * lib/question-normalize.js
 * ---------------------------------------------------------------------------
 * Chuan hoa cau hoi client gui ve khi xuat file: thay token anh bang data URI
 * cua phien, sanitize HTML, validate tung cau.
 */

var shortAnswer = require('./short-answer');
var questionContent = require('./question-content');
var normalizePoints = questionContent.normalizePoints;
var parsePointsInput = questionContent.parsePointsInput;
var sanitizeHtml = questionContent.sanitizeHtml;

var IMG_TOKEN_RE = /\[\[IMG:(\d+)\]\]/g;

// Tách mọi <img src="data:..."> trong câu hỏi/đáp án thành token [[IMG:n]]
function extractImagesToTokens(questions) {
    var pool = {};   // token -> dataUri
    var order = [];  // thứ tự token để tái sử dụng ảnh trùng
    var seq = 0;

    function walk(html) {
        if (!html || typeof html !== 'string') return html;
        return html.replace(/<img\b[^>]*?src\s*=\s*("([^"]*)"|'([^']*)')[^>]*>/gi, function (full, _q, dq, sq) {
            var src = dq !== undefined ? dq : sq;
            if (!/^data:image\//i.test(src)) return '';
            for (var i = 0; i < order.length; i++) {
                if (pool[order[i]] === src) return '[[IMG:' + order[i] + ']]';
            }
            var token = String(seq++);
            pool[token] = src;
            order.push(token);
            return '[[IMG:' + token + ']]';
        });
    }

    questions.forEach(function (q) {
        q.question = walk(q.question);
        if (q.options) q.options = q.options.map(walk);
    });
    return pool;
}

// Thay token [[IMG:n]] bằng lại ảnh gốc từ pool của phiên (không tin dữ liệu client)
function expandTokens(html, pool) {
    if (!html || typeof html !== 'string') return html || '';
    return html.replace(IMG_TOKEN_RE, function (m, n) {
        return Object.prototype.hasOwnProperty.call(pool, n) ? '<img src="' + pool[n] + '" />' : '';
    });
}

// Chuẩn hoá 1 câu hỏi do client gửi về (đã qua chỉnh sửa ở bản xem trước)
function normalizeEditedQuestion(raw, qNum) {
    if (!raw || typeof raw !== 'object') return { error: 'Câu ' + qNum + ': dữ liệu không hợp lệ' };
    var type = raw.type;
    if (type !== 'single' && type !== 'multi' && type !== 'truefalse' && type !== 'shortanswer') {
        return { error: 'Câu ' + qNum + ': loại câu hỏi không hợp lệ' };
    }
    var out = {
        type: type,
        question: String(raw.question == null ? '' : raw.question),
        options: Array.isArray(raw.options) ? raw.options.map(function (o) { return String(o == null ? '' : o); }) : [],
        correct: 0,
        correctList: null,
        correctTF: null,
        correctAnswer: null,
        image: null,
        points: normalizePoints(raw.points)
    };
    if (type === 'single') {
        out.correct = (typeof raw.correct === 'number') ? raw.correct : -1;
    } else if (type === 'multi') {
        out.correctList = Array.isArray(raw.correctList)
            ? raw.correctList.filter(function (v) { return typeof v === 'number'; })
            : [];
        out.correct = out.correctList.length ? out.correctList[0] : -1;
    } else if (type === 'truefalse') {
        out.correctTF = Array.isArray(raw.correctTF)
            ? raw.correctTF.map(function (v) { return v === true ? true : (v === false ? false : null); })
            : [];
    } else {
        out.options = [];
        out.correctAnswer = shortAnswer.normalize(raw.correctAnswer);
    }
    return { question: out };
}

// Chuan hoa danh sach cau hoi client gui ve (dung chung cho luu va xuat offline):
// thay token anh, sanitize HTML, validate tung cau. Cau chua hop le vao `skipped`.
function prepareIncomingQuestions(session, incoming) {
    var pool = session.pool || {};
    var saved = [];
    var skipped = [];
    var excluded = 0;

    incoming.forEach(function (raw, i) {
        var num = i + 1;
        if (raw && raw.include === false) { excluded++; return; }

        var norm = normalizeEditedQuestion(raw, num);
        if (norm.error) { skipped.push({ index: i, reason: norm.error }); return; }
        var q = norm.question;

        // Thay token ảnh -> data URI gốc (chỉ lấy từ pool của phiên)
        q.question = expandTokens(q.question, pool);
        q.options = q.options.map(function (o) { return expandTokens(o, pool); });

        // Làm sạch HTML
        q.question = sanitizeHtml(q.question);
        q.options = q.options.map(function (o) { return sanitizeHtml(o); });

        // Đúng/Sai: bắt buộc mọi phát biểu đã được chọn (không lưu dữ liệu bỏ trống)
        if (q.type === 'truefalse') {
            var undecided = q.correctTF.filter(function (v) { return v !== true && v !== false; }).length;
            if (undecided > 0) {
                skipped.push({ index: i, reason: 'Câu ' + num + ': còn ' + undecided + ' phát biểu chưa chọn Đúng/Sai' });
                return;
            }
        }

        var errs = [];
        var normalized = validateAndNormalizeQuestion(q, num, errs);
        if (!normalized) { skipped.push({ index: i, reason: errs.join('; ') }); return; }
        saved.push(normalized);
    });

    return { saved: saved, skipped: skipped, excluded: excluded };
}

function validateAndNormalizeQuestion(q, qNum, errors) {
    if (!q.question || q.question.trim() === '') {
        errors.push('Câu ' + qNum + ': Thiếu nội dung câu hỏi');
        return null;
    }

    var type = q.type || 'single';
    var points = 1;
    if (q.points !== undefined && q.points !== null) {
        points = parsePointsInput(q.points);
        if (points === null) {
            errors.push('Câu ' + qNum + ': Điểm phải là số lớn hơn 0 và không quá 100');
            return null;
        }
    }

    if (type === 'shortanswer') {
        var answer = shortAnswer.normalize(q.correctAnswer);
        if (!answer) {
            errors.push('Câu ' + qNum + ': Đáp án trả lời ngắn không hợp lệ hoặc để trống');
            return null;
        }
        return {
            type: 'shortanswer',
            question: String(q.question).trim(),
            options: [],
            correct: 0,
            correctList: null,
            correctTF: null,
            correctAnswer: answer,
            image: q.image || null,
            points: points
        };
    }

    if (!q.options || !Array.isArray(q.options) || q.options.length < 2) {
        errors.push('Câu ' + qNum + ': Thiếu hoặc không đủ đáp án (cần ít nhất 2)');
        return null;
    }

    if (type === 'single') {
        if (typeof q.correct !== 'number' || q.correct < 0 || q.correct >= q.options.length) {
            errors.push('Câu ' + qNum + ': Đáp án đúng không hợp lệ');
            return null;
        }
        return {
            type: 'single',
            question: String(q.question).trim(),
            options: q.options.map(function(o) { return String(o).trim(); }),
            correct: q.correct,
            correctList: null,
            correctTF: null,
            image: q.image || null,
            points: points
        };
    }

    if (type === 'multi') {
        if (!Array.isArray(q.correctList) || q.correctList.length === 0) {
            errors.push('Câu ' + qNum + ': Loại nhiều đáp án cần có correctList');
            return null;
        }
        for (var i = 0; i < q.correctList.length; i++) {
            if (typeof q.correctList[i] !== 'number' || q.correctList[i] < 0 || q.correctList[i] >= q.options.length) {
                errors.push('Câu ' + qNum + ': correctList chứa chỉ số không hợp lệ');
                return null;
            }
        }
        return {
            type: 'multi',
            question: String(q.question).trim(),
            options: q.options.map(function(o) { return String(o).trim(); }),
            correct: q.correctList[0],
            correctList: q.correctList,
            correctTF: null,
            image: q.image || null,
            points: points
        };
    }

    if (type === 'truefalse') {
        if (!Array.isArray(q.correctTF) || q.correctTF.length !== q.options.length) {
            errors.push('Câu ' + qNum + ': Loại đúng/sai cần correctTF cùng độ dài options');
            return null;
        }
        return {
            type: 'truefalse',
            question: String(q.question).trim(),
            options: q.options.map(function(o) { return String(o).trim(); }),
            correct: 0,
            correctList: null,
            correctTF: q.correctTF,
            image: q.image || null,
            points: points
        };
    }

    errors.push('Câu ' + qNum + ': Loại câu hỏi không hợp lệ (' + type + ')');
    return null;
}

module.exports = {
    IMG_TOKEN_RE: IMG_TOKEN_RE,
    extractImagesToTokens: extractImagesToTokens,
    expandTokens: expandTokens,
    normalizeEditedQuestion: normalizeEditedQuestion,
    prepareIncomingQuestions: prepareIncomingQuestions,
    validateAndNormalizeQuestion: validateAndNormalizeQuestion,
    normalizePoints: normalizePoints
};
