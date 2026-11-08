'use strict';
/**
 * lib/question-parser.js
 * ---------------------------------------------------------------------------
 * Phan tich HTML do mammoth xuat ra thanh cau hoi trac nghiem, dung-sai,
 * tra loi ngan. KHONG tu bia dap an: cau thieu dap an duoc danh needsReview.
 */

var answerTables = require('./answer-tables');
var shortAnswer = require('./short-answer');
var sanitizeHtml = require('./question-content').sanitizeHtml;

// ========== PARSE CÂU HỎI TỪ HTML (mammoth output) ==========
function stripHtmlTags(html) {
    return html.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim();
}

function removeCorrectMarkers(html) {
    return html.replace(/\*|\[x\]|\[X\]/g, '')
        .replace(/\(đúng\)|\(dung\)/gi, '')
        // Bo the gach chan: gach chan trong 1 dong dap an duoc coi la dau hieu
        // "dap an dung" (giong dau "*"), nen khong giu lai trong noi dung cau hoi.
        .replace(/<\/?u\b[^>]*>/gi, '')
        .trim();
}

// Dong dap an co bi GACH CHAN khong? (mammoth xuat the <u> khi bat styleMap "u => u")
var UNDERLINE_MARKER_RX = /<u\b[^>]*>[\s\S]*?<\/u>/i;

function hasUnderlineMarker(html) {
    return UNDERLINE_MARKER_RX.test(String(html || ''));
}

// Regex nhận dạng marker "đúng" cho truefalse: (đúng) (dung) (Đ) (D) (d) (đ) (T) (t) (Y) (y)
var TF_TRUE_REGEX = /\(đúng\)|\(dung\)|\(Đ\)|\(đ\)|\(D\)|\(d\)|\(T\)|\(t\)|\(Y\)|\(y\)/i;
// Regex nhận dạng marker "sai" cho truefalse: (sai) (S) (s) (F) (f) (N) (n)
var TF_FALSE_REGEX = /\(sai\)|\(S\)|\(s\)|\(F\)|\(f\)|\(N\)|\(n\)/i;
// Regex xóa tất cả TF markers khỏi HTML
var TF_ALL_MARKERS = /\(đúng\)|\(dung\)|\(Đ\)|\(đ\)|\(D\)|\(d\)|\(T\)|\(t\)|\(Y\)|\(y\)|\(sai\)|\(S\)|\(s\)|\(F\)|\(f\)|\(N\)|\(n\)/gi;

function removeTFMarkers(html) {
    return html.replace(TF_ALL_MARKERS, '').trim();
}

// ---------- Bo tien to ("Câu 80.", "D.") trong HTML cua Word ----------
// Word hay IN DAM / GACH CHAN / to mau cho chinh tien to, nen HTML co dang
//   "<strong>Câu 80.</strong> noi dung..."
//   "<u>D</u><strong>.</strong> noi dung..."
//   "<strong>D. noi dung...</strong>"
// Vi vay khong the dung regex tren chuoi HTML: phai tim tien to tren VAN BAN
// THUAN (bo qua the) roi cat dung vi tri tuong ung trong HTML. Neu khong lam
// vay, "Câu 80." se bi giu lai trong noi dung cau hoi -> trong ban xem truoc
// trong nhu cau 71 bi "long" cau 80 (loi giao vien da bao).

// Ban do: moi ky tu cua van ban thuan -> vi tri tuong ung trong HTML
function mapHtmlText(html) {
    var chars = [];
    var i = 0;
    while (i < html.length) {
        var c = html.charAt(i);
        if (c === '<') {
            var e = html.indexOf('>', i);
            if (e === -1) break;
            i = e + 1;
            continue;
        }
        if (c === '&') {
            var ent = html.slice(i).match(/^&(amp|lt|gt|quot|apos|#39|nbsp);/i);
            if (ent) {
                var name = ent[1].toLowerCase();
                var ch = ' ';
                if (name === 'amp') ch = '&';
                else if (name === 'lt') ch = '<';
                else if (name === 'gt') ch = '>';
                else if (name === 'quot') ch = '"';
                else if (name !== 'nbsp') ch = "'";
                chars.push({ text: ch, start: i, end: i + ent[0].length });
                i += ent[0].length;
                continue;
            }
        }
        chars.push({ text: c, start: i, end: i + 1 });
        i++;
    }
    var text = '';
    for (var k = 0; k < chars.length; k++) text += chars[k].text;
    return { text: text, chars: chars, htmlLen: html.length };
}

var VOID_HTML_TAGS = { br: 1, img: 1, hr: 1, input: 1, wbr: 1, col: 1 };

// Can bang the: bo the dong le, tu dong dong cac the mo con lai.
// Dung sau khi cat tien to (the boc quanh tien to bi cat mat, the dong cua no
// con lai -> phai bo di).
function balanceHtml(html) {
    var out = '';
    var stack = [];
    var re = /<(\/?)([a-zA-Z][a-zA-Z0-9]*)\b([^>]*)>/g;
    var last = 0;
    var m;
    while ((m = re.exec(html)) !== null) {
        out += html.slice(last, m.index);
        last = re.lastIndex;
        var name = m[2].toLowerCase();
        if (VOID_HTML_TAGS[name] || /\/$/.test(String(m[3] || ''))) { out += m[0]; continue; }
        if (m[1] !== '/') { stack.push(name); out += m[0]; continue; }
        var at = stack.lastIndexOf(name);
        if (at === -1) continue;               // the dong khong co the mo -> bo
        for (var k = stack.length - 1; k > at; k--) out += '</' + stack[k] + '>';
        out += '</' + name + '>';
        stack.length = at;
    }
    out += html.slice(last);
    for (var k2 = stack.length - 1; k2 >= 0; k2--) out += '</' + stack[k2] + '>';
    return out;
}

// Bo `n` ky tu dau tien cua van ban thuan (kem cac the boc quanh chung)
function cutHtmlPrefix(html, n) {
    if (n <= 0) return html;
    var map = mapHtmlText(html);
    if (map.chars.length < n) return '';
    return balanceHtml(html.slice(map.chars[n - 1].end));
}

// Giu lai phan HTML truoc vi tri `n` cua van ban thuan (bo duoi "Đáp án: ...").
function cutHtmlSuffix(html, n) {
    if (n <= 0) return '';
    var map = mapHtmlText(html);
    if (n >= map.chars.length) return html;
    return balanceHtml(html.slice(0, map.chars[n].start));
}

// "Câu 5. Tính 2+3. Đáp án: 5" => dap an nam ngay cuoi dong cau hoi.
var INLINE_ANSWER_RE = /\s*(?:Đáp\s*án(?:\s*đúng)?|ĐA|Kết\s*quả|Answer)\s*[:：]\s*(.+?)\s*$/i;

function findInlineAnswer(html) {
    var map = mapHtmlText(html);
    var m = INLINE_ANSWER_RE.exec(map.text);
    if (!m || m.index === 0) return null;
    return { start: m.index, value: shortAnswer.normalize(m[1]), raw: m[1] };
}

function stripOptionPrefix(html) {
    // Bo the gach chan cua Word TRUOC khi bo "A." / "B)" vi gach chan trong 1
    // dong dap an duoc coi la dau hieu "dap an dung" (giong dau "*").
    var s = String(html).replace(/<\/?u\b[^>]*>/gi, '');
    var m = mapHtmlText(s).text.match(/^\s*([A-Za-z])[\.\)]\s*/);
    if (!m) return s.trim();
    return cutHtmlPrefix(s, m[0].length).trim();
}

function stripQuestionPrefix(html) {
    // Loai bo "Câu N:" / "N." o dau, ke ca khi bi in dam/gach chan
    var s = String(html);
    var m = mapHtmlText(s).text.match(/^\s*(C(?:âu|au)\s*\d+\s*[\.:]\s*|\d+\s*[\.:]\s*)/i);
    if (!m) return s.trim();
    return cutHtmlPrefix(s, m[0].length).trim();
}

// Bo cac the dinh dang "trang tri" cua Word (in dam/nghieng/gach chan/mau) khoi
// noi dung cau hoi & dap an: chung chi lam ban ban xem truoc (hien ra chu
// "<strong>") va khong mang thong tin cho de thi. Giu nguyen anh, khoi code,
// chi so tren/duoi va cac token [[IMG:n]] / MATHPLACEHOLDER.
var INLINE_FMT_TAG_RX = /<\/?(?:b|strong|i|em|u|span|font|mark)\b[^>]*>/gi;

function stripInlineFormatting(html) {
    return String(html == null ? '' : html)
        .replace(INLINE_FMT_TAG_RX, '')
        .replace(/&nbsp;/g, ' ')
        .trim();
}

// Chuoi cac the MO dang bao quanh vi tri `pos` (de khi cat 1 dong thanh nhieu
// dap an, moi dap an van giu du the dinh dang cua no, ke ca the gach chan)
function openTagPrefixAt(html, pos) {
    var stack = [];
    var re = /<(\/?)([a-zA-Z][a-zA-Z0-9]*)\b([^>]*)>/g;
    var m;
    while ((m = re.exec(html)) !== null) {
        if (m.index >= pos) break;
        var name = m[2].toLowerCase();
        if (VOID_HTML_TAGS[name] || /\/$/.test(String(m[3] || ''))) continue;
        if (m[1] !== '/') { stack.push(m[0]); continue; }
        for (var k = stack.length - 1; k >= 0; k--) {
            if (new RegExp('^<' + name + '\\b', 'i').test(stack[k])) { stack.length = k; break; }
        }
    }
    return stack.join('');
}

// Tach 1 doan chua NHIEU dap an cung dong. Nhieu file Word (de 2025) trinh bay
// 2 cot bang khoang trang nen 1 doan co ca A, B (hoac ca A, B, C, D):
//   "A. Campuchia      B. Mianma      C. Brunay      D. Malaixia"
//   "A. Cộng đồng ASEAN.B. Hiệp hội Đông Nam Á."
// Neu khong tach, ca dong bi coi la 1 dap an -> cau do bi danh dau "can kiem
// tra" va dap an dung co the bi gan sai cho (gach chan nam o B nhung lai tinh
// cho A). Tra ve null neu doan chi co 1 dap an.
function splitInlineOptions(paraHtml) {
    var map = mapHtmlText(paraHtml);
    var text = map.text;
    var rx = /(^|[\s\u00a0]|[\.\)])([A-Za-z])[\.\)]/g;
    var marks = [];
    var m;
    while ((m = rx.exec(text)) !== null) {
        var after = text.charAt(rx.lastIndex);
        if (after !== '' && !/[\s\u00a0]/.test(after)) continue;   // "A.abc" khong phai nhan dap an
        marks.push({ letter: m[2].toUpperCase(), pos: m.index + m[1].length });
    }
    if (marks.length < 2) return null;
    if (text.slice(0, marks[0].pos).trim() !== '') return null;    // doan phai bat dau bang nhan dap an

    // Chi lay chuoi nhan LIEN TIEP tang dan: A, B, C, D...
    var run = [marks[0]];
    var expect = marks[0].letter.charCodeAt(0) + 1;
    for (var k = 1; k < marks.length; k++) {
        if (marks[k].letter.charCodeAt(0) === expect) { run.push(marks[k]); expect++; }
    }
    if (run.length < 2) return null;

    var segs = [];
    for (var s2 = 0; s2 < run.length; s2++) {
        var fromChar = run[s2].pos;
        var toChar = (s2 + 1 < run.length) ? run[s2 + 1].pos : map.chars.length;
        var htmlStart = map.chars[fromChar].start;
        var htmlEnd = (toChar < map.chars.length) ? map.chars[toChar].start : map.htmlLen;
        segs.push({
            letter: run[s2].letter,
            html: openTagPrefixAt(paraHtml, htmlStart) + paraHtml.slice(htmlStart, htmlEnd)
        });
    }
    return segs;
}

function stripTypeTags(html) {
    return html.replace(/\[MULTI\]/gi, '')
               .replace(/\[TRẢ LỜI NGẮN\]/gi, '')
               .replace(/\[TRA LOI NGAN\]/gi, '')
               .replace(/\[SHORT\]/gi, '')
               .replace(/\[ĐÚNG\/SAI\]/gi, '')
               .replace(/\[DUNGSAI\]/gi, '')
               .replace(/\[DUNG\/SAI\]/gi, '')
               .trim();
}

function buildQuestionObj(type, question, options, correctAnswer, correctList, correctTF, opts, correctAnswerText) {
    opts = opts || {};
    var answerDetected;
    var shortAnswerText = type === 'shortanswer' ? shortAnswer.normalize(correctAnswerText) : null;
    if (type === 'multi') {
        answerDetected = !!(correctList && correctList.length > 0);
    } else if (type === 'truefalse') {
        answerDetected = !!(correctTF && correctTF.length > 0) &&
            correctTF.every(function (v) { return v === true || v === false; });
    } else if (type === 'shortanswer') {
        answerDetected = !!shortAnswerText;
    } else {
        answerDetected = correctAnswer >= 0;
    }

    var obj = {
        type: type,
        question: sanitizeHtml(stripInlineFormatting(question)),
        options: type === 'shortanswer' ? [] : options.map(function(o) { return sanitizeHtml(stripInlineFormatting(o)); }),
        correct: correctAnswer >= 0 ? correctAnswer : 0,
        correctList: type === 'multi' ? correctList : null,
        correctTF: type === 'truefalse' ? correctTF : null,
        correctAnswer: shortAnswerText,
        image: null
    };

    // Chế độ "nhập đề có xem trước": KHÔNG tự gán đáp án A khi chưa nhận diện được,
    // mà đánh dấu để giáo viên kiểm tra.
    if (opts.trackReview) {
        obj.answerDetected = answerDetected;
        if (!answerDetected && type === 'single') obj.correct = -1;
        obj.needsReview = false;
        obj.reviewReasons = [];
    }
    return obj;
}

// Doi 1 cau "single khong co phuong an" thanh tra loi ngan (giu co needsReview
// neu chua co dap an). Dung lai dung logic danh gia cua pushCurrent().
function toShortAnswer(q) {
    q.type = 'shortanswer';
    q.options = [];
    q.correct = -1;
    q.correctList = null;
    q.correctTF = null;
    if (q.correctAnswer == null) q.correctAnswer = '';
    q.answerDetected = !!q.correctAnswer;
    if (Array.isArray(q.reviewReasons)) {
        q.reviewReasons = q.reviewReasons.filter(function (r) {
            return r.indexOf('cần ít nhất 2') < 0 && r.indexOf('Chưa xác định được đáp án đúng') < 0;
        });
        if (!q.answerDetected) q.reviewReasons.push('Chưa có đáp án trả lời ngắn — giáo viên cần nhập.');
        q.needsReview = q.reviewReasons.length > 0;
    }
}

// Chi doi khi TOAN BO cau cua phan khong co phuong an nao: tranh bien mot cau
// trac nghiem bi loi doc dap an thanh cau tra loi ngan.
function promoteOptionlessSections(questions) {
    var groups = {};
    questions.forEach(function (q) {
        var key = (q.importMeta && q.importMeta.sectionIndex != null) ? q.importMeta.sectionIndex : -1;
        (groups[key] = groups[key] || []).push(q);
    });
    Object.keys(groups).forEach(function (key) {
        var group = groups[key];
        var optionless = group.length > 0 && group.every(function (q) {
            return (q.type || 'single') === 'single' && !(q.options && q.options.length);
        });
        if (optionless) group.forEach(toShortAnswer);
    });
}

function parseQuestionsFromHtml(html, opts) {
    opts = opts || {};
    var parsedQuestions = [];

    // Bang dap an o cuoi file: tach khoi HTML de khong thanh cau hoi, va
    // giu lai cau truc bang de ghep dap an sau khi phan tich xong cau hoi.
    var tableResult = answerTables.extract(html);
    html = tableResult.cleanedHtml;

    // Mammoth xuat cac doan dung style Heading thanh <h1>..<h6>. Neu chi tach
    // cau hoi theo <p> thi toan bo cau nam trong heading bi bo qua am tham
    // (de that "Su-Trac nghiem.docx": mat 33/120 cau vi chung o <h1>).
    html = html.replace(/<(\/?)h[1-6]\b[^>]*>/gi, '<$1p>');

    var paragraphs = [];

    // Merge <pre> blocks into the preceding <p> so code stays attached to its question/option
    // Split HTML into tokens: either a <pre>...</pre> block or a <p>...</p> block
    var tokens = [];
    var remaining = html;
    var preRe = /<pre>([\s\S]*?)<\/pre>/i;
    var pRe = /<p[^>]*>([\s\S]*?)<\/p>/i;
    var tableRe = /<table\b[^>]*>[\s\S]*?<\/table>/i;
    while (remaining.length > 0) {
        var preIdx = remaining.search(/<pre>/i);
        var pIdx = remaining.search(/<p[^>]*>/i);
        var tableIdx = remaining.search(/<table\b/i);
        if (preIdx === -1 && pIdx === -1 && tableIdx === -1) break;
        // Bang noi dung (bang bien thien, bang so lieu) phai duoc giu nguyen
        // HTML, khong duoc tach thanh cac <p> ben trong o.
        if (tableIdx !== -1 &&
            (preIdx === -1 || tableIdx < preIdx) &&
            (pIdx === -1 || tableIdx < pIdx)) {
            var tm = remaining.match(tableRe);
            if (!tm) break;
            tokens.push({ type: 'table', content: tm[0], raw: tm[0] });
            remaining = remaining.slice(remaining.indexOf(tm[0]) + tm[0].length);
        } else if (pIdx === -1 || (preIdx !== -1 && preIdx < pIdx)) {
            // next token is a <pre>
            var pm = remaining.match(/<pre>([\s\S]*?)<\/pre>/i);
            if (!pm) break;
            tokens.push({ type: 'pre', content: pm[1], raw: pm[0] });
            remaining = remaining.slice(remaining.indexOf(pm[0]) + pm[0].length);
        } else {
            // next token is a <p>
            var pm2 = remaining.match(/<p[^>]*>([\s\S]*?)<\/p>/i);
            if (!pm2) break;
            tokens.push({ type: 'p', content: pm2[1], raw: pm2[0] });
            remaining = remaining.slice(remaining.indexOf(pm2[0]) + pm2[0].length);
        }
    }

    // Merge each <pre> into the previous <p> token
    var merged = [];
    for (var ti = 0; ti < tokens.length; ti++) {
        if (tokens[ti].type === 'pre') {
            if (merged.length > 0) {
                merged[merged.length - 1].content += '<pre>' + tokens[ti].content + '</pre>';
            }
            // else: orphan <pre> at start, skip
        } else {
            merged.push({ content: tokens[ti].content });
        }
    }

    for (var mi = 0; mi < merged.length; mi++) {
        // For <p> content: replace <br> with space for plain-text detection
        // but keep the full HTML (including any appended <pre>) for storage
        var inner = merged[mi].content.replace(/<br\s*\/?>/gi, ' ').trim();
        if (inner) paragraphs.push(merged[mi].content);
    }

    // Gom các dòng giữa [CODE] và [/CODE] thành <pre> block gắn vào paragraph trước đó
    var mergedParagraphs = [];
    var codeBuffer = [];
    var inCode = false;
    for (var pi = 0; pi < paragraphs.length; pi++) {
        var plainLine = stripHtmlTags(paragraphs[pi]).trim();
        if (plainLine === '[CODE]') {
            inCode = true;
            codeBuffer = [];
        } else if (plainLine === '[/CODE]') {
            inCode = false;
            var preBlock = '<pre>' + codeBuffer.join('\n') + '</pre>';
            if (mergedParagraphs.length > 0) {
                mergedParagraphs[mergedParagraphs.length - 1] += preBlock;
            } else {
                mergedParagraphs.push(preBlock);
            }
            codeBuffer = [];
        } else if (inCode) {
            // Dòng code: lấy plain text (strip tất cả tags kể cả <strong> từ Word bold), rồi escape HTML
            var codeLine = stripHtmlTags(paragraphs[pi])
                .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'")
                .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
            codeBuffer.push(codeLine);
        } else {
            mergedParagraphs.push(paragraphs[pi]);
        }
    }
    paragraphs = mergedParagraphs;

    // Fallback nếu không có thẻ <p>
    if (paragraphs.length === 0) {
        return parseQuestionsFromText(html.replace(/<[^>]+>/g, '\n'), opts);
    }

    var currentQuestion = null;
    var currentOptions = [];
    var currentType = 'single';
    var correctAnswer = -1;
    var correctList = [];
    var correctTF = [];
    var tfMarkerSeen = false;
    var currentReasons = [];
    var currentShortAnswer = '';
    // Dap an ghi ngay cuoi dong cau hoi / o rieng ngay sau cau hoi, nhung chua
    // biet day la tra loi ngan hay trac nghiem cho den khi het cau.
    var inlineAnswer = null;
    var pendingShortAnswer = '';
    var pendingAnswerHtml = null;
    // Loai cau mac dinh theo "Phần" dang mo (de 2025: Phần II = câu đúng/sai)
    var sectionType = null;
    var sectionIndex = -1;
    var currentSection = { key: null, label: '', index: -1 };
    var currentQuestionNumber = null;

    // "Phần II. Câu trắc nghiệm đúng sai. ..." la TIÊU ĐỀ PHẦN, khong phai noi
    // dung cau hoi. Truoc day dong nay bi coi la "đoạn chưa nhận dạng được" =>
    // tao ly do "cần kiểm tra" gia cho cau dung ngay truoc no.
    var SECTION_RE = /^\s*(?:Phần|PHẦN|Phan)\s+[IVXLC0-9]+\s*(?:[\.:\-–—]|\s|$)/i;
    // Tieu de phan khong co "Phần": "CÁC CÂU HỎI TRẢ LỜI NGẮN", "TRẮC NGHIỆM TRẢ LỜI NGẮN"
    var SHORT_SECTION_RE = /^\s*(?:CÁC\s+)?(?:CÂU\s+HỎI\s+|TRẮC\s+NGHIỆM\s+)*TRẢ\s+LỜI\s+NGẮN\s*[\.:]?\s*$/i;

    function sectionTypeOf(text) {
        if (/trả\s+lời\s+ngắn/i.test(text)) return 'shortanswer';
        if (/trắc nghiệm\s+đúng\s*[-\/–—]?\s*sai/i.test(text)) return 'truefalse';
        if (/^\s*PHẦN\s+[IVXLC0-9]+\s*[\.:\-–—]?\s*(?:CÂU\s+)?(?:HỎI\s+)?ĐÚNG\s*[-\/–—]?\s*SAI/i.test(text)) return 'truefalse';
        // "Trong mỗi ý a) b) c) d) ... học sinh chọn đúng hoặc sai" (đề 2025).
        if (/chọn\s+đúng\s+(?:hoặc|hay)\s+sai|đúng\s+(?:hoặc|hay)\s+sai/i.test(text)) return 'truefalse';
        if (/trắc nghiệm\s+nhiều phương án/i.test(text)) return 'single';
        return null;
    }

    // Đẩy câu hỏi đang dựng vào kết quả.
    // Chế độ xem trước (opts.trackReview): giữ cả câu thiếu đáp án/đáp án <2 và
    // gắn cờ "cần kiểm tra" thay vì âm thầm bỏ hoặc tự bịa đáp án.
    function pushCurrent() {
        if (currentQuestion === null) return;

        // Cau khong co phuong an nao + co dap an viet sau cau hoi => tra loi ngan.
        if (currentType === 'single' && currentOptions.length === 0 && (inlineAnswer || pendingShortAnswer)) {
            currentType = 'shortanswer';
            if (inlineAnswer) {
                currentShortAnswer = inlineAnswer.value || '';
                currentQuestion = cutHtmlSuffix(currentQuestion, inlineAnswer.start).trim();
                if (!inlineAnswer.value && opts.trackReview) {
                    currentReasons.push('Đáp án trong file không hợp lệ: "' + String(inlineAnswer.raw).slice(0, 40) + '" — giáo viên cần nhập lại.');
                }
            } else {
                currentShortAnswer = pendingShortAnswer;
                if (pendingAnswerHtml) {
                    var at = currentQuestion.lastIndexOf(pendingAnswerHtml);
                    if (at >= 0) currentQuestion = currentQuestion.slice(0, at).trim();
                }
            }
        }

        if (currentType === 'truefalse') {
            // Chuan hoa correctTF cho khop so dap an (khong de lo hong)
            var nTF = currentOptions.length;
            for (var t = 0; t < nTF; t++) {
                if (typeof correctTF[t] !== 'boolean') correctTF[t] = null;
            }
            correctTF.length = nTF;
            if (!tfMarkerSeen && nTF > 0) {
                for (var t2 = 0; t2 < nTF; t2++) correctTF[t2] = null;
                currentReasons.push('Chưa nhận ra phát biểu nào đúng/sai — giáo viên cần chọn.');
            }
        }

        var isShortAnswer = currentType === 'shortanswer';
        var enoughOptions = isShortAnswer || currentOptions.length >= 2;
        if (!opts.trackReview && !enoughOptions) return;

        var obj = buildQuestionObj(currentType, currentQuestion, currentOptions, correctAnswer, correctList, correctTF, opts, currentShortAnswer);
        obj.importMeta = {
            sectionKey: currentSection.key,
            sectionLabel: currentSection.label,
            sectionIndex: currentSection.index,
            sectionType: sectionType || currentType,
            number: currentQuestionNumber
        };

        if (opts.trackReview) {
            var reasons = currentReasons.slice();
            var qPlain = stripHtmlTags(obj.question).trim();
            var qHasImg = /<img\b/i.test(obj.question);
            if (qPlain === '' && qHasImg) {
                reasons.push('Câu hỏi chỉ có hình ảnh — cần bổ sung nội dung chữ.');
            } else if (qPlain === '') {
                reasons.push('Thiếu nội dung câu hỏi.');
            }
            if (!isShortAnswer && !enoughOptions) {
                reasons.push('Chỉ có ' + currentOptions.length + ' đáp án (cần ít nhất 2).');
            }
            if (isShortAnswer) {
                if (!obj.correctAnswer) reasons.push('Chưa có đáp án trả lời ngắn — giáo viên cần nhập.');
            } else if (!obj.answerDetected) {
                reasons.push('Chưa xác định được đáp án đúng — giáo viên cần chọn.');
            }
            obj.reviewReasons = reasons;
            obj.needsReview = reasons.length > 0;
        }
        parsedQuestions.push(obj);
    }

    // Ap 1 dong dap an (co the la 1 phan sau khi tach dong chua nhieu dap an)
    function applyOption(letterRaw, optParaHtml) {
        var optIndex = String(letterRaw).toUpperCase().charCodeAt(0) - 65;
        if (optIndex < 0 || optIndex > 25) return;
        var optPlain = stripHtmlTags(optParaHtml);
        var optHtml = stripOptionPrefix(optParaHtml).trim();

        if (currentType === 'single') {
            if (optPlain.indexOf('*') >= 0 || /\(đúng\)|\(dung\)/i.test(optPlain) || hasUnderlineMarker(optParaHtml)) {
                correctAnswer = optIndex;
            }
            optHtml = removeCorrectMarkers(optHtml);
        } else if (currentType === 'multi') {
            if (optPlain.indexOf('*') >= 0 || /\(đúng\)|\(dung\)/i.test(optPlain) || hasUnderlineMarker(optParaHtml)) {
                correctList.push(optIndex);
            }
            optHtml = removeCorrectMarkers(optHtml);
        } else if (currentType === 'truefalse') {
            // Gach chan duoi 1 phat bieu = phat bieu DUNG (cung quy uoc voi
            // gach chan chan duoi dap an o phan I cua de)
            var marked = hasUnderlineMarker(optParaHtml) || TF_TRUE_REGEX.test(optPlain);
            correctTF[optIndex] = marked;
            if (marked) tfMarkerSeen = true;
            optHtml = removeTFMarkers(optHtml);
        }

        while (currentOptions.length < optIndex) currentOptions.push('');
        // Khong de 1 dong dap an RONG ghi de len noi dung da co
        if (optHtml !== '' || !currentOptions[optIndex]) currentOptions[optIndex] = optHtml;
    }

    for (var i = 0; i < paragraphs.length; i++) {
        var paraHtml = paragraphs[i];
        var plainText = stripHtmlTags(paraHtml);

        // Bang noi dung trong de (bang bien thien, bang so lieu): giu nguyen
        // HTML de file xuat ra hien dung cau truc thay vi noi thanh chuoi.
        if (/^\s*<table\b/i.test(paraHtml)) {
            if (currentQuestion !== null) {
                currentQuestion += ' ' + paraHtml;
                if (opts.trackReview && currentOptions.length > 0) {
                    currentReasons.push('Có bảng nằm sau các đáp án — kiểm tra vị trí.');
                }
            }
            continue;
        }

        // Dong ket thuc de thi: phan sau chi la huong dan/đáp án, khong phai
        // noi dung cau hoi (file de that hay de "MÃ ĐỀ..." va bang đáp án ở đây).
        if (/^(?:[-–—]+\s*)?HẾT\s*[\.\!\:]?\s*(?:[-–—]+)?$/i.test(plainText.trim())) break;

        // Tieu de phan -> chot cau dang dung, doi loai cau mac dinh cho phan sau
        if (SECTION_RE.test(plainText) || SHORT_SECTION_RE.test(plainText)) {
            pushCurrent();
            currentQuestion = null;
            currentOptions = [];
            correctAnswer = -1;
            correctList = [];
            correctTF = [];
            tfMarkerSeen = false;
            currentReasons = [];
            currentShortAnswer = '';
            inlineAnswer = null;
            pendingShortAnswer = '';
            pendingAnswerHtml = null;
            currentType = 'single';
            sectionType = sectionTypeOf(plainText);
            if (sectionType) currentType = sectionType;

            sectionIndex++;
            var secMatch = plainText.match(/(?:Phần|PHAN|Phan)\s+([IVXLC]+|\d+)/i);
            currentSection = {
                key: secMatch ? answerTables.sectionKey(secMatch[1]) : null,
                label: secMatch ? 'Phần ' + String(secMatch[1]).toUpperCase() : plainText.trim().slice(0, 40),
                index: sectionIndex
            };
            currentQuestionNumber = null;
            continue;
        }

        var questionMatch = plainText.match(/^(C(?:âu|au)\s*\d+[\.:]\s*|\d+[\.:]\s*)(.*)/i);
        if (questionMatch) {
            pushCurrent();
            var numMatch = questionMatch[1].match(/\d+/);
            currentQuestionNumber = numMatch ? parseInt(numMatch[0], 10) : null;
            currentType = sectionType || 'single';
            correctAnswer = -1;
            correctList = [];
            correctTF = [];
            tfMarkerSeen = false;
            currentOptions = [];
            currentReasons = [];
            currentShortAnswer = '';
            inlineAnswer = null;
            pendingShortAnswer = '';
            pendingAnswerHtml = null;

            if (/\[MULTI\]/i.test(plainText)) {
                currentType = 'multi';
            } else if (/\[TRẢ LỜI NGẮN\]|\[TRA LOI NGAN\]|\[SHORT\]/i.test(plainText)) {
                currentType = 'shortanswer';
            } else if (/\[ĐÚNG\/SAI\]|\[DUNGSAI\]|\[DUNG\/SAI\]/i.test(plainText)) {
                currentType = 'truefalse';
            }

            // Lấy inner HTML của câu hỏi, bỏ prefix "Câu N:" và type tags
            currentQuestion = stripQuestionPrefix(stripTypeTags(paraHtml)).trim();

            // "Câu 7. Tính 2+3. Đáp án: 5": dap an o cuoi dong cau hoi.
            if (currentType === 'shortanswer') {
                var inline = findInlineAnswer(currentQuestion);
                if (inline) {
                    currentShortAnswer = inline.value || '';
                    currentQuestion = cutHtmlSuffix(currentQuestion, inline.start).trim();
                    if (!inline.value && opts.trackReview) {
                        currentReasons.push('Đáp án trong file không hợp lệ: "' + String(inline.raw).slice(0, 40) + '" — giáo viên cần nhập lại.');
                    }
                }
            } else if (currentType === 'single') {
                inlineAnswer = findInlineAnswer(currentQuestion);
            }
            continue;
        }

        // Dap an cua cau tra loi ngan: "Đáp án: 3", "Đáp án đúng: 2,5", "Kết quả: 4"
        if (currentQuestion !== null && currentType !== 'truefalse' && currentType !== 'multi') {
            var ansMatch = plainText.match(/^\s*\(?\s*(?:Đáp\s*án(?:\s*đúng)?|ĐA|Kết\s*quả|Answer)\s*[:：]\s*(.+?)\s*\)?\s*$/i);
            if (ansMatch) {
                var detected = shortAnswer.normalize(ansMatch[1]);
                if (currentType === 'shortanswer') {
                    currentShortAnswer = detected || '';
                    if (!detected && opts.trackReview) {
                        currentReasons.push('Đáp án trong file không hợp lệ: "' + String(ansMatch[1]).slice(0, 40) + '" — giáo viên cần nhập lại.');
                    }
                } else if (currentOptions.length === 0 && !pendingShortAnswer) {
                    pendingShortAnswer = detected || '';
                    if (!detected && opts.trackReview) {
                        currentReasons.push('Đáp án trong file không hợp lệ: "' + String(ansMatch[1]).slice(0, 40) + '" — giáo viên cần nhập lại.');
                    }
                }
                continue;
            }
        }

        // Mot dong chi ghi so ngay sau cau hoi (khong co nhan) => dap an tra loi ngan.
        if (currentQuestion !== null && currentType !== 'truefalse' && currentType !== 'multi' &&
            !currentShortAnswer && !pendingShortAnswer &&
            (currentType === 'shortanswer' || currentOptions.length === 0)) {
            var numAns = shortAnswer.normalize(plainText.trim());
            if (numAns) {
                if (currentType === 'shortanswer') {
                    currentShortAnswer = numAns;
                } else {
                    pendingShortAnswer = numAns;
                    pendingAnswerHtml = paraHtml;
                    currentQuestion += ' ' + paraHtml;
                }
                continue;
            }
        }

        var optionMatch = plainText.match(/^([A-Za-z])[\.\)]\s*(.*)/);
        var optSegments = null;
        if (optionMatch && currentQuestion !== null) {
            // 1 doan co the chua nhieu dap an (A. ... B. ... C. ... D. ...)
            optSegments = splitInlineOptions(paraHtml) || [{ letter: optionMatch[1], html: paraHtml }];
        }
        if (optSegments) {
            for (var si = 0; si < optSegments.length; si++) {
                applyOption(optSegments[si].letter, optSegments[si].html);
            }
        } else if (currentQuestion !== null) {
            var tailHasImg = /<img\b/i.test(paraHtml);
            if (currentOptions.length === 0) {
                currentQuestion += ' ' + paraHtml;
            } else if (tailHasImg && plainText.trim() === '') {
                // Đoạn chỉ có hình ảnh nằm sau các đáp án -> gắn vào câu hỏi và đánh dấu vị trí
                currentQuestion += ' ' + paraHtml;
                if (opts.trackReview) {
                    currentReasons.push('Có hình ảnh nằm riêng — kiểm tra xem thuộc câu hỏi hay đáp án nào.');
                }
            } else if (opts.trackReview && plainText.trim() !== '') {
                // Đoạn văn bản không nhận dạng được -> đánh dấu, không tự đoán
                currentReasons.push('Đoạn chưa nhận dạng được: "' + plainText.trim().slice(0, 60) + '"');
            }
        }
    }

    pushCurrent();

    // Co de khong ghi loai cau o tieu de phan (vd "PHẦN III. (1,5 điểm) Học sinh
    // trả lời từ câu 1 đến câu 6."). Phan ma KHONG cau nao co phuong an A/B/C/D
    // chac chan la tra loi ngan, khong phai cau trac nghiem bi thieu dap an.
    promoteOptionlessSections(parsedQuestions);

    // Ghep dap an tu bang cuoi file vao cau hoi da phan tich.
    var tableReport = answerTables.apply(parsedQuestions, tableResult.tables);
    if (opts) {
        opts.answerTableReport = {
            tables: tableReport.tables,
            warnings: tableResult.warnings.concat(tableReport.warnings)
        };
    }

    return parsedQuestions;
}

// ========== KIỂM TRA TÍNH ĐẦY ĐỦ SAU KHI PHÂN TÍCH ==========
// Chỉ BÁO CÁO, không tự tạo câu giả để bù số lượng thiếu.
function analyzeImportCompleteness(questions) {
    var report = {
        detected: questions.length,
        numbers: [],
        missingNumbers: [],
        duplicateNumbers: [],
        emptyQuestions: [],
        missingOptions: [],
        unparsedContent: [],
        imageCount: 0,
        unresolvedImages: [],
        warnings: []
    };

    function labelOf(q, i) {
        var n = q.importMeta && q.importMeta.number;
        return (n === null || n === undefined) ? ('câu #' + (i + 1)) : ('Câu ' + n);
    }

    var groups = {};
    questions.forEach(function (q, i) {
        var meta = q.importMeta || {};
        var label = labelOf(q, i);
        var questionHtml = String(q.question || '');
        var options = Array.isArray(q.options) ? q.options : [];
        var type = q.type || 'single';

        report.numbers.push(meta.number === undefined ? null : meta.number);

        var plain = stripHtmlTags(questionHtml).trim();
        var hasImg = /<img\b/i.test(questionHtml);
        if (plain === '' && !hasImg) {
            report.emptyQuestions.push(label);
        }

        if (type === 'shortanswer') {
            if (!q.correctAnswer) report.missingOptions.push(label + ' (thiếu đáp án trả lời ngắn)');
        } else if (options.length < 2) {
            report.missingOptions.push(label + ' (chỉ có ' + options.length + ' phương án)');
        }

        (q.reviewReasons || []).forEach(function (r) {
            if (/chưa nhận dạng được/i.test(String(r))) {
                report.unparsedContent.push(label + ': ' + String(r));
            }
        });

        var allHtml = questionHtml + ' ' + options.join(' ');
        var imgs = allHtml.match(/<img\b[^>]*>/gi) || [];
        report.imageCount += imgs.length;
        imgs.forEach(function (tag) {
            var src = (tag.match(/\bsrc\s*=\s*["']?([^"'\s>]*)/i) || [])[1] || '';
            if (!/^(data:|https?:)/i.test(src)) {
                report.unresolvedImages.push(label + ' (' + (src ? src.slice(0, 40) : 'không có src') + ')');
            }
        });

        // Đánh số được kiểm tra theo từng phần (đề có thể đánh số lại từ 1 ở Phần II).
        var groupKey = String(meta.sectionIndex);
        if (!groups[groupKey]) groups[groupKey] = [];
        if (typeof meta.number === 'number') groups[groupKey].push(meta.number);
    });

    Object.keys(groups).forEach(function (key) {
        var nums = groups[key].slice().sort(function (a, b) { return a - b; });
        if (nums.length === 0) return;
        var seen = {};
        nums.forEach(function (n) {
            if (seen[n]) {
                if (report.duplicateNumbers.indexOf(n) === -1) report.duplicateNumbers.push(n);
            }
            seen[n] = true;
        });
        for (var n = nums[0]; n < nums[nums.length - 1]; n++) {
            if (!seen[n]) report.missingNumbers.push(n);
        }
    });

    report.warnings.push('Nhận diện được ' + report.detected + ' câu hỏi.');
    if (report.missingNumbers.length) {
        report.warnings.push('Thiếu số thứ tự câu: ' + report.missingNumbers.join(', ') + '.');
    }
    if (report.duplicateNumbers.length) {
        report.warnings.push('Số thứ tự câu bị trùng: ' + report.duplicateNumbers.join(', ') + '.');
    }
    if (report.emptyQuestions.length) {
        report.warnings.push('Câu không có nội dung: ' + report.emptyQuestions.join(', ') + '.');
    }
    if (report.missingOptions.length) {
        report.warnings.push('Câu thiếu phương án/đáp án: ' + report.missingOptions.join(', ') + '.');
    }
    if (report.unparsedContent.length) {
        report.warnings.push('Nội dung chưa phân tích được: ' + report.unparsedContent.length + ' đoạn (xem từng câu trong danh sách cần kiểm tra).');
    }
    if (report.unresolvedImages.length) {
        report.warnings.push('Hình ảnh chưa xử lý được: ' + report.unresolvedImages.join(', ') + '.');
    }
    return report;
}

// Câu trả lời ngắn do file Word không kèm đáp án thì KHÔNG lưu thẳng:
// giáo viên phải nhập đáp án ở bước xem trước (tránh câu chấm sai âm thầm).
function dropUnansweredShortAnswers(list, warnings) {
    return list.filter(function (q, i) {
        if ((q.type || 'single') === 'shortanswer' && !q.correctAnswer) {
            if (warnings) warnings.push('Câu ' + (i + 1) + ': Câu trả lời ngắn chưa có đáp án — đã bỏ qua (dùng "Nhập Word/Ảnh" để nhập đáp án trước khi lưu).');
            return false;
        }
        return true;
    });
}

// Hàm parse câu hỏi từ text
function parseQuestionsFromText(text, opts) {
    opts = opts || {};
    var questions = [];
    var lines = text.split('\n').map(function(l) { return l.trim(); }).filter(function(l) { return l; });
    
    var currentQuestion = null;
    var currentOptions = [];
    var correctAnswer = -1;

    function pushQ() {
        var detected = correctAnswer >= 0;
        var enough = currentOptions.length >= 2;
        var item = {
            question: currentQuestion,
            options: currentOptions,
            correct: detected ? correctAnswer : 0,
            image: null
        };
        if (opts.trackReview) {
            item.type = 'single';
            item.answerDetected = detected;
            if (!detected) item.correct = -1;
            var reasons = [];
            if (!enough) reasons.push('Chỉ có ' + currentOptions.length + ' đáp án (cần ít nhất 2).');
            if (!detected) reasons.push('Chưa xác định được đáp án đúng — giáo viên cần chọn.');
            item.reviewReasons = reasons;
            item.needsReview = reasons.length > 0;
        }
        questions.push(item);
    }
    
    for (var i = 0; i < lines.length; i++) {
        var line = lines[i];
        
        // Kiểm tra nếu là câu hỏi (bắt đầu bằng "Câu X:" hoặc "Câu X." hoặc số)
        var questionMatch = line.match(/^(C(?:âu|au)\s*\d+[\.:]\s*|^\d+[\.:]\s*)(.*)/i);
        
        if (questionMatch) {
            // Lưu câu hỏi trước đó
            if (currentQuestion && (currentOptions.length >= 2 || opts.trackReview)) {
                pushQ();
            }
            
            currentQuestion = questionMatch[2] || '';
            currentOptions = [];
            correctAnswer = -1;
            continue;
        }
        
        // Kiểm tra nếu là đáp án (A. B. C. D. hoặc A) B) C) D))
        var optionMatch = line.match(/^([A-Za-z])[\.\)]\s*(.*)/);
        
        if (optionMatch && currentQuestion) {
            var optionText = optionMatch[2];
            var optionIndex = optionMatch[1].toUpperCase().charCodeAt(0) - 65;
            
            // Kiểm tra đáp án đúng (có dấu * hoặc [x] hoặc (đúng))
            if (optionText.includes('*') || optionText.includes('[x]') || optionText.includes('[X]') || 
                optionText.toLowerCase().includes('(đúng)') || optionText.toLowerCase().includes('(dung)')) {
                correctAnswer = optionIndex;
                optionText = optionText.replace(/\*|\[x\]|\[X\]|\(đúng\)|\(dung\)/gi, '').trim();
            }
            
            // Đảm bảo đủ số option
            while (currentOptions.length < optionIndex) {
                currentOptions.push('');
            }
            currentOptions[optionIndex] = optionText;
        }
        // Nếu không match và đang có câu hỏi, có thể là phần tiếp của câu hỏi
        else if (currentQuestion && currentOptions.length === 0 && !line.match(/^[A-Za-z][\.\)]/)) {
            currentQuestion += ' ' + line;
        }
    }
    
    // Lưu câu hỏi cuối cùng
    if (currentQuestion && (currentOptions.length >= 2 || opts.trackReview)) {
        pushQ();
    }
    
    return questions;
}

module.exports = {
    parseQuestionsFromHtml: parseQuestionsFromHtml,
    parseQuestionsFromText: parseQuestionsFromText,
    analyzeImportCompleteness: analyzeImportCompleteness,
    dropUnansweredShortAnswers: dropUnansweredShortAnswers,
    stripHtmlTags: stripHtmlTags
};
