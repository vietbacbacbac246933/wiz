'use strict';
/**
 * lib/pdf-import.js
 * ---------------------------------------------------------------------------
 * Chuyen file .pdf (lop van ban) thanh HTML <p> giong dau ra cua mammoth, de
 * dung chung questionParser.parseQuestionsFromHtml voi file Word.
 *
 * Gioi han da biet: PDF khong giu duoc dinh dang (gach chan dap an, cong thuc
 * OMML, anh nhung) nen chi doc duoc chu. Cau thieu dau hieu dap an se duoc
 * danh dau "can kiem tra" — khong tu doan.
 */

// Require thang file lib de tranh doan debug cua pdf-parse/index.js: khi
// server.js la module chinh, no tu doc file test khong ton tai va crash.
var pdfParse = require('pdf-parse/lib/pdf-parse.js');

function escapeHtml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// Moi dong van ban -> 1 <p>. Dong bi ngat giua cau van duoc parser noi vao cau
// hoi; dong ngat giua dap an se bi danh dau "can kiem tra".
function textToHtml(text) {
    return String(text)
        .replace(/\r\n?/g, '\n')
        .replace(/[\u00AD\u200B-\u200D\uFEFF]/g, '')
        .split('\n')
        .map(function (line) { return line.trim(); })
        .filter(function (line) { return line !== ''; })
        .map(function (line) { return '<p>' + escapeHtml(line) + '</p>'; })
        .join('\n');
}

/**
 * pdfToHtml(buffer) -> Promise<{ value: html, pages: number, warnings: string[] }>
 * Nem loi co thong bao ro rang khi PDF la ban scan hoac bi khoa.
 */
function pdfToHtml(buffer) {
    // Phai copy sang Uint8Array moi: Buffer nho lay tu pool 8KB nen `.buffer`
    // chua ca pool va pdf.js doc sai offset xref -> "bad XRef entry".
    var bytes = new Uint8Array(buffer.length);
    bytes.set(buffer);
    return pdfParse(bytes.buffer).then(function (data) {
        var text = data && data.text ? data.text : '';
        if (text.replace(/\s+/g, '').length < 10) {
            throw new Error('PDF không có nội dung chữ (có thể là bản scan/ảnh hoặc file bị khóa). ' +
                'Hãy dùng file Word, hoặc chụp ảnh đề và nhập từ ảnh.');
        }
        return {
            value: textToHtml(text),
            pages: data.numpages || 0,
            warnings: ['PDF chỉ đọc được phần chữ: mất định dạng gạch chân, công thức toán và hình ảnh. ' +
                'Câu nào không thấy dấu đáp án (*, (đúng)) sẽ được đánh dấu "cần kiểm tra".']
        };
    });
}

module.exports = {
    pdfToHtml: pdfToHtml
};
