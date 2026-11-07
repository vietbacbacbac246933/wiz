'use strict';
/**
 * lib/docx-import.js
 * ---------------------------------------------------------------------------
 * Chuyen file .docx thanh HTML: mammoth, OMML -> LaTeX, anh nhung -> data URI.
 */

var mammoth = require('mammoth');
var JSZip = require('jszip');
var xmljs = require('xml-js');

// ========== OMML -> LATEX CONVERTER ==========
// Escape ky tu dac biet cua LaTeX trong phan van ban cua cong thuc.
// Chu co dau (vi du "nghiem") phai nam trong \text{} de KaTeX hien dung dau.
function ommlTextToLatex(text) {
    var s = String(text == null ? '' : text);
    if (!s) return '';
    var hasNonAscii = /[^\x00-\x7F]/.test(s);
    // Chu co dau (Latin mo rong / tieng Viet) phai vao \text{}; ky hieu Hy Lap
    // va toan hoc (α, β, Δ, ±, ∑...) thi giu nguyen trong math mode.
    var hasTextLetters = /[\u00C0-\u024F\u1E00-\u1EFF]/.test(s);
    if (hasNonAscii && hasTextLetters) {
        return '\\text{' + s.replace(/[\\{}%#$&_]/g, '\\$&') + '}';
    }
    return s
        .replace(/\\/g, '\\backslash ')
        .replace(/([{}%#$&_])/g, '\\$1')
        .replace(/\^/g, '\\^{}')
        .replace(/~/g, '\\~{}');
}

function ommlDelimiterToLatex(chr) {
    var map = { '{': '\\{', '}': '\\}', '<': '\\langle', '>': '\\rangle' };
    return map[chr] || chr;
}

function ommlNodeToLatex(node) {
    if (!node) return '';
    if (node.type === 'text') return node.text || '';
    var name = (node.name || '').replace('m:', '');
    var children = node.elements || [];
    function child(n) { return children.find(function(c){ return (c.name||'').replace('m:','') === n; }); }
    function childrenOf(n) { var c = child(n); return c ? (c.elements || []) : []; }
    function latex(nodes) { return (nodes||[]).map(ommlNodeToLatex).join(''); }
    function latexChild(n) { return latex(childrenOf(n)); }
    switch(name) {
        case 'oMath': case 'oMathPara': return latex(children);
        case 'f': return '\\frac{' + latexChild('num') + '}{' + latexChild('den') + '}';
        case 'sSup': return latexChild('e') + '^{' + latexChild('sup') + '}';
        case 'sSub': return latexChild('e') + '_{' + latexChild('sub') + '}';
        case 'sSubSup': return latexChild('e') + '_{' + latexChild('sub') + '}^{' + latexChild('sup') + '}';
        case 'rad': {
            var pr = child('radPr');
            var hideEl = pr && (pr.elements||[]).find(function(c){ return (c.name||'').replace('m:','') === 'degHide'; });
            var hideAttr = hideEl && hideEl.attributes ? (hideEl.attributes['m:val'] || hideEl.attributes.val) : '';
            var hide = hideEl && hideAttr !== '0' && hideAttr !== 'false';
            var deg = latexChild('deg'); var e = latexChild('e');
            if (hide || !deg.trim()) return '\\sqrt{' + e + '}';
            return '\\sqrt[' + deg + ']{' + e + '}';
        }
        case 'nary': {
            var nPr = child('naryPr');
            var chrEl = nPr && (nPr.elements||[]).find(function(c){ return (c.name||'').replace('m:','') === 'chr'; });
            var chr = chrEl && chrEl.attributes && (chrEl.attributes['m:val'] || chrEl.attributes.val);
            var symMap = {
                '\u222B': '\\int', '\u222C': '\\iint', '\u222D': '\\iiint',
                '\u2211': '\\sum', '\u220F': '\\prod', '\u2210': '\\coprod',
                '\u222E': '\\oint', '\u22C3': '\\bigcup', '\u22C2': '\\bigcap'
            };
            var sym = symMap[chr] || '\\int';
            var sub2 = latexChild('sub'); var sup2 = latexChild('sup'); var e2 = latexChild('e');
            return sym + (sub2 ? '_{'+sub2+'}' : '') + (sup2 ? '^{'+sup2+'}' : '') + ' ' + e2;
        }
        case 'd': {
            var dPr = child('dPr');
            function getDPrChr(tag, def) {
                var el = dPr && (dPr.elements||[]).find(function(c){ return (c.name||'').replace('m:','') === tag; });
                return el && el.attributes ? (el.attributes['m:val'] || el.attributes.val || def) : def;
            }
            var beg = getDPrChr('begChr','('); var end = getDPrChr('endChr',')');
            var inner = children.filter(function(c){ return (c.name||'').replace('m:','') === 'e'; })
                .map(function(c){ return latex(c.elements||[]); }).join(', ');
            if (!beg && !end) return inner;
            // \left/\right can '.' khi mot ben khong co dau ngoac.
            return '\\left' + (beg ? ommlDelimiterToLatex(beg) : '.') +
                inner + '\\right' + (end ? ommlDelimiterToLatex(end) : '.');
        }
        case 'limLow': return latexChild('e') + '_{' + latexChild('lim') + '}';
        case 'limUpp': return latexChild('e') + '^{' + latexChild('lim') + '}';
        case 'm': {
            var rows = children.filter(function(c){ return (c.name||'').replace('m:','') === 'mr'; });
            return '\\begin{pmatrix}' + rows.map(function(row){
                return (row.elements||[]).filter(function(c){ return (c.name||'').replace('m:','') === 'e'; })
                    .map(function(c){ return latex(c.elements||[]); }).join(' & ');
            }).join(' \\\\ ') + '\\end{pmatrix}';
        }
        case 'func': return latexChild('fName') + ' ' + latexChild('e');
        case 'acc': {
            var aPr = child('accPr');
            var aChrEl = aPr && (aPr.elements||[]).find(function(c){ return (c.name||'').replace('m:','') === 'chr'; });
            var aChr = aChrEl && aChrEl.attributes ? (aChrEl.attributes['m:val'] || aChrEl.attributes.val) : '';
            var accMap = {'\u0302':'\\hat','\u0303':'\\tilde','\u0304':'\\bar','\u20D7':'\\vec','\u0307':'\\dot','\u0308':'\\ddot'};
            return (accMap[aChr]||'\\hat') + '{' + latexChild('e') + '}';
        }
        case 'bar': {
            var bPr = child('barPr');
            var posEl = bPr && (bPr.elements||[]).find(function(c){ return (c.name||'').replace('m:','') === 'pos'; });
            var pos2 = posEl && posEl.attributes ? (posEl.attributes['m:val'] || posEl.attributes.val) : 'top';
            return pos2 === 'bot' ? '\\underline{' + latexChild('e') + '}' : '\\overline{' + latexChild('e') + '}';
        }
        case 'eqArr': {
            var eqs = children.filter(function(c){ return (c.name||'').replace('m:','') === 'e'; });
            return '\\begin{aligned}' + eqs.map(function(c){ return latex(c.elements||[]); }).join(' \\\\ ') + '\\end{aligned}';
        }
        case 'r': case 'e': case 'num': case 'den': case 'sup': case 'sub':
        case 'deg': case 'lim': case 'fName': return latex(children);
        case 't': return ommlTextToLatex(children.map(function(c){ return c.text || ''; }).join(''));
        case 'rPr': case 'fPr': case 'sSupPr': case 'sSubPr': case 'sSubSupPr':
        case 'radPr': case 'naryPr': case 'dPr': case 'mPr': case 'mrPr':
        case 'limLowPr': case 'limUppPr': case 'funcPr': case 'accPr':
        case 'ctrlPr': case 'eqArrPr': case 'boxPr': case 'barPr': return '';
        default: return latex(children);
    }
}

function ommlStringToLatex(ommlStr) {
    try {
        var obj = xmljs.xml2js(ommlStr, { compact: false, ignoreDeclaration: true });
        var root = obj.elements && obj.elements[0];
        return ommlNodeToLatex(root);
    } catch(e) { return null; }
}

// Sua XML hong do file .docx tao bang tay: ky tu "<", ">", "&" nam truc tiep
// trong <w:t>/<m:t> lam xmldom (mammoth) bao loi parse. Word that luon escape.
function escapeTextNodes(xml) {
    return xml.replace(/<(w|m):t\b([^>]*)>([\s\S]*?)<\/\1:t>/g, function (full, prefix, attrs, text) {
        var fixed = text
            .replace(/&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');
        return '<' + prefix + ':t' + attrs + '>' + fixed + '</' + prefix + ':t>';
    });
}

// Pre-process docx buffer: thay OMML bằng placeholder text để mammoth giữ lại vị trí
function preprocessDocxMath(buffer) {
    return JSZip.loadAsync(buffer).then(function(zip) {
        var docFile = zip.file('word/document.xml');
        if (!docFile) return { buffer: buffer, mathMap: {} };
        return docFile.async('string').then(function(xml) {
            var fixed = escapeTextNodes(xml);
            if (xml.indexOf('m:oMath') === -1) {
                if (fixed === xml) return { buffer: buffer, mathMap: {} };
                zip.file('word/document.xml', fixed);
                return zip.generateAsync({ type: 'nodebuffer' }).then(function (newBuf) {
                    return { buffer: newBuf, mathMap: {} };
                });
            }
            xml = fixed;
            var mathMap = {};
            var counter = 0;
            var result = '';
            var i = 0;
            while (i < xml.length) {
                var startTag = xml.indexOf('<m:oMath', i);
                if (startTag === -1) { result += xml.slice(i); break; }
                result += xml.slice(i, startTag);
                var tagEnd = xml.indexOf('>', startTag);
                var tagName = xml.slice(startTag + 1, tagEnd).split(/[\s>]/)[0];
                var closeTag = '</' + tagName + '>';
                var depth = 1, pos = tagEnd + 1;
                while (depth > 0 && pos < xml.length) {
                    var nextOpen  = xml.indexOf('<' + tagName, pos);
                    var nextClose = xml.indexOf(closeTag, pos);
                    if (nextClose === -1) break;
                    if (nextOpen !== -1 && nextOpen < nextClose) { depth++; pos = nextOpen + 1; }
                    else { depth--; pos = (depth === 0) ? nextClose + closeTag.length : nextClose + 1; }
                }
                var ommlStr = xml.slice(startTag, pos);
                var latexStr = ommlStringToLatex(ommlStr);
                if (latexStr !== null) {
                    // oMathPara = cong thuc doc lap; oMath = cong thuc noi dong.
                    var wrapped = (tagName === 'm:oMathPara')
                        ? '\\[' + latexStr + '\\]'
                        : '\\(' + latexStr + '\\)';
                    var placeholder = 'MATHPLACEHOLDER' + counter + 'END';
                    mathMap[placeholder] = wrapped;
                    result += '<w:r><w:t xml:space="preserve"> ' + placeholder + ' </w:t></w:r>';
                    counter++;
                } else {
                    result += ommlStr;
                }
                i = pos;
            }
            zip.file('word/document.xml', result);
            return zip.generateAsync({ type: 'nodebuffer' }).then(function(newBuf) {
                return { buffer: newBuf, mathMap: mathMap };
            });
        });
    });
}

// ========== MAMMOTH HTML HELPER ==========
function mammothToHtml(buffer) {
    return preprocessDocxMath(buffer).then(function(data) {
        return mammoth.convertToHtml(
            { buffer: data.buffer },
            {
                styleMap: [
                    "p[style-name='Code'] => pre",
                    // Giu lai thong tin GACH CHAN cua Word (mammoth mac dinh bo qua).
                    // Nhieu de thi danh dau dap an dung bang cach gach chan chu cai
                    // dap an -> can thong tin nay de nhan dien dung dap an.
                    "u => u"
                ],
                convertImage: mammoth.images.imgElement(function(image) {
                    return image.read('base64').then(function(imageBase64) {
                        return { src: 'data:' + image.contentType + ';base64,' + imageBase64 };
                    });
                })
            }
        ).then(function(result) {
            var html = result.value;
            Object.keys(data.mathMap).forEach(function(placeholder) {
                // Escape HTML truoc khi chen: LaTeX co the chua "<", ">" hoac "&".
                var latex = data.mathMap[placeholder]
                    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
                html = html.split(placeholder).join(latex);
            });
            // Escape HTML inside <pre> so code is displayed as text, not rendered as HTML
            html = html.replace(/<pre>([\s\S]*?)<\/pre>/gi, function(match, inner) {
                var escaped = inner
                    .replace(/&amp;/g, '&')   // decode any double-encoding first
                    .replace(/&lt;/g, '<')
                    .replace(/&gt;/g, '>')
                    .replace(/&quot;/g, '"')
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/<br\s*\/?>/gi, '\n');  // convert <br> back to newlines
                return '<pre>' + escaped + '</pre>';
            });
            return { value: html, messages: result.messages };
        });
    });
}

module.exports = {
    mammothToHtml: mammothToHtml
};
