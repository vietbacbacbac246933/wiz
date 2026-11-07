'use strict';
/**
 * Helper cho test: tu khoi dong server o cong 3999 neu chua chay, va upload
 * file .docx len /api/extract.
 */

const { spawn } = require('child_process');
const path = require('path');

const PORT = 3999;
const BASE = process.env.TEST_BASE || ('http://127.0.0.1:' + PORT);
const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

let child = null;

async function ensureServer() {
    try {
        const r = await fetch(BASE + '/api/health');
        if (r.ok) return BASE;
    } catch (e) { /* chua chay */ }
    child = spawn(process.execPath, ['server.js'], {
        cwd: path.resolve(__dirname, '..'),
        env: Object.assign({}, process.env, { PORT: String(PORT) }),
        stdio: 'ignore'
    });
    for (let i = 0; i < 60; i++) {
        await new Promise(r => setTimeout(r, 200));
        try {
            const r = await fetch(BASE + '/api/health');
            if (r.ok) return BASE;
        } catch (e) { /* dang khoi dong */ }
    }
    throw new Error('Server khong khoi dong duoc o ' + BASE);
}

function stopServer() {
    if (child) child.kill();
    child = null;
}

async function uploadDocx(buffer, name) {
    const fd = new FormData();
    fd.append('file', new Blob([buffer], { type: DOCX_TYPE }), name || 'de.docx');
    const res = await fetch(BASE + '/api/extract', { method: 'POST', body: fd });
    return { status: res.status, json: await res.json().catch(function () { return null; }) };
}

module.exports = { BASE, ensureServer, stopServer, uploadDocx };
