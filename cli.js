#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');
const http = require('http');
const puppeteer = require('puppeteer-core');

const MIME_TYPES = {
    '.html': 'text/html; charset=utf-8',
    '.js': 'application/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.png': 'image/png',
    '.svg': 'image/svg+xml',
    '.ico': 'image/x-icon'
};

function printHelp() {
    console.log(`
Antigravity Doc Porter CLI — Convert Markdown (+ Mermaid / PlantUML / MathJax) to .docx

Usage:
  md-porter <input.md> [--out <output.docx>]     Convert a single Markdown file
  md-porter [dir] --all [--out-dir <dir>] [-r]    Convert all .md files in a directory

Options:
  -o, --out <file>       Output .docx path (single-file mode; defaults to <input>.docx)
  -a, --all              Convert all .md files in target directory (defaults to current dir)
  -d, --out-dir <dir>    Output directory for --all mode (defaults to each file's folder)
  -r, --recursive        Recurse into subdirectories in --all mode
  -h, --help             Show this help message

Environment:
  CHROME_PATH            Optional path to Chrome/Edge executable
`);
}

function parseArgs(argv) {
    const args = {
        target: null,
        out: null,
        outDir: null,
        all: false,
        recursive: false,
        help: false
    };

    for (let i = 0; i < argv.length; i++) {
        const a = argv[i];
        if (a === '-h' || a === '--help') {
            args.help = true;
        } else if (a === '-a' || a === '--all') {
            args.all = true;
        } else if (a === '-r' || a === '--recursive') {
            args.recursive = true;
        } else if (a === '-o' || a === '--out') {
            args.out = argv[++i];
        } else if (a === '-d' || a === '--out-dir') {
            args.outDir = argv[++i];
        } else if (!a.startsWith('-')) {
            if (!args.target) args.target = a;
        } else {
            throw new Error(`Unknown option: ${a}`);
        }
    }
    return args;
}

function findBrowserExecutable() {
    if (process.env.CHROME_PATH && fs.existsSync(process.env.CHROME_PATH)) {
        return process.env.CHROME_PATH;
    }

    const candidates = process.platform === 'win32'
        ? [
            'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
            'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
            'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
            'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe'
        ]
        : process.platform === 'darwin'
            ? [
                '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
                '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge'
            ]
            : [
                '/usr/bin/google-chrome',
                '/usr/bin/chromium',
                '/usr/bin/chromium-browser',
                '/usr/bin/microsoft-edge'
            ];

    for (const p of candidates) {
        if (fs.existsSync(p)) return p;
    }
    throw new Error('Could not find Chrome or Edge. Set CHROME_PATH environment variable.');
}

function startStaticServer(rootDir) {
    return new Promise((resolve, reject) => {
        const server = http.createServer((req, res) => {
            const urlPath = (req.url || '/').split('?')[0];
            const relPath = urlPath === '/' ? 'index.html' : urlPath.replace(/^\/+/, '');
            const filePath = path.resolve(rootDir, relPath);

            if (!filePath.startsWith(path.resolve(rootDir))) {
                res.writeHead(403);
                res.end('Forbidden');
                return;
            }

            fs.readFile(filePath, (err, data) => {
                if (err) {
                    res.writeHead(404);
                    res.end('Not Found');
                    return;
                }
                const ext = path.extname(filePath).toLowerCase();
                res.writeHead(200, { 'Content-Type': MIME_TYPES[ext] || 'application/octet-stream' });
                res.end(data);
            });
        });

        server.on('error', reject);
        server.listen(0, '127.0.0.1', () => {
            const addr = server.address();
            resolve({ server, port: addr.port });
        });
    });
}

function collectMarkdownFiles(dir, recursive) {
    const results = [];
    const entries = fs.readdirSync(dir, { withFileTypes: true });
    for (const entry of entries) {
        if (entry.name === 'node_modules' || entry.name.startsWith('.')) continue;
        const fullPath = path.join(dir, entry.name);
        if (entry.isDirectory()) {
            if (recursive) results.push(...collectMarkdownFiles(fullPath, true));
        } else if (entry.isFile() && /\.md$/i.test(entry.name)) {
            results.push(fullPath);
        }
    }
    return results;
}

async function convertFile(page, mdPath, outPath) {
    const mdContent = fs.readFileSync(mdPath, 'utf8');

    const base64Docx = await page.evaluate(async (markdown) => {
        const input = document.getElementById('markdown-input');
        input.value = markdown;

        if (window.MathJax && window.MathJax.startup && window.MathJax.startup.promise) {
            await window.MathJax.startup.promise;
        }

        await window.renderDocument();

        // Wait for any PlantUML diagrams (which use a 500ms debounce + fetch)
        const pumlTargets = document.querySelectorAll('.plantuml-render-target');
        if (pumlTargets.length > 0) {
            const deadline = Date.now() + 20000;
            while (Date.now() < deadline) {
                const allDone = Array.from(pumlTargets).every(
                    el => el.querySelector('svg') || el.textContent.includes('PlantUML Render Error')
                );
                if (allDone) break;
                await new Promise(r => setTimeout(r, 150));
            }
        }

        return await new Promise((resolve, reject) => {
            const origSaveAs = window.saveAs;
            window.saveAs = (blob) => {
                window.saveAs = origSaveAs;
                const reader = new FileReader();
                reader.onloadend = () => {
                    const dataUrl = reader.result;
                    resolve(dataUrl.split(',')[1]);
                };
                reader.onerror = () => reject(new Error('Failed to read generated DOCX blob'));
                reader.readAsDataURL(blob);
            };

            const btn = document.getElementById('docx-btn');
            if (!btn) {
                reject(new Error('Export DOCX button not found in UI'));
                return;
            }
            btn.click();
        });
    }, mdContent);

    fs.mkdirSync(path.dirname(outPath), { recursive: true });
    fs.writeFileSync(outPath, Buffer.from(base64Docx, 'base64'));
}

async function main() {
    let args;
    try {
        args = parseArgs(process.argv.slice(2));
    } catch (err) {
        console.error(`Error: ${err.message}`);
        printHelp();
        process.exit(1);
    }

    if (args.help || (!args.target && !args.all)) {
        printHelp();
        process.exit(args.help ? 0 : 1);
    }

    const jobs = [];
    if (args.all) {
        const targetDir = path.resolve(process.cwd(), args.target || '.');
        if (!fs.existsSync(targetDir) || !fs.statSync(targetDir).isDirectory()) {
            console.error(`Error: Directory not found: ${targetDir}`);
            process.exit(1);
        }
        const mdFiles = collectMarkdownFiles(targetDir, args.recursive);
        if (mdFiles.length === 0) {
            console.log(`No .md files found in ${targetDir}`);
            return;
        }
        for (const mdFile of mdFiles) {
            const rel = path.relative(targetDir, mdFile);
            const outFile = args.outDir
                ? path.resolve(process.cwd(), args.outDir, rel.replace(/\.md$/i, '.docx'))
                : mdFile.replace(/\.md$/i, '.docx');
            jobs.push({ mdFile, outFile });
        }
    } else {
        const mdFile = path.resolve(process.cwd(), args.target);
        if (!fs.existsSync(mdFile) || !fs.statSync(mdFile).isFile()) {
            console.error(`Error: File not found: ${mdFile}`);
            process.exit(1);
        }
        const outFile = args.out
            ? path.resolve(process.cwd(), args.out)
            : mdFile.replace(/\.md$/i, '.docx');
        jobs.push({ mdFile, outFile });
    }

    const executablePath = findBrowserExecutable();
    const { server, port } = await startStaticServer(__dirname);
    let browser;

    try {
        browser = await puppeteer.launch({
            executablePath,
            headless: true,
            args: ['--no-sandbox', '--disable-setuid-sandbox']
        });

        const page = await browser.newPage();
        await page.goto(`http://127.0.0.1:${port}/`, { waitUntil: 'networkidle0', timeout: 30000 });

        for (const job of jobs) {
            process.stdout.write(`Converting ${path.relative(process.cwd(), job.mdFile)} -> ${path.relative(process.cwd(), job.outFile)} ... `);
            await convertFile(page, job.mdFile, job.outFile);
            console.log('done');
        }
    } finally {
        if (browser) await browser.close();
        server.close();
    }
}

main().catch((err) => {
    console.error(`Failed: ${err.message}`);
    process.exit(1);
});
