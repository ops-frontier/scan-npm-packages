#!/usr/bin/env node
/**
 * scan-npm-packages
 * GitHub 組織内の全リポジトリを巡回し、汚染された npm パッケージを検出するツール。
 */

import { readFileSync } from 'fs';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';
import ExcelJS from 'exceljs';

const __dirname = dirname(fileURLToPath(import.meta.url));

// ---- 設定 ----
const GITHUB_TOKEN = process.env.GITHUB_TOKEN;
const GITHUB_ORG = process.env.GITHUB_ORGANIZATION;
const CONCURRENCY = 5;
const LOCK_FILENAMES = new Set(['package-lock.json', 'yarn.lock', 'pnpm-lock.yaml']);
const GH_API = 'https://api.github.com';

if (!GITHUB_TOKEN) {
  console.error('Error: GITHUB_TOKEN 環境変数が設定されていません。');
  process.exit(1);
}
if (!GITHUB_ORG) {
  console.error('Error: GITHUB_ORGANIZATION 環境変数が設定されていません。');
  process.exit(1);
}

// ---- package-list.md のパース ----
function loadPackageList() {
  const content = readFileSync(join(__dirname, 'package-list.md'), 'utf8');
  const pkgs = new Set();
  const contaminatedVersions = new Map();
  for (const line of content.split('\n')) {
    if (!line.startsWith('|')) continue;
    const cols = line.split('|').map(c => c.trim()).filter(Boolean);
    if (!cols[0] || cols[0] === 'パッケージ名' || /^-+$/.test(cols[0])) continue;
    pkgs.add(cols[0]);
    if (cols[1]) {
      contaminatedVersions.set(cols[0], new Set(cols[1].split(',').map(v => v.trim()).filter(Boolean)));
    }
  }
  return { pkgs, contaminatedVersions };
}

// ---- GitHub API ----
async function ghFetch(path) {
  const url = path.startsWith('http') ? path : `${GH_API}${path}`;
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${GITHUB_TOKEN}`,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
    },
  });
  if (res.status === 404) return null;
  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`GitHub API ${res.status} (${url}): ${body}`);
  }
  return res.json();
}

async function listAllRepos(org) {
  const repos = [];
  let page = 1;
  for (;;) {
    const data = await ghFetch(
      `/orgs/${encodeURIComponent(org)}/repos?per_page=100&page=${page}&type=all`
    );
    if (!data || data.length === 0) break;
    for (const r of data) repos.push({ name: r.name, branch: r.default_branch });
    if (data.length < 100) break;
    page++;
  }
  return repos;
}

async function getRepoTree(org, repo, branch) {
  const data = await ghFetch(
    `/repos/${encodeURIComponent(org)}/${encodeURIComponent(repo)}/git/trees/${encodeURIComponent(branch)}?recursive=1`
  );
  if (!data || data.message) return null;
  return data;
}

async function getBlobText(org, repo, sha) {
  const blob = await ghFetch(
    `/repos/${encodeURIComponent(org)}/${encodeURIComponent(repo)}/git/blobs/${sha}`
  );
  if (!blob) return null;
  return Buffer.from(blob.content.replace(/\s/g, ''), 'base64').toString('utf8');
}

// ---- ロックファイルのパーサー ----

/**
 * package-lock.json をパースして対象パッケージを検索する。
 * lockfileVersion 1 (npm v6) と 2/3 (npm v7+) の両方に対応。
 */
function parsePackageLock(content) {
  const results = [];
  let data;
  try { data = JSON.parse(content); } catch {
    console.warn('package-lock.json のパースに失敗しました。');
    return results;
}

  if (data.packages) {
    // npm v7+ (lockfileVersion 2/3): キーは "node_modules/pkg" の形式
    for (const [key, info] of Object.entries(data.packages)) {
      if (!key || typeof info.version !== 'string') continue;
      const idx = key.lastIndexOf('node_modules/');
      if (idx === -1) continue;
      const name = key.slice(idx + 'node_modules/'.length);
      results.push({ name, version: info.version });
    }
  } else if (data.dependencies) {
    // npm v6 (lockfileVersion 1): 再帰的な dependencies オブジェクト
    const seen = new Set();
    function walk(deps) {
      for (const [name, info] of Object.entries(deps)) {
        if (typeof info.version !== 'string') continue;
        const k = `${name}@${info.version}`;
        if (seen.has(k)) continue;
        seen.add(k);
        results.push({ name, version: info.version });
        if (info.dependencies) walk(info.dependencies);
      }
    }
    walk(data.dependencies);
  }

  return dedupe(results);
}

/**
 * yarn.lock をパースして対象パッケージを検索する。
 * v1 および berry (v2+) の両形式に対応。
 */
function parseYarnLock(content) {
  const results = [];
  const seen = new Set();
  const lines = content.split('\n');

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // インデントあり・空行・コメント行はヘッダーではない
    if (!line || line.startsWith(' ') || line.startsWith('\t') || line.startsWith('#')) continue;

    // ヘッダー行を取得 (例: "@scope/pkg@^1.0.0, @scope/pkg@~1.0.0":)
    let header;
    const quotedMatch = line.match(/^"(.+)":\s*$/);
    if (quotedMatch) {
      header = quotedMatch[1];
    } else if (line.endsWith(':')) {
      header = line.slice(0, -1).trim();
    } else {
      continue;
    }

    // コンマ区切りの複数スペックをパース
    const specifiers = header.split(/,\s*/);
    const matchedNames = new Set();

    for (const spec of specifiers) {
      const s = spec.trim().replace(/^"|"$/g, '');
      let name;
      if (s.startsWith('@')) {
        // スコープ付き: @scope/name@specifier
        const m = s.match(/^(@[^/]+\/[^@]+)@/);
        if (m) name = m[1];
      } else {
        const atIdx = s.indexOf('@');
        if (atIdx > 0) name = s.slice(0, atIdx);
      }
      if (name) matchedNames.add(name);
    }

    if (matchedNames.size === 0) continue;

    // 直後のインデント行から version を取得
    for (let j = i + 1; j < Math.min(i + 15, lines.length); j++) {
      const vl = lines[j];
      if (vl && !vl.startsWith(' ') && !vl.startsWith('\t')) break;
      const vm = vl.match(/^\s+version:?\s+"?([^\s"]+)"?\s*$/);
      if (vm) {
        const version = vm[1];
        for (const name of matchedNames) {
          const k = `${name}@${version}`;
          if (!seen.has(k)) { seen.add(k); results.push({ name, version }); }
        }
        break;
      }
    }
  }
  return results;
}

/**
 * pnpm-lock.yaml をパースして対象パッケージを検索する。
 * v5 / v6 / v9 の各形式に対応。
 */
function parsePnpmLock(content) {
  const results = [];
  const seen = new Set();

  for (const line of content.split('\n')) {
    // パッケージキーはインデントあり・末尾がコロン
    if (!/^\s/.test(line) || !line.trimEnd().endsWith(':')) continue;

    let key = line.trim().replace(/:$/, '');

    // 囲みクォートを除去 (' または ")
    if ((key.startsWith("'") && key.endsWith("'")) ||
        (key.startsWith('"') && key.endsWith('"'))) {
      key = key.slice(1, -1);
    }

    // 先頭スラッシュを除去 (v5/v6 形式: "/pkg@version")
    if (key.startsWith('/')) key = key.slice(1);

    let name, version;
    if (key.startsWith('@')) {
      // スコープ付き: @scope/name@version または @scope/name/version
      const m = key.match(/^(@[^/]+\/[^@/(]+)[@/]([^(]+)/);
      if (m) { name = m[1]; version = m[2]; }
    } else {
      // スコープなし: name@version または name/version
      const m = key.match(/^([^@/(]+)[@/]([^(]+)/);
      if (m) { name = m[1]; version = m[2]; }
    }

    if (!name || !version) continue;

    // ピア依存サフィックス "(peer@version)" を除去
    version = version.replace(/\(.*\)$/, '').trim();

    const k = `${name}@${version}`;
    if (!seen.has(k)) { seen.add(k); results.push({ name, version }); }
  }
  return results;
}

function dedupe(arr) {
  const seen = new Set();
  return arr.filter(({ name, version }) => {
    const k = `${name}@${version}`;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
}

// ---- 並列実行リミッター ----
function createLimiter(n) {
  let active = 0;
  const queue = [];
  function drain() {
    while (active < n && queue.length) {
      active++;
      const { fn, resolve, reject } = queue.shift();
      fn().then(resolve, reject).finally(() => { active--; drain(); });
    }
  }
  return fn => new Promise((resolve, reject) => { queue.push({ fn, resolve, reject }); drain(); });
}

// ---- リポジトリのスキャン ----
async function scanRepo(org, repo) {
  const findings = [];

  const tree = await getRepoTree(org, repo.name, repo.branch);
  if (!tree) return findings;

  const lockFiles = tree.tree.filter(
    item => item.type === 'blob' && LOCK_FILENAMES.has(item.path.split('/').pop())
  );

  for (const file of lockFiles) {
    const content = await getBlobText(org, repo.name, file.sha);
    if (!content) continue;

    const filename = file.path.split('/').pop();
    let found;
    if (filename === 'package-lock.json') found = parsePackageLock(content);
    else if (filename === 'yarn.lock') found = parseYarnLock(content);
    else if (filename === 'pnpm-lock.yaml') found = parsePnpmLock(content);
    else continue;

    for (const { name, version } of found) {
      findings.push({ repo: repo.name, file: file.path, name, version });
    }
  }
  return findings;
}

// ---- Excel 出力 ----
async function writeExcel(findings, pkgs, contaminatedVersions) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.addRow(['リポジトリ', 'ファイルパス', 'パッケージ名', 'パッケージバージョン', '対象パッケージ数', '汚染パッケージ数']);
  for (const { repo, file, name, version } of findings) {
    const isTarget = pkgs.has(name) ? 1 : 0;
    const contamVersions = contaminatedVersions.get(name) ?? new Set();
    const isContaminated = contamVersions.has(version) ? 1 : 0;
    sheet.addRow([repo, file, name, version, isTarget, isContaminated]);
  }
  await workbook.xlsx.writeFile('shai-hulud-scan.xlsx');
  console.log('Excel ファイルを出力しました: shai-hulud-scan.xlsx');
}

// ---- メイン ----
async function main() {
  const { pkgs, contaminatedVersions } = loadPackageList();
  console.log(`対象パッケージ数: ${pkgs.size} (package-list.md より読み込み)`);
  console.log(`組織 "${GITHUB_ORG}" のリポジトリを取得中...`);

  const repos = await listAllRepos(GITHUB_ORG);
  if (repos.length === 0) {
    console.log('リポジトリが見つかりませんでした。');
    return;
  }
  console.log(`リポジトリ数: ${repos.length} (並列数: ${CONCURRENCY})\n`);

  const limit = createLimiter(CONCURRENCY);
  let done = 0;

  const allFindings = (await Promise.all(
    repos.map(repo =>
      limit(async () => {
        try {
          const findings = await scanRepo(GITHUB_ORG, repo);
          done++;
          process.stderr.write(`\r進捗: ${done}/${repos.length} - ${repo.name}                               `);
          return findings;
        } catch (err) {
          done++;
          process.stderr.write(`\r進捗: ${done}/${repos.length} - ${repo.name} [エラー: ${err.message}]   `);
          return [];
        }
      })
    )
  )).flat();

  process.stderr.write('\n\n');

  // コンソール出力は package-list.md に含まれるパッケージのみ
  const consoleFindings = allFindings.filter(f => pkgs.has(f.name));

  if (consoleFindings.length === 0) {
    console.log('汚染されたパッケージは見つかりませんでした。');
  } else {
    console.log(`汚染パッケージへの参照が ${consoleFindings.length} 件見つかりました:\n`);
    for (const { repo, file, name, version } of consoleFindings) {
      console.log(`Repository : ${repo}`);
      console.log(`File       : ${file}`);
      console.log(`Package    : ${name}@${version}`);
      console.log('');
    }
  }

  // Excel には全パッケージを出力
  await writeExcel(allFindings, pkgs, contaminatedVersions);
}

main().catch(err => {
  console.error('\n致命的なエラー:', err.message);
  process.exit(1);
});
