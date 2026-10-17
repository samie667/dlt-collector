#!/usr/bin/env node
'use strict';
/**
 * 大乐透开奖数据收集器
 *
 * 数据源：中国体彩网官方接口（webapi.sporttery.cn）
 * 用法：
 *   node dlt.js latest                     查看最新一期开奖结果并入库
 *   node dlt.js fetch [--pages 3] [--all]  抓取历史数据（增量，--all 全量）
 *   node dlt.js show [--limit 10]          查看本地已存储数据
 *   node dlt.js stats [--limit 100]        号码出现频率统计
 * 数据保存在脚本同目录 data/ 下（CSV + JSON），按期号去重。
 * 仅使用 Node.js 内置模块（Node >= 18），无需安装依赖。
 */

const fs = require('fs');
const path = require('path');

const API_URL = 'https://webapi.sporttery.cn/gateway/lottery/getHistoryPageListV1.qry';
const HEADERS = {
  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
  'Referer': 'https://www.lottery.gov.cn/',
};

const DATA_DIR = path.join(__dirname, 'data');
const CSV_PATH = path.join(DATA_DIR, 'dlt_history.csv');
const JSON_PATH = path.join(DATA_DIR, 'dlt_history.json');

const CSV_FIELDS = [
  '期号', '开奖日期',
  '前区1', '前区2', '前区3', '前区4', '前区5',
  '后区1', '后区2',
  '销售额', '奖池余额', '一等奖注数', '一等奖单注奖金',
];

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/** 抓取一页开奖数据，返回原始记录数组 */
async function fetchPage(pageNo, pageSize = 30, retries = 3) {
  const url = `${API_URL}?gameNo=85&provinceId=0&pageSize=${pageSize}&is498=1&pageNo=${pageNo}`;
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      const resp = await fetch(url, {
        headers: HEADERS,
        signal: AbortSignal.timeout(15000),
      });
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const payload = await resp.json();
      return (payload && payload.value && payload.value.list) || [];
    } catch (e) {
      if (attempt === retries) throw e;
      const wait = attempt * 2;
      console.error(`  第 ${pageNo} 页请求失败(${e.message})，${wait}s 后重试...`);
      await sleep(wait * 1000);
    }
  }
  return [];
}

/** 把接口记录解析为一行存储数据 */
function parseDraw(item) {
  const nums = String(item.lotteryDrawResult || '').trim().split(/\s+/);
  const row = {
    '期号': item.lotteryDrawNum || '',
    '开奖日期': item.lotteryDrawTime || '',
    '销售额': item.totalSaleAmount || '',
    '奖池余额': item.poolBalanceAfterdraw || '',
    '一等奖注数': '',
    '一等奖单注奖金': '',
  };
  for (let i = 0; i < 5; i++) row[`前区${i + 1}`] = nums[i] || '';
  for (let i = 0; i < 2; i++) row[`后区${i + 1}`] = nums[5 + i] || '';
  for (const p of item.prizeLevelList || []) {
    if (p.prizeLevel === '一等奖') {
      row['一等奖注数'] = p.stakeCount || '';
      row['一等奖单注奖金'] = p.stakeAmount || '';
      break;
    }
  }
  return row;
}

// ---------- CSV 读写（字段含英文逗号，需转义；带 BOM 方便 Excel 打开） ----------

function csvEscape(v) {
  const s = String(v ?? '');
  return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

function parseCsvLine(line) {
  const out = [];
  let cur = '', inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (inQuotes) {
      if (c === '"' && line[i + 1] === '"') { cur += '"'; i++; }
      else if (c === '"') inQuotes = false;
      else cur += c;
    } else if (c === '"') inQuotes = true;
    else if (c === ',') { out.push(cur); cur = ''; }
    else cur += c;
  }
  out.push(cur);
  return out;
}

/** 读取本地 CSV，返回 Map: 期号 -> row */
function loadLocal() {
  const rows = new Map();
  if (!fs.existsSync(CSV_PATH)) return rows;
  const text = fs.readFileSync(CSV_PATH, 'utf-8').replace(/^﻿/, '');
  const lines = text.split(/\r?\n/).filter((l) => l.trim());
  if (lines.length < 2) return rows;
  const header = parseCsvLine(lines[0]);
  for (const line of lines.slice(1)) {
    const values = parseCsvLine(line);
    const row = {};
    header.forEach((h, i) => { row[h] = values[i] ?? ''; });
    if (row['期号']) rows.set(row['期号'], row);
  }
  return rows;
}

/** 按期号倒序写回 CSV 和 JSON */
function saveAll(rowsMap) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  const ordered = [...rowsMap.values()].sort((a, b) => b['期号'].localeCompare(a['期号']));
  const lines = [CSV_FIELDS.map(csvEscape).join(',')];
  for (const row of ordered) lines.push(CSV_FIELDS.map((f) => csvEscape(row[f])).join(','));
  fs.writeFileSync(CSV_PATH, '﻿' + lines.join('\r\n') + '\r\n', 'utf-8');
  fs.writeFileSync(JSON_PATH, JSON.stringify(ordered, null, 2), 'utf-8');
}

function printRows(rows) {
  for (const r of rows) {
    const front = [1, 2, 3, 4, 5].map((i) => r[`前区${i}`]).join(' ');
    const back = [1, 2].map((i) => r[`后区${i}`]).join(' ');
    console.log(`${r['期号']}  ${r['开奖日期']}  前区[${front}] 后区[${back}]  销售额:${r['销售额'] || '-'}  奖池:${r['奖池余额'] || '-'}`);
  }
}

// ---------- 子命令 ----------

async function cmdLatest() {
  const items = await fetchPage(1, 1);
  if (!items.length) { console.log('未获取到数据'); return; }
  const row = parseDraw(items[0]);
  console.log('最新一期开奖结果：');
  printRows([row]);
  const local = loadLocal();
  if (!local.has(row['期号'])) {
    local.set(row['期号'], row);
    saveAll(local);
    console.log(`\n已入库，本地共 ${local.size} 期。`);
  } else {
    console.log('\n该期已在本地库中。');
  }
}

async function cmdFetch(args) {
  const local = loadLocal();
  console.log(`本地已有 ${local.size} 期数据，开始增量抓取...`);
  let newCount = 0, page = 1;
  const maxPages = args.all ? Infinity : (args.pages || 3);
  const pageSize = args['page-size'] || 30;
  while (page <= maxPages) {
    const items = await fetchPage(page, pageSize);
    if (!items.length) { console.log(`第 ${page} 页无数据，抓取结束。`); break; }
    let pageNew = 0;
    for (const it of items) {
      const row = parseDraw(it);
      if (row['期号'] && !local.has(row['期号'])) {
        local.set(row['期号'], row);
        pageNew++;
      }
    }
    newCount += pageNew;
    console.log(`第 ${page} 页：${items.length} 条，新增 ${pageNew} 条（${items[items.length - 1].lotteryDrawNum} 期）`);
    if (pageNew === 0) { console.log('已追平本地数据，停止。'); break; }
    page++;
    await sleep(500); // 礼貌限速
  }
  saveAll(local);
  console.log(`\n完成：新增 ${newCount} 期，本地共 ${local.size} 期。`);
  console.log(`CSV: ${CSV_PATH}\nJSON: ${JSON_PATH}`);
}

function cmdShow(args) {
  const local = loadLocal();
  if (!local.size) { console.log('本地暂无数据，请先运行: node dlt.js fetch'); return; }
  const ordered = [...local.values()].sort((a, b) => b['期号'].localeCompare(a['期号']));
  const limit = args.limit || 10;
  console.log(`本地共 ${ordered.length} 期，最新 ${limit} 期：`);
  printRows(ordered.slice(0, limit));
}

function cmdStats(args) {
  const local = loadLocal();
  if (!local.size) { console.log('本地暂无数据，请先运行: node dlt.js fetch'); return; }
  const limit = args.limit || 100;
  const ordered = [...local.values()]
    .sort((a, b) => b['期号'].localeCompare(a['期号']))
    .slice(0, limit);
  const frontCnt = new Map(), backCnt = new Map();
  const bump = (m, k) => { if (k) m.set(k, (m.get(k) || 0) + 1); };
  for (const r of ordered) {
    for (let i = 1; i <= 5; i++) bump(frontCnt, r[`前区${i}`]);
    for (let i = 1; i <= 2; i++) bump(backCnt, r[`后区${i}`]);
  }
  const fmt = (arr) => arr.map(([n, c]) => `${n}(${c})`).join('  ');
  const byHot = (m) => [...m.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
  const byCold = (m) => [...m.entries()].sort((a, b) => a[1] - b[1] || a[0].localeCompare(b[0]));
  console.log(`统计范围：最近 ${ordered.length} 期（${ordered[ordered.length - 1]['期号']} ~ ${ordered[0]['期号']}）\n`);
  console.log('前区热号 TOP10：', fmt(byHot(frontCnt).slice(0, 10)));
  console.log('前区冷号 TOP10：', fmt(byCold(frontCnt).slice(0, 10)));
  console.log('后区热号 TOP5： ', fmt(byHot(backCnt).slice(0, 5)));
  console.log('后区冷号 TOP5： ', fmt(byCold(backCnt).slice(0, 5)));
}

// ---------- 入口 ----------

function parseArgs(argv) {
  const args = { _: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a.startsWith('--')) {
      const key = a.slice(2);
      const next = argv[i + 1];
      if (next !== undefined && !next.startsWith('--')) {
        args[key] = isNaN(Number(next)) ? next : Number(next);
        i++;
      } else {
        args[key] = true;
      }
    } else {
      args._.push(a);
    }
  }
  return args;
}

function usage() {
  console.log(`大乐透开奖数据收集器（数据源：中国体彩网）

用法：node dlt.js <命令> [参数]

命令：
  latest                     查看最新一期开奖结果并入库
  fetch [--pages 3] [--all]  抓取历史数据（默认 3 页，--all 全量）
  show [--limit 10]          查看本地已存储数据
  stats [--limit 100]        号码出现频率统计`);
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const cmd = args._[0];
  try {
    if (cmd === 'latest') await cmdLatest();
    else if (cmd === 'fetch') await cmdFetch(args);
    else if (cmd === 'show') cmdShow(args);
    else if (cmd === 'stats') cmdStats(args);
    else usage();
  } catch (e) {
    console.error(`执行失败：${e.message}`);
    process.exit(1);
  }
}

main();
