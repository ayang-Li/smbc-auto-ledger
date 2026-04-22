/**
 * emailTemplates.gs — SMBC 月报 / 周报 HTML 邮件模板
 *
 * 按月份自动轮换主题（同一双月，周报 + 月报同主题）：
 *   1–2  月 → c6  八十年代账本（年关清账感）
 *   3–4  月 → c4  古诗主题（春意）
 *   5–6  月 → c1  自然现代宋体（明快）
 *   7–8  月 → a   纯净英文新闻纸（盛夏冷静，白底）
 *   9–10 月 → c2  中文报纸社论（秋日读报）
 *   11–12月 → e7  维多利亚粗 slab（入冬厚重）
 *
 * 入口（由 code.gs sendEmailReport_ 调用）：
 *   pickTemplateId_(date)                 → 'c6'|'c4'|'c1'|'a'|'c2'|'e7'
 *   buildViewModel_(summary, includeFixed)→ 视图数据对象
 *   buildPlainText_(vm, title)            → 纯文本 body
 *   buildEmailHtml_(vm, title, tplId)     → HTML body
 *
 * Gmail 兼容要点：
 *   - 所有 CSS inline，不用 <style>（Gmail 移动端会剥离 head 样式）
 *   - 布局用 <table>，不用 flex
 *   - 只依赖 iOS/macOS/Windows 系统字体栈（不 @font-face）
 *   - 无 JS / button / onclick
 *   - 邮件末尾附 <pre> 等宽"纯文本长按复制区"
 */

/* ========= 轮换调度 ========= */

const EMAIL_TEMPLATES = ['c6', 'c4', 'c1', 'a', 'c2', 'e7'];

function pickTemplateId_(date) {
  return EMAIL_TEMPLATES[Math.floor(date.getMonth() / 2)];
}


/* ========= 视图模型 ========= */

function buildViewModel_(summary, includeFixed) {
  const days = Math.round((summary.end - summary.start) / 86400000);
  const avgDaily = days > 0 ? Math.round(summary.total / days) : 0;

  const grouped = {};
  const used = new Set();
  CATEGORY_GROUPS.forEach(function (g) {
    const amt = g.cats.reduce(function (s, c) { return s + (summary.byCat[c] || 0); }, 0);
    if (amt !== 0) { grouped[g.label] = amt; g.cats.forEach(function (c) { used.add(c); }); }
  });
  Object.entries(summary.byCat).forEach(function (kv) {
    if (!used.has(kv[0]) && kv[1] !== 0) grouped[kv[0]] = kv[1];
  });
  const sortedGroups = Object.entries(grouped).sort(function (a, b) { return b[1] - a[1]; });

  let fixedTotal = 0;
  if (includeFixed) FIXED_COSTS.forEach(function (f) { fixedTotal += f.amount; });
  const grandTotal = summary.total + (includeFixed ? fixedTotal : 0);

  const startStr = Utilities.formatDate(summary.start, TIME_ZONE, 'M月d日');
  const endStr   = Utilities.formatDate(new Date(summary.end - 1), TIME_ZONE, 'M月d日');
  const startDate = new Date(summary.start);

  return {
    total: summary.total,
    days: days,
    avgDaily: avgDaily,
    sortedGroups: sortedGroups,
    fixedCosts: includeFixed ? FIXED_COSTS : [],
    fixedTotal: fixedTotal,
    grandTotal: grandTotal,
    periodStr: startStr + ' — ' + endStr,
    startStr: startStr,
    endStr: endStr,
    includeFixed: includeFixed,
    year: startDate.getFullYear(),
    month: startDate.getMonth() + 1
  };
}


/* ========= 纯文本（作为 plain body 和底部复制区的内容） ========= */

function buildPlainText_(vm, title) {
  const padEnd = function (s, n) {
    let w = 0;
    const str = String(s);
    for (let i = 0; i < str.length; i++) {
      w += (str.charCodeAt(i) > 127 ? 2 : 1);
    }
    return str + ' '.repeat(Math.max(0, n - w));
  };
  const padStart = function (s, n) {
    let w = 0;
    const str = String(s);
    for (let i = 0; i < str.length; i++) {
      w += (str.charCodeAt(i) > 127 ? 2 : 1);
    }
    return ' '.repeat(Math.max(0, n - w)) + str;
  };

  const W = 16;
  const L = [];
  L.push(title);
  L.push('周期: ' + vm.periodStr);
  L.push('');
  L.push(padEnd(vm.includeFixed ? '月合计' : '周合计', W) + padStart(_yen(vm.total), 12));
  L.push(padEnd('日均', W) + padStart(_yen(vm.avgDaily), 12));
  L.push('');
  L.push('─── 分类明细 ───');
  vm.sortedGroups.forEach(function (entry) {
    const label = entry[0], amt = entry[1];
    const pct = _pct(amt, vm.total);
    L.push(padEnd(label, W) + padStart(_yen(amt), 12) + '   ' + pct + '%');
  });

  if (vm.includeFixed && vm.fixedCosts.length > 0) {
    L.push('');
    L.push('─── 固定估算（未入账）───');
    vm.fixedCosts.forEach(function (f) {
      L.push(padEnd(f.label, W) + padStart(_yen(f.amount), 12));
    });
    L.push(padEnd('固定小计', W) + padStart(_yen(vm.fixedTotal), 12));
    L.push('');
    L.push(padEnd('本月总计（含估算）', W) + padStart(_yen(vm.grandTotal), 12));
  }

  L.push('');
  L.push('—— auto-generated ——');
  return L.join('\n');
}


/* ========= 路由 ========= */

function buildEmailHtml_(vm, title, tplId) {
  const map = {
    a:  buildHtml_a_,
    c1: buildHtml_c1_,
    c2: buildHtml_c2_,
    c4: buildHtml_c4_,
    c6: buildHtml_c6_,
    e7: buildHtml_e7_
  };
  const fn = map[tplId] || buildHtml_a_;
  return _wrapHtml_(fn(vm, title));
}


/* ========= 公共 helper ========= */

function _yen(n) { return '¥' + Number(n).toLocaleString(); }

function _pct(part, total) {
  return total > 0 ? ((part / total) * 100).toFixed(0) : '0';
}

function _esc(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function _pad2(n) { return n < 10 ? '0' + n : String(n); }

function _monthEn(m) {
  return ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][m - 1];
}

function _monthZh(m) {
  return ['一','二','三','四','五','六','七','八','九','十','十一','十二'][m - 1];
}

function _wrapHtml_(inner) {
  return '<!DOCTYPE html><html><head>' +
         '<meta name="viewport" content="width=device-width, initial-scale=1">' +
         '<meta name="color-scheme" content="light">' +
         '<meta name="supported-color-schemes" content="light">' +
         '</head><body style="margin:0; padding:0; background:#f2f2f2;">' +
         '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f2f2f2;"><tr><td align="center">' +
         '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width:560px;"><tr><td>' +
         inner +
         '</td></tr></table></td></tr></table></body></html>';
}

function _copyBox_(plainText, opts) {
  const o = opts || {};
  const bg         = o.bg         || '#ffffff';
  const border     = o.border     || '#333333';
  const label      = o.label      || '▎ 纯文本版 · 长按任意处 → 全选 → 复制';
  const labelColor = o.labelColor || border;
  const labelFont  = o.labelFont  || "-apple-system,'PingFang SC','Helvetica Neue',sans-serif";
  const textColor  = o.textColor  || '#1a1a1a';
  return '<div style="margin-top:28px; padding:14px 16px; background:' + bg + '; border:1.5px dashed ' + border + '; border-radius:2px;">' +
         '<div style="font-family:' + labelFont + '; font-size:11px; letter-spacing:0.2em; color:' + labelColor + '; margin-bottom:10px; font-weight:700;">' + label + '</div>' +
         '<pre style="margin:0; font-family:\'Courier New\',\'Consolas\',monospace; font-size:12px; line-height:1.75; color:' + textColor + '; white-space:pre-wrap; word-wrap:break-word;">' + _esc(plainText) + '</pre>' +
         '</div>';
}


/* ========= A · 纯净英文新闻纸（白底 WSJ 风） ========= */

function buildHtml_a_(vm, title) {
  const plain = buildPlainText_(vm, title);

  const rows = vm.sortedGroups.map(function (entry, i, arr) {
    const label = entry[0], amt = entry[1];
    const sep = i < arr.length - 1 ? ' border-bottom:1px dotted #8a7355;' : '';
    return '<tr>' +
           '<td style="padding:11px 0;' + sep + '">' + _esc(label) + '</td>' +
           '<td align="right" style="' + sep + '">' + _yen(amt) + '</td>' +
           '<td align="right" style="color:#5a4a30; font-style:italic; width:52px;' + sep + '">' + _pct(amt, vm.total) + '%</td>' +
           '</tr>';
  }).join('');

  const fixedRows = vm.fixedCosts.map(function (f) {
    return '<tr><td style="padding:6px 0;">' + _esc(f.label) + '</td>' +
           '<td align="right" style="font-style:italic;">' + _yen(f.amount) + '</td></tr>';
  }).join('');

  const fixedBlock = vm.includeFixed ?
    '<div style="border-top:1px solid #1a1a1a; margin-top:20px; padding-top:16px;">' +
    '<div style="font-size:12px; letter-spacing:0.3em; color:#5a4a30; margin-bottom:10px;">FIXED · ESTIMATED</div>' +
    '<table width="100%" style="font-size:15px;">' + fixedRows + '</table>' +
    '</div>' : '';

  const grandBlock = vm.includeFixed ?
    '<div style="text-align:center; margin-top:20px; padding:16px 0; border-top:3px solid #1a1a1a; border-bottom:3px solid #1a1a1a;">' +
    '<div style="font-size:12px; letter-spacing:0.3em; color:#5a4a30;">GRAND TOTAL</div>' +
    '<div style="font-size:30px; font-weight:700; margin-top:4px;">' + _yen(vm.grandTotal) + '</div>' +
    '</div>' : '';

  const sumLabel = vm.includeFixed ? 'TOTAL' : 'WEEK';
  const leadLine = vm.includeFixed
    ? _monthEn(vm.month) + ' Spending Settles<br>at ' + _yen(vm.total)
    : 'This Week at ' + _yen(vm.total);

  return '<div style="padding:32px 24px; font-family:Georgia,\'Times New Roman\',serif; color:#1a1a1a; background:#ffffff;">' +
         '<div style="text-align:center; border-top:3px solid #1a1a1a; border-bottom:1px solid #1a1a1a; padding:14px 0;">' +
         '<div style="font-size:11px; letter-spacing:0.4em; color:#5a4a30;">VOL. III · NO. ' + _pad2(vm.month) + '</div>' +
         '<div style="font-size:34px; font-weight:700; margin:6px 0; font-style:italic; letter-spacing:-0.01em; line-height:1;">The Monthly Ledger</div>' +
         '<div style="font-size:12px; letter-spacing:0.25em; color:#5a4a30;">' + vm.periodStr + '</div>' +
         '</div>' +

         '<div style="text-align:center; margin:24px 0 18px;">' +
         '<div style="font-size:12px; letter-spacing:0.3em; color:#5a4a30;">LEAD</div>' +
         '<div style="font-size:22px; font-weight:700; font-style:italic; margin-top:6px; line-height:1.3;">' + leadLine + '</div>' +
         '</div>' +

         '<table width="100%" style="border-top:1px solid #1a1a1a; border-bottom:1px solid #1a1a1a; margin-bottom:20px;"><tr>' +
         '<td align="center" style="width:50%; border-right:1px solid #bba487; padding:18px 0;">' +
         '<div style="font-size:12px; letter-spacing:0.25em; color:#5a4a30;">' + sumLabel + '</div>' +
         '<div style="font-size:26px; font-weight:700; margin-top:4px;">' + _yen(vm.total) + '</div>' +
         '</td>' +
         '<td align="center" style="width:50%; padding:18px 0;">' +
         '<div style="font-size:12px; letter-spacing:0.25em; color:#5a4a30;">DAILY</div>' +
         '<div style="font-size:26px; font-weight:700; margin-top:4px;">' + _yen(vm.avgDaily) + '</div>' +
         '</td></tr></table>' +

         '<div style="font-size:12px; letter-spacing:0.3em; color:#5a4a30; margin-bottom:12px;">CATEGORY BREAKDOWN</div>' +
         '<table width="100%" style="font-size:16px;">' + rows + '</table>' +

         fixedBlock +
         grandBlock +

         '<div style="text-align:center; margin-top:16px; font-size:11px; letter-spacing:0.25em; color:#5a4a30; font-style:italic;">— auto-generated —</div>' +

         _copyBox_(plain, { bg: '#fafafa', border: '#1a1a1a', labelColor: '#5a4a30' }) +
         '</div>';
}


/* ========= C1 · 自然现代宋体 ========= */

function buildHtml_c1_(vm, title) {
  const plain = buildPlainText_(vm, title);

  const rows = vm.sortedGroups.map(function (entry, i, arr) {
    const label = entry[0], amt = entry[1];
    const sep = i < arr.length - 1 ? ' border-bottom:1px dotted #8a7355;' : '';
    return '<tr>' +
           '<td style="padding:11px 0;' + sep + '">' + _esc(label) + '</td>' +
           '<td align="right" style="' + sep + '">' + _yen(amt) + '</td>' +
           '<td align="right" style="color:#5a4a30; font-style:italic; width:52px;' + sep + '">' + _pct(amt, vm.total) + '%</td>' +
           '</tr>';
  }).join('');

  const fixedRows = vm.fixedCosts.map(function (f) {
    return '<tr><td style="padding:6px 0;">' + _esc(f.label) + '</td>' +
           '<td align="right">' + _yen(f.amount) + '</td></tr>';
  }).join('');

  const fixedBlock = vm.includeFixed ?
    '<div style="border-top:1px solid #1a1a1a; margin-top:20px; padding-top:16px;">' +
    '<div style="font-size:12px; letter-spacing:0.25em; color:#5a4a30; margin-bottom:10px;">—— 固定估算（未入账） ——</div>' +
    '<table width="100%" style="font-size:15px;">' + fixedRows + '</table>' +
    '</div>' : '';

  const grandBlock = vm.includeFixed ?
    '<div style="text-align:center; margin-top:20px; padding:16px 0; border-top:3px solid #1a1a1a; border-bottom:3px solid #1a1a1a;">' +
    '<div style="font-size:12px; letter-spacing:0.25em; color:#5a4a30;">本月总计</div>' +
    '<div style="font-size:30px; font-weight:700; margin-top:4px;">' + _yen(vm.grandTotal) + '</div>' +
    '</div>' : '';

  const sumLabel = vm.includeFixed ? '月合计' : '周合计';
  const leadLine = vm.includeFixed
    ? vm.month + ' 月共支出 ' + _yen(vm.total)
    : '上周共支出 ' + _yen(vm.total);
  const subtitle = '日均 ' + _yen(vm.avgDaily);

  return '<div style="padding:32px 24px; font-family:\'Songti SC\',\'STSong\',\'SimSun\',serif; color:#1a1a1a; background:#f4ede0;">' +
         '<div style="text-align:center; border-top:3px solid #1a1a1a; border-bottom:1px solid #1a1a1a; padding:14px 0;">' +
         '<div style="font-size:12px; letter-spacing:0.3em; color:#5a4a30;">' + vm.year + ' 年 · 第 ' + _monthZh(vm.month) + ' 期</div>' +
         '<div style="font-size:34px; font-weight:700; margin:6px 0; letter-spacing:0.1em; line-height:1.1;">月度账本</div>' +
         '<div style="font-size:13px; letter-spacing:0.2em; color:#5a4a30;">' + vm.periodStr + '</div>' +
         '</div>' +

         '<div style="text-align:center; margin:24px 0 18px;">' +
         '<div style="font-size:12px; letter-spacing:0.3em; color:#5a4a30;">—— 本期要点 ——</div>' +
         '<div style="font-size:21px; font-weight:700; margin-top:8px; line-height:1.4;">' + leadLine + '</div>' +
         '<div style="font-size:13px; color:#5a4a30; margin-top:6px;">' + subtitle + '</div>' +
         '</div>' +

         '<table width="100%" style="border-top:1px solid #1a1a1a; border-bottom:1px solid #1a1a1a; margin-bottom:20px;"><tr>' +
         '<td align="center" style="width:50%; border-right:1px solid #bba487; padding:18px 0;">' +
         '<div style="font-size:12px; letter-spacing:0.2em; color:#5a4a30;">' + sumLabel + '</div>' +
         '<div style="font-size:26px; font-weight:700; margin-top:4px;">' + _yen(vm.total) + '</div>' +
         '</td>' +
         '<td align="center" style="width:50%; padding:18px 0;">' +
         '<div style="font-size:12px; letter-spacing:0.2em; color:#5a4a30;">日均</div>' +
         '<div style="font-size:26px; font-weight:700; margin-top:4px;">' + _yen(vm.avgDaily) + '</div>' +
         '</td></tr></table>' +

         '<div style="font-size:12px; letter-spacing:0.25em; color:#5a4a30; margin-bottom:12px;">—— 分类明细 ——</div>' +
         '<table width="100%" style="font-size:16px;">' + rows + '</table>' +

         fixedBlock +
         grandBlock +

         '<div style="text-align:center; margin-top:16px; font-size:11px; letter-spacing:0.2em; color:#5a4a30;">—— 自动生成 ——</div>' +

         _copyBox_(plain, { bg: '#fafafa', border: '#1a1a1a', labelColor: '#5a4a30' }) +
         '</div>';
}


/* ========= C2 · 中文报纸社论 ========= */

function buildHtml_c2_(vm, title) {
  const plain = buildPlainText_(vm, title);

  const rows = vm.sortedGroups.map(function (entry, i, arr) {
    const label = entry[0], amt = entry[1];
    const sep = i < arr.length - 1 ? ' border-bottom:1px dotted #a88c60;' : '';
    return '<tr>' +
           '<td style="padding:11px 0;' + sep + '">' + _esc(label) + '</td>' +
           '<td align="right" style="font-weight:700;' + sep + '">' + _yen(amt) + '</td>' +
           '<td align="right" style="color:#8b1a1a; font-weight:700; width:52px;' + sep + '">' + _pct(amt, vm.total) + '%</td>' +
           '</tr>';
  }).join('');

  const fixedRows = vm.fixedCosts.map(function (f) {
    return '<tr><td style="padding:6px 0;">' + _esc(f.label) + '</td>' +
           '<td align="right">' + _yen(f.amount) + '</td></tr>';
  }).join('');

  const fixedBlock = vm.includeFixed ?
    '<div style="margin-top:16px; background:#5a4230; color:#f4ede0; padding:5px 12px; font-size:12px; letter-spacing:0.3em;">■ 固定估算 · 未入账</div>' +
    '<table width="100%" style="font-size:15px; margin-top:8px;">' + fixedRows + '</table>' : '';

  const grandBlock = vm.includeFixed ?
    '<div style="text-align:center; margin-top:18px; padding:14px 0; border-top:5px solid #1a1410; border-bottom:1px solid #1a1410;">' +
    '<div style="font-size:12px; letter-spacing:0.3em; color:#5a4230;">· 本月总计 ·</div>' +
    '<div style="font-size:32px; font-weight:900; margin-top:4px;">' + _yen(vm.grandTotal) + '</div>' +
    '</div>' : '';

  const sumLabel = vm.includeFixed ? '月合计' : '周合计';
  const lead = vm.includeFixed
    ? vm.month + ' 月消费 ' + _yen(vm.total)
    : '上周消费 ' + _yen(vm.total);

  return '<div style="padding:32px 24px; font-family:\'Songti SC\',\'STSong\',\'SimSun\',serif; color:#1a1410; background:#f4ede0;">' +
         '<div style="text-align:center; border-top:5px solid #1a1410; padding-top:12px;">' +
         '<div style="font-size:12px; letter-spacing:0.35em; color:#5a4230;">创刊于 二〇二三</div>' +
         '</div>' +
         '<div style="text-align:center; border-bottom:1px solid #1a1410; padding-bottom:12px;">' +
         '<div style="font-size:42px; font-weight:900; margin:6px 0; letter-spacing:0.15em; line-height:1;">月 计 报</div>' +
         '<div style="font-size:11px; letter-spacing:0.3em; color:#5a4230; margin-top:12px;">' + vm.year + ' 年 · 第 ' + _monthZh(vm.month) + ' 期 · ' + vm.periodStr + '</div>' +
         '</div>' +

         '<div style="background:#1a1410; color:#f4ede0; padding:6px 12px;">' +
         '<table width="100%"><tr>' +
         '<td style="font-size:11px; letter-spacing:0.25em;">▌本期要目</td>' +
         '<td align="right" style="font-size:11px; letter-spacing:0.25em;">定价 无</td>' +
         '</tr></table></div>' +

         '<div style="text-align:center; margin:22px 0 16px;">' +
         '<div style="font-size:12px; letter-spacing:0.3em; color:#8b1a1a; font-weight:700;">▎消费观察</div>' +
         '<div style="font-size:22px; font-weight:900; margin-top:6px; line-height:1.3;">' + lead + '</div>' +
         '<div style="font-size:13px; color:#5a4230; margin-top:6px;">日均支出 ' + _yen(vm.avgDaily) + '</div>' +
         '</div>' +

         '<table width="100%" style="border-top:1px solid #1a1410; border-bottom:1px solid #1a1410; margin-bottom:18px;"><tr>' +
         '<td align="center" style="width:50%; border-right:1px solid #a88c60; padding:16px 0;">' +
         '<div style="font-size:12px; letter-spacing:0.25em; color:#5a4230;">' + sumLabel + '</div>' +
         '<div style="font-size:26px; font-weight:900; margin-top:4px;">' + _yen(vm.total) + '</div>' +
         '</td>' +
         '<td align="center" style="width:50%; padding:16px 0;">' +
         '<div style="font-size:12px; letter-spacing:0.25em; color:#5a4230;">日均</div>' +
         '<div style="font-size:26px; font-weight:900; margin-top:4px;">' + _yen(vm.avgDaily) + '</div>' +
         '</td></tr></table>' +

         '<div style="background:#1a1410; color:#f4ede0; padding:5px 12px; font-size:12px; letter-spacing:0.3em;">■ 分类详情</div>' +
         '<table width="100%" style="font-size:16px; margin-top:8px;">' + rows + '</table>' +

         fixedBlock +
         grandBlock +

         '<div style="text-align:center; margin-top:14px; font-size:11px; letter-spacing:0.3em; color:#5a4230;">—— 截稿 ——</div>' +

         _copyBox_(plain, { bg: '#fafafa', border: '#8b1a1a', labelColor: '#8b1a1a' }) +
         '</div>';
}


/* ========= C4 · 古诗主题（楷体 + 宣纸） ========= */

function buildHtml_c4_(vm, title) {
  const plain = buildPlainText_(vm, title);

  const rows = vm.sortedGroups.map(function (entry, i, arr) {
    const label = entry[0], amt = entry[1];
    const sep = i < arr.length - 1 ? ' border-bottom:1px dotted #b89368;' : '';
    return '<tr>' +
           '<td style="padding:12px 0;' + sep + '">' + _esc(label) + '</td>' +
           '<td align="right" style="' + sep + '">' + _yen(amt) + '</td>' +
           '<td align="right" style="color:#a01818; width:52px;' + sep + '">' + _pct(amt, vm.total) + '%</td>' +
           '</tr>';
  }).join('');

  const fixedRows = vm.fixedCosts.map(function (f) {
    return '<tr><td style="padding:6px 0;">' + _esc(f.label) + '</td>' +
           '<td align="right">' + _yen(f.amount) + '</td></tr>';
  }).join('');

  const fixedBlock = vm.includeFixed ?
    '<div style="border-top:1px solid #2a1810; margin-top:20px; padding-top:16px;">' +
    '<div style="font-size:12px; letter-spacing:0.35em; color:#7a4a28; margin-bottom:10px; text-align:center;">❦ 隱費（不錄於冊） ❦</div>' +
    '<table width="100%" style="font-size:15px;">' + fixedRows + '</table>' +
    '</div>' : '';

  const grandBlock = vm.includeFixed ?
    '<div style="text-align:center; margin-top:20px; padding:18px 0; border-top:1px solid #2a1810; border-bottom:1px solid #2a1810;">' +
    '<div style="font-size:12px; letter-spacing:0.4em; color:#7a4a28;">—— 總　計 ——</div>' +
    '<div style="font-size:32px; font-weight:700; margin-top:6px;">' + _yen(vm.grandTotal) + '</div>' +
    '<div style="font-size:13px; color:#7a4a28; margin-top:10px; letter-spacing:0.12em; font-style:italic;">金風玉露一相逢</div>' +
    '</div>' : '';

  const poem = _c4Poem_(vm.month);
  const sumLabel = vm.includeFixed ? '月　計' : '周　計';

  return '<div style="padding:34px 26px; font-family:\'Kaiti SC\',\'STKaiti\',\'KaiTi\',\'Songti SC\',serif; color:#2a1810; background:#f1e5cb;">' +
         '<div style="text-align:center; border-top:1px solid #2a1810; border-bottom:1px solid #2a1810; padding:16px 0;">' +
         '<div style="font-size:12px; letter-spacing:0.4em; color:#7a4a28;">' + vm.year + ' 年 · 第 ' + _monthZh(vm.month) + ' 期</div>' +
         '<div style="font-size:42px; font-weight:700; margin:10px 0; letter-spacing:0.3em; line-height:1;">' + _monthZh(vm.month) + ' 月 記</div>' +
         '<div style="font-size:13px; letter-spacing:0.3em; color:#7a4a28; font-style:italic;">' + vm.periodStr + '</div>' +
         '</div>' +

         '<div style="text-align:center; margin:24px 0 18px; padding:0 10px;">' +
         '<div style="font-size:17px; font-weight:400; line-height:2; letter-spacing:0.12em; color:#2a1810;">' +
         poem.line1 + '<br>' +
         '<span style="color:#a01818;">' + poem.line2 + '</span>' +
         '</div>' +
         '<div style="font-size:14px; color:#7a4a28; margin-top:10px; letter-spacing:0.1em;">是月之費 ' + _yen(vm.total) + '</div>' +
         '</div>' +

         '<table width="100%" style="border-top:1px solid #2a1810; border-bottom:1px solid #2a1810; margin-bottom:20px;"><tr>' +
         '<td align="center" style="width:50%; border-right:1px solid #b89368; padding:16px 0;">' +
         '<div style="font-size:12px; letter-spacing:0.3em; color:#7a4a28;">' + sumLabel + '</div>' +
         '<div style="font-size:26px; font-weight:700; margin-top:4px;">' + _yen(vm.total) + '</div>' +
         '</td>' +
         '<td align="center" style="width:50%; padding:16px 0;">' +
         '<div style="font-size:12px; letter-spacing:0.3em; color:#7a4a28;">日　凡</div>' +
         '<div style="font-size:26px; font-weight:700; margin-top:4px;">' + _yen(vm.avgDaily) + '</div>' +
         '</td></tr></table>' +

         '<div style="font-size:12px; letter-spacing:0.35em; color:#7a4a28; margin-bottom:14px; text-align:center;">❦ 諸項 ❦</div>' +
         '<table width="100%" style="font-size:16px;">' + rows + '</table>' +

         fixedBlock +
         grandBlock +

         '<div style="text-align:center; margin-top:16px; font-size:12px; letter-spacing:0.4em; color:#7a4a28;">—— 自 記 ——</div>' +

         _copyBox_(plain, { bg: '#fbf5e1', border: '#a01818', labelColor: '#7a4a28' }) +
         '</div>';
}

// 每月对应诗句（按轮换表，c4 主要命中 3-4 月；其他月份作 fallback）
function _c4Poem_(m) {
  const map = {
    1:  { line1: '爆竹聲中一歲除', line2: '新春伊始 吾理舊帳' },
    2:  { line1: '不知細葉誰裁出', line2: '二月春風 吾耗銀兩' },
    3:  { line1: '春風又綠江南岸', line2: '當此之時 吾耗銀兩' },
    4:  { line1: '清明時節雨紛紛', line2: '當此之時 吾耗銀兩' },
    5:  { line1: '綠樹陰濃夏日長', line2: '初夏之月 吾耗銀兩' },
    6:  { line1: '接天蓮葉無窮碧', line2: '仲夏將至 吾耗銀兩' },
    7:  { line1: '小荷才露尖尖角', line2: '盛夏之月 吾耗銀兩' },
    8:  { line1: '荷風送香氣', line2: '三伏將盡 吾耗銀兩' },
    9:  { line1: '空山新雨後', line2: '初秋之月 吾耗銀兩' },
    10: { line1: '停車坐愛楓林晚', line2: '深秋已至 吾耗銀兩' },
    11: { line1: '霜葉紅於二月花', line2: '霜月之期 吾耗銀兩' },
    12: { line1: '瑞雪兆豐年', line2: '歲末將至 吾理此帳' }
  };
  return map[m] || map[3];
}


/* ========= C6 · 八十年代账本（红格线 + 仿宋） ========= */

function buildHtml_c6_(vm, title) {
  const plain = buildPlainText_(vm, title);

  const bodyRows = vm.sortedGroups.map(function (entry, i) {
    const label = entry[0], amt = entry[1];
    const bg = (i % 2 === 0) ? '#fbf5e1' : 'transparent';
    return '<tr style="background:' + bg + ';">' +
           '<td style="padding:10px; border-bottom:1px solid #8b1a1a; border-right:1px dotted #8b1a1a;">' + _esc(label) + '</td>' +
           '<td style="padding:10px; text-align:right; border-bottom:1px solid #8b1a1a; border-right:1px dotted #8b1a1a; font-weight:700;">' + Number(amt).toLocaleString() + '</td>' +
           '<td style="padding:10px; text-align:right; border-bottom:1px solid #8b1a1a; color:#8b1a1a; font-weight:700;">' + _pct(amt, vm.total) + '%</td>' +
           '</tr>';
  }).join('');

  const fixedRows = vm.fixedCosts.map(function (f) {
    return '<tr><td style="padding:4px 0;">' + _esc(f.label) + '</td>' +
           '<td align="right">' + _yen(f.amount) + '</td></tr>';
  }).join('');

  const fixedBlock = vm.includeFixed ?
    '<div style="margin-top:14px; border:1px dashed #8b1a1a; padding:10px 12px;">' +
    '<div style="font-size:12px; letter-spacing:0.2em; color:#8b1a1a; margin-bottom:6px;">※ 估算项（未入账）</div>' +
    '<table width="100%" style="font-size:14px;">' + fixedRows +
    '<tr style="border-top:1px solid #8b1a1a;"><td style="padding:7px 0; font-weight:700;">估算合计</td><td align="right" style="font-weight:700;">' + _yen(vm.fixedTotal) + '</td></tr>' +
    '</table></div>' : '';

  const grandBlock = vm.includeFixed ?
    '<div style="margin-top:14px; background:#1a1410; color:#f3ead0; padding:12px 16px;">' +
    '<table width="100%"><tr>' +
    '<td style="font-size:13px; letter-spacing:0.25em;">本月合计（含估算）</td>' +
    '<td align="right" style="font-size:26px; font-weight:700;">' + _yen(vm.grandTotal) + '</td>' +
    '</tr></table></div>' : '';

  const reportTitle = vm.includeFixed ? '个人消费月报表' : '个人消费周报表';

  return '<div style="padding:24px 20px; font-family:\'Fangsong SC\',\'STFangsong\',\'FangSong\',\'Songti SC\',serif; color:#1a1410; background:#f3ead0;">' +
         '<div style="text-align:center;">' +
         '<div style="font-size:11px; letter-spacing:0.3em; color:#8b1a1a;">家　庭　财　务</div>' +
         '<div style="font-size:28px; font-weight:700; margin:6px 0; letter-spacing:0.25em; color:#1a1410;">' + reportTitle + '</div>' +
         '<div style="font-size:11px; letter-spacing:0.2em; color:#5a4230;">编号：SMBC－' + vm.year + '－' + _pad2(vm.month) + '　　表样：甲式</div>' +
         '</div>' +

         '<div style="border:2px solid #8b1a1a; margin-top:14px; background:#fbf5e1;">' +
         '<div style="background:#8b1a1a; color:#fbf5e1; padding:6px 12px; font-size:12px; letter-spacing:0.2em;">■ 填报说明</div>' +
         '<table width="100%" style="font-size:13px;">' +
         '<tr><td style="padding:5px 12px; border-bottom:1px dotted #8b1a1a; width:80px;">报表期间</td><td style="padding:5px 12px; border-bottom:1px dotted #8b1a1a;">' + vm.periodStr + '</td></tr>' +
         '<tr><td style="padding:5px 12px; border-bottom:1px dotted #8b1a1a;">填 报 人</td><td style="padding:5px 12px; border-bottom:1px dotted #8b1a1a;">SMBC 自动记账</td></tr>' +
         '<tr><td style="padding:5px 12px; border-bottom:1px dotted #8b1a1a;">单　　位</td><td style="padding:5px 12px; border-bottom:1px dotted #8b1a1a;">日元（円）</td></tr>' +
         '<tr><td style="padding:5px 12px;">出表日期</td><td style="padding:5px 12px;">' + Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy年 M月 d日') + '</td></tr>' +
         '</table></div>' +

         '<div style="margin-top:16px; border:2px solid #1a1410;">' +
         '<table width="100%" style="font-size:14px; border-collapse:collapse;">' +
         '<thead style="background:#8b1a1a; color:#fbf5e1;">' +
         '<tr>' +
         '<th style="padding:8px 10px; text-align:left; font-weight:400; letter-spacing:0.15em; font-size:12px; border-right:1px solid #fbf5e1;">项目</th>' +
         '<th style="padding:8px 10px; text-align:right; font-weight:400; letter-spacing:0.1em; font-size:12px; border-right:1px solid #fbf5e1;">金额（円）</th>' +
         '<th style="padding:8px 10px; text-align:right; font-weight:400; letter-spacing:0.1em; font-size:12px; width:52px;">占比</th>' +
         '</tr></thead>' +
         '<tbody>' + bodyRows + '</tbody>' +
         '<tfoot><tr style="background:#8b1a1a; color:#fbf5e1;">' +
         '<td style="padding:10px; font-weight:700; letter-spacing:0.2em; border-right:1px solid #fbf5e1;">小　计</td>' +
         '<td style="padding:10px; text-align:right; font-weight:700; border-right:1px solid #fbf5e1;">' + Number(vm.total).toLocaleString() + '</td>' +
         '<td style="padding:10px; text-align:right; font-weight:700;">100%</td>' +
         '</tr></tfoot></table></div>' +

         fixedBlock +
         grandBlock +

         '<div style="margin-top:14px; border-top:2px solid #8b1a1a; padding-top:10px;">' +
         '<table width="100%" style="font-size:12px; color:#5a4230;"><tr>' +
         '<td>经办：<span style="border-bottom:1px solid #5a4230; padding:0 22px;">GAS</span></td>' +
         '<td>复核：<span style="border-bottom:1px solid #5a4230; padding:0 22px;">自动</span></td>' +
         '<td align="right">' + Utilities.formatDate(new Date(), TIME_ZONE, 'MM/dd') + '</td>' +
         '</tr></table></div>' +

         _copyBox_(plain, { bg: '#fbf5e1', border: '#8b1a1a', labelColor: '#8b1a1a' }) +
         '</div>';
}


/* ========= E7 · 维多利亚粗 slab ========= */

function buildHtml_e7_(vm, title) {
  const plain = buildPlainText_(vm, title);

  const rows = vm.sortedGroups.map(function (entry, i, arr) {
    const label = entry[0], amt = entry[1];
    const sep = i < arr.length - 1 ? ' border-bottom:1px dotted #a88c60;' : '';
    return '<tr>' +
           '<td style="padding:11px 0;' + sep + '">' + _esc(label) + '</td>' +
           '<td align="right" style="font-weight:700;' + sep + '">' + _yen(amt) + '</td>' +
           '<td align="right" style="color:#5a4230; font-style:italic; width:52px;' + sep + '">' + _pct(amt, vm.total) + '%</td>' +
           '</tr>';
  }).join('');

  const fixedRows = vm.fixedCosts.map(function (f) {
    return '<tr><td style="padding:6px 0;">' + _esc(f.label) + '</td>' +
           '<td align="right" style="font-style:italic;">' + _yen(f.amount) + '</td></tr>';
  }).join('');

  const fixedBlock = vm.includeFixed ?
    '<div style="border-top:1px solid #1a1410; margin-top:20px; padding-top:16px;">' +
    '<div style="font-size:12px; letter-spacing:0.3em; color:#5a4230; margin-bottom:10px; font-weight:700;">▸ FIXED · ESTIMATED</div>' +
    '<table width="100%" style="font-size:15px;">' + fixedRows + '</table>' +
    '</div>' : '';

  const grandBlock = vm.includeFixed ?
    '<div style="text-align:center; margin-top:20px; padding:16px 0; border-top:5px solid #1a1410; border-bottom:1px solid #1a1410;">' +
    '<div style="font-size:12px; letter-spacing:0.3em; color:#5a4230; font-weight:700;">· GRAND TOTAL ·</div>' +
    '<div style="font-size:34px; font-weight:900; margin-top:4px; letter-spacing:-0.01em;">' + _yen(vm.grandTotal) + '</div>' +
    '</div>' : '';

  const sumLabel = vm.includeFixed ? 'TOTAL' : 'WEEK';
  const leadLine = vm.includeFixed
    ? _monthEn(vm.month) + ' Holds<br>at ' + _yen(vm.total)
    : 'This Week at ' + _yen(vm.total);

  return '<div style="padding:32px 24px; font-family:\'Playfair Display\',Georgia,\'Times New Roman\',serif; color:#1a1410; background:#ebe0c6;">' +
         '<div style="text-align:center; border-top:5px solid #1a1410; padding-top:12px;">' +
         '<div style="font-size:11px; letter-spacing:0.5em; color:#5a4230;">· ESTABLISHED · MMXXIII ·</div>' +
         '</div>' +
         '<div style="text-align:center; border-bottom:1px solid #1a1410; padding-bottom:12px;">' +
         '<div style="font-size:40px; font-weight:900; letter-spacing:-0.01em; margin-top:6px; line-height:0.95;">The Monthly<br>Ledger</div>' +
         '<div style="font-size:11px; letter-spacing:0.35em; color:#5a4230; margin-top:12px;">VOL. III · ' + _monthEn(vm.month).toUpperCase() + ' · ' + vm.year + '</div>' +
         '</div>' +

         '<div style="text-align:center; margin:24px 0 18px;">' +
         '<div style="font-size:12px; letter-spacing:0.3em; color:#5a4230; font-weight:700;">▸ DISPATCH</div>' +
         '<div style="font-size:22px; font-weight:900; margin-top:6px; line-height:1.25;">' + leadLine + '</div>' +
         '</div>' +

         '<table width="100%" style="border-top:1px solid #1a1410; border-bottom:1px solid #1a1410; margin-bottom:20px;"><tr>' +
         '<td align="center" style="width:50%; border-right:1px solid #a88c60; padding:18px 0;">' +
         '<div style="font-size:12px; letter-spacing:0.3em; color:#5a4230; font-weight:700;">' + sumLabel + '</div>' +
         '<div style="font-size:26px; font-weight:900; margin-top:4px;">' + _yen(vm.total) + '</div>' +
         '</td>' +
         '<td align="center" style="width:50%; padding:18px 0;">' +
         '<div style="font-size:12px; letter-spacing:0.3em; color:#5a4230; font-weight:700;">DAILY</div>' +
         '<div style="font-size:26px; font-weight:900; margin-top:4px;">' + _yen(vm.avgDaily) + '</div>' +
         '</td></tr></table>' +

         '<div style="font-size:12px; letter-spacing:0.3em; color:#5a4230; margin-bottom:12px; font-weight:700;">▸ PARTICULARS</div>' +
         '<table width="100%" style="font-size:16px;">' + rows + '</table>' +

         fixedBlock +
         grandBlock +

         '<div style="text-align:center; margin-top:16px; font-size:11px; letter-spacing:0.3em; color:#5a4230; font-style:italic;">— finis —</div>' +

         _copyBox_(plain, { bg: '#fafafa', border: '#1a1410', labelColor: '#5a4230' }) +
         '</div>';
}
