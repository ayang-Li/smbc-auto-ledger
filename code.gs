/**
 * ==============================
 * SMBC 自动记账 + 月度/周度账单系统
 * ==============================
 *
 * 功能：
 * 1. 自动从 Gmail 读取 SMBC 消费邮件
 * 2. 自动解析金额 / 时间 / 地点
 * 3. 自动分类：scene / type / time_tag / 分类
 * 4. 自动写入 Sheet「取引」
 * 5. 定期发送消费报告（周一发周报，1号发月报）
 */

/* ========= 核心配置 ========= */
const SHEET_NAME = '取引';
const TIME_ZONE = Session.getScriptTimeZone();

// 取引表的标准列顺序（数据写入与表头必须严格一致）
// O–R 四列为 2026-06 新增：修正商家 / 一级分类 / 二级分类 / 健康标签
const SHEET_HEADERS = [
  '日付', '時刻', '金額', '利用先', '承認番号',
  'scene', 'type', 'time_tag',
  '分类', '信任度', '规则名', '手动修正', '最终分类', '备注',
  '修正商家', '一级分类', '二级分类', '健康标签'
];

const FIXED_COSTS = [
  // 房租已改由「口座引き落とし事前お知らせ」邮件自动读取，不再手动写入
  { label: '光热水煤(推定)', amount: 15000 },
  { label: '通信/iCloud', amount: 3500 }
];

// 月报用：将细分类合并为大类显示
const CATEGORY_GROUPS = [
  { label: '餐饮(外食/外卖)', cats: ['午餐', '晚餐', '早餐', '夜宵/外卖', '外卖', '晚饭'] },
  { label: '便利店/超市',     cats: ['便利店/杂项', '超市/买菜', '超市买菜', '日常买菜', 'supermarket'] },
  { label: '交通',           cats: ['交通费'] },
  { label: '订阅服务',        cats: ['订阅服务'] },
  { label: '饮料',           cats: ['饮料'] },
  // 房租两类（已于 2026-04-23 统一历史数据，砍掉所有旧 alias）：
  //   '房租(浦安・銀行引落)'    — 浦安 predebit_rent，银行引落直接扣
  //   '房租(エポス代扣・横浜)'  — 横浜 predebit_rent_epos，管理公司指定エポス卡代扣
  { label: '房租',           cats: ['房租(浦安・銀行引落)', '房租(エポス代扣・横浜)'] },
  { label: '水道光热',        cats: ['水道光熱費'] },
  { label: '信用卡还款',      cats: ['信用卡还款'] },
  { label: '银行引落',        cats: ['银行引落'] },
  { label: '日常消费/杂项',   cats: ['日常消费', '大额支出'] },
  { label: '日常购物',        cats: ['日常购物'] },
  { label: '退款',           cats: ['退款/返还'] },
  // 资金调拨不计入消费统计（generateSummary_ 中直接跳过）
];



/* ========= Gmail 导入主任务 (最新在前版) ========= */
function importSmbcDebitMails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  // 自愈：确保表头行含全部标准列名（补齐之前漏写的 O–R 列头）
  ensureSheetHeaders_(sheet);

  const threads = GmailApp.search('from:(smbc-debit@smbc-card.com OR SMBC_service@dn.smbc.co.jp) newer_than:7d', 0, 50);
  Logger.log('找到邮件线程数: ' + threads.length);

  const existingIds = loadExistingIds_(sheet);
  Logger.log('已有记录去重key数: ' + existingIds.size);
  
  const rows = [];

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const body = getFullBody_(msg);

      let pairs = [];
      const single = parseMail_(body);
      if (single) {
        pairs = [{ parsed: single, classified: classify_(single) }];
        Logger.log('[parser] card_notification: ' + single.place);
      } else {
        const bankWd = parseBankWithdrawalMail_(body);
        if (bankWd) {
          pairs = [{ parsed: bankWd, classified: classifyBankWithdrawal_(bankWd) }];
          Logger.log('[parser] bank_withdrawal: ' + bankWd.place);
        } else {
          const preDebitItems = parsePreDebitNoticeMail_(body);
          if (preDebitItems) {
            pairs = preDebitItems.map(p => ({ parsed: p, classified: classifyPreDebitNotice_(p) }));
            Logger.log('[parser] pre_debit_notice: ' + preDebitItems.length + '件');
          } else {
            Logger.log('[parser] WARN: 全パーサー不一致 subject=' + msg.getSubject());
          }
        }
      }

      if (pairs.length === 0) return;

      pairs.forEach(({ parsed, classified }) => {
        const dedupeKey = [
          formatDate_(parsed.date),
          formatTime_(parsed.date),
          parsed.amount,
          Number(parsed.approval) || parsed.approval
        ].join('_');

        if (existingIds.has(dedupeKey)) {
          Logger.log('跳过(已存在): ' + dedupeKey);
          return;
        }

        Logger.log('新增: ' + dedupeKey + ' | ' + parsed.place);
        rows.push([
          formatDate_(parsed.date),
          formatTime_(parsed.date),
          parsed.amount,
          parsed.place,
          parsed.approval,
          classified.scene,
          classified.type,
          classified.timeTag,
          classified.category,
          classified.conf,
          classified.rule,
          '',
          classified.category,
          classified.note || '',
          classified.merchant || '',
          classified.cat1 || '',
          classified.cat2 || '',
          classified.health || ''
        ]);

        existingIds.add(dedupeKey);
      });
    });
  });

  Logger.log('本次新增行数: ' + rows.length);
  
  if (rows.length > 0) {
    // 【坑1解决】新抓到的数据也要内部倒序（最新的在前），保证插入后顺序是对的
    rows.sort((a, b) => new Date(b[0] + ' ' + b[1]) - new Date(a[0] + ' ' + a[1]));
    
    // 【坑2解决】在第 2 行（表头下方）插入空行
    sheet.insertRowsBefore(2, rows.length);
    
    // 【坑3解决】定位到第 2 行开始写入
    const targetRange = sheet.getRange(2, 1, rows.length, rows[0].length);
    
    // 清除可能从表头继承的背景色、粗体等格式
    targetRange.clearFormat(); 
    
    // 写入数据
    targetRange.setValues(rows);
    
    // 重新刷一下金额列的货币格式（第3列）
    sheet.getRange(2, 3, rows.length, 1).setNumberFormat('¥#,##0');

    // --- 隔月字体变灰色 ---
    sheet.clearConditionalFormatRules();
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($A2<>"", TEXT($A2,"yyyy-mm")<>TEXT(TODAY(),"yyyy-mm"))')
      .setFontColor('#999999')
      .setRanges([sheet.getRange("A2:R5000")])
      .build();
    sheet.setConditionalFormatRules([rule]);
    // -----------------------
    
    Logger.log('成功在顶端插入新数据 ✨');
  }
}

/* ========= 邮件解析工具 ========= */
function parseMail_(body) {
  const dateMatch = body.match(/◇利用日[^\n]*?([0-9]{4}\/[0-9]{2}\/[0-9]{2}[ \t]+[0-9]{2}:[0-9]{2}:[0-9]{2})/);
  const amountMatch = body.match(/◇利用金額[^\n]*?(-?[0-9,]+)円/);
  const placeMatch = body.match(/◇利用先[^\n]*?\n?[ \t]*(.+)/);
  const approvalMatch = body.match(/◇承認番号[^\n]*?([0-9]+)/);

  if (!dateMatch || !amountMatch) return null;

  const date = parseDate_(dateMatch[1].trim());
  const amount = Number(amountMatch[1].replace(/,/g, ''));
  const place = placeMatch
    ? placeMatch[1].replace(/[\r\n].*/, '').replace(/^[\s:：　]+/, '').trim()
    : '';

  return {
    date,
    amount,
    place,
    approval: approvalMatch ? approvalMatch[1] : ''
  };
}

/* ========= 分类规则 v3 ========= */

const PLACE_RULES = [
  // 1. 退款
  {
    name: 'refund',
    match: tx => tx.amount < 0,
    out: tx => ({
      scene: '退款', type: '退款', timeTag: classifyTimeTag_(tx.date),
      category: '退款/返还', conf: 'high', rule: 'refund',
      merchant: '', cat1: '退款', cat2: '', health: '', note: ''
    })
  },

  // 2. 订阅
  {
    name: 'subs_apple',
    match: tx => includesAny_(tx.placeNorm, ['APPLE.COM/BILL']),
    out: tx => ({
      scene: '订阅', type: '订阅', timeTag: classifyTimeTag_(tx.date),
      category: '订阅服务', conf: 'high', rule: 'subs_apple',
      merchant: 'Apple', cat1: '订阅服务', cat2: 'App/内容订阅', health: '',
      note: tx.amount === 3400 ? 'Claude会员' : 'Apple订阅'
    })
  },
  {
    name: 'subs_google',
    match: tx => includesAny_(tx.placeNorm, ['GOOGLE', 'YOUTUBE']),
    out: tx => ({
      scene: '订阅', type: '订阅', timeTag: classifyTimeTag_(tx.date),
      category: '订阅服务', conf: 'high', rule: 'subs_google',
      merchant: 'Google/YouTube', cat1: '订阅服务', cat2: 'App/内容订阅', health: '', note: ''
    })
  },
  {
    name: 'subs_direct',
    match: tx => includesAny_(tx.placeNorm, ['NETFLIX', 'SPOTIFY', 'CLAUDE.AI', 'OPENAI']),
    out: tx => ({
      scene: '订阅', type: '订阅', timeTag: classifyTimeTag_(tx.date),
      category: '订阅服务', conf: 'high', rule: 'subs_direct',
      merchant: tx.placeNorm, cat1: '订阅服务', cat2: 'App/内容订阅', health: '', note: ''
    })
  },

  // 3. 资金调拨（不计入日常消费统计）
  {
    name: 'place_wise',
    // 'WISE' 是 4 字符，防止 OTHERWISE/LIKEWISE 等单词误判：要求金额 >= 5000 或 TRANSFERWISE 精确词
    match: tx => includesAny_(tx.placeNorm, ['TRANSFERWISE']) ||
                 (tx.amount >= 5000 && includesAny_(tx.placeNorm, ['WISE'])),
    out: tx => ({
      scene: '资金调拨', type: '国际转账', timeTag: classifyTimeTag_(tx.date),
      category: '资金调拨', conf: 'high', rule: 'place_wise',
      merchant: 'Wise', cat1: '资金调拨', cat2: '国际跨境转账', health: '',
      note: '不计入日常消费（仅手续费单独计入金融服务）'
    })
  },
  {
    name: 'place_paypaybank_card',
    match: tx => includesAny_(tx.placeNorm, ['PAYPAYBANK', 'PAYPAY BANK']),
    out: tx => ({
      scene: '资金调拨', type: '电子钱包充值', timeTag: classifyTimeTag_(tx.date),
      category: '资金调拨', conf: 'high', rule: 'place_paypaybank_card',
      merchant: 'PayPay银行', cat1: '资金调拨', cat2: '电子钱包充值', health: '',
      note: '账户内互转，不计入消费，避免双重记账'
    })
  },

  // 4. 外卖
  {
    name: 'place_uber_eats',
    match: tx => includesAny_(tx.placeNorm, ['UBER * EATS', 'UBER EATS']),
    out: tx => {
      const timeTag = classifyTimeTag_(tx.date);
      return {
        scene: '外卖', type: '正餐', timeTag,
        category: timeTagToMealCategory_(timeTag, '外卖'),
        conf: 'high', rule: 'place_uber_eats',
        merchant: 'Uber Eats', cat1: '餐饮外卖', cat2: '线上外卖平台',
        health: '🔄 视底层商家而定', note: ''
      };
    }
  },
  {
    name: 'place_uber_pending',
    match: tx => includesAny_(tx.placeNorm, ['UBER * PENDING', 'UBER PENDING']),
    out: tx => {
      const timeTag = classifyTimeTag_(tx.date);
      if (tx.amount <= 200) {
        return {
          scene: '未知', type: '小额消费', timeTag,
          category: '便利店/杂项', conf: 'low', rule: 'uber_pending_small',
          merchant: 'Uber', cat1: '待分类', cat2: '小额/未知', health: '', note: ''
        };
      }
      return {
        scene: '外卖', type: '正餐', timeTag,
        category: timeTagToMealCategory_(timeTag, '外卖'),
        conf: 'mid', rule: 'uber_pending_meal',
        merchant: 'Uber Eats', cat1: '餐饮外卖', cat2: '线上外卖平台',
        health: '🔄 视底层商家而定', note: ''
      };
    }
  },

  // 5. 便利店
  {
    name: 'place_cvs',
    match: tx => includesAny_(tx.placeNorm, ['FAMILYMART', 'SEVEN-ELEVEN', 'LAWSON', 'MINISTOP']),
    out: tx => {
      const convType = classifyConvenienceType_(tx.amount);
      return {
        scene: '便利店', type: convType,
        timeTag: classifyTimeTag_(tx.date),
        category: '便利店/杂项', conf: 'high', rule: 'place_cvs',
        merchant: cvsMerchantName_(tx.placeNorm),
        cat1: '便利店', cat2: convType, health: '', note: ''
      };
    }
  },

  // 6. 超市
  {
    name: 'place_supermarket_celsior',
    match: tx => includesAny_(tx.placeNorm, ['CELSIOR']),
    out: tx => ({
      scene: '超市', type: classifySupermarketType_(tx.amount),
      timeTag: classifyTimeTag_(tx.date),
      category: '超市/买菜', conf: 'high', rule: 'place_supermarket_celsior',
      merchant: 'CELSIOR', cat1: '日常购物', cat2: '生鲜超市/买菜', health: '',
      note: 'CELSIOR=超市'
    })
  },
  {
    name: 'place_aeon',
    match: tx => includesAny_(tx.placeNorm, ['AEON', 'MAIBASUKE', 'MY BASKET', 'MYBASKET']),
    out: tx => ({
      scene: '超市', type: classifySupermarketType_(tx.amount),
      timeTag: classifyTimeTag_(tx.date),
      category: '超市/买菜', conf: 'high', rule: 'place_aeon',
      merchant: '永旺/My Basket (まいばすけっと)',
      cat1: '日常购物', cat2: '生鲜超市/买菜', health: '',
      note: '自炊食材/日常补给'
    })
  },

  // 7. 餐厅
  {
    name: 'place_katsuya',
    match: tx => includesAny_(tx.placeNorm, ['KATSUYA']),
    out: tx => {
      const timeTag = classifyTimeTag_(tx.date);
      return {
        scene: '餐厅', type: '正餐', timeTag,
        category: timeTagToMealCategory_(timeTag, '餐厅'),
        conf: 'high', rule: 'place_katsuya',
        merchant: '吉豚屋 (かつや)', cat1: '餐饮正餐', cat2: '炸物/日式猪排',
        health: '🔴 不太健康', note: '高油炸高卡'
      };
    }
  },
  {
    name: 'place_mcdonalds',
    match: tx => includesAny_(tx.placeNorm, ['MCDONALD']),
    out: tx => {
      const timeTag = classifyTimeTag_(tx.date);
      const isLight = tx.amount <= 500;
      return {
        scene: '餐厅', type: isLight ? '轻食' : '正餐', timeTag,
        category: timeTagToMealCategory_(timeTag, '餐厅'),
        conf: 'high', rule: 'place_mcdonalds',
        merchant: "麦当劳 (McDonald's)",
        cat1: isLight ? '餐饮轻食' : '餐饮正餐', cat2: '西式快餐',
        health: '🔴 不太健康', note: '油炸/精制碳水'
      };
    }
  },
  {
    name: 'place_yakinikukingu',
    match: tx => includesAny_(tx.placeNorm, ['YAKINIKUKINGU', 'YAKINIKU KINGU', 'YAKINIKUKING']),
    out: tx => {
      const timeTag = classifyTimeTag_(tx.date);
      return {
        scene: '餐厅', type: '正餐', timeTag,
        category: timeTagToMealCategory_(timeTag, '餐厅'),
        conf: 'high', rule: 'place_yakinikukingu',
        merchant: '烧肉王 (焼肉きんぐ)', cat1: '餐饮正餐', cat2: '日式烧肉放题',
        health: '🟡 注意放纵', note: '高蛋白/高热量'
      };
    }
  },
  {
    name: 'place_sukiya',
    match: tx => includesAny_(tx.placeNorm, ['SUKIYA']),
    out: tx => {
      const timeTag = classifyTimeTag_(tx.date);
      return {
        scene: '餐厅', type: '正餐', timeTag,
        category: timeTagToMealCategory_(timeTag, '餐厅'),
        conf: 'high', rule: 'place_sukiya',
        merchant: '食其家 (すき家)', cat1: '餐饮正餐', cat2: '日式牛丼快餐',
        health: '🟡 普通快餐', note: '碳水偏高'
      };
    }
  },
  {
    name: 'place_nakau',
    match: tx => includesAny_(tx.placeNorm, ['NAKAU']),
    out: tx => {
      const timeTag = classifyTimeTag_(tx.date);
      return {
        scene: '餐厅', type: '正餐', timeTag,
        category: timeTagToMealCategory_(timeTag, '餐厅'),
        conf: 'high', rule: 'place_nakau',
        merchant: '奈卡乌 (なか卯)', cat1: '餐饮正餐', cat2: '日式快餐/乌冬',
        health: '🟡 普通快餐', note: '碳水偏高'
      };
    }
  },
  {
    name: 'place_kitchen_origin',
    match: tx => includesAny_(tx.placeNorm, ['KITCHENORIGIN', 'KITCHEN ORIGIN', 'KITCHEN-ORIGIN']),
    out: tx => {
      const timeTag = classifyTimeTag_(tx.date);
      return {
        scene: '餐厅', type: '轻食', timeTag,
        category: timeTagToMealCategory_(timeTag, '餐厅'),
        conf: 'high', rule: 'place_kitchen_origin',
        merchant: 'Kitchen Origin (和田町店)', cat1: '餐饮轻食', cat2: '便当/外带熟食',
        health: '🟢 普通正餐', note: '低客单价，多为单点小菜/饭团'
      };
    }
  },
  {
    name: 'place_syanhaitei',
    match: tx => includesAny_(tx.placeNorm, ['SYANHAITEI', 'SHANHAI', 'SHANGHAITEI']),
    out: tx => {
      const timeTag = classifyTimeTag_(tx.date);
      return {
        scene: '餐厅', type: '正餐', timeTag,
        category: timeTagToMealCategory_(timeTag, '餐厅'),
        conf: 'high', rule: 'place_syanhaitei',
        merchant: '上海亭 (木场店)', cat1: '餐饮正餐', cat2: '中华料理',
        health: '🟡 普通正餐', note: '油脂可能偏高'
      };
    }
  },
  {
    name: 'place_dinii_banrao',
    match: tx => includesAny_(tx.placeNorm, ['DINII', 'BANRAO']),
    out: tx => {
      const timeTag = classifyTimeTag_(tx.date);
      return {
        scene: '餐厅', type: '正餐', timeTag,
        category: timeTagToMealCategory_(timeTag, '餐厅'),
        conf: 'high', rule: 'place_dinii_banrao',
        merchant: '泰料店 (Banrao)', cat1: '餐饮正餐', cat2: '东南亚料理',
        health: '🟢 普通正餐', note: 'Dinii为点餐系统，banrao为实际商家'
      };
    }
  },

  // 8. 零售 / 药妆 / 综合
  {
    name: 'place_katsumata',
    match: tx => includesAny_(tx.placeNorm, ['KATSUMATA']),
    out: tx => ({
      scene: '购物', type: '日化/药妆', timeTag: classifyTimeTag_(tx.date),
      category: '日常购物', conf: 'high', rule: 'place_katsumata',
      merchant: 'Katsumata 药妆店 (和田町店)',
      cat1: '日常购物', cat2: '医药/日化药妆', health: '',
      note: '个人护理/生活用品'
    })
  },
  {
    name: 'place_bigbox',
    match: tx => includesAny_(tx.placeNorm, ['BIGBOX', 'BIG BOX']),
    out: tx => ({
      scene: '购物', type: '综合零售', timeTag: classifyTimeTag_(tx.date),
      category: '日常消费', conf: 'mid', rule: 'place_bigbox',
      merchant: '高田马场 BIGBOX',
      cat1: '综合娱乐', cat2: '商业综合体/零售', health: '',
      note: '复合型消费'
    })
  },

  // 9. 交通
  {
    name: 'place_suica',
    match: tx => includesAny_(tx.placeNorm, ['MOBILE SUICA']),
    out: tx => {
      const isGreenCar = tx.amount === 750;
      const isLarge = tx.amount >= 500;
      return {
        scene: '交通', type: isLarge ? '车费/特急' : '通勤',
        timeTag: classifyTimeTag_(tx.date),
        category: '交通费', conf: 'high', rule: 'place_suica',
        merchant: 'Mobile Suica', cat1: '交通出行',
        cat2: isGreenCar ? '普通列车绿色车厢票' : (isLarge ? '特急/长途' : '通勤交通'),
        health: '',
        note: isGreenCar ? '¥750 Green Car グリーン車' : (isLarge ? '可能是グリーン車/特急' : '')
      };
    }
  },
  {
    name: 'place_icoca',
    match: tx => includesAny_(tx.placeNorm, ['MOBILE ICOCA']),
    out: tx => ({
      scene: '交通', type: '通勤', timeTag: classifyTimeTag_(tx.date),
      category: '交通费', conf: 'high', rule: 'place_icoca',
      merchant: 'Mobile ICOCA', cat1: '交通出行', cat2: '通勤交通', health: '', note: ''
    })
  },

  // 10. 支付通道（系统无法穿透底层商家，需手动补充「手动修正」列）
  {
    name: 'place_payment_gateway',
    match: tx => includesAny_(tx.placeNorm, ['IDデビット', 'SQ*SQUARE', 'SQ*']),
    out: tx => ({
      scene: '未知', type: '支付通道', timeTag: classifyTimeTag_(tx.date),
      category: '日常消费', conf: 'low', rule: 'place_payment_gateway',
      merchant: '', cat1: '待分类', cat2: '技术网关通道', health: '',
      note: '⚠️ 请手动补充底层商家到「手动修正」列'
    })
  }
];

/* ========= 主分类函数 ========= */
function classify_(tx) {
  const normalized = normalizeTx_(tx);

  for (const rule of PLACE_RULES) {
    if (rule.match(normalized)) {
      return rule.out(normalized);
    }
  }

  return classifyFallback_(normalized);
}

/* ========= fallback ========= */
function classifyFallback_(tx) {
  const timeTag = classifyTimeTag_(tx.date);
  const amount = tx.amount;

  if (amount >= 8000) {
    return {
      scene: '未知', type: '大额消费', timeTag,
      category: '大额支出', conf: 'low', rule: 'big_amount',
      merchant: '', cat1: '大额支出', cat2: '', health: '', note: ''
    };
  }

  if (tx.hour >= 11 && tx.hour < 14 && amount >= 600 && amount <= 3000) {
    return {
      scene: '未知', type: '正餐', timeTag,
      category: '午餐', conf: 'mid', rule: 'time_lunch',
      merchant: '', cat1: '餐饮正餐', cat2: '未知', health: '', note: ''
    };
  }

  if (tx.hour >= 17 && tx.hour < 21 && amount >= 800) {
    return {
      scene: '未知', type: '正餐', timeTag,
      category: '晚餐', conf: 'mid', rule: 'time_dinner',
      merchant: '', cat1: '餐饮正餐', cat2: '未知', health: '', note: ''
    };
  }

  if (tx.hour >= 22 || tx.hour < 4) {
    return {
      scene: '未知', type: amount <= 300 ? '饮料/小额' : '正餐', timeTag,
      category: amount <= 300 ? '便利店/杂项' : '夜宵/外卖',
      conf: 'mid', rule: 'time_late',
      merchant: '', cat1: amount <= 300 ? '便利店' : '餐饮外卖', cat2: '未知', health: '', note: ''
    };
  }

  if (amount > 100 && amount <= 250) {
    return {
      scene: '未知', type: '饮料/小额', timeTag,
      category: '饮料', conf: 'mid', rule: 'price_drink',
      merchant: '', cat1: '饮料', cat2: '', health: '', note: ''
    };
  }

  return {
    scene: '未知', type: amount <= 300 ? '小额消费' : '日常消费', timeTag,
    category: '日常消费', conf: 'low', rule: 'default',
    merchant: '', cat1: '日常消费', cat2: '', health: '', note: ''
  };
}

/* ========= 归一化工具 ========= */
function normalizeTx_(tx) {
  return {
    ...tx,
    placeNorm: normalizePlace_(tx.place || ''),
    hour: tx.date.getHours()
  };
}

function normalizePlace_(place) {
  return String(place || '')
    .toUpperCase()
    .replace(/[：:]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function includesAny_(text, keywords) {
  return keywords.some(k => text.includes(k));
}

function cvsMerchantName_(placeNorm) {
  if (placeNorm.includes('FAMILYMART')) return 'FamilyMart';
  if (placeNorm.includes('SEVEN-ELEVEN')) return 'Seven-Eleven';
  if (placeNorm.includes('LAWSON')) return 'Lawson';
  if (placeNorm.includes('MINISTOP')) return 'Ministop';
  return '便利店';
}

/* ========= 标签工具 ========= */
function classifyTimeTag_(date) {
  const hour = date.getHours();

  if (hour >= 5 && hour < 10) return '早餐';
  if (hour >= 10 && hour < 15) return '午餐';
  if (hour >= 15 && hour < 18) return '下午';
  if (hour >= 18 && hour < 22) return '晚餐';
  return '夜间';
}

function classifyConvenienceType_(amount) {
  if (amount <= 220) return '饮料/小食';
  if (amount <= 700) return '轻食';
  return '杂项';
}

function classifySupermarketType_(amount) {
  if (amount <= 500) return '补货';
  if (amount <= 2000) return '日常买菜';
  return '集中采购';
}

function timeTagToMealCategory_(timeTag, scene) {
  if (scene === '外卖') {
    if (timeTag === '午餐') return '午餐';
    if (timeTag === '晚餐') return '晚餐';
    if (timeTag === '夜间') return '夜宵/外卖';
    return '外卖';
  }

  if (scene === '餐厅') {
    if (timeTag === '午餐') return '午餐';
    if (timeTag === '晚餐') return '晚餐';
    if (timeTag === '夜间') return '夜宵/外卖';
    return '日常消费';
  }

  return '日常消费';
}

/* ========= 报告系统 ========= */

// 发送周报
function sendWeeklyExpenseReport() {
  const now = new Date();
  const oneWeekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  const summary = generateSummary_(oneWeekAgo, now);

  const title = `[周报] 上周消费周报 (${formatDate_(oneWeekAgo)} ~ ${formatDate_(now)})`;
  sendEmailReport_(title, summary);
}

// 发送本月截至今日的报告（测试 / 月中预览用）
function sendCurrentMonthExpenseReport() {
  const now = new Date();
  const firstThis = new Date(now.getFullYear(), now.getMonth(), 1);
  const summary = generateSummary_(firstThis, now);

  const ym = Utilities.formatDate(firstThis, 'Asia/Tokyo', 'yyyy-MM');
  const title = `[月报-预览] ${ym} 本月消费截至 ${formatDate_(now)}`;
  sendEmailReport_(title, summary, true);
}

// 发送月报
function sendLastMonthExpenseReport() {
  const now = new Date();
  const firstThis = new Date(now.getFullYear(), now.getMonth(), 1);
  const firstLast = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const summary = generateSummary_(firstLast, firstThis);

  const ym = Utilities.formatDate(firstLast, 'Asia/Tokyo', 'yyyy-MM');
  const title = `[月报] 月度消费报告 ${ym}`;
  sendEmailReport_(title, summary, true);
}

// 汇总逻辑
function generateSummary_(start, end) {
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME)
    .getDataRange()
    .getValues();

  const headers = data[0];
  const dateCol = headers.indexOf('日付');
  const amtCol = headers.indexOf('金額');
  const finalCatCol = headers.indexOf('最终分类');
  const catCol = headers.indexOf('分类');
  const oldCatCol = headers.indexOf('分類'); // 旧列名（迁移前数据在此列）
  const ruleCol = headers.indexOf('规则名');

  // 防御: 必须列缺失时提前报错
  if (dateCol < 0 || amtCol < 0) {
    Logger.log('[generateSummary_] WARN: 列缺失 dateCol=' + dateCol + ' amtCol=' + amtCol);
  }

  let total = 0;
  const byCat = {};
  let skipNoCat = 0, skipBadDate = 0, skipOutOfRange = 0, fallbackOldCol = 0;

  for (let i = 1; i < data.length; i++) {
    const d = data[i][dateCol];
    const amt = data[i][amtCol];
    // 优先读 最终分类(M) → 分类(I) → 旧分類(O)，兼容迁移前老数据
    const finalCatVal = finalCatCol >= 0 ? data[i][finalCatCol] : '';
    const catVal      = catCol >= 0      ? data[i][catCol]      : '';
    const oldCatVal   = oldCatCol >= 0   ? data[i][oldCatCol]   : '';
    const usedOldCol  = !finalCatVal && !catVal && !!oldCatVal;
    const cat = (finalCatVal || catVal || oldCatVal || '').toString().trim();
    if (usedOldCol) fallbackOldCol++;

    if (!cat) { skipNoCat++; continue; }

    const dObj = (d instanceof Date) ? d : new Date(String(d) + 'T00:00:00');
    if (isNaN(dObj.getTime())) { skipBadDate++; continue; }
    if (dObj < start || dObj >= end) { skipOutOfRange++; continue; }

    // 资金调拨（Wise/PayPay充值等）不计入消费统计，避免双重记账
    if (cat === '资金调拨') continue;

    total += amt;
    byCat[cat] = (byCat[cat] || 0) + amt;
  }

  Logger.log(
    '[generateSummary_] 期間=' + Utilities.formatDate(start, Session.getScriptTimeZone(), 'MM/dd') +
    '~' + Utilities.formatDate(end, Session.getScriptTimeZone(), 'MM/dd') +
    ' | 計入=' + Object.values(byCat).reduce((a,b) => a+b, 0) + '円' +
    ' | skip(no_cat=' + skipNoCat + ' bad_date=' + skipBadDate + ' out_of_range=' + skipOutOfRange + ')' +
    ' | oldCol_fallback=' + fallbackOldCol
  );

  return { total, byCat, start, end };
}

// 邮件发送：路由到 emailTemplates.gs 里的 HTML 模板，按月份自动轮换主题。
// 每封邮件同时带 HTML body（好看）和 plain body（fallback，也嵌入在 HTML 底部作长按复制区）。
function sendEmailReport_(title, summary, includeFixed = false) {
  const vm    = buildViewModel_(summary, includeFixed);
  const tplId = pickTemplateId_(new Date());
  const html  = buildEmailHtml_(vm, title, tplId);
  const plain = buildPlainText_(vm, title);

  GmailApp.sendEmail(
    Session.getActiveUser().getEmail(),
    title,
    plain,
    { htmlBody: html, name: 'SMBC 记账助手' }
  );

  Logger.log('[sendEmailReport_] 已发送 · template=' + tplId + ' · includeFixed=' + includeFixed);
}

/* ========= 触发器设置 ========= */
function setupTriggers() {
  const funcs = [
    'importSmbcDebitMails',
    'sendLastMonthExpenseReport',
    'sendWeeklyExpenseReport'
  ];

  ScriptApp.getProjectTriggers().forEach(t => {
    if (funcs.includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('importSmbcDebitMails')
    .timeBased()
    .everyHours(1)
    .create();

  ScriptApp.newTrigger('sendWeeklyExpenseReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  ScriptApp.newTrigger('sendLastMonthExpenseReport')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();
    
}

/* ========= 通用工具函数 ========= */

function loadExistingIds_(sheet) {
  const set = new Set();
  const last = sheet.getLastRow();
  if (last < 2) return set;

  const values = sheet.getRange(2, 1, last - 1, 5).getValues();
  values.forEach(r => {
    const timeVal = r[1] instanceof Date
      ? Utilities.formatDate(r[1], TIME_ZONE, 'HH:mm:ss')
      : String(r[1]);
    const key = [
      (r[0] instanceof Date ? Utilities.formatDate(r[0], TIME_ZONE, 'yyyy-MM-dd') : String(r[0])),
      timeVal,
      r[2],
      Number(r[4]) || r[4]
    ].join('_');
    set.add(key);
  });

  return set;
}

function parseDate_(s) {
  const parts = s.split(/\s+/);
  const d = parts[0].split('/').map(Number);
  const t = parts[1].split(':').map(Number);
  return new Date(d[0], d[1] - 1, d[2], t[0], t[1], t[2]);
}

function formatDate_(d) {
  return Utilities.formatDate(d, TIME_ZONE, 'yyyy-MM-dd');
}

function formatTime_(d) {
  return Utilities.formatDate(d, TIME_ZONE, 'HH:mm:ss');
}

function formatDateTime_(d) {
  return Utilities.formatDate(d, TIME_ZONE, 'yyyyMMdd_HHmmss');
}

function restructureTransactionSheetPreserveMetadata() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`找不到工作表: ${SHEET_NAME}`);

  const HEADER_ROW = 1;

  // 目标列顺序（你真正要用的）—— 与数据写入共用同一份定义
  const desiredHeaders = SHEET_HEADERS;

  // 不想要但先保留元数据的列：移到右边并隐藏，不直接删除
  const archiveHeaders = ['id', '利用日', '分類', '信頼度', 'ルール名'];

  // 读取当前表头
  let headers = getHeaderRow_(sh, HEADER_ROW);

  // 1) 先补齐缺失列（加在最右边，后面再移动到正确位置）
  const missingDesired = desiredHeaders.filter(h => !headers.includes(h));
  if (missingDesired.length > 0) {
    sh.insertColumnsAfter(sh.getLastColumn(), missingDesired.length);
    sh.getRange(HEADER_ROW, sh.getLastColumn() - missingDesired.length + 1, 1, missingDesired.length)
      .setValues([missingDesired]);
    headers = getHeaderRow_(sh, HEADER_ROW);
  }

  // 2) 把目标列按 desiredHeaders 顺序移动到前面
  //    用整列 move，尽量保留格式/备注/验证等元数据
  desiredHeaders.forEach((headerName, targetIdxZeroBased) => {
    moveHeaderColumnToIndex_(sh, HEADER_ROW, headerName, targetIdxZeroBased + 1);
  });

  // 3) 把 archiveHeaders（id / 利用日）移动到最右侧，但不删除
  archiveHeaders.forEach(headerName => {
    if (getHeaderIndex_(sh, HEADER_ROW, headerName) > 0) {
      moveHeaderColumnToIndex_(sh, HEADER_ROW, headerName, sh.getLastColumn());
    }
  });

  // 4) 隐藏 archiveHeaders，对外看起来就像“删了”，但元数据还在
  headers = getHeaderRow_(sh, HEADER_ROW);
  archiveHeaders.forEach(headerName => {
    const col = headers.indexOf(headerName) + 1;
    if (col > 0) sh.hideColumns(col);
  });

  SpreadsheetApp.flush();
  Logger.log('Sheet 列结构已重构完成（保留元数据版）');
}

/**
 * 把指定表头所在列移动到 targetCol（1-based）
 * 通过整列移动，尽量保留该列所有元数据
 */
function moveHeaderColumnToIndex_(sheet, headerRow, headerName, targetCol) {
  const currentCol = getHeaderIndex_(sheet, headerRow, headerName);
  if (currentCol <= 0) return;
  if (currentCol === targetCol) return;

  const maxRows = sheet.getMaxRows();
  const colRange = sheet.getRange(1, currentCol, maxRows, 1);

  // Apps Script 的 moveColumns 是“插到 destinationIndex 前面”
  // 如果当前列在 targetCol 左边，移动后目标位会左缩 1，所以需要 +1 修正
  const destination = currentCol < targetCol ? targetCol + 1 : targetCol;

  sheet.moveColumns(colRange, destination);
}

/**
 * 确保表头行（第1行）与标准列名 SHEET_HEADERS 完全一致。
 * 数据写入按固定列序占满 A–R 全 18 列，所以这 18 个表头必须严格对应；
 * 任何不一致都是残留旧列名（如 分類/信頼度）或漏填，统一校正。
 * 仅校正前 18 列，第 19 列及之后（隐藏的存档列）保持不动。
 */
function ensureSheetHeaders_(sheet) {
  if (!sheet) return;
  const lastCol = Math.max(sheet.getLastColumn(), SHEET_HEADERS.length);
  const current = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const changedCols = [];  // 被校正的列号（1-based）
  for (let i = 0; i < SHEET_HEADERS.length; i++) {
    if (String(current[i]).trim() !== SHEET_HEADERS[i]) {
      Logger.log(`表头校正 第${i + 1}列: "${current[i]}" → "${SHEET_HEADERS[i]}"`);
      current[i] = SHEET_HEADERS[i];
      changedCols.push(i + 1);
    }
  }
  if (changedCols.length === 0) return;

  sheet.getRange(1, 1, 1, current.length).setValues([current]);

  // 让新校正的表头格继承既有表头样式（灰底白字加粗）：
  // 取一个“未被改动”的表头列当模板；整行都是新建时无模板可继承，跳过。
  let templateCol = 0;
  for (let i = 0; i < SHEET_HEADERS.length; i++) {
    if (changedCols.indexOf(i + 1) === -1) { templateCol = i + 1; break; }
  }
  if (templateCol > 0) {
    const template = sheet.getRange(1, templateCol);
    changedCols.forEach(col => template.copyTo(sheet.getRange(1, col), { formatOnly: true }));
  }
  Logger.log('表头自愈：已校正列名' + (templateCol > 0 ? '并继承表头样式' : ''));
}

/**
 * 获取 header 行（裁掉右侧空白）
 */
function getHeaderRow_(sheet, headerRow) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
}

/**
 * 根据表头名获取列号（1-based），找不到返回 -1
 */
function getHeaderIndex_(sheet, headerRow, headerName) {
  const headers = getHeaderRow_(sheet, headerRow);
  const idx = headers.indexOf(headerName);
  return idx >= 0 ? idx + 1 : -1;
}

function parseBankWithdrawalMail_(body) {
  // 只处理“口座引落し / 出金”这类邮件
  if (!/口座引落し/.test(body) || !/出金額/.test(body)) return null;

  const dateMatch = body.match(/出金日\s*：\s*([0-9]{4})年([0-9]{2})月([0-9]{2})日/);
  const amountMatch = body.match(/出金額\s*：\s*([0-9,]+)円/);
  const contentMatch = body.match(/内容\s*：\s*(.+)/);
  const sentAtMatch = body.match(/（([0-9]{4})年([0-9]{2})月([0-9]{2})日([0-9]{2})時([0-9]{2})分現在/);

  if (!dateMatch || !amountMatch) return null;

  const y = Number(dateMatch[1]);
  const m = Number(dateMatch[2]) - 1;
  const d = Number(dateMatch[3]);

  // 邮件里没有交易秒级时间，就优先取“現在”里的时分；没有就给 12:00
  let hh = 12;
  let mm = 0;
  if (sentAtMatch) {
    hh = Number(sentAtMatch[4]);
    mm = Number(sentAtMatch[5]);
  }

  const date = new Date(y, m, d, hh, mm, 0);
  const amount = Number(amountMatch[1].replace(/,/g, ''));
  const place = contentMatch ? String(contentMatch[1]).trim() : '銀行引落';

  return {
    date,
    amount,
    place,
    approval: 'BANK_' + Utilities.formatDate(date, TIME_ZONE, 'yyyyMMdd_HHmm') + '_' + place
  };
}

/**
 * 解析「口座引き落とし事前お知らせ」邮件
 * 一封邮件可能包含多条 ◆明細，返回数组；解析失败返回 null
 */
function parsePreDebitNoticeMail_(body) {
  if (!/口座引落予定日/.test(body)) return null;

  const scheduledDateMatch = body.match(/口座引落予定日[：:　 \t]*([0-9]{4})年([0-9]{2})月([0-9]{2})日/);
  if (!scheduledDateMatch) return null;

  const y  = Number(scheduledDateMatch[1]);
  const mo = Number(scheduledDateMatch[2]) - 1;
  const d  = Number(scheduledDateMatch[3]);

  // 用邮件里「現在」时间戳精确去重（时分）
  const sentAtMatch = body.match(/([0-9]{4})年([0-9]{2})月([0-9]{2})日([0-9]{2})時([0-9]{2})分現在/);
  const hh = sentAtMatch ? Number(sentAtMatch[4]) : 9;
  const mm = sentAtMatch ? Number(sentAtMatch[5]) : 0;

  const deliveryNumMatch = body.match(/配信番号[：:　 \t]*([0-9\-]+)/);
  const deliveryNum = deliveryNumMatch ? deliveryNumMatch[1] : 'UNK';

  // 按 ◆明細N 切块，每块独立解析
  const items = [];
  const chunks = body.split(/◆明細[0-9０-９]+/);
  chunks.shift(); // 丢掉第一块（明細之前的内容）

  chunks.forEach((chunk, idx) => {
    const amtMatch = chunk.match(/引落金額[：:][　 \t]*([0-9,]+)円/);
    const cntMatch = chunk.match(/内容[　 \t]*[：:][　 \t]*(.+)/);
    if (!amtMatch || !cntMatch) return;

    const amount = Number(amtMatch[1].replace(/,/g, ''));
    const place  = String(cntMatch[1]).trim();
    const date   = new Date(y, mo, d, hh, mm, 0);
    items.push({
      date,
      amount,
      place,
      approval: `PREDEBIT_${y}${String(mo + 1).padStart(2, '0')}${String(d).padStart(2, '0')}_${deliveryNum}_${idx + 1}`
    });
  });

  return items.length > 0 ? items : null;
}

function classifyPreDebitNotice_(tx) {
  const norm = String(tx.place || '')
    .toUpperCase()
    .replace(/[（）()]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  // 浦安住宅，银行引落直接扣（SMBC(ヤチン(セーフテイ…）
  if (norm.includes('ヤチン') || norm.includes('家賃')) {
    return { scene: '固定支出', type: '房租', timeTag: '月次', category: '房租(浦安・銀行引落)', conf: 'high', rule: 'predebit_rent', merchant: '浦安住宅', cat1: '固定支出', cat2: '房租/银行引落', health: '', note: tx.place };
  }
  // 横浜住宅，管理公司指定用エポス卡代扣房租（非普通信用卡消费）
  if (norm.includes('エポス')) {
    return { scene: '固定支出', type: '房租', timeTag: '月次', category: '房租(エポス代扣・横浜)', conf: 'high', rule: 'predebit_rent_epos', merchant: 'エポス卡代扣', cat1: '固定支出', cat2: '房租/信用卡代扣', health: '', note: tx.place };
  }
  if (norm.includes('ラクテン') || norm.includes('RAKUTEN') ||
      (norm.includes('カード') && norm.includes('サービス'))) {
    return { scene: '固定支出', type: '信用卡还款', timeTag: '月次', category: '信用卡还款', conf: 'high', rule: 'predebit_card', merchant: '楽天カード', cat1: '固定支出', cat2: '信用卡还款', health: '', note: tx.place };
  }
  if (norm.includes('水道')) {
    return { scene: '固定支出', type: '水道費', timeTag: '月次', category: '水道光熱費', conf: 'high', rule: 'predebit_water', merchant: '水道局', cat1: '固定支出', cat2: '水道光热', health: '', note: tx.place };
  }
  return { scene: '固定支出', type: '银行引落', timeTag: '月次', category: '银行引落', conf: 'mid', rule: 'predebit_default', merchant: '', cat1: '固定支出', cat2: '银行引落/其他', health: '', note: tx.place };
}

function classifyBankWithdrawal_(tx) {
  const placeNorm = String(tx.place || '').toUpperCase();

  if (placeNorm.includes('PAYPAY')) {
    return {
      scene: '资金调拨', type: '电子钱包充值', timeTag: classifyTimeTag_(tx.date),
      category: '资金调拨', conf: 'high', rule: 'bank_paypay',
      merchant: 'PayPay', cat1: '资金调拨', cat2: '电子钱包充值', health: '',
      note: '银行账户出金到PAYPAY，不计入消费，避免双重记账'
    };
  }

  return {
    scene: '资金出金', type: '账户出金', timeTag: classifyTimeTag_(tx.date),
    category: '日常消费', conf: 'mid', rule: 'bank_withdrawal',
    merchant: '', cat1: '日常消费', cat2: '银行出金/其他', health: '',
    note: '银行账户出金'
  };
}

function getFullBody_(msg) {
  let body = msg.getPlainBody();
  if (body.length < 100) {
    body = msg.getBody().replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ');
  }
  return body;
}

/**
 * 【终极加固版】日次 Sheet 设置
 * 目标：倒序排列、跨月置灰、合计项精准锁定 6-8 行、备注随动
 */
function setupDailySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('日次');
  if (!sh) { sh = ss.insertSheet('日次'); } else { sh.clear(); sh.clearConditionalFormatRules(); }

  // --- A/B 列：核心消费数据 ---
  // 1. 表头
  sh.getRange('A1:B1').setValues([['日付', '当日合计(円)']]).setFontWeight('bold').setBackground('#f3f3f3');

  // 2. A2 倒序日期公式 (保持日期对象，不转文本)
  sh.getRange('A2').setFormula('=IFERROR(SORT(UNIQUE(FILTER(INDIRECT("\'取引\'!A2:A"), INDIRECT("\'取引\'!A2:A")<>"")), 1, FALSE), "")');
  
  // 【修复 1】强行锁定 A 列格式，防止变成 46107
  sh.getRange('A2:A1000').setNumberFormat('yyyy-mm-dd');

  // 3. B2 每日求和公式
  // 【修复 2】删掉了 TEXT() 函数，直接用 A2:A 进行数值对比，确保 SUMIF 匹配成功
  sh.getRange('B2').setFormula('=ARRAYFORMULA(IF(A2:A="","",SUMIF(取引!A:A, A2:A, 取引!C:C)))');
  sh.getRange('B2:B1000').setNumberFormat('¥#,##0');

  // 4. 置灰规则
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"", TEXT($A2,"yyyy-mm")<>TEXT(TODAY(),"yyyy-mm"))')
    .setFontColor('#999999')
    .setRanges([sh.getRange('A2:B1000')])
    .build();
  sh.setConditionalFormatRules([rule]);


  // --- D/E 列：固定支出与汇总 (精准对齐 6, 7, 8) ---
  sh.getRange('D1').setValue('── 固定支出(估算参考) ──').setFontWeight('bold');
  
  FIXED_COSTS.forEach((f, i) => {
    sh.getRange(i + 2, 4).setValue(f.label);
    sh.getRange(i + 2, 5).setValue(f.amount);
  });

  // 第 6 行：固定估算小计
  sh.getRange(6, 4).setValue('固定估算小计');
  sh.getRange(6, 5).setFormula('=SUM(E2:E3)');
  sh.getRange(6, 4).setNote('光热水煤 + 通信/iCloud 的手动估算值，不计入账本。');

  // 第 7 行：当月取引实际合计
  sh.getRange(7, 4).setValue('当月取引实际合计');
  sh.getRange(7, 5).setFormula('=SUMIFS(取引!C:C,取引!A:A,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),取引!A:A,"<"&DATE(YEAR(TODAY()),MONTH(TODAY())+1,1))');
  sh.getRange(7, 4).setNote('账本当月全部实际交易：刷卡消费 + 银行出金 + 房租等。');

  // 第 8 行：月花销总计
  sh.getRange(8, 4).setValue('月花销当前总计(估算固定+实际)');
  sh.getRange(8, 5).setFormula('=E6+E7'); 
  sh.getRange(8, 4).setNote('固定估算小计 + 当月取引实际合计。');


  // --- 格式美化 ---
  sh.getRange('E2:E8').setNumberFormat('¥#,##0');
  sh.getRange('D6:E8').setFontWeight('bold');
  sh.setColumnWidth(1, 110);
  sh.setColumnWidth(2, 120);
  sh.setColumnWidth(4, 220); 
  sh.setColumnWidth(5, 110);

  SpreadsheetApp.flush();
  Logger.log('日次 Sheet Bug 已修复，格式已焊死。🫡');
}

// importFixedCosts 已废弃：固定支出估算仅在 FIXED_COSTS 配置和月报中使用，
// 不再写入取引，避免污染日次每日合计和当月实际总计。


/* ========= 测试工具 ========= */

/**
 * 一次性发送全部 6 套模板的测试邮件（月报版，含固定估算）。
 * 用当月真实数据生成 summary，循环每个 tplId 各发一封，subject 里标注模板 id。
 * 方便在手机端一次对比所有主题的视觉效果。
 *
 * 跑法：Apps Script 编辑器函数下拉选中本函数 → 运行。
 *       ~3 秒后收到 6 封邮件（每封间隔 500ms，防触发 Gmail quota 限流）。
 */
function testSendAllTemplates() {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), 1);
  const summary = generateSummary_(start, now);
  const me = Session.getActiveUser().getEmail();

  let sent = 0;
  let failed = 0;

  EMAIL_TEMPLATES.forEach(function (tplId) {
    try {
      const vm = buildViewModel_(summary, true);  // 月报模式，看全量区块（FIXED + GRAND TOTAL）
      const title = '[TEST ' + tplId + '] 模板预览 · ' + vm.periodStr;
      const html = buildEmailHtml_(vm, title, tplId);
      const plain = buildPlainText_(vm, title);

      GmailApp.sendEmail(me, title, plain, {
        htmlBody: html,
        name: 'SMBC 模板测试'
      });
      sent++;
      Logger.log('[test] sent: ' + tplId);
    } catch (err) {
      failed++;
      Logger.log('[test] FAILED for ' + tplId + ': ' + err.message);
    }
    Utilities.sleep(500);  // 防 Gmail 短时间密集发送触发 quota 限流
  });

  Logger.log('[test] 完成：成功 ' + sent + ' / 失败 ' + failed + ' / 共 ' + EMAIL_TEMPLATES.length);
}


/* ========= 一次性数据迁移工具 ========= */

/**
 * 入口 1：先跑 dry-run，只打印会改什么，不真写入 Sheet。
 * 在 Apps Script 编辑器的函数下拉里选中本函数后点运行。
 */
function runRentMigrationDryRun() {
  migrateRentCategory_(true);
}

/**
 * 入口 2：dry-run 确认无误后，选中本函数运行，真写入 Sheet。
 * 跑完成功后，手动编辑 CATEGORY_GROUPS 里房租那一项，砍掉历史 alias：
 *   { label: '房租', cats: ['房租(浦安・銀行引落)', '房租(エポス代扣・横浜)'] }
 */
function runRentMigrationForReal() {
  migrateRentCategory_(false);
}

/**
 * 一次性迁移：把历史房租分类统一成 canonical 值
 *
 * 背景：早期 predebit_rent 的 category 只写 '房租'，中期改为 '房租(浦安)'，
 *       现在统一为 '房租(浦安・銀行引落)'（含支付方式）。
 *       epos 规则也经历了 '房租(横浜和田町)' → '房租(エポス代扣・横浜)' 的演变。
 *       CATEGORY_GROUPS 暂时保留历史别名兼容聚合，但 Sheet 里多种字符串共存是技术债。
 *       本函数一次性把所有历史值回填成新 canonical 值，跑完后即可砍掉 alias。
 *
 * ※ 函数名带 _ 后缀 = Apps Script 私有函数，不出现在运行下拉。
 *   入口请用上面两个 wrapper：runRentMigrationDryRun / runRentMigrationForReal
 *
 * 迁移规则（按 规则名 K 列精确匹配，不是字符串模糊匹配，确定性 100%）：
 *   predebit_rent       → '房租(浦安・銀行引落)'
 *   predebit_rent_epos  → '房租(エポス代扣・横浜)'
 *
 * 最终分类(M) 列同步策略：
 *   仅当 M === 旧 I 值时同步（说明是自动继承，没被手动修正过）
 *   若 M 被手动改过（不等于自动分类值），则完全不动，尊重人工修正
 */
function migrateRentCategory_(dryRun) {
  dryRun = dryRun !== false;  // 默认 dry-run；必须显式传 false 才真写入

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const catCol = headers.indexOf('分类');
  const ruleCol = headers.indexOf('规则名');
  const finalCatCol = headers.indexOf('最终分类');

  if (catCol < 0 || ruleCol < 0) {
    Logger.log('[migrate] 列缺失，中止。catCol=' + catCol + ' ruleCol=' + ruleCol);
    return;
  }

  const RULE_TO_CAT = {
    'predebit_rent':      '房租(浦安・銀行引落)',
    'predebit_rent_epos': '房租(エポス代扣・横浜)'
  };

  const updates = [];

  for (let i = 1; i < data.length; i++) {
    const rule = String(data[i][ruleCol] || '').trim();
    const cat = String(data[i][catCol] || '').trim();
    const finalCat = finalCatCol >= 0 ? String(data[i][finalCatCol] || '').trim() : '';

    const target = RULE_TO_CAT[rule];
    if (!target) continue;

    const needCatUpdate = (cat !== target);
    const needFinalUpdate = (finalCatCol >= 0 && finalCat === cat && finalCat !== target);

    if (needCatUpdate || needFinalUpdate) {
      updates.push({
        row: i + 1,
        oldCat: cat, newCat: target,
        oldFinal: finalCat, newFinal: needFinalUpdate ? target : finalCat,
        updateFinal: needFinalUpdate
      });
    }
  }

  Logger.log('[migrate] 扫描完成：需更新 ' + updates.length + ' 行（dryRun=' + dryRun + '）');
  updates.slice(0, 30).forEach(u => {
    Logger.log('  row ' + u.row + ': 分类 "' + u.oldCat + '" → "' + u.newCat + '"' +
               (u.updateFinal ? ' | 最终分类 "' + u.oldFinal + '" → "' + u.newFinal + '"' : ''));
  });
  if (updates.length > 30) Logger.log('  ... 仅显示前 30 行，共 ' + updates.length + ' 行');

  if (dryRun) {
    Logger.log('[migrate] dry-run 模式，未写入 Sheet。确认无误后：');
    Logger.log('           新建一个小函数  function _doMigrateRent() { migrateRentCategory_(false); }');
    Logger.log('           选中 _doMigrateRent 运行即可真跑');
    return;
  }

  updates.forEach(u => {
    sheet.getRange(u.row, catCol + 1).setValue(u.newCat);
    if (u.updateFinal) sheet.getRange(u.row, finalCatCol + 1).setValue(u.newFinal);
  });
  SpreadsheetApp.flush();

  Logger.log('[migrate] 完成，已更新 ' + updates.length + ' 行');
  Logger.log('[migrate] 下一步：手动编辑 code.gs 顶部 CATEGORY_GROUPS 里房租那一项，砍掉历史 alias：');
  Logger.log("             { label: '房租', cats: ['房租(浦安・銀行引落)', '房租(エポス代扣・横浜)'] }");
}

