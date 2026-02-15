// ==========================================
// 1. プルダウン（年・月）の選択肢を作成する
// ==========================================
const setupDropdowns = () => {
  const currentYear = new Date().getFullYear();
  const years = [document.getElementById('startYear'), document.getElementById('endYear')];
  const months = [document.getElementById('startMonth'), document.getElementById('endMonth')];

  // 年の選択肢 (2015年〜現在)
  years.forEach(select => {
    for (let y = currentYear; y >= 2015; y--) {
      select.add(new Option(`${y}年`, y));
    }
  });

  // 月の選択肢 (01月〜12月)
  months.forEach(select => {
    for (let m = 1; m <= 12; m++) {
      const mm = m.toString().padStart(2, '0');
      select.add(new Option(`${m}月`, mm));
    }
  });
};

// 起動時に実行
setupDropdowns();

// 共通: タイムスタンプ生成 (YYYYMMDD-HHMM)
const getFormattedTimestamp = () => {
  const now = new Date();
  return now.getFullYear() +
    (now.getMonth() + 1).toString().padStart(2, '0') +
    now.getDate().toString().padStart(2, '0') + "-" +
    now.getHours().toString().padStart(2, '0') +
    now.getMinutes().toString().padStart(2, '0');
};

// ==========================================
// 定数定義
// ==========================================
const SALES_SCAN_MONTHS_BEFORE = 12; // 振込遅延を考慣したスキャン期間2
const BANK_FEE_THRESHOLD = 30000;    // 銀行振込手数料の閾値
const BANK_FEE_HIGH = 300;           // 30,000円以上の振込手数料
const BANK_FEE_LOW = 200;            // 30,000円未満の振込手数料

// 会計ソフトの選択によって入力欄をグレーアウトさせる
const toggleCreditInput = () => {
  const formatSelect = document.getElementById('exportFormat');
  const creditInput = document.getElementById('creditAccount');

  if (formatSelect.value === 'simple') {
    // 単式簿記ならグレーアウト
    creditInput.disabled = true;
    creditInput.style.backgroundColor = '#e9e9e9'; // 薄いグレー
    creditInput.style.color = '#999';            // 文字を薄く
  } else {
    // 複式簿記なら有効化
    creditInput.disabled = false;
    creditInput.style.backgroundColor = '#fff';
    creditInput.style.color = '#000';
  }
};

// 選択が変わった時に実行
document.getElementById('exportFormat').addEventListener('change', toggleCreditInput);

// 起動時にも一度実行して現在の状態を反映
toggleCreditInput();

// ==========================================
// 2. CSV抽出ボタンの処理
// ==========================================
document.getElementById('extract').addEventListener('click', async () => {
  const creditAccount = document.getElementById('creditAccount').value;
  const exportFormat = document.getElementById('exportFormat').value; // 会計ソフトの形式

  const sYear = document.getElementById('startYear').value;
  const sMonth = document.getElementById('startMonth').value;
  const eYear = document.getElementById('endYear').value;
  const eMonth = document.getElementById('endMonth').value;

  // --- 期間の入力チェック ---
  if (sYear && sMonth && eYear && eMonth) {
    const startVal = parseInt(sYear + sMonth);
    const endVal = parseInt(eYear + eMonth);

    if (startVal > endVal) {
      alert("【無効な期間】\n終了月が開始月よりも過去になっています。\n期間を正しく設定してください。");
      return;
    }
  }

  // 期間文字列の作成 (例: "2025-01")
  const startMonth = (sYear && sMonth) ? `${sYear}-${sMonth}` : "";
  const endMonth = (eYear && eMonth) ? `${eYear}-${eMonth}` : "";

  // ブラウザのタブを取得してスクリプトを実行
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  if (!tab.url.includes("accounts.booth.pm/orders")) {
    alert("BOOTHの購入履歴ページを開いた状態で実行してください。");
    return;
  }

  const formatSelect = document.getElementById('exportFormat');
  const exportFormatLabel = formatSelect.options[formatSelect.selectedIndex].text;
  const timestamp = getFormattedTimestamp();

  chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: exportAllPagesToCsv,
    args: [creditAccount, startMonth, endMonth, exportFormat, exportFormatLabel, timestamp]
  });
});

// ==========================================
// 証憑画像の一括保存処理（修正・安定版）
// ==========================================
document.getElementById('captureImages').addEventListener('click', async () => {
  const sYear = document.getElementById('startYear').value;
  const sMonth = document.getElementById('startMonth').value;
  const eYear = document.getElementById('endYear').value;
  const eMonth = document.getElementById('endMonth').value;

  // HTMLに追加したプルダウンから値を取得
  const filenameFormat = document.getElementById('filenameFormat').value;

  if (sYear && sMonth && eYear && eMonth) {
    if (parseInt(sYear + sMonth) > parseInt(eYear + eMonth)) {
      alert("【無効な期間】\n終了月が開始月よりも過去になっています。\n期間を正しく設定してください。");
      return;
    }
  }

  const startMonth = (sYear && sMonth) ? `${sYear}-${sMonth}` : "";
  const endMonth = (eYear && eMonth) ? `${eYear}-${eMonth}` : "";

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  if (!tab.url.includes("accounts.booth.pm/orders")) {
    alert("BOOTHの購入履歴ページを開いた状態で実行してください。");
    return;
  }


  const timestamp = getFormattedTimestamp();
  const uniqueId = timestamp;

  chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: async (sMonthStr, eMonthStr, filenameFormat, uniqueId) => {
      const msg = "解析には少し時間がかかります。\n解析完了後、注文詳細ページが自動で開き撮影が始まります。\n完了までしばらくお待ちください。\n\n※進行状況は F12キー（コンソール）で確認できます。";
      if (!window.confirm(msg)) return;

      console.log("%c=== 証憑画像解析開始 ===", "color: blue; font-weight: bold;");

      const parseDate = (str, isEnd) => {
        if (!str) return null;
        const [y, m] = str.split('-');
        return isEnd ? new Date(y, m, 0, 23, 59, 59) : new Date(y, parseInt(m) - 1, 1, 0, 0, 0);
      };
      const startDate = parseDate(sMonthStr, false);
      const endDate = parseDate(eMonthStr, true);

      let orderMap = new Map();
      let nextUrl = "https://accounts.booth.pm/orders?page=1";
      let stopProcessing = false;
      let page = 1;

      while (nextUrl && !stopProcessing && page <= 50) {
        const res = await fetch(nextUrl, { credentials: "include" });
        const doc = new DOMParser().parseFromString(await res.text(), "text/html");
        const links = Array.from(doc.querySelectorAll("a[href*='/orders/']"));
        const pageOrderEntries = [...new Set(links.map(l => {
          const m = l.href.match(/\/orders\/(\d+)$/);
          return m ? { id: m[1], url: l.href } : null;
        }).filter(x => x))];

        for (const order of pageOrderEntries) {
          if (orderMap.has(order.id)) continue;
          try {
            const dRes = await fetch(order.url, { credentials: "include" });
            const dDoc = new DOMParser().parseFromString(await dRes.text(), "text/html");
            const dateStr = dDoc.body.innerText.match(/\d{4}[\/\-]\d{2}[\/\-]\d{2}/)?.[0] || "";
            if (!dateStr) continue;

            const orderDate = new Date(dateStr.replace(/-/g, '/') + " 00:00:00");
            if (startDate && orderDate < startDate) { stopProcessing = true; break; }
            if (endDate && orderDate > endDate) continue;
            if (dDoc.querySelector('.order-state.cancelled')) continue;

            // ★金額の取得を追加
            const priceMatch = dDoc.body.innerText.match(/お支払金額.*?¥\s*([\d,]+)/);
            const price = priceMatch ? priceMatch[1].replace(/,/g, "") : "0";

            orderMap.set(order.id, {
              id: order.id,
              date: dateStr.replace(/[\/\-]/g, ""), // 20250901形式
              price: price
            });
            console.log(`[解析完了] ID:${order.id} 金額:${price}`);
          } catch (e) { }
        }
        if (stopProcessing) break;
        const nextBtn = doc.querySelector('a[rel="next"], a.next_page');
        nextUrl = (nextBtn && nextBtn.href) ? nextBtn.href : null;
        page++;
        await new Promise(r => setTimeout(r, 400));
      }

      const orders = Array.from(orderMap.values()).filter(v => v !== null);

      if (orders.length === 0) {
        alert("該当する注文が見つかりませんでした。"); return;
      }

      // バックグラウンドへ送信 (コンテンツスクリプトから直接)
      chrome.runtime.sendMessage({
        action: "start_relay_capture",
        orders: orders,
        format: filenameFormat, // 引数で受け取る
        start: sMonthStr,       // 引数で受け取る (sMonthStr/eMonthStr)
        end: eMonthStr,
        uniqueId: uniqueId      // 引数で受け取る
      });

    },
    args: [startMonth, endMonth, filenameFormat, uniqueId]
  });
});

// ==========================================
// 2.5 売上管理仕訳CSV生成 (振込日基準の精密抽出版)
// ==========================================
document.getElementById('extractSales').addEventListener('click', async () => {
  const sYear = parseInt(document.getElementById('startYear').value);
  const sMonth = parseInt(document.getElementById('startMonth').value);
  const eYear = parseInt(document.getElementById('endYear').value);
  const eMonth = parseInt(document.getElementById('endMonth').value);
  const depositAcc = document.getElementById('salesDebitAccount').value;
  const payoutMethod = document.getElementById('payoutMethod').value;

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab.url.includes("manage.booth.pm")) {
    alert("BOOTHのショップ管理画面を開いた状態で実行してください。");
    return;
  }

  if (!sYear || !sMonth || !eYear || !eMonth) {
    alert("期間を選択してください。");
    return;
  }

  const startVal = sYear * 100 + sMonth;
  const endVal = eYear * 100 + eMonth;

  if (startVal > endVal) {
    alert("【無効な期間】\n終了月が開始月よりも過去になっています。\n期間を正しく設定してください。");
    return;
  }

  const exportFormat = document.getElementById('exportFormat').value;
  const formatSelect = document.getElementById('exportFormat');
  const exportFormatLabel = formatSelect.options[formatSelect.selectedIndex].text;
  const timestamp = getFormattedTimestamp();

  chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: async (sY, sM, eY, eM, dAcc, timestamp, exportFormat, exportFormatLabel, feeType) => {
      try {
        const msg = `解析を開始します。\n形式: ${exportFormatLabel}\nログはF12キーのコンソールで確認できます。`;
        if (!window.confirm(msg)) return;
        console.log(`=== 解析開始 (${sY}/${sM} 〜 ${eY}/${eM}) ===`);

        // インジェクション先で必要な定数を定義
        const BANK_FEE_THRESHOLD = 30000;
        const BANK_FEE_HIGH = 300;
        const BANK_FEE_LOW = 200;

        // ユーザーが指定した期間の開始日と終了日
        const startLimit = new Date(sY, sM - 1, 1);
        const endLimit = new Date(eY, eM, 0, 23, 59, 59); // 月の最終日を正しく取得

        // 判定用:振込履歴を遡るための月リスト (指定期間の12ヶ月前から終了月まで)
        const scanMonths = [];
        let curY = sY, curM = sM;
        // 振込遅延を考慮して1年前からスキャン
        let scanY = sY - 1, scanM = sM;
        while (scanY < eY || (scanY === eY && scanM <= eM)) {
          scanMonths.push({ y: scanY, m: scanM });
          scanM++; if (scanM > 12) { scanM = 1; scanY++; }
        }

        const getValueByLabel = (doc, labelText) => {
          const labels = Array.from(doc.querySelectorAll('.u-tpg-caption2'));
          const targetLabel = labels.find(el => el.innerText.includes(labelText));
          return targetLabel?.nextElementSibling?.innerText.trim() || null;
        };

        // --- A. 基本設定の取得 ---
        let shopName = "ショップ";
        try {
          const sRes = await fetch("https://manage.booth.pm/settings", { credentials: "include" });
          const sDoc = new DOMParser().parseFromString(await sRes.text(), "text/html");
          shopName = sDoc.querySelector('input#shop_name')?.value || "ショップ";
        } catch (e) {
          console.warn('ショップ名の取得に失敗しました:', e);
        }

        // let feeType = "bank";
        // try {
        //   const pRes = await fetch("https://manage.booth.pm/payout_account", { credentials: "include" });
        //   const pDoc = new DOMParser().parseFromString(await pRes.text(), "text/html");
        //   const payoutMethod = Array.from(pDoc.querySelectorAll('.l-row')).find(row => row.innerText.includes("振込方法"));
        //   if (payoutMethod && payoutMethod.innerText.includes("PayPal")) feeType = "paypal";
        // } catch (e) {
        //   console.warn('振込方法の取得に失敗しました:', e);
        // }

        // --- B. 全スキャン対象月の解析 ---
        const salesData = [];
        for (const item of scanMonths) {
          const url = `https://manage.booth.pm/sales/${item.y}/${item.m.toString().padStart(2, '0')}`;
          try {
            const res = await fetch(url, { credentials: "include" });
            const doc = new DOMParser().parseFromString(await res.text(), "text/html");

            const netIncome = parseInt(doc.querySelector('.u-tpg-title2 b')?.innerText.replace(/[¥,]/g, "")) || 0;
            if (netIncome === 0) continue;

            const totalSales = parseInt(Array.from(doc.querySelectorAll('.lo-grid-cell')).find(el => el.innerText.trim() === "総売上")?.parentElement.querySelector('.text-right')?.innerText.replace(/[¥,]/g, "")) || 0;
            const kDateRaw = getValueByLabel(doc, "受取金額確定日") || "";
            const fDateRaw = getValueByLabel(doc, "振込日") || "";

            const fMatch = fDateRaw.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
            const fDateObj = fMatch ? new Date(fMatch[1], fMatch[2] - 1, fMatch[3]) : null;

            salesData.push({
              yearMonth: `${item.y}年${item.m}月`,
              total: totalSales,
              fee: totalSales - netIncome,
              net: netIncome,
              kDate: kDateRaw.replace(/[年月日]/g, '/').replace(/\/$/, ''),
              fDateObj: fDateObj,
              fDateStr: fMatch ? `${fMatch[1]}/${fMatch[2]}/${fMatch[3]}` : "未定"
            });
            console.log(`[スキャン] ${item.y}/${item.m}: 振込予定=${fMatch ? fMatch[0] : "未定"}`);
          } catch (e) {
            console.warn(`売上データの取得に失敗しました (${item.y}/${item.m}):`, e);
          }
          await new Promise(r => setTimeout(r, 400));
        }

        // --- C. 仕訳データの構築とソート ---
        let allEntries = [];
        const furikomiGroup = new Map();

        salesData.forEach(r => {
          // 1. 月末売上仕訳
          const kDateObj = new Date(r.kDate);
          const saleMonthLastDayObj = new Date(kDateObj.getFullYear(), kDateObj.getMonth() + 1, 0);

          if (saleMonthLastDayObj >= startLimit && saleMonthLastDayObj <= endLimit) {
            allEntries.push({
              type: 'sales',
              dateObj: saleMonthLastDayObj,
              data: r
            });
          }

          // 2. 振込時仕訳
          if (r.fDateObj && r.fDateObj >= startLimit && r.fDateObj <= endLimit) {
            const fDateStr = r.fDateStr;
            if (!furikomiGroup.has(fDateStr)) {
              furikomiGroup.set(fDateStr, { netSum: 0, months: [], dateObj: r.fDateObj });
            }
            const group = furikomiGroup.get(fDateStr);
            group.netSum += r.net;
            group.months.push(r.yearMonth);
          }
        });

        // 振込グループを展開してエントリに追加
        furikomiGroup.forEach((data, fDate) => {
          allEntries.push({
            type: 'transfer',
            dateObj: data.dateObj,
            data: data,
            fDateStr: fDate
          });
        });

        // 日付順にソート
        allEntries.sort((a, b) => a.dateObj - b.dateObj);

        // --- D. CSV生成 (共通フォーマット対応) ---
        let csvRows = [];

        // (D-1) ヘッダー定義
        const getHeader = () => {
          switch (exportFormat) {
            case 'freee':
              return ['[表題行]', '日付', '伝票番号', '決算整理仕訳', '借方勘定科目', '借方科目コード', '借方補助科目', '借方取引先', '借方取引先コード', '借方部門', '借方品目', '借方メモタグ', '借方セグメント1', '借方セグメント2', '借方セグメント3', '借方金額', '借方税区分', '借方税額', '貸方勘定科目', '貸方科目コード', '貸方補助科目', '貸方取引先', '貸方取引先コード', '貸方部門', '貸方品目', '貸方メモタグ', '貸方セグメント1', '貸方セグメント2', '貸方セグメント3', '貸方金額', '貸方税区分', '貸方税額', '摘要'];
            case 'mf':
              return ['取引No', '取引日', '借方勘定科目', '借方補助科目', '借方部門', '借方取引先', '借方税区分', '借方インボイス', '借方金額(円)', '借方税額', '貸方勘定科目', '貸方補助科目', '貸方部門', '貸方取引先', '貸方税区分', '貸方インボイス', '貸方金額(円)', '貸方税額', '摘要', '仕訳メモ', 'タグ', 'MF仕訳タイプ', '決算整理仕訳'];
            case 'yayoi':
              return ['識別フラグ', '伝票No', '決算', '取引日付', '借方勘定科目', '借方補助科目', '借方部門', '借方税区分', '借方金額', '借方税金額', '貸方勘定科目', '貸方補助科目', '貸方部門', '貸方税区分', '貸方金額', '貸方税金額', '摘要', '番号', '期日', 'タイプ', '生成元', '仕訳メモ', '付箋1', '付箋2', '調整', '借方取引先名', '貸方取引先名'];
            case 'simple':
              return ['注文番号', '日付', '項目', '収入', '支出', '摘要', '振込先'];
            default:
              return ['注文番号', '日付', '借方勘定科目', '借方金額', '貸方勘定科目', '貸方金額', '摘要'];
          }
        };

        csvRows.push(getHeader().join(','));

        // (D-2) 行追加関数
        // debitPartner, creditPartner を追加引数として受け取る
        const addRow = (flag, id, date, debitK, debitA, creditK, creditA, rem, debitPartner = "", creditPartner = "") => {
          let row = [];
          const dateStr = date.replace(/-/g, '/');

          switch (exportFormat) {
            case 'freee':
              row = ['[明細行]', dateStr, id, '', debitK, '', '', debitPartner, '', '', '', '', '', '', '', debitA, '', '', creditK, '', '', creditPartner, '', '', '', '', '', '', '', creditA, '', '', rem];
              break;
            case 'mf':
              row = [id, dateStr, debitK, '', '', debitPartner, '', '', debitA, '', creditK, '', '', creditPartner, '', '', creditA, '', rem, '', '', '', ''];
              break;
            case 'yayoi':
              row = [flag, id, '', dateStr, debitK, '', '', '', debitA, '', creditK, '', '', '', creditA, '', rem, '', '', '', '', '', '', '', '', debitPartner, creditPartner];
              break;
            case 'simple':
              // 簡易帳簿形式: 注文番号, 日付, 項目, 収入, 支出, 摘要, 振込先
              // debitKに項目名、debitAに収入金額、creditAに支出金額、creditKに振込先を格納
              row = [id, dateStr, debitK, debitA, creditA, rem, creditK];
              break;
            default:
              row = [id, dateStr, debitK, debitA, creditK, creditA, rem];
          }
          csvRows.push(row.join(','));
        };

        // (D-3) ID生成と行追加ループ
        // IDカウンター管理: "YYYYMM" -> sequence number
        const idCounters = new Map();

        const generateId = (dateObj) => {
          const y = dateObj.getFullYear(); // 2025
          const m = dateObj.getMonth() + 1; // 1
          const yy = y.toString().slice(-2);
          const mm = m.toString().padStart(2, '0');
          const key = `${yy}${mm}`;

          let count = idCounters.get(key) || 0;
          count++;
          idCounters.set(key, count);

          return `${key}${count.toString().padStart(3, '0')}`; // YYMMCCC
        };

        allEntries.forEach(entry => {
          const id = generateId(entry.dateObj);
          const dateStr = entry.dateObj.toLocaleDateString('ja-JP', { year: 'numeric', month: '2-digit', day: '2-digit' }).replace(/\//g, '-');

          if (entry.type === 'sales') {
            // 月末売上仕訳
            const r = entry.data;
            const summaryMain = `BOOTH売上${r.yearMonth}(${shopName})`;
            const summaryFee = `BOOTH手数料(${shopName})`;
            const flag = (exportFormat === 'yayoi') ? '2110' : '';

            if (exportFormat === 'simple') {
              // 簡易帳簿形式: 売上を収入に記録
              addRow(flag, id, dateStr, '売上(BOOTH)', r.total, '', '', summaryMain);
              // 手数料を支出に記録
              if (r.fee > 0) {
                addRow(flag, id, dateStr, '支払手数料(BOOTH)', '', '', r.fee, summaryFee);
              }
            } else {
              // 複式簿記: 1行目: 売掛金 / 売上高
              addRow(flag, id, dateStr, "売掛金", r.net, "売上高", r.total, summaryMain, "株式会社ピクシブ", "ECサイト売上(BOOTH)");
              // 2行目: 支払手数料 / (空)
              addRow(flag, id, dateStr, "支払手数料", r.fee, "", "", summaryFee, "株式会社ピクシブ", "");
            }

          } else if (entry.type === 'transfer') {
            // 振込時仕訳
            const data = entry.data;
            const tFee = (feeType === "bank") ? (data.netSum >= BANK_FEE_THRESHOLD ? BANK_FEE_HIGH : BANK_FEE_LOW) : 0;
            const finalAmt = data.netSum - tFee;
            const monthInfo = data.months.join("・");
            const summaryTransfer = `BOOTH売上振込${monthInfo}(${shopName})`;
            const summaryTransferFee = `BOOTH振込手数料(${shopName})`;
            const flag = (exportFormat === 'yayoi') ? '2000' : '';

            if (exportFormat === 'simple') {
              // 簡易帳簿形式: 振込入金を収入に記録、振込先に預金名を記録
              addRow(flag, id, dateStr, '振込入金(BOOTH)', finalAmt, dAcc, '', summaryTransfer);
              // 振込手数料を支出に記録
              if (tFee > 0) {
                addRow(flag, id, dateStr, '支払手数料(BOOTH)', '', '', tFee, summaryTransferFee);
              }
            } else {
              // 複式簿記: 1行目: 預金 / 売掛金
              addRow(flag, id, dateStr, dAcc, finalAmt, "売掛金", data.netSum, summaryTransfer, "", "株式会社ピクシブ");
              // 2行目: 支払手数料 / (空)
              if (tFee > 0) {
                addRow(flag, id, dateStr, "支払手数料", tFee, "", "", summaryTransferFee, "株式会社ピクシブ", "");
              }
            }
          }
        });

        const csvContent = "\uFEFF" + csvRows.join("\n");
        const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);

        const pad = (n) => n.toString().padStart(2, '0');
        const startStr = `${sY}-${pad(sM)}`;
        const endStr = `${eY}-${pad(eM)}`;

        // booth_sales_simple_2025-01_to_2025-12_20260214-0446.csv
        a.download = `booth_sales_${exportFormat}_${startStr}_to_${endStr}_${timestamp}.csv`;
        a.click();

        alert(`抽出完了！\n合計 ${allEntries.length} 件の仕訳データを保存しました。`);
      } catch (e) {
        console.error("CSV生成中に致命的なエラーが発生しました:", e);
        alert("エラーが発生したためCSVを生成できませんでした。詳細はコンソール（F12）を確認してください。");
      }
    },
    args: [sYear, sMonth, eYear, eMonth, depositAcc, timestamp, exportFormat, exportFormatLabel, payoutMethod]
  });
});

// ==========================================
// 3. メイン処理 (BOOTHのページ内で動く関数)
// ==========================================
async function exportAllPagesToCsv(creditAccountName, startMonthStr, endMonthStr, exportFormat, exportFormatLabel, timestamp) {
  console.log("=== BOOTH CSV出力処理開始 ===");
  console.log(`形式: ${exportFormat}, 貸方: ${creditAccountName}, 期間: ${startMonthStr}～${endMonthStr}`);

  let csvRows = [];
  let successCount = 0; // 成功した注文数
  const clientName = "株式会社ピクシブ";

  // --------------------------------------------------
  // (A) ヘッダー定義
  // --------------------------------------------------
  const getHeader = () => {
    switch (exportFormat) {
      case 'freee':
        return ['[表題行]', '日付', '伝票番号', '決算整理仕訳', '借方勘定科目', '借方科目コード', '借方補助科目', '借方取引先', '借方取引先コード', '借方部門', '借方品目', '借方メモタグ', '借方セグメント1', '借方セグメント2', '借方セグメント3', '借方金額', '借方税区分', '借方税額', '貸方勘定科目', '貸方科目コード', '貸方補助科目', '貸方取引先', '貸方取引先コード', '貸方部門', '貸方品目', '貸方メモタグ', '貸方セグメント1', '貸方セグメント2', '貸方セグメント3', '貸方金額', '貸方税区分', '貸方税額', '摘要', '決済方法'];
      case 'mf':
        return ['取引No', '取引日', '借方勘定科目', '借方補助科目', '借方部門', '借方取引先', '借方税区分', '借方インボイス', '借方金額(円)', '借方税額', '貸方勘定科目', '貸方補助科目', '貸方部門', '貸方取引先', '貸方税区分', '貸方インボイス', '貸方金額(円)', '貸方税額', '摘要', '仕訳メモ', 'タグ', 'MF仕訳タイプ', '決算整理仕訳'];
      case 'yayoi':
        return ['識別フラグ', '伝票No', '決算', '取引日付', '借方勘定科目', '借方補助科目', '借方部門', '借方税区分', '借方金額', '借方税金額', '貸方勘定科目', '貸方補助科目', '貸方部門', '貸方税区分', '貸方金額', '貸方税金額', '摘要', '番号', '期日', 'タイプ', '生成元', '仕訳メモ', '付箋1', '付箋2', '調整', '借方取引先名', '貸方取引先名'];
      case 'simple':
        return ['注文番号', '日付', '項目', '収入', '支出', '摘要'];
      default: // double (仕訳帳形式)
        return ['注文番号', '日付', '借方勘定科目', '借方金額', '貸方勘定科目', '貸方金額', '摘要', '決済方法'];
    }
  };

  csvRows.push(getHeader().join(','));

  // --------------------------------------------------
  // (B) データ行作成用関数
  // --------------------------------------------------
  const addRow = (flag, id, date, debitK, debitA, creditK, creditA, rem, method) => {
    let row = [];
    // 日付を YYYY/MM/DD 形式に統一
    const dateStr = date.replace(/-/g, '/');

    switch (exportFormat) {
      case 'freee':
        row = ['[明細行]', dateStr, id, '', debitK, '', '', clientName, '', '', '', '', '', '', '', debitA, '', '', creditK, '', '', '', '', '', '', '', '', '', '', creditA, '', '', rem, method];
        break;

      case 'mf':
        row = [id, dateStr, debitK, '', '', clientName, '', '', debitA, '', creditK, '', '', '', '', '', creditA, '', rem, '', '', '', ''];
        break;

      case 'yayoi':
        row = [flag, id, '', dateStr, debitK, '', '', '', debitA, '', creditK, '', '', '', creditA, '', rem, '', '', '', '', '', '', '', '', clientName, ''];
        break;

      case 'simple':
        // 簡易帳簿形式: 注文番号, 日付, 項目, 収入, 支出, 摘要
        row = [id, dateStr, '消耗品費(BOOTH)', '', debitA, rem];
        break;

      default: // double (仕訳帳形式)
        row = [id, dateStr, debitK, debitA, creditK, creditA, rem, method];
    }
    // 配列をカンマ区切り文字列にして追加
    csvRows.push(row.join(','));
  };

  // --------------------------------------------------
  // (C) メインループ処理
  // --------------------------------------------------
  const paymentMethods = ["pixivcoban", "ピクシブかんたん決済", "PayPal決済", "クレジットカード", "楽天ペイ", "銀行・コンビニ決済"];
  const processedOrderIds = new Set();

  let startDate = null;
  let endDate = null;
  if (startMonthStr) {
    startDate = new Date(startMonthStr + "-01");
    startDate.setHours(0, 0, 0, 0);
  }
  if (endMonthStr) {
    const [y, m] = endMonthStr.split('-');
    endDate = new Date(y, m, 0); // 月末日
    endDate.setHours(23, 59, 59, 999);
  }

  if (!window.confirm(`解析を開始します。\n形式: ${exportFormatLabel}\nログはF12キーのコンソールで確認できます。`)) return;

  let nextUrl = "https://accounts.booth.pm/orders?page=1";
  let pageCount = 1;
  let stopProcessing = false;

  while (nextUrl && !stopProcessing) {
    console.log(`\n========== ページ ${pageCount} を処理中 ==========`);

    try {
      const response = await fetch(nextUrl, { credentials: "include" });
      const html = await response.text();
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, "text/html");

      const pageLinks = Array.from(doc.querySelectorAll("a[href*='/orders/']"));
      const orderIdsInPage = [...new Set(pageLinks.map(link => {
        const match = link.href.match(/\/orders\/(\d+)$/);
        return match ? match[1] : null;
      }).filter(id => id))];

      console.log(`注文候補数: ${orderIdsInPage.length}件`);

      for (let i = 0; i < orderIdsInPage.length; i++) {
        const id = orderIdsInPage[i];

        try {
          const detailUrl = `https://accounts.booth.pm/orders/${id}`;
          const dRes = await fetch(detailUrl, { credentials: "include" });
          const dHtml = await dRes.text();
          const dDoc = parser.parseFromString(dHtml, "text/html");

          const bodyText = dDoc.body.innerText;
          const dateMatch = bodyText.match(/\d{4}[\/\-]\d{2}[\/\-]\d{2}/);
          const dateStrRaw = dateMatch ? dateMatch[0] : "";
          const orderDate = new Date(dateStrRaw);

          if (startDate && orderDate < startDate) {
            console.log(`[期間外:終了] ${dateStrRaw} - 解析を終了します`);
            stopProcessing = true;
            break;
          }
          if (endDate && orderDate > endDate) {
            console.log(`[期間外:スキップ] ${dateStrRaw}`);
            continue;
          }
          if (processedOrderIds.has(id)) continue;

          if (dDoc.querySelector('.order-state.cancelled')) {
            processedOrderIds.add(id);
            continue;
          }

          console.log(`抽出中... ID:${id} 日付:${dateStrRaw}`);

          let foundMethod = "不明";
          for (const method of paymentMethods) {
            if (bodyText.includes(method)) { foundMethod = method; break; }
          }

          const extractPrice = (regex) => {
            const m = dHtml.match(regex);
            return m ? parseInt(m[1].replace(/,/g, "")) : 0;
          };

          const totalAmount = extractPrice(/お支払金額.*?¥\s*([\d,]+)/);
          const fee = extractPrice(/支払手数料.*?¥\s*([\d,]+)/);
          const shipping = extractPrice(/送料.*?¥\s*([\d,]+)/);
          const boost = extractPrice(/BOOST.*?¥\s*([\d,]+)/);

          const itemLinks = Array.from(dDoc.querySelectorAll("a[href*='/items/']"));
          let productNames = [];
          const foundItems = new Set();
          itemLinks.forEach(link => {
            const name = link.innerText.trim();
            if (name && !name.includes('メッセージ') && !foundItems.has(name)) {
              productNames.push(name.replace(/,/g, ' '));
              foundItems.add(name);
            }
          });
          let remarks = productNames.join(' / ');
          if (shipping > 0) remarks += ` (送料含)`;
          if (boost > 0) remarks += ` (BOOST含)`;

          // --------------------------------------------------
          // (D) CSV行の追加処理
          // --------------------------------------------------
          if (fee > 0) {
            const suppliesAmount = totalAmount - fee;
            if (exportFormat === 'yayoi') {
              addRow('2110', id, dateStrRaw, '消耗品費', suppliesAmount, creditAccountName, totalAmount, remarks, foundMethod);
              addRow('2101', id, dateStrRaw, '支払手数料', fee, '', '', 'BOOTH支払手数料', foundMethod);
            } else {
              addRow('', id, dateStrRaw, '消耗品費', suppliesAmount, creditAccountName, totalAmount, remarks, foundMethod);
              addRow('', id, dateStrRaw, '支払手数料', fee, '', '', 'BOOTH支払手数料', foundMethod);
            }
          } else {
            const flag = (exportFormat === 'yayoi') ? '2000' : '';
            addRow(flag, id, dateStrRaw, '消耗品費', totalAmount, creditAccountName, totalAmount, remarks, foundMethod);
          }

          processedOrderIds.add(id);
          successCount++;

          await new Promise(r => setTimeout(r, 500));

        } catch (e) {
          console.error(`ID:${id} 解析失敗`, e);
        }
      }

      if (stopProcessing) break;
      const nextBtn = doc.querySelector('a[rel="next"], a.next_page');
      if (nextBtn && nextBtn.href) {
        nextUrl = nextBtn.href;
        pageCount++;
      } else {
        nextUrl = null;
      }

    } catch (e) {
      console.error("ページ取得エラー", e);
      nextUrl = null;
    }
  }

  // --------------------------------------------------
  // (E) ファイル保存
  // --------------------------------------------------
  if (successCount === 0) {
    alert("指定期間に該当する注文は見つかりませんでした。");
    return;
  }

  const csvContent = "\uFEFF" + csvRows.join("\n");
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement("a");
  const today = new Date().toISOString().split('T')[0];

  let periodStr = (startMonthStr && endMonthStr) ? `${startMonthStr}_to_${endMonthStr}` : `history_${today}`;
  // booth_order_simple_2025-01_to_2025-12_20260214-0446.csv
  let fileName = `booth_order_${exportFormat}_${periodStr}_${timestamp}.csv`;

  a.href = URL.createObjectURL(blob);
  a.download = fileName;
  a.click();

  alert(`抽出完了！\n合計 ${successCount} 件の注文データを保存しました。`);
}