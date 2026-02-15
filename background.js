chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === "start_relay_capture") {
    
    // 期間文字列を作成
    const periodStr = (request.start && request.end) 
                      ? `${request.start}_to_${request.end}` 
                      : "all_period";

    // フォルダ名を組み立て
    // 例: booth_証憑_2025-09_to_2025-09_20260214-1437
    const folderName = `booth_証憑_${periodStr}_${request.uniqueId}`;
    
    processIds(request.orders, request.format, folderName);
    sendResponse({ status: "started" });
  }
  return true;
});
async function processIds(orders, format, folderName) {
  for (let i = 0; i < orders.length; i++) {
    const order = orders[i];
    let tab = null;

    try {
      tab = await chrome.tabs.create({ url: `https://accounts.booth.pm/orders/${order.id}`, active: true });

      await new Promise((resolve) => {
        const listener = (tabId, info) => {
          if (tabId === tab.id && info.status === 'complete') {
            chrome.tabs.onUpdated.removeListener(listener);
            resolve();
          }
        };
        chrome.tabs.onUpdated.addListener(listener);
        setTimeout(resolve, 15000);
      });

      await new Promise(r => setTimeout(r, 4000));

      let fileName = `booth_order_${order.id}.png`;
      if (format === 'dencho') {
        fileName = `${order.date}_株式会社ピクシブ_${order.price}_注文詳細_${order.id}.png`;
      }

      const target = { tabId: tab.id };
      await chrome.debugger.attach(target, "1.3");
      const { contentSize } = await chrome.debugger.sendCommand(target, "Page.getLayoutMetrics");
      
      // 撮影実行とデータ取得
      const response = await chrome.debugger.sendCommand(target, "Page.captureScreenshot", {
        clip: { x: 0, y: 0, width: contentSize.width, height: contentSize.height, scale: 1 },
        captureBeyondViewport: true
      });

      // エラー箇所修正：response.data を使用
      if (response && response.data) {
        await chrome.downloads.download({
          url: "data:image/png;base64," + response.data,
          filename: `${folderName}/${fileName}` // フォルダ名を付与
        });
      }

      await chrome.debugger.detach(target);
      await chrome.tabs.remove(tab.id);
    } catch (e) {
      console.error("Error:", e);
      if (tab?.id) await chrome.tabs.remove(tab.id);
    }
    await new Promise(r => setTimeout(r, 1000));
  }
}