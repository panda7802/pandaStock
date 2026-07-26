import fs from "node:fs/promises";
import path from "node:path";

const outDir = path.resolve("raw_data");
await fs.mkdir(outDir, { recursive: true });

const targets = [
  ["QQQ", "105.QQQ"],
  ["DIA", "106.DIA"],
  ["SPY", "106.SPY"],
  ["KO", "106.KO"],
  ["BRK.B", "106.BRK.B"],
];

function urlFor(secid) {
  const params = new URLSearchParams({
    secid,
    klt: "101",
    fqt: "1",
    lmt: "1000000",
    end: "20500101",
    iscca: "1",
    fields1: "f1,f2,f3,f4,f5,f6,f7,f8",
    fields2: "f51,f52,f53,f54,f55,f56,f57,f58,f59,f60,f61",
  });
  return `https://push2his.eastmoney.com/api/qt/stock/kline/get?${params}`;
}

async function fetchWithRetry(url, attempts = 12) {
  let lastError;
  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      const response = await fetch(url, {
        headers: {
          "User-Agent": "Mozilla/5.0",
          Referer: "https://quote.eastmoney.com/",
        },
      });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      return await response.text();
    } catch (error) {
      lastError = error;
      if (attempt < attempts) {
        await new Promise((resolve) => setTimeout(resolve, Math.min(15000, 1000 * 2 ** Math.min(attempt, 4))));
      }
    }
  }
  throw lastError;
}

for (const [symbol, secid] of targets) {
  const output = path.join(outDir, `${symbol}_eastmoney.json`);
  const text = await fetchWithRetry(urlFor(secid));
  JSON.parse(text);
  await fs.writeFile(output, text, "utf8");
  const payload = JSON.parse(text);
  const rows = payload?.data?.klines ?? [];
  console.log(`${symbol}: ${rows.length} rows; ${rows.at(0) ?? "no data"}; ${rows.at(-1) ?? "no data"}`);
}
