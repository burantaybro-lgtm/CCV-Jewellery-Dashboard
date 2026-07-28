import { NextResponse } from "next/server";

const TROY_OUNCE_GRAMS = 31.1034768;

export async function GET() {
  try {
    const [goldResponse, silverResponse, fxResponse] = await Promise.all([
      fetch("https://api.gold-api.com/price/XAU", { headers: { Accept: "application/json" } }),
      fetch("https://api.gold-api.com/price/XAG", { headers: { Accept: "application/json" } }),
      fetch("https://open.er-api.com/v6/latest/USD", { headers: { Accept: "application/json" } })
    ]);
    if (!goldResponse.ok || !silverResponse.ok || !fxResponse.ok) throw new Error("Price service unavailable");
    const gold = await goldResponse.json() as { price?: number };
    const silver = await silverResponse.json() as { price?: number };
    const fx = await fxResponse.json() as { rates?: { NZD?: number } };
    if (!gold.price || !silver.price || !fx.rates?.NZD) throw new Error("Incomplete price response");
    return NextResponse.json({
      gold: Number(((gold.price * fx.rates.NZD) / TROY_OUNCE_GRAMS).toFixed(2)),
      silver: Number(((silver.price * fx.rates.NZD) / TROY_OUNCE_GRAMS).toFixed(3)),
      updatedAt: new Date().toISOString().slice(0, 10),
      source: "daily"
    }, { headers: { "Cache-Control": "public, max-age=3600, s-maxage=86400" } });
  } catch {
    return NextResponse.json({ error: "Daily price service unavailable" }, { status: 503 });
  }
}
