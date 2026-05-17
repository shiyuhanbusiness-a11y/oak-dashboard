export const runtime = "nodejs";
export const dynamic = "force-dynamic";

function normalizeGoogleSheetCsvUrl(url: string) {
  try {
    const parsed = new URL(url);
    const matched = parsed.pathname.match(/\/spreadsheets\/d\/([^/]+)/);
    if (!parsed.hostname.includes("docs.google.com") || !matched?.[1]) {
      return url;
    }

    const exportUrl = new URL(`https://docs.google.com/spreadsheets/d/${matched[1]}/export`);
    exportUrl.searchParams.set("format", "csv");
    const gid = parsed.searchParams.get("gid");
    if (gid) {
      exportUrl.searchParams.set("gid", gid);
    }
    return exportUrl.toString();
  } catch {
    return url;
  }
}

export async function GET() {
  const rawUrl = process.env.NEXT_PUBLIC_USERS_CSV_URL ?? process.env.NEXT_PUBLIC_SHEET_CSV_URL;
  if (!rawUrl) {
    return Response.json(
      { error: "服务端未配置 NEXT_PUBLIC_USERS_CSV_URL（兼容 NEXT_PUBLIC_SHEET_CSV_URL）。" },
      { status: 500 }
    );
  }

  try {
    const csvUrl = normalizeGoogleSheetCsvUrl(rawUrl);
    const upstreamResponse = await fetch(csvUrl, {
      method: "GET",
      cache: "no-store",
    });

    if (!upstreamResponse.ok) {
      const details = await upstreamResponse.text();
      return Response.json(
        {
          error: "拉取 CSV 数据失败。",
          status: upstreamResponse.status,
          details: details.slice(0, 500),
        },
        { status: 502 }
      );
    }

    const csvText = await upstreamResponse.text();
    return new Response(csvText, {
      status: 200,
      headers: {
        "Content-Type": "text/csv; charset=utf-8",
        "Cache-Control": "no-store",
      },
    });
  } catch (error) {
    return Response.json(
      {
        error: "代理拉取 CSV 异常。",
        details: error instanceof Error ? error.message : "未知错误",
      },
      { status: 500 }
    );
  }
}
