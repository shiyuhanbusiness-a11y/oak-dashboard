export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type ProcessPayload = {
  query?: string;
};

export async function POST(request: Request) {
  const webhookUrl = process.env.N8N_WEBHOOK_URL;

  if (!webhookUrl) {
    return Response.json(
      { error: "服务端未配置 N8N_WEBHOOK_URL。" },
      { status: 500 }
    );
  }

  let payload: ProcessPayload;
  try {
    payload = await request.json();
  } catch {
    return Response.json(
      { error: "请求体格式错误，请使用 JSON 格式。" },
      { status: 400 }
    );
  }

  const query = payload.query?.trim();
  if (!query) {
    return Response.json({ error: "query 不能为空。" }, { status: 400 });
  }

  try {
    const upstreamResponse = await fetch(webhookUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ query }),
      cache: "no-store",
    });

    const contentType = upstreamResponse.headers.get("content-type") ?? "";
    const responseText = await upstreamResponse.text();

    if (!upstreamResponse.ok) {
      console.error("[/api/process] 上游 n8n 返回非 2xx", {
        upstreamStatus: upstreamResponse.status,
        upstreamStatusText: upstreamResponse.statusText,
        queryPreview: query.slice(0, 100),
        details: responseText.slice(0, 500),
      });

      return Response.json(
        {
          error: "调用 n8n Webhook 失败。",
          status: upstreamResponse.status,
          upstreamStatus: upstreamResponse.status,
          upstreamStatusText: upstreamResponse.statusText,
          queryPreview: query.slice(0, 100),
          details: responseText.slice(0, 500),
        },
        { status: 502 }
      );
    }

    if (contentType.includes("application/json")) {
      try {
        const data = JSON.parse(responseText);
        return Response.json({ success: true, data });
      } catch {
        return Response.json({
          success: true,
          data: responseText,
        });
      }
    }

    return Response.json({
      success: true,
      data: responseText,
    });
  } catch (error) {
    console.error("[/api/process] 请求异常", {
      queryPreview: query.slice(0, 100),
      message: error instanceof Error ? error.message : "未知错误",
    });

    return Response.json(
      {
        error: "代理请求异常。",
        queryPreview: query.slice(0, 100),
        details: error instanceof Error ? error.message : "未知错误",
      },
      { status: 500 }
    );
  }
}
