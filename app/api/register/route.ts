export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request) {
  const webhookUrl = process.env.N8N_REGISTER_WEBHOOK_URL;

  if (!webhookUrl) {
    return Response.json(
      { error: "服务端未配置 N8N_REGISTER_WEBHOOK_URL。" },
      { status: 500 }
    );
  }

  let payload: unknown;
  try {
    payload = await request.json();
  } catch {
    return Response.json(
      { error: "请求体格式错误，请使用 JSON 格式。" },
      { status: 400 }
    );
  }

  try {
    const upstreamResponse = await fetch(webhookUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
      cache: "no-store",
    });

    const contentType = upstreamResponse.headers.get("content-type") ?? "";
    const responseText = await upstreamResponse.text();

    if (!upstreamResponse.ok) {
      return Response.json(
        {
          error: "调用注册 Webhook 失败。",
          status: upstreamResponse.status,
          details: responseText.slice(0, 500),
        },
        { status: 502 }
      );
    }

    if (contentType.includes("application/json")) {
      try {
        return Response.json({ success: true, data: JSON.parse(responseText) });
      } catch {
        return Response.json({ success: true, data: responseText });
      }
    }

    return Response.json({ success: true, data: responseText });
  } catch (error) {
    return Response.json(
      {
        error: "注册代理请求异常。",
        details: error instanceof Error ? error.message : "未知错误",
      },
      { status: 500 }
    );
  }
}
