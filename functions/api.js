// Cloudflare Pages Function for /api
export async function onRequest(context) {
  const { request, env } = context;
  const url = new URL(request.url);

  // Handle CORS Preflight
  if (request.method === "OPTIONS") {
    return new Response(null, {
      status: 204,
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Requested-With",
        "Access-Control-Max-Age": "86400",
      },
    });
  }

  // Determine backend URL from Cloudflare environment variable or active default
  const defaultBackend = "https://ais-pre-5a6afjgobl6oxlzhzq46nb-961275344911.asia-southeast1.run.app";
  const rawBackend = (env && env.BACKEND_URL) ? env.BACKEND_URL : defaultBackend;
  const backendBase = rawBackend.replace(/\/+$/, "");

  // Build target URL
  const targetUrl = new URL(url.pathname + url.search, backendBase);

  // Prepare headers for proxying
  const forwardHeaders = new Headers(request.headers);
  forwardHeaders.set("Host", targetUrl.host);
  forwardHeaders.delete("cf-connecting-ip");
  forwardHeaders.delete("cf-ipcountry");
  forwardHeaders.delete("cf-ray");
  forwardHeaders.delete("cf-visitor");

  const isGetOrHead = request.method === "GET" || request.method === "HEAD";

  const fetchOptions = {
    method: request.method,
    headers: forwardHeaders,
    redirect: "follow",
  };

  if (!isGetOrHead) {
    try {
      fetchOptions.body = await request.arrayBuffer();
    } catch (e) {
      // Body might be empty or already consumed
    }
  }

  try {
    const backendResponse = await fetch(targetUrl.toString(), fetchOptions);

    // Create a new response so we can safely attach CORS headers
    const responseHeaders = new Headers(backendResponse.headers);
    responseHeaders.set("Access-Control-Allow-Origin", "*");
    responseHeaders.set("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
    responseHeaders.set("Access-Control-Allow-Headers", "Content-Type, Authorization, X-Requested-With");

    return new Response(backendResponse.body, {
      status: backendResponse.status,
      statusText: backendResponse.statusText,
      headers: responseHeaders,
    });
  } catch (error) {
    return new Response(
      JSON.stringify({
        status: "error",
        message: "Gagal menghubungkan ke backend melalui Cloudflare Pages Function: " + (error?.message || error),
        target: targetUrl.toString(),
      }),
      {
        status: 502,
        headers: {
          "Content-Type": "application/json",
          "Access-Control-Allow-Origin": "*",
        },
      }
    );
  }
}
