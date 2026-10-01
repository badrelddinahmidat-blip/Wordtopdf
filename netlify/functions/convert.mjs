// دالة Netlify: تحويل Word إلى PDF عبر محرّك ConvertAPI (نفس جودة Microsoft Office)
// المفتاح السري يُحفظ في متغير البيئة CONVERTAPI_SECRET ولا يظهر أبداً للمتصفح.

const API = "https://v2.convertapi.com";
const ALLOWED = ["docx", "doc", "rtf", "odt", "dotx", "dot", "docm"];

const json = (data, status = 200) =>
  new Response(JSON.stringify(data), {
    status,
    headers: { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" },
  });

const authHeaders = () => ({
  Authorization: `Bearer ${process.env.CONVERTAPI_SECRET}`,
  Accept: "application/json",
});

async function readError(res) {
  try {
    const t = await res.text();
    try {
      const j = JSON.parse(t);
      return j.Message || j.message || t;
    } catch {
      return t;
    }
  } catch {
    return `HTTP ${res.status}`;
  }
}

export default async (req) => {
  if (!process.env.CONVERTAPI_SECRET) {
    return json({ error: "لم يتم ضبط المفتاح CONVERTAPI_SECRET في إعدادات Netlify." }, 500);
  }

  const url = new URL(req.url);
  const action = url.searchParams.get("action") || "start";

  try {
    if (req.method === "POST" && action === "upload") {
      const name = decodeURIComponent(req.headers.get("x-file-name") || "document.docx");
      const body = await req.arrayBuffer();
      const res = await fetch(`${API}/upload`, {
        method: "POST",
        headers: {
          ...authHeaders(),
          "Content-Type": "application/octet-stream",
          "Content-Disposition": `inline; filename="${encodeURIComponent(name)}"`,
        },
        body,
      });
      if (!res.ok) return json({ error: await readError(res) }, 502);
      const data = await res.json();
      return json({ fileId: data.FileId });
    }

    if (req.method === "POST" && action === "start") {
      const { fileId, ext, pdfa } = await req.json();
      const from = String(ext || "docx").toLowerCase();
      if (!fileId) return json({ error: "معرّف الملف مفقود." }, 400);
      if (!ALLOWED.includes(from)) return json({ error: "صيغة غير مدعومة." }, 400);

      const params = [
        { Name: "File", FileValue: { Id: fileId } },
        { Name: "StoreFile", Value: true },
      ];
      if (pdfa) params.push({ Name: "Pdfa", Value: true });

      const res = await fetch(`${API}/async/convert/${from}/to/pdf`, {
        method: "POST",
        headers: { ...authHeaders(), "Content-Type": "application/json" },
        body: JSON.stringify({ Parameters: params }),
      });
      if (!res.ok) return json({ error: await readError(res) }, 502);
      const data = await res.json();
      return json({ jobId: data.JobId });
    }

    if (req.method === "GET" && action === "status") {
      const job = url.searchParams.get("job") || "";
      if (!/^[a-z0-9]{8,64}$/i.test(job)) return json({ error: "معرّف مهمة غير صالح." }, 400);

      const res = await fetch(`${API}/async/job/${job}`, { headers: authHeaders() });
      if (res.status === 202) return json({ status: "pending" });
      if (res.status === 404) return json({ status: "error", error: "انتهت صلاحية المهمة، أعد المحاولة." });
      if (!res.ok) return json({ status: "error", error: await readError(res) });

      const data = await res.json();
      const file = data.Files && data.Files[0];
      if (!file || !file.Url) return json({ status: "error", error: "لم يتم إنشاء ملف PDF." });
      return json({ status: "done", url: file.Url, name: file.FileName, size: file.FileSize });
    }

    return json({ error: "طلب غير معروف." }, 404);
  } catch (err) {
    return json({ error: err?.message || "حدث خطأ غير متوقع." }, 500);
  }
};

export const config = { path: "/api/convert" };
