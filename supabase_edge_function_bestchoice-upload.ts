// 베스트 초이스 업로드 API (비밀번호 인증) — 오로지교육 /bestchoice-upload 페이지에서 호출
// actions: sign(업로드 URL 발급) / notify(GitHub Actions 트리거) / status(이력·상태)   — v1.1: 스토리지 키 ASCII 처리
import { createClient } from "npm:@supabase/supabase-js@2";

const BUCKET = "bestchoice-inbox";
const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const json = (b: unknown, s = 200) =>
  new Response(JSON.stringify(b), { status: s, headers: { ...CORS, "Content-Type": "application/json" } });

const sb = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);

async function cfg(): Promise<Record<string, string>> {
  const { data, error } = await sb.from("bestchoice_config").select("key,value");
  if (error) throw error;
  return Object.fromEntries((data ?? []).map((r: any) => [r.key, r.value]));
}
const safeName = (n: string) => n.replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, " ").trim().slice(0, 120);

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  if (req.method !== "POST") return json({ error: "POST only" }, 405);
  let body: any;
  try { body = await req.json(); } catch { return json({ error: "잘못된 요청" }, 400); }

  const c = await cfg();
  if (!body.passcode || body.passcode !== c.passcode) return json({ error: "비밀번호가 올바르지 않습니다." }, 401);

  try {
    if (body.action === "sign") {
      const name = safeName(String(body.filename || "upload.pptx"));
      if (!/\.pptx$/i.test(name)) return json({ error: "PPTX 파일만 업로드할 수 있습니다." }, 400);
      // 스토리지 키는 ASCII만 허용 → 타임스탬프 키로 저장, 원본 파일명은 notify 때 전달
      const stamp = new Date().toISOString().replace(/[-:T.Z]/g, "").slice(0, 17);
      const path = `${stamp}.pptx`;
      const { data, error } = await sb.storage.from(BUCKET).createSignedUploadUrl(path);
      if (error) throw error;
      return json({ path, signedUrl: data.signedUrl, token: data.token });
    }

    if (body.action === "notify") {
      const path = String(body.path || "");
      const name = safeName(String(body.filename || "upload.pptx"));
      if (!path) return json({ error: "path 누락" }, 400);
      const { data: dl, error: e1 } = await sb.storage.from(BUCKET).createSignedUrl(path, 7200);
      if (e1) throw e1;
      const r = await fetch(`https://api.github.com/repos/${c.gh_repo}/dispatches`, {
        method: "POST",
        headers: { Authorization: `Bearer ${c.gh_pat}`, Accept: "application/vnd.github+json", "Content-Type": "application/json", "User-Agent": "bestchoice-upload" },
        body: JSON.stringify({ event_type: "bestchoice_upload", client_payload: { url: dl.signedUrl, name } }),
      });
      const ok = r.status === 204;
      await sb.from("bestchoice_uploads").insert({ file_name: name, storage_path: path, size_bytes: body.size ?? null, dispatched: ok, note: ok ? null : `GitHub ${r.status}: ${(await r.text()).slice(0, 200)}` });
      if (!ok) return json({ error: `변환 요청 실패 (GitHub ${r.status})` }, 502);
      return json({ ok: true });
    }

    if (body.action === "status") {
      const { data: ups } = await sb.from("bestchoice_uploads").select("file_name,size_bytes,uploaded_at,dispatched,note").order("uploaded_at", { ascending: false }).limit(10);
      const gh = await fetch(`https://api.github.com/repos/${c.gh_repo}/actions/runs?per_page=5`, {
        headers: { Authorization: `Bearer ${c.gh_pat}`, Accept: "application/vnd.github+json", "User-Agent": "bestchoice-upload" },
      }).then((r) => r.json()).catch(() => ({}));
      const runs = (gh.workflow_runs ?? []).map((x: any) => ({ status: x.status, conclusion: x.conclusion, created_at: x.created_at, url: x.html_url, title: x.display_title }));
      const latest = await fetch(`https://raw.githubusercontent.com/${c.gh_repo}/main/data/latest.json?ts=${Date.now()}`).then((r) => r.json()).then((d) => ({ month: d.month, key: d.key, generated_at: d.generated_at, source: d.source })).catch(() => null);
      return json({ uploads: ups ?? [], runs, latest });
    }

    return json({ error: "알 수 없는 action" }, 400);
  } catch (e) {
    return json({ error: String((e as any)?.message || e) }, 500);
  }
});
