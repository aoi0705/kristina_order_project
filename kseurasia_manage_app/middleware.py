import re
from django.http import HttpResponse, HttpResponseRedirect
from django.conf import settings
from django.core import signing

class SitePassMiddleware:
    """
    サイト共通パスコードゲート。
    - 有効時、署名付きクッキー 'site_gate' が無い/不正ならパスコード入力画面を返す
    - 正答時に 'site_gate' を付与して以降通過
    """
    COOKIE_NAME = "site_gate"
    FORM_PATH = "/_gate"
    LOGOUT_PATH = "/_gate/logout"

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        # 無効なら素通り
        if not getattr(settings, "SITE_GATE_ENABLED", False):
            return self.get_response(request)

        path = request.path or "/"

        # 許可パス（静的・ファビコン・ゲート自身など）
        allow_patterns = getattr(settings, "SITE_GATE_ALLOWLIST", [
            r"^/static/",
            r"^/favicon\.ico$",
            r"^/healthz$",
            r"^%s$" % re.escape(self.FORM_PATH),
            r"^%s$" % re.escape(self.LOGOUT_PATH),
        ])
        if any(re.match(p, path) for p in allow_patterns):
            # ゲートのPOST/GET処理は下で別ハンドリング
            if path.startswith(self.FORM_PATH):
                return self._handle_form(request)
            if path == self.LOGOUT_PATH:
                return self._logout_and_redirect(request)
            return self.get_response(request)

        # 署名付きクッキー確認
        token = request.COOKIES.get(self.COOKIE_NAME)
        if token:
            try:
                data = signing.loads(token, max_age=getattr(settings, "SITE_GATE_MAX_AGE", 60*60*12))
                if data.get("ok"):
                    return self.get_response(request)
            except signing.BadSignature:
                pass  # フォールバックでフォーム表示へ

        # 未認証 → 入力フォームへ（元URLを保持）
        return HttpResponseRedirect(f"{self.FORM_PATH}?next={request.get_full_path()}")

    # --- helpers ---

    def _handle_form(self, request):
        # 正答ならクッキー付与して next へ
        if request.method == "POST":
            pwd = (request.POST.get("passcode") or "").strip()
            if pwd and pwd == getattr(settings, "SITE_GATE_PASSWORD", ""):
                resp = HttpResponseRedirect(request.GET.get("next") or "/")
                value = signing.dumps({"ok": True})
                secure = request.is_secure()
                resp.set_cookie(
                    self.COOKIE_NAME, value,
                    max_age=getattr(settings, "SITE_GATE_MAX_AGE", 60*60*12),
                    httponly=True, secure=secure, samesite="Lax",
                )
                return resp
            # 不正→フォーム再表示（エラー）
            return self._render_form(error=True, next_url=request.GET.get("next") or "/")

        # GET：フォーム表示
        return self._render_form(next_url=request.GET.get("next") or "/")

    def _logout_and_redirect(self, request):
        resp = HttpResponseRedirect("/")
        resp.delete_cookie(self.COOKIE_NAME)
        return resp

    def _render_form(self, error=False, next_url="/"):
        html = f"""
<!doctype html><meta charset="utf-8">
<title>Access Gate</title>
<style>
  body{{font-family:system-ui,-apple-system,Segoe UI,Roboto,"Noto Sans JP",sans-serif;background:#f6f7fb;display:grid;place-items:center;height:100vh;margin:0}}
  .card{{background:#fff;border:1px solid #e5e7eb;border-radius:12px;box-shadow:0 4px 18px rgba(0,0,0,.06);padding:24px;min-width:320px}}
  .title{{font-weight:700;color:#0c4a6e;margin-bottom:8px}}
  .err{{color:#b91c1c;font-size:12px;margin:6px 0}}
  input[type=password]{{width:100%;height:40px;border:1px solid #e5e7eb;border-radius:8px;padding:8px 10px}}
  button{{margin-top:12px;width:100%;height:40px;border-radius:8px;border:0;background:#0ea5e9;color:#fff;font-weight:600;cursor:pointer}}
  .hint{{font-size:12px;color:#6b7280;margin-top:8px;text-align:center}}
</style>
<div class="card">
  <div class="title">Access Gate</div>
  {"<div class='err'>パスコードが違います。</div>" if error else ""}
  <form method="post" action="{self.FORM_PATH}?next={next_url}">
    <input type="password" name="passcode" placeholder="パスコード" autofocus required>
    <button type="submit">ログイン</button>
  </form>
</div>"""
        return HttpResponse(html)