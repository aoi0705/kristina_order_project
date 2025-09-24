import base64, re
from django.http import HttpResponse
from django.conf import settings

class BasicAuthMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        # 無効なら素通り
        if not getattr(settings, "BASIC_AUTH_ENABLED", False):
            return self.get_response(request)

        # 静的ファイル等は素通り（必要最小限）
        allow_patterns = getattr(settings, "BASIC_AUTH_PATH_ALLOWLIST", [
            r"^/static/", r"^/favicon\.ico$",
        ])
        path = request.path or "/"
        if any(re.match(p, path) for p in allow_patterns):
            return self.get_response(request)

        auth = request.META.get("HTTP_AUTHORIZATION", "")
        if auth.startswith("Basic "):
            try:
                decoded = base64.b64decode(auth.split(" ", 1)[1]).decode("utf-8")
                username, _, password = decoded.partition(":")
            except Exception:
                return self._unauthorized()
            if (username == getattr(settings, "BASIC_AUTH_USERNAME", "")
                and password == getattr(settings, "BASIC_AUTH_PASSWORD", "")):
                return self.get_response(request)

        return self._unauthorized()

    @staticmethod
    def _unauthorized():
        resp = HttpResponse("Authorization Required", status=401)
        resp["WWW-Authenticate"] = 'Basic realm="Restricted"'
        return resp