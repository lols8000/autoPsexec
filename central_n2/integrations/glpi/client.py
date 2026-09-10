from __future__ import annotations

import json
import urllib.error
import urllib.parse
import urllib.request
from typing import Any


class GLPIError(RuntimeError):
    pass


class GLPIClient:
    def __init__(
        self,
        base_url: str,
        app_token: str,
        user_token: str,
        *,
        timeout: int = 20,
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self.app_token = app_token
        self.user_token = user_token
        self.timeout = max(1, int(timeout))
        self.session_token: str | None = None

    def _headers(self, *, session: bool = True) -> dict[str, str]:
        headers = {
            "Content-Type": "application/json",
            "App-Token": self.app_token,
        }
        if session and self.session_token:
            headers["Session-Token"] = self.session_token
        elif not session:
            headers["Authorization"] = f"user_token {self.user_token}"
        return headers

    def _request(
        self,
        method: str,
        path: str,
        *,
        payload: dict[str, Any] | None = None,
        session: bool = True,
    ) -> Any:
        if not self.base_url:
            raise GLPIError("glpi_api.base_url não configurado.")

        url = f"{self.base_url}/{path.lstrip('/')}"
        data = (
            None
            if payload is None
            else json.dumps(payload).encode("utf-8")
        )
        request = urllib.request.Request(
            url,
            data=data,
            method=method,
            headers=self._headers(session=session),
        )

        try:
            with urllib.request.urlopen(
                request,
                timeout=self.timeout,
            ) as response:
                raw = response.read().decode("utf-8")
        except urllib.error.HTTPError as exc:
            try:
                detail = exc.read().decode(
                    "utf-8",
                    errors="replace",
                )
            except Exception:
                detail = ""
            suffix = f": {detail[:500]}" if detail else ""
            raise GLPIError(
                f"GLPI HTTP {exc.code}{suffix}"
            ) from exc
        except urllib.error.URLError as exc:
            raise GLPIError(
                f"Falha de conexão com GLPI: {exc.reason}"
            ) from exc
        except OSError as exc:
            raise GLPIError(
                f"Falha de comunicação com GLPI: {exc}"
            ) from exc

        if not raw:
            return None
        try:
            return json.loads(raw)
        except json.JSONDecodeError as exc:
            raise GLPIError(
                "GLPI retornou uma resposta que não é JSON válido."
            ) from exc

    def init_session(self) -> str:
        data = self._request(
            "GET",
            "initSession",
            session=False,
        )
        token = (
            data.get("session_token")
            if isinstance(data, dict)
            else None
        )
        if not token:
            raise GLPIError("GLPI não retornou session_token.")
        self.session_token = str(token)
        return self.session_token

    def kill_session(self) -> None:
        if not self.session_token:
            return
        try:
            self._request("GET", "killSession")
        finally:
            self.session_token = None

    def search_computer(self, name: str) -> Any:
        if not self.session_token:
            self.init_session()
        query = urllib.parse.urlencode(
            {
                "criteria[0][field]": 1,
                "criteria[0][searchtype]": "contains",
                "criteria[0][value]": name,
            }
        )
        return self._request(
            "GET",
            f"search/Computer?{query}",
        )

    def add_ticket_followup(
        self,
        ticket_id: int,
        content: str,
    ) -> Any:
        if not self.session_token:
            self.init_session()
        return self._request(
            "POST",
            "TicketFollowup",
            payload={
                "input": {
                    "items_id": int(ticket_id),
                    "content": content,
                }
            },
        )
