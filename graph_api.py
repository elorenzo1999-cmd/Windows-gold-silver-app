"""Microsoft Graph API v1.0 client."""

import requests

_BASE = "https://graph.microsoft.com/v1.0"


class GraphError(Exception):
    def __init__(self, status_code: int, message: str):
        self.status_code = status_code
        super().__init__(f"HTTP {status_code}: {message}")


class GraphAPI:
    def __init__(self, token: str):
        self._session = requests.Session()
        self._session.headers.update({
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        })

    def _request(self, method: str, path: str, json=None) -> dict | list | None:
        url = _BASE + path if not path.startswith("http") else path
        response = self._session.request(method, url, json=json)
        if response.status_code == 204:
            return None
        if not response.ok:
            try:
                msg = response.json().get("error", {}).get("message", response.text)
            except Exception:
                msg = response.text
            raise GraphError(response.status_code, msg)
        return response.json()

    def execute(self, method: str, endpoint: str, body: dict | None = None):
        """Generic Graph API call. endpoint should start with /."""
        if not endpoint.startswith("/"):
            endpoint = "/" + endpoint
        return self._request(method.upper(), endpoint, json=body)

    # ── Users ──────────────────────────────────────────────────────────────

    def get_users(self) -> list[dict]:
        """Return all users (pages automatically followed)."""
        users = []
        path = "/users?$select=id,displayName,userPrincipalName,jobTitle,department,accountEnabled&$top=999"
        while path:
            data = self._request("GET", path)
            users.extend(data.get("value", []))
            path = data.get("@odata.nextLink")
        return users

    def get_user(self, user_id: str) -> dict:
        return self._request("GET", f"/users/{user_id}")

    def create_user(self, payload: dict) -> dict:
        return self._request("POST", "/users", json=payload)

    def update_user(self, user_id: str, payload: dict) -> None:
        self._request("PATCH", f"/users/{user_id}", json=payload)

    def delete_user(self, user_id: str) -> None:
        self._request("DELETE", f"/users/{user_id}")

    # ── Licenses ───────────────────────────────────────────────────────────

    def get_subscribed_skus(self) -> list[dict]:
        """Return all license SKUs for the tenant."""
        data = self._request("GET", "/subscribedSkus")
        return data.get("value", [])

    def get_user_licenses(self, user_id: str) -> list[dict]:
        data = self._request("GET", f"/users/{user_id}/licenseDetails")
        return data.get("value", [])

    def assign_license(self, user_id: str, sku_id: str) -> dict:
        payload = {
            "addLicenses": [{"skuId": sku_id, "disabledPlans": []}],
            "removeLicenses": [],
        }
        return self._request("POST", f"/users/{user_id}/assignLicense", json=payload)

    def remove_license(self, user_id: str, sku_id: str) -> dict:
        payload = {
            "addLicenses": [],
            "removeLicenses": [sku_id],
        }
        return self._request("POST", f"/users/{user_id}/assignLicense", json=payload)
