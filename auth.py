"""Microsoft 365 authentication — OAuth2 auth code flow (supports MFA, isolated session)."""

from urllib.parse import parse_qs, urlparse

import msal


# Azure PowerShell well-known public client — no app registration needed
_CLIENT_ID = "1950a258-227b-4e31-a9cf-717495945fc2"
_AUTHORITY = "https://login.microsoftonline.com/organizations"
_SCOPES = ["https://graph.microsoft.com/.default"]
_REDIRECT_URI = "http://localhost"


class AuthError(Exception):
    pass


class M365Auth:
    def __init__(self):
        self._app = msal.PublicClientApplication(
            client_id=_CLIENT_ID,
            authority=_AUTHORITY,
        )
        self._flow: dict | None = None
        self._token: str | None = None
        self._username: str | None = None

    def start_auth_flow(self, login_hint: str = "") -> str:
        """Initialise the auth code flow and return the Microsoft login URL.

        The caller should navigate an embedded browser to this URL.
        login_hint pre-fills the username field on the login page (optional).
        """
        kwargs = {}
        if login_hint:
            kwargs["login_hint"] = login_hint

        self._flow = self._app.initiate_auth_code_flow(
            scopes=_SCOPES,
            redirect_uri=_REDIRECT_URI,
            **kwargs,
        )
        return self._flow["auth_uri"]

    def complete_auth_flow(self, redirect_url: str) -> str:
        """Complete the flow using the redirect URL (http://localhost?code=…).

        Returns the access token on success; raises AuthError on failure.
        """
        if not self._flow:
            raise AuthError("No active auth flow. Call start_auth_flow() first.")

        parsed = urlparse(redirect_url)
        auth_response = {k: v[0] for k, v in parse_qs(parsed.query).items()}

        result = self._app.acquire_token_by_auth_code_flow(self._flow, auth_response)
        self._flow = None

        if "access_token" in result:
            self._token = result["access_token"]
            claims = result.get("id_token_claims") or {}
            self._username = claims.get("preferred_username", "unknown")
            return self._token

        raise AuthError(result.get("error_description", "Authentication failed."))

    def logout(self):
        self._token = None
        self._username = None
        self._flow = None

    @property
    def token(self) -> str | None:
        return self._token

    @property
    def username(self) -> str | None:
        return self._username

    @property
    def is_authenticated(self) -> bool:
        return self._token is not None
