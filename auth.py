"""Microsoft 365 authentication using MSAL interactive browser flow (supports MFA)."""

import msal


# Azure PowerShell well-known public client — no app registration needed
_CLIENT_ID = "1950a258-227b-4e31-a9cf-717495945fc2"
_AUTHORITY = "https://login.microsoftonline.com/organizations"
_SCOPES = ["https://graph.microsoft.com/.default"]


class AuthError(Exception):
    pass


class M365Auth:
    def __init__(self):
        self._app = msal.PublicClientApplication(
            client_id=_CLIENT_ID,
            authority=_AUTHORITY,
        )
        self._token: str | None = None
        self._username: str | None = None

    def login_interactive(self, login_hint: str = "") -> str:
        """Open a browser window for interactive login (supports MFA).

        login_hint pre-fills the username in the browser (optional).
        Returns the access token on success; raises AuthError on failure.
        """
        kwargs = {}
        if login_hint:
            kwargs["login_hint"] = login_hint

        result = self._app.acquire_token_interactive(
            scopes=_SCOPES,
            **kwargs,
        )

        if "access_token" in result:
            self._token = result["access_token"]
            # Use the UPN from the token claims if available, else fall back to hint
            claims = result.get("id_token_claims") or {}
            self._username = claims.get("preferred_username") or login_hint or "unknown"
            return self._token

        error = result.get("error", "unknown_error")
        description = result.get("error_description", "Authentication failed.")
        raise AuthError(f"{error}: {description}")

    def logout(self):
        self._token = None
        self._username = None

    @property
    def token(self) -> str | None:
        return self._token

    @property
    def username(self) -> str | None:
        return self._username

    @property
    def is_authenticated(self) -> bool:
        return self._token is not None
