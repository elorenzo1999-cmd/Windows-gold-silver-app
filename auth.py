"""Microsoft 365 authentication using MSAL ROPC flow."""

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

    def login(self, username: str, password: str) -> str:
        """Authenticate with username/password. Returns access token.

        Raises AuthError on failure.
        Note: Accounts with MFA enabled cannot use this flow.
        """
        result = self._app.acquire_token_by_username_password(
            username=username,
            password=password,
            scopes=_SCOPES,
        )
        if "access_token" in result:
            self._token = result["access_token"]
            self._username = username
            return self._token

        error = result.get("error", "unknown_error")
        description = result.get("error_description", "Authentication failed.")
        if "AADSTS50076" in description or "MFA" in description.upper():
            raise AuthError(
                "This account requires Multi-Factor Authentication (MFA). "
                "Username/password login is not supported for MFA-enabled accounts."
            )
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
