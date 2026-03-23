#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote

try:
    import msal
except ImportError:  # pragma: no cover - exercised only when dependencies are missing.
    msal = None

try:
    import requests
except ImportError:  # pragma: no cover - exercised only when dependencies are missing.
    requests = None

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
DEFAULT_SCOPES = ("Mail.ReadWrite",)
DEFAULT_SELECT = "id,subject,receivedDateTime,from,sender,bodyPreview,webLink"
RESERVED_SCOPES = frozenset({"openid", "offline_access", "profile"})
DEFAULT_ALLOWED_SENDERS = ("john.doe@example.com", "noreply@example.com")
DEFAULT_SUBJECT_KEYWORDS = ("verification code", "login code")
DEFAULT_TOKEN_CACHE_FILE = ".tokens/msal_cache.json"
VALID_LOG_LEVELS = ("CRITICAL", "ERROR", "WARNING", "INFO", "DEBUG")
APP_REGISTRATION_URL = "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade"


class ConfigError(ValueError):
    pass


class GraphError(RuntimeError):
    pass


@dataclass(frozen=True)
class EnvSettings:
    client_id: str
    tenant_id: str
    scopes: tuple[str, ...]
    allowed_senders: tuple[str, ...]
    allowed_domains: tuple[str, ...]
    subject_keywords: tuple[str, ...]
    body_keywords: tuple[str, ...]
    move_all: bool
    poll_seconds: int
    lookback_hours: int
    scan_limit: int
    dry_run: bool
    token_cache_file: str
    log_level: str


@dataclass(frozen=True)
class Config:
    client_id: str
    tenant_id: str
    scopes: tuple[str, ...]
    poll_seconds: int
    scan_limit: int
    lookback_hours: int
    move_all: bool
    dry_run: bool
    allowed_senders: frozenset[str]
    allowed_domains: frozenset[str]
    subject_keywords: tuple[str, ...]
    body_keywords: tuple[str, ...]
    token_cache_file: Path
    log_level: str

    @classmethod
    def from_env(cls) -> "Config":
        client_id = os.getenv("OUTLOOK_CLIENT_ID", "").strip()
        if not client_id:
            raise ConfigError("OUTLOOK_CLIENT_ID is required.")

        tenant_id = os.getenv("OUTLOOK_TENANT_ID", "common").strip() or "common"
        scopes = normalize_scopes(parse_csv(os.getenv("OUTLOOK_SCOPES")) or list(DEFAULT_SCOPES))
        poll_seconds = parse_int("OUTLOOK_POLL_SECONDS", os.getenv("OUTLOOK_POLL_SECONDS"), 300, minimum=30)
        scan_limit = parse_int("OUTLOOK_SCAN_LIMIT", os.getenv("OUTLOOK_SCAN_LIMIT"), 50, minimum=1)
        lookback_hours = parse_int(
            "OUTLOOK_LOOKBACK_HOURS",
            os.getenv("OUTLOOK_LOOKBACK_HOURS"),
            168,
            minimum=1,
        )
        move_all = parse_bool("OUTLOOK_MOVE_ALL", os.getenv("OUTLOOK_MOVE_ALL"), False)
        dry_run = parse_bool("OUTLOOK_DRY_RUN", os.getenv("OUTLOOK_DRY_RUN"), False)

        allowed_senders = frozenset(normalize_address(value) for value in parse_csv(os.getenv("OUTLOOK_ALLOWED_SENDERS")))
        allowed_domains = frozenset(normalize_domain(value) for value in parse_csv(os.getenv("OUTLOOK_ALLOWED_DOMAINS")))
        subject_keywords = tuple(value.casefold() for value in parse_csv(os.getenv("OUTLOOK_SUBJECT_KEYWORDS")))
        body_keywords = tuple(value.casefold() for value in parse_csv(os.getenv("OUTLOOK_BODY_KEYWORDS")))

        if not move_all and not any((allowed_senders, allowed_domains, subject_keywords, body_keywords)):
            raise ConfigError(
                "Set at least one allowlist/keyword rule, or set OUTLOOK_MOVE_ALL=true to move every message from Junk."
            )

        token_cache_file = Path(os.getenv("OUTLOOK_TOKEN_CACHE_FILE", DEFAULT_TOKEN_CACHE_FILE)).expanduser()
        log_level = os.getenv("OUTLOOK_LOG_LEVEL", "INFO").strip().upper() or "INFO"
        if log_level not in VALID_LOG_LEVELS:
            raise ConfigError("OUTLOOK_LOG_LEVEL must be one of CRITICAL, ERROR, WARNING, INFO, DEBUG.")

        return cls(
            client_id=client_id,
            tenant_id=tenant_id,
            scopes=scopes,
            poll_seconds=poll_seconds,
            scan_limit=scan_limit,
            lookback_hours=lookback_hours,
            move_all=move_all,
            dry_run=dry_run,
            allowed_senders=allowed_senders,
            allowed_domains=allowed_domains,
            subject_keywords=subject_keywords,
            body_keywords=body_keywords,
            token_cache_file=token_cache_file,
            log_level=log_level,
        )


def parse_csv(raw_value: str | None) -> list[str]:
    if not raw_value:
        return []
    return [part.strip() for part in raw_value.split(",") if part.strip()]


def normalize_scopes(scope_values: list[str]) -> tuple[str, ...]:
    normalized: list[str] = []
    for scope in scope_values:
        stripped = scope.strip()
        if not stripped:
            continue
        if stripped.casefold() in RESERVED_SCOPES:
            continue
        normalized.append(stripped)

    if not normalized:
        raise ConfigError("OUTLOOK_SCOPES must include at least one non-reserved Microsoft Graph delegated scope.")

    return tuple(normalized)


def stringify_csv(values: tuple[str, ...] | list[str]) -> str:
    return ",".join(values)


def stringify_bool(value: bool) -> str:
    return "true" if value else "false"


def parse_bool(name: str, raw_value: str | None, default: bool) -> bool:
    if raw_value is None or not raw_value.strip():
        return default

    normalized = raw_value.strip().casefold()
    if normalized in {"1", "true", "yes", "on"}:
        return True
    if normalized in {"0", "false", "no", "off"}:
        return False
    raise ConfigError(f"{name} must be true/false, yes/no, on/off, or 1/0.")


def parse_int(name: str, raw_value: str | None, default: int, minimum: int | None = None) -> int:
    if raw_value is None or not raw_value.strip():
        value = default
    else:
        try:
            value = int(raw_value)
        except ValueError as exc:
            raise ConfigError(f"{name} must be an integer.") from exc

    if minimum is not None and value < minimum:
        raise ConfigError(f"{name} must be >= {minimum}.")
    return value


def normalize_address(value: str) -> str:
    return value.strip().casefold()


def normalize_domain(value: str) -> str:
    return value.strip().casefold().lstrip("@")


def parse_iso_datetime(raw_value: str | None) -> datetime | None:
    if not raw_value:
        return None
    try:
        return datetime.fromisoformat(raw_value.replace("Z", "+00:00"))
    except ValueError:
        return None


def parse_dotenv_file(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if not path.exists():
        return values

    for line_number, raw_line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue

        if line.startswith("export "):
            line = line[7:].strip()

        if "=" not in line:
            raise ConfigError(f"Invalid line in {path} at {line_number}: expected KEY=VALUE.")

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip()
        if not key:
            raise ConfigError(f"Invalid line in {path} at {line_number}: empty key.")

        if value[:1] == value[-1:] and value[:1] in {"'", '"'}:
            value = value[1:-1]

        values[key] = value

    return values


def infer_account_type(tenant_id: str) -> str:
    return "personal" if tenant_id.strip().casefold() == "consumers" else "work"


def parse_existing_bool(raw_value: str | None, default: bool) -> bool:
    try:
        return parse_bool("value", raw_value, default)
    except ConfigError:
        return default


def parse_existing_int(raw_value: str | None, default: int, minimum: int | None = None) -> int:
    try:
        return parse_int("value", raw_value, default, minimum=minimum)
    except ConfigError:
        return default


def build_wizard_defaults(existing_values: dict[str, str]) -> EnvSettings:
    has_existing_matchers = any(
        existing_values.get(key, "").strip()
        for key in (
            "OUTLOOK_ALLOWED_SENDERS",
            "OUTLOOK_ALLOWED_DOMAINS",
            "OUTLOOK_SUBJECT_KEYWORDS",
            "OUTLOOK_BODY_KEYWORDS",
        )
    ) or parse_existing_bool(existing_values.get("OUTLOOK_MOVE_ALL"), False)

    default_senders = tuple(parse_csv(existing_values.get("OUTLOOK_ALLOWED_SENDERS")))
    default_subject_keywords = tuple(parse_csv(existing_values.get("OUTLOOK_SUBJECT_KEYWORDS")))

    if not has_existing_matchers:
        default_senders = DEFAULT_ALLOWED_SENDERS
        default_subject_keywords = DEFAULT_SUBJECT_KEYWORDS

    try:
        scopes = normalize_scopes(parse_csv(existing_values.get("OUTLOOK_SCOPES")) or list(DEFAULT_SCOPES))
    except ConfigError:
        scopes = DEFAULT_SCOPES

    return EnvSettings(
        client_id=existing_values.get("OUTLOOK_CLIENT_ID", "").strip(),
        tenant_id=existing_values.get("OUTLOOK_TENANT_ID", "consumers").strip() or "consumers",
        scopes=scopes,
        allowed_senders=default_senders,
        allowed_domains=tuple(parse_csv(existing_values.get("OUTLOOK_ALLOWED_DOMAINS"))),
        subject_keywords=default_subject_keywords,
        body_keywords=tuple(parse_csv(existing_values.get("OUTLOOK_BODY_KEYWORDS"))),
        move_all=parse_existing_bool(existing_values.get("OUTLOOK_MOVE_ALL"), False),
        poll_seconds=parse_existing_int(existing_values.get("OUTLOOK_POLL_SECONDS"), 300, minimum=30),
        lookback_hours=parse_existing_int(existing_values.get("OUTLOOK_LOOKBACK_HOURS"), 168, minimum=1),
        scan_limit=parse_existing_int(existing_values.get("OUTLOOK_SCAN_LIMIT"), 50, minimum=1),
        dry_run=parse_existing_bool(existing_values.get("OUTLOOK_DRY_RUN"), False),
        token_cache_file=existing_values.get("OUTLOOK_TOKEN_CACHE_FILE", DEFAULT_TOKEN_CACHE_FILE).strip()
        or DEFAULT_TOKEN_CACHE_FILE,
        log_level=(existing_values.get("OUTLOOK_LOG_LEVEL", "INFO").strip().upper() or "INFO"),
    )


def render_env_file(settings: EnvSettings) -> str:
    return "\n".join(
        [
            "# Generated by outlook_junk_mover.py --configure",
            f"OUTLOOK_CLIENT_ID={settings.client_id}",
            f"OUTLOOK_TENANT_ID={settings.tenant_id}",
            f"OUTLOOK_SCOPES={stringify_csv(settings.scopes)}",
            "",
            "# Message matching rules",
            f"OUTLOOK_ALLOWED_SENDERS={stringify_csv(settings.allowed_senders)}",
            f"OUTLOOK_ALLOWED_DOMAINS={stringify_csv(settings.allowed_domains)}",
            f"OUTLOOK_SUBJECT_KEYWORDS={stringify_csv(settings.subject_keywords)}",
            f"OUTLOOK_BODY_KEYWORDS={stringify_csv(settings.body_keywords)}",
            f"OUTLOOK_MOVE_ALL={stringify_bool(settings.move_all)}",
            "",
            "# Polling",
            f"OUTLOOK_POLL_SECONDS={settings.poll_seconds}",
            f"OUTLOOK_LOOKBACK_HOURS={settings.lookback_hours}",
            f"OUTLOOK_SCAN_LIMIT={settings.scan_limit}",
            f"OUTLOOK_DRY_RUN={stringify_bool(settings.dry_run)}",
            "",
            "# Runtime",
            f"OUTLOOK_TOKEN_CACHE_FILE={settings.token_cache_file}",
            f"OUTLOOK_LOG_LEVEL={settings.log_level}",
            "",
        ]
    )


def prompt_text(
    label: str,
    *,
    default: str | None = None,
    required: bool = False,
    allow_clear: bool = False,
    validator: Any | None = None,
) -> str:
    while True:
        suffix = f" [{default}]" if default else ""
        raw_value = input(f"{label}{suffix}: ").strip()

        if raw_value == "-" and allow_clear:
            return ""
        if not raw_value:
            if default is not None:
                value = default
            elif required:
                print("A value is required.")
                continue
            else:
                return ""
        else:
            value = raw_value

        if validator is None:
            return value

        try:
            validator(value)
        except ConfigError as exc:
            print(exc)
            continue
        return value


def prompt_choice(label: str, options: list[tuple[str, str]], default: str) -> str:
    display = ", ".join(f"{index}. {description}" for index, (_, description) in enumerate(options, start=1))
    mapping = {str(index): value for index, (value, _) in enumerate(options, start=1)}
    descriptions = {value: description for value, description in options}

    while True:
        raw_value = input(f"{label} [{descriptions[default]}] ({display}): ").strip()
        if not raw_value:
            return default

        normalized = raw_value.casefold()
        if raw_value in mapping:
            return mapping[raw_value]
        for value, description in options:
            if normalized in {value.casefold(), description.casefold()}:
                return value

        print("Choose one of the listed options.")


def prompt_bool(label: str, default: bool) -> bool:
    suffix = "Y/n" if default else "y/N"
    while True:
        raw_value = input(f"{label} [{suffix}]: ").strip()
        if not raw_value:
            return default

        normalized = raw_value.casefold()
        if normalized in {"y", "yes", "true", "1", "on"}:
            return True
        if normalized in {"n", "no", "false", "0", "off"}:
            return False

        print("Enter yes or no.")


def prompt_csv(label: str, default: tuple[str, ...]) -> tuple[str, ...]:
    default_display = stringify_csv(default)
    raw_value = prompt_text(label, default=default_display, allow_clear=True)
    return tuple(parse_csv(raw_value))


def prompt_int_value(label: str, default: int, minimum: int) -> int:
    while True:
        raw_value = prompt_text(label, default=str(default))
        try:
            return parse_int(label, raw_value, default, minimum=minimum)
        except ConfigError as exc:
            print(exc)


def prompt_log_level(default: str) -> str:
    while True:
        raw_value = prompt_text("Log level", default=default)
        normalized = raw_value.upper()
        if normalized in VALID_LOG_LEVELS:
            return normalized
        print(f"Choose one of: {', '.join(VALID_LOG_LEVELS)}.")


def validate_client_id(value: str) -> None:
    if not value.strip():
        raise ConfigError("OUTLOOK_CLIENT_ID is required.")


def print_registration_guide() -> None:
    print("Register or update your Microsoft Entra app before continuing:")
    print(f"1. Open {APP_REGISTRATION_URL}")
    print("2. Create an app registration, or open an existing one.")
    print("3. Copy the Application (client) ID from the app Overview page.")
    print("4. In Authentication, enable public client flows.")
    print("5. In API permissions, add Microsoft Graph delegated permission Mail.ReadWrite.")
    print("")


def print_auth_section_guide(account_type: str) -> None:
    print("Authentication settings")
    if account_type == "personal":
        print("Use an app that supports personal Microsoft accounts.")
        print(
            "In Supported account types, choose either "
            "`Accounts in any organizational directory and personal Microsoft accounts` "
            "or `Personal Microsoft accounts only`."
        )
        print("This wizard will write OUTLOOK_TENANT_ID=consumers.")
    else:
        print("Use a tenant-specific value for OUTLOOK_TENANT_ID.")
        print("That can be your tenant GUID or a verified domain like contoso.onmicrosoft.com.")
    print("")


def print_matching_section_guide() -> None:
    print("Message matching rules")
    print("Messages are moved out of Junk only when they match at least one rule, unless you enable move-all.")
    print("Allowed sender emails are the safest option; domains and keywords are broader matches.")
    print("")


def print_polling_section_guide() -> None:
    print("Polling settings")
    print("Poll interval controls how often Junk is scanned.")
    print("Lookback limits how old a message can be and still be moved.")
    print("Scan limit caps how many recent Junk messages are checked each run.")
    print("")


def run_onboarding_wizard(config_path: Path) -> int:
    if not sys.stdin.isatty():
        print("The configuration wizard requires an interactive terminal.", file=sys.stderr)
        return 2

    try:
        existing_values = parse_dotenv_file(config_path)
    except ConfigError as exc:
        print(f"Could not parse {config_path}: {exc}", file=sys.stderr)
        return 2

    defaults = build_wizard_defaults(existing_values)

    print("Outlook Junk Mail Mover configuration wizard")
    print("Press Enter to keep the value in brackets. Enter '-' to clear optional CSV fields.")
    print("This wizard will collect app auth settings, message matching rules, and polling settings, then write the config file.")
    print("")
    print_registration_guide()

    account_type = prompt_choice(
        "Mailbox type",
        [
            ("personal", "Personal Outlook.com/Hotmail/Live"),
            ("work", "Work or school Microsoft 365"),
        ],
        default=infer_account_type(defaults.tenant_id),
    )
    print("")
    print_auth_section_guide(account_type)
    client_id = prompt_text(
        "Microsoft Entra application (client) ID",
        default=defaults.client_id or None,
        required=True,
        validator=validate_client_id,
    )

    if account_type == "personal":
        tenant_id = "consumers"
        print("Using OUTLOOK_TENANT_ID=consumers for personal Microsoft accounts.")
        print("Your app registration must support personal Microsoft accounts.")
    else:
        tenant_default = defaults.tenant_id if defaults.tenant_id not in {"common", "consumers"} else ""
        tenant_id = prompt_text(
            "Tenant ID or verified domain",
            default=tenant_default or None,
            required=True,
        )

    print("")
    print_matching_section_guide()
    while True:
        allowed_senders = prompt_csv("Allowed sender emails (comma-separated)", defaults.allowed_senders)
        allowed_domains = prompt_csv("Allowed sender domains (comma-separated)", defaults.allowed_domains)
        subject_keywords = prompt_csv("Subject keywords (comma-separated)", defaults.subject_keywords)
        body_keywords = prompt_csv("Body keywords (comma-separated)", defaults.body_keywords)
        move_all = prompt_bool("Move all recent messages from Junk", defaults.move_all)

        if move_all or any((allowed_senders, allowed_domains, subject_keywords, body_keywords)):
            break

        print("Configure at least one allowlist/keyword rule, or enable move-all.")

    print("")
    print_polling_section_guide()
    poll_seconds = prompt_int_value("Poll interval in seconds", defaults.poll_seconds, minimum=30)
    lookback_hours = prompt_int_value("Lookback window in hours", defaults.lookback_hours, minimum=1)
    scan_limit = prompt_int_value("Max junk messages to scan each run", defaults.scan_limit, minimum=1)
    dry_run = prompt_bool("Dry run only (log matches without moving)", defaults.dry_run)
    token_cache_file = prompt_text("Token cache file", default=defaults.token_cache_file)
    log_level = prompt_log_level(defaults.log_level if defaults.log_level in VALID_LOG_LEVELS else "INFO")

    settings = EnvSettings(
        client_id=client_id,
        tenant_id=tenant_id,
        scopes=defaults.scopes,
        allowed_senders=allowed_senders,
        allowed_domains=allowed_domains,
        subject_keywords=subject_keywords,
        body_keywords=body_keywords,
        move_all=move_all,
        poll_seconds=poll_seconds,
        lookback_hours=lookback_hours,
        scan_limit=scan_limit,
        dry_run=dry_run,
        token_cache_file=token_cache_file,
        log_level=log_level,
    )

    print("")
    print(f"Config file: {config_path}")
    print(f"Client ID: {settings.client_id}")
    print(f"Tenant: {settings.tenant_id}")
    print(f"Allowed senders: {stringify_csv(settings.allowed_senders) or '<none>'}")
    print(f"Allowed domains: {stringify_csv(settings.allowed_domains) or '<none>'}")
    print(f"Subject keywords: {stringify_csv(settings.subject_keywords) or '<none>'}")
    print(f"Body keywords: {stringify_csv(settings.body_keywords) or '<none>'}")
    print(f"Move all: {stringify_bool(settings.move_all)}")
    print(f"Poll seconds: {settings.poll_seconds}")
    print(f"Lookback hours: {settings.lookback_hours}")
    print(f"Scan limit: {settings.scan_limit}")
    print(f"Dry run: {stringify_bool(settings.dry_run)}")
    print(f"Token cache file: {settings.token_cache_file}")
    print(f"Log level: {settings.log_level}")
    print("")

    if not prompt_bool(f"Write these settings to {config_path}", True):
        print("Wizard cancelled; .env was not updated.")
        return 1

    config_path.parent.mkdir(parents=True, exist_ok=True)
    config_path.write_text(render_env_file(settings), encoding="utf-8")

    print(f"Wrote {config_path}")
    print("Next: run `uv run outlook_junk_mover.py --once` to test authentication and one scan.")
    return 0


def ensure_runtime_dependencies() -> None:
    missing: list[str] = []
    if msal is None:
        missing.append("msal")
    if requests is None:
        missing.append("requests")
    if missing:
        joined = ", ".join(missing)
        raise RuntimeError(f"Missing required packages: {joined}. Install them with `pip install -r requirements.txt`.")


def format_device_flow_error(flow: dict[str, Any], tenant_id: str) -> str:
    error = str(flow.get("error") or "unknown_error")
    description = str(flow.get("error_description") or "No error description returned.")
    details = [f"{error}: {description}"]

    if "AADSTS50059" in description or "AADSTS90133" in description:
        if tenant_id == "common":
            details.append(
                "If this is a personal Outlook.com/Hotmail/Live account, set OUTLOOK_TENANT_ID=consumers."
            )
            details.append(
                "If this is a work/school account, set OUTLOOK_TENANT_ID to your tenant GUID or verified domain "
                "instead of `common`."
            )
        else:
            details.append(
                "For work/school accounts, use a tenant-specific authority instead of "
                "`common`/`organizations`/`consumers`."
            )
            details.append(
                "Set OUTLOOK_TENANT_ID to your tenant GUID or verified domain, for example "
                "`contoso.onmicrosoft.com`."
            )
    elif "AADSTS700016" in description and tenant_id in {"common", "consumers"}:
        details.append(
            "This app registration is not available for Microsoft personal accounts."
        )
        details.append(
            "For Outlook.com/Hotmail/Live, change the app registration supported account types to include "
            "personal Microsoft accounts, or create a new app that does."
        )

    return " ".join(details)


def load_dotenv(path: Path) -> None:
    for key, value in parse_dotenv_file(path).items():
        os.environ.setdefault(key, value)


def get_message_addresses(message: dict[str, Any]) -> set[str]:
    addresses: set[str] = set()
    for field in ("from", "sender"):
        container = message.get(field) or {}
        email = (container.get("emailAddress") or {}).get("address")
        if email:
            addresses.add(normalize_address(email))
    return addresses


def matches_message(message: dict[str, Any], config: Config) -> bool:
    if config.move_all:
        return True

    addresses = get_message_addresses(message)
    if addresses & config.allowed_senders:
        return True

    for address in addresses:
        if "@" in address and address.rsplit("@", 1)[1] in config.allowed_domains:
            return True

    subject = str(message.get("subject") or "").casefold()
    body_preview = str(message.get("bodyPreview") or "").casefold()

    if any(keyword in subject for keyword in config.subject_keywords):
        return True

    if any(keyword in body_preview for keyword in config.body_keywords):
        return True

    return False


def is_recent_enough(message: dict[str, Any], lookback_hours: int) -> bool:
    received_at = parse_iso_datetime(message.get("receivedDateTime"))
    if received_at is None:
        return False
    cutoff = datetime.now(timezone.utc) - timedelta(hours=lookback_hours)
    return received_at >= cutoff


def describe_message(message: dict[str, Any]) -> str:
    addresses = sorted(get_message_addresses(message))
    sender = addresses[0] if addresses else "<unknown>"
    subject = str(message.get("subject") or "").strip() or "<no subject>"
    received = str(message.get("receivedDateTime") or "<unknown time>")
    return f"{subject} | {sender} | {received}"


class OutlookGraphClient:
    def __init__(self, config: Config) -> None:
        ensure_runtime_dependencies()
        self.config = config
        self.cache = msal.SerializableTokenCache()
        self._load_cache()
        self.app = msal.PublicClientApplication(
            client_id=config.client_id,
            authority=f"https://login.microsoftonline.com/{config.tenant_id}",
            token_cache=self.cache,
        )
        self.session = requests.Session()

    def _load_cache(self) -> None:
        if not self.config.token_cache_file.exists():
            return
        self.cache.deserialize(self.config.token_cache_file.read_text(encoding="utf-8"))

    def _save_cache(self) -> None:
        if not self.cache.has_state_changed:
            return
        self.config.token_cache_file.parent.mkdir(parents=True, exist_ok=True)
        self.config.token_cache_file.write_text(self.cache.serialize(), encoding="utf-8")

    def get_access_token(self) -> str:
        accounts = self.app.get_accounts()
        result: dict[str, Any] | None = None
        if accounts:
            result = self.app.acquire_token_silent(scopes=list(self.config.scopes), account=accounts[0])
            self._save_cache()

        if not result or "access_token" not in result:
            flow = self.app.initiate_device_flow(scopes=list(self.config.scopes))
            if "user_code" not in flow:
                raise GraphError(f"Failed to start device-code authentication flow. {format_device_flow_error(flow, self.config.tenant_id)}")

            print(flow["message"], file=sys.stderr)
            result = self.app.acquire_token_by_device_flow(flow)
            self._save_cache()

        if not result or "access_token" not in result:
            error = result.get("error") if result else "unknown_error"
            description = result.get("error_description") if result else "No error description returned."
            raise GraphError(f"Authentication failed: {error}: {description}")

        return result["access_token"]

    def request(
        self,
        method: str,
        url_or_path: str,
        *,
        params: dict[str, Any] | None = None,
        json_body: dict[str, Any] | None = None,
        expected_statuses: tuple[int, ...] = (200,),
    ) -> Any:
        url = url_or_path if url_or_path.startswith("http") else f"{GRAPH_ROOT}{url_or_path}"
        last_error: GraphError | None = None

        for attempt in range(3):
            token = self.get_access_token()
            headers = {
                "Authorization": f"Bearer {token}",
                "Accept": "application/json",
            }

            response = self.session.request(
                method=method,
                url=url,
                params=params,
                json=json_body,
                headers=headers,
                timeout=30,
            )

            if response.status_code in expected_statuses:
                if response.status_code == 204 or not response.content:
                    return None
                return response.json()

            if response.status_code in {429, 500, 502, 503, 504}:
                retry_after = response.headers.get("Retry-After")
                wait_seconds = int(retry_after) if retry_after and retry_after.isdigit() else 2 ** attempt
                last_error = GraphError(f"Transient Graph API error {response.status_code}: {response.text}")
                logging.warning("Graph API returned %s, retrying in %ss.", response.status_code, wait_seconds)
                time.sleep(wait_seconds)
                continue

            error_message = extract_error_message(response)
            raise GraphError(f"Graph API {method} {url} failed with {response.status_code}: {error_message}")

        raise last_error or GraphError("Graph API request failed after retries.")

    def list_junk_messages(self, limit: int) -> list[dict[str, Any]]:
        results: list[dict[str, Any]] = []
        next_url: str | None = "/me/mailFolders/junkemail/messages"
        params: dict[str, Any] | None = {
            "$select": DEFAULT_SELECT,
            "$orderby": "receivedDateTime desc",
            "$top": min(limit, 100),
        }

        while next_url and len(results) < limit:
            payload = self.request("GET", next_url, params=params, expected_statuses=(200,))
            results.extend(payload.get("value", []))
            next_url = payload.get("@odata.nextLink")
            params = None

        return results[:limit]

    def move_message_to_inbox(self, message_id: str) -> dict[str, Any]:
        encoded_message_id = quote(message_id, safe="")
        return self.request(
            "POST",
            f"/me/messages/{encoded_message_id}/move",
            json_body={"destinationId": "inbox"},
            expected_statuses=(201,),
        )


def extract_error_message(response: requests.Response) -> str:
    try:
        payload = response.json()
    except json.JSONDecodeError:
        return response.text.strip() or "<no response body>"

    error = payload.get("error") or {}
    code = error.get("code")
    message = error.get("message")
    if code and message:
        return f"{code}: {message}"
    if message:
        return str(message)
    return response.text.strip() or "<no response body>"


def run_once(client: OutlookGraphClient, config: Config) -> tuple[int, int]:
    messages = client.list_junk_messages(config.scan_limit)
    moved = 0
    matched = 0

    for message in messages:
        if not is_recent_enough(message, config.lookback_hours):
            continue

        if not matches_message(message, config):
            logging.debug("Skipping unmatched junk message: %s", describe_message(message))
            continue

        matched += 1
        if config.dry_run:
            logging.info("Dry run: would move %s", describe_message(message))
            continue

        message_id = message.get("id")
        if not message_id:
            logging.warning("Skipping matched message without an id: %s", describe_message(message))
            continue

        moved_message = client.move_message_to_inbox(message_id)
        logging.info("Moved %s", describe_message(moved_message))
        moved += 1

    logging.info("Scan complete. Matched %s message(s); moved %s.", matched, moved)
    return matched, moved


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Poll Outlook Junk Email and move allowlisted messages back into Inbox."
    )
    parser.add_argument(
        "--config",
        default=".env",
        help="Path to a .env-style config file. Defaults to ./.env",
    )
    parser.add_argument(
        "--configure",
        action="store_true",
        help="Run the interactive configuration wizard and update the config file, then exit.",
    )
    parser.add_argument(
        "--wizard",
        action="store_true",
        dest="configure",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--once",
        action="store_true",
        help="Run a single scan and exit instead of polling forever.",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    config_path = Path(args.config).expanduser()

    if args.configure:
        return run_onboarding_wizard(config_path)

    try:
        load_dotenv(config_path)
        config = Config.from_env()
    except ConfigError as exc:
        print(f"Configuration error: {exc}", file=sys.stderr)
        print(
            f"Run `uv run outlook_junk_mover.py --configure --config {config_path}` to create or update the config.",
            file=sys.stderr,
        )
        return 2

    logging.basicConfig(
        level=getattr(logging, config.log_level),
        format="%(asctime)s %(levelname)s %(message)s",
    )

    try:
        client = OutlookGraphClient(config)
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        return 2

    try:
        while True:
            try:
                run_once(client, config)
            except GraphError:
                logging.exception("Scan failed due to a Graph API error.")
            except requests.RequestException:
                logging.exception("Scan failed due to a network error.")

            if args.once:
                return 0

            time.sleep(config.poll_seconds)
    except KeyboardInterrupt:
        logging.info("Stopped by user.")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
