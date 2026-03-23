import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path
from tempfile import TemporaryDirectory

from outlook_junk_mover import (
    Config,
    EnvSettings,
    build_wizard_defaults,
    is_recent_enough,
    matches_message,
    normalize_scopes,
    parse_dotenv_file,
    render_env_file,
)


def build_config(**overrides):
    defaults = dict(
        client_id="client-id",
        tenant_id="common",
        scopes=("Mail.ReadWrite", "offline_access"),
        poll_seconds=300,
        scan_limit=50,
        lookback_hours=168,
        move_all=False,
        dry_run=False,
        allowed_senders=frozenset(),
        allowed_domains=frozenset(),
        subject_keywords=(),
        body_keywords=(),
        token_cache_file=Path(".tokens/msal_cache.json"),
        log_level="INFO",
    )
    defaults.update(overrides)
    return Config(**defaults)


class MatchesMessageTests(unittest.TestCase):
    def test_normalize_scopes_strips_reserved_scopes(self):
        self.assertEqual(normalize_scopes(["Mail.ReadWrite", "offline_access", "openid"]), ("Mail.ReadWrite",))

    def test_matches_exact_sender_case_insensitively(self):
        config = build_config(allowed_senders=frozenset({"john.doe@example.com"}))
        message = {
            "from": {"emailAddress": {"address": "John.Doe@Www.Example.com"}},
            "subject": "Your code",
            "bodyPreview": "",
        }
        self.assertTrue(matches_message(message, config))

    def test_matches_domain(self):
        config = build_config(allowed_domains=frozenset({"example.com"}))
        message = {
            "sender": {"emailAddress": {"address": "alerts@example.com"}},
            "subject": "Your code",
            "bodyPreview": "",
        }
        self.assertTrue(matches_message(message, config))

    def test_matches_subject_keyword(self):
        config = build_config(subject_keywords=("verification code",))
        message = {
            "from": {"emailAddress": {"address": "random@example.com"}},
            "subject": "Your Verification Code",
            "bodyPreview": "",
        }
        self.assertTrue(matches_message(message, config))

    def test_matches_body_keyword(self):
        config = build_config(body_keywords=("one-time code",))
        message = {
            "from": {"emailAddress": {"address": "random@example.com"}},
            "subject": "Sign-in alert",
            "bodyPreview": "Your one-time code is 123456.",
        }
        self.assertTrue(matches_message(message, config))

    def test_move_all_overrides_filters(self):
        config = build_config(move_all=True)
        message = {
            "from": {"emailAddress": {"address": "spam@example.com"}},
            "subject": "spam",
            "bodyPreview": "spam",
        }
        self.assertTrue(matches_message(message, config))

    def test_unmatched_message_returns_false(self):
        config = build_config(allowed_senders=frozenset({"john.doe@example.com"}))
        message = {
            "from": {"emailAddress": {"address": "spam@example.com"}},
            "subject": "spam",
            "bodyPreview": "spam",
        }
        self.assertFalse(matches_message(message, config))

    def test_recent_filter_accepts_fresh_message(self):
        message = {
            "receivedDateTime": (datetime.now(timezone.utc) - timedelta(minutes=30)).isoformat(),
        }
        self.assertTrue(is_recent_enough(message, lookback_hours=1))

    def test_recent_filter_rejects_old_message(self):
        message = {
            "receivedDateTime": (datetime.now(timezone.utc) - timedelta(hours=3)).isoformat(),
        }
        self.assertFalse(is_recent_enough(message, lookback_hours=1))

    def test_parse_dotenv_file_reads_export_and_quotes(self):
        with TemporaryDirectory() as tmpdir:
            env_path = Path(tmpdir) / ".env"
            env_path.write_text(
                "\n".join(
                    [
                        "export OUTLOOK_CLIENT_ID='client-id'",
                        'OUTLOOK_TENANT_ID="consumers"',
                        "# comment",
                        "OUTLOOK_ALLOWED_SENDERS=john.doe@example.com,noreply@example.com",
                    ]
                ),
                encoding="utf-8",
            )

            self.assertEqual(
                parse_dotenv_file(env_path),
                {
                    "OUTLOOK_CLIENT_ID": "client-id",
                    "OUTLOOK_TENANT_ID": "consumers",
                    "OUTLOOK_ALLOWED_SENDERS": "john.doe@example.com,noreply@example.com",
                },
            )

    def test_build_wizard_defaults_prefills_example_matching_rules_for_new_config(self):
        defaults = build_wizard_defaults({})
        self.assertEqual(defaults.tenant_id, "consumers")
        self.assertEqual(defaults.allowed_senders, ("john.doe@example.com", "noreply@example.com"))
        self.assertEqual(defaults.subject_keywords, ("verification code", "login code"))

    def test_render_env_file_writes_expected_values(self):
        settings = EnvSettings(
            client_id="client-id",
            tenant_id="consumers",
            scopes=("Mail.ReadWrite",),
            allowed_senders=("john.doe@example.com",),
            allowed_domains=("example.com",),
            subject_keywords=("verification code",),
            body_keywords=("one-time code",),
            move_all=False,
            poll_seconds=120,
            lookback_hours=24,
            scan_limit=25,
            dry_run=True,
            token_cache_file=".tokens/msal_cache.json",
            log_level="DEBUG",
        )

        rendered = render_env_file(settings)

        self.assertIn("OUTLOOK_CLIENT_ID=client-id", rendered)
        self.assertIn("OUTLOOK_TENANT_ID=consumers", rendered)
        self.assertIn("OUTLOOK_SCOPES=Mail.ReadWrite", rendered)
        self.assertIn("OUTLOOK_ALLOWED_SENDERS=john.doe@example.com", rendered)
        self.assertIn("OUTLOOK_ALLOWED_DOMAINS=example.com", rendered)
        self.assertIn("OUTLOOK_DRY_RUN=true", rendered)
        self.assertIn("OUTLOOK_LOG_LEVEL=DEBUG", rendered)


if __name__ == "__main__":
    unittest.main()
