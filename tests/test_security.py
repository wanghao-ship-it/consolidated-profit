import sys
import unittest
from pathlib import Path

SCRIPTS = Path(__file__).parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from renheng_web.security import SessionTokens, hash_password, verify_password


class SecurityTests(unittest.TestCase):
    def test_password_hash_is_salted_and_verifiable(self):
        first = hash_password("a sufficiently long password")
        second = hash_password("a sufficiently long password")
        self.assertNotEqual(first, second)
        self.assertTrue(verify_password("a sufficiently long password", first))
        self.assertFalse(verify_password("wrong password", first))

    def test_session_token_only_resolves_signed_user_id(self):
        sessions = SessionTokens("s" * 48, max_age_seconds=60)
        token = sessions.issue({"user_id": "user-123", "role": "ADMIN"})
        self.assertEqual(sessions.read_user_id(token), "user-123")
        self.assertIsNone(sessions.read_user_id("tampered" + token))
        self.assertIsNone(sessions.read_user_id(None))

    def test_session_secret_must_not_be_short(self):
        with self.assertRaises(ValueError):
            SessionTokens("too-short")


if __name__ == "__main__":
    unittest.main()
