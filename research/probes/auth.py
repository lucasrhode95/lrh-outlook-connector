"""Device-code sign-in for the research probes.

    python research/probes/auth.py graph      # Outlook Mobile -> Graph Mail.Read
    python research/probes/auth.py outlook     # One Outlook Web -> outlook.office.com (OWS)
    python research/probes/auth.py search    # One Outlook Web -> outlook.office.com/search (Substrate)
    python research/probes/auth.py --status

'search' is first attempted silently with the 'outlook' refresh token (same client),
so a second sign-in is usually unnecessary.
"""

from __future__ import annotations

import argparse

from common import PROFILES, SignInRequired, claims, device_code_sign_in, emit, get_token, token_status


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("profile", nargs="?", choices=sorted(PROFILES))
    parser.add_argument("--status", action="store_true")
    parser.add_argument("--force", action="store_true", help="Run device code even if a silent token works.")
    args = parser.parse_args()
    if args.status or not args.profile:
        emit(token_status())
        return 0
    source = "device_code"
    if not args.force:
        try:
            get_token(args.profile)
            source = "silent"
        except SignInRequired:
            device_code_sign_in(args.profile)
    else:
        device_code_sign_in(args.profile)
    status = token_status().get(args.profile, {})
    emit({"profile": args.profile, "token_source": source, **status})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
