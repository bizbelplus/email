"""Email and data validators."""

from __future__ import annotations

import re
from dataclasses import dataclass


@dataclass
class EmailValidationResult:
    """Result of email validation."""

    email: str
    is_valid: bool
    error: str | None = None
    warnings: list[str] | None = None

    @property
    def is_clean(self) -> bool:
        """Check if valid with no warnings."""
        return self.is_valid and (not self.warnings or len(self.warnings) == 0)


# RFC 5322 simplified regex for basic email validation
EMAIL_REGEX = re.compile(
    r"^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$"
)


def validate_email(email: str, check_mx: bool = False) -> EmailValidationResult:
    """
    Validate email address.

    Args:
        email: Email address to validate
        check_mx: If True, check MX records (requires dnspython)

    Returns:
        EmailValidationResult with validation status and details
    """
    original_email = email
    email = email.strip()
    warnings: list[str] = []

    # Check for empty
    if not email:
        return EmailValidationResult(original_email, is_valid=False, error="Email is empty")

    # Check for common typos/issues
    if email.count("@") != 1:
        return EmailValidationResult(original_email, is_valid=False, error="Email must contain exactly one @ symbol")

    # Basic regex check
    if not EMAIL_REGEX.match(email):
        return EmailValidationResult(original_email, is_valid=False, error="Email format is invalid")

    # Check for consecutive dots
    if ".." in email:
        warnings.append("Email contains consecutive dots")

    # Check if starts/ends with dot
    local_part, domain = email.rsplit("@", 1)
    if local_part.startswith(".") or local_part.endswith("."):
        warnings.append("Local part starts or ends with dot")

    if domain.startswith(".") or domain.endswith("."):
        warnings.append("Domain starts or ends with dot")

    # Check domain length (max 255 characters)
    if len(domain) > 255:
        warnings.append("Domain part exceeds 255 characters")

    # Check local part length (max 64 characters)
    if len(local_part) > 64:
        warnings.append("Local part exceeds 64 characters")

    # Optional MX record check
    if check_mx:
        try:
            import dns.resolver  # type: ignore

            try:
                dns.resolver.resolve(domain, "MX")
            except (dns.resolver.NXDOMAIN, dns.resolver.NoAnswer):
                warnings.append(f"MX records not found for domain: {domain}")
            except Exception as error:
                warnings.append(f"Could not check MX records: {error}")
        except ImportError:
            pass  # dnspython not installed, skip MX check

    return EmailValidationResult(
        email=email,
        is_valid=True,
        warnings=warnings if warnings else None,
    )


def validate_csv_recipients(csv_path: str) -> dict:
    """
    Validate all emails in a CSV file.

    Returns dict with:
    - valid: list of valid emails
    - invalid: list of (email, error)
    - with_warnings: list of (email, warnings)
    """
    from pathlib import Path

    from .recipients import load_recipients

    path = Path(csv_path)
    try:
        recipients = load_recipients(path)
    except Exception as error:
        return {"error": str(error), "valid": [], "invalid": [], "with_warnings": []}

    valid = []
    invalid = []
    with_warnings = []

    for recipient in recipients:
        result = validate_email(recipient.email, check_mx=False)
        if not result.is_valid:
            invalid.append((recipient.email, result.error))
        elif result.warnings:
            with_warnings.append((recipient.email, result.warnings))
        else:
            valid.append(recipient.email)

    return {
        "valid": valid,
        "invalid": invalid,
        "with_warnings": with_warnings,
        "total_valid": len(valid),
        "total_invalid": len(invalid),
        "total_warnings": len(with_warnings),
    }
