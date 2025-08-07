from .config import get_config, set_config
from .outlook import (
    get_flagged_emails,
    get_flagged_emails_in_month,
    is_outlook_installed,
    get_flagged_emails_in_month_pst,
)

__all__ = [
    "get_config",
    "set_config",
    "is_outlook_installed",
    "get_flagged_emails",
    "get_flagged_emails_in_month",
    "get_flagged_emails_in_month_pst",
]
