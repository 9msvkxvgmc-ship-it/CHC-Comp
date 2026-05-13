from dataclasses import dataclass, field


@dataclass
class Config:
    excluded_signers: list[str] = field(default_factory=list)
    reclassify_as_supervising_md: list[str] = field(default_factory=list)
    add_to_app_list: list[str] = field(default_factory=list)
    omit: list[str] = field(default_factory=list)
    app_credential_patterns: list[str] = field(
        default_factory=lambda: [
            r"\bPA-C\b", r"\bPA\b", r"\bNP\b", r"\bAGNP-C\b",
            r"\bCRNP\b", r"\bFNP\b", r"\bANP\b", r"\bDNP\b",
            r"\bAPRN\b", r"\bMSN\b",
        ]
    )
    wrvu_file: str | None = None  # path to CMS RVU CSV; None = use built-in table
    location_filter: str | None = None  # if set, only include charges from this location


CHCWM_Q1_2026 = Config(
    excluded_signers=["Hogan, Jennifer", "Fancovic, Lisa", "Monetza, Liz"],
    reclassify_as_supervising_md=["Eastburg MD, Luke", "Fabricius MD, William", "Wong MD, Frances"],
    add_to_app_list=["Tol, Margie"],
    omit=["Porter PhD, Jeff"],
)
