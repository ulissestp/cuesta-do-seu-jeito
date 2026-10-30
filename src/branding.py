APP_NAME = "CUESTA DO SEU JEITO - RT"
APP_NAME_PDF = "CUETA DO SEU JEITO"
APP_TAGLINE = "Roteiros Turísticos"
APP_VERSION = "1.0 Beta"
APP_VERSION_SHORT = "1.0-beta"


def desktop_window_title() -> str:
    return f"{APP_NAME} — v{APP_VERSION}"


def desktop_header_text() -> str:
    return f"{APP_NAME} · {APP_TAGLINE} · v{APP_VERSION}"


def pdf_version_label() -> str:
    return f"Versão {APP_VERSION}"
