APP_NAME = "CUESTA DO SEU JEITO - Roteiros Turísticos"
APP_NAME_PDF = "CUESTA DO SEU JEITO"
APP_TAGLINE = "Roteiros Turísticos"
APP_VERSION = "1.1"
APP_VERSION_SHORT = "1.1"


def desktop_window_title() -> str:
    return f"{APP_NAME} v{APP_VERSION}"


def desktop_header_text() -> str:
    return desktop_window_title()


def pdf_version_label() -> str:
    return f"Versão {APP_VERSION}"
