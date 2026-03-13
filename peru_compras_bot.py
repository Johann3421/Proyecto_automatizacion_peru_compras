# -*- coding: utf-8 -*-
"""Entrypoint principal de Peru Compras Bot."""

import sys

from peru_compras_bot_app.automation import main_cli
from peru_compras_bot_app.gui import iniciar_interfaz


def main():
    if "--cli" in sys.argv:
        main_cli()
    else:
        iniciar_interfaz()


if __name__ == "__main__":
    main()