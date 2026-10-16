#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from generar_informe_Version3 import generate_report
import sys
import traceback

if __name__ == '__main__':
    try:
        ok = generate_report()
        if ok:
            print("Proceso finalizado correctamente.")
        else:
            print("Proceso finalizado con errores.")
    except Exception:
        traceback.print_exc()
        sys.exit(1)
