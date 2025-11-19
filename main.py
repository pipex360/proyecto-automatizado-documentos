#!/usr/bin/env python3
"""
Script principal para ejecutar el Automatizador de Documentos Word

Uso:
    python main.py
"""
import sys
import os

# Agregar directorio src al path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from interfaz_grafica import main

if __name__ == "__main__":
    print("=" * 60)
    print("    AUTOMATIZADOR DE DOCUMENTOS WORD")
    print("=" * 60)
    print()
    print("Iniciando aplicación...")
    print()

    try:
        main()
    except KeyboardInterrupt:
        print("\nAplicación cerrada por el usuario")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ Error al ejecutar la aplicación: {e}")
        sys.exit(1)
