#!/usr/bin/env python3
"""
descargador_odata.py - Descarga datos de Datos.gov.co (OData v4) con logica condicional por reporte.

REGLAS DE REPORTE
-----------------
- --reporte "jbjy-vk9h" -> descarga directa a ContratosElectronicos.csv (bloques de 20000).
- --reporte "dmgg-8hin" -> descarga directa SIN CRUCE (bloques de 20000).
- --reporte <cualquier otro> -> streaming + cruce por id_contrato / id_contrato
"""

import argparse
import csv
import os
import sys
import time
from urllib.parse import quote

import requests

# ---------------------------------------------------------------------------
# Configuracion global
# ---------------------------------------------------------------------------

BASE_URL = "https://www.datos.gov.co/api/odata/v4"
DEFAULT_TOP = 20000           # bloque por peticion
DEFAULT_DELAY = 1.0           # segundos entre peticiones (respeto al servidor)
REQUEST_TIMEOUT = 600         # 10 minutos por peticion (para paginas lentas)
CONTRATOS_DEFAULT = "ContratosElectronicos.csv"

# ---------------------------------------------------------------------------
# Helpers URL
# ---------------------------------------------------------------------------

def _encode_filter(raw: str) -> str:
    """Codifica el valor de $filter: espacios -> %20, pero deja ':' sin codificar (como el navegador)."""
    return quote(raw, safe=":")


def build_first_url(reporte: str, filter_expr: str | None, top: int) -> str:
    """URL inicial con $filter, $top y $skip=0 explicito (como usa el navegador)."""
    seg = f"{reporte}"
    parts = [f"{BASE_URL}/{seg}"]
    qs: list[str] = []
    if filter_expr:
        qs.append(f"$filter={_encode_filter(filter_expr)}")
    if top:
        qs.append(f"$top={top}")
    qs.append("$skip=0")   # explicito, como en la consulta del navegador
    if qs:
        parts.append("?" + "&".join(qs))
    return "".join(parts)


# ---------------------------------------------------------------------------
# Descarga de una sola pagina
# ---------------------------------------------------------------------------

def fetch_page(url: str, session: requests.Session) -> dict:
    """Obtiene una pagina JSON del servicio. Lanza en caso de error."""
    resp = session.get(url, timeout=REQUEST_TIMEOUT)
    if resp.status_code != 200:
        raise RuntimeError(
            f"HTTP {resp.status_code} en {url}:\n{resp.text[:500]}"
        )
    return resp.json()


def extract_next_link(page: dict) -> str | None:
    """Devuelve @odata.nextLink o None si se termino."""
    return page.get("@odata.nextLink")


def extract_count(page: dict) -> int | None:
    """Devuelve @odata.count si esta presente."""
    return page.get("@odata.count")


# ---------------------------------------------------------------------------
# Normalizacion de registros / filas de CSV
# ---------------------------------------------------------------------------

def normalize_record(record: dict) -> dict:
    """Limpia un registro OData: None -> "", \n literales -> saltos reales."""
    out: dict = {}
    for k, v in record.items():
        if isinstance(v, str):
            out[k] = v.replace("\\n", "\n").strip()
        elif v is None:
            out[k] = ""
        else:
            out[k] = v
    return out


def id_str(v) -> str:
    """Convierte un valor a cadena normalizada (blancos/None -> "")."""
    if v is None:
        return ""
    if isinstance(v, str):
        return v.strip()
    return str(v).strip()


# ---------------------------------------------------------------------------
# Lectura de ContratosElectronicos.csv (tabla de referencia)
# ---------------------------------------------------------------------------

def load_reference_csv(path: str) -> set[str]:
    """
    Lee el CSV de ContratosElectronicos.csv y devuelve un conjunto con los
    valores distintos de la columna id_contrato (normalizados a cadena).
    """
    if not os.path.isfile(path):
        raise FileNotFoundError(f"No se encontro la tabla de referencia: {path}")

    ids: set[str] = set()
    with open(path, newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        if not reader.fieldnames:
            return ids
        if "id_contrato" not in reader.fieldnames:
            raise RuntimeError(
                f"El CSV de referencia '{path}' no contiene la columna 'id_contrato'. "
                f"Columnas disponibles: {reader.fieldnames}"
            )
        for row in reader:
            val = id_str(row.get("id_contrato"))
            if val:
                ids.add(val)
    return ids


# ---------------------------------------------------------------------------
# Caso A - descarga directa (jbjy-vk9h y dmgg-8hin)
# ---------------------------------------------------------------------------

def download_direct(
    reporte: str,
    filter_expr: str | None,
    output: str,
    top: int,
    delay: float,
    verbose: bool,
) -> int:
    """
    Descarga todos los registros en bloques de *top* y los escribe en CSV.
    Retorna el total descargado.
    """
    session = requests.Session()
    session.headers.update({
        "User-Agent": "SecopODataDownloader/1.0 (datos.gov.co)",
        "Accept": "application/json",
    })

    url = build_first_url(reporte, filter_expr, top)
    all_keys: list[str] = []
    total = 0
    batch = 0

    if verbose:
        print(f"[URL] URL inicial: {url}")
        print(f"[PAQ] Bloques de {top} registros")
        print(f"[OUT] Salida: {output}")
        print("-" * 60)

    with open(output, "w", newline="", encoding="utf-8-sig") as fh:
        writer: csv.DictWriter | None = None

        while url:
            batch += 1
            if verbose:
                snippet = url[:110]
                print(f"\n[ ESP ] Bloque {batch} - {snippet}...")

            page = fetch_page(url, session)
            records = page.get("value", [])

            if not records:
                if verbose:
                    print("  -> Sin registros. Terminando.")
                break

            # Primera pagina: capturar encabezados
            if not all_keys:
                all_keys = list(records[0].keys())
                writer = csv.DictWriter(fh, fieldnames=all_keys)
                writer.writeheader()
                if verbose:
                    print(f"  [HDR] {len(all_keys)} columnas detectadas")

            # Normalizar y escribir
            writer.writerows(normalize_record(r) for r in records)
            total += len(records)
            if verbose:
                print(f"  [OK] {len(records)} registros (acumulado: {total})")

            # Siguiente
            nxt = extract_next_link(page)
            url = nxt if nxt else None
            if url and delay:
                time.sleep(delay)

    if verbose:
        sz = os.path.getsize(output) if os.path.isfile(output) else 0
        print("\n" + "=" * 60)
        print(f"[OK]  Descarga directa completa.")
        print(f"    Registros: {total}")
        print(f"    Archivo: {output}")
        print(f"    Tamano: {sz:,} bytes")
        print("=" * 60)

    return total


# ---------------------------------------------------------------------------
# Caso B - descarga streaming + cruce por id_contrato con ContratosElectronicos.csv
# ---------------------------------------------------------------------------

def download_and_cross(
    reporte: str,
    filter_expr: str | None,
    output: str,
    top: int,
    delay: float,
    referencia_path: str,
    verbose: bool,
) -> int:
    """
    Descarga el reporte pagina por pagina (streaming, sin acumular en memoria),
    carga ContratosElectronicos.csv como tabla de referencia, cruza por id_contrato
    y guarda solo las filas coincidentes en *output* inmediatamente.

    Retorna el numero de filas escritas en *output*.

    NOTA: dmgg-8hin NO usa esta funcion; es descarga directa sin cruce.
    """
    # 1. Cargar tabla de referencia
    if verbose:
        print(f"[CAR] Cargando tabla de referencia: {referencia_path}")
    ref_set = load_reference_csv(referencia_path)

    if verbose:
        print(f"   -> {len(ref_set):,} valores de id_contrato en referencia")

    # 2. Inicializar descarga paginada
    session = requests.Session()
    session.headers.update({
        "User-Agent": "SecopODataDownloader/1.0 (datos.gov.co)",
        "Accept": "application/json",
    })

    url = build_first_url(reporte, filter_expr, top)
    all_keys: list[str] = []
    total = 0         # registros descargados (todos, sin filtrar)
    matched = 0       # registros que pasaron el cruce y se escribieron
    batch = 0

    if verbose:
        print(f"[URL] URL inicial: {url}")
        print(f"[PAQ] Descargando paginado + cruce en streaming (bloques de {top})...")
        if not ref_set:
            print("   [ATENCION] Tabla de referencia vacia. Resultado sera 0 filas.")
        print("-" * 60)

    with open(output, "w", newline="", encoding="utf-8-sig") as fh:
        writer: csv.DictWriter | None = None

        first_page = True
        while url:
            batch += 1
            if verbose:
                snippet = url[:110]
                print(f"\n[ ESP ] Bloque {batch} - {snippet}...")

            page = fetch_page(url, session)
            records = page.get("value", [])

            if not records:
                if verbose:
                    print("  -> Sin registros. Terminando.")
                break

            # Primera pagina: capturar columnas y escribir encabezado
            if first_page:
                all_keys = list(records[0].keys())
                if "id_contrato" not in all_keys:
                    raise RuntimeError(
                        f"El reporte '{reporte}' no tiene la columna 'id_contrato'. "
                        f"Columnas: {all_keys}. No se puede realizar el cruce."
                    )
                writer = csv.DictWriter(fh, fieldnames=all_keys)
                writer.writeheader()
                first_page = False
                if verbose:
                    print(f"  [HDR] {len(all_keys)} columnas detectadas, "
                          f"encabezado escrito (cruce por id_contrato)")

            # Procesar registro por registro (streaming: no acumulamos en memoria)
            for rec in records:
                total += 1
                valor_cruce = id_str(rec.get("id_contrato"))
                if valor_cruce in ref_set:
                    writer.writerow(normalize_record(rec))
                    matched += 1

            if verbose:
                print(f"  [OK] {len(records)} registros "
                      f"(acumulado descargados: {total}, "
                      f"coincidentes escritos: {matched})")

            nxt = extract_next_link(page)
            url = nxt if nxt else None
            if url and delay:
                time.sleep(delay)

    # 3. Reporte final
    sz = os.path.getsize(output) if os.path.isfile(output) else 0
    if verbose:
        if not ref_set:
            print("\n[ATENCION] Tabla de referencia vacia. "
                  "El resultado del cruce sera 0 filas (archivo con encabezados).")
        else:
            print(f"\n[OK] Cruce: {matched:,} de {total:,} registros coinciden con id_contrato")
        print("\n" + "=" * 60)
        print(f"[OK]  Descarga + cruce completado (streaming).")
        print(f"    Reporte:                                    {reporte}")
        print(f"    Registros descargados (total, sin filtrar): {total:,}")
        print(f"    Valores en referencia (id_contrato):          {len(ref_set):,}")
        print(f"    Registros despues del cruce (escritos):     {matched:,}")
        print(f"    Archivo de salida:                          {output}")
        print(f"    Tamano:                                     {sz:,} bytes")
        print("=" * 60)

    if matched == 0 and total > 0:
        # hubo registros pero ninguno coincidio
        print(f"\n[ATENCION] Descargados {total:,} registros pero 0 coinciden con "
              f"id_contrato en {referencia_path}.",
              file=sys.stderr)

    return matched


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Descarga OData de Datos.gov.co con logica condicional por reporte.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
REGLAS DE REPORTE
  --reporte "jbjy-vk9h"
      -> Descarga directa a ContratosElectronicos.csv (bloques de 20000).
      -> Es el modo que descarga TODO y lo guarda tal cual.

  --reporte "dmgg-8hin"
      -> Descarga directa SIN CRUCE (bloques de 20000).
      -> Descarga todos los registros del reporte sin cruzar con
        ContratosElectronicos.csv.
      -> Ejemplo:
        python descargador_odata.py \\
            --reporte "dmgg-8hin" \\
            --filter "nit_entidad eq 900971006 and fecha_carga ge '2026-01-01T00:00:00' and fecha_carga le '2026-01-31T23:59:59'" \\
            --output "ArchivosDescargadosDesde2025202601"

  --reporte <cualquier otro>
      -> Streaming + cruce por id_contrato / id_contrato.
      -> Ejemplo:
        python descargador_odata.py \\
            --reporte "cb9c-h8sn" \\
            --filter "fecharegistro ge '2026-01-01T00:00:00' and fecharegistro le '2026-01-31T23:59:59'" \\
            --output "Adiciones202601.csv" \\
            --top 5000

NOTA IMPORTANTE
  El Caso B requiere que ContratosElectronicos.csv EXISTE. Ejecuta primero:
    python descargador_odata.py --reporte jbjy-vk9h --filter "nit_entidad eq 900971006"

  IMPORTANTE: CODIFICACION DE URL
    El servidor OData de Datos.gov.co rechaza urllib.parse.urlencode porque
    convierte espacios en '+' y los dos puntos ':' en '%3A'. El script usa
    urllib.parse.quote(filter_raw, safe=":") que deja ':' intacto y usa %20
    para espacios, igual que en la barra del navegador.

        """,
    )

    p.add_argument("--reporte", "-r", required=True,
                   help="Identificador del dataset OData (ej. jbjy-vk9h, cb9c-h8sn, dmgg-8hin)")
    p.add_argument("--filter", "-f", default=None,
                   help="Filtro OData opcional (ej. 'nit_entidad eq 900971006')")
    p.add_argument("--top", "-t", type=int, default=DEFAULT_TOP,
                   help=f"Registros por peticion (default: {DEFAULT_TOP})")
    p.add_argument("--output", "-o", default=None,
                   help="Archivo de salida (default automatico segun reporte)")
    p.add_argument("--contratos-path", default=CONTRATOS_DEFAULT,
                   help=f"Ruta al CSV maestro ContratosElectronicos.csv "
                        f"(default: {CONTRATOS_DEFAULT})")
    p.add_argument("--delay", type=float, default=DEFAULT_DELAY,
                   help=f"Pausa entre peticiones en s (default: {DEFAULT_DELAY})")
    p.add_argument("--no-delay", action="store_true",
                   help="Sin pausa entre peticiones")
    p.add_argument("--quiet", "-q", action="store_true",
                   help="No mostrar progreso")

    return p.parse_args()


def default_output(reporte: str, is_direct: bool) -> str:
    """Nombre de salida por defecto segun reglas."""
    if is_direct:
        # Caso A: jbjy-vk9h -> ContratosElectronicos.csv, dmgg-8hin -> dmgg-8hin.csv
        if reporte == "jbjy-vk9h":
            return CONTRATOS_DEFAULT
        if reporte == "dmgg-8hin":
            return "dmgg-8hin.csv"
        # otro reporte directo (si se agrega en el futuro) -> <reporte>.csv
        return f"{reporte}.csv"
    # Caso B: <reporte>_cruce.csv  (si no se paso --output)
    base = reporte.replace("/", "_").replace("\\", "_")
    return f"{base}_cruce.csv"


def main() -> None:
    args = parse_args()
    delay = 0.0 if args.no_delay else args.delay
    verbose = not args.quiet
    is_direct = (args.reporte in ("jbjy-vk9h", "dmgg-8hin"))

    # Determinar output
    output = args.output or default_output(args.reporte, is_direct)

    if is_direct:
        # Descarga directa sin cruce (jbjy-vk9h o dmgg-8hin)
        if args.reporte == "jbjy-vk9h" and args.output:
            if verbose:
                print(f"[INFO] jbjy-vk9h pero --output={args.output}. "
                      f"Se escribe aqui en lugar de {CONTRATOS_DEFAULT}.")
        try:
            total = download_direct(
                reporte=args.reporte,
                filter_expr=args.filter,
                output=output,
                top=args.top,
                delay=delay,
                verbose=verbose,
            )
            if total == 0:
                print("[ATENCION] No se descargaron registros.", file=sys.stderr)
                sys.exit(1)
        except requests.exceptions.RequestException as e:
            print(f"\n[ERROR] Error de conexion: {e}", file=sys.stderr)
            sys.exit(2)
        except RuntimeError as e:
            print(f"\n[ERROR] Error: {e}", file=sys.stderr)
            sys.exit(3)
    else:
        # Caso B - requiere ContratosElectronicos.csv existente
        if not os.path.isfile(args.contratos_path):
            print(f"\n[ERROR] ERROR: no existe '{args.contratos_path}'. "
                  f"Ejecute primero: python descargador_odata.py "
                  f"--reporte jbjy-vk9h --filter 'nit_entidad eq 900971006'",
                  file=sys.stderr)
            sys.exit(4)

        try:
            total = download_and_cross(
                reporte=args.reporte,
                filter_expr=args.filter,
                output=output,
                top=args.top,
                delay=delay,
                referencia_path=args.contratos_path,
                verbose=verbose,
            )
            if total == 0:
                print("\n[ATENCION] Resultado del cruce: 0 filas escritas.",
                      file=sys.stderr)
                # No es error fatal, pero avisamos
        except FileNotFoundError as e:
            print(f"\n[ERROR] {e}", file=sys.stderr)
            sys.exit(4)
        except RuntimeError as e:
            print(f"\n[ERROR] Error: {e}", file=sys.stderr)
            sys.exit(3)
        except requests.exceptions.RequestException as e:
            print(f"\n[ERROR] Error de conexion: {e}", file=sys.stderr)
            sys.exit(2)


if __name__ == "__main__":
    main()