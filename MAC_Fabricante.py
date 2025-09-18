import pandas as pd
import re
import time
import requests
import json
from pathlib import Path

API_URL = "https://api.macvendors.com/"

def clean_mac(mac: str) -> str | None:
    """Remove separadores e retorna os 12 hex do MAC em maiúsculas."""
    if pd.isna(mac):
        return None
    s = re.sub(r'[^0-9A-Fa-f]', '', str(mac))
    s = s.upper()
    return s[:12] if len(s) >= 12 else None

def mac_to_int(mac: str) -> int:
    try:
        return int(mac, 16)
    except Exception:
        return -1

def fetch_vendor(key: str, retry: int = 2, sleep_secs: float = 1.0) -> str:
    """Consulta a API do MACVendors."""
    url = API_URL + key
    for attempt in range(retry + 1):
        try:
            r = requests.get(url, timeout=10)
        except Exception:
            if attempt == retry:
                return "Desconhecido"
            time.sleep(sleep_secs)
            continue

        if r.status_code == 200:
            return r.text.strip()
        elif r.status_code in (404, 204):
            return "Desconhecido"
        elif r.status_code == 429:  # limite de taxa
            time.sleep(max(2.0, sleep_secs * 2))
            continue
        else:
            if attempt == retry:
                return "Desconhecido"
            time.sleep(sleep_secs)
    return "Desconhecido"

def load_cache(cache_path: Path) -> dict[str, str]:
    if cache_path.is_file():
        try:
            with cache_path.open("r", encoding="utf-8") as f:
                data = json.load(f)
            return {k.upper()[:6]: v for k, v in data.items() if isinstance(k, str)}
        except Exception:
            return {}
    return {}

def save_cache(cache_path: Path, cache: dict[str, str]):
    try:
        with cache_path.open("w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def main(path_in: str, path_out: str, cache_file: str = "oui_cache.json"):
    print(">> Lendo planilha...")
    df = pd.read_excel(path_in)

    if "MAC" not in df.columns:
        raise ValueError("A planilha precisa ter a coluna 'MAC'.")

    # Normalizar MAC e gerar OUI
    df["MAC_clean"] = df["MAC"].apply(clean_mac)
    df["OUI"] = df["MAC_clean"].str[:6]

    # Carregar cache existente
    cache_path = Path(cache_file) if cache_file else None
    vendor_by_oui: dict[str, str] = load_cache(cache_path) if cache_path else {}
    if vendor_by_oui:
        print(f">> Cache carregado: {len(vendor_by_oui)} OUIs")

    # Determinar OUIs
    ouis_all = sorted([x for x in df["OUI"].dropna().unique()])
    to_fetch = [o for o in ouis_all if o not in vendor_by_oui]
    print(f">> OUIs únicos: {len(ouis_all)} | Em cache: {len(ouis_all)-len(to_fetch)} | A consultar: {len(to_fetch)}")

    for i, oui in enumerate(to_fetch, 1):
        vendor = fetch_vendor(oui)
        vendor_by_oui[oui] = vendor
        print(f"[fetch {i}/{len(to_fetch)}] {oui} -> {vendor}")
        time.sleep(1.0)  # respeitar limite da API

    if cache_path:
        save_cache(cache_path, vendor_by_oui)
        print(f">> Cache salvo: {cache_path}")

    # Atribuir fabricante
    df["Fabricante"] = df["OUI"].map(vendor_by_oui)

    # Ordenar pelo MAC (usar série temporária, não manter coluna)
    mac_int = df["MAC_clean"].apply(mac_to_int)
    df_sorted = df.assign(_mac_int=mac_int).sort_values(by=["_mac_int", "MAC_clean"]).drop(columns=["_mac_int"], errors="ignore")

    # Salvar resultado somente com as colunas desejadas
    required_cols = ["Login", "MAC", "Fabricante"]
    missing = [c for c in required_cols if c not in df_sorted.columns]
    if missing:
        raise ValueError(f"Colunas obrigatórias ausentes no dataframe final: {missing}")

    output_df = df_sorted[required_cols]
    Path(path_out).parent.mkdir(parents=True, exist_ok=True)
    output_df.to_excel(path_out, index=False)
    print(f">> Arquivo gerado: {path_out} (colunas: {', '.join(required_cols)})")


    # Função Main
if __name__ == "__main__":
    main("Login_inicio.xlsx", "login_final.xlsx")
