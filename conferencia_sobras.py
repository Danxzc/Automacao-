"""
Conferência de Sobras x Faltas
==============================
Lê um arquivo de FALTAS (artigos a procurar) e um CONSOLIDADO (saída da
extração WAVE) e, para cada falta, procura uma loja com SOBRA suficiente
do mesmo artigo (SKU).

Regras:
  - Vínculo: ARTIGO (faltas) == SKU (consolidado)
  - Uma sobra precisa cobrir a falta INTEIRA (sobra disponível >= qtd falta).
    Não divide uma falta entre duas lojas.
  - Uma mesma sobra pode atender várias faltas até esgotar
    (ex: loja com sobra 6 cobre uma falta de 3 e depois outra de 3).
  - Procura primeiro na data principal; o que não achar, procura nas
    outras datas e marca "VERIFICAR" para conferência manual.
  - Histórico: lembra entre execuções o quanto de cada sobra já foi
    comprometido (controle por DATA + CVR + Código da loja + SKU).
    Outro dia ou outro CVR = sobra independente.

Saída: planilha no mesmo formato do modelo, com colunas de FALTA e SOBRA.
"""

from __future__ import annotations

import logging
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd

# ── Diretório base: funciona como .py e como .exe (PyInstaller) ───────────────
SCRIPT_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent

HISTORICO_FILE = SCRIPT_DIR / "historico_sobras.csv"

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-8s %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(str(SCRIPT_DIR / "conferencia_sobras.log"), encoding="utf-8"),
    ],
)
logger = logging.getLogger("conferencia")

# ── Colunas esperadas ─────────────────────────────────────────────────────────
# No arquivo de FALTAS (o que você preenche)
COL_FALTA = {
    "cod":    "COD",
    "loja":   "LOJA",
    "cvr":    "CVR",
    "volume": "VOLUME",
    "qtd":    "QTD PCS",
    "artigo": "ARTIGO",
}
# No CONSOLIDADO (saída da extração WAVE)
COL_CONS = {
    "cvr":   "CVR",
    "cod":   "Código",
    "loja":  "Nome Loja",
    "data":  "Recebimento Visão Loja",
    "sku":   "SKU",
    "sobra": "Sobra",
}


# ─────────────────────────────────────────────────────────────────────────────
# Janela de configuração (Tkinter)
# ─────────────────────────────────────────────────────────────────────────────

def abrir_tela_configuracao() -> Optional[dict]:
    """Janela para escolher os arquivos de entrada e a pasta de saída."""
    import tkinter as tk
    from tkinter import filedialog, messagebox

    cfg = {"ok": False, "faltas": "", "consolidado": "", "saida": str(SCRIPT_DIR)}

    root = tk.Tk()
    root.title("Conferência de Sobras x Faltas")
    root.resizable(False, False)
    root.configure(bg="#f5f5f5")
    root.update_idletasks()
    w, h = 640, 320
    x = (root.winfo_screenwidth() - w) // 2
    y = (root.winfo_screenheight() - h) // 2
    root.geometry(f"{w}x{h}+{x}+{y}")

    tk.Frame(root, bg="#002868", height=50).pack(fill="x")
    tk.Label(root, text="🔍  Conferência de Sobras x Faltas",
             bg="#002868", fg="white", font=("Segoe UI", 13, "bold"),
             pady=12).place(relx=0.5, y=25, anchor="center")

    body = tk.Frame(root, bg="#f5f5f5", padx=20, pady=20)
    body.pack(fill="both", expand=True)

    vars_ = {
        "faltas":      tk.StringVar(),
        "consolidado": tk.StringVar(),
        "saida":       tk.StringVar(value=str(SCRIPT_DIR)),
    }

    def linha(label, key, tipo):
        row = tk.Frame(body, bg="#f5f5f5")
        row.pack(fill="x", pady=7)
        tk.Label(row, text=label, bg="#f5f5f5", font=("Segoe UI", 9),
                 anchor="w", width=22).pack(side="left")
        tk.Entry(row, textvariable=vars_[key], width=42,
                 font=("Segoe UI", 9), relief="solid", bd=1).pack(side="left")

        def browse():
            if tipo == "pasta":
                p = filedialog.askdirectory(title=label)
            else:
                p = filedialog.askopenfilename(
                    title=label, filetypes=[("Excel", "*.xlsx *.xls")])
            if p:
                vars_[key].set(p)

        tk.Button(row, text="📂", command=browse, font=("Segoe UI", 9),
                  relief="solid", bd=1, bg="#e8e8e8", cursor="hand2",
                  padx=4).pack(side="left", padx=(4, 0))

    linha("Arquivo de FALTAS:",   "faltas",      "arquivo")
    linha("Arquivo CONSOLIDADO:", "consolidado", "arquivo")
    linha("Pasta de saída:",      "saida",       "pasta")

    def confirmar():
        f, c, s = vars_["faltas"].get().strip(), vars_["consolidado"].get().strip(), vars_["saida"].get().strip()
        if not all([f, c, s]):
            messagebox.showwarning("Atenção", "Preencha todos os campos.")
            return
        if not Path(f).is_file():
            messagebox.showerror("Erro", f"Arquivo de faltas não encontrado:\n{f}")
            return
        if not Path(c).is_file():
            messagebox.showerror("Erro", f"Consolidado não encontrado:\n{c}")
            return
        cfg.update(ok=True, faltas=f, consolidado=c, saida=s)
        root.destroy()

    footer = tk.Frame(root, bg="#f5f5f5", pady=10)
    footer.pack()
    tk.Button(footer, text="✖  Cancelar", command=root.destroy,
              font=("Segoe UI", 10), bg="#e0e0e0", fg="#333",
              relief="solid", bd=1, padx=6, pady=5, cursor="hand2").pack(side="left", padx=10)
    tk.Button(footer, text="▶  Processar", command=confirmar,
              font=("Segoe UI", 10, "bold"), bg="#002868", fg="white",
              relief="flat", padx=6, pady=5, cursor="hand2").pack(side="left", padx=10)

    root.mainloop()
    return cfg if cfg["ok"] else None


# ─────────────────────────────────────────────────────────────────────────────
# Carregamento e normalização
# ─────────────────────────────────────────────────────────────────────────────

def normalizar_codigo(v) -> str:
    """Converte para string limpa, removendo '.0' de floats lidos do Excel."""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _validar_colunas(df: pd.DataFrame, esperadas: dict, nome_arquivo: str) -> None:
    faltando = [c for c in esperadas.values() if c not in df.columns]
    if faltando:
        logger.error("Colunas faltando em %s: %s", nome_arquivo, faltando)
        logger.error("Colunas disponíveis: %s", list(df.columns))
        sys.exit(1)


def carregar_consolidado(caminho: str) -> pd.DataFrame:
    df = pd.read_excel(caminho, dtype=str)
    df.columns = df.columns.str.strip()
    _validar_colunas(df, COL_CONS, "CONSOLIDADO")
    df[COL_CONS["sobra"]] = pd.to_numeric(df[COL_CONS["sobra"]], errors="coerce").fillna(0).astype(int)
    for k in ("sku", "cvr", "cod"):
        df[COL_CONS[k]] = df[COL_CONS[k]].apply(normalizar_codigo)
    df[COL_CONS["data"]] = df[COL_CONS["data"]].astype(str).str.strip()
    return df


# O consolidado (SKU) traz sempre 4 zeros à direita que não existem no
# artigo digitado nas faltas. Completamos o artigo com esses zeros antes
# de comparar — assim não há risco de casar errado (não cortamos o SKU).
ZEROS_SKU = "0000"


def carregar_faltas(caminho: str) -> pd.DataFrame:
    df = pd.read_excel(caminho, dtype=str)
    df.columns = df.columns.str.strip()
    _validar_colunas(df, COL_FALTA, "FALTAS")
    df[COL_FALTA["qtd"]] = pd.to_numeric(df[COL_FALTA["qtd"]], errors="coerce").fillna(0).astype(int)
    # artigo original (como digitado) preservado para o relatório
    df["_ARTIGO_ORIG"] = df[COL_FALTA["artigo"]].apply(normalizar_codigo)
    # artigo com zeros à direita, usado só internamente para casar com o SKU
    df[COL_FALTA["artigo"]] = df["_ARTIGO_ORIG"] + ZEROS_SKU
    return df


# ─────────────────────────────────────────────────────────────────────────────
# Histórico (persistência entre execuções)
# ─────────────────────────────────────────────────────────────────────────────

def carregar_historico() -> dict:
    """Retorna { (data, cvr, cod, sku): qtd_comprometida }."""
    hist: dict = {}
    if HISTORICO_FILE.exists():
        df = pd.read_csv(HISTORICO_FILE, dtype=str)
        for _, r in df.iterrows():
            chave = (str(r["data"]), str(r["cvr"]), str(r["cod"]), str(r["sku"]))
            hist[chave] = int(float(r["comprometido"]))
        logger.info("Histórico carregado: %d sobra(s) com uso registrado.", len(hist))
    return hist


def salvar_historico(historico: dict) -> None:
    linhas = [
        {"data": k[0], "cvr": k[1], "cod": k[2], "sku": k[3], "comprometido": v}
        for k, v in historico.items()
    ]
    pd.DataFrame(linhas).to_csv(HISTORICO_FILE, index=False, encoding="utf-8-sig")
    logger.info("Histórico salvo em: %s", HISTORICO_FILE.name)


# ─────────────────────────────────────────────────────────────────────────────
# Motor de matching
# ─────────────────────────────────────────────────────────────────────────────

def processar(faltas: pd.DataFrame, consolidado: pd.DataFrame,
              historico: dict) -> tuple[list[dict], dict]:
    comprometido: dict = defaultdict(int)
    for chave, qtd in historico.items():
        comprometido[chave] = qtd

    datas = sorted(consolidado[COL_CONS["data"]].dropna().unique())
    if not datas:
        logger.warning("Consolidado não tem datas válidas.")
    data_principal = datas[0] if datas else None
    outras_datas = [d for d in datas if d != data_principal]

    resultados: list[dict] = []
    pendentes: list[int] = []

    def disponivel(row, sku: str) -> int:
        chave = (row[COL_CONS["data"]], row[COL_CONS["cvr"]], row[COL_CONS["cod"]], sku)
        return int(row[COL_CONS["sobra"]]) - comprometido[chave]

    def casar(falta, datas_alvo) -> Optional[dict]:
        sku = falta[COL_FALTA["artigo"]]
        qtd = falta[COL_FALTA["qtd"]]
        if qtd <= 0:
            return None
        for data in datas_alvo:
            cand = consolidado[
                (consolidado[COL_CONS["sku"]] == sku) &
                (consolidado[COL_CONS["data"]] == data) &
                (consolidado[COL_CONS["sobra"]] > 0)
            ]
            for _, row in cand.iterrows():
                if disponivel(row, sku) >= qtd:
                    chave = (data, row[COL_CONS["cvr"]], row[COL_CONS["cod"]], sku)
                    comprometido[chave] += qtd
                    return {
                        "data":  data,
                        "cvr":   row[COL_CONS["cvr"]],
                        "cod":   row[COL_CONS["cod"]],
                        "loja":  row[COL_CONS["loja"]],
                        "qtd":   qtd,
                    }
        return None

    def montar(falta, match, obs) -> dict:
        linha = {
            "COD(FALTA)":       falta[COL_FALTA["cod"]],
            "LOJA(FALTA)":      falta[COL_FALTA["loja"]],
            "CVR(FALTA)":       falta[COL_FALTA["cvr"]],
            "VOLUME(FALTA)":    falta[COL_FALTA["volume"]],
            "QTD PÇS (FALTA)":  falta[COL_FALTA["qtd"]],
            "ARTIGO":           falta["_ARTIGO_ORIG"],
        }
        if match:
            linha.update({
                "COD(SOBRA)":       match["cod"],
                "LOJA(SOBRA)":      match["loja"],
                "CVR(SOBRA)":       match["cvr"],
                "QTD PÇS (SOBRA)":  match["qtd"],
                "DATA SOBRA":       match["data"],
                "OBS":              obs,
            })
        else:
            linha.update({
                "COD(SOBRA)": "", "LOJA(SOBRA)": "", "CVR(SOBRA)": "",
                "QTD PÇS (SOBRA)": "", "DATA SOBRA": "", "OBS": obs,
            })
        return linha

    # 1ª passada — data principal
    for idx, falta in faltas.iterrows():
        match = casar(falta, [data_principal]) if data_principal else None
        if match:
            resultados.append((idx, montar(falta, match, "")))
        else:
            pendentes.append(idx)

    # 2ª passada — pendentes nas outras datas
    for idx in pendentes:
        falta = faltas.loc[idx]
        match = casar(falta, outras_datas)
        if match:
            obs = f"SOBRA EM DATA DIFERENTE ({match['data']}) - VERIFICAR"
            resultados.append((idx, montar(falta, match, obs)))
        else:
            resultados.append((idx, montar(falta, None, "SEM SOBRA")))

    # reordena pela ordem original das faltas
    resultados.sort(key=lambda t: t[0])
    linhas = [linha for _, linha in resultados]

    return linhas, dict(comprometido)


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    cfg = abrir_tela_configuracao()
    if cfg is None:
        logger.info("Cancelado pelo usuário.")
        return

    logger.info("Faltas      : %s", cfg["faltas"])
    logger.info("Consolidado : %s", cfg["consolidado"])
    logger.info("Saída       : %s", cfg["saida"])

    faltas      = carregar_faltas(cfg["faltas"])
    consolidado = carregar_consolidado(cfg["consolidado"])
    historico   = carregar_historico()

    logger.info("Processando %d falta(s) contra %d linha(s) de consolidado...",
                len(faltas), len(consolidado))

    linhas, historico_novo = processar(faltas, consolidado, historico)

    df_out = pd.DataFrame(linhas, columns=[
        "COD(FALTA)", "LOJA(FALTA)", "CVR(FALTA)", "VOLUME(FALTA)",
        "QTD PÇS (FALTA)", "ARTIGO",
        "COD(SOBRA)", "LOJA(SOBRA)", "CVR(SOBRA)", "QTD PÇS (SOBRA)",
        "DATA SOBRA", "OBS",
    ])

    # Estatísticas
    com_sobra  = (df_out["OBS"] == "").sum()
    verificar  = df_out["OBS"].str.contains("VERIFICAR", na=False).sum()
    sem_sobra  = (df_out["OBS"] == "SEM SOBRA").sum()

    ts    = datetime.now().strftime("%Y%m%d_%H%M%S")
    saida = Path(cfg["saida"]) / f"conferencia_sobras_{ts}.xlsx"
    df_out.to_excel(saida, index=False)

    salvar_historico(historico_novo)

    logger.info("=" * 55)
    logger.info("✅ Casadas no dia        : %d", com_sobra)
    logger.info("⚠  Sobra em outra data   : %d (marcadas VERIFICAR)", verificar)
    logger.info("❌ Sem sobra             : %d", sem_sobra)
    logger.info("📄 Resultado salvo em    : %s", saida.name)

    try:
        import tkinter.messagebox as mb
        mb.showinfo(
            "Concluído",
            f"Processamento finalizado!\n\n"
            f"✅ Casadas no dia: {com_sobra}\n"
            f"⚠ Sobra em outra data: {verificar}\n"
            f"❌ Sem sobra: {sem_sobra}\n\n"
            f"Arquivo: {saida.name}",
        )
    except Exception:
        pass


if __name__ == "__main__":
    main()
