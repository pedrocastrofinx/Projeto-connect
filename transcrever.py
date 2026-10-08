#!/usr/bin/env python3
"""
FINX · Transcrição das ligações da IA (Natasha)

Roda no SEU computador (as gravações ficam atrás do login do 3C Plus; o painel HTML não consegue baixá-las).
Gera o arquivo transcricoes.csv — é só soltá-lo no painel (aba IA Natasha), junto com as outras planilhas.

Instalação (uma vez):
    pip install faster-whisper            # transcrição offline (precisa do ffmpeg só se o áudio for estranho)
    pip install openpyxl                  # opcional: para ler também a planilha "agente IA perdidos por motivo"

Uso — baixando sozinho do 3C Plus:
    python transcrever.py --csv telefonia.csv --xlsx agente_ia_perdidos.xlsx --token SEU_TOKEN

Uso — com uma pasta de áudios que você já baixou (o nome do arquivo deve conter o id da ligação ou o código da gravação):
    python transcrever.py --csv telefonia.csv --pasta C:\\gravacoes

Opções úteis:
    --min 10        só ligações com 10 s de fala ou mais (padrão). Mais curtas quase sempre são o cliente desligando.
    --max 50        faz só as 50 maiores (bom para testar)
    --modelo small  tiny | base | small | medium | large-v3 (small é o equilíbrio; medium é melhor e mais lento)
    --saida transcricoes.csv

Pode interromper (Ctrl+C) e rodar de novo: continua de onde parou.
O texto fica só no seu computador e no navegador; nada é enviado para fora.
"""
import argparse
import csv
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

BASES = ["https://app.3c.plus", "https://finx.3c.plus"]  # a API pode responder pelo domínio da empresa; o script tenta os dois
URL_GRAVACAO = "/api/v1/calls/{id}/recording"


def segundos(v):
    v = str(v or "").strip()
    m = re.match(r"^(\d+):(\d{2}):(\d{2})$", v)
    if m:
        return int(m[1]) * 3600 + int(m[2]) * 60 + int(m[3])
    m = re.match(r"^(\d+):(\d{2})$", v)
    if m:
        return int(m[1]) * 60 + int(m[2])
    try:
        return int(float(v.replace(",", ".")))
    except ValueError:
        return 0


def chave(nome):
    return re.sub(r"[^a-z0-9]", "", str(nome or "").lower())


def ler_csv(caminho):
    """CSV da telefonia (3C Plus): _id, speaking_time, record_name."""
    with open(caminho, encoding="utf-8-sig", newline="") as f:
        amostra = f.read(4096)
        f.seek(0)
        sep = ";" if amostra.count(";") >= amostra.count(",") else ","
        linhas = list(csv.reader(f, delimiter=sep))
    cab = [chave(c) for c in linhas[0]]

    def col(*nomes):
        for n in nomes:
            if n in cab:
                return cab.index(n)
        return -1

    ci, cf, cr = col("id"), col("speakingtime"), col("recordname")
    cn, cd = col("number"), col("calldate")
    if ci < 0:
        sys.exit("O CSV não tem a coluna _id (esperado: exportação do Relatório de Ligações do 3C Plus).")
    out = {}
    for r in linhas[1:]:
        if len(r) <= ci or not r[ci].strip():
            continue
        out[r[ci].strip()] = {
            "dur": segundos(r[cf]) if cf >= 0 and cf < len(r) else 0,
            "rec": r[cr].strip() if cr >= 0 and cr < len(r) else "",
            "num": re.sub(r"\D", "", r[cn])[-11:] if cn >= 0 and cn < len(r) else "",
            "quando": r[cd].strip() if cd >= 0 and cd < len(r) else "",
        }
    return out


def ler_xlsx(caminho):
    """Planilha 'agente IA perdidos por motivo': colunas IA - id / IA - speaking_time / IA - record_name (listas separadas por ' | ')."""
    try:
        import openpyxl
    except ImportError:
        print("  (openpyxl não instalado: pulando a planilha. Instale com: pip install openpyxl)")
        return {}
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    ws = wb.active
    linhas = ws.iter_rows(values_only=True)
    cab = [chave(c) for c in next(linhas)]
    ix = lambda n: cab.index(n) if n in cab else -1
    ci, cf, cr = ix("iaid"), ix("iaspeakingtime"), ix("iarecordname")
    cn, cd = ix("ianumber"), ix("iacalldate")
    if ci < 0:
        return {}
    out = {}
    for r in linhas:
        ids = [x.strip() for x in str(r[ci] or "").split("|")]
        fs = [x.strip() for x in str(r[cf] if cf >= 0 and r[cf] is not None else "").split("|")]
        rs = [x.strip() for x in str(r[cr] if cr >= 0 and r[cr] is not None else "").split("|")]
        ns = [re.sub(r"\D", "", x)[-11:] for x in str(r[cn] if cn >= 0 and r[cn] is not None else "").split("|")]
        ds = [x.strip() for x in str(r[cd] if cd >= 0 and r[cd] is not None else "").split("|")]
        for k, i in enumerate(ids):
            if not i:
                continue
            out[i] = {"dur": segundos(fs[k]) if k < len(fs) else 0, "rec": rs[k] if k < len(rs) else "", "num": ns[k] if k < len(ns) else "", "quando": ds[k] if k < len(ds) else ""}
    return out


def dia_iso(quando):
    """'07/10/2026 14:16:03' ou '2026-10-07 11:21:59' -> '2026-10-07' (vazio se não entender)."""
    m = re.match(r"^(\d{2})/(\d{2})/(\d{4})", str(quando or "").strip())
    if m:
        return "%s-%s-%s" % (m[3], m[2], m[1])
    m = re.match(r"^(\d{4}-\d{2}-\d{2})", str(quando or "").strip())
    return m[1] if m else ""


def quando_br(quando):
    """Mostra sempre dd/mm/aaaa hh:mm:ss, venha a data como vier."""
    q = str(quando or "").strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})(.*)$", q)
    return "%s/%s/%s%s" % (m[3], m[2], m[1], m[4]) if m else q


def indexar_pasta(pasta):
    arquivos = []
    for raiz, _, nomes in os.walk(pasta):
        for n in nomes:
            if n.lower().endswith((".mp3", ".wav", ".ogg", ".m4a", ".mp4", ".flac", ".opus", ".webm", ".aac", ".gsm")):
                arquivos.append(os.path.join(raiz, n))
    return arquivos


def achar_na_pasta(arquivos, id_, rec):
    codigo = rec.rsplit("-", 1)[-1] if rec else ""  # ex.: 2026/10/05/5531973371050-Cs825vHQiv -> Cs825vHQiv
    base = os.path.basename(rec) if rec else ""
    for a in arquivos:
        nome = os.path.basename(a)
        if id_ in nome or (base and base in nome) or (len(codigo) >= 8 and codigo in nome):
            return a
    return None


def baixar(id_, token, destino, bases):
    """Usa o SEU token (nunca cookies do navegador). Tenta cada domínio e cada forma de enviar o token;
    se o 3C responder 401/403, é falta de permissão da conta: o script não tenta contornar."""
    caminho = URL_GRAVACAO.format(id=urllib.parse.quote(id_))
    ultimo = None
    for base in bases:
        url = base.rstrip("/") + caminho
        for u, cab in ((url, {"Authorization": "Bearer " + token}), (url + "?api_token=" + urllib.parse.quote(token), {})):
            try:
                req = urllib.request.Request(u, headers=dict(cab, **{"User-Agent": "finx-transcrever"}))
                with urllib.request.urlopen(req, timeout=60) as resp:
                    dados = resp.read()
                if len(dados) < 1000:
                    raise ValueError("resposta pequena demais para ser áudio")
                with open(destino, "wb") as f:
                    f.write(dados)
                return destino
            except urllib.error.HTTPError as e:
                ultimo = "HTTP %s em %s" % (e.code, base)
            except (urllib.error.URLError, ValueError) as e:
                ultimo = "%s em %s" % (e, base)
    raise RuntimeError("não consegui baixar (%s). Se for 401/403, a sua conta não tem permissão: peça ao administrador do 3C." % ultimo)


def carregar_modelo(nome):
    try:
        from faster_whisper import WhisperModel
    except ImportError:
        sys.exit("Falta instalar a transcrição: pip install faster-whisper")
    try:
        return WhisperModel(nome, device="auto", compute_type="default")
    except Exception:
        return WhisperModel(nome, device="cpu", compute_type="int8")


def transcrever(modelo, caminho):
    segmentos, _ = modelo.transcribe(caminho, language="pt", vad_filter=True, beam_size=1, condition_on_previous_text=False)
    return " ".join(s.text.strip() for s in segmentos).strip()


def ja_feitas(saida):
    if not os.path.exists(saida):
        return set()
    with open(saida, encoding="utf-8-sig", newline="") as f:
        return {r["id"] for r in csv.DictReader(f, delimiter=";") if r.get("id")}


def main():
    ap = argparse.ArgumentParser(description="Transcreve as ligações da IA para o painel FINX.")
    ap.add_argument("--csv", help="CSV da telefonia (Relatório de Ligações do 3C Plus)")
    ap.add_argument("--xlsx", help="planilha agente IA perdidos por motivo (opcional)")
    ap.add_argument("--token", help="o SEU token de API do 3C Plus, para baixar as gravações")
    ap.add_argument("--base", help="domínio do 3C da sua empresa (ex.: https://finx.3c.plus); por padrão tenta app.3c.plus e finx.3c.plus")
    ap.add_argument("--pasta", help="pasta com áudios já baixados")
    ap.add_argument("--ordem", choices=["dia", "fala"], default="dia", help="dia (padrão): hoje primeiro, depois ontem, anteontem…; dentro de cada dia, as ligações mais longas primeiro. fala: só pela duração, de todos os dias")
    ap.add_argument("--min", type=int, default=10, help="mínimo de segundos de fala (padrão 10)")
    ap.add_argument("--max", type=int, default=0, help="limitar a quantidade (0 = todas)")
    ap.add_argument("--modelo", default="small")
    ap.add_argument("--saida", default="transcricoes.csv")
    ap.add_argument("--audios", default="gravacoes_baixadas", help="onde guardar os áudios baixados")
    ap.add_argument("--listar", type=int, default=0, metavar="N", help="não transcreve: lista as N ligações mais longas que AINDA NÃO têm áudio na --pasta (telefone, data e duração), para você baixar no 3C; salva em faltam_baixar.csv")
    ap.add_argument("--simular", action="store_true", help="não transcreve de verdade (teste do fluxo)")
    a = ap.parse_args()

    if not a.csv and not a.xlsx:
        sys.exit("Informe --csv e/ou --xlsx.")
    for arq in (a.csv, a.xlsx):
        if arq and not os.path.isfile(arq):
            sys.exit("Arquivo não encontrado: %s" % arq)
    if a.pasta and not os.path.isdir(a.pasta):
        sys.exit("Pasta não encontrada: %s" % a.pasta)
    if not a.token and not a.pasta and not a.simular and not a.listar:
        sys.exit("Informe --token (baixar do 3C Plus) ou --pasta (áudios já baixados).")

    ligacoes = {}
    if a.xlsx:
        ligacoes.update(ler_xlsx(a.xlsx))
    if a.csv:
        for k, v in ler_csv(a.csv).items():
            ant = ligacoes.get(k, {})
            ligacoes[k] = {"dur": max(v["dur"], ant.get("dur", 0)), "rec": v["rec"] or ant.get("rec", ""), "num": v["num"] or ant.get("num", ""), "quando": v["quando"] or ant.get("quando", "")}

    feitas = ja_feitas(a.saida)
    fila = sorted(((k, v) for k, v in ligacoes.items() if v["dur"] >= a.min and k not in feitas), key=(lambda kv: (-int(dia_iso(kv[1].get("quando")).replace("-", "") or 0), -kv[1]["dur"])) if a.ordem == "dia" else (lambda kv: -kv[1]["dur"]))
    if a.max:
        fila = fila[: a.max]
    total_s = sum(v["dur"] for _, v in fila)
    print("%d ligações lidas · %d já transcritas · %d na fila (>= %d s de fala, cerca de %.1f h de conversa)" % (len(ligacoes), len(feitas), len(fila), a.min, total_s / 3600))
    if not fila:
        print("Nada a fazer.")
        return

    arquivos = indexar_pasta(a.pasta) if a.pasta else []
    if a.listar:
        faltam = [(k, v) for k, v in fila if not (a.pasta and achar_na_pasta(arquivos, k, v["rec"]))][: a.listar]
        print("\nPara baixar no 3C (filtre pelo telefone e clique na setinha de download). Ordem: %s" % ("hoje primeiro, depois ontem, anteontem…; em cada dia, as mais longas primeiro" if a.ordem == "dia" else "das mais longas para as mais curtas"))
        with open("faltam_baixar.csv", "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["ordem", "telefone", "data_hora", "fala", "id"])
            for n, (k, v) in enumerate(faltam, 1):
                fala = "%d:%02d" % (v["dur"] // 60, v["dur"] % 60)
                w.writerow([n, v.get("num", ""), quando_br(v.get("quando", "")), fala, k])
                print("  %3d. %s  %s  fala %s" % (n, v.get("num", "?"), quando_br(v.get("quando", "")), fala))
        print("\nLista salva em faltam_baixar.csv. Já com áudio na pasta: %d de %d." % (len(fila) - len([1 for k, v in fila if not (a.pasta and achar_na_pasta(arquivos, k, v["rec"]))]), len(fila)))
        return
    if a.token:
        os.makedirs(a.audios, exist_ok=True)
    modelo = None if a.simular else carregar_modelo(a.modelo)

    novo = not os.path.exists(a.saida)
    inicio, feitas_agora, falhas = time.time(), 0, 0
    with open(a.saida, "a", encoding="utf-8-sig" if novo else "utf-8", newline="") as f:
        w = csv.writer(f, delimiter=";", quoting=csv.QUOTE_ALL)
        if novo:
            w.writerow(["id", "duracao_s", "texto"])
        for n, (id_, v) in enumerate(fila, 1):
            try:
                if a.simular:
                    texto = "(simulação) transcrição da ligação %s" % id_[-6:]
                else:
                    caminho = achar_na_pasta(arquivos, id_, v["rec"]) if a.pasta else None
                    if not caminho and a.token:
                        destino = os.path.join(a.audios, id_ + ".mp3")
                        caminho = destino if os.path.exists(destino) else baixar(id_, a.token, destino, ([a.base] if a.base else []) + BASES)
                    if not caminho:
                        raise RuntimeError("áudio não encontrado na pasta")
                    texto = transcrever(modelo, caminho)
                w.writerow([id_, v["dur"], texto])
                f.flush()
                feitas_agora += 1
            except KeyboardInterrupt:
                print("\nInterrompido. Rode de novo para continuar de onde parou.")
                return
            except Exception as e:  # uma ligação com problema não derruba as outras
                falhas += 1
                print("  [%d/%d] %s: %s" % (n, len(fila), id_, e))
                if falhas >= 5 and feitas_agora == 0:
                    sys.exit("5 falhas seguidas sem nenhum sucesso: confira o token / a pasta.")
                continue
            dec = time.time() - inicio
            resta = dec / n * (len(fila) - n)
            if sys.stdout.isatty() or n % 10 == 0 or n == len(fila):
                print("  [%d/%d] ok · faltam ~%d min" % (n, len(fila), resta / 60), end="\r" if sys.stdout.isatty() else "\n")
    print("\nPronto: %d transcritas, %d com problema. Arquivo: %s" % (feitas_agora, falhas, a.saida))
    print("Agora solte esse arquivo no painel (aba IA Natasha).")


if __name__ == "__main__":
    main()
