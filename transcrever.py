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
    --min 1         (padrão) todas as ligações com alguma fala. Use --min 10 para só as de 10 s ou mais; --min 0 para absolutamente todas.
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


DICA = "Ligação da IA Natasha, da Dryve Assinaturas de Veículos, com um cliente. Fala de motorista de aplicativo, Uber, 99, assinatura de carro, financiamento, aluguel."


def transcrever(modelo, caminho, dica=False):
    segmentos, _ = modelo.transcribe(caminho, language="pt", vad_filter=True, beam_size=1, condition_on_previous_text=False, initial_prompt=(DICA if dica else None))
    return " ".join(s.text.strip() for s in segmentos).strip()


def ja_feitas(saida):
    if not os.path.exists(saida):
        return set()
    with open(saida, encoding="utf-8-sig", newline="") as f:
        return {r["id"] for r in csv.DictReader(f, delimiter=";") if r.get("id")}


DRYVE_BASE = "https://app.dryve.pro/dashboard/opportunities/"
PII = re.compile(r"nome|name|cpf|mail|lead_?phones?$|^pedidos ?- ?lead_phone", re.I)  # nunca vão para a planilha final


def ler_transcricoes(caminho):
    if not os.path.exists(caminho):
        return {}
    with open(caminho, encoding="utf-8-sig", newline="") as f:
        return {r["id"]: r for r in csv.DictReader(f, delimiter=";") if r.get("id")}


def fmt_seg(n):
    try:
        n = int(float(n))
    except (TypeError, ValueError):
        return ""
    return "%d:%02d" % (n // 60, n % 60)


def montar_planilha(xlsx, transc, saida, todas_colunas=False):
    """Uma linha por ligação transcrita e uma por lead, com número do lead e as informações da planilha 'agente IA perdidos por motivo'
    (sem nome nem CPF do cliente). Abas: Por lead, Por ligação e Prompt sugerido para a IA."""
    try:
        import openpyxl
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("  (openpyxl não instalado: não deu para montar a planilha completa. Instale com: pip install openpyxl)")
        return None
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    ws = wb.active
    it = ws.iter_rows(values_only=True)
    cab_orig = [str(c or "") for c in next(it)]
    cab = [chave(c) for c in cab_orig]
    ix = lambda n: cab.index(n) if n in cab else -1
    I = {k: ix(k) for k in ["pedidosreferencecode", "telefone", "pedidosid", "pedidoslostreasondescription", "pedidoslostdetails", "pedidosboardstepname",
                           "pedidosorigem", "pedidosdatalead", "perdidosdataperdido", "pedidosdataperdido", "pedidoscloser", "pedidossdr",
                           "iaid", "iacalldate", "iaqualification", "iaspeakingtime", "iareadableamdstatustext", "iareadablehangupcausetext",
                           "iareadablestatustext", "iarecordname"]}
    if I["iaid"] < 0 or I["pedidosreferencecode"] < 0:
        print("  (a planilha não tem as colunas esperadas: reference_code e IA - id)")
        return None
    extras = []
    if todas_colunas:
        usadas = set(k for k, v in I.items() if v >= 0)
        extras = [i for i, c in enumerate(cab_orig) if c and not PII.search(c) and cab[i] not in usadas and not cab[i].startswith(("ia", "ura"))]

    def val(r, k):
        i = I.get(k, -1)
        return "" if i < 0 or i >= len(r) or r[i] is None else str(r[i]).strip()

    def lista(r, k):
        return [x.strip() for x in val(r, k).split("|")]

    def pega(l, k):
        return l[k] if k < len(l) else ""

    base_cols = ["Lead", "Telefone", "Link Dryve", "Motivo registrado", "Detalhe do motivo", "Etapa", "Origem", "Data do lead", "Data perdido", "Closer", "SDR"]
    ex_cols = [cab_orig[i] for i in extras]
    por_lig, por_lead = [], []
    for r in it:
        cod = val(r, "pedidosreferencecode")
        if not cod:
            continue
        uuid = val(r, "pedidosid")
        base = [cod, val(r, "telefone"), (DRYVE_BASE + uuid) if uuid else "", val(r, "pedidoslostreasondescription"), val(r, "pedidoslostdetails"),
                val(r, "pedidosboardstepname"), val(r, "pedidosorigem"), val(r, "pedidosdatalead"), val(r, "perdidosdataperdido") or val(r, "pedidosdataperdido"),
                val(r, "pedidoscloser"), val(r, "pedidossdr")] + [("" if r[i] is None else str(r[i])) for i in extras]
        ids, dts, qs, fs = lista(r, "iaid"), lista(r, "iacalldate"), lista(r, "iaqualification"), lista(r, "iaspeakingtime")
        ams, cs, sts, recs = lista(r, "iareadableamdstatustext"), lista(r, "iareadablehangupcausetext"), lista(r, "iareadablestatustext"), lista(r, "iarecordname")
        chamadas = []
        for k, cid in enumerate(ids):
            if not cid or cid not in transc:
                continue
            c = {"id": cid, "quando": quando_br(pega(dts, k)), "fala": segundos(pega(fs, k)), "qual": pega(qs, k), "amd": pega(ams, k) if len(ams) > 1 else (ams[0] if ams else ""),
                 "causa": pega(cs, k) if len(cs) > 1 else (cs[0] if cs else ""), "status": pega(sts, k) if len(sts) > 1 else (sts[0] if sts else ""),
                 "rec": pega(recs, k), "texto": (transc[cid].get("texto") or "").strip()}
            chamadas.append(c)
        if not chamadas:
            continue
        chamadas.sort(key=lambda c: (dia_iso(c["quando"]), c["quando"]))
        for c in chamadas:
            por_lig.append(base + [c["quando"], c["fala"], fmt_seg(c["fala"]), c["qual"], c["amd"], c["causa"], c["status"], c["rec"], c["id"], c["texto"][:32000]])
        maior = max(chamadas, key=lambda c: c["fala"])
        quals = [c["qual"] for c in chamadas if c["qual"] and c["qual"] != "-"]
        texto = "\n\n".join("[%s · fala %s · qualificação: %s · atendeu: %s · desligamento: %s]\n%s" % (c["quando"], fmt_seg(c["fala"]), c["qual"] or "-", c["amd"] or "-", c["causa"] or "-", c["texto"]) for c in chamadas)
        por_lead.append(base + [len(ids) if ids != [""] else 0, len(chamadas), maior["fala"], fmt_seg(maior["fala"]), maior["quando"], (maior["qual"] if maior["qual"] and maior["qual"] != "-" else (quals[0] if quals else "")), texto[:32000]])
    por_lead.sort(key=lambda x: (-int(dia_iso(x[len(base_cols) + len(ex_cols) + 4]).replace("-", "") or 0), -x[len(base_cols) + len(ex_cols) + 2]))
    por_lig.sort(key=lambda x: (-int(dia_iso(x[len(base_cols) + len(ex_cols)]).replace("-", "") or 0), -x[len(base_cols) + len(ex_cols) + 1]))

    out = openpyxl.Workbook()
    def aba(nome, cabecalho, linhas, larg_txt):
        w = out.create_sheet(nome)
        w.append(cabecalho)
        for lin in linhas:
            w.append(lin)
        for c in w[1]:
            c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="092A5E"); c.alignment = Alignment(wrap_text=True, vertical="center")
        w.freeze_panes = "B2"
        for i, c in enumerate(cabecalho, 1):
            w.column_dimensions[get_column_letter(i)].width = larg_txt if c.startswith("Transcrição") else (28 if c in ("Motivo registrado", "Detalhe do motivo", "Link Dryve") else 16)
        for row in w.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        w.auto_filter.ref = w.dimensions
        return w
    out.remove(out.active)
    aba("Por lead", base_cols + ex_cols + ["Ligações do lead", "Ligações transcritas", "Maior fala (s)", "Maior fala", "Data da maior fala", "Qualificação da IA (principal)", "Transcrição (todas as ligações do lead)"], por_lead, 90)
    aba("Por ligação", base_cols + ex_cols + ["Data/hora da ligação", "Fala (s)", "Fala", "Qualificação da IA", "Atendeu (detecção)", "Desligamento", "Situação", "Gravação", "Id da ligação", "Transcrição"], por_lig, 80)
    pr = out.create_sheet("Prompt para a IA")
    for linha in ["COMO USAR",
                  "1) Aba 'Por lead': uma linha por lead, com todas as transcrições das ligações dele (a conversa não separa quem fala: a IA Natasha fala primeiro).",
                  "2) Cole a linha (ou várias) numa IA junto com o prompt abaixo, ou anexe a planilha.",
                  "", "PROMPT SUGERIDO",
                  "Você é auditor de qualidade comercial de uma empresa de assinatura de veículos. As ligações foram feitas pela IA de voz 'Natasha', que qualifica leads. "
                  "Para CADA lead, valide se o MOTIVO REGISTRADO ('Motorista de App') está correto, com base na transcrição, e como a IA se comportou. Responda EXATAMENTE neste formato, uma linha por lead:",
                  "LEAD: <número> | CORRETO: SIM / NÃO / SEM DADOS | COMPORTAMENTO DA IA: BOA / MÉDIA / RUIM | QUALIFICAÇÃO DA IA CORRETA: SIM / NÃO (qual seria) | POR QUÊ: <1 a 3 frases> | MOTIVO SUGERIDO: <só se NÃO> | EVIDÊNCIA: <trecho curto>",
                  "",
                  "REGRA DO MOTIVO: 'Motorista de App' é correto quando o cliente CONFIRMA que trabalha como motorista de aplicativo (Uber, 99 etc.) e é por isso que não segue com a assinatura. "
                  "Não é correto quando o cliente nem chegou a falar disso (desligou logo, não atendeu, ligação caiu), demonstrou interesse ou pediu retorno, ou o motivo real é outro. "
                  "Atenção: a transcrição não separa quem fala; a pergunta 'você é motorista de aplicativo?' pode ser da IA, não do cliente."]:
        pr.append([linha])
    pr.column_dimensions["A"].width = 150
    for row in pr.iter_rows():
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    out.save(saida)
    return len(por_lead), len(por_lig)


def main():
    ap = argparse.ArgumentParser(description="Transcreve as ligações da IA para o painel FINX.")
    ap.add_argument("--csv", help="CSV da telefonia (Relatório de Ligações do 3C Plus)")
    ap.add_argument("--xlsx", help="planilha agente IA perdidos por motivo (opcional)")
    ap.add_argument("--token", help="o SEU token de API do 3C Plus, para baixar as gravações")
    ap.add_argument("--base", help="domínio do 3C da sua empresa (ex.: https://finx.3c.plus); por padrão tenta app.3c.plus e finx.3c.plus")
    ap.add_argument("--pasta", help="pasta com áudios já baixados")
    ap.add_argument("--ordem", choices=["dia", "fala"], default="dia", help="dia (padrão): hoje primeiro, depois ontem, anteontem…; dentro de cada dia, as ligações mais longas primeiro. fala: só pela duração, de todos os dias")
    ap.add_argument("--min", type=int, default=1, help="mínimo de segundos de fala (padrão 1 = todas as ligações em que alguém falou; use --min 10 para só as mais longas, ou --min 0 para absolutamente todas, inclusive as sem conversa)")
    ap.add_argument("--max", type=int, default=0, help="limitar a quantidade (0 = todas)")
    ap.add_argument("--modelo", default="small")
    ap.add_argument("--saida", default="transcricoes.csv")
    ap.add_argument("--audios", default="gravacoes_baixadas", help="onde guardar os áudios baixados")
    ap.add_argument("--listar", type=int, default=0, metavar="N", help="não transcreve: lista as N ligações mais longas que AINDA NÃO têm áudio na --pasta (telefone, data e duração), para você baixar no 3C; salva em faltam_baixar.csv")
    ap.add_argument("--planilha", default="transcricoes_completo.xlsx", help="planilha final (número do lead + informações + transcrição), gerada ao terminar")
    ap.add_argument("--so-planilha", action="store_true", help="não transcreve: só monta a planilha a partir do --saida já existente")
    ap.add_argument("--todas-colunas", action="store_true", help="inclui na planilha final todas as colunas da planilha agente IA perdidos (menos nome/CPF do cliente)")
    ap.add_argument("--dica", action="store_true", help="ajuda o programa a acertar nomes como Dryve Assinaturas, Uber, 99 (teste com o testar_audio.py --dica antes)")
    ap.add_argument("--simular", action="store_true", help="não transcreve de verdade (teste do fluxo)")
    a = ap.parse_args()

    if not a.csv and not a.xlsx:
        sys.exit("Informe --csv e/ou --xlsx.")
    for arq in (a.csv, a.xlsx):
        if arq and not os.path.isfile(arq):
            sys.exit("Arquivo não encontrado: %s" % arq)
    if a.pasta and not os.path.isdir(a.pasta):
        sys.exit("Pasta não encontrada: %s" % a.pasta)
    if not a.token and not a.pasta and not a.simular and not a.listar and not a.so_planilha:
        sys.exit("Informe --token (baixar do 3C Plus) ou --pasta (áudios já baixados).")

    if a.so_planilha:
        if not a.xlsx:
            sys.exit("--so-planilha precisa de --xlsx (planilha agente IA perdidos por motivo).")
        r = montar_planilha(a.xlsx, ler_transcricoes(a.saida), a.planilha, a.todas_colunas)
        print("Planilha: %s (%s leads, %s ligações)" % ((a.planilha,) + (r or (0, 0))) if r else "Não foi possível montar a planilha.")
        return
    ligacoes = {}
    if a.xlsx:
        ligacoes.update(ler_xlsx(a.xlsx))
    if a.csv:
        for k, v in ler_csv(a.csv).items():
            ant = ligacoes.get(k, {})
            ligacoes[k] = {"dur": max(v["dur"], ant.get("dur", 0)), "rec": v["rec"] or ant.get("rec", ""), "num": v["num"] or ant.get("num", ""), "quando": v["quando"] or ant.get("quando", "")}

    feitas = ja_feitas(a.saida)
    fila = sorted(((k, v) for k, v in ligacoes.items() if v["dur"] >= a.min and k not in feitas), key=(lambda kv: (-int(dia_iso(kv[1].get("quando")).replace("-", "") or 0), -kv[1]["dur"])) if a.ordem == "dia" else (lambda kv: -kv[1]["dur"]))
    if a.max and not (a.pasta and not a.listar):
        fila = fila[: a.max]
    total_s = sum(v["dur"] for _, v in fila)
    print("%d ligações lidas · %d já transcritas · %d na fila (>= %d s de fala, cerca de %.1f h de conversa)" % (len(ligacoes), len(feitas), len(fila), a.min, total_s / 3600))
    if not fila:
        print("Nada a fazer.")
        return

    arquivos = indexar_pasta(a.pasta) if a.pasta else []
    if a.pasta and not a.listar:
        # com a pasta de áudios, a fila passa a ser só o que TEM áudio (as mais longas/de hoje primeiro entre elas)
        com_audio = [(k, v) for k, v in fila if achar_na_pasta(arquivos, k, v["rec"])]
        print("%d áudios na pasta; %d deles casam com ligações da fila (as demais ligações ficam para quando você baixar mais)." % (len(arquivos), len(com_audio)))
        fila = com_audio[: a.max] if a.max else com_audio
        if not fila:
            sys.exit("Nenhum áudio da pasta casa com as ligações. Confira se os nomes são os originais do 3C (aaaa_mm_dd_telefone-código.mp3).")
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
                    texto = transcrever(modelo, caminho, a.dica)
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
    if a.xlsx:
        r = montar_planilha(a.xlsx, ler_transcricoes(a.saida), a.planilha, a.todas_colunas)
        if r:
            print("Planilha completa para a IA analisar: %s (%d leads, %d ligações). Nela: número do lead, telefone, link do Dryve, motivo, qualificação, fala, desligamento e a transcrição." % (a.planilha, r[0], r[1]))


if __name__ == "__main__":
    main()
