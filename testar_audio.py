#!/usr/bin/env python3
"""
Teste rápido: transcreve UM áudio e mostra o texto na tela (não gera arquivo, não usa CSV).

    python testar_audio.py audios\\2026_10_08_5516993571179-tUWO75Xrgr.mp3

Na primeira vez baixa o modelo de voz (~500 MB, precisa de internet). Depois roda offline.
Opcional: --modelo tiny|base|small|medium  (small é o padrão)
"""
import sys
import time


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    modelo = "small"
    if "--modelo" in sys.argv:
        modelo = sys.argv[sys.argv.index("--modelo") + 1]
        args = [a for a in args if a != modelo]
    if not args:
        sys.exit("Uso: python testar_audio.py CAMINHO_DO_AUDIO.mp3")
    try:
        from faster_whisper import WhisperModel
    except ImportError:
        sys.exit("Falta instalar: python -m pip install faster-whisper")

    t0 = time.time()
    print("Carregando o modelo '%s' (na 1ª vez baixa da internet)..." % modelo)
    try:
        m = WhisperModel(modelo, device="auto", compute_type="default")
    except Exception:
        m = WhisperModel(modelo, device="cpu", compute_type="int8")
    t1 = time.time()
    print("Modelo pronto em %.0f s. Transcrevendo %s ..." % (t1 - t0, args[0]))
    segs, info = m.transcribe(args[0], language="pt", vad_filter=True, beam_size=1, condition_on_previous_text=False)
    texto = []
    for s in segs:
        print("[%5.1fs] %s" % (s.start, s.text.strip()))
        texto.append(s.text.strip())
    t2 = time.time()
    print("\n--- RESULTADO ---")
    print("Áudio: %.0f s · transcrição levou %.0f s (%.1fx o tempo real) · %d caracteres" % (info.duration, t2 - t1, (t2 - t1) / max(info.duration, 1), len(" ".join(texto))))


if __name__ == "__main__":
    main()
