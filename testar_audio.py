#!/usr/bin/env python3
"""
Teste rápido: transcreve UM áudio e mostra o texto na tela (não gera arquivo, não usa CSV).

    python testar_audio.py audios\\2026_10_08_5516993571179-tUWO75Xrgr.mp3

Na primeira vez baixa o modelo de voz (~500 MB, precisa de internet). Depois roda offline.
Opcional: --modelo tiny|base|small|medium  (small é o padrão)
          --dica   ajuda o programa a acertar nomes como "Dryve Assinaturas", Uber, 99 (compare com e sem)
"""
import os
import sys
import time


DICA = "Ligação da IA Natasha, da Dryve Assinaturas de Veículos, com um cliente. Fala de motorista de aplicativo, Uber, 99, assinatura de carro, financiamento, aluguel."


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    usar_dica = "--dica" in sys.argv
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
    gpu = False
    try:
        import ctranslate2
        gpu = ctranslate2.get_cuda_device_count() > 0
    except Exception:
        pass
    try:
        m = WhisperModel(modelo, device="cuda", compute_type="float16") if gpu else WhisperModel(modelo, device="cpu", compute_type="int8", cpu_threads=os.cpu_count() or 4)
    except Exception:
        m = WhisperModel(modelo, device="cpu", compute_type="int8")
    print("Rodando na placa de vídeo." if gpu else "Rodando no processador (int8).")
    t1 = time.time()
    print("Modelo pronto em %.0f s. Transcrevendo %s ..." % (t1 - t0, args[0]))
    segs, info = m.transcribe(args[0], language="pt", vad_filter=True, beam_size=1, condition_on_previous_text=False, initial_prompt=(DICA if usar_dica else None))
    texto = []
    for s in segs:
        print("[%5.1fs] %s" % (s.start, s.text.strip()))
        texto.append(s.text.strip())
    t2 = time.time()
    print("\n--- RESULTADO ---")
    print("Áudio: %.0f s · transcrição levou %.0f s (%.1fx o tempo real) · %d caracteres" % (info.duration, t2 - t1, (t2 - t1) / max(info.duration, 1), len(" ".join(texto))))


if __name__ == "__main__":
    main()
