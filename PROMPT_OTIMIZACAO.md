# Prompt de otimização geral — Connect (Central de Atendimento)

> Cole o texto abaixo no Claude Code **dentro do repositório que tem o código-fonte do app Electron** (o `projeto-connect` só guarda releases, o painel `finx_monitoria.html` e o `auth-callback.html`).

---

## Contexto

Você vai otimizar o **Connect** ("Central de Atendimento"), um app desktop Windows em **Electron** que reúne **vários WhatsApp Web em uma janela só**. O app é instalado por usuário em `%LOCALAPPDATA%\Programs\Central de Atendimento\Connect\Connect.exe` (AppUserModelID `com.qualidade.centralwa`), usa login Microsoft (Entra ID / SSO, com redirect para `auth-callback.html`) e carrega o painel `finx_monitoria.html` (painel de auditorias e metas, JS puro, dados em `localStorage`).

**Problema:** o app trava e fica pesado nas máquinas dos usuários, que são variadas e muitas vezes modestas (4 GB de RAM, CPU de entrada, HDD, antivírus ativo). O objetivo é que ele rode **liso em qualquer máquina**, sem perder funcionalidade.

## Regras do trabalho

1. **Meça antes de mexer.** Primeiro crie uma linha de base e só depois otimize. Sem número antes/depois, a otimização não conta.
2. **Uma mudança por vez, por tema**, com commit separado e mensagem dizendo o ganho medido.
3. **Não quebre nada:** login SSO, notificações, badge de mensagens não lidas, sessões persistentes dos WhatsApps (o usuário não pode ter que escanear QR de novo), auto-update, atalho e instalação por usuário.
4. **Não enfraqueça a segurança** para ganhar velocidade (mantenha `contextIsolation: true`, `nodeIntegration: false`, `sandbox` onde já estiver, CSP, validação de IPC).
5. Se algo for arriscado ou mudar comportamento visível, **pare e pergunte** em vez de decidir sozinho.
6. Teste também em modo "máquina fraca" (limite CPU/RAM, ex.: VM com 2 vCPU/4 GB ou throttling de CPU 4x no DevTools).

## Fase 1 — Diagnóstico (entregue um relatório antes de alterar código)

- Mapeie a arquitetura: processo main, preload, como cada WhatsApp é carregado (`<webview>`, `BrowserView`/`WebContentsView`, janelas), partições de sessão, IPC, timers, listeners, logs, auto-update.
- Meça e registre numa tabela (ideal: 1, 3, 6 e 10 contas abertas):
  - tempo de abertura até a janela utilizável e até a primeira conta pronta;
  - RAM total (somando todos os processos do Electron, via `app.getAppMetrics()`);
  - CPU em repouso e durante troca de conta;
  - uso de disco/IO na inicialização;
  - tamanho do instalador, do `app.asar` e de `node_modules` empacotados;
  - tempo de resposta da UI (long tasks > 50 ms no Performance do DevTools).
- Rode `electron --inspect` / `chrome://tracing` / `--trace-startup` e aponte os **10 maiores gargalos**, em ordem de impacto.

## Fase 2 — Otimizações a investigar (priorize pelo que o diagnóstico confirmar)

**A. Multi-WhatsApp (maior impacto esperado)**
- **Carregamento preguiçoso:** só abrir a conta quando o usuário clicar nela; no início, abrir só a última ativa. Opção de "carregar em sequência" com intervalo para não estourar CPU no boot.
- **Hibernação de contas inativas:** descarregar (destruir o `webContents`, mantendo a partição/sessão) contas sem uso há X minutos e sem notificações pendentes; recarregar ao clicar. Deixar X configurável e um modo "economia" automático em máquinas fracas.
- **Background throttling:** `webContents.setBackgroundThrottling(true)` nas contas ocultas, `backgroundThrottling` ativo no `webPreferences`, e pausar mídia/animações das ocultas.
- Verificar se há **uma instância pesada por conta sem necessidade**: avaliar `BrowserView`/`WebContentsView` em vez de `<webview>`, e compartilhar processo de renderização quando fizer sentido (sem misturar sessões).
- Limitar `setInterval`/polling por conta (badge, título, notificação): trocar polling por observação de eventos (`MutationObserver` no título, `page-title-updated`), com debounce.
- Limpar caches de contas hibernadas e definir limite de cache em disco da sessão.

**B. Inicialização**
- Splash/janela mostrada imediatamente e conteúdo carregado depois; `show: false` + `ready-to-show`.
- Reduzir trabalho antes do `app.whenReady()`; `require` tardio de módulos pesados; evitar leitura/escrita síncrona (`fs.*Sync`) no main.
- Empacotar com **asar**, remover dependências e arquivos mortos, `devDependencies` fora do pacote, locales não usados (`electronLanguages`), minificar o código próprio, considerar V8 snapshot/bytecode cache.
- Checagem de atualização e telemetria **depois** do app utilizável, nunca bloqueando a abertura.

**C. Processo main e IPC**
- Nada de trabalho pesado no main thread: mover para `utilityProcess`/worker. Trocar `ipcRenderer.sendSync`/`remote` por `invoke` assíncrono.
- Debounce/coalescência de mensagens IPC frequentes (badge, contadores, estado de contas).
- Logs: nível `warn` em produção, rotação com tamanho máximo, escrita assíncrona e em lote.

**D. Renderização e GPU**
- Verificar aceleração por hardware; oferecer fallback automático/configurável para GPUs problemáticas (`app.disableHardwareAcceleration()` apenas como opção detectada, não padrão).
- Revisar `app.commandLine.appendSwitch` (flags úteis: limitar renderers, `disable-renderer-backgrounding` NÃO deve ser usado; evitar flags que aumentam consumo).
- CSS/UI da própria casca: evitar `backdrop-filter`, sombras e animações custosas, layout thrash, listas grandes sem virtualização; respeitar `prefers-reduced-motion` e criar um "modo leve".

**E. Memória e estabilidade**
- Caçar vazamentos: listeners não removidos, referências a `webContents` destruídos, timers órfãos, `BrowserWindow` não liberada.
- Detectar `render-process-gone`/`unresponsive` e recuperar a conta sozinho (recarregar só ela) em vez de travar o app.
- Alerta/ação automática quando a RAM do app passar de um limite (ex.: hibernar as contas menos usadas).

**F. Disco, rede e atualização**
- Auto-update com **diferencial** (blockmap), fora do horário de uso, em segundo plano, sem travar a UI.
- Reduzir gravações em disco frequentes (configurações, estado); gravar com debounce e de forma atômica.
- Excluir a pasta de dados do app da varredura do antivírus é decisão do usuário/TI: apenas **documente** a recomendação, não faça isso no código.

## Fase 3 — Painel `finx_monitoria.html` (achados já levantados, com linha aproximada)

Corrigir e medir (use um conjunto de dados grande, ex.: 20 mil auditorias e 50 mil linhas de base):

1. **`indiceLeads()` (~l.1130):** para cada auditoria faz `b.filter(...)` na base do lead → custo O(n·m). Como `b` já está ordenado por data, usar busca binária ou ponteiro incremental.
2. **`nomeOperadorAtual()` (~l.1544):** `estado.operadores.find(...)` a cada chamada, usada dentro de filtro e comparador de ordenação do histórico. Criar `Map` id→nome reconstruído só quando `estado.rev` mudar.
3. **`calcEquipe()` (~l.1292):** a cada render recalcula `normBusca(b.avaliado)` (NFD + regex) para toda a base e reconstrói o `Set` de chaves. Pré-calcular o nome normalizado ao sanitizar a base e memoizar o resultado por `estado.rev` + data.
4. **Histórico (~l.2735, `filtro()`):** cada tecla nos filtros chama `renderHistorico()` (filtra, ordena e repopula o `<select>` de operadores). Aplicar **debounce de ~150 ms**, repopular o select só quando operadores/auditorias mudarem, e cachear o resultado filtrado.
5. **Busca da base (`#b-busca`, ~l.2791):** sem debounce; `renderBase()` ainda ordena todas as datas a cada vez. Debounce e cache de período.
6. **`tick()` (~l.2801):** roda a cada 1 s e, a cada virada de minuto, refaz a aba inteira (`renderGeral`/`renderRegistrar`). Pausar quando `document.hidden`, atualizar só o relógio por segundo (e só se o texto mudou) e redesenhar a aba com `requestAnimationFrame` apenas se algum dado dependente do horário mudou.
7. **Persistência (`salvarChave`, ~l.1059):** `JSON.stringify` do array inteiro e escrita síncrona no `localStorage` a cada alteração, mais duas leituras extras em `atualizarSalvo()`. Agrupar gravações (debounce/`requestIdleCallback`), guardar o timestamp em memória e avaliar IndexedDB se os dados crescerem.
8. **Animação `.aba{animation:entrar}`:** replay a cada troca de aba; respeitar `prefers-reduced-motion` e permitir desligar no modo leve.
9. Garantir que listas longas continuem **paginadas/virtualizadas** e que o carregamento inicial não bloqueie (parse da `BASE_SEMENTE` só na primeira execução).

## Critérios de aceite

- RAM total com 6 contas abertas: **redução de pelo menos 30%** sobre a linha de base (alvo: contas inativas hibernadas ocupam quase nada).
- Janela utilizável em **< 3 s** em máquina fraca (cold start) e **< 1,5 s** (warm start).
- CPU em repouso **< 2%** com todas as contas abertas e sem atividade.
- Nenhuma long task > 100 ms na troca de conta e na navegação do painel com dados grandes.
- Zero regressão funcional: checklist manual de login SSO, troca de conta, notificação, badge, QR persistido, auto-update, instalação/atalho.
- Relatório final com tabela **antes × depois**, lista do que foi feito, do que foi descartado (e por quê) e do que ficou como sugestão futura.

## Entrega

Trabalhe em uma branch própria, com commits pequenos por tema (A…F, painel), não faça merge nem publique release sem aprovação, e ao final apresente o relatório antes/depois e as instruções para gerar e testar o instalador.
