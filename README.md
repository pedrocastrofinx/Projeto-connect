# Projeto-connec
Releases do Central de Atendimento

## monitoring-rescue.html

Arquivo único (abre direto no navegador, sem servidor e sem internet) com dois módulos:

- **Resgate de Monitoria** — cruzamento Ligações × Kanban (inalterado).
- **Cadência de Ligações** — usa só a base de ligações. Avalia cada operador + data + lead:
  cadência seguida com 3 ou mais ligações no dia, ou com 1 ligação `effective` = 1 acima de 01:30
  (01:30 exato não conta). Inclui visão geral, ranking por operador (volume mínimo configurável),
  leads por dia/semana/mês, detalhamento com linha do tempo, oportunidades de correção,
  relatório para e-mail, imagem PNG do ranking e exportações XLSX/CSV/HTML.
