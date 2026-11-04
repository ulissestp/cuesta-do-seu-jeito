# Ajuda — CUESTA DO SEU JEITO · RT

Guia rápido dentro do aplicativo. Manual completo: **Manual do usuário** (PDF ou Markdown no menu **Ajuda**).

---

## Fluxo recomendado

1. **Perfil** — dados do grupo, calendário (check-in / check-out), cidades, questionários Plog e Cohen. Clique em **Gravar na planilha**.
2. **Inventário** — importe o `.xlsx` (modelo em `data\Modelo_Inventario_Turistico.xlsx` ou botão **Salvar modelo**).
3. **Roteiro PDF** — **Gerar roteiro e PDF**; use **Abrir PDF** ou a pasta `output`.
4. **Ficha anamnese** — PDF para impressão (anamnese + termo em páginas separadas).

---

## Barra superior

| Ação | Efeito |
|------|--------|
| **Limpar informações** | Apaga o perfil e o roteiro calculado; **mantém** o inventário importado. |
| **Nova planilha** | Cria outra sessão Excel (`Projeto_Integrador_*.xlsx` em `output`). |
| Nome do arquivo | Planilha de sessão atual (perfil + inventário). |

---

## Aba 1 — Perfil

- **Plog** e **Cohen**: cada pergunta precisa de nota **1 a 5** (não deixe em branco) para gerar o roteiro.
- **Cidades**: marque ao menos uma (Botucatu, Pardinho, Bofete). Só entram atrativos dessas cidades.
- **Ritmo**: Relaxado (2 visitas/dia), Moderado (4) ou Intenso (6) — almoço não conta.
- Após gravar, confira **Média Plog**, **Média Cohen** e **fator restritivo** (tamanho do grupo).

---

## Aba 2 — Inventário

- Colunas essenciais: horários por dia da semana e feriado, **Cidade**, **Maps** (Google Maps), **Imagem** (foto no PDF), **Ingresso** (valor opcional).
- **Verificar atualização** (aba Inventário): confere **programa (.exe)** no GitHub Releases e **inventário** (planilha remota). Mescla só diferenças (novos/alterados) e remove duplicatas (mesmo ID ou nome+cidade). Não roda ao abrir o app.
- Depois de gerar um roteiro, linhas usadas ficam destacadas; em **Novo ID** você pode trocar um destino e gerar de novo.
- Sem links **Maps**, deslocamentos ficam menos precisos (o programa avisa na geração do PDF).

---

## Aba 3 — Roteiro PDF

- Exige perfil completo e inventário compatível com as cidades.
- Regras resumidas: primeira visita por volta de **9h30**, almoço a partir de **12h00** (1h30), fim até **18h00**, deslocamento máximo **40 min**.
- PDFs ficam em `output\` com nome `Roteiro_NomeDoGrupo_HHMM.pdf`.

---

## Aba 4 — Ficha anamnese

- Gera `Ficha_Anamnese_*.pdf` em `output\` para preenchimento manual do grupo.

---

## Pasta `output`

| Arquivo | Conteúdo |
|---------|----------|
| `Projeto_Integrador_*.xlsx` | Sessão (perfil + inventário) |
| `Roteiro_*.pdf` | Roteiros gerados |
| `Ficha_Anamnese_*.pdf` | Fichas de anamnese |
| `ultima_sessao.json` | Reabre a última planilha na próxima execução |

---

## Problemas comuns

| Situação | Solução |
|----------|---------|
| Mensagem de perfil incompleto | Preencha todas as notas Plog/Cohen e **Gravar na planilha**. |
| Inventário vazio | Importe o `.xlsx` na aba 2. |
| Nenhuma atração para as cidades | Confira a coluna **Cidade** no Excel (ex.: `Botucatu - SP`). |
| PDF lento na primeira vez | Normal ao resolver muitos links do Maps; mantenha internet. |

---

## Atalhos (desktop)

- **F1** — abre esta ajuda.
- Menu **Ajuda** — manual em PDF, manual Markdown e pasta de documentação.

Versão do programa: ver título da janela ou rodapé dos PDFs.
