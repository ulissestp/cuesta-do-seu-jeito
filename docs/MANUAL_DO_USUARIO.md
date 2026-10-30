# Manual do usuário — CUESTA DO SEU JEITO · RT

**Versão do programa:** 1.0 Beta  
**Produto:** gerador de roteiros turísticos personalizados (sem inteligência artificial).

---

## 1. O que o programa faz

O **CUESTA DO SEU JEITO - RT** monta roteiros dia a dia a partir de:

1. **Perfil do grupo** — dados gerais, cidades, questionários **Plog** (aventura) e **Cohen** (independência).
2. **Inventário turístico** — lista de atrativos com horários, pesos, descrição e link do Google Maps.
3. **Motor de roteiro** — escolhe atrativos compatíveis com o perfil, ordena deslocamentos e monta horários.
4. **Saídas em PDF** — roteiro para o cliente e ficha de anamnese (termo em página separada).

Tudo fica gravado numa **planilha Excel de sessão** (`Projeto_Integrador_*.xlsx` na pasta `output`).

---

## 2. Como instalar e abrir

### Opção A — Executável (recomendado)

1. Use `dist\RoteiroTurismo.exe`.
2. Na primeira execução, o programa cria uma planilha de sessão em `output\` (mesma pasta do `.exe`, se você copiou só o executável).
3. Ícone e versão **1.0 Beta** aparecem na janela e nas propriedades do arquivo no Windows.

### Opção B — Python (equipe técnica)

```powershell
pip install -r requirements.txt
python desktop_main.py
```

---

## 3. Visão geral da interface

Quatro abas, na ordem sugerida de uso:

| Aba | Função |
|-----|--------|
| **1. Perfil** | Dados do grupo, cidades, Plog, Cohen; gravar na planilha |
| **2. Inventário** | Importar `.xlsx` com atrativos; visualizar tabela |
| **3. Roteiro PDF** | Gerar roteiro, pré-visualizar texto, abrir PDF |
| **4. Ficha anamnese** | PDF de anamnese + termo (página 2) |

**Barra superior**

- **Limpar informações** — zera o perfil na planilha (não apaga o inventário importado).
- **Nova planilha** — inicia outra sessão Excel.
- Nome do arquivo de sessão atual (canto direito).

---

## 4. Aba 1 — Perfil do turista

### 4.1 Dados gerais

Preencha:

- Nome do grupo / contato  
- Número de pessoas  
- Duração da viagem (dias)  
- Perfil predominante (Casais, Famílias, etc.)  
- Ritmo de viagem (Relaxado, Moderado, Intenso)

### 4.2 Cidades do roteiro

Marque **ao menos uma** entre:

- Botucatu  
- Pardinho  
- Bofete  

Somente atrativos cuja **cidade** (coluna do inventário) corresponder às cidades marcadas entram no roteiro.

### 4.3 Eixos Plog e Cohen

Para cada pergunta, escolha nota **de 1 a 5** (textos de ajuda aparecem abaixo de cada item).

Clique em **Gravar na planilha** antes de gerar o roteiro.

### 4.4 Seção 4 — Score automático

Após gravar, o programa exibe:

- **Média Plog** e **Média Cohen**  
- **Fator restritivo** (ligado ao tamanho do grupo: baixo / médio / alto)

Esses valores orientam quais atrativos são elegíveis e aparecem no PDF.

---

## 5. Aba 2 — Inventário turístico

### Importar

1. **Importar inventário (.xlsx)** — escolha um arquivo com aba de inventário (modelo em `data\Modelo_Inventario_Turistico.xlsx`).
2. Os dados são copiados para a planilha de sessão.
3. A tabela na tela lista ID, nome, horários, capacidade, Plog, Cohen.

### Colunas importantes (modelo completo)

| Coluna | Descrição |
|--------|-----------|
| ID | Identificador único |
| Nome da Atração | Título no roteiro |
| Categoria | Tema do dia (parcial) |
| Horário de Abertura / Fechamento | Respeitados na programação |
| Tempo Médio (horas) | Duração da visita |
| Peso Plog / Cohen | Compatibilidade com o grupo |
| Descrição Breve | Texto no PDF |
| **Cidade** | Filtro Botucatu / Pardinho / Bofete |
| **Maps** | Link Google Maps (recomendado) |

**Salvar modelo de inventário** exporta um arquivo vazio para você preencher.

### Links do Google Maps

- Use links do Maps (de preferência **maps.app.goo.gl** ou URL completa).
- Com **internet**, o programa expande links curtos e calcula **distância e tempo** entre atrativos.
- Sem link ou sem rede, usa aproximação por cidade (menos preciso).

---

## 6. Aba 3 — Roteiro PDF

### Antes de gerar

O programa exige perfil **completo** e inventário **não vazio** para as cidades escolhidas. Se faltar algo, uma mensagem lista os campos pendentes.

### Gerar

1. **Gerar roteiro e PDF** — calcula o roteiro e grava o PDF (pode levar alguns segundos na primeira vez, enquanto resolve links do Maps).
2. **Abrir PDF** — último roteiro gerado.
3. **Abrir pasta de saída** — pasta `output`.

Nome típico: `Roteiro_NomeDoGrupo_HHMM.pdf`.

### Conteúdo do PDF do roteiro

- Capa com logo, versão **1.0 Beta**, cliente e duração  
- Quadro-resumo do perfil e índices  
- **Linha do tempo por dia** — visitas, deslocamentos (km e minutos), almoço, encerramento  
- Miniatura de **mapa** por bloco (Google Static Maps se houver chave API; senão OpenStreetMap)  
- Sugestões de **hospedagem** (não entram na grade horária)  
- Análise técnica (Plog, Cohen, fator restritivo)

### Regras de horário (resumo)

| Regra | Valor |
|-------|--------|
| Primeira visita do dia | Início **até 9h30** (atractivo deve abrir até esse horário) |
| Almoço | Janela a partir de **12h00**, duração **1h30** |
| Fim das atividades | Até **18h00** |
| Deslocamento | Calculado pela **distância** entre coordenadas dos links Maps (~45 km/h + margem) |
| Visitas por dia | Depende do ritmo: Relaxado 2 · Moderado 4 · Intenso 6 (almoço não conta) |

A ordem das visitas no dia considera proximidade geográfica e compatibilidade de scores.

---

## 7. Aba 4 — Ficha anamnese

Gera PDF para preenchimento / arquivo do grupo:

- Dados clínicos e checkboxes para marcar à mão  
- **Termo** em **página separada**  
- Cabeçalho com logo e **Grupo: [nome]**  
- Rodapé com versão do programa  

Arquivo: `Ficha_Anamnese_NomeDoGrupo_HHMM.pdf`.

---

## 8. Pasta de saída (`output`)

| Conteúdo | Descrição |
|----------|-----------|
| `Projeto_Integrador_*.xlsx` | Sessão (perfil + inventário) |
| `Roteiro_*.pdf` | Roteiros gerados |
| `Ficha_Anamnese_*.pdf` | Anamneses |
| `ultima_sessao.json` | Lembra qual planilha abrir na próxima execução |

---

## 9. Configuração opcional — miniaturas de mapa no PDF

Para mapas estáticos do Google (melhor qualidade):

1. Copie `data/google_maps_api_key.txt.example` para `google_maps_api_key.txt` na pasta `data` (ou ao lado do `.exe` em `dist\data` se aplicável).
2. Cole a chave da API **Google Maps Static Maps** (uma linha, sem `#`).
3. Sem chave, o programa usa mapa OpenStreetMap na coordenada do atrativo.

---

## 10. Solução de problemas

| Problema | O que fazer |
|----------|-------------|
| “Preencha todos os campos…” | Complete aba 1 e clique **Gravar na planilha** |
| “Inventário vazio” | Importe inventário na aba 2 |
| “Nenhuma atração para as cidades…” | Confira coluna **Cidade** no Excel (ex.: “Botucatu - SP”) |
| Deslocamentos iguais ou 10–15 min | Inclua links **Maps** válidos; mantenha internet na 1ª geração |
| PDF demora ou trava | Normal na 1ª vez com muitos links; aguarde ou reduza atrativos |
| Não consigo recompilar o `.exe` | Feche `RoteiroTurismo.exe` antes do build |
| Ícone antigo no atalho | Windows pode cachear; use o `.exe` em `dist\` ou recrie o atalho |

---

## 11. Contato e evolução

Versão atual: **1.0 Beta**. Novas versões alteram principalmente `APP_VERSION` em `src/branding.py` e o instalador desktop.

Para dúvidas sobre regras de negócio (Plog/Cohen, cidades, capacidade), alinhe com a equipe pedagógica / produto do projeto integrador.
