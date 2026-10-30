# Documentação técnica — CUESTA DO SEU JEITO · RT

**Versão:** 1.0 Beta (`src/branding.py` → `APP_VERSION`)

---

## 1. Arquitetura

```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────────┐
│ desktop_main.py │     │ app.py           │     │ scripts/*.ps1       │
│ (CustomTkinter) │     │ (Streamlit)      │     │ build, publish      │
└────────┬────────┘     └────────┬─────────┘     └─────────────────────┘
         │                       │
         └───────────┬───────────┘
                     ▼
         ┌───────────────────────┐
         │ src/roteiro_service   │  gerar_roteiro_da_planilha()
         └───────────┬───────────┘
                     │
     ┌───────────────┼───────────────┐
     ▼               ▼               ▼
┌─────────┐   ┌──────────────┐  ┌────────────┐
│ excel_io│   │itinerary_eng.│  │ group_scores│
│ profile │   │ build_itinerary│  │ fator, médias│
│ inventory│  └──────┬───────┘  └────────────┘
└─────────┘          │
                     ▼
              ┌──────────────┐
              │ geo_distance │  haversine, expand Maps URL
              │ maps_links   │
              └──────────────┘
                     │
         ┌───────────┴───────────┐
         ▼                       ▼
  ┌─────────────┐         ┌──────────────┐
  │ pdf_export  │         │ anamnese_pdf │
  │ + attraction│         │              │
  │   _images   │         └──────────────┘
  └─────────────┘
```

**Fonte de verdade:** planilha Excel de sessão (duas abas principais).

| Aba | Constante | Conteúdo |
|-----|-----------|----------|
| Perfil do Turista (Grupos) | `PROFILE_SHEET` | Campos + scores Plog/Cohen |
| Inventário Turístico | `INVENTORY_SHEET` | `Attraction` rows |

---

## 2. Módulos (`src/`)

| Módulo | Responsabilidade |
|--------|------------------|
| `app_paths.py` | Raiz do projeto / `_MEIPASS` (PyInstaller), `data/`, `output/` |
| `branding.py` | Nomes, versão, títulos de janela/PDF |
| `app_icon.py` | `logo_cuesta.ico` gerado a partir de `logo_cuesta.png` |
| `session_store.py` | Persiste caminho da última planilha (`ultima_sessao.json`) |
| `form_constants.py` | Opções de formulário (blank, scores 1–5) |
| `excel_io.py` | Leitura/escrita Excel, `TouristProfile`, `Attraction`, templates |
| `group_scores.py` | `GroupScores`, texto de análise técnica |
| `itinerary_engine.py` | Filtragem, agenda diária, `ItineraryResult`, markdown |
| `geo_distance.py` | `GeoContext`, coordenadas, minutos de viagem |
| `maps_links.py` | Sanitização URL, `expand_maps_url` (cache LRU), links PDF |
| `maps_static_thumbnail.py` | Static Maps / OSM para miniaturas no PDF |
| `attraction_images.py` | Integração thumbnail ↔ bloco de agenda |
| `pdf_export.py` | Layout ReportLab do roteiro |
| `anamnese_pdf.py` | Ficha + termo (PageBreak) |
| `roteiro_service.py` | Orquestra validação + `build_itinerary` |

---

## 3. Fluxo de geração do roteiro

1. `read_profile(workbook)` → `TouristProfile`
2. `validate_profile_for_roteiro` — lista campos obrigatórios
3. `read_inventory` + `filter_inventory_by_cities`
4. `GroupScores.from_profile` — médias e fator restritivo
5. `build_itinerary(profile, inventory, scores)`

### 3.1 Filtragem de atrativos

- Pesos Plog/Cohen vs médias do grupo (margem ±1 conforme fator).
- Capacidade vs `numero_pessoas` quando fator alto.
- Categorias: gastronomia → almoço; hospedagem → sugestão de pernoite, fora da timeline de visitas.

### 3.2 Agenda (`build_itinerary`)

Por dia:

1. `day_start_reference` — centróide das cidades do perfil ou média das coords.
2. Manhã até 12h — `_schedule_visits_phase` com `stop_before=lunch_cutoff`.
3. Bloco almoço (1h30) — `_pick_lunch_gastronomy` ou pausa genérica.
4. Tarde — mesma fase com `start_from=lunch_attr`.
5. Marcador de encerramento / tarde livre.
6. Sugestão de hospedagem (18h, bloco informativo).

**Primeira visita:** `DAY_LATEST_FIRST_START = 09:30`; candidatos devem abrir até 9h30 e iniciar visita até 9h30.

**Deslocamento:** `GeoContext.travel_minutes` — haversine entre `coords_resolved` (URL Maps expandida ou centróide de cidade). Fórmula em `travel_minutes_between` (~45 km/h + 10 min, clamp 15–120 min).

**Ordem:** `_schedule_from_pool` ordena candidatos por `_geo_sort_key` (distância + pesos).

### 3.3 Links Maps

`expand_maps_url` (HTTP redirect, timeout 12s, `@lru_cache(maxsize=512)`) em `maps_links.py`.  
Parsing: padrões `@lat,lng`, `!3d!4d`, `q=lat,lng` em `geo_distance._parse_coords_from_url`.

---

## 4. PDF

### Roteiro (`build_pdf`)

- `SimpleDocTemplate` A4, chrome em `_draw_page_chrome`
- Timeline por `_timeline_row` — cores por `kind` (visit, transit, lunch, end, lodging)
- Thumbnails: `attraction_thumbnail_flowable` → `fetch_map_thumbnail_bytes`

### Anamnese (`build_anamnese_pdf`)

- Cabeçalho custom `_draw_anamnese_page_chrome` (grupo no header)
- Termo após `PageBreak`
- Checkboxes: `_CHECK = "(&nbsp;&nbsp;&nbsp;)"` no ReportLab

---

## 5. Desktop (`desktop_main.py`)

- **Thread** para geração de PDF do roteiro (evita freeze UI) com botão desabilitado + status.
- Planilha: `initialize_session_workbook` copia template de `data/`.
- Treeview (`ttk`) para inventário.

---

## 6. Empacotamento PyInstaller

| Artefato | Função |
|----------|--------|
| `RoteiroTurismo.spec` | One-file exe, `datas=data/`, exclui streamlit |
| `build/version_info.txt` | Metadados Windows (Product version 1.0 Beta) |
| `scripts/build_desktop.ps1` | pip install, `ensure_app_icon`, pyinstaller |

Comando:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_desktop.ps1
```

**Frozen paths:** `project_root()` → `sys._MEIPASS`; gravações em `writable_root()` (pasta do `.exe`).

---

## 7. Dependências

```
streamlit>=1.32.0      # só app web
openpyxl>=3.1.2
reportlab>=4.1.0
customtkinter>=5.2.2
Pillow                   # transitivo (ícone, thumbnails)
numpy                    # transitivo (tiles OSM)
pyinstaller              # build apenas
```

---

## 8. Testes manuais sugeridos

1. Perfil completo + `Modelo_Inventario_Turistico.xlsx` com Maps → verificar deslocamentos variados no PDF.
2. Cidades sem atrativos → mensagem de erro clara.
3. Anamnese → 2 páginas, termo separado.
4. Build exe com programa fechado → ícone logo + versão nas propriedades do arquivo.

---

## 9. Pontos de extensão

| Mudança | Onde editar |
|---------|-------------|
| Nova versão | `src/branding.py`, `build/version_info.txt` |
| Nova cidade | `CIDADES_ROTEIRO`, `CITY_CENTROIDS` em `geo_distance.py` |
| Regras de horário | `itinerary_engine.py` (constantes no topo) |
| Colunas inventário | `FULL_INVENTORY_HEADERS`, parsers em `excel_io.py` |
| Texto análise técnica | `group_scores.build_technical_analysis` (nota: texto ainda menciona “30 min” genérico; tempos reais vêm de `geo_distance`) |

---

## 10. Estrutura de dados principais

```python
@dataclass
class TouristProfile:  # excel_io
    nome_grupo, numero_pessoas, duracao_dias, perfil, ritmo, cidades, plog_*, cohen_*

@dataclass
class Attraction:
    id, nome, categoria, horarios, tempo_horas, peso_plog, peso_cohen, descricao, cidade, maps_url

@dataclass
class ItineraryResult:
    profile, scores, selected_attractions, days: list[DayPlan], justification, lodging_suggestions
```

`ScheduleBlock.kind`: `visit | transit | lunch | end | lodging`.
