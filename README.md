# CUESTA DO SEU JEITO — Roteiros Turísticos (RT)

**Versão 1.0 Beta** · Gerador de roteiros personalizados com base no perfil do grupo (Plog / Cohen), inventário em Excel e exportação em PDF.

Aplicativo **desktop para Windows** (recomendado). Versão web opcional via Streamlit. **Não utiliza IA** — a lógica é determinística e auditável.

---

## Documentação

| Documento | Público | Conteúdo |
|-----------|---------|----------|
| [Manual do usuário](docs/MANUAL_DO_USUARIO.md) | Operadores, guias, equipe comercial | Instalação, abas do programa, planilha, PDFs, dicas |
| [Manual em PDF](docs/Manual_CUESTA_RT_v1.0_Beta.pdf) | Operadores | Texto do manual + **espaços reservados** para colar capturas de tela |
| [Documentação técnica](docs/DOCUMENTACAO_TECNICA.md) | Desenvolvedores | Arquitetura, módulos, motor de roteiro, build |

---

## Início rápido (Windows)

### Executável

1. Abra `dist\RoteiroTurismo.exe` (ou copie a pasta `dist` para onde preferir).
2. Preencha **Perfil** → importe **Inventário** → gere **Roteiro PDF** e/ou **Ficha anamnese**.
3. Arquivos gerados ficam em `output\` (ao lado do `.exe` quando empacotado).

### Desenvolvimento

```powershell
cd C:\Users\USI-SEVEN\Developer\roteiro-turismo
pip install -r requirements.txt
python desktop_main.py
```

Ou duplo clique em `Iniciar_Roteiro.bat` (a partir da raiz do projeto).

### Compilar o `.exe`

Feche o programa antes de compilar.

```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_desktop.ps1
```

Saída: `dist\RoteiroTurismo.exe` (logo e versão **1.0 Beta** embutidos).

### Streamlit (opcional)

```powershell
streamlit run app.py
```

---

## Estrutura do repositório

```
roteiro-turismo/
├── desktop_main.py      # Interface desktop (CustomTkinter)
├── app.py               # Interface web (Streamlit)
├── data/                # Logo, modelos Excel, chave Maps (opcional)
├── dist/                # Executável gerado
├── docs/                # Manual e documentação técnica
├── output/              # Planilhas de sessão e PDFs (modo dev)
├── scripts/             # build_desktop.ps1, utilitários
└── src/                 # Lógica de negócio (Excel, roteiro, PDF)
```

---

## Modelos de planilha

| Arquivo | Uso |
|---------|-----|
| `data/Projeto_Integrador_Turismo_Final.xlsx` | Modelo de sessão (perfil + abas) |
| `data/Modelo_Inventario_Turistico.xlsx` | Inventário de referência (coluna **Maps** recomendada) |

---

## Requisitos

- **Windows 10/11** (desktop)
- **Python 3.11+** (desenvolvimento)
- **Internet** (opcional mas recomendada): expandir links curtos do Google Maps para calcular distâncias entre atrativos

---

## Licença e créditos

Projeto integrador — CUESTA DO SEU JEITO. Consulte a equipe responsável pelo uso comercial e distribuição.
