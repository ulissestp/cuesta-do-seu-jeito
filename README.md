# Roteiro Turístico — Projeto Integrador

Formulário **desktop** (recomendado) ou web (Streamlit), **sem IA**.

## Versão desktop (Windows)

### Executar direto (desenvolvimento)

Duplo clique em **`Iniciar_Roteiro.bat`** ou:

```powershell
cd C:\Users\USI-SEVEN\Developer\roteiro-turismo
pip install -r requirements.txt
python desktop_main.py
```

### Gerar executável (.exe)

```powershell
cd C:\Users\USI-SEVEN\Developer\roteiro-turismo
powershell -ExecutionPolicy Bypass -File scripts\build_desktop.ps1
```

O arquivo ficará em `dist\RoteiroTurismo.exe`. Copie a pasta `dist` ou apenas o `.exe` (a planilha modelo já vem embutida).

## Etapas do programa

1. **Perfil** — seções 1–3 + score automático (seção 4).
2. **Inventário** — importar `.xlsx`.
3. **Roteiro PDF** — gera PDF com cabeçalho, quadro-resumo e linha do tempo.

## Versão web (opcional)

```powershell
streamlit run app.py
```

## Lógica do roteiro

- **Média Plog** / **Média Cohen** / **fator restritivo** (seção 4).
- Horários de funcionamento de cada atração.
- **Modelo de inventário:** `data/Modelo_Inventario_Turistico.xlsx` (ID, Atração, Abre, Fecha, Cap., Plog, Cohen).
- Descrição breve no roteiro e no PDF.
- Almoço 12h–13h30, deslocamento 30 min, fim às 18h.

Saídas em `output/`. Modelo: `data/Projeto_Integrador_Turismo_Final.xlsx`.

O **PDF final** inclui cabeçalho personalizado, quadro-resumo do cliente e linha do tempo por dia.
