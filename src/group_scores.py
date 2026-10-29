from __future__ import annotations

from dataclasses import dataclass

from src.excel_io import TouristProfile, compute_group_scores


@dataclass
class GroupScores:
    """Espelha a seção 4 da planilha — SCORE DO GRUPO E FATORES RESTRITIVOS."""

    media_plog: float
    media_cohen: float
    fator_restritivo: str

    @classmethod
    def from_profile(cls, profile: TouristProfile) -> GroupScores:
        raw = compute_group_scores(profile)
        return cls(
            media_plog=raw["media_plog"],
            media_cohen=raw["media_cohen"],
            fator_restritivo=raw["fator_restritivo"],
        )


def _perfil_psicografico_label(plog: float, cohen: float) -> str:
    """Rótulo descritivo alinhado ao roteiro modelo (mid-centric etc.)."""

    def eixo(v: float) -> str:
        if v < 2.0:
            return "centrado"
        if v <= 3.0:
            return "mid-centric"
        return "explorador"

    p, c = eixo(plog), eixo(cohen)
    if p == c:
        return f'perfil "{p}"'
    return f"perfil Plog {p} / Cohen {c}"


def build_technical_analysis(scores: GroupScores, selected_weights: list[tuple[float, float]]) -> str:
    weights_txt = ", ".join(f"{p:.1f}/{c:.1f}" for p, c in selected_weights[:6])
    label = _perfil_psicografico_label(scores.media_plog, scores.media_cohen)
    return (
        f"O roteiro gerado atende ao nível psicográfico de aventura e independência do grupo ({label}). "
        f"As atrações selecionadas respeitam margem ±1 em relação à Média Plog ({scores.media_plog:.2f}) "
        f"e Média Cohen ({scores.media_cohen:.2f}), considerando o fator restritivo "
        f"\"{scores.fator_restritivo}\" (tamanho do grupo). "
        f"Pesos Plog/Cohen das atrações: {weights_txt or 'n/d'}. "
        "A logística respeita almoço entre 12h00 e 13h30 (1h30), deslocamentos de 30 minutos "
        "e encerramento diário até 18h00."
    )
