"""
Agentes do sistema de geração de relatórios V3A.

Orquestrador  → valida input e coordena o fluxo
PPT Builder   → gera o PPTX (chama ppt_builder.py)
Diretor de Arte → revisa estrutura e estética via Claude API
"""
import os
import json
from dataclasses import dataclass, field


# ── Mensagens entre agentes ──────────────────────────────────────────────────

@dataclass
class Message:
    sender: str
    receiver: str
    payload: dict


@dataclass
class AgentResult:
    success: bool
    data: dict = field(default_factory=dict)
    error: str = ""


# ── Orquestrador ─────────────────────────────────────────────────────────────

class OrchestratorAgent:
    """
    Ponto de entrada do pipeline.
    Valida input → aciona Builder → aciona Diretor de Arte → devolve resultado.
    """

    def __init__(self):
        self.builder = PPTBuilderAgent()
        self.art_director = ArtDirectorAgent()

    def process(self, subtitle: str, pieces: list[dict]) -> AgentResult:
        print("[Orquestrador] Iniciando pipeline")

        # Validação básica
        validation = self._validate(subtitle, pieces)
        if not validation.success:
            return validation

        print(f"[Orquestrador] Input válido — {len(pieces)} peça(s)")

        # Aciona o Builder
        build_result = self.builder.build(subtitle, pieces)
        if not build_result.success:
            return build_result

        pptx_bytes: bytes = build_result.data["pptx_bytes"]
        report: dict = build_result.data["report"]

        print("[Orquestrador] PPTX gerado, enviando para Diretor de Arte")

        # Aciona o Diretor de Arte (até 2 tentativas)
        for attempt in range(1, 3):
            review = self.art_director.review(
                report=report,
                n_pieces=len(pieces),
                subtitle=subtitle,
            )

            if review.data.get("approved"):
                print(f"[Orquestrador] Aprovado pelo Diretor de Arte (tentativa {attempt})")
                return AgentResult(
                    success=True,
                    data={
                        "pptx_bytes": pptx_bytes,
                        "art_director_message": review.data.get("message", "Aprovado."),
                        "issues": [],
                    },
                )

            issues = review.data.get("issues", [])
            print(f"[Orquestrador] Reprovado (tentativa {attempt}): {issues}")

            if attempt < 2:
                # Tenta corrigir e rebuildar (por enquanto, loga e tenta de novo)
                build_result = self.builder.build(subtitle, pieces, hints=issues)
                if not build_result.success:
                    return build_result
                pptx_bytes = build_result.data["pptx_bytes"]
                report = build_result.data["report"]

        # Após 2 tentativas sem aprovação, entrega com aviso
        return AgentResult(
            success=True,
            data={
                "pptx_bytes": pptx_bytes,
                "art_director_message": "Entregue com ressalvas. Verifique os itens abaixo.",
                "issues": review.data.get("issues", []),
            },
        )

    def _validate(self, subtitle: str, pieces: list) -> AgentResult:
        errors = []
        if not subtitle or not subtitle.strip():
            errors.append("Subtítulo da capa não pode estar vazio.")
        if not pieces:
            errors.append("Informe pelo menos uma peça.")
        for i, p in enumerate(pieces):
            if not p.get("image_path"):
                errors.append(f"Peça {i+1}: imagem ausente.")
            if not p.get("filename"):
                errors.append(f"Peça {i+1}: nome de arquivo ausente.")
        if errors:
            return AgentResult(success=False, error="; ".join(errors))
        return AgentResult(success=True)


# ── PPT Builder ───────────────────────────────────────────────────────────────

class PPTBuilderAgent:
    """Gera o PPTX usando ppt_builder.py (lógica determinística)."""

    def build(self, subtitle: str, pieces: list[dict], hints: list = None) -> AgentResult:
        from app.ppt_builder import build_report, inspect_report

        print(f"[Builder] Gerando PPTX — {len(pieces)} peça(s)")
        try:
            pptx_bytes = build_report(subtitle, pieces)
            report = inspect_report(pptx_bytes)
            return AgentResult(success=True, data={"pptx_bytes": pptx_bytes, "report": report})
        except Exception as e:
            return AgentResult(success=False, error=f"Erro ao gerar PPTX: {e}")


# ── Diretor de Arte ───────────────────────────────────────────────────────────

class ArtDirectorAgent:
    """
    Revisa a estrutura do PPTX gerado.
    Usa Claude API se ANTHROPIC_API_KEY estiver configurada;
    caso contrário, faz verificação determinística.
    """

    EXPECTED_FONTS = {"Geologica", "Geologica ExtraBold", "Geologica Light"}

    def review(self, report: dict, n_pieces: int, subtitle: str) -> AgentResult:
        api_key = os.getenv("ANTHROPIC_API_KEY")
        if api_key:
            return self._review_with_claude(report, n_pieces, subtitle, api_key)
        return self._review_deterministic(report, n_pieces, subtitle)

    # ── Revisão determinística (sem LLM) ─────────────────────────────────────

    def _review_deterministic(self, report: dict, n_pieces: int, subtitle: str) -> AgentResult:
        issues = []
        expected_slides = n_pieces + 3  # capa + peças + total + último

        if report["total_slides"] != expected_slides:
            issues.append(
                f"Esperado {expected_slides} slides, encontrado {report['total_slides']}."
            )

        slides = report["slides"]

        # Slide 0 — capa: deve conter parte do subtítulo
        cover_texts = " ".join(slides[0]["texts"]).upper()
        sub_check = subtitle.split("|")[0].strip().upper()
        if sub_check and sub_check not in cover_texts:
            issues.append(f"Capa não contém o subtítulo esperado '{sub_check}'.")

        # Slides 1..n_pieces — peças: devem ter imagem
        for i in range(1, n_pieces + 1):
            if i < len(slides) and not slides[i]["has_image"]:
                issues.append(f"Slide {i+1} (peça) não tem imagem.")

        # Penúltimo slide — deve conter "total"
        total_slide = slides[n_pieces + 1] if n_pieces + 1 < len(slides) else None
        if total_slide:
            total_text = " ".join(total_slide["texts"]).lower()
            if "total" not in total_text:
                issues.append("Slide de total não contém a palavra 'total'.")
            if str(n_pieces) not in total_text:
                issues.append(f"Slide de total não menciona o número {n_pieces}.")

        # Último slide — deve conter texto V3A
        last_slide = slides[-1] if slides else None
        if last_slide:
            last_text = " ".join(last_slide["texts"]).upper()
            if "V3A" not in last_text and "ATIVAÇÃO" not in last_text:
                issues.append("Último slide parece não ser o slide fixo V3A.")

        # Fontes — verifica nos slides de peça
        for i in range(1, n_pieces + 1):
            if i < len(slides):
                for font in slides[i]["fonts"]:
                    if font and not any(f.lower() in font.lower() for f in ["Geologica", "Calibri"]):
                        issues.append(f"Slide {i+1}: fonte inesperada '{font}'.")

        approved = len(issues) == 0
        message = "Aprovado" if approved else f"Revisão necessária: {'; '.join(issues)}"
        print(f"[Diretor de Arte] {'OK Aprovado' if approved else 'REPROVADO'}")
        if issues:
            for issue in issues:
                print(f"  - {issue}")

        return AgentResult(success=True, data={"approved": approved, "issues": issues, "message": message})

    # ── Revisão com Claude API ────────────────────────────────────────────────

    def _review_with_claude(self, report: dict, n_pieces: int, subtitle: str, api_key: str) -> AgentResult:
        try:
            import anthropic
            client = anthropic.Anthropic(api_key=api_key)

            prompt = f"""Você é o Diretor de Arte da agência V3A. Revise a estrutura do relatório PPTX gerado.

PARÂMETROS ESPERADOS:
- Número de peças: {n_pieces}
- Total de slides: {n_pieces + 3} (capa + {n_pieces} peça(s) + total + slide V3A)
- Subtítulo da capa deve conter: "{subtitle}"
- Todos os slides de peça devem ter imagem
- Fontes esperadas: Geologica (ExtraBold, Light, Regular)
- Penúltimo slide deve dizer "Total de {n_pieces:02d} peças"
- Último slide deve ser o slide fixo da V3A (contém "ATIVAÇÃO" e "V3A")

RELATÓRIO DO PPTX GERADO:
{json.dumps(report, ensure_ascii=False, indent=2)}

Responda APENAS com JSON válido:
{{
  "approved": true|false,
  "issues": ["lista de problemas encontrados, vazia se aprovado"],
  "message": "mensagem curta para o usuário"
}}"""

            response = client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=512,
                messages=[{"role": "user", "content": prompt}],
            )

            text = response.content[0].text.strip()
            # Extrai JSON da resposta
            json_match = re.search(r'\{.*\}', text, re.DOTALL)
            if json_match:
                result = json.loads(json_match.group())
                return AgentResult(success=True, data=result)
        except Exception as e:
            print(f"[Diretor de Arte] Erro na API Claude: {e} — usando revisão determinística")

        # Fallback para revisão determinística
        return self._review_deterministic(report, n_pieces, subtitle)


import re  # necessário para _review_with_claude
