"""
Modulo de Validacao de Placas Mercosul

Valida o formato e autenticidade das placas do padrao Mercosul.
Formato padrao: LLLNLNN (3 letras + 1 numero + 1 letra + 2 numeros)
Exemplo: ABC1D23
"""

import re
from typing import Tuple, Dict, List, Optional
from dataclasses import dataclass
from enum import Enum


class ValidationStatus(Enum):
    """Status de validacao da placa"""
    VALID = "valida"
    SUSPECT = "suspeita"
    INVALID = "invalida"


@dataclass
class ValidationResult:
    """Resultado da validacao de uma placa"""
    status: ValidationStatus
    plate_text: str
    normalized_text: str
    confidence: float
    issues: List[str]
    details: Dict


class PlateValidator:
    """
    Validador de placas do Mercosul (Brasil).

    Caracteristicas das placas Mercosul:
    - Formato: LLLNLNN (ex: ABC1D23)
    - Letras maiusculas (A-Z, exceto acentuadas)
    - Numeros de 0-9
    - Algumas combinacoes sao restritas
    """

    # Padrao regex para placa Mercosul
    MERCOSUL_PATTERN = r'^[A-Z]{3}[0-9][A-Z][0-9]{2}$'

    # Padrao antigo brasileiro (para referencia)
    OLD_BRAZIL_PATTERN = r'^[A-Z]{3}[0-9]{4}$'

    # Letras que podem ser confundidas com numeros
    CONFUSABLE_LETTERS = {
        'O': '0', 'I': '1', 'Z': '2', 'S': '5',
        'G': '6', 'B': '8', 'Q': '0'
    }

    # Numeros que podem ser confundidos com letras
    CONFUSABLE_NUMBERS = {
        '0': 'O', '1': 'I', '2': 'Z', '5': 'S',
        '6': 'G', '8': 'B'
    }

    # Combinacoes proibidas/restritas (exemplos)
    RESTRICTED_COMBINATIONS = [
        'XXX',  # Placas de teste
        'TST',  # Placas de teste
    ]

    # Prefixos especiais por estado brasileiro (3 primeiras letras)
    STATE_PREFIXES = {
        'AC': ['NAA', 'NEZ'],
        'AL': ['OJA', 'OMZ'],
        'AP': ['NFA', 'NFZ'],
        'AM': ['NGA', 'NOZ'],
        'BA': ['NPA', 'NSZ', 'ONA', 'OSZ', 'JJA', 'JSZ'],
        'CE': ['HUA', 'HZZ', 'NTA', 'NXZ'],
        'DF': ['JHA', 'JKZ', 'OAA', 'OEZ'],
        'ES': ['MTA', 'MXZ', 'PPA', 'PTZ'],
        'GO': ['NZA', 'OIZ', 'OUA', 'OZZ'],
        'MA': ['HSA', 'HTZ', 'NYA', 'OAZ'],
        'MT': ['JKA', 'JMZ', 'PUA', 'PZZ', 'QAA', 'QDZ'],
        'MS': ['HTA', 'HUZ', 'NGA', 'NRZ'],
        'MG': ['GUA', 'HRZ', 'OFA', 'ONZ'],
        'PA': ['JTA', 'JVZ', 'QEA', 'QGZ'],
        'PB': ['MYA', 'MSZ', 'OTA', 'OUZ'],
        'PR': ['AAA', 'BEZ', 'AUA', 'BAZ'],
        'PE': ['NAA', 'NPZ', 'OVA', 'PAZ'],
        'PI': ['NQA', 'NSZ', 'OBB', 'OEZ'],
        'RJ': ['KMA', 'LVZ', 'LWA', 'LZZ'],
        'RN': ['MZA', 'NAZ', 'QHA', 'QJZ'],
        'RS': ['IIA', 'JGZ', 'IRA', 'JSZ'],
        'RO': ['NBA', 'NEZ', 'QKA', 'QMZ'],
        'RR': ['NFA', 'NFZ', 'QNA', 'QNZ'],
        'SC': ['LZA', 'MMZ', 'MMA', 'MOZ'],
        'SP': ['BFA', 'GHZ', 'CAA', 'FZZ', 'GGG', 'GTZ'],
        'SE': ['JSA', 'JTZ', 'PBA', 'PEZ'],
        'TO': ['JUA', 'JVZ', 'PFA', 'POZ'],
    }

    # Caracteres similares que indicam possivel adulteracao
    SIMILAR_CHARS = {
        'A': ['4'],
        'B': ['8', '3'],
        'C': ['(', '<'],
        'D': ['0', 'O'],
        'E': ['3'],
        'F': ['7'],
        'G': ['6', '9'],
        'H': ['#'],
        'I': ['1', 'l', '!', '|'],
        'J': [']'],
        'K': ['X'],
        'L': ['1', 'I'],
        'M': ['N'],
        'N': ['M'],
        'O': ['0', 'Q', 'D'],
        'P': ['9'],
        'Q': ['O', '0', '9'],
        'R': ['2'],
        'S': ['5', '$'],
        'T': ['7', '+'],
        'U': ['V'],
        'V': ['U', 'Y'],
        'W': ['VV'],
        'X': ['K'],
        'Y': ['V'],
        'Z': ['2'],
        '0': ['O', 'D', 'Q'],
        '1': ['I', 'l', '!', '|', 'L'],
        '2': ['Z', 'R'],
        '3': ['B', 'E', '8'],
        '4': ['A'],
        '5': ['S', '$'],
        '6': ['G', '9', 'b'],
        '7': ['T', 'F'],
        '8': ['B', '3'],
        '9': ['P', 'Q', '6', 'g'],
    }

    def __init__(self, strict_mode: bool = False):
        """
        Inicializa o validador.

        Args:
            strict_mode: Se True, aplica validacoes mais rigorosas
        """
        self.strict_mode = strict_mode
        self._compile_patterns()

    def _compile_patterns(self):
        """Compila os padroes regex para melhor performance"""
        self.mercosul_regex = re.compile(self.MERCOSUL_PATTERN)
        self.old_brazil_regex = re.compile(self.OLD_BRAZIL_PATTERN)

    def normalize_plate(self, plate_text: str) -> str:
        """
        Normaliza o texto da placa removendo caracteres invalidos.

        Args:
            plate_text: Texto bruto da placa

        Returns:
            Texto normalizado (apenas letras e numeros, maiusculas)
        """
        # Remove espacos, hifens e caracteres especiais
        normalized = re.sub(r'[^A-Za-z0-9]', '', plate_text)
        return normalized.upper()

    def is_mercosul_format(self, plate_text: str) -> bool:
        """
        Verifica se a placa esta no formato Mercosul.

        Args:
            plate_text: Texto normalizado da placa

        Returns:
            True se estiver no formato Mercosul
        """
        return bool(self.mercosul_regex.match(plate_text))

    def is_old_brazil_format(self, plate_text: str) -> bool:
        """
        Verifica se a placa esta no formato antigo brasileiro.

        Args:
            plate_text: Texto normalizado da placa

        Returns:
            True se estiver no formato antigo
        """
        return bool(self.old_brazil_regex.match(plate_text))

    def check_restricted_combinations(self, plate_text: str) -> List[str]:
        """
        Verifica se a placa contem combinacoes restritas.

        Args:
            plate_text: Texto normalizado da placa

        Returns:
            Lista de problemas encontrados
        """
        issues = []
        prefix = plate_text[:3]

        if prefix in self.RESTRICTED_COMBINATIONS:
            issues.append(f"Prefixo restrito detectado: {prefix}")

        # Verifica sequencias suspeitas
        if plate_text == plate_text[0] * len(plate_text):
            issues.append("Todos os caracteres sao iguais")

        # Verifica sequencias alfabeticas ou numericas obvias
        if plate_text[:3] in ['ABC', 'XYZ', 'QWE']:
            issues.append(f"Sequencia de letras suspeita: {plate_text[:3]}")

        if plate_text[3:] in ['1234', '0000', '9999', '1111']:
            issues.append(f"Sequencia de numeros suspeita: {plate_text[3:]}")

        return issues

    def detect_character_confusion(self, plate_text: str) -> List[str]:
        """
        Detecta possiveis confusoes entre caracteres similares.

        Args:
            plate_text: Texto da placa (pode nao estar normalizado)

        Returns:
            Lista de possiveis confusoes detectadas
        """
        issues = []

        # Posicoes esperadas: 0-2 letras, 3 numero, 4 letra, 5-6 numeros
        expected_types = ['L', 'L', 'L', 'N', 'L', 'N', 'N']

        for i, (char, expected) in enumerate(zip(plate_text.upper(), expected_types)):
            if expected == 'L' and char.isdigit():
                if char in self.CONFUSABLE_NUMBERS:
                    issues.append(
                        f"Posicao {i}: '{char}' pode ser '{self.CONFUSABLE_NUMBERS[char]}'"
                    )
            elif expected == 'N' and char.isalpha():
                if char in self.CONFUSABLE_LETTERS:
                    issues.append(
                        f"Posicao {i}: '{char}' pode ser '{self.CONFUSABLE_LETTERS[char]}'"
                    )

        return issues

    def calculate_validity_score(
        self,
        plate_text: str,
        ocr_confidence: float = 1.0
    ) -> Tuple[float, List[str]]:
        """
        Calcula um score de validade para a placa.

        Args:
            plate_text: Texto normalizado da placa
            ocr_confidence: Confianca do OCR (0-1)

        Returns:
            Tupla com (score 0-1, lista de issues)
        """
        score = 1.0
        issues = []

        # Penalidade por tamanho incorreto
        if len(plate_text) != 7:
            score -= 0.5
            issues.append(f"Tamanho incorreto: {len(plate_text)} caracteres (esperado: 7)")

        # Penalidade por formato incorreto
        if not self.is_mercosul_format(plate_text):
            score -= 0.3
            issues.append("Formato nao corresponde ao padrao Mercosul")

        # Verifica combinacoes restritas
        restriction_issues = self.check_restricted_combinations(plate_text)
        if restriction_issues:
            score -= 0.2 * len(restriction_issues)
            issues.extend(restriction_issues)

        # Detecta confusao de caracteres
        confusion_issues = self.detect_character_confusion(plate_text)
        if confusion_issues:
            score -= 0.1 * len(confusion_issues)
            issues.extend(confusion_issues)

        # Aplica confianca do OCR
        score *= ocr_confidence

        # Garante que o score esta entre 0 e 1
        score = max(0.0, min(1.0, score))

        return score, issues

    def validate(
        self,
        plate_text: str,
        ocr_confidence: float = 1.0
    ) -> ValidationResult:
        """
        Valida uma placa completa e retorna o resultado detalhado.

        Args:
            plate_text: Texto da placa (bruto ou normalizado)
            ocr_confidence: Confianca do OCR (0-1)

        Returns:
            ValidationResult com status, score e detalhes
        """
        normalized = self.normalize_plate(plate_text)
        score, issues = self.calculate_validity_score(normalized, ocr_confidence)

        # Determina o status baseado no score
        if score >= 0.8:
            status = ValidationStatus.VALID
        elif score >= 0.5:
            status = ValidationStatus.SUSPECT
        else:
            status = ValidationStatus.INVALID

        # Monta detalhes adicionais
        details = {
            'is_mercosul_format': self.is_mercosul_format(normalized),
            'is_old_format': self.is_old_brazil_format(normalized),
            'length': len(normalized),
            'expected_length': 7,
            'letters': ''.join(c for c in normalized if c.isalpha()),
            'numbers': ''.join(c for c in normalized if c.isdigit()),
        }

        return ValidationResult(
            status=status,
            plate_text=plate_text,
            normalized_text=normalized,
            confidence=score,
            issues=issues,
            details=details
        )

    def suggest_corrections(self, plate_text: str) -> List[str]:
        """
        Sugere possiveis correcoes para uma placa invalida.

        Args:
            plate_text: Texto da placa

        Returns:
            Lista de possiveis placas corretas
        """
        normalized = self.normalize_plate(plate_text)
        suggestions = []

        if len(normalized) != 7:
            return suggestions

        # Tenta corrigir confusoes comuns
        expected_types = ['L', 'L', 'L', 'N', 'L', 'N', 'N']
        corrected = list(normalized)

        for i, (char, expected) in enumerate(zip(normalized, expected_types)):
            if expected == 'L' and char.isdigit():
                # Tenta converter numero para letra
                if char in self.CONFUSABLE_NUMBERS:
                    corrected[i] = self.CONFUSABLE_NUMBERS[char]
            elif expected == 'N' and char.isalpha():
                # Tenta converter letra para numero
                if char in self.CONFUSABLE_LETTERS:
                    corrected[i] = self.CONFUSABLE_LETTERS[char]

        corrected_text = ''.join(corrected)
        if corrected_text != normalized and self.is_mercosul_format(corrected_text):
            suggestions.append(corrected_text)

        return suggestions


def validate_plate(plate_text: str, ocr_confidence: float = 1.0) -> ValidationResult:
    """
    Funcao de conveniencia para validar uma placa.

    Args:
        plate_text: Texto da placa
        ocr_confidence: Confianca do OCR

    Returns:
        Resultado da validacao
    """
    validator = PlateValidator()
    return validator.validate(plate_text, ocr_confidence)
