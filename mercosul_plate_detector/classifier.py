"""
Modulo de Classificacao de Tipos de Placa Mercosul

Classifica placas por categoria baseado em:
- Cores do fundo e caracteres
- Padroes de prefixo
- Caracteristicas visuais

Tipos de placa no Brasil (Mercosul):
- Particular: fundo branco, letras/numeros pretos
- Comercial: fundo vermelho, letras/numeros brancos
- Oficial: fundo azul, letras/numeros brancos
- Diplomatico: fundo azul, letras/numeros brancos, prefixos especiais
- Especial: fundo verde, letras/numeros brancos (taxi, escolar, etc)
- Colecionador: fundo cinza, letras/numeros pretos
"""

import numpy as np
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass
from enum import Enum

try:
    import cv2
    CV2_AVAILABLE = True
except ImportError:
    CV2_AVAILABLE = False


class PlateType(Enum):
    """Tipos de placa do Mercosul Brasil"""
    PARTICULAR = "particular"
    COMERCIAL = "comercial"
    OFICIAL = "oficial"
    DIPLOMATICO = "diplomatico"
    ESPECIAL = "especial"
    COLECIONADOR = "colecionador"
    DESCONHECIDO = "desconhecido"


@dataclass
class PlateColor:
    """Representa uma cor no espaco HSV com tolerancia"""
    name: str
    h_min: int
    h_max: int
    s_min: int
    s_max: int
    v_min: int
    v_max: int


@dataclass
class ClassificationResult:
    """Resultado da classificacao de uma placa"""
    plate_type: PlateType
    confidence: float
    detected_bg_color: str
    detected_char_color: str
    details: Dict
    alternative_types: List[Tuple[PlateType, float]]


class PlateClassifier:
    """
    Classificador de tipos de placa do Mercosul.

    Utiliza analise de cores e padroes de texto para
    determinar a categoria da placa.
    """

    # Cores de fundo das placas (HSV)
    BACKGROUND_COLORS = {
        'branco': PlateColor('branco', 0, 180, 0, 30, 200, 255),
        'azul': PlateColor('azul', 100, 130, 100, 255, 50, 255),
        'vermelho': PlateColor('vermelho', 0, 10, 100, 255, 50, 255),
        'vermelho2': PlateColor('vermelho2', 170, 180, 100, 255, 50, 255),
        'verde': PlateColor('verde', 40, 80, 100, 255, 50, 255),
        'cinza': PlateColor('cinza', 0, 180, 0, 50, 100, 200),
    }

    # Cores dos caracteres (HSV)
    CHARACTER_COLORS = {
        'preto': PlateColor('preto', 0, 180, 0, 255, 0, 50),
        'branco': PlateColor('branco', 0, 180, 0, 30, 200, 255),
    }

    # Mapeamento de cores para tipos de placa
    COLOR_TO_TYPE = {
        ('branco', 'preto'): PlateType.PARTICULAR,
        ('vermelho', 'branco'): PlateType.COMERCIAL,
        ('azul', 'branco'): PlateType.OFICIAL,  # ou DIPLOMATICO
        ('verde', 'branco'): PlateType.ESPECIAL,
        ('cinza', 'preto'): PlateType.COLECIONADOR,
    }

    # Prefixos diplomaticos conhecidos
    DIPLOMATIC_PREFIXES = [
        'CMD',  # Corpo Diplomatico
        'CC',   # Corpo Consular
        'CD',   # Corpo Diplomatico
        'OI',   # Organismo Internacional
    ]

    # Prefixos de veiculos oficiais por orgao
    OFFICIAL_PREFIXES = {
        'federal': ['PF', 'PRF', 'RF'],
        'estadual': ['PM', 'PC', 'BM', 'CB'],
        'militar': ['EB', 'MB', 'FAB'],
    }

    # Prefixos especiais (transporte publico, taxi, escolar)
    SPECIAL_PREFIXES = [
        'TX',  # Taxi
        'ESC', # Escolar
        'TUR', # Turismo
    ]

    def __init__(self, use_neural: bool = False):
        """
        Inicializa o classificador.

        Args:
            use_neural: Se True, usa classificador neural (requer modelo treinado)
        """
        self.use_neural = use_neural
        self.neural_model = None

        if use_neural:
            self._load_neural_model()

    def _load_neural_model(self):
        """Carrega o modelo neural de classificacao"""
        # Placeholder para carregamento do modelo
        pass

    def _detect_color_regions(
        self,
        image: np.ndarray,
        color: PlateColor
    ) -> np.ndarray:
        """
        Detecta regioes de uma cor especifica na imagem.

        Args:
            image: Imagem em formato BGR
            color: Cor a ser detectada

        Returns:
            Mascara binaria das regioes detectadas
        """
        if not CV2_AVAILABLE:
            raise ImportError("OpenCV nao esta instalado")

        hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)

        lower = np.array([color.h_min, color.s_min, color.v_min])
        upper = np.array([color.h_max, color.s_max, color.v_max])

        mask = cv2.inRange(hsv, lower, upper)

        return mask

    def _calculate_color_percentage(
        self,
        image: np.ndarray,
        color: PlateColor
    ) -> float:
        """
        Calcula a porcentagem de pixels de uma cor na imagem.

        Args:
            image: Imagem em formato BGR
            color: Cor a ser analisada

        Returns:
            Porcentagem de pixels (0-1)
        """
        mask = self._detect_color_regions(image, color)
        total_pixels = mask.shape[0] * mask.shape[1]
        color_pixels = np.count_nonzero(mask)

        return color_pixels / total_pixels if total_pixels > 0 else 0.0

    def detect_background_color(
        self,
        image: np.ndarray
    ) -> Tuple[str, float]:
        """
        Detecta a cor de fundo predominante da placa.

        Args:
            image: Imagem da placa em formato BGR

        Returns:
            Tupla com (nome_da_cor, confianca)
        """
        if not CV2_AVAILABLE:
            return ('desconhecido', 0.0)

        color_scores = {}

        for color_name, color in self.BACKGROUND_COLORS.items():
            # Vermelho precisa de tratamento especial (wrap no HSV)
            if color_name == 'vermelho2':
                continue

            score = self._calculate_color_percentage(image, color)
            color_scores[color_name] = score

        # Combina as duas faixas de vermelho
        if 'vermelho' in color_scores:
            vermelho2_score = self._calculate_color_percentage(
                image, self.BACKGROUND_COLORS['vermelho2']
            )
            color_scores['vermelho'] += vermelho2_score

        # Encontra a cor predominante
        if not color_scores:
            return ('desconhecido', 0.0)

        best_color = max(color_scores, key=color_scores.get)
        confidence = color_scores[best_color]

        return (best_color, confidence)

    def detect_character_color(
        self,
        image: np.ndarray,
        background_color: str
    ) -> Tuple[str, float]:
        """
        Detecta a cor dos caracteres da placa.

        Args:
            image: Imagem da placa em formato BGR
            background_color: Cor de fundo detectada

        Returns:
            Tupla com (nome_da_cor, confianca)
        """
        if not CV2_AVAILABLE:
            return ('desconhecido', 0.0)

        # Caracteres sao geralmente pretos ou brancos
        char_colors = {}

        for color_name, color in self.CHARACTER_COLORS.items():
            score = self._calculate_color_percentage(image, color)
            char_colors[color_name] = score

        # Normalmente a cor dos caracteres e oposta ao fundo
        if background_color in ['branco', 'cinza']:
            expected = 'preto'
        else:
            expected = 'branco'

        # Ajusta a confianca baseado na expectativa
        for color_name in char_colors:
            if color_name == expected:
                char_colors[color_name] *= 1.2  # Bonus para cor esperada

        best_color = max(char_colors, key=char_colors.get)
        confidence = min(1.0, char_colors[best_color])

        return (best_color, confidence)

    def check_diplomatic_pattern(self, plate_text: str) -> bool:
        """
        Verifica se a placa tem padrao diplomatico.

        Args:
            plate_text: Texto da placa normalizado

        Returns:
            True se for padrao diplomatico
        """
        if len(plate_text) < 2:
            return False

        for prefix in self.DIPLOMATIC_PREFIXES:
            if plate_text.startswith(prefix):
                return True

        return False

    def check_official_pattern(self, plate_text: str) -> Tuple[bool, Optional[str]]:
        """
        Verifica se a placa tem padrao oficial.

        Args:
            plate_text: Texto da placa normalizado

        Returns:
            Tupla com (is_official, orgao)
        """
        if len(plate_text) < 2:
            return (False, None)

        for category, prefixes in self.OFFICIAL_PREFIXES.items():
            for prefix in prefixes:
                if plate_text.startswith(prefix):
                    return (True, category)

        return (False, None)

    def check_special_pattern(self, plate_text: str) -> bool:
        """
        Verifica se a placa tem padrao especial.

        Args:
            plate_text: Texto da placa normalizado

        Returns:
            True se for padrao especial
        """
        if len(plate_text) < 2:
            return False

        for prefix in self.SPECIAL_PREFIXES:
            if plate_text.startswith(prefix):
                return True

        return False

    def classify_by_color(
        self,
        image: np.ndarray
    ) -> Tuple[PlateType, float, Dict]:
        """
        Classifica o tipo de placa baseado nas cores.

        Args:
            image: Imagem da placa em formato BGR

        Returns:
            Tupla com (tipo, confianca, detalhes)
        """
        if not CV2_AVAILABLE:
            return (PlateType.DESCONHECIDO, 0.0, {})

        bg_color, bg_conf = self.detect_background_color(image)
        char_color, char_conf = self.detect_character_color(image, bg_color)

        color_key = (bg_color, char_color)
        details = {
            'background_color': bg_color,
            'background_confidence': bg_conf,
            'character_color': char_color,
            'character_confidence': char_conf,
        }

        # Busca o tipo pela combinacao de cores
        if color_key in self.COLOR_TO_TYPE:
            plate_type = self.COLOR_TO_TYPE[color_key]
            confidence = (bg_conf + char_conf) / 2
        else:
            # Tenta match parcial
            for key, ptype in self.COLOR_TO_TYPE.items():
                if key[0] == bg_color:
                    plate_type = ptype
                    confidence = bg_conf * 0.7
                    break
            else:
                plate_type = PlateType.DESCONHECIDO
                confidence = 0.0

        return (plate_type, confidence, details)

    def classify_by_text(
        self,
        plate_text: str
    ) -> Tuple[PlateType, float, Dict]:
        """
        Classifica o tipo de placa baseado no texto.

        Args:
            plate_text: Texto da placa normalizado

        Returns:
            Tupla com (tipo, confianca, detalhes)
        """
        details = {
            'is_diplomatic': False,
            'is_official': False,
            'official_category': None,
            'is_special': False,
        }

        # Verifica padroes especiais no texto
        if self.check_diplomatic_pattern(plate_text):
            details['is_diplomatic'] = True
            return (PlateType.DIPLOMATICO, 0.9, details)

        is_official, official_cat = self.check_official_pattern(plate_text)
        if is_official:
            details['is_official'] = True
            details['official_category'] = official_cat
            return (PlateType.OFICIAL, 0.8, details)

        if self.check_special_pattern(plate_text):
            details['is_special'] = True
            return (PlateType.ESPECIAL, 0.8, details)

        # Se nao detectou padrao especial, assume particular
        return (PlateType.PARTICULAR, 0.5, details)

    def classify(
        self,
        image: Optional[np.ndarray] = None,
        plate_text: Optional[str] = None
    ) -> ClassificationResult:
        """
        Classifica o tipo de placa combinando analise de cor e texto.

        Args:
            image: Imagem da placa em formato BGR (opcional)
            plate_text: Texto da placa (opcional)

        Returns:
            ClassificationResult com tipo e detalhes
        """
        color_result = None
        text_result = None
        alternatives = []

        # Classificacao por cor
        if image is not None and CV2_AVAILABLE:
            color_type, color_conf, color_details = self.classify_by_color(image)
            color_result = (color_type, color_conf, color_details)
            if color_type != PlateType.DESCONHECIDO:
                alternatives.append((color_type, color_conf))

        # Classificacao por texto
        if plate_text:
            text_type, text_conf, text_details = self.classify_by_text(plate_text)
            text_result = (text_type, text_conf, text_details)
            if text_type != PlateType.DESCONHECIDO:
                alternatives.append((text_type, text_conf))

        # Combina os resultados
        final_type = PlateType.DESCONHECIDO
        final_confidence = 0.0
        bg_color = 'desconhecido'
        char_color = 'desconhecido'
        details = {}

        if color_result and text_result:
            # Prioriza classificacao por texto para tipos especiais
            if text_result[0] in [PlateType.DIPLOMATICO, PlateType.OFICIAL, PlateType.ESPECIAL]:
                final_type = text_result[0]
                final_confidence = (text_result[1] + color_result[1]) / 2
            # Caso contrario, usa a classificacao por cor
            elif color_result[0] != PlateType.DESCONHECIDO:
                # Se a cor indica azul mas texto nao e diplomatico/oficial
                if color_result[0] == PlateType.OFICIAL and not text_result[2].get('is_diplomatic'):
                    final_type = PlateType.OFICIAL
                else:
                    final_type = color_result[0]
                final_confidence = color_result[1]
            else:
                final_type = text_result[0]
                final_confidence = text_result[1]

            bg_color = color_result[2].get('background_color', 'desconhecido')
            char_color = color_result[2].get('character_color', 'desconhecido')
            details = {**color_result[2], **text_result[2]}

        elif color_result:
            final_type = color_result[0]
            final_confidence = color_result[1]
            bg_color = color_result[2].get('background_color', 'desconhecido')
            char_color = color_result[2].get('character_color', 'desconhecido')
            details = color_result[2]

        elif text_result:
            final_type = text_result[0]
            final_confidence = text_result[1]
            details = text_result[2]

        # Remove duplicatas das alternativas
        seen = set()
        unique_alternatives = []
        for alt in alternatives:
            if alt[0] not in seen and alt[0] != final_type:
                seen.add(alt[0])
                unique_alternatives.append(alt)

        return ClassificationResult(
            plate_type=final_type,
            confidence=final_confidence,
            detected_bg_color=bg_color,
            detected_char_color=char_color,
            details=details,
            alternative_types=unique_alternatives
        )


def classify_plate(
    image: Optional[np.ndarray] = None,
    plate_text: Optional[str] = None
) -> ClassificationResult:
    """
    Funcao de conveniencia para classificar uma placa.

    Args:
        image: Imagem da placa (opcional)
        plate_text: Texto da placa (opcional)

    Returns:
        Resultado da classificacao
    """
    classifier = PlateClassifier()
    return classifier.classify(image, plate_text)
