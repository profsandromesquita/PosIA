"""
Modulo de Deteccao de Adulteracao de Placas

Detecta sinais de adulteracao em placas do Mercosul atraves de:
- Analise de consistencia de fontes
- Deteccao de bordas irregulares
- Analise de textura e cor
- Verificacao de proporcoes e alinhamento
- Deteccao de sobreposicoes e colagens
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


class TamperingType(Enum):
    """Tipos de adulteracao detectaveis"""
    NONE = "nenhuma"
    CHAR_REPLACEMENT = "substituicao_caractere"
    STICKER_OVERLAY = "adesivo_sobreposto"
    PAINT_ALTERATION = "alteracao_pintura"
    FONT_INCONSISTENCY = "inconsistencia_fonte"
    ALIGNMENT_ISSUE = "problema_alinhamento"
    EDGE_TAMPERING = "adulteracao_bordas"
    COLOR_INCONSISTENCY = "inconsistencia_cor"
    TEXTURE_ANOMALY = "anomalia_textura"
    UNKNOWN = "desconhecida"


@dataclass
class TamperingIndicator:
    """Indicador individual de adulteracao"""
    type: TamperingType
    confidence: float
    location: Optional[Tuple[int, int, int, int]]  # x, y, w, h
    description: str
    severity: str  # 'baixa', 'media', 'alta'


@dataclass
class TamperingResult:
    """Resultado completo da analise de adulteracao"""
    is_tampered: bool
    overall_confidence: float
    tampering_score: float  # 0-1, quanto maior mais adulterado
    indicators: List[TamperingIndicator]
    details: Dict
    recommendations: List[str]


class TamperingDetector:
    """
    Detector de adulteracao em placas do Mercosul.

    Utiliza multiplas tecnicas de analise de imagem para
    detectar sinais de adulteracao:
    - Analise de frequencia (FFT)
    - Deteccao de bordas
    - Analise de histograma de cores
    - Deteccao de contornos irregulares
    - Analise de textura (LBP, GLCM)
    """

    # Proporcoes padrao da placa Mercosul (mm)
    PLATE_WIDTH_MM = 400
    PLATE_HEIGHT_MM = 130
    CHAR_HEIGHT_MM = 63
    CHAR_WIDTH_AVG_MM = 45

    # Limites de deteccao
    EDGE_THRESHOLD = 100
    COLOR_VARIANCE_THRESHOLD = 30
    ALIGNMENT_TOLERANCE = 5  # pixels

    def __init__(self, sensitivity: str = 'media'):
        """
        Inicializa o detector.

        Args:
            sensitivity: Nivel de sensibilidade ('baixa', 'media', 'alta')
        """
        self.sensitivity = sensitivity
        self._set_thresholds()

    def _set_thresholds(self):
        """Define os limites baseado na sensibilidade"""
        if self.sensitivity == 'baixa':
            self.edge_threshold = 150
            self.color_var_threshold = 50
            self.alignment_tolerance = 10
        elif self.sensitivity == 'alta':
            self.edge_threshold = 50
            self.color_var_threshold = 15
            self.alignment_tolerance = 3
        else:  # media
            self.edge_threshold = 100
            self.color_var_threshold = 30
            self.alignment_tolerance = 5

    def _preprocess_image(self, image: np.ndarray) -> np.ndarray:
        """
        Pre-processa a imagem para analise.

        Args:
            image: Imagem BGR

        Returns:
            Imagem pre-processada
        """
        if not CV2_AVAILABLE:
            return image

        # Redimensiona para tamanho padrao mantendo proporcao
        target_width = 400
        aspect_ratio = image.shape[0] / image.shape[1]
        target_height = int(target_width * aspect_ratio)

        resized = cv2.resize(image, (target_width, target_height))

        return resized

    def detect_edge_irregularities(
        self,
        image: np.ndarray
    ) -> List[TamperingIndicator]:
        """
        Detecta irregularidades nas bordas dos caracteres.

        Args:
            image: Imagem da placa

        Returns:
            Lista de indicadores de adulteracao
        """
        if not CV2_AVAILABLE:
            return []

        indicators = []

        # Converte para escala de cinza
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

        # Aplica Canny edge detection
        edges = cv2.Canny(gray, 50, 150)

        # Encontra contornos
        contours, _ = cv2.findContours(
            edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
        )

        for i, contour in enumerate(contours):
            area = cv2.contourArea(contour)
            perimeter = cv2.arcLength(contour, True)

            if perimeter == 0:
                continue

            # Calcula circularidade (regularidade da forma)
            circularity = 4 * np.pi * area / (perimeter * perimeter)

            # Aproxima o contorno
            epsilon = 0.02 * perimeter
            approx = cv2.approxPolyDP(contour, epsilon, True)

            # Verifica se o contorno tem forma muito irregular
            if len(approx) > 15 and circularity < 0.3:
                x, y, w, h = cv2.boundingRect(contour)

                # Ignora contornos muito pequenos
                if w < 10 or h < 10:
                    continue

                indicators.append(TamperingIndicator(
                    type=TamperingType.EDGE_TAMPERING,
                    confidence=0.6,
                    location=(x, y, w, h),
                    description=f"Borda irregular detectada na regiao ({x}, {y})",
                    severity='media'
                ))

        return indicators

    def detect_color_inconsistencies(
        self,
        image: np.ndarray
    ) -> List[TamperingIndicator]:
        """
        Detecta inconsistencias de cor entre caracteres.

        Args:
            image: Imagem da placa

        Returns:
            Lista de indicadores de adulteracao
        """
        if not CV2_AVAILABLE:
            return []

        indicators = []

        # Converte para HSV
        hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)

        # Divide a imagem em regioes (aproximadamente onde estariam os caracteres)
        h, w = image.shape[:2]
        num_regions = 7  # 7 caracteres na placa Mercosul

        region_width = w // num_regions
        color_stats = []

        for i in range(num_regions):
            x_start = i * region_width
            x_end = (i + 1) * region_width
            region = hsv[:, x_start:x_end]

            # Calcula estatisticas de cor
            mean_h = np.mean(region[:, :, 0])
            mean_s = np.mean(region[:, :, 1])
            mean_v = np.mean(region[:, :, 2])
            std_v = np.std(region[:, :, 2])

            color_stats.append({
                'region': i,
                'mean_h': mean_h,
                'mean_s': mean_s,
                'mean_v': mean_v,
                'std_v': std_v,
                'x_start': x_start,
                'x_end': x_end
            })

        # Compara regioes adjacentes
        for i in range(len(color_stats) - 1):
            curr = color_stats[i]
            next_r = color_stats[i + 1]

            # Calcula diferenca de valores
            v_diff = abs(curr['mean_v'] - next_r['mean_v'])
            s_diff = abs(curr['mean_s'] - next_r['mean_s'])

            # Se a diferenca e muito grande, pode indicar adulteracao
            if v_diff > self.color_var_threshold or s_diff > self.color_var_threshold:
                indicators.append(TamperingIndicator(
                    type=TamperingType.COLOR_INCONSISTENCY,
                    confidence=min(1.0, (v_diff + s_diff) / 100),
                    location=(curr['x_start'], 0, region_width * 2, h),
                    description=f"Inconsistencia de cor entre caracteres {i+1} e {i+2}",
                    severity='alta' if v_diff > 50 else 'media'
                ))

        return indicators

    def detect_texture_anomalies(
        self,
        image: np.ndarray
    ) -> List[TamperingIndicator]:
        """
        Detecta anomalias de textura que podem indicar adesivos ou colagens.

        Args:
            image: Imagem da placa

        Returns:
            Lista de indicadores de adulteracao
        """
        if not CV2_AVAILABLE:
            return []

        indicators = []

        # Converte para escala de cinza
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

        # Aplica filtro Laplaciano para detectar variacao de textura
        laplacian = cv2.Laplacian(gray, cv2.CV_64F)
        laplacian_var = laplacian.var()

        # Calcula Local Binary Pattern simplificado
        h, w = gray.shape
        block_size = 20
        texture_values = []

        for y in range(0, h - block_size, block_size):
            for x in range(0, w - block_size, block_size):
                block = gray[y:y+block_size, x:x+block_size]
                block_var = np.var(block)
                texture_values.append({
                    'x': x, 'y': y,
                    'variance': block_var
                })

        if not texture_values:
            return indicators

        # Calcula media e desvio padrao das texturas
        variances = [t['variance'] for t in texture_values]
        mean_var = np.mean(variances)
        std_var = np.std(variances)

        # Detecta blocos com textura muito diferente
        for t in texture_values:
            z_score = abs(t['variance'] - mean_var) / (std_var + 1e-6)

            if z_score > 2.5:  # Outlier significativo
                indicators.append(TamperingIndicator(
                    type=TamperingType.TEXTURE_ANOMALY,
                    confidence=min(1.0, z_score / 4),
                    location=(t['x'], t['y'], block_size, block_size),
                    description=f"Anomalia de textura detectada em ({t['x']}, {t['y']})",
                    severity='alta' if z_score > 4 else 'media'
                ))

        return indicators

    def detect_font_inconsistencies(
        self,
        char_images: List[np.ndarray]
    ) -> List[TamperingIndicator]:
        """
        Detecta inconsistencias de fonte entre caracteres.

        Args:
            char_images: Lista de imagens de caracteres individuais

        Returns:
            Lista de indicadores de adulteracao
        """
        if not CV2_AVAILABLE or len(char_images) < 2:
            return []

        indicators = []

        # Calcula caracteristicas de cada caractere
        char_features = []
        for i, char_img in enumerate(char_images):
            if char_img is None or char_img.size == 0:
                continue

            gray = cv2.cvtColor(char_img, cv2.COLOR_BGR2GRAY) \
                if len(char_img.shape) == 3 else char_img

            # Caracteristicas basicas
            h, w = gray.shape
            aspect_ratio = w / h if h > 0 else 0

            # Contagem de pixels
            _, binary = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY_INV)
            pixel_density = np.count_nonzero(binary) / (w * h) if w * h > 0 else 0

            # Momentos de Hu (invariantes a escala e rotacao)
            moments = cv2.moments(binary)
            hu_moments = cv2.HuMoments(moments).flatten()

            char_features.append({
                'index': i,
                'aspect_ratio': aspect_ratio,
                'pixel_density': pixel_density,
                'hu_moments': hu_moments
            })

        if len(char_features) < 2:
            return indicators

        # Compara caracteristicas entre caracteres
        densities = [f['pixel_density'] for f in char_features]
        mean_density = np.mean(densities)
        std_density = np.std(densities)

        for f in char_features:
            if std_density > 0:
                z_score = abs(f['pixel_density'] - mean_density) / std_density

                if z_score > 2.0:
                    indicators.append(TamperingIndicator(
                        type=TamperingType.FONT_INCONSISTENCY,
                        confidence=min(1.0, z_score / 3),
                        location=None,
                        description=f"Caractere {f['index']+1} com fonte inconsistente",
                        severity='alta' if z_score > 3 else 'media'
                    ))

        return indicators

    def detect_alignment_issues(
        self,
        char_boxes: List[Tuple[int, int, int, int]]
    ) -> List[TamperingIndicator]:
        """
        Detecta problemas de alinhamento entre caracteres.

        Args:
            char_boxes: Lista de bounding boxes dos caracteres (x, y, w, h)

        Returns:
            Lista de indicadores de adulteracao
        """
        if len(char_boxes) < 2:
            return []

        indicators = []

        # Calcula linha base (y + h de cada caractere)
        baselines = [y + h for x, y, w, h in char_boxes]
        toplines = [y for x, y, w, h in char_boxes]

        mean_baseline = np.mean(baselines)
        mean_topline = np.mean(toplines)

        # Verifica desvio de cada caractere
        for i, (x, y, w, h) in enumerate(char_boxes):
            baseline = y + h
            topline = y

            baseline_dev = abs(baseline - mean_baseline)
            topline_dev = abs(topline - mean_topline)

            if baseline_dev > self.alignment_tolerance or topline_dev > self.alignment_tolerance:
                max_dev = max(baseline_dev, topline_dev)
                indicators.append(TamperingIndicator(
                    type=TamperingType.ALIGNMENT_ISSUE,
                    confidence=min(1.0, max_dev / (self.alignment_tolerance * 2)),
                    location=(x, y, w, h),
                    description=f"Caractere {i+1} desalinhado ({max_dev:.1f}px)",
                    severity='alta' if max_dev > self.alignment_tolerance * 2 else 'media'
                ))

        return indicators

    def detect_overlay_patterns(
        self,
        image: np.ndarray
    ) -> List[TamperingIndicator]:
        """
        Detecta padroes de sobreposicao (adesivos, fitas).

        Args:
            image: Imagem da placa

        Returns:
            Lista de indicadores de adulteracao
        """
        if not CV2_AVAILABLE:
            return []

        indicators = []

        # Converte para escala de cinza
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

        # Aplica threshold adaptativo
        thresh = cv2.adaptiveThreshold(
            gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY, 11, 2
        )

        # Detecta retangulos que podem ser adesivos
        contours, _ = cv2.findContours(
            thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE
        )

        image_area = image.shape[0] * image.shape[1]

        for contour in contours:
            area = cv2.contourArea(contour)

            # Ignora contornos muito pequenos ou muito grandes
            if area < image_area * 0.01 or area > image_area * 0.3:
                continue

            # Verifica se e aproximadamente retangular
            x, y, w, h = cv2.boundingRect(contour)
            rect_area = w * h
            rectangularity = area / rect_area if rect_area > 0 else 0

            # Se for muito retangular, pode ser um adesivo
            if rectangularity > 0.85:
                # Verifica se a regiao tem borda definida
                roi = gray[y:y+h, x:x+w]
                edges = cv2.Canny(roi, 50, 150)
                edge_density = np.count_nonzero(edges) / (w * h) if w * h > 0 else 0

                if edge_density > 0.1:  # Muitas bordas na regiao
                    indicators.append(TamperingIndicator(
                        type=TamperingType.STICKER_OVERLAY,
                        confidence=min(1.0, rectangularity * edge_density * 10),
                        location=(x, y, w, h),
                        description=f"Possivel adesivo detectado em ({x}, {y})",
                        severity='alta'
                    ))

        return indicators

    def analyze(
        self,
        image: np.ndarray,
        char_images: Optional[List[np.ndarray]] = None,
        char_boxes: Optional[List[Tuple[int, int, int, int]]] = None
    ) -> TamperingResult:
        """
        Realiza analise completa de adulteracao.

        Args:
            image: Imagem da placa
            char_images: Lista de imagens de caracteres (opcional)
            char_boxes: Lista de bounding boxes dos caracteres (opcional)

        Returns:
            TamperingResult com todas as analises
        """
        if not CV2_AVAILABLE:
            return TamperingResult(
                is_tampered=False,
                overall_confidence=0.0,
                tampering_score=0.0,
                indicators=[],
                details={'error': 'OpenCV nao disponivel'},
                recommendations=['Instale OpenCV: pip install opencv-python']
            )

        # Pre-processa a imagem
        processed = self._preprocess_image(image)

        # Coleta todos os indicadores
        all_indicators = []

        # Analise de bordas
        edge_indicators = self.detect_edge_irregularities(processed)
        all_indicators.extend(edge_indicators)

        # Analise de cor
        color_indicators = self.detect_color_inconsistencies(processed)
        all_indicators.extend(color_indicators)

        # Analise de textura
        texture_indicators = self.detect_texture_anomalies(processed)
        all_indicators.extend(texture_indicators)

        # Analise de sobreposicao
        overlay_indicators = self.detect_overlay_patterns(processed)
        all_indicators.extend(overlay_indicators)

        # Analise de fonte (se imagens de caracteres fornecidas)
        if char_images:
            font_indicators = self.detect_font_inconsistencies(char_images)
            all_indicators.extend(font_indicators)

        # Analise de alinhamento (se bounding boxes fornecidos)
        if char_boxes:
            align_indicators = self.detect_alignment_issues(char_boxes)
            all_indicators.extend(align_indicators)

        # Calcula scores finais
        if not all_indicators:
            return TamperingResult(
                is_tampered=False,
                overall_confidence=0.9,
                tampering_score=0.0,
                indicators=[],
                details={'analysis_complete': True},
                recommendations=[]
            )

        # Score baseado em quantidade e severidade dos indicadores
        high_severity = sum(1 for i in all_indicators if i.severity == 'alta')
        medium_severity = sum(1 for i in all_indicators if i.severity == 'media')
        low_severity = sum(1 for i in all_indicators if i.severity == 'baixa')

        tampering_score = min(1.0, (
            high_severity * 0.3 +
            medium_severity * 0.15 +
            low_severity * 0.05
        ))

        # Confianca geral baseada na consistencia dos indicadores
        confidences = [i.confidence for i in all_indicators]
        overall_confidence = np.mean(confidences) if confidences else 0.0

        # Determina se a placa foi adulterada
        is_tampered = tampering_score > 0.3 or high_severity >= 2

        # Gera recomendacoes
        recommendations = []
        if is_tampered:
            recommendations.append("Verificacao manual recomendada")

            if high_severity > 0:
                recommendations.append("Indicadores de alta severidade detectados")

            tampering_types = set(i.type for i in all_indicators)
            if TamperingType.STICKER_OVERLAY in tampering_types:
                recommendations.append("Verificar presenca de adesivos sobrepostos")
            if TamperingType.COLOR_INCONSISTENCY in tampering_types:
                recommendations.append("Verificar uniformidade de cor dos caracteres")
            if TamperingType.FONT_INCONSISTENCY in tampering_types:
                recommendations.append("Verificar consistencia da fonte dos caracteres")

        # Monta detalhes
        details = {
            'total_indicators': len(all_indicators),
            'high_severity_count': high_severity,
            'medium_severity_count': medium_severity,
            'low_severity_count': low_severity,
            'tampering_types_detected': [i.type.value for i in all_indicators],
            'analysis_sensitivity': self.sensitivity
        }

        return TamperingResult(
            is_tampered=is_tampered,
            overall_confidence=overall_confidence,
            tampering_score=tampering_score,
            indicators=all_indicators,
            details=details,
            recommendations=recommendations
        )


def detect_tampering(
    image: np.ndarray,
    sensitivity: str = 'media'
) -> TamperingResult:
    """
    Funcao de conveniencia para detectar adulteracao.

    Args:
        image: Imagem da placa
        sensitivity: Nivel de sensibilidade

    Returns:
        Resultado da analise
    """
    detector = TamperingDetector(sensitivity=sensitivity)
    return detector.analyze(image)
