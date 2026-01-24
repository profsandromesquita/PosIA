"""
Modulo de Deteccao e OCR de Placas Mercosul

Responsavel por:
- Detectar placas em imagens usando YOLOv8 ou modelos customizados
- Extrair texto das placas usando OCR (EasyOCR, Tesseract)
- Integrar com modulos de validacao, classificacao e adulteracao
"""

import os
import numpy as np
from typing import Dict, List, Tuple, Optional, Union
from dataclasses import dataclass
from pathlib import Path

try:
    import cv2
    CV2_AVAILABLE = True
except ImportError:
    CV2_AVAILABLE = False

try:
    from ultralytics import YOLO
    YOLO_AVAILABLE = True
except ImportError:
    YOLO_AVAILABLE = False

try:
    import easyocr
    EASYOCR_AVAILABLE = True
except ImportError:
    EASYOCR_AVAILABLE = False

try:
    import pytesseract
    TESSERACT_AVAILABLE = True
except ImportError:
    TESSERACT_AVAILABLE = False

from .validator import PlateValidator, ValidationResult
from .classifier import PlateClassifier, ClassificationResult, PlateType
from .tampering import TamperingDetector, TamperingResult


@dataclass
class DetectionBox:
    """Bounding box de uma deteccao"""
    x1: int
    y1: int
    x2: int
    y2: int
    confidence: float
    class_id: int
    class_name: str


@dataclass
class CharacterDetection:
    """Deteccao de um caractere individual"""
    char: str
    confidence: float
    box: Tuple[int, int, int, int]  # x, y, w, h


@dataclass
class PlateDetectionResult:
    """Resultado completo da deteccao de uma placa"""
    # Deteccao
    detected: bool
    plate_image: Optional[np.ndarray]
    detection_box: Optional[DetectionBox]
    detection_confidence: float

    # OCR
    plate_text: str
    ocr_confidence: float
    characters: List[CharacterDetection]

    # Validacao
    validation: Optional[ValidationResult]

    # Classificacao
    classification: Optional[ClassificationResult]

    # Adulteracao
    tampering: Optional[TamperingResult]

    # Metadados
    processing_time_ms: float
    details: Dict


class PlateDetector:
    """
    Detector completo de placas do Mercosul.

    Combina deteccao por deep learning (YOLO) com OCR
    e analise de validacao/classificacao/adulteracao.
    """

    # Configuracoes padrao de pre-processamento
    DEFAULT_PLATE_SIZE = (400, 130)  # largura x altura
    MIN_PLATE_SIZE = (50, 15)
    MAX_PLATE_SIZE = (800, 260)

    # Caracteres validos para OCR
    ALLOWED_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

    def __init__(
        self,
        model_path: Optional[str] = None,
        ocr_engine: str = 'easyocr',
        use_gpu: bool = False,
        confidence_threshold: float = 0.5,
        enable_validation: bool = True,
        enable_classification: bool = True,
        enable_tampering_detection: bool = True
    ):
        """
        Inicializa o detector de placas.

        Args:
            model_path: Caminho para modelo YOLO customizado (opcional)
            ocr_engine: Motor de OCR ('easyocr', 'tesseract', 'both')
            use_gpu: Usar GPU para processamento
            confidence_threshold: Limite minimo de confianca para deteccao
            enable_validation: Habilitar validacao de formato
            enable_classification: Habilitar classificacao de tipo
            enable_tampering_detection: Habilitar deteccao de adulteracao
        """
        self.model_path = model_path
        self.ocr_engine = ocr_engine
        self.use_gpu = use_gpu
        self.confidence_threshold = confidence_threshold
        self.enable_validation = enable_validation
        self.enable_classification = enable_classification
        self.enable_tampering_detection = enable_tampering_detection

        # Inicializa componentes
        self.yolo_model = None
        self.ocr_reader = None
        self.validator = None
        self.classifier = None
        self.tampering_detector = None

        self._initialize_components()

    def _initialize_components(self):
        """Inicializa os componentes do detector"""
        # Inicializa modelo YOLO
        if YOLO_AVAILABLE:
            if self.model_path and os.path.exists(self.model_path):
                self.yolo_model = YOLO(self.model_path)
            else:
                # Usa modelo pre-treinado generico
                # Em producao, usar modelo treinado especificamente para placas
                self.yolo_model = None

        # Inicializa OCR
        if self.ocr_engine in ['easyocr', 'both'] and EASYOCR_AVAILABLE:
            self.ocr_reader = easyocr.Reader(
                ['pt', 'en'],
                gpu=self.use_gpu
            )

        # Inicializa validador
        if self.enable_validation:
            self.validator = PlateValidator()

        # Inicializa classificador
        if self.enable_classification:
            self.classifier = PlateClassifier()

        # Inicializa detector de adulteracao
        if self.enable_tampering_detection:
            self.tampering_detector = TamperingDetector()

    def _preprocess_plate_image(
        self,
        image: np.ndarray,
        enhance: bool = True
    ) -> np.ndarray:
        """
        Pre-processa imagem da placa para OCR.

        Args:
            image: Imagem da placa
            enhance: Aplicar melhorias de contraste

        Returns:
            Imagem pre-processada
        """
        if not CV2_AVAILABLE:
            return image

        # Redimensiona para tamanho padrao
        h, w = image.shape[:2]
        target_w, target_h = self.DEFAULT_PLATE_SIZE

        # Mantem proporcao
        scale = min(target_w / w, target_h / h)
        new_w = int(w * scale)
        new_h = int(h * scale)

        resized = cv2.resize(image, (new_w, new_h))

        if not enhance:
            return resized

        # Converte para escala de cinza
        if len(resized.shape) == 3:
            gray = cv2.cvtColor(resized, cv2.COLOR_BGR2GRAY)
        else:
            gray = resized

        # Aplica CLAHE para melhorar contraste
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        enhanced = clahe.apply(gray)

        # Aplica denoising
        denoised = cv2.fastNlMeansDenoising(enhanced, None, 10, 7, 21)

        # Converte de volta para BGR para compatibilidade
        result = cv2.cvtColor(denoised, cv2.COLOR_GRAY2BGR)

        return result

    def _detect_plate_region(
        self,
        image: np.ndarray
    ) -> List[DetectionBox]:
        """
        Detecta regioes de placa na imagem usando YOLO ou metodos tradicionais.

        Args:
            image: Imagem de entrada

        Returns:
            Lista de bounding boxes detectados
        """
        detections = []

        if self.yolo_model is not None:
            # Usa YOLO para deteccao
            results = self.yolo_model(image, conf=self.confidence_threshold)

            for result in results:
                boxes = result.boxes
                for box in boxes:
                    x1, y1, x2, y2 = map(int, box.xyxy[0])
                    conf = float(box.conf[0])
                    cls_id = int(box.cls[0])
                    cls_name = result.names[cls_id]

                    detections.append(DetectionBox(
                        x1=x1, y1=y1, x2=x2, y2=y2,
                        confidence=conf,
                        class_id=cls_id,
                        class_name=cls_name
                    ))

        else:
            # Fallback: usa metodos tradicionais de CV
            detections = self._detect_plate_traditional(image)

        return detections

    def _detect_plate_traditional(
        self,
        image: np.ndarray
    ) -> List[DetectionBox]:
        """
        Detecta placas usando metodos tradicionais de visao computacional.

        Args:
            image: Imagem de entrada

        Returns:
            Lista de bounding boxes
        """
        if not CV2_AVAILABLE:
            return []

        detections = []

        # Converte para escala de cinza
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

        # Aplica blur para reduzir ruido
        blurred = cv2.GaussianBlur(gray, (5, 5), 0)

        # Detecta bordas
        edges = cv2.Canny(blurred, 50, 150)

        # Dilata as bordas para conectar componentes
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
        dilated = cv2.dilate(edges, kernel, iterations=2)

        # Encontra contornos
        contours, _ = cv2.findContours(
            dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
        )

        img_h, img_w = image.shape[:2]

        for contour in contours:
            # Aproxima o contorno
            epsilon = 0.02 * cv2.arcLength(contour, True)
            approx = cv2.approxPolyDP(contour, epsilon, True)

            # Verifica se tem 4 vertices (retangulo)
            if len(approx) >= 4:
                x, y, w, h = cv2.boundingRect(contour)

                # Verifica proporcao (placas tem ~3:1)
                aspect_ratio = w / h if h > 0 else 0

                if 2.0 < aspect_ratio < 5.0:
                    # Verifica tamanho minimo e maximo
                    if (self.MIN_PLATE_SIZE[0] < w < self.MAX_PLATE_SIZE[0] and
                        self.MIN_PLATE_SIZE[1] < h < self.MAX_PLATE_SIZE[1]):

                        # Calcula confianca baseada na regularidade do contorno
                        area = cv2.contourArea(contour)
                        rect_area = w * h
                        rectangularity = area / rect_area if rect_area > 0 else 0

                        confidence = rectangularity * 0.7  # Confianca baseada na forma

                        if confidence > 0.3:
                            detections.append(DetectionBox(
                                x1=x, y1=y, x2=x+w, y2=y+h,
                                confidence=confidence,
                                class_id=0,
                                class_name='plate'
                            ))

        # Ordena por confianca
        detections.sort(key=lambda d: d.confidence, reverse=True)

        # Aplica Non-Maximum Suppression simples
        detections = self._apply_nms(detections, iou_threshold=0.5)

        return detections

    def _apply_nms(
        self,
        detections: List[DetectionBox],
        iou_threshold: float = 0.5
    ) -> List[DetectionBox]:
        """
        Aplica Non-Maximum Suppression para remover deteccoes redundantes.

        Args:
            detections: Lista de deteccoes
            iou_threshold: Limite de IoU para supressao

        Returns:
            Lista filtrada de deteccoes
        """
        if not detections:
            return []

        # Ordena por confianca
        sorted_dets = sorted(detections, key=lambda d: d.confidence, reverse=True)
        keep = []

        while sorted_dets:
            current = sorted_dets.pop(0)
            keep.append(current)

            remaining = []
            for det in sorted_dets:
                iou = self._calculate_iou(current, det)
                if iou < iou_threshold:
                    remaining.append(det)

            sorted_dets = remaining

        return keep

    def _calculate_iou(self, box1: DetectionBox, box2: DetectionBox) -> float:
        """Calcula Intersection over Union entre dois boxes"""
        x1 = max(box1.x1, box2.x1)
        y1 = max(box1.y1, box2.y1)
        x2 = min(box1.x2, box2.x2)
        y2 = min(box1.y2, box2.y2)

        intersection = max(0, x2 - x1) * max(0, y2 - y1)

        area1 = (box1.x2 - box1.x1) * (box1.y2 - box1.y1)
        area2 = (box2.x2 - box2.x1) * (box2.y2 - box2.y1)

        union = area1 + area2 - intersection

        return intersection / union if union > 0 else 0

    def _extract_plate_image(
        self,
        image: np.ndarray,
        box: DetectionBox
    ) -> np.ndarray:
        """
        Extrai a regiao da placa da imagem.

        Args:
            image: Imagem completa
            box: Bounding box da placa

        Returns:
            Imagem recortada da placa
        """
        return image[box.y1:box.y2, box.x1:box.x2]

    def _ocr_easyocr(
        self,
        image: np.ndarray
    ) -> Tuple[str, float, List[CharacterDetection]]:
        """
        Executa OCR usando EasyOCR.

        Args:
            image: Imagem da placa

        Returns:
            Tupla com (texto, confianca, caracteres)
        """
        if not EASYOCR_AVAILABLE or self.ocr_reader is None:
            return ('', 0.0, [])

        results = self.ocr_reader.readtext(
            image,
            allowlist=self.ALLOWED_CHARS,
            detail=1
        )

        if not results:
            return ('', 0.0, [])

        # Combina resultados
        full_text = ''
        total_conf = 0.0
        characters = []

        for (bbox, text, conf) in results:
            full_text += text
            total_conf += conf

            # Extrai bounding box
            x_coords = [p[0] for p in bbox]
            y_coords = [p[1] for p in bbox]
            x, y = int(min(x_coords)), int(min(y_coords))
            w = int(max(x_coords) - x)
            h = int(max(y_coords) - y)

            for i, char in enumerate(text):
                char_width = w // len(text) if len(text) > 0 else w
                characters.append(CharacterDetection(
                    char=char,
                    confidence=conf,
                    box=(x + i * char_width, y, char_width, h)
                ))

        avg_conf = total_conf / len(results) if results else 0.0

        return (full_text.upper(), avg_conf, characters)

    def _ocr_tesseract(
        self,
        image: np.ndarray
    ) -> Tuple[str, float, List[CharacterDetection]]:
        """
        Executa OCR usando Tesseract.

        Args:
            image: Imagem da placa

        Returns:
            Tupla com (texto, confianca, caracteres)
        """
        if not TESSERACT_AVAILABLE or not CV2_AVAILABLE:
            return ('', 0.0, [])

        # Converte para escala de cinza se necessario
        if len(image.shape) == 3:
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        else:
            gray = image

        # Aplica threshold
        _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

        # Configuracao do Tesseract para placas
        config = '--psm 7 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'

        try:
            # OCR com detalhes
            data = pytesseract.image_to_data(
                thresh,
                config=config,
                output_type=pytesseract.Output.DICT
            )

            text = ''
            total_conf = 0.0
            count = 0
            characters = []

            for i, char_text in enumerate(data['text']):
                if char_text.strip():
                    conf = int(data['conf'][i])
                    if conf > 0:
                        text += char_text
                        total_conf += conf / 100.0
                        count += 1

                        characters.append(CharacterDetection(
                            char=char_text,
                            confidence=conf / 100.0,
                            box=(
                                data['left'][i],
                                data['top'][i],
                                data['width'][i],
                                data['height'][i]
                            )
                        ))

            avg_conf = total_conf / count if count > 0 else 0.0

            return (text.upper(), avg_conf, characters)

        except Exception:
            return ('', 0.0, [])

    def _perform_ocr(
        self,
        image: np.ndarray
    ) -> Tuple[str, float, List[CharacterDetection]]:
        """
        Executa OCR na imagem da placa.

        Args:
            image: Imagem da placa

        Returns:
            Tupla com (texto, confianca, caracteres)
        """
        # Pre-processa a imagem
        processed = self._preprocess_plate_image(image)

        if self.ocr_engine == 'easyocr':
            return self._ocr_easyocr(processed)
        elif self.ocr_engine == 'tesseract':
            return self._ocr_tesseract(processed)
        elif self.ocr_engine == 'both':
            # Usa ambos e combina resultados
            easy_result = self._ocr_easyocr(processed)
            tess_result = self._ocr_tesseract(processed)

            # Escolhe o resultado com maior confianca
            if easy_result[1] >= tess_result[1]:
                return easy_result
            else:
                return tess_result
        else:
            return ('', 0.0, [])

    def detect(
        self,
        image: Union[np.ndarray, str],
        return_all: bool = False
    ) -> Union[PlateDetectionResult, List[PlateDetectionResult]]:
        """
        Detecta e analisa placas na imagem.

        Args:
            image: Imagem (numpy array) ou caminho para arquivo
            return_all: Se True, retorna todas as placas detectadas

        Returns:
            PlateDetectionResult ou lista de resultados
        """
        import time
        start_time = time.time()

        # Carrega imagem se for caminho
        if isinstance(image, str):
            if not CV2_AVAILABLE:
                raise ImportError("OpenCV necessario para carregar imagens")
            image = cv2.imread(image)
            if image is None:
                raise ValueError(f"Nao foi possivel carregar a imagem")

        # Detecta regioes de placa
        detections = self._detect_plate_region(image)

        if not detections:
            elapsed = (time.time() - start_time) * 1000
            return PlateDetectionResult(
                detected=False,
                plate_image=None,
                detection_box=None,
                detection_confidence=0.0,
                plate_text='',
                ocr_confidence=0.0,
                characters=[],
                validation=None,
                classification=None,
                tampering=None,
                processing_time_ms=elapsed,
                details={'error': 'Nenhuma placa detectada'}
            )

        results = []

        for detection in detections:
            # Extrai imagem da placa
            plate_image = self._extract_plate_image(image, detection)

            # Executa OCR
            plate_text, ocr_conf, characters = self._perform_ocr(plate_image)

            # Validacao
            validation = None
            if self.enable_validation and self.validator:
                validation = self.validator.validate(plate_text, ocr_conf)

            # Classificacao
            classification = None
            if self.enable_classification and self.classifier:
                classification = self.classifier.classify(plate_image, plate_text)

            # Deteccao de adulteracao
            tampering = None
            if self.enable_tampering_detection and self.tampering_detector:
                char_boxes = [c.box for c in characters] if characters else None
                tampering = self.tampering_detector.analyze(
                    plate_image,
                    char_boxes=char_boxes
                )

            elapsed = (time.time() - start_time) * 1000

            result = PlateDetectionResult(
                detected=True,
                plate_image=plate_image,
                detection_box=detection,
                detection_confidence=detection.confidence,
                plate_text=plate_text,
                ocr_confidence=ocr_conf,
                characters=characters,
                validation=validation,
                classification=classification,
                tampering=tampering,
                processing_time_ms=elapsed,
                details={
                    'detection_method': 'yolo' if self.yolo_model else 'traditional',
                    'ocr_engine': self.ocr_engine,
                }
            )

            results.append(result)

        if return_all:
            return results
        else:
            return results[0] if results else None

    def detect_from_video(
        self,
        video_path: str,
        frame_skip: int = 5,
        max_frames: Optional[int] = None
    ) -> List[PlateDetectionResult]:
        """
        Detecta placas em um video.

        Args:
            video_path: Caminho para o arquivo de video
            frame_skip: Processar a cada N frames
            max_frames: Limite maximo de frames a processar

        Returns:
            Lista de deteccoes unicas
        """
        if not CV2_AVAILABLE:
            raise ImportError("OpenCV necessario para processar video")

        cap = cv2.VideoCapture(video_path)
        if not cap.isOpened():
            raise ValueError(f"Nao foi possivel abrir o video: {video_path}")

        results = []
        seen_plates = set()
        frame_count = 0

        while cap.isOpened():
            ret, frame = cap.read()
            if not ret:
                break

            if max_frames and frame_count >= max_frames:
                break

            if frame_count % frame_skip == 0:
                detections = self.detect(frame, return_all=True)

                for det in detections:
                    if det.detected and det.plate_text:
                        # Evita duplicatas
                        if det.plate_text not in seen_plates:
                            seen_plates.add(det.plate_text)
                            results.append(det)

            frame_count += 1

        cap.release()
        return results


def detect_plate(
    image: Union[np.ndarray, str],
    ocr_engine: str = 'easyocr'
) -> PlateDetectionResult:
    """
    Funcao de conveniencia para detectar placa em uma imagem.

    Args:
        image: Imagem ou caminho
        ocr_engine: Motor de OCR

    Returns:
        Resultado da deteccao
    """
    detector = PlateDetector(ocr_engine=ocr_engine)
    return detector.detect(image)
