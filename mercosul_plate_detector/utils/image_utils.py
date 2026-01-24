"""
Utilitarios de processamento de imagem
"""

import numpy as np
from typing import Tuple, Optional, List, Union
from pathlib import Path

try:
    import cv2
    CV2_AVAILABLE = True
except ImportError:
    CV2_AVAILABLE = False


def load_image(path: Union[str, Path]) -> Optional[np.ndarray]:
    """
    Carrega uma imagem do disco.

    Args:
        path: Caminho para a imagem

    Returns:
        Imagem em formato numpy BGR ou None se falhar
    """
    if not CV2_AVAILABLE:
        raise ImportError("OpenCV necessario para carregar imagens")

    image = cv2.imread(str(path))
    return image


def save_image(image: np.ndarray, path: Union[str, Path]) -> bool:
    """
    Salva uma imagem no disco.

    Args:
        image: Imagem em formato numpy
        path: Caminho de destino

    Returns:
        True se salvar com sucesso
    """
    if not CV2_AVAILABLE:
        raise ImportError("OpenCV necessario para salvar imagens")

    return cv2.imwrite(str(path), image)


def resize_image(
    image: np.ndarray,
    target_size: Tuple[int, int],
    keep_aspect: bool = True
) -> np.ndarray:
    """
    Redimensiona uma imagem.

    Args:
        image: Imagem original
        target_size: (largura, altura) desejada
        keep_aspect: Manter proporcao

    Returns:
        Imagem redimensionada
    """
    if not CV2_AVAILABLE:
        return image

    h, w = image.shape[:2]
    target_w, target_h = target_size

    if keep_aspect:
        scale = min(target_w / w, target_h / h)
        new_w = int(w * scale)
        new_h = int(h * scale)
    else:
        new_w, new_h = target_w, target_h

    resized = cv2.resize(image, (new_w, new_h))

    return resized


def draw_detection(
    image: np.ndarray,
    box: Tuple[int, int, int, int],
    label: str,
    color: Tuple[int, int, int] = (0, 255, 0),
    thickness: int = 2
) -> np.ndarray:
    """
    Desenha uma deteccao na imagem.

    Args:
        image: Imagem original
        box: Bounding box (x1, y1, x2, y2)
        label: Texto do label
        color: Cor BGR
        thickness: Espessura das linhas

    Returns:
        Imagem com deteccao desenhada
    """
    if not CV2_AVAILABLE:
        return image

    result = image.copy()
    x1, y1, x2, y2 = box

    # Desenha retangulo
    cv2.rectangle(result, (x1, y1), (x2, y2), color, thickness)

    # Desenha fundo do label
    label_size = cv2.getTextSize(label, cv2.FONT_HERSHEY_SIMPLEX, 0.6, 2)[0]
    cv2.rectangle(
        result,
        (x1, y1 - label_size[1] - 10),
        (x1 + label_size[0], y1),
        color,
        -1
    )

    # Desenha texto
    cv2.putText(
        result,
        label,
        (x1, y1 - 5),
        cv2.FONT_HERSHEY_SIMPLEX,
        0.6,
        (0, 0, 0),
        2
    )

    return result


def create_plate_mosaic(
    plate_images: List[np.ndarray],
    cols: int = 4,
    plate_size: Tuple[int, int] = (200, 65),
    padding: int = 10,
    background_color: Tuple[int, int, int] = (50, 50, 50)
) -> np.ndarray:
    """
    Cria um mosaico de imagens de placas.

    Args:
        plate_images: Lista de imagens de placas
        cols: Numero de colunas
        plate_size: Tamanho de cada placa no mosaico
        padding: Espacamento entre placas
        background_color: Cor de fundo

    Returns:
        Imagem do mosaico
    """
    if not CV2_AVAILABLE:
        raise ImportError("OpenCV necessario")

    if not plate_images:
        return np.zeros((100, 100, 3), dtype=np.uint8)

    n_images = len(plate_images)
    rows = (n_images + cols - 1) // cols

    pw, ph = plate_size

    # Calcula dimensoes do mosaico
    mosaic_w = cols * pw + (cols + 1) * padding
    mosaic_h = rows * ph + (rows + 1) * padding

    # Cria imagem de fundo
    mosaic = np.full((mosaic_h, mosaic_w, 3), background_color, dtype=np.uint8)

    # Adiciona cada placa
    for i, plate in enumerate(plate_images):
        row = i // cols
        col = i % cols

        x = padding + col * (pw + padding)
        y = padding + row * (ph + padding)

        # Redimensiona placa
        resized = cv2.resize(plate, (pw, ph))

        # Insere no mosaico
        mosaic[y:y+ph, x:x+pw] = resized

    return mosaic


def enhance_plate_image(image: np.ndarray) -> np.ndarray:
    """
    Melhora a qualidade da imagem da placa para OCR.

    Args:
        image: Imagem da placa

    Returns:
        Imagem melhorada
    """
    if not CV2_AVAILABLE:
        return image

    # Converte para escala de cinza
    if len(image.shape) == 3:
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    else:
        gray = image

    # Aplica CLAHE
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(gray)

    # Remove ruido
    denoised = cv2.fastNlMeansDenoising(enhanced, None, 10, 7, 21)

    # Binariza
    _, binary = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    return binary


def extract_characters(
    plate_image: np.ndarray,
    num_chars: int = 7
) -> List[np.ndarray]:
    """
    Extrai imagens individuais dos caracteres da placa.

    Args:
        plate_image: Imagem da placa
        num_chars: Numero esperado de caracteres

    Returns:
        Lista de imagens de caracteres
    """
    if not CV2_AVAILABLE:
        return []

    # Melhora a imagem
    enhanced = enhance_plate_image(plate_image)

    # Encontra contornos
    contours, _ = cv2.findContours(
        enhanced, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
    )

    # Filtra e ordena por posicao x
    char_boxes = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)

        # Filtra por tamanho (caracteres tem proporcao ~2:3)
        if h > plate_image.shape[0] * 0.3 and w > 5:
            char_boxes.append((x, y, w, h))

    # Ordena por posicao x
    char_boxes.sort(key=lambda b: b[0])

    # Extrai caracteres
    characters = []
    for x, y, w, h in char_boxes[:num_chars]:
        char_img = plate_image[y:y+h, x:x+w]
        characters.append(char_img)

    return characters
