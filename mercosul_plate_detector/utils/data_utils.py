"""
Utilitarios para manipulacao de dados e anotacoes
"""

import json
import csv
import random
import string
from typing import Dict, List, Tuple, Optional, Any
from pathlib import Path
from dataclasses import asdict
import numpy as np

try:
    import cv2
    CV2_AVAILABLE = True
except ImportError:
    CV2_AVAILABLE = False

try:
    from PIL import Image, ImageDraw, ImageFont
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


def generate_plate_text() -> str:
    """
    Gera um texto de placa Mercosul valido aleatorio.

    Returns:
        Texto da placa no formato LLLNLNN
    """
    letters = string.ascii_uppercase
    digits = string.digits

    # Formato Mercosul: LLL N L NN
    plate = (
        random.choice(letters) +
        random.choice(letters) +
        random.choice(letters) +
        random.choice(digits) +
        random.choice(letters) +
        random.choice(digits) +
        random.choice(digits)
    )

    return plate


def generate_synthetic_plate(
    plate_text: Optional[str] = None,
    plate_type: str = 'particular',
    size: Tuple[int, int] = (400, 130),
    add_noise: bool = False
) -> np.ndarray:
    """
    Gera uma imagem sintetica de placa Mercosul.

    Args:
        plate_text: Texto da placa (gera aleatorio se None)
        plate_type: Tipo da placa (afeta as cores)
        size: Tamanho da imagem (largura, altura)
        add_noise: Adicionar ruido para simular condicoes reais

    Returns:
        Imagem da placa sintetica
    """
    if not PIL_AVAILABLE:
        # Fallback com OpenCV
        if not CV2_AVAILABLE:
            raise ImportError("PIL ou OpenCV necessario")

        return _generate_plate_cv2(plate_text, plate_type, size, add_noise)

    if plate_text is None:
        plate_text = generate_plate_text()

    # Cores por tipo de placa
    colors = {
        'particular': {'bg': (255, 255, 255), 'text': (0, 0, 0)},
        'comercial': {'bg': (255, 0, 0), 'text': (255, 255, 255)},
        'oficial': {'bg': (0, 0, 255), 'text': (255, 255, 255)},
        'diplomatico': {'bg': (0, 0, 200), 'text': (255, 255, 255)},
        'especial': {'bg': (0, 128, 0), 'text': (255, 255, 255)},
        'colecionador': {'bg': (128, 128, 128), 'text': (0, 0, 0)},
    }

    color_scheme = colors.get(plate_type, colors['particular'])

    # Cria imagem
    img = Image.new('RGB', size, color_scheme['bg'])
    draw = ImageDraw.Draw(img)

    # Tenta carregar fonte (usa default se falhar)
    try:
        # Tenta fontes comuns
        font_size = int(size[1] * 0.6)
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", font_size)
    except (OSError, IOError):
        font = ImageFont.load_default()

    # Calcula posicao do texto (centralizado)
    try:
        bbox = draw.textbbox((0, 0), plate_text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
    except AttributeError:
        # Versao antiga do PIL
        text_width, text_height = draw.textsize(plate_text, font=font)

    x = (size[0] - text_width) // 2
    y = (size[1] - text_height) // 2

    # Desenha texto
    draw.text((x, y), plate_text, fill=color_scheme['text'], font=font)

    # Desenha borda
    draw.rectangle([(0, 0), (size[0]-1, size[1]-1)], outline=(0, 0, 0), width=2)

    # Converte para numpy
    plate_array = np.array(img)

    # Converte RGB para BGR (OpenCV)
    plate_array = plate_array[:, :, ::-1]

    # Adiciona ruido se solicitado
    if add_noise and CV2_AVAILABLE:
        noise = np.random.normal(0, 10, plate_array.shape).astype(np.uint8)
        plate_array = cv2.add(plate_array, noise)

    return plate_array


def _generate_plate_cv2(
    plate_text: Optional[str],
    plate_type: str,
    size: Tuple[int, int],
    add_noise: bool
) -> np.ndarray:
    """Gera placa usando apenas OpenCV"""
    if plate_text is None:
        plate_text = generate_plate_text()

    # Cores BGR
    colors = {
        'particular': {'bg': (255, 255, 255), 'text': (0, 0, 0)},
        'comercial': {'bg': (0, 0, 255), 'text': (255, 255, 255)},
        'oficial': {'bg': (255, 0, 0), 'text': (255, 255, 255)},
        'especial': {'bg': (0, 128, 0), 'text': (255, 255, 255)},
        'colecionador': {'bg': (128, 128, 128), 'text': (0, 0, 0)},
    }

    color_scheme = colors.get(plate_type, colors['particular'])

    # Cria imagem
    img = np.full((size[1], size[0], 3), color_scheme['bg'], dtype=np.uint8)

    # Adiciona texto
    font = cv2.FONT_HERSHEY_SIMPLEX
    font_scale = 2.0
    thickness = 3

    text_size = cv2.getTextSize(plate_text, font, font_scale, thickness)[0]
    x = (size[0] - text_size[0]) // 2
    y = (size[1] + text_size[1]) // 2

    cv2.putText(img, plate_text, (x, y), font, font_scale, color_scheme['text'], thickness)

    # Borda
    cv2.rectangle(img, (0, 0), (size[0]-1, size[1]-1), (0, 0, 0), 2)

    if add_noise:
        noise = np.random.normal(0, 10, img.shape).astype(np.uint8)
        img = cv2.add(img, noise)

    return img


def create_annotation(
    image_path: str,
    boxes: List[Tuple[int, int, int, int]],
    labels: List[str],
    image_size: Tuple[int, int],
    format: str = 'yolo'
) -> str:
    """
    Cria anotacao para uma imagem.

    Args:
        image_path: Caminho da imagem
        boxes: Lista de bounding boxes [(x1, y1, x2, y2), ...]
        labels: Lista de labels
        image_size: (largura, altura) da imagem
        format: Formato da anotacao ('yolo', 'coco', 'voc')

    Returns:
        String com anotacao ou dict para COCO
    """
    img_w, img_h = image_size

    if format == 'yolo':
        lines = []
        for (x1, y1, x2, y2), label in zip(boxes, labels):
            # Converte para formato YOLO normalizado
            x_center = ((x1 + x2) / 2) / img_w
            y_center = ((y1 + y2) / 2) / img_h
            width = (x2 - x1) / img_w
            height = (y2 - y1) / img_h

            # Class ID (0 para plate)
            class_id = 0

            lines.append(f"{class_id} {x_center:.6f} {y_center:.6f} {width:.6f} {height:.6f}")

        return '\n'.join(lines)

    elif format == 'coco':
        annotations = []
        for i, ((x1, y1, x2, y2), label) in enumerate(zip(boxes, labels)):
            annotations.append({
                'id': i,
                'image_id': image_path,
                'category_id': 0,
                'bbox': [x1, y1, x2 - x1, y2 - y1],
                'area': (x2 - x1) * (y2 - y1),
                'iscrowd': 0
            })
        return json.dumps(annotations)

    elif format == 'voc':
        # Formato Pascal VOC XML simplificado
        xml_template = """<annotation>
    <filename>{filename}</filename>
    <size>
        <width>{width}</width>
        <height>{height}</height>
    </size>
    {objects}
</annotation>"""

        objects = []
        for (x1, y1, x2, y2), label in zip(boxes, labels):
            obj = f"""<object>
        <name>{label}</name>
        <bndbox>
            <xmin>{x1}</xmin>
            <ymin>{y1}</ymin>
            <xmax>{x2}</xmax>
            <ymax>{y2}</ymax>
        </bndbox>
    </object>"""
            objects.append(obj)

        return xml_template.format(
            filename=Path(image_path).name,
            width=img_w,
            height=img_h,
            objects='\n    '.join(objects)
        )

    else:
        raise ValueError(f"Formato nao suportado: {format}")


def export_results_json(
    results: List[Dict[str, Any]],
    output_path: str,
    pretty: bool = True
) -> None:
    """
    Exporta resultados para arquivo JSON.

    Args:
        results: Lista de resultados
        output_path: Caminho do arquivo de saida
        pretty: Formatar JSON com indentacao
    """
    with open(output_path, 'w', encoding='utf-8') as f:
        if pretty:
            json.dump(results, f, indent=2, ensure_ascii=False, default=str)
        else:
            json.dump(results, f, ensure_ascii=False, default=str)


def export_results_csv(
    results: List[Dict[str, Any]],
    output_path: str,
    columns: Optional[List[str]] = None
) -> None:
    """
    Exporta resultados para arquivo CSV.

    Args:
        results: Lista de resultados
        output_path: Caminho do arquivo de saida
        columns: Lista de colunas a incluir (todas se None)
    """
    if not results:
        return

    if columns is None:
        columns = list(results[0].keys())

    with open(output_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=columns, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(results)


def generate_synthetic_dataset(
    output_dir: str,
    num_images: int = 100,
    plate_types: Optional[List[str]] = None,
    with_augmentation: bool = True
) -> None:
    """
    Gera um dataset sintetico de placas para treinamento.

    Args:
        output_dir: Diretorio de saida
        num_images: Numero de imagens a gerar
        plate_types: Tipos de placa a incluir
        with_augmentation: Aplicar augmentacoes
    """
    if plate_types is None:
        plate_types = ['particular', 'comercial', 'oficial', 'especial', 'colecionador']

    output_path = Path(output_dir)
    images_dir = output_path / 'images'
    labels_dir = output_path / 'labels'

    images_dir.mkdir(parents=True, exist_ok=True)
    labels_dir.mkdir(parents=True, exist_ok=True)

    for i in range(num_images):
        # Escolhe tipo aleatorio
        plate_type = random.choice(plate_types)

        # Gera placa
        plate_text = generate_plate_text()
        plate_img = generate_synthetic_plate(
            plate_text,
            plate_type,
            add_noise=with_augmentation
        )

        # Cria imagem de fundo aleatorio
        bg_size = (800, 600)
        background = np.random.randint(50, 200, (bg_size[1], bg_size[0], 3), dtype=np.uint8)

        # Posicao aleatoria para a placa
        plate_h, plate_w = plate_img.shape[:2]
        max_x = bg_size[0] - plate_w - 10
        max_y = bg_size[1] - plate_h - 10
        x = random.randint(10, max(10, max_x))
        y = random.randint(10, max(10, max_y))

        # Insere placa no fundo
        background[y:y+plate_h, x:x+plate_w] = plate_img

        # Salva imagem
        img_filename = f"plate_{i:05d}.jpg"
        img_path = images_dir / img_filename

        if CV2_AVAILABLE:
            cv2.imwrite(str(img_path), background)

        # Cria anotacao YOLO
        annotation = create_annotation(
            str(img_path),
            [(x, y, x + plate_w, y + plate_h)],
            ['plate'],
            bg_size,
            format='yolo'
        )

        label_filename = f"plate_{i:05d}.txt"
        label_path = labels_dir / label_filename
        label_path.write_text(annotation)

    print(f"Dataset sintetico gerado: {num_images} imagens em {output_dir}")


def load_dataset_stats(data_dir: str) -> Dict[str, Any]:
    """
    Carrega estatisticas de um dataset.

    Args:
        data_dir: Diretorio do dataset

    Returns:
        Dicionario com estatisticas
    """
    data_path = Path(data_dir)

    stats = {
        'total_images': 0,
        'total_labels': 0,
        'splits': {},
        'classes': {}
    }

    for split in ['train', 'val', 'test']:
        images_dir = data_path / 'images' / split
        labels_dir = data_path / 'labels' / split

        if images_dir.exists():
            num_images = len(list(images_dir.glob('*.jpg'))) + \
                        len(list(images_dir.glob('*.png')))
            stats['splits'][split] = {'images': num_images}
            stats['total_images'] += num_images

        if labels_dir.exists():
            num_labels = len(list(labels_dir.glob('*.txt')))
            if split in stats['splits']:
                stats['splits'][split]['labels'] = num_labels
            stats['total_labels'] += num_labels

    return stats
