"""
Modulo de Treinamento de Modelos para Deteccao de Placas Mercosul

Fornece funcionalidades para:
- Preparacao de datasets
- Data augmentation especifico para placas
- Treinamento de modelos YOLO para deteccao
- Treinamento de classificadores de tipo de placa
- Avaliacao e metricas
"""

import os
import json
import random
import shutil
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Callable
from dataclasses import dataclass
import numpy as np

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
    import torch
    import torch.nn as nn
    import torch.optim as optim
    from torch.utils.data import Dataset, DataLoader
    TORCH_AVAILABLE = True
except ImportError:
    TORCH_AVAILABLE = False


@dataclass
class TrainingConfig:
    """Configuracao de treinamento"""
    # Paths
    data_dir: str
    output_dir: str
    model_name: str = 'mercosul_plate_detector'

    # Treinamento
    epochs: int = 100
    batch_size: int = 16
    learning_rate: float = 0.001
    patience: int = 10  # Early stopping

    # Modelo
    base_model: str = 'yolov8n.pt'  # Modelo base YOLO
    image_size: int = 640

    # Data augmentation
    augmentation: bool = True
    rotation_range: float = 5.0
    brightness_range: Tuple[float, float] = (0.8, 1.2)
    noise_level: float = 0.02

    # Validacao
    val_split: float = 0.2
    test_split: float = 0.1


@dataclass
class TrainingMetrics:
    """Metricas de treinamento"""
    epoch: int
    train_loss: float
    val_loss: float
    precision: float
    recall: float
    mAP50: float
    mAP50_95: float


class PlateDataAugmenter:
    """
    Augmentacao de dados especifica para placas de veiculo.

    Aplica transformacoes que simulam condicoes reais:
    - Variacao de iluminacao
    - Rotacao leve (placas nem sempre estao perfeitamente alinhadas)
    - Ruido (cameras de baixa qualidade)
    - Blur de movimento
    - Variacao de perspectiva
    """

    def __init__(self, config: TrainingConfig):
        self.config = config

    def apply_rotation(self, image: np.ndarray) -> np.ndarray:
        """Aplica rotacao aleatoria"""
        if not CV2_AVAILABLE:
            return image

        angle = random.uniform(
            -self.config.rotation_range,
            self.config.rotation_range
        )

        h, w = image.shape[:2]
        center = (w // 2, h // 2)

        matrix = cv2.getRotationMatrix2D(center, angle, 1.0)
        rotated = cv2.warpAffine(image, matrix, (w, h))

        return rotated

    def apply_brightness(self, image: np.ndarray) -> np.ndarray:
        """Aplica variacao de brilho"""
        if not CV2_AVAILABLE:
            return image

        factor = random.uniform(*self.config.brightness_range)
        adjusted = cv2.convertScaleAbs(image, alpha=factor, beta=0)

        return adjusted

    def apply_noise(self, image: np.ndarray) -> np.ndarray:
        """Adiciona ruido gaussiano"""
        noise_level = self.config.noise_level
        noise = np.random.normal(0, noise_level * 255, image.shape).astype(np.uint8)
        noisy = cv2.add(image, noise)

        return noisy

    def apply_blur(self, image: np.ndarray) -> np.ndarray:
        """Aplica blur aleatorio"""
        if not CV2_AVAILABLE:
            return image

        blur_type = random.choice(['gaussian', 'motion', 'none'])

        if blur_type == 'gaussian':
            ksize = random.choice([3, 5])
            blurred = cv2.GaussianBlur(image, (ksize, ksize), 0)
        elif blur_type == 'motion':
            ksize = random.randint(3, 7)
            kernel = np.zeros((ksize, ksize))
            kernel[ksize // 2, :] = 1 / ksize
            blurred = cv2.filter2D(image, -1, kernel)
        else:
            blurred = image

        return blurred

    def apply_perspective(self, image: np.ndarray) -> np.ndarray:
        """Aplica transformacao de perspectiva leve"""
        if not CV2_AVAILABLE:
            return image

        h, w = image.shape[:2]

        # Pontos originais
        pts1 = np.float32([[0, 0], [w, 0], [0, h], [w, h]])

        # Pontos com pequena variacao aleatoria
        offset = 10
        pts2 = np.float32([
            [random.randint(0, offset), random.randint(0, offset)],
            [w - random.randint(0, offset), random.randint(0, offset)],
            [random.randint(0, offset), h - random.randint(0, offset)],
            [w - random.randint(0, offset), h - random.randint(0, offset)]
        ])

        matrix = cv2.getPerspectiveTransform(pts1, pts2)
        transformed = cv2.warpPerspective(image, matrix, (w, h))

        return transformed

    def augment(
        self,
        image: np.ndarray,
        probability: float = 0.5
    ) -> np.ndarray:
        """
        Aplica augmentacoes aleatorias na imagem.

        Args:
            image: Imagem original
            probability: Probabilidade de aplicar cada augmentacao

        Returns:
            Imagem augmentada
        """
        result = image.copy()

        if random.random() < probability:
            result = self.apply_rotation(result)

        if random.random() < probability:
            result = self.apply_brightness(result)

        if random.random() < probability * 0.5:
            result = self.apply_noise(result)

        if random.random() < probability * 0.3:
            result = self.apply_blur(result)

        if random.random() < probability * 0.3:
            result = self.apply_perspective(result)

        return result


class PlateDataset:
    """
    Gerenciador de dataset para treinamento.

    Estrutura esperada do dataset:
    data_dir/
        images/
            train/
            val/
            test/
        labels/
            train/
            val/
            test/
    """

    def __init__(self, data_dir: str, config: TrainingConfig):
        self.data_dir = Path(data_dir)
        self.config = config
        self.augmenter = PlateDataAugmenter(config)

    def setup_directories(self):
        """Cria estrutura de diretorios"""
        for split in ['train', 'val', 'test']:
            (self.data_dir / 'images' / split).mkdir(parents=True, exist_ok=True)
            (self.data_dir / 'labels' / split).mkdir(parents=True, exist_ok=True)

    def convert_annotation(
        self,
        annotation: Dict,
        image_size: Tuple[int, int],
        format_from: str = 'coco',
        format_to: str = 'yolo'
    ) -> str:
        """
        Converte anotacao entre formatos.

        Args:
            annotation: Dicionario de anotacao
            image_size: (largura, altura) da imagem
            format_from: Formato de origem ('coco', 'voc', 'labelme')
            format_to: Formato de destino ('yolo')

        Returns:
            String com anotacao no formato destino
        """
        img_w, img_h = image_size

        if format_from == 'coco':
            # COCO: [x, y, width, height] em pixels
            x, y, w, h = annotation['bbox']
            class_id = annotation.get('category_id', 0)

        elif format_from == 'voc':
            # VOC/Pascal: [xmin, ymin, xmax, ymax] em pixels
            xmin = annotation['xmin']
            ymin = annotation['ymin']
            xmax = annotation['xmax']
            ymax = annotation['ymax']
            x, y = xmin, ymin
            w, h = xmax - xmin, ymax - ymin
            class_id = annotation.get('class_id', 0)

        elif format_from == 'labelme':
            # LabelMe: pontos do poligono
            points = annotation['points']
            xs = [p[0] for p in points]
            ys = [p[1] for p in points]
            x, y = min(xs), min(ys)
            w, h = max(xs) - x, max(ys) - y
            class_id = 0

        else:
            raise ValueError(f"Formato nao suportado: {format_from}")

        if format_to == 'yolo':
            # YOLO: [class_id, x_center, y_center, width, height] normalizados
            x_center = (x + w / 2) / img_w
            y_center = (y + h / 2) / img_h
            w_norm = w / img_w
            h_norm = h / img_h

            return f"{class_id} {x_center:.6f} {y_center:.6f} {w_norm:.6f} {h_norm:.6f}"

        else:
            raise ValueError(f"Formato de destino nao suportado: {format_to}")

    def split_dataset(
        self,
        images_dir: str,
        labels_dir: str,
        annotation_format: str = 'yolo'
    ):
        """
        Divide o dataset em train/val/test.

        Args:
            images_dir: Diretorio com imagens
            labels_dir: Diretorio com anotacoes
            annotation_format: Formato das anotacoes
        """
        images_path = Path(images_dir)
        labels_path = Path(labels_dir)

        # Lista todas as imagens
        image_files = list(images_path.glob('*.jpg')) + \
                     list(images_path.glob('*.png')) + \
                     list(images_path.glob('*.jpeg'))

        # Embaralha
        random.shuffle(image_files)

        # Calcula indices de divisao
        n_total = len(image_files)
        n_test = int(n_total * self.config.test_split)
        n_val = int(n_total * self.config.val_split)

        test_files = image_files[:n_test]
        val_files = image_files[n_test:n_test + n_val]
        train_files = image_files[n_test + n_val:]

        # Move arquivos
        for split, files in [('test', test_files), ('val', val_files), ('train', train_files)]:
            for img_file in files:
                # Copia imagem
                dest_img = self.data_dir / 'images' / split / img_file.name
                shutil.copy2(img_file, dest_img)

                # Copia label (se existir)
                label_file = labels_path / f"{img_file.stem}.txt"
                if label_file.exists():
                    dest_label = self.data_dir / 'labels' / split / label_file.name
                    shutil.copy2(label_file, dest_label)

        print(f"Dataset dividido: {len(train_files)} train, {len(val_files)} val, {len(test_files)} test")

    def generate_augmented_samples(
        self,
        num_samples: int = 5
    ):
        """
        Gera amostras augmentadas para o conjunto de treino.

        Args:
            num_samples: Numero de amostras augmentadas por imagem
        """
        if not CV2_AVAILABLE:
            print("OpenCV necessario para augmentacao")
            return

        train_images_dir = self.data_dir / 'images' / 'train'
        train_labels_dir = self.data_dir / 'labels' / 'train'

        image_files = list(train_images_dir.glob('*.jpg')) + \
                     list(train_images_dir.glob('*.png'))

        for img_file in image_files:
            image = cv2.imread(str(img_file))
            if image is None:
                continue

            label_file = train_labels_dir / f"{img_file.stem}.txt"
            label_content = ''
            if label_file.exists():
                label_content = label_file.read_text()

            for i in range(num_samples):
                # Gera versao augmentada
                aug_image = self.augmenter.augment(image)

                # Salva com novo nome
                aug_name = f"{img_file.stem}_aug{i}{img_file.suffix}"
                aug_path = train_images_dir / aug_name
                cv2.imwrite(str(aug_path), aug_image)

                # Copia label (mesmas coordenadas, transformacoes sao leves)
                if label_content:
                    aug_label_path = train_labels_dir / f"{img_file.stem}_aug{i}.txt"
                    aug_label_path.write_text(label_content)

        print(f"Amostras augmentadas geradas: {len(image_files) * num_samples}")

    def create_data_yaml(self) -> str:
        """
        Cria arquivo data.yaml para treinamento YOLO.

        Returns:
            Caminho do arquivo criado
        """
        yaml_content = f"""
# Dataset de Placas Mercosul
path: {self.data_dir.absolute()}
train: images/train
val: images/val
test: images/test

# Classes
names:
  0: plate
  1: plate_particular
  2: plate_comercial
  3: plate_oficial
  4: plate_diplomatico
  5: plate_especial
  6: plate_colecionador

# Numero de classes
nc: 7
"""
        yaml_path = self.data_dir / 'data.yaml'
        yaml_path.write_text(yaml_content)

        return str(yaml_path)


class PlateTrainer:
    """
    Treinador de modelos para deteccao de placas.

    Suporta:
    - Treinamento de modelos YOLO para deteccao
    - Fine-tuning de modelos pre-treinados
    - Treinamento de classificadores customizados
    """

    def __init__(self, config: TrainingConfig):
        self.config = config
        self.dataset = PlateDataset(config.data_dir, config)
        self.metrics_history: List[TrainingMetrics] = []

    def setup(self):
        """Prepara ambiente de treinamento"""
        # Cria diretorios
        self.dataset.setup_directories()
        Path(self.config.output_dir).mkdir(parents=True, exist_ok=True)

        # Cria arquivo de configuracao
        self.data_yaml_path = self.dataset.create_data_yaml()

        print(f"Ambiente de treinamento preparado em: {self.config.output_dir}")

    def prepare_data(
        self,
        images_dir: str,
        labels_dir: str,
        augment: bool = True
    ):
        """
        Prepara os dados para treinamento.

        Args:
            images_dir: Diretorio com imagens
            labels_dir: Diretorio com labels
            augment: Gerar dados augmentados
        """
        # Divide dataset
        self.dataset.split_dataset(images_dir, labels_dir)

        # Gera augmentacoes
        if augment and self.config.augmentation:
            self.dataset.generate_augmented_samples()

    def train_yolo(
        self,
        resume: bool = False,
        weights_path: Optional[str] = None
    ) -> str:
        """
        Treina modelo YOLO para deteccao de placas.

        Args:
            resume: Continuar treinamento anterior
            weights_path: Caminho para pesos iniciais

        Returns:
            Caminho do melhor modelo treinado
        """
        if not YOLO_AVAILABLE:
            raise ImportError("Ultralytics YOLO necessario para treinamento")

        # Carrega modelo base ou customizado
        if weights_path and os.path.exists(weights_path):
            model = YOLO(weights_path)
        else:
            model = YOLO(self.config.base_model)

        # Configura treinamento
        results = model.train(
            data=self.data_yaml_path,
            epochs=self.config.epochs,
            batch=self.config.batch_size,
            imgsz=self.config.image_size,
            patience=self.config.patience,
            save=True,
            project=self.config.output_dir,
            name=self.config.model_name,
            exist_ok=True,
            resume=resume,
            lr0=self.config.learning_rate,
            verbose=True
        )

        # Caminho do melhor modelo
        best_model_path = Path(self.config.output_dir) / self.config.model_name / 'weights' / 'best.pt'

        print(f"Treinamento concluido. Melhor modelo: {best_model_path}")

        return str(best_model_path)

    def evaluate(self, model_path: str) -> Dict:
        """
        Avalia o modelo no conjunto de teste.

        Args:
            model_path: Caminho do modelo a avaliar

        Returns:
            Dicionario com metricas
        """
        if not YOLO_AVAILABLE:
            raise ImportError("Ultralytics YOLO necessario para avaliacao")

        model = YOLO(model_path)

        # Avalia no conjunto de teste
        results = model.val(
            data=self.data_yaml_path,
            split='test'
        )

        metrics = {
            'precision': float(results.box.mp),
            'recall': float(results.box.mr),
            'mAP50': float(results.box.map50),
            'mAP50_95': float(results.box.map),
            'confusion_matrix': results.confusion_matrix.matrix.tolist() if hasattr(results, 'confusion_matrix') else None
        }

        return metrics

    def export_model(
        self,
        model_path: str,
        format: str = 'onnx'
    ) -> str:
        """
        Exporta modelo para formato de producao.

        Args:
            model_path: Caminho do modelo treinado
            format: Formato de exportacao ('onnx', 'torchscript', 'tflite')

        Returns:
            Caminho do modelo exportado
        """
        if not YOLO_AVAILABLE:
            raise ImportError("Ultralytics YOLO necessario para exportacao")

        model = YOLO(model_path)

        # Exporta
        export_path = model.export(format=format)

        print(f"Modelo exportado para: {export_path}")

        return export_path


class PlateClassifierTrainer:
    """
    Treinador de classificador de tipos de placa.

    Usa CNN para classificar placas em categorias:
    - Particular
    - Comercial
    - Oficial
    - Diplomatico
    - Especial
    - Colecionador
    """

    def __init__(self, config: TrainingConfig):
        self.config = config
        self.num_classes = 6  # Tipos de placa

    def create_model(self) -> 'nn.Module':
        """
        Cria modelo CNN para classificacao.

        Returns:
            Modelo PyTorch
        """
        if not TORCH_AVAILABLE:
            raise ImportError("PyTorch necessario para classificador")

        class PlateClassifierCNN(nn.Module):
            def __init__(self, num_classes):
                super().__init__()

                # Backbone simples
                self.features = nn.Sequential(
                    nn.Conv2d(3, 32, 3, padding=1),
                    nn.BatchNorm2d(32),
                    nn.ReLU(),
                    nn.MaxPool2d(2),

                    nn.Conv2d(32, 64, 3, padding=1),
                    nn.BatchNorm2d(64),
                    nn.ReLU(),
                    nn.MaxPool2d(2),

                    nn.Conv2d(64, 128, 3, padding=1),
                    nn.BatchNorm2d(128),
                    nn.ReLU(),
                    nn.MaxPool2d(2),

                    nn.Conv2d(128, 256, 3, padding=1),
                    nn.BatchNorm2d(256),
                    nn.ReLU(),
                    nn.AdaptiveAvgPool2d((4, 4))
                )

                # Classificador
                self.classifier = nn.Sequential(
                    nn.Flatten(),
                    nn.Linear(256 * 4 * 4, 512),
                    nn.ReLU(),
                    nn.Dropout(0.5),
                    nn.Linear(512, num_classes)
                )

            def forward(self, x):
                x = self.features(x)
                x = self.classifier(x)
                return x

        return PlateClassifierCNN(self.num_classes)

    def train(
        self,
        train_loader: 'DataLoader',
        val_loader: 'DataLoader',
        epochs: Optional[int] = None
    ) -> str:
        """
        Treina o classificador.

        Args:
            train_loader: DataLoader de treino
            val_loader: DataLoader de validacao
            epochs: Numero de epocas (usa config se nao especificado)

        Returns:
            Caminho do melhor modelo
        """
        if not TORCH_AVAILABLE:
            raise ImportError("PyTorch necessario para treinamento")

        epochs = epochs or self.config.epochs
        device = torch.device('cuda' if torch.cuda.is_available() else 'cpu')

        model = self.create_model().to(device)
        criterion = nn.CrossEntropyLoss()
        optimizer = optim.Adam(model.parameters(), lr=self.config.learning_rate)
        scheduler = optim.lr_scheduler.ReduceLROnPlateau(optimizer, patience=5)

        best_val_acc = 0.0
        best_model_path = Path(self.config.output_dir) / 'plate_classifier_best.pt'

        for epoch in range(epochs):
            # Treino
            model.train()
            train_loss = 0.0
            train_correct = 0
            train_total = 0

            for images, labels in train_loader:
                images, labels = images.to(device), labels.to(device)

                optimizer.zero_grad()
                outputs = model(images)
                loss = criterion(outputs, labels)
                loss.backward()
                optimizer.step()

                train_loss += loss.item()
                _, predicted = outputs.max(1)
                train_total += labels.size(0)
                train_correct += predicted.eq(labels).sum().item()

            train_acc = train_correct / train_total

            # Validacao
            model.eval()
            val_loss = 0.0
            val_correct = 0
            val_total = 0

            with torch.no_grad():
                for images, labels in val_loader:
                    images, labels = images.to(device), labels.to(device)
                    outputs = model(images)
                    loss = criterion(outputs, labels)

                    val_loss += loss.item()
                    _, predicted = outputs.max(1)
                    val_total += labels.size(0)
                    val_correct += predicted.eq(labels).sum().item()

            val_acc = val_correct / val_total

            scheduler.step(val_loss)

            print(f"Epoch {epoch+1}/{epochs} - Train Acc: {train_acc:.4f} - Val Acc: {val_acc:.4f}")

            # Salva melhor modelo
            if val_acc > best_val_acc:
                best_val_acc = val_acc
                torch.save(model.state_dict(), best_model_path)

        print(f"Melhor acuracia de validacao: {best_val_acc:.4f}")
        print(f"Modelo salvo em: {best_model_path}")

        return str(best_model_path)


def create_training_config(
    data_dir: str,
    output_dir: str,
    **kwargs
) -> TrainingConfig:
    """
    Cria configuracao de treinamento.

    Args:
        data_dir: Diretorio dos dados
        output_dir: Diretorio de saida
        **kwargs: Parametros adicionais

    Returns:
        TrainingConfig
    """
    return TrainingConfig(
        data_dir=data_dir,
        output_dir=output_dir,
        **kwargs
    )
