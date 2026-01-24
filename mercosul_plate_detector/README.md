# Sistema de Deteccao de Placas do Mercosul

Sistema completo em Python para deteccao, validacao, classificacao e analise de placas de veiculos no padrao Mercosul (Brasil).

## Funcionalidades

- **Deteccao de Placas**: Detecta placas em imagens e videos usando YOLOv8 ou metodos tradicionais de visao computacional
- **OCR**: Extrai o texto da placa usando EasyOCR ou Tesseract
- **Validacao**: Valida se a placa esta no formato correto (LLLNLNN) e identifica problemas
- **Classificacao de Tipo**: Identifica o tipo da placa (particular, comercial, oficial, diplomatico, especial, colecionador)
- **Deteccao de Adulteracao**: Analisa a placa em busca de sinais de adulteracao
- **Treinamento**: Permite treinar modelos customizados com seus proprios dados

## Instalacao

```bash
# Clone o repositorio
git clone https://github.com/profsandromesquita/PosIA.git
cd PosIA

# Instale as dependencias
pip install -r mercosul_plate_detector/requirements.txt

# Instale o Tesseract (opcional, para OCR alternativo)
# Ubuntu/Debian:
sudo apt-get install tesseract-ocr tesseract-ocr-por

# macOS:
brew install tesseract
```

## Uso Rapido

### Via Python

```python
from mercosul_plate_detector import PlateDetector, PlateValidator, PlateClassifier

# Validar uma placa
from mercosul_plate_detector.validator import validate_plate

resultado = validate_plate("ABC1D23")
print(f"Status: {resultado.status.value}")  # valida, suspeita, ou invalida
print(f"Confianca: {resultado.confidence:.2%}")

# Detectar placa em imagem
detector = PlateDetector(ocr_engine='easyocr')
resultado = detector.detect("caminho/para/imagem.jpg")

if resultado.detected:
    print(f"Placa: {resultado.plate_text}")
    print(f"Tipo: {resultado.classification.plate_type.value}")
    print(f"Adulterada: {resultado.tampering.is_tampered}")
```

### Via Linha de Comando

```bash
# Detectar placa em imagem
python -m mercosul_plate_detector.main detect --image foto.jpg

# Validar texto de placa
python -m mercosul_plate_detector.main validate --text "ABC1D23"

# Classificar tipo de placa
python -m mercosul_plate_detector.main classify --text "ABC1D23" --image placa.jpg

# Treinar modelo customizado
python -m mercosul_plate_detector.main train \
    --data-dir ./meus_dados \
    --output-dir ./meu_modelo \
    --epochs 100
```

## Estrutura do Projeto

```
mercosul_plate_detector/
├── __init__.py          # Modulo principal
├── main.py              # CLI
├── validator.py         # Validacao de placas
├── classifier.py        # Classificacao de tipos
├── tampering.py         # Deteccao de adulteracao
├── detector.py          # Deteccao e OCR
├── trainer.py           # Treinamento de modelos
├── requirements.txt     # Dependencias
├── utils/
│   ├── image_utils.py   # Utilitarios de imagem
│   └── data_utils.py    # Utilitarios de dados
├── examples/
│   └── example_usage.py # Exemplos de uso
├── models/              # Modelos treinados
├── weights/             # Pesos dos modelos
└── data/                # Dados de treinamento
```

## Formato das Placas Mercosul

O sistema suporta o padrao de placas Mercosul brasileiro:

- **Formato**: `LLLNLNN` (3 letras + 1 numero + 1 letra + 2 numeros)
- **Exemplo**: `ABC1D23`

### Tipos de Placa

| Tipo | Cor de Fundo | Cor dos Caracteres | Descricao |
|------|--------------|-------------------|-----------|
| Particular | Branco | Preto | Veiculos particulares |
| Comercial | Vermelho | Branco | Veiculos de aluguel/comerciais |
| Oficial | Azul | Branco | Veiculos oficiais do governo |
| Diplomatico | Azul | Branco | Veiculos do corpo diplomatico |
| Especial | Verde | Branco | Taxis, transporte escolar |
| Colecionador | Cinza | Preto | Veiculos antigos/colecao |

## Treinamento de Modelo Customizado

### Preparando os Dados

1. Organize suas imagens e anotacoes no formato YOLO:

```
meus_dados/
├── images/
│   ├── train/
│   ├── val/
│   └── test/
└── labels/
    ├── train/
    ├── val/
    └── test/
```

2. Formato do arquivo de label (YOLO):
```
# class_id x_center y_center width height (valores normalizados 0-1)
0 0.5 0.5 0.3 0.1
```

### Gerando Dataset Sintetico

```python
from mercosul_plate_detector.utils.data_utils import generate_synthetic_dataset

generate_synthetic_dataset(
    output_dir="./dataset_sintetico",
    num_images=1000,
    plate_types=['particular', 'comercial', 'oficial'],
    with_augmentation=True
)
```

### Treinando o Modelo

```python
from mercosul_plate_detector.trainer import PlateTrainer, TrainingConfig

config = TrainingConfig(
    data_dir="./meus_dados",
    output_dir="./meu_modelo",
    epochs=100,
    batch_size=16,
    image_size=640
)

trainer = PlateTrainer(config)
trainer.setup()
trainer.prepare_data("./imagens", "./labels")
best_model = trainer.train_yolo()

# Avalia o modelo
metrics = trainer.evaluate(best_model)
print(f"mAP@50: {metrics['mAP50']:.4f}")

# Exporta para producao
trainer.export_model(best_model, format='onnx')
```

## Deteccao de Adulteracao

O sistema analisa diversos indicadores de adulteracao:

- **Inconsistencia de cor**: Diferenca de cor entre caracteres
- **Anomalia de textura**: Padroes irregulares indicando adesivos
- **Bordas irregulares**: Contornos de caracteres alterados
- **Inconsistencia de fonte**: Caracteres com fontes diferentes
- **Problemas de alinhamento**: Caracteres desalinhados

```python
from mercosul_plate_detector.tampering import TamperingDetector
import cv2

detector = TamperingDetector(sensitivity='alta')
imagem = cv2.imread("placa_suspeita.jpg")

resultado = detector.analyze(imagem)

if resultado.is_tampered:
    print("ALERTA: Placa possivelmente adulterada!")
    print(f"Score de adulteracao: {resultado.tampering_score:.2%}")

    for indicador in resultado.indicators:
        print(f"  - {indicador.type.value}: {indicador.description}")
```

## API de Alto Nivel

### PlateDetector

```python
detector = PlateDetector(
    model_path="./modelo_customizado.pt",  # Modelo YOLO customizado
    ocr_engine='easyocr',                   # 'easyocr', 'tesseract', ou 'both'
    use_gpu=True,                           # Usar GPU
    confidence_threshold=0.5,               # Limite de confianca
    enable_validation=True,                 # Habilitar validacao
    enable_classification=True,             # Habilitar classificacao
    enable_tampering_detection=True         # Habilitar deteccao de adulteracao
)

# Detectar em imagem
resultado = detector.detect("imagem.jpg")

# Detectar em video
resultados = detector.detect_from_video(
    "video.mp4",
    frame_skip=5,      # Processar a cada 5 frames
    max_frames=1000    # Maximo de frames
)
```

### PlateValidator

```python
validator = PlateValidator(strict_mode=True)

resultado = validator.validate("ABC1D23", ocr_confidence=0.95)
print(resultado.status)      # ValidationStatus.VALID
print(resultado.confidence)  # 0.95
print(resultado.issues)      # Lista de problemas encontrados

# Sugerir correcoes
sugestoes = validator.suggest_corrections("AB01D23")
print(sugestoes)  # ['ABO1D23'] - O foi confundido com 0
```

### PlateClassifier

```python
classifier = PlateClassifier()

# Classificar por imagem
resultado = classifier.classify(image=imagem)

# Classificar por texto (padroes especiais)
resultado = classifier.classify(plate_text="CMD1A23")  # Diplomatico

print(resultado.plate_type)       # PlateType.DIPLOMATICO
print(resultado.detected_bg_color)  # 'azul'
```

## Dependencias

- Python 3.8+
- OpenCV 4.5+
- PyTorch 1.9+
- Ultralytics (YOLOv8)
- EasyOCR ou Tesseract
- NumPy
- Pillow

## Limitacoes

- O sistema e otimizado para placas brasileiras do padrao Mercosul
- A deteccao de adulteracao e heuristica e pode ter falsos positivos/negativos
- O desempenho do OCR depende da qualidade da imagem
- Para melhor precisao, recomenda-se treinar um modelo customizado com dados reais

## Contribuicao

Contribuicoes sao bem-vindas! Por favor:

1. Fork o repositorio
2. Crie uma branch para sua feature (`git checkout -b feature/nova-feature`)
3. Commit suas alteracoes (`git commit -am 'Adiciona nova feature'`)
4. Push para a branch (`git push origin feature/nova-feature`)
5. Abra um Pull Request

## Licenca

Este projeto esta licenciado sob a Licenca MIT - veja o arquivo LICENSE para detalhes.

## Autor

Desenvolvido como parte do projeto PosIA.
