"""
Sistema de Detecao e Analise de Placas do Mercosul

Este modulo fornece funcionalidades para:
- Detecao de placas em imagens
- Leitura OCR do texto da placa
- Validacao do formato da placa Mercosul
- Classificacao do tipo de placa
- Deteccao de adulteracao

Autor: Claude AI
Data: 2026-01-24
"""

from .validator import PlateValidator
from .classifier import PlateClassifier
from .detector import PlateDetector
from .tampering import TamperingDetector
from .trainer import PlateTrainer

__version__ = "1.0.0"
__all__ = [
    "PlateValidator",
    "PlateClassifier",
    "PlateDetector",
    "TamperingDetector",
    "PlateTrainer"
]
