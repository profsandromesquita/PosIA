#!/usr/bin/env python3
"""
Exemplos de uso do Sistema de Deteccao de Placas do Mercosul

Este script demonstra as principais funcionalidades do sistema:
1. Validacao de placas
2. Classificacao de tipos
3. Deteccao de adulteracao
4. Deteccao completa em imagens
5. Geracao de dados sinteticos
"""

import sys
from pathlib import Path

# Adiciona o diretorio pai ao path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from mercosul_plate_detector.validator import PlateValidator, validate_plate
from mercosul_plate_detector.classifier import PlateClassifier, classify_plate, PlateType
from mercosul_plate_detector.tampering import TamperingDetector
from mercosul_plate_detector.detector import PlateDetector
from mercosul_plate_detector.utils.data_utils import (
    generate_plate_text,
    generate_synthetic_plate,
    generate_synthetic_dataset
)


def exemplo_validacao():
    """Exemplo de validacao de placas"""
    print("\n" + "=" * 60)
    print("EXEMPLO 1: VALIDACAO DE PLACAS")
    print("=" * 60)

    # Cria validador
    validator = PlateValidator()

    # Testa diferentes placas
    placas_teste = [
        "ABC1D23",      # Valida - formato Mercosul
        "XYZ9A99",      # Valida - formato Mercosul
        "ABC1234",      # Invalida - formato antigo
        "AB1CD23",      # Invalida - formato incorreto
        "ABC-1D23",     # Sera normalizada
        "abc1d23",      # Sera normalizada para maiusculas
        "O0I1Z2",       # Confusao de caracteres
    ]

    for placa in placas_teste:
        resultado = validator.validate(placa)

        print(f"\nPlaca: {placa}")
        print(f"  Normalizada: {resultado.normalized_text}")
        print(f"  Status: {resultado.status.value}")
        print(f"  Confianca: {resultado.confidence:.2%}")

        if resultado.issues:
            print("  Problemas:")
            for issue in resultado.issues:
                print(f"    - {issue}")

        # Sugere correcoes se invalida
        if resultado.status.value != 'valida':
            sugestoes = validator.suggest_corrections(placa)
            if sugestoes:
                print(f"  Sugestoes: {', '.join(sugestoes)}")


def exemplo_classificacao():
    """Exemplo de classificacao de tipos de placa"""
    print("\n" + "=" * 60)
    print("EXEMPLO 2: CLASSIFICACAO DE TIPOS DE PLACA")
    print("=" * 60)

    classifier = PlateClassifier()

    # Gera placas sinteticas de diferentes tipos
    tipos = ['particular', 'comercial', 'oficial', 'especial', 'colecionador']

    for tipo in tipos:
        # Gera imagem sintetica
        placa_texto = generate_plate_text()
        placa_imagem = generate_synthetic_plate(placa_texto, tipo)

        # Classifica
        resultado = classifier.classify(placa_imagem, placa_texto)

        print(f"\nPlaca: {placa_texto} (tipo real: {tipo})")
        print(f"  Tipo detectado: {resultado.plate_type.value}")
        print(f"  Confianca: {resultado.confidence:.2%}")
        print(f"  Cor de fundo: {resultado.detected_bg_color}")
        print(f"  Cor caracteres: {resultado.detected_char_color}")


def exemplo_adulteracao():
    """Exemplo de deteccao de adulteracao"""
    print("\n" + "=" * 60)
    print("EXEMPLO 3: DETECCAO DE ADULTERACAO")
    print("=" * 60)

    detector = TamperingDetector(sensitivity='media')

    # Gera placa normal
    placa_normal = generate_synthetic_plate("ABC1D23", "particular")

    # Analisa
    resultado = detector.analyze(placa_normal)

    print("\nPlaca normal:")
    print(f"  Adulterada: {'Sim' if resultado.is_tampered else 'Nao'}")
    print(f"  Score de adulteracao: {resultado.tampering_score:.2%}")
    print(f"  Indicadores: {len(resultado.indicators)}")

    # Simula placa adulterada (adicionando ruido artificial)
    import numpy as np
    try:
        import cv2
        placa_adulterada = placa_normal.copy()
        # Adiciona um "adesivo" simulado
        placa_adulterada[40:90, 50:100] = np.random.randint(200, 255, (50, 50, 3), dtype=np.uint8)

        resultado_adulterada = detector.analyze(placa_adulterada)

        print("\nPlaca com alteracao:")
        print(f"  Adulterada: {'Sim' if resultado_adulterada.is_tampered else 'Nao'}")
        print(f"  Score de adulteracao: {resultado_adulterada.tampering_score:.2%}")
        print(f"  Indicadores: {len(resultado_adulterada.indicators)}")

        if resultado_adulterada.recommendations:
            print("  Recomendacoes:")
            for rec in resultado_adulterada.recommendations:
                print(f"    - {rec}")
    except ImportError:
        print("  (OpenCV necessario para simulacao de adulteracao)")


def exemplo_deteccao_completa():
    """Exemplo de deteccao completa em imagem"""
    print("\n" + "=" * 60)
    print("EXEMPLO 4: DETECCAO COMPLETA")
    print("=" * 60)

    try:
        import cv2
        import numpy as np

        # Cria uma imagem de teste com placa
        imagem = np.zeros((400, 600, 3), dtype=np.uint8)
        imagem[:] = (100, 100, 100)  # Fundo cinza

        # Adiciona placa
        placa = generate_synthetic_plate("BRA2E19", "particular")
        ph, pw = placa.shape[:2]
        x, y = 150, 150
        imagem[y:y+ph, x:x+pw] = placa

        # Salva imagem temporaria
        temp_path = "/tmp/teste_placa.jpg"
        cv2.imwrite(temp_path, imagem)

        print(f"\nImagem de teste criada: {temp_path}")

        # Detecta
        detector = PlateDetector(
            ocr_engine='easyocr',
            enable_validation=True,
            enable_classification=True,
            enable_tampering_detection=True
        )

        resultado = detector.detect(temp_path)

        if resultado.detected:
            print(f"\n[+] Placa detectada: {resultado.plate_text}")
            print(f"    Confianca OCR: {resultado.ocr_confidence:.2%}")

            if resultado.validation:
                print(f"    Validacao: {resultado.validation.status.value}")

            if resultado.classification:
                print(f"    Tipo: {resultado.classification.plate_type.value}")

            if resultado.tampering:
                print(f"    Adulteracao: {'Sim' if resultado.tampering.is_tampered else 'Nao'}")

            print(f"    Tempo de processamento: {resultado.processing_time_ms:.2f}ms")
        else:
            print("\n[!] Nenhuma placa detectada")

    except ImportError:
        print("  (OpenCV necessario para este exemplo)")


def exemplo_geracao_dataset():
    """Exemplo de geracao de dataset sintetico"""
    print("\n" + "=" * 60)
    print("EXEMPLO 5: GERACAO DE DATASET SINTETICO")
    print("=" * 60)

    output_dir = "/tmp/mercosul_dataset_demo"

    print(f"\nGerando dataset sintetico em: {output_dir}")

    generate_synthetic_dataset(
        output_dir=output_dir,
        num_images=10,  # Apenas 10 para demonstracao
        plate_types=['particular', 'comercial', 'oficial'],
        with_augmentation=True
    )

    # Lista arquivos gerados
    from pathlib import Path
    output_path = Path(output_dir)

    images = list((output_path / 'images').glob('*.jpg'))
    labels = list((output_path / 'labels').glob('*.txt'))

    print(f"  Imagens geradas: {len(images)}")
    print(f"  Labels geradas: {len(labels)}")

    if images:
        print(f"\n  Exemplo de arquivo de imagem: {images[0].name}")

    if labels:
        print(f"  Exemplo de label:")
        print(f"    {labels[0].read_text()}")


def exemplo_api_simples():
    """Exemplo de uso da API simples"""
    print("\n" + "=" * 60)
    print("EXEMPLO 6: API SIMPLES")
    print("=" * 60)

    # Validacao rapida
    resultado = validate_plate("ABC1D23")
    print(f"\nvalidate_plate('ABC1D23'):")
    print(f"  Status: {resultado.status.value}")
    print(f"  Confianca: {resultado.confidence:.2%}")

    # Classificacao rapida (apenas por texto)
    resultado = classify_plate(plate_text="PF1A23")
    print(f"\nclassify_plate('PF1A23'):")
    print(f"  Tipo: {resultado.plate_type.value}")

    # Geracao de texto de placa aleatorio
    for _ in range(5):
        placa = generate_plate_text()
        print(f"\nPlaca aleatoria: {placa}")


def main():
    """Executa todos os exemplos"""
    print("\n" + "#" * 60)
    print("# SISTEMA DE DETECCAO DE PLACAS DO MERCOSUL - EXEMPLOS #")
    print("#" * 60)

    exemplo_validacao()
    exemplo_classificacao()
    exemplo_adulteracao()
    exemplo_api_simples()

    # Exemplos que requerem OpenCV
    try:
        import cv2
        exemplo_deteccao_completa()
        exemplo_geracao_dataset()
    except ImportError:
        print("\n[AVISO] Alguns exemplos foram pulados pois OpenCV nao esta instalado")

    print("\n" + "=" * 60)
    print("FIM DOS EXEMPLOS")
    print("=" * 60 + "\n")


if __name__ == '__main__':
    main()
