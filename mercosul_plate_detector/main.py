#!/usr/bin/env python3
"""
Script Principal - Sistema de Deteccao de Placas do Mercosul

Uso via linha de comando para detectar e analisar placas em imagens/videos.

Exemplos de uso:
    # Detectar placa em uma imagem
    python main.py detect --image caminho/para/imagem.jpg

    # Detectar placas em video
    python main.py detect --video caminho/para/video.mp4

    # Validar texto de placa
    python main.py validate --text "ABC1D23"

    # Treinar modelo customizado
    python main.py train --data-dir ./data --output-dir ./models

    # Avaliar modelo
    python main.py evaluate --model ./models/best.pt
"""

import argparse
import json
import sys
from pathlib import Path
from typing import Optional

try:
    import cv2
    CV2_AVAILABLE = True
except ImportError:
    CV2_AVAILABLE = False

from .detector import PlateDetector, PlateDetectionResult
from .validator import PlateValidator, validate_plate
from .classifier import PlateClassifier, classify_plate, PlateType
from .tampering import TamperingDetector, detect_tampering
from .trainer import PlateTrainer, TrainingConfig


def print_detection_result(result: PlateDetectionResult, verbose: bool = False):
    """Imprime resultado da deteccao de forma formatada"""
    print("\n" + "=" * 60)
    print("RESULTADO DA DETECCAO DE PLACA MERCOSUL")
    print("=" * 60)

    if not result.detected:
        print("\n[!] Nenhuma placa detectada na imagem")
        return

    print(f"\n[+] PLACA DETECTADA: {result.plate_text}")
    print(f"    Confianca OCR: {result.ocr_confidence:.2%}")
    print(f"    Confianca Deteccao: {result.detection_confidence:.2%}")

    # Validacao
    if result.validation:
        v = result.validation
        status_emoji = "✓" if v.status.value == "valida" else "⚠" if v.status.value == "suspeita" else "✗"
        print(f"\n[VALIDACAO] {status_emoji} Status: {v.status.value.upper()}")
        print(f"    Confianca: {v.confidence:.2%}")
        print(f"    Texto Normalizado: {v.normalized_text}")

        if v.issues:
            print("    Problemas encontrados:")
            for issue in v.issues:
                print(f"      - {issue}")

        if verbose and v.details:
            print("    Detalhes:")
            print(f"      - Formato Mercosul: {'Sim' if v.details.get('is_mercosul_format') else 'Nao'}")
            print(f"      - Letras: {v.details.get('letters', 'N/A')}")
            print(f"      - Numeros: {v.details.get('numbers', 'N/A')}")

    # Classificacao
    if result.classification:
        c = result.classification
        print(f"\n[CLASSIFICACAO] Tipo: {c.plate_type.value.upper()}")
        print(f"    Confianca: {c.confidence:.2%}")
        print(f"    Cor de Fundo: {c.detected_bg_color}")
        print(f"    Cor dos Caracteres: {c.detected_char_color}")

        if c.alternative_types:
            print("    Tipos alternativos:")
            for alt_type, alt_conf in c.alternative_types:
                print(f"      - {alt_type.value}: {alt_conf:.2%}")

    # Adulteracao
    if result.tampering:
        t = result.tampering
        tampering_emoji = "⚠ ADULTERADA" if t.is_tampered else "✓ AUTENTICA"
        print(f"\n[ADULTERACAO] {tampering_emoji}")
        print(f"    Score de Adulteracao: {t.tampering_score:.2%}")
        print(f"    Confianca Geral: {t.overall_confidence:.2%}")

        if t.indicators:
            print(f"    Indicadores detectados: {len(t.indicators)}")
            if verbose:
                for ind in t.indicators[:5]:  # Mostra apenas os 5 primeiros
                    print(f"      - {ind.type.value}: {ind.description} (severidade: {ind.severity})")

        if t.recommendations:
            print("    Recomendacoes:")
            for rec in t.recommendations:
                print(f"      - {rec}")

    print(f"\n[TEMPO] Processamento: {result.processing_time_ms:.2f}ms")
    print("=" * 60 + "\n")


def cmd_detect(args):
    """Comando para detectar placa em imagem ou video"""
    detector = PlateDetector(
        model_path=args.model,
        ocr_engine=args.ocr_engine,
        use_gpu=args.gpu,
        confidence_threshold=args.confidence
    )

    if args.image:
        print(f"Processando imagem: {args.image}")
        result = detector.detect(args.image)
        print_detection_result(result, verbose=args.verbose)

        # Salva resultado em JSON se solicitado
        if args.output:
            output_data = {
                'detected': result.detected,
                'plate_text': result.plate_text,
                'ocr_confidence': result.ocr_confidence,
                'detection_confidence': result.detection_confidence,
                'validation': {
                    'status': result.validation.status.value if result.validation else None,
                    'confidence': result.validation.confidence if result.validation else None,
                    'issues': result.validation.issues if result.validation else []
                } if result.validation else None,
                'classification': {
                    'type': result.classification.plate_type.value if result.classification else None,
                    'confidence': result.classification.confidence if result.classification else None
                } if result.classification else None,
                'tampering': {
                    'is_tampered': result.tampering.is_tampered if result.tampering else None,
                    'score': result.tampering.tampering_score if result.tampering else None
                } if result.tampering else None
            }

            with open(args.output, 'w') as f:
                json.dump(output_data, f, indent=2, ensure_ascii=False)
            print(f"Resultado salvo em: {args.output}")

        # Salva imagem anotada se solicitado
        if args.save_image and result.detected and CV2_AVAILABLE:
            import cv2
            image = cv2.imread(args.image)
            box = result.detection_box

            # Desenha bounding box
            cv2.rectangle(image, (box.x1, box.y1), (box.x2, box.y2), (0, 255, 0), 2)

            # Adiciona texto
            label = f"{result.plate_text} ({result.ocr_confidence:.0%})"
            cv2.putText(image, label, (box.x1, box.y1 - 10),
                       cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)

            output_path = args.save_image
            cv2.imwrite(output_path, image)
            print(f"Imagem anotada salva em: {output_path}")

    elif args.video:
        print(f"Processando video: {args.video}")
        results = detector.detect_from_video(
            args.video,
            frame_skip=args.frame_skip,
            max_frames=args.max_frames
        )

        print(f"\n{len(results)} placas detectadas no video:\n")
        for i, result in enumerate(results):
            print(f"{i+1}. {result.plate_text} (confianca: {result.ocr_confidence:.2%})")

        if args.output:
            output_data = [
                {
                    'plate_text': r.plate_text,
                    'ocr_confidence': r.ocr_confidence,
                    'validation_status': r.validation.status.value if r.validation else None,
                    'plate_type': r.classification.plate_type.value if r.classification else None,
                    'is_tampered': r.tampering.is_tampered if r.tampering else None
                }
                for r in results
            ]
            with open(args.output, 'w') as f:
                json.dump(output_data, f, indent=2, ensure_ascii=False)
            print(f"\nResultados salvos em: {args.output}")

    else:
        print("Erro: especifique --image ou --video")
        sys.exit(1)


def cmd_validate(args):
    """Comando para validar texto de placa"""
    result = validate_plate(args.text, ocr_confidence=args.confidence)

    print("\n" + "=" * 50)
    print("VALIDACAO DE PLACA MERCOSUL")
    print("=" * 50)
    print(f"\nTexto Original: {args.text}")
    print(f"Texto Normalizado: {result.normalized_text}")
    print(f"\nStatus: {result.status.value.upper()}")
    print(f"Confianca: {result.confidence:.2%}")

    if result.issues:
        print("\nProblemas encontrados:")
        for issue in result.issues:
            print(f"  - {issue}")

    print("\nDetalhes:")
    for key, value in result.details.items():
        print(f"  {key}: {value}")

    # Sugere correcoes se invalida
    if result.status.value != 'valida':
        validator = PlateValidator()
        suggestions = validator.suggest_corrections(args.text)
        if suggestions:
            print("\nSugestoes de correcao:")
            for sug in suggestions:
                print(f"  - {sug}")

    print("=" * 50 + "\n")


def cmd_classify(args):
    """Comando para classificar tipo de placa"""
    if args.image and CV2_AVAILABLE:
        import cv2
        image = cv2.imread(args.image)
        result = classify_plate(image=image, plate_text=args.text)
    else:
        result = classify_plate(plate_text=args.text)

    print("\n" + "=" * 50)
    print("CLASSIFICACAO DE PLACA MERCOSUL")
    print("=" * 50)
    print(f"\nTipo de Placa: {result.plate_type.value.upper()}")
    print(f"Confianca: {result.confidence:.2%}")
    print(f"Cor de Fundo Detectada: {result.detected_bg_color}")
    print(f"Cor dos Caracteres: {result.detected_char_color}")

    if result.alternative_types:
        print("\nTipos alternativos:")
        for alt_type, alt_conf in result.alternative_types:
            print(f"  - {alt_type.value}: {alt_conf:.2%}")

    print("=" * 50 + "\n")


def cmd_train(args):
    """Comando para treinar modelo"""
    config = TrainingConfig(
        data_dir=args.data_dir,
        output_dir=args.output_dir,
        model_name=args.model_name or 'mercosul_detector',
        epochs=args.epochs,
        batch_size=args.batch_size,
        image_size=args.image_size,
        base_model=args.base_model or 'yolov8n.pt'
    )

    trainer = PlateTrainer(config)

    print("\n" + "=" * 50)
    print("TREINAMENTO DE MODELO DE DETECCAO DE PLACAS")
    print("=" * 50)

    # Setup
    trainer.setup()

    # Prepara dados se fornecido diretorio de imagens
    if args.images_dir and args.labels_dir:
        print("\nPreparando dados...")
        trainer.prepare_data(args.images_dir, args.labels_dir)

    # Treina
    print("\nIniciando treinamento...")
    best_model = trainer.train_yolo(resume=args.resume)

    print(f"\nModelo treinado salvo em: {best_model}")

    # Avalia
    print("\nAvaliando modelo...")
    metrics = trainer.evaluate(best_model)

    print("\nMetricas finais:")
    print(f"  Precision: {metrics['precision']:.4f}")
    print(f"  Recall: {metrics['recall']:.4f}")
    print(f"  mAP@50: {metrics['mAP50']:.4f}")
    print(f"  mAP@50-95: {metrics['mAP50_95']:.4f}")

    # Exporta se solicitado
    if args.export_format:
        print(f"\nExportando para formato {args.export_format}...")
        exported = trainer.export_model(best_model, format=args.export_format)
        print(f"Modelo exportado: {exported}")

    print("=" * 50 + "\n")


def cmd_evaluate(args):
    """Comando para avaliar modelo"""
    config = TrainingConfig(
        data_dir=args.data_dir,
        output_dir='./eval_output'
    )

    trainer = PlateTrainer(config)
    trainer.data_yaml_path = trainer.dataset.create_data_yaml()

    print("\n" + "=" * 50)
    print("AVALIACAO DE MODELO")
    print("=" * 50)

    metrics = trainer.evaluate(args.model)

    print(f"\nModelo: {args.model}")
    print(f"\nMetricas:")
    print(f"  Precision: {metrics['precision']:.4f}")
    print(f"  Recall: {metrics['recall']:.4f}")
    print(f"  mAP@50: {metrics['mAP50']:.4f}")
    print(f"  mAP@50-95: {metrics['mAP50_95']:.4f}")

    print("=" * 50 + "\n")


def main():
    parser = argparse.ArgumentParser(
        description='Sistema de Deteccao de Placas do Mercosul',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )

    subparsers = parser.add_subparsers(dest='command', help='Comandos disponiveis')

    # Comando detect
    detect_parser = subparsers.add_parser('detect', help='Detectar placa em imagem/video')
    detect_parser.add_argument('--image', '-i', help='Caminho da imagem')
    detect_parser.add_argument('--video', '-v', help='Caminho do video')
    detect_parser.add_argument('--model', '-m', help='Caminho do modelo YOLO customizado')
    detect_parser.add_argument('--ocr-engine', default='easyocr',
                              choices=['easyocr', 'tesseract', 'both'],
                              help='Motor de OCR')
    detect_parser.add_argument('--confidence', '-c', type=float, default=0.5,
                              help='Limite de confianca para deteccao')
    detect_parser.add_argument('--gpu', action='store_true', help='Usar GPU')
    detect_parser.add_argument('--output', '-o', help='Arquivo JSON de saida')
    detect_parser.add_argument('--save-image', help='Salvar imagem anotada')
    detect_parser.add_argument('--verbose', action='store_true', help='Saida detalhada')
    detect_parser.add_argument('--frame-skip', type=int, default=5,
                              help='Processar a cada N frames (video)')
    detect_parser.add_argument('--max-frames', type=int, help='Maximo de frames (video)')

    # Comando validate
    validate_parser = subparsers.add_parser('validate', help='Validar texto de placa')
    validate_parser.add_argument('--text', '-t', required=True, help='Texto da placa')
    validate_parser.add_argument('--confidence', '-c', type=float, default=1.0,
                                help='Confianca do OCR (0-1)')

    # Comando classify
    classify_parser = subparsers.add_parser('classify', help='Classificar tipo de placa')
    classify_parser.add_argument('--text', '-t', help='Texto da placa')
    classify_parser.add_argument('--image', '-i', help='Imagem da placa')

    # Comando train
    train_parser = subparsers.add_parser('train', help='Treinar modelo')
    train_parser.add_argument('--data-dir', '-d', required=True, help='Diretorio dos dados')
    train_parser.add_argument('--output-dir', '-o', required=True, help='Diretorio de saida')
    train_parser.add_argument('--images-dir', help='Diretorio com imagens para preparar')
    train_parser.add_argument('--labels-dir', help='Diretorio com labels para preparar')
    train_parser.add_argument('--model-name', help='Nome do modelo')
    train_parser.add_argument('--epochs', type=int, default=100, help='Numero de epocas')
    train_parser.add_argument('--batch-size', type=int, default=16, help='Tamanho do batch')
    train_parser.add_argument('--image-size', type=int, default=640, help='Tamanho da imagem')
    train_parser.add_argument('--base-model', help='Modelo YOLO base')
    train_parser.add_argument('--resume', action='store_true', help='Continuar treinamento')
    train_parser.add_argument('--export-format', choices=['onnx', 'torchscript', 'tflite'],
                             help='Formato de exportacao')

    # Comando evaluate
    eval_parser = subparsers.add_parser('evaluate', help='Avaliar modelo')
    eval_parser.add_argument('--model', '-m', required=True, help='Caminho do modelo')
    eval_parser.add_argument('--data-dir', '-d', required=True, help='Diretorio dos dados')

    args = parser.parse_args()

    if args.command == 'detect':
        cmd_detect(args)
    elif args.command == 'validate':
        cmd_validate(args)
    elif args.command == 'classify':
        cmd_classify(args)
    elif args.command == 'train':
        cmd_train(args)
    elif args.command == 'evaluate':
        cmd_evaluate(args)
    else:
        parser.print_help()


if __name__ == '__main__':
    main()
