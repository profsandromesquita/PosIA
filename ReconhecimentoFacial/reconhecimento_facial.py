"""
Reconhecimento Facial com OpenCV
================================
Sistema de cadastro e reconhecimento facial usando OpenCV e LBPH.

Funcionalidades:
  1. Cadastrar rosto a partir de imagem (arquivo)
  2. Cadastrar rosto via webcam (captura ao vivo)
  3. Treinar o modelo de reconhecimento
  4. Reconhecimento facial em tempo real via webcam

Estrutura de pastas:
  dataset/         -> fotos cadastradas (uma subpasta por pessoa)
  modelo/          -> modelo treinado salvo em disco

Autor: Pós-graduação IA - UniAteneu
"""

import os
import sys
import platform
import glob as glob_mod
import cv2
import numpy as np

# ---------------------------------------------------------------------------
# Configuracoes
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATASET_DIR = os.path.join(BASE_DIR, "dataset")
MODELO_DIR = os.path.join(BASE_DIR, "modelo")
MODELO_PATH = os.path.join(MODELO_DIR, "modelo_lbph.yml")
LABELS_PATH = os.path.join(MODELO_DIR, "labels.npy")

# Classificador Haar Cascade para deteccao de rosto
HAAR_CASCADE = cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
face_cascade = cv2.CascadeClassifier(HAAR_CASCADE)

# Parametros de deteccao
SCALE_FACTOR = 1.3
MIN_NEIGHBORS = 5
MIN_SIZE = (80, 80)

# Dimensao padrao para as imagens de treino
FACE_WIDTH = 200
FACE_HEIGHT = 200

# Backends de captura para tentar (em ordem de prioridade por SO)
if platform.system() == "Windows":
    BACKENDS = [
        ("DirectShow", cv2.CAP_DSHOW),
        ("MSMF", cv2.CAP_MSMF),
        ("Default", cv2.CAP_ANY),
    ]
elif platform.system() == "Linux":
    BACKENDS = [
        ("V4L2", cv2.CAP_V4L2),
        ("GStreamer", cv2.CAP_GSTREAMER),
        ("Default", cv2.CAP_ANY),
    ]
else:  # macOS
    BACKENDS = [
        ("AVFoundation", cv2.CAP_AVFOUNDATION),
        ("Default", cv2.CAP_ANY),
    ]

# Indices de camera para tentar
CAMERA_INDICES = [0, 1, 2]


# ---------------------------------------------------------------------------
# Funcoes auxiliares
# ---------------------------------------------------------------------------
def garantir_diretorios():
    """Cria os diretorios necessarios se nao existirem."""
    os.makedirs(DATASET_DIR, exist_ok=True)
    os.makedirs(MODELO_DIR, exist_ok=True)


def detectar_rostos(imagem_gray):
    """Detecta rostos em uma imagem em escala de cinza."""
    return face_cascade.detectMultiScale(
        imagem_gray,
        scaleFactor=SCALE_FACTOR,
        minNeighbors=MIN_NEIGHBORS,
        minSize=MIN_SIZE,
    )


def recortar_rosto(imagem_gray, x, y, w, h):
    """Recorta e redimensiona um rosto detectado."""
    rosto = imagem_gray[y : y + h, x : x + w]
    return cv2.resize(rosto, (FACE_WIDTH, FACE_HEIGHT))


def diagnosticar_camera():
    """Exibe informacoes de diagnostico sobre cameras disponiveis."""
    print("\n[DIAGNOSTICO] Verificando cameras disponiveis...")
    print(f"  Sistema Operacional: {platform.system()} {platform.release()}")
    print(f"  OpenCV versao: {cv2.__version__}")

    # No Linux, verificar dispositivos de video
    if platform.system() == "Linux":
        dispositivos = glob_mod.glob("/dev/video*")
        if dispositivos:
            print(f"  Dispositivos encontrados: {', '.join(sorted(dispositivos))}")
            for dev in sorted(dispositivos):
                permissao = os.access(dev, os.R_OK)
                print(f"    {dev} - leitura: {'OK' if permissao else 'SEM PERMISSAO'}")
            if not all(os.access(d, os.R_OK) for d in dispositivos):
                print("\n  [DICA] Para liberar permissao, execute:")
                print("    sudo usermod -aG video $USER")
                print("    (depois faca logout e login novamente)")
        else:
            print("  Nenhum dispositivo /dev/video* encontrado.")
            print("\n  [DICA] Possiveis solucoes:")
            print("    - Verifique se a webcam esta conectada")
            print("    - Tente: sudo modprobe uvcvideo")
            print("    - Verifique: lsusb | grep -i cam")

    # No Windows
    elif platform.system() == "Windows":
        print("  [DICA] Possiveis solucoes:")
        print("    - Verifique se a webcam esta ativada no Gerenciador de Dispositivos")
        print("    - Feche outros programas que usam a camera (Teams, Zoom, etc.)")
        print("    - Verifique Configuracoes > Privacidade > Camera")

    # Tentar abrir com cada combinacao de indice e backend
    print("\n  Tentando combinacoes de indice/backend:")
    for idx in CAMERA_INDICES:
        for nome_backend, backend in BACKENDS:
            cap = cv2.VideoCapture(idx, backend)
            if cap.isOpened():
                ret, frame = cap.read()
                if ret and frame is not None:
                    h, w = frame.shape[:2]
                    print(f"    [OK] Indice {idx} + {nome_backend}: {w}x{h}")
                else:
                    print(f"    [--] Indice {idx} + {nome_backend}: abriu mas sem frame")
                cap.release()
            else:
                print(f"    [--] Indice {idx} + {nome_backend}: nao abriu")
    print()


def abrir_camera():
    """
    Tenta abrir a webcam testando multiplos indices e backends.
    Retorna (cap, descricao) ou (None, mensagem_de_erro).
    """
    for idx in CAMERA_INDICES:
        for nome_backend, backend in BACKENDS:
            cap = cv2.VideoCapture(idx, backend)
            if cap.isOpened():
                ret, frame = cap.read()
                if ret and frame is not None:
                    print(f"[INFO] Camera aberta: indice {idx}, backend {nome_backend}")
                    return cap, f"indice {idx} ({nome_backend})"
                cap.release()

    return None, "Nenhuma camera disponivel"


def carregar_mapa_labels():
    """Retorna o mapeamento id -> nome a partir das pastas do dataset."""
    nomes = sorted(
        [d for d in os.listdir(DATASET_DIR)
         if os.path.isdir(os.path.join(DATASET_DIR, d))]
    )
    return {i: nome for i, nome in enumerate(nomes)}


# ---------------------------------------------------------------------------
# 1. Cadastro via arquivo de imagem
# ---------------------------------------------------------------------------
def cadastrar_por_imagem(nome, caminho_imagem):
    """
    Cadastra um rosto a partir de um arquivo de imagem.

    Args:
        nome: nome da pessoa
        caminho_imagem: caminho para o arquivo de imagem
    """
    garantir_diretorios()

    if not os.path.isfile(caminho_imagem):
        print(f"[ERRO] Arquivo nao encontrado: {caminho_imagem}")
        return False

    imagem = cv2.imread(caminho_imagem)
    if imagem is None:
        print(f"[ERRO] Nao foi possivel ler a imagem: {caminho_imagem}")
        return False

    gray = cv2.cvtColor(imagem, cv2.COLOR_BGR2GRAY)
    rostos = detectar_rostos(gray)

    if len(rostos) == 0:
        print("[ERRO] Nenhum rosto detectado na imagem.")
        return False

    if len(rostos) > 1:
        print(f"[AVISO] {len(rostos)} rostos detectados. Usando o maior.")

    # Seleciona o maior rosto (maior area)
    x, y, w, h = max(rostos, key=lambda r: r[2] * r[3])
    rosto = recortar_rosto(gray, x, y, w, h)

    # Salvar
    pasta_pessoa = os.path.join(DATASET_DIR, nome)
    os.makedirs(pasta_pessoa, exist_ok=True)
    qtd = len(os.listdir(pasta_pessoa))
    nome_arquivo = f"{nome}_{qtd + 1}.jpg"
    caminho_salvar = os.path.join(pasta_pessoa, nome_arquivo)
    cv2.imwrite(caminho_salvar, rosto)

    print(f"[OK] Rosto cadastrado: {caminho_salvar}")

    # Mostrar preview
    imagem_preview = imagem.copy()
    cv2.rectangle(imagem_preview, (x, y), (x + w, y + h), (0, 255, 0), 2)
    cv2.putText(
        imagem_preview, nome, (x, y - 10),
        cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 255, 0), 2,
    )
    cv2.imshow("Rosto Cadastrado", imagem_preview)
    cv2.waitKey(2000)
    cv2.destroyAllWindows()

    return True


# ---------------------------------------------------------------------------
# 2. Cadastro via webcam
# ---------------------------------------------------------------------------
def cadastrar_por_webcam(nome, num_capturas=10):
    """
    Cadastra rostos capturando imagens pela webcam.

    Args:
        nome: nome da pessoa
        num_capturas: quantidade de fotos a capturar (default: 10)
    """
    garantir_diretorios()

    pasta_pessoa = os.path.join(DATASET_DIR, nome)
    os.makedirs(pasta_pessoa, exist_ok=True)
    qtd_existente = len(os.listdir(pasta_pessoa))

    cap, desc = abrir_camera()
    if cap is None:
        print("[ERRO] Nao foi possivel abrir a webcam.")
        print("[INFO] Execute a opcao 6 (Diagnosticar camera) para mais detalhes.")
        return False

    print(f"[INFO] Capturando {num_capturas} fotos de '{nome}'.")
    print("[INFO] Pressione 'q' para cancelar.")

    capturas = 0
    while capturas < num_capturas:
        ret, frame = cap.read()
        if not ret:
            print("[ERRO] Falha ao capturar frame da webcam.")
            break

        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        rostos = detectar_rostos(gray)

        for (x, y, w, h) in rostos:
            cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
            progresso = f"{capturas + 1}/{num_capturas}"
            cv2.putText(
                frame, f"{nome} - {progresso}", (x, y - 10),
                cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2,
            )

        cv2.imshow("Cadastro - Webcam (q = cancelar)", frame)

        # Captura automatica quando detecta exatamente 1 rosto
        if len(rostos) == 1:
            x, y, w, h = rostos[0]
            rosto = recortar_rosto(gray, x, y, w, h)
            capturas += 1
            nome_arquivo = f"{nome}_{qtd_existente + capturas}.jpg"
            caminho = os.path.join(pasta_pessoa, nome_arquivo)
            cv2.imwrite(caminho, rosto)
            print(f"  Captura {capturas}/{num_capturas}: {nome_arquivo}")
            cv2.waitKey(300)  # Pequena pausa entre capturas

        if cv2.waitKey(1) & 0xFF == ord("q"):
            print("[INFO] Cadastro cancelado pelo usuario.")
            break

    cap.release()
    cv2.destroyAllWindows()

    if capturas > 0:
        print(f"[OK] {capturas} fotos salvas para '{nome}'.")
        return True
    else:
        print("[AVISO] Nenhuma foto capturada.")
        return False


# ---------------------------------------------------------------------------
# 3. Treinamento do modelo
# ---------------------------------------------------------------------------
def treinar_modelo():
    """
    Treina o modelo LBPH com todas as imagens cadastradas no dataset.
    Salva o modelo e o mapeamento de labels em disco.
    """
    garantir_diretorios()

    mapa_labels = carregar_mapa_labels()
    if not mapa_labels:
        print("[ERRO] Nenhuma pessoa cadastrada no dataset.")
        return False

    nome_para_id = {nome: id_ for id_, nome in mapa_labels.items()}

    faces = []
    labels = []

    for nome, id_ in nome_para_id.items():
        pasta = os.path.join(DATASET_DIR, nome)
        for arquivo in os.listdir(pasta):
            caminho = os.path.join(pasta, arquivo)
            imagem = cv2.imread(caminho, cv2.IMREAD_GRAYSCALE)
            if imagem is None:
                continue
            imagem = cv2.resize(imagem, (FACE_WIDTH, FACE_HEIGHT))
            faces.append(imagem)
            labels.append(id_)

    if len(faces) == 0:
        print("[ERRO] Nenhuma imagem valida encontrada no dataset.")
        return False

    print(f"[INFO] Treinando com {len(faces)} imagens de {len(nome_para_id)} pessoa(s)...")

    recognizer = cv2.face.LBPHFaceRecognizer_create()
    recognizer.train(faces, np.array(labels))
    recognizer.save(MODELO_PATH)

    # Salvar mapeamento de labels
    np.save(LABELS_PATH, mapa_labels)

    print(f"[OK] Modelo treinado e salvo em: {MODELO_PATH}")
    print(f"[OK] Pessoas cadastradas: {', '.join(nome_para_id.keys())}")
    return True


# ---------------------------------------------------------------------------
# 4. Reconhecimento facial em tempo real
# ---------------------------------------------------------------------------
def reconhecer_em_tempo_real(confianca_minima=70):
    """
    Executa reconhecimento facial em tempo real via webcam.

    Args:
        confianca_minima: limiar de confianca (quanto menor, mais restrito).
                          Valores tipicos: 50-80. Default: 70.
    """
    if not os.path.isfile(MODELO_PATH):
        print("[ERRO] Modelo nao encontrado. Execute o treinamento primeiro.")
        return

    if not os.path.isfile(LABELS_PATH):
        print("[ERRO] Arquivo de labels nao encontrado.")
        return

    # Carregar modelo e labels
    recognizer = cv2.face.LBPHFaceRecognizer_create()
    recognizer.read(MODELO_PATH)
    mapa_labels = np.load(LABELS_PATH, allow_pickle=True).item()

    cap, desc = abrir_camera()
    if cap is None:
        print("[ERRO] Nao foi possivel abrir a webcam.")
        print("[INFO] Execute a opcao 6 (Diagnosticar camera) para mais detalhes.")
        return

    print("[INFO] Reconhecimento facial iniciado.")
    print("[INFO] Pressione 'q' para sair.")

    while True:
        ret, frame = cap.read()
        if not ret:
            break

        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        rostos = detectar_rostos(gray)

        for (x, y, w, h) in rostos:
            rosto = recortar_rosto(gray, x, y, w, h)

            id_previsto, confianca = recognizer.predict(rosto)

            # LBPH: confianca = distancia (quanto menor, melhor)
            porcentagem = max(0, 100 - confianca)

            if porcentagem >= confianca_minima:
                nome = mapa_labels.get(id_previsto, "Desconhecido")
                cor = (0, 255, 0)  # Verde
                texto = f"{nome} ({porcentagem:.0f}%)"
            else:
                cor = (0, 0, 255)  # Vermelho
                texto = f"Desconhecido ({porcentagem:.0f}%)"

            cv2.rectangle(frame, (x, y), (x + w, y + h), cor, 2)
            cv2.putText(
                frame, texto, (x, y - 10),
                cv2.FONT_HERSHEY_SIMPLEX, 0.7, cor, 2,
            )

        # Instrucoes na tela
        cv2.putText(
            frame, "Pressione 'q' para sair", (10, 30),
            cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 1,
        )

        cv2.imshow("Reconhecimento Facial", frame)

        if cv2.waitKey(1) & 0xFF == ord("q"):
            break

    cap.release()
    cv2.destroyAllWindows()
    print("[INFO] Reconhecimento finalizado.")


# ---------------------------------------------------------------------------
# 5. Listar pessoas cadastradas
# ---------------------------------------------------------------------------
def listar_cadastrados():
    """Lista todas as pessoas cadastradas no dataset."""
    garantir_diretorios()
    mapa = carregar_mapa_labels()
    if not mapa:
        print("[INFO] Nenhuma pessoa cadastrada ainda.")
        return

    print("\n=== Pessoas Cadastradas ===")
    for id_, nome in mapa.items():
        pasta = os.path.join(DATASET_DIR, nome)
        qtd = len([f for f in os.listdir(pasta) if f.endswith(".jpg")])
        print(f"  ID {id_}: {nome} ({qtd} foto(s))")
    print()


# ---------------------------------------------------------------------------
# Menu interativo
# ---------------------------------------------------------------------------
def menu():
    """Menu principal do sistema."""
    garantir_diretorios()

    while True:
        print("\n" + "=" * 50)
        print("  SISTEMA DE RECONHECIMENTO FACIAL")
        print("  OpenCV + LBPH Face Recognizer")
        print("=" * 50)
        print("  1. Cadastrar rosto (arquivo de imagem)")
        print("  2. Cadastrar rosto (webcam)")
        print("  3. Treinar modelo")
        print("  4. Reconhecimento em tempo real")
        print("  5. Listar pessoas cadastradas")
        print("  6. Diagnosticar camera")
        print("  0. Sair")
        print("=" * 50)

        opcao = input("  Escolha uma opcao: ").strip()

        if opcao == "1":
            nome = input("  Nome da pessoa: ").strip()
            if not nome:
                print("[ERRO] Nome nao pode ser vazio.")
                continue
            caminho = input("  Caminho da imagem: ").strip()
            cadastrar_por_imagem(nome, caminho)

        elif opcao == "2":
            nome = input("  Nome da pessoa: ").strip()
            if not nome:
                print("[ERRO] Nome nao pode ser vazio.")
                continue
            try:
                n = input("  Quantidade de capturas (default=10): ").strip()
                num = int(n) if n else 10
            except ValueError:
                num = 10
            cadastrar_por_webcam(nome, num)

        elif opcao == "3":
            treinar_modelo()

        elif opcao == "4":
            try:
                c = input("  Confianca minima % (default=70): ").strip()
                conf = float(c) if c else 70
            except ValueError:
                conf = 70
            reconhecer_em_tempo_real(conf)

        elif opcao == "5":
            listar_cadastrados()

        elif opcao == "6":
            diagnosticar_camera()

        elif opcao == "0":
            print("[INFO] Encerrando o sistema. Ate logo!")
            sys.exit(0)

        else:
            print("[ERRO] Opcao invalida.")


# ---------------------------------------------------------------------------
# Ponto de entrada
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    menu()
