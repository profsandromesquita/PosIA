# Reconhecimento Facial com OpenCV

Sistema de cadastro e reconhecimento facial usando Python, OpenCV e o algoritmo LBPH (Local Binary Patterns Histograms).

## Requisitos

- Python 3.8+
- Webcam (para cadastro via webcam e reconhecimento em tempo real)

## Instalacao

```bash
cd ReconhecimentoFacial
pip install -r requirements.txt
```

## Como Usar

Execute o script principal:

```bash
python reconhecimento_facial.py
```

O menu interativo oferece 5 opcoes:

### 1. Cadastrar rosto (arquivo de imagem)

Cadastra uma pessoa a partir de um arquivo de imagem existente (JPG, PNG, etc.).

- Informe o **nome da pessoa**
- Informe o **caminho da imagem**
- O sistema detecta o rosto automaticamente, recorta e salva na pasta `dataset/<nome>/`

### 2. Cadastrar rosto (webcam)

Captura multiplas fotos do rosto pela webcam.

- Informe o **nome da pessoa**
- Informe a **quantidade de capturas** (padrao: 10)
- Posicione o rosto em frente a webcam
- As capturas sao feitas automaticamente quando um rosto e detectado
- Pressione **q** para cancelar

> Dica: quanto mais fotos com angulos e iluminacoes diferentes, melhor o reconhecimento.

### 3. Treinar modelo

Treina o modelo LBPH com todas as imagens cadastradas.

- Deve ser executado **apos cadastrar** pelo menos uma pessoa
- Deve ser **re-executado** sempre que novos rostos forem cadastrados

### 4. Reconhecimento em tempo real

Abre a webcam e identifica rostos em tempo real.

- Rostos reconhecidos aparecem com **retangulo verde** e o nome da pessoa
- Rostos desconhecidos aparecem com **retangulo vermelho**
- A porcentagem indica o nivel de confianca
- Pressione **q** para sair

### 5. Listar pessoas cadastradas

Exibe todas as pessoas cadastradas com a quantidade de fotos.

## Estrutura de Pastas

```
ReconhecimentoFacial/
├── reconhecimento_facial.py    # Script principal
├── requirements.txt            # Dependencias
├── DocReconhecimentoFacial.md  # Esta documentacao
├── dataset/                    # Fotos cadastradas
│   ├── Maria/
│   │   ├── Maria_1.jpg
│   │   └── Maria_2.jpg
│   └── Joao/
│       ├── Joao_1.jpg
│       └── Joao_2.jpg
└── modelo/                     # Modelo treinado
    ├── modelo_lbph.yml
    └── labels.npy
```

## Fluxo de Uso Tipico

1. Cadastre pelo menos 2 pessoas (opcao 1 ou 2)
2. Treine o modelo (opcao 3)
3. Inicie o reconhecimento em tempo real (opcao 4)

## Algoritmo Utilizado

**LBPH (Local Binary Patterns Histograms):**

- Funciona bem com poucas imagens de treino
- Robusto a variacoes de iluminacao
- Nao necessita de GPU
- Adequado para aplicacoes em tempo real

## Tecnologias

| Componente | Tecnologia |
|---|---|
| Linguagem | Python 3.8+ |
| Visao Computacional | OpenCV |
| Deteccao de Rosto | Haar Cascade Classifier |
| Reconhecimento | LBPH Face Recognizer |
| Dados | NumPy |
