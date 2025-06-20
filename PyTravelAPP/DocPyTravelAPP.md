# PyTravelAPP

![PyTravelAPP Logo](logo.png)

**PyTravelAPP** é uma aplicação desktop interativa desenvolvida em Python com PyQt5, voltada para análise de catálogos de pacotes de viagem oferecidos por agências e operadoras de turismo. O sistema permite extrair, visualizar, analisar e exportar dados de pacotes a partir de arquivos `.docx` ou `.pdf` de forma simples, amigável e eficiente.

---

## Índice

- [Visão Geral](#visão-geral)
- [Funcionalidades](#funcionalidades)
- [Arquitetura e Tecnologias](#arquitetura-e-tecnologias)
- [Instalação](#instalação)
- [Como Usar](#como-usar)
- [Descrição dos Componentes](#descrição-dos-componentes)
- [Fluxo da Aplicação](#fluxo-da-aplicação)
- [Processamento do Catálogo](#processamento-do-catálogo)
- [Exportação de Dados](#exportação-de-dados)
- [Estrutura de Diretórios](#estrutura-de-diretórios)
- [Personalização](#personalização)
- [Requisitos e Dependências](#requisitos-e-dependências)
- [Licença](#licença)
- [Contato](#contato)

---

## Visão Geral

O PyTravelAPP foi projetado para facilitar a análise de pacotes de viagem comercializados, identificando categorias, cruzamentos entre elas e exportando as informações de forma estruturada, permitindo a tomada de decisões estratégicas por profissionais de turismo.

---

## Funcionalidades

- **Interface gráfica moderna e intuitiva**: Desenvolvida com PyQt5.
- **Importação de arquivos DOCX e PDF**: Suporte a documentos de texto e catálogos em PDF.
- **Visualização de PDFs como imagem**: Exibição das páginas em miniatura.
- **Extração automática do texto**: Utilizando `docx2txt` e `PyMuPDF`.
- **Análise e categorização dos pacotes**: Classifica pacotes em praias, capitais, interior, avião, ônibus e navio.
- **Cruzamento entre categorias**: Identifica pacotes que pertencem a múltiplas categorias (ex: "capitais_praianas").
- **Exportação para CSV**: Salva os dados analisados em CSV pronto para Excel.
- **Feedback visual ao usuário**: Mensagens de sucesso, confirmação e controle de erros.

---

## Arquitetura e Tecnologias

- **Linguagem:** Python 3
- **Interface gráfica:** [PyQt5](https://riverbankcomputing.com/software/pyqt/intro)
- **Leitura DOCX:** [docx2txt](https://github.com/ankushshah89/python-docx2txt)
- **Leitura PDF:** [PyMuPDF (fitz)](https://pymupdf.readthedocs.io/)
- **Manipulação de dados:** [pandas](https://pandas.pydata.org/)
- **Exportação CSV:** pandas

---

## Instalação

1. **Clone o repositório**
   ```bash
   git clone https://github.com/profsandromesquita/PosIA.git
   cd PosIA/PyTravelAPP
   ```

2. **Instale as dependências**
   ```bash
   pip install -r requirements.txt
   ```
   > Se o arquivo `requirements.txt` não existir, instale manualmente:
   ```bash
   pip install PyQt5 pandas docx2txt pymupdf
   ```

3. **Adicione o arquivo de logo (opcional)**
   - Certifique-se que `logo.png` está na mesma pasta do script principal.

---

## Como Usar

1. Execute o app:
   ```bash
   python PyTravelAPP-1.0.py
   ```
2. Na interface:
   - Clique em **Selecionar Arquivo** e escolha um arquivo `.docx` ou `.pdf` de catálogo de pacotes.
   - Visualize o conteúdo extraído (e, no caso de PDFs, as miniaturas das páginas).
   - Clique em **Confirmar e Analisar** para processar o catálogo.
   - Escolha o local e nome do arquivo para exportar os dados em `.csv`.
   - Ao final, decida se deseja analisar outro arquivo ou encerrar.

---

## Descrição dos Componentes

### 1. PyTravelAnalyzer (Classe Principal)

- **Herda:** `QWidget`
- **Responsável por:**
  - Montar e gerenciar a interface.
  - Controle dos botões, área de texto e visualização de PDF.
  - Orquestrar a leitura, análise, exportação e reinicialização do app.

### 2. Funções Principais
### [__init__()](https://github.com/profsandromesquita/PosIA/blob/35248c65ba8724a1e9026b75d59ad91b106820f3/PyTravelAPP/PyTravelAPP-1.0.py#L14)
- Monta toda a interface gráfica.  
- **Linhas**: 14-22

### [init_ui()](https://github.com/profsandromesquita/PosIA/blob/35248c65ba8724a1e9026b75d59ad91b106820f3/PyTravelAPP/PyTravelAPP-1.0.py#L24)
- Esta função monta todos os elementos da interface gráfica do aplicativo, organizando o layout principal e todos os componentes visuais como botões, áreas de texto, visualização de PDF e cabeçalhos.
- **Linhas**: 24-84

### [button_style()](https://github.com/profsandromesquita/PosIA/blob/35248c65ba8724a1e9026b75d59ad91b106820f3/PyTravelAPP/PyTravelAPP-1.0.py#L86)
- Estilização de botões. 
- **Linhas**: 86-91

### [load_file()](https://github.com/profsandromesquita/PosIA/blob/35248c65ba8724a1e9026b75d59ad91b106820f3/PyTravelAPP/PyTravelAPP-1.0.py#L93)
- Abre diálogo para seleção do arquivo e inicia leitura e exibição do conteúdo.
-  Esta função abre o seletor de arquivos, permitindo que o usuário escolha um rquivo do tipo .docx ou .pdf. Após a seleção, ela processa e exibe o conteúdo do documento, habilitando o botão de análise.
- **Linhas**: 93-108

### [show_pdf_as_images()](https://github.com/profsandromesquita/PosIA/blob/35248c65ba8724a1e9026b75d59ad91b106820f3/PyTravelAPP/PyTravelAPP-1.0.py#L110)
- Renderiza páginas do PDF como imagens.
- Esta função renderiza as páginas de um documento PDF como imagens e as adiciona ao layout de visualização da interface gráfica.
- **Linhas**: 110-123

### [read_file()](https://github.com/profsandromesquita/PosIA/blob/35248c65ba8724a1e9026b75d59ad91b106820f3/PyTravelAPP/PyTravelAPP-1.0.py#L125)
- Lê e extrai texto do arquivo selecionado.
- **Linhas**: 125-131

### [analyze_file()](https://github.com/profsandromesquita/PosIA/blob/35248c65ba8724a1e9026b75d59ad91b106820f3/PyTravelAPP/PyTravelAPP-1.0.py#L133)
- Confirmação de análise, processa dados e exporta para CSV.
- Esta função lida com a análise do texto extraído do documento carregado, exportando os dados processados para um arquivo CSV. Além disso, exibe caixas de diálogo para confirmar a análise e oferecer ao usuário a opção de realizar uma nova análise ou encerrar a aplicação.
- **Linhas**: 133-187

### [processar_catalogo()](https://github.com/profsandromesquita/PosIA/blob/35248c65ba8724a1e9026b75d59ad91b106820f3/PyTravelAPP/PyTravelAPP-1.0.py#L189)
- Analisa e categoriza os pacotes por seções e tipos de transporte.
- **Linhas**: 189-219

### [reset_app()](https://github.com/profsandromesquita/PosIA/blob/35248c65ba8724a1e9026b75d59ad91b106820f3/PyTravelAPP/PyTravelAPP-1.0.py#L221)
- Limpa o estado do app para nova análise.
-  Esta função reinicia os campos e o estado da interface gráfica do aplicativo, permitindo ao usuário iniciar uma nova análise a partir do zero.
- **Linhas**: 221-231

### [__name__ == '__main__'](https://github.com/profsandromesquita/PosIA/blob/35248c65ba8724a1e9026b75d59ad91b106820f3/PyTravelAPP/PyTravelAPP-1.0.py#L233)
-   
- **Linhas**: 233-237


---

## Fluxo da Aplicação

1. **Início:** Tela principal com logo, título e botão de seleção.
2. **Seleção de arquivo:** Usuário seleciona um `.docx` ou `.pdf`.
3. **Visualização:** Exibe o nome do arquivo, conteúdo extraído e, se PDF, miniaturas das páginas.
4. **Análise:** Usuário confirma análise, dados são processados e categorizados.
5. **Exportação:** Usuário escolhe local para salvar o CSV.
6. **Reinício/Ou Encerrar:** Usuário pode reiniciar para novo arquivo ou fechar o app.

---

## Processamento do Catálogo

O processamento consiste em:
- **Identificação de seções:** Procura por palavras-chave como "praianas", "capitais", "interior", "avião", "ônibus", "navio".
- **Coleta dos pacotes:** Linhas (geralmente com parênteses) após cada seção são consideradas como pacotes daquela categoria.
- **Cruzamentos automáticos:** Gera listas como "capitais_praianas", "praias_onibus", "interior_aviao", "capitais_navio", identificando pacotes que aparecem em múltiplas categorias.

O resultado é um dicionário em que cada chave representa uma categoria ou cruzamento e os valores são listas de pacotes.

---

## Exportação de Dados

Os dados processados são exportados em formato `.csv` no padrão:

| PRAIAS | CAPITAIS | INTERIOR | AVIAO | ONIBUS | NAVIO | CAPITAIS_PRAIANAS | PRAIAS_ONIBUS | INTERIOR_AVIAO | CAPITAIS_NAVIO |
|--------|----------|----------|-------|--------|-------|-------------------|---------------|----------------|----------------|
| ...    | ...      | ...      | ...   | ...    | ...   | ...               | ...           | ...            | ...            |

Cada coluna pode ter listas de pacotes (linhas podem ficar em branco se não houver dados).

---

## Estrutura de Diretórios

```
PyTravelAPP/
├── PyTravelAPP-1.0.py
├── logo.png
├── requirements.txt
├── README.md
└── (outros arquivos gerados)
```

---

## Personalização

- **Logo:** Edite ou substitua o arquivo `logo.png` para personalizar a identidade visual.
- **Palavras-chave:** Para suportar novas categorias, edite a função `processar_catalogo`.
- **Estilo:** Modifique estilos diretamente nos métodos de estilização PyQt5.

---

## Requisitos e Dependências

- Python 3.6+
- PyQt5
- pandas
- docx2txt
- PyMuPDF (fitz)

---

## Licença

Este projeto é distribuído sob a licença MIT. Consulte o arquivo [`LICENSE`](../LICENSE) para mais detalhes.

---

## Contato

- Autor: [@profsandromesquita](https://github.com/profsandromesquita)
- Dúvidas, sugestões ou contribuições são bem-vindas via [Issues](https://github.com/profsandromesquita/PosIA/issues).

---
