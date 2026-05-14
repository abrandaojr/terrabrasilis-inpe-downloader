# TerraBrasilis/INPE Public Data Downloader

Repositório para listar, baixar, retomar downloads interrompidos e validar arquivos ZIP públicos disponíveis na página oficial do TerraBrasilis/INPE.

Fonte dos dados:

https://terrabrasilis.dpi.inpe.br/en/download-files/

Este repositório fornece apenas o código. Ele não redistribui bases do TerraBrasilis/INPE. Os arquivos são baixados diretamente da fonte oficial e armazenados localmente na máquina do usuário.

## Estrutura

```text
terrabrasilis-inpe-downloader/
├── README.md
├── LICENSE
├── requirements.txt
├── .gitignore
├── CITATION.cff
├── scripts/
│   └── download_terrabrasilis.py
└── data/
    ├── raw/
    └── metadata/
```

## Instalação

Crie um ambiente virtual e instale as dependências:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

No Linux ou macOS:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Uso básico

Rodar com confirmação manual antes do download:

```bash
python scripts/download_terrabrasilis.py
```

Rodar sem confirmação interativa:

```bash
python scripts/download_terrabrasilis.py --yes
```

Definir uma pasta local de saída:

```bash
python scripts/download_terrabrasilis.py --root data/raw/terrabrasilis --yes
```

Forçar renderização dinâmica com Selenium, caso a página deixe de expor os links no HTML estático:

```bash
python scripts/download_terrabrasilis.py --dynamic --yes
```

## Saídas locais

Por padrão, os dados são salvos em:

```text
data/raw/terrabrasilis/<YYYY-MM-DD>/
```

Dentro dessa pasta, o script organiza os arquivos por bioma e categoria quando essa informação pode ser inferida da URL.

O script também salva metadados da raspagem:

```text
terrabrasilis_zips.csv
terrabrasilis_zips.json
validation_report.json
```

Esses arquivos locais não devem ser enviados ao GitHub.

## Aviso sobre uso dos dados

Este repositório não redistribui bases do TerraBrasilis/INPE. Os scripts apenas automatizam o download direto da fonte oficial.

Usuários são responsáveis por verificar os termos de uso, requisitos de citação e condições de uso definidos pelo INPE/TerraBrasilis antes de publicar, compartilhar ou interpretar produtos derivados.

Os resultados derivados desses dados devem ser validados antes de uso em relatórios, publicações, painéis, análises técnicas ou processos de decisão.

## Licença

O código deste repositório é disponibilizado sob a licença MIT.

Essa licença se aplica apenas ao código desenvolvido neste repositório. Ela não se aplica às bases de dados do TerraBrasilis/INPE ou a qualquer outro dado baixado por meio dos scripts.
