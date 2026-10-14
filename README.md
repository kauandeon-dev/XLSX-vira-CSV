<div align="center">

# 📊 ➜ 📄 XLSX vira CSV

**Converta pastas inteiras de planilhas Excel em CSV com um único comando.**

[![Python](https://img.shields.io/badge/python-3.8%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![openpyxl](https://img.shields.io/badge/dep-openpyxl-217346?logo=microsoftexcel&logoColor=white)](https://pypi.org/project/openpyxl/)
[![CLI](https://img.shields.io/badge/interface-CLI-black?logo=gnubash&logoColor=white)](#-uso)
[![Plataformas](https://img.shields.io/badge/plataformas-Linux%20%7C%20macOS%20%7C%20Windows-lightgrey)](#)
[![Licença: MIT](https://img.shields.io/badge/licen%C3%A7a-MIT-yellow.svg)](LICENSE)

</div>

---

Chega de abrir planilha por planilha e clicar em *Salvar como → CSV*. O `xlsx2csv.py` varre uma pasta (ou subpastas), converte todos os `.xlsx`/`.xlsm` que encontrar e te entrega um resumo no final. Um arquivo, uma dependência, zero configuração.

```console
$ python xlsx2csv.py ./relatorios -r -o ./csv --all-sheets
[OK]   relatorio.xlsx :: Vendas -> relatorio__Vendas.csv (4 linhas)
[OK]   relatorio.xlsx :: Resumo Q1 (2026) -> relatorio__Resumo_Q1_2026.csv (1 linhas)
[OK]   outro.xlsx :: Sheet -> outro.csv (1 linhas)

Concluído: 3 gerado(s), 0 pulado(s), 0 falha(s).
```

## ✨ Recursos

- 📁 **Conversão em lote** — aponta pra uma pasta e pronto; também aceita um arquivo único
- 🔁 **Modo recursivo** — varre subpastas e **espelha a estrutura** na pasta de saída
- 📑 **Todas as abas** — exporta só a primeira (padrão) ou cada aba num CSV separado
- 🧮 **Valores, não fórmulas** — exporta o resultado calculado das células
- 📅 **Datas padronizadas** — `datetime` vira ISO 8601 (`2026-01-05` ou `2026-01-06 09:15:00`)
- 🔢 **Números limpos** — `10.0` vira `10`, booleanos viram `1`/`0`
- 🧹 **Ignora lixo** — pula linhas totalmente vazias e arquivos temporários do Excel (`~$arquivo.xlsx`)
- 🛡️ **Seguro por padrão** — não sobrescreve CSV existente sem `--overwrite`
- 🪟 **Abre direito no Excel** — encoding padrão `utf-8-sig` (com BOM), acentos aparecem certinhos
- ⚡ **Leve** — abre as planilhas em modo *read-only*, aguenta arquivos grandes sem estourar memória

## 📦 Instalação

```bash
git clone https://github.com/kauandeon-dev/XLSX-vira-CSV.git
cd XLSX-vira-CSV
pip install -r requirements.txt
```

> Só depende do [`openpyxl`](https://pypi.org/project/openpyxl/). Se preferir: `pip install openpyxl`.

## 🚀 Uso

```bash
python xlsx2csv.py [entrada] [opções]
```

Sem argumentos, converte todos os `.xlsx`/`.xlsm` da pasta atual e salva os CSVs ao lado deles.

### Opções

| Opção | Descrição | Padrão |
|---|---|---|
| `entrada` | Pasta ou arquivo de origem | `.` (pasta atual) |
| `-o`, `--saida` | Pasta de destino | mesma pasta da origem |
| `-r`, `--recursivo` | Varre subpastas | desligado |
| `-d`, `--delim` | Delimitador de campos | `,` |
| `-e`, `--encoding` | Encoding do CSV gerado | `utf-8-sig` |
| `--all-sheets` | Exporta todas as abas, não só a primeira | desligado |
| `--quote-all` | Coloca aspas em todos os campos | desligado |
| `--overwrite` | Sobrescreve CSVs já existentes | desligado |

### Exemplos

**Converter um único arquivo**
```bash
python xlsx2csv.py planilha.xlsx
```

**Converter uma pasta inteira, jogando os CSVs em outro lugar**
```bash
python xlsx2csv.py ~/Downloads/planilhas -o ~/Downloads/csv
```

**Varrer subpastas mantendo a estrutura**
```bash
python xlsx2csv.py ./dados -r -o ./saida
```
```
dados/                      saida/
├── jan.xlsx         ➜      ├── jan.csv
└── 2025/                   └── 2025/
    └── dez.xlsx                └── dez.csv
```

**Exportar todas as abas**
```bash
python xlsx2csv.py vendas.xlsx --all-sheets
# vendas__Janeiro.csv, vendas__Fevereiro.csv, vendas__Resumo.csv ...
```

**CSV no padrão Excel brasileiro (ponto e vírgula)**
```bash
python xlsx2csv.py ./planilhas -d ";"
```

**TSV (separado por tab)**
```bash
python xlsx2csv.py ./planilhas -d $'\t'
```

**Encoding legado pra sistemas antigos**
```bash
python xlsx2csv.py ./planilhas -e latin-1
```

**Reprocessar tudo do zero**
```bash
python xlsx2csv.py ./planilhas -r -o ./csv --all-sheets --overwrite
```

## 🏷️ Nomes dos arquivos gerados

| Situação | Nome do CSV |
|---|---|
| Padrão (só a primeira aba) | `arquivo.csv` |
| `--all-sheets` e o arquivo tem **várias** abas | `arquivo__NomeDaAba.csv` |
| `--all-sheets` e o arquivo tem **uma** aba | `arquivo.csv` |

Caracteres especiais no nome da aba viram `_` — `Resumo Q1 (2026)` vira `Resumo_Q1_2026`.

## 🔄 Como os valores são convertidos

| Tipo na célula | Exemplo no Excel | No CSV |
|---|---|---|
| Data | `05/01/2026` | `2026-01-05` |
| Data e hora | `06/01/2026 09:15` | `2026-01-06 09:15:00` |
| Hora | `14:30` | `14:30:00` |
| Número inteiro | `10,0` | `10` |
| Número decimal | `12,5` | `12.5` |
| Booleano | `VERDADEIRO` / `FALSO` | `1` / `0` |
| Fórmula | `=SOMA(A1:A3)` | resultado calculado |
| Vazio | | *(campo vazio)* |

## ⚠️ Observações

- **Fórmulas:** o script lê o último valor calculado que o Excel salvou no arquivo. Planilhas geradas por outros programas que nunca foram abertas e salvas no Excel/LibreOffice podem vir com as células de fórmula vazias.
- **Formatos suportados:** `.xlsx` e `.xlsm`. Arquivos `.xls` (Excel 97-2003) não são suportados — abra e salve como `.xlsx` antes.
- **Caracteres fora do encoding:** se você escolher um encoding como `latin-1` e a planilha tiver caracteres que ele não representa, eles são trocados por `?` em vez de quebrar a conversão.
- **Erros não param o lote:** se um arquivo estiver corrompido ou protegido, ele entra como `[ERRO]` no log e o script segue para o próximo.

## 📜 Licença

Distribuído sob a licença MIT. Veja [`LICENSE`](LICENSE) para mais detalhes.

---

<div align="center">

Feito por [**kauandeon-dev**](https://github.com/kauandeon-dev)

</div>
