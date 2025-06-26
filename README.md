# automacao-de-excel

# 🧹 Filtro de Registros Sem Bairro (Excel)

Este projeto contém um script Python que **lê uma planilha Excel** com registros de pessoas e **filtra automaticamente todos os registros onde a coluna "BAIRRO" está vazia, nula ou preenchida apenas com espaços**.

O objetivo é facilitar a limpeza de dados, separando os registros incompletos em um novo arquivo para revisão.

---

## ✅ O que o script faz

1. **Abre o arquivo Excel original** (ex: `japeri.xlsx`).
2. **Remove espaços extras dos nomes das colunas** (ex: "BAIRRO " → "BAIRRO").
3. **Verifica a existência da coluna `BAIRRO`**.
4. **Conta quantos registros têm bairro preenchido e quantos estão vazios**.
5. **Filtra todos os registros sem bairro** (vazios, nulos ou só com espaços).
6. **Exibe uma amostra dos registros sem bairro no terminal**.
7. **Salva esses registros em um novo arquivo Excel** chamado `sem_bairro.xlsx`, na aba `Registros Sem Bairro`.
8. **Formata automaticamente a largura das colunas** no arquivo gerado.
9. **Adiciona informações no topo do arquivo** com:
   - Nome do arquivo original
   - Data e hora da extração
   - Total de registros sem bairro

---

## 📂 Arquivos gerados

- **`sem_bairro.xlsx`** → contém apenas os registros que precisam de revisão (sem informação de bairro).
- Aba criada: `Registros Sem Bairro`

---

## 🛠️ Requisitos

Você precisa ter as seguintes bibliotecas instaladas:

```bash
pip install pandas openpyxl


No terminal, execute:
python sembairro.py


Exemplo de saída no terminal

📊 Total de registros: 4449
❌ Registros sem bairro: 2046
✅ Registros com bairro: 2403
📈 Taxa de completude: 54.0%
```
