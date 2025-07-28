# ComunicaPJE

# 📄 EMATER-MG - Coletor de Processos (via API do PJe)

Este script Python consulta automaticamente a API pública do PJe para buscar processos com parte **EMATER-MG** e extrai os dados principais (número do processo, tribunal, data, advogados, etc.). Em seguida, gera uma planilha `.xlsx` organizada com os resultados, destacando quais advogados (baseado em uma lista de OABs) estão envolvidos.

---

## ✨ Funcionalidades

- Busca processos da **EMATER-MG** com data de ontem.
- Filtra por advogados com base em OABs específicas.
- Exporta os dados para uma planilha Excel.
- Interface para escolher onde salvar o arquivo (via janela).
- Compatível com advogados com OAB em formatos como `MG156872A`.

---

## ✅ Pré-requisitos

- Python 3.8 ou superior instalado na máquina (apenas para rodar sem `.exe`).
- Acesso à internet com IP **brasileiro** (a API bloqueia IPs estrangeiros).

---

## 🐍 Usando ambiente virtual (`venv`)

1. **Crie o ambiente virtual:**

```bash
python -m venv venv
```

2. **Ative o ambiente:**

- No Windows:
  ```bash
  venv\Scripts\activate
  ```

- No Linux/Mac:
  ```bash
  source venv/bin/activate
  ```

3. **Instale as dependências:**

```bash
pip install -r requirements.txt
```

---

## ▶️ Executando o script manualmente

Após ativar o `venv` e instalar as dependências:

```bash
python main.py
```

---

## 🛠 Gerando o executável `.exe` com auto-py-to-exe

### 1. Instale o auto-py-to-exe

```bash
pip install auto-py-to-exe
```

### 2. Rode o auto-py-to-exe

```bash
auto-py-to-exe
```

### 3. Configurações recomendadas:

- **Script Location**: selecione o `main.py`
- Marque **One File**
- Marque **Window Based** (para rodar sem console preto) ou **Console Based** (se quiser ver logs)
- (Opcional) Adicione um ícone `.ico` em **"Icon"**
- Clique em **Convert .py to .exe**

O executável será gerado na pasta `output`.

---
