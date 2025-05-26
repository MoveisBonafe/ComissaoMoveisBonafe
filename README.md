# Móveis Bonafé - Sistema de Comissão

Sistema automatizado para processamento de planilhas Excel e geração de documentos Word de comissão.

## Deploy no Railway

### Passo 1: Preparar o Repositório
1. Faça commit de todos os arquivos no seu repositório Git
2. Certifique-se que os arquivos `Procfile` e `railway.json` estão incluídos

### Passo 2: Deploy no Railway
1. Acesse [railway.app](https://railway.app)
2. Clique em "Start a New Project"
3. Conecte sua conta GitHub
4. Selecione "Deploy from GitHub repo"
5. Escolha este repositório
6. Railway detectará automaticamente que é uma aplicação Python/Flask

### Passo 3: Configuração
- Railway irá automaticamente instalar as dependências
- A aplicação estará disponível em uma URL `.railway.app`
- Sem necessidade de configurações adicionais

### Recursos do Sistema
- ✅ Processamento em lote de múltiplos arquivos Excel
- ✅ Cálculos automáticos de comissão
- ✅ Geração automática de documentos Word
- ✅ Download instantâneo dos resultados
- ✅ Interface moderna com drag & drop

### Tecnologias Utilizadas
- Flask (Python)
- Bootstrap 5
- Feather Icons
- openpyxl (processamento Excel)
- python-docx (geração Word)

### Estrutura do Projeto
```
├── main.py              # Aplicação principal
├── app.py               # Configuração Flask
├── utils/               # Utilitários
│   ├── excel_processor.py
│   ├── word_processor.py
│   └── calculations.py
├── templates/           # Templates HTML
├── static/              # CSS e JavaScript
├── templates_word/      # Template Word fixo
└── uploads/             # Arquivos processados
```