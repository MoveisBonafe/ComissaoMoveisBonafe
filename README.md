# Móveis Bonafé - Sistema de Comissão

Sistema automatizado para processamento de planilhas Excel e geração de documentos Word de comissão.

## Deploy em Plataformas

### Railway (Recomendado)
1. Acesse [railway.app](https://railway.app)
2. Login com GitHub
3. "New Project" → "Deploy from GitHub repo"
4. Selecione este repositório
5. Deploy automático com `railway.json`

### Render (Alternativa)
1. Acesse [render.com](https://render.com)
2. Login com GitHub  
3. "New Web Service"
4. Selecione este repositório
5. Configure:
   - **Build Command:** `pip install -r requirements-render.txt`
   - **Start Command:** `gunicorn --bind 0.0.0.0:$PORT main:app`

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