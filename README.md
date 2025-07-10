# Modelling Tools Flask App

A scalable Flask application that provides web interfaces for PELMOex and TOXSWAex tools using Flask blueprints.

## 🏗️ Architecture

This application uses **Flask blueprints** for modular, scalable development:

```
ModellingToolsFlask/
├── app.py                    # Main entry point
├── templates/
│   └── index.html           # Landing page with tool selection
├── pelmoex/                 # PELMOex Blueprint
│   ├── __init__.py
│   ├── routes.py            # PELMOex routes and API endpoints
│   ├── extractor.py         # PELMO data extraction logic
│   └── templates/
│       └── pelmoex/
│           └── index.html   # PELMOex web interface
├── toxswaex/                # TOXSWAex Blueprint
│   ├── __init__.py
│   ├── routes.py            # TOXSWAex routes and API endpoints
│   ├── extractor.py         # TOXSWA data extraction logic
│   └── templates/
│       └── toxswaex/
│           └── index.html   # TOXSWAex web interface
└── static/                  # Shared static assets
```

## 🚀 Getting Started

### Prerequisites
- Python 3.7+
- Flask
- xlsxwriter

### Installation
1. Clone the repository
2. Install dependencies:
   ```bash
   pip install flask xlsxwriter
   ```

### Running the Application
```bash
python app.py
```

The application will be available at `http://localhost:5000`

## 📱 Usage

### Landing Page (`/`)
- Modern interface with cards for each tool
- Click on a tool to access its specific interface

### PELMOex Tool (`/pelmoex/`)
- Extract data from PELMO FOCUS projects
- Supports multiple project selection
- Excel export functionality
- Parametric limit highlighting

### TOXSWAex Tool (`/toxswaex/`)
- Extract data from TOXSWA SWASH projects
- Parent and metabolite compound support
- RAC value exceedance highlighting
- Areic deposition comparison
- Excel export with detailed sheets

## 🔧 Development

### Adding New Tools

To add a new tool (e.g., "NEWTOOLex"):

1. **Create the blueprint structure:**
   ```bash
   mkdir newtoolex
   mkdir newtoolex/templates/newtoolex
   mkdir newtoolex/static
   ```

2. **Create the files:**
   - `newtoolex/__init__.py` - Package initialization
   - `newtoolex/routes.py` - Routes and API endpoints
   - `newtoolex/extractor.py` - Data extraction logic
   - `newtoolex/templates/newtoolex/index.html` - Web interface

3. **Register the blueprint in `app.py`:**
   ```python
   from newtoolex.routes import newtoolex_bp
   app.register_blueprint(newtoolex_bp, url_prefix='/newtoolex')
   ```

4. **Add to the landing page in `templates/index.html`**

### Blueprint Structure

Each tool blueprint follows this pattern:

```python
# routes.py
from flask import Blueprint, render_template, request, jsonify
from .extractor import ToolExtractor

tool_bp = Blueprint('tool', __name__, 
                   template_folder='templates',
                   static_folder='static')

@tool_bp.route('/')
def tool_index():
    return render_template('tool/index.html')

# API endpoints...
```

## 🎯 Benefits of This Architecture

### ✅ **Scalability**
- Each tool is independent and can be developed separately
- Easy to add new tools without affecting existing ones
- Blueprints can be moved to separate microservices later

### ✅ **Maintainability**
- Clear separation of concerns
- Modular code structure
- Easy to test individual components

### ✅ **User Experience**
- Single entry point for all tools
- Consistent navigation and theming
- Shared static assets and layouts

### ✅ **Future-Proof**
- Ready for authentication and user management
- Can easily add shared features (logging, analytics, etc.)
- Prepared for deployment scaling

## 🔄 Migration from Original Apps

The original standalone apps (`PELMOex/app.py` and `TOXSWAex/app.py`) have been refactored into:

1. **Extraction Logic** → `pelmoex/extractor.py` and `toxswaex/extractor.py`
2. **Routes and API** → `pelmoex/routes.py` and `toxswaex/routes.py`
3. **Templates** → `pelmoex/templates/pelmoex/index.html` and `toxswaex/templates/toxswaex/index.html`

## 🚀 Deployment

### Development
```bash
python app.py
```

### Production
```bash
# Using gunicorn
gunicorn -w 4 -b 0.0.0.0:5000 app:app

# Using waitress (Windows)
waitress-serve --host=0.0.0.0 --port=5000 app:app
```

## 📝 API Endpoints

### PELMOex
- `POST /pelmoex/scan_directory` - Scan for PELMO projects
- `POST /pelmoex/extract_data` - Extract PELMO data
- `POST /pelmoex/export_excel` - Export to Excel
- `GET /pelmoex/get_table_data` - Get current table data

### TOXSWAex
- `POST /toxswaex/scan_directory` - Scan for TOXSWA projects
- `POST /toxswaex/extract_data` - Extract TOXSWA data
- `POST /toxswaex/export_excel` - Export to Excel
- `GET /toxswaex/get_table_data` - Get current table data

## 🤝 Contributing

1. Create a new branch for your feature
2. Follow the blueprint structure for new tools
3. Ensure all templates use consistent styling
4. Test thoroughly before submitting

## 📄 License

This project is licensed under the MIT License. 