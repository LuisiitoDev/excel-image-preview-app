# Excel & Image Preview App

Una web app estática que permite **previsualizar archivos `.xlsx` e imágenes** directamente en el navegador, sin servidor ni backend.

## ✨ Características

- 📊 **Excel (.xlsx)** — parsea el archivo en el cliente con [SheetJS](https://sheetjs.com/), renderiza una tabla con selector de hojas
- 🖼️ **Imágenes** — muestra `image/*` (PNG, JPG, JPEG, WEBP, GIF) en vista previa grande
- 📋 **Metadatos** — nombre, tamaño y tipo MIME del archivo
- ❌ **Manejo de errores** — formatos no soportados y archivos corruptos informados al usuario
- 🎨 **Interfaz limpia** — tema oscuro, responsiva, lista para GitHub Pages

## 🚀 Uso local

```bash
git clone https://github.com/LuisiitoDev/excel-image-preview-app.git
cd excel-image-preview-app
# Abre index.html en tu navegador
open index.html    # macOS
xdg-open index.html  # Linux
start index.html   # Windows
```

## 🌐 GitHub Pages

1. Ve a **Settings → Pages** del repositorio
2. En *Build and deployment* → Source: **Deploy from a branch**
3. Branch: `main` / carpeta: `/ (root)`
4. Guarda y espera ~1 minuto

Tu app estará disponible en:  
`https://luisiitodev.github.io/excel-image-preview-app/`

## 📁 Estructura

```
excel-image-preview-app/
├── index.html   # Marcado HTML
├── styles.css   # Estilos
└── app.js       # Lógica de la aplicación
```

## 🛠️ Tecnologías

- HTML5, CSS3, JavaScript (ES2020+)
- [SheetJS (xlsx)](https://cdn.sheetjs.com/) — CDN, sin npm
