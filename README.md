# Cummins Perú · Dashboard Comercial Zona Sur

Dashboard de inteligencia comercial para Cummins Perú - Zona Sur (K27 Arequipa, K11 Huancayo, K23 Tacna).

## 🚀 Deploy automático

Cada vez que haces `git push` a `main`, GitHub Actions compila y despliega automáticamente.

**URL del dashboard:** `https://TU-USUARIO.github.io/cummins-dashboard/`

## 💻 Desarrollo local

```bash
npm install
npm run dev
```

Abre http://localhost:5173

## 📦 Build manual

```bash
npm run build
```

Los archivos compilados quedan en `/dist`.

## 📊 Uso

1. Abre el dashboard en el navegador
2. Arrastra tus archivos Excel (ZPAVX005) o usa el botón **+ EXCELS**
3. Nombra los archivos como `2026_04_K27.xlsx` para detección automática de período y sucursal
4. Filtra, explora y analiza

## 🔥 Firebase (opcional)

Para habilitar subida a la nube y acceso compartido, edita `src/App.jsx` y reemplaza los valores en `FB_CONFIG` con los de tu proyecto Firebase.

## 📁 Estructura

```
cummins-dashboard/
├── src/
│   ├── App.jsx       ← Dashboard completo
│   └── main.jsx      ← Entrada React
├── index.html
├── vite.config.js
├── package.json
└── .github/
    └── workflows/
        └── deploy.yml  ← Auto-deploy a GitHub Pages
```
