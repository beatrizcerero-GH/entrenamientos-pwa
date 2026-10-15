# 🏋️ Registro de Entrenamientos

App web para registrar entrenamientos con pesas, con base de datos en Google Sheets.
Este repositorio aloja la **carcasa PWA**, que permite **instalar la app en el móvil**
(icono propio, nombre y pantalla completa) sobre la aplicación de Google Apps Script.

## 🧩 Cómo funciona

- **Motor + base de datos** → Google Apps Script + Google Sheets (carpeta `apps-script/`).
- **Carcasa instalable** → esta web, publicada con **GitHub Pages**, que abre el motor a pantalla completa.

## 🚀 Puesta en marcha

### 1. Backend (una sola vez)
1. Crea una hoja de **Google Sheets** → menú **Extensiones › Apps Script**.
2. Pega `apps-script/Code.gs` y crea un archivo HTML llamado `Index` con `apps-script/Index.html`.
3. Ejecuta la función `setup` una vez (autoriza permisos).
4. **Implementar › Nueva implementación › Aplicación web** → *Ejecutar como: yo* · *Acceso: cualquiera con cuenta de Google*.
5. Copia la **URL** que termina en `/exec`.

### 2. Carcasa PWA (este repo)
1. Abre `index.html` y pega tu URL `/exec` en la variable **`APP_URL`** (arriba del todo).
2. Sube el repo a GitHub.
3. En GitHub: **Settings › Pages › Branch: `main` / carpeta `/root`** → Guardar.
4. En 1-2 minutos tendrás una URL tipo `https://tuusuario.github.io/entrenamientos-pwa/`.

### 3. Instalar en el móvil
- **Android (Chrome):** abre la URL → menú ⋮ → *Instalar aplicación* / *Añadir a pantalla de inicio*.
- **iPhone (Safari):** abre la URL → *Compartir* → *Añadir a pantalla de inicio*.

## 📁 Estructura

```
index.html              carcasa PWA (aquí va tu APP_URL)
manifest.webmanifest    nombre, icono y modo pantalla completa
sw.js                   service worker (hace la app instalable)
icon-192 / 512 / apple  iconos de la app
apps-script/            código del backend (Code.gs + Index.html)
```
