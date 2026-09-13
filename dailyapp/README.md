# Mis Finanzas

App de finanzas personales, con inicio de sesión y sincronización real entre
dispositivos vía Supabase.

## Qué cambió respecto a la versión de Claude

- El almacenamiento ahora usa Supabase (base de datos propia) en vez de
  `window.storage` — funciona igual dentro y fuera de Claude, y sincroniza
  de verdad entre PC y celular.
- Se agregó una pantalla de inicio de sesión (email + contraseña) para que
  tus datos sean solo tuyos.
- El valor del dólar/UF ahora se obtiene directo de mindicador.cl (API
  pública oficial de Chile), sin necesitar IA.
- Se quitaron las 4 funciones que usaban la API de Claude (sugerir tipo de
  cuenta, mejorar detección de columnas al importar, clasificar cuentas
  nuevas al importar, "Pregúntale a Claude") — todo lo demás sigue igual.

## Cómo publicar esta app (una sola vez)

### 1. Crear un repositorio en GitHub

1. Ve a **https://github.com** → crea una cuenta si no tienes.
2. Click en **"New repository"** → nómbralo (ej. `mis-finanzas`) → **Create repository**.
3. En tu computador, sube estos archivos a ese repositorio. La forma más
   simple si no usas git antes: en la página del repo recién creado, busca
   el link **"uploading an existing file"** y arrastra ahí TODOS los
   archivos de esta carpeta (manteniendo la subcarpeta `src/`).

### 2. Conectarlo a Vercel (gratis)

1. Ve a **https://vercel.com** → **Sign up** → elige "Continue with GitHub"
   (así quedan conectados automáticamente).
2. Click en **"Add New..." → "Project"**.
3. Busca tu repositorio `mis-finanzas` y dale **Import**.
4. Vercel detecta automáticamente que es un proyecto Vite — deja todo tal
   cual y presiona **Deploy**.
5. Espera 1-2 minutos. Al terminar, te da una URL (algo como
   `mis-finanzas.vercel.app`) — esa es tu app, ya funcionando y accesible
   desde cualquier dispositivo.

### 3. Usarla en el celular como una app

1. Abre esa URL en el navegador de tu celular.
2. En Chrome (Android): menú (⋮) → **"Agregar a pantalla de inicio"**.
   En Safari (iPhone): botón compartir (□↑) → **"Agregar a pantalla de inicio"**.
3. Te queda un ícono como cualquier app — al abrirlo entra directo a la web app.

### 4. Actualizaciones futuras

Cuando sigamos cambiando la app en el chat con Claude, te voy a entregar el
archivo `App.jsx` actualizado. Para publicar la nueva versión: reemplaza el
archivo `src/App.jsx` en tu repositorio de GitHub (subiendo el nuevo encima
del viejo) — Vercel vuelve a publicar solo, automáticamente, en 1-2 minutos.

## Primer uso

Al entrar por primera vez, verás una pantalla para "Crear cuenta" — usa tu
email y una contraseña. Supabase te va a mandar un correo de confirmación;
ábrelo y luego vuelve a iniciar sesión. Desde ahí, tus datos quedan
guardados en tu base de datos y disponibles en cualquier dispositivo donde
inicies sesión con la misma cuenta.
