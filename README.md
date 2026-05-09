# Sistema de Control de Asistencia
### Con foto, geolocalización, base de datos y reporte diario por correo

---

## ¿Qué incluye este sistema?

- ✅ Login separado para administrador y empleados
- 📷 Foto tomada en el momento de marcar (anti-clonación)
- 📍 Geolocalización GPS al marcar entrada/salida
- 🗄️ Base de datos SQLite persistente
- 📊 Panel admin con stats, historial y gestión de empleados
- 📧 Reporte Excel diario automático por Gmail a las 19:00
- ⬇️ Descarga de Excel con fotos miniatura incluidas
- 🔍 Log de todas las acciones del sistema

---

## INSTALACIÓN LOCAL (para probar en tu PC)

### Requisitos
- Node.js 18 o superior → https://nodejs.org

### Pasos

```bash
# 1. Entra a la carpeta del proyecto
cd control-asistencia

# 2. Instala las dependencias
npm install --ignore-scripts

# 3. Crea tu archivo de configuración
cp .env.example .env
# En Windows: copy .env.example .env

# 4. Edita el archivo .env con tu editor (Notepad, VS Code, etc.)
# Como mínimo debes configurar GMAIL_USER y GMAIL_APP_PASSWORD

# 5. Inicia el servidor
npm start

# 6. Abre tu navegador en:
# http://localhost:3000
```

> ⚠️ **Por qué `--ignore-scripts`:** La base de datos usa `sql.js`, que es JavaScript puro y no necesita compilación. Sin ese flag, en algunos sistemas Windows o Mac puede dar un error de compilación nativa. Si aun así falla, prueba:
> ```bash
> npm install --legacy-peer-deps --ignore-scripts
> ```

**Credenciales de inicio (demo):**
- Admin: `admin@empresa.com` / `admin123`
- Empleado: `ana@empresa.com` / `empleado123`

> ⚠️ Cambia las contraseñas inmediatamente después de instalar.

---

## CONFIGURAR GMAIL PARA EL REPORTE

Para que el sistema envíe correos desde tu Gmail necesitas una "Contraseña de aplicación" (no tu contraseña normal):

1. Ve a https://myaccount.google.com
2. Clic en **Seguridad** (menú izquierdo)
3. Asegúrate de tener **Verificación en 2 pasos** activada
4. Busca **Contraseñas de aplicaciones** y haz clic
5. Selecciona "Otra (nombre personalizado)" → escribe "Asistencia"
6. Copia la contraseña de 16 caracteres que genera Google
7. Pégala en tu `.env` como `GMAIL_APP_PASSWORD=xxxx-xxxx-xxxx-xxxx`

---

## DESPLIEGUE EN RAILWAY (gratis, recomendado)

Railway es la opción más fácil. Puedes tener el sistema online en ~5 minutos.

### Paso 1: Crear cuenta
- Ve a https://railway.app
- Regístrate con tu cuenta de GitHub (es gratis)

### Paso 2: Subir el código a GitHub
```bash
# En la carpeta del proyecto:
git init
git add .
git commit -m "Sistema de asistencia"
```
- Ve a https://github.com/new
- Crea un repositorio nuevo (puede ser privado)
- Sigue las instrucciones para subir el código

### Paso 3: Desplegar en Railway
1. En Railway, clic en **New Project**
2. Selecciona **Deploy from GitHub repo**
3. Elige tu repositorio
4. Railway detecta Node.js automáticamente y despliega

### Paso 4: Configurar variables de entorno
En Railway, ve a tu proyecto → pestaña **Variables** → agrega:

| Variable | Valor |
|---|---|
| `JWT_SECRET` | Una cadena larga aleatoria (ej: `mi-sistema-2024-secreto-xyz`) |
| `GMAIL_USER` | `tucorreo@gmail.com` |
| `GMAIL_APP_PASSWORD` | La contraseña de app de Google |
| `ADMIN_EMAIL` | Correo donde llega el reporte |

### Paso 5: Persistencia de datos (importante)
SQLite guarda los datos en un archivo. Para que no se pierdan en Railway:

1. En tu proyecto Railway → **Add Plugin** → **Volume**
2. Monta el volumen en `/app`
3. Agrega la variable: `DB_PATH=/app/asistencia.db`

### Paso 6: Acceder al sistema
Railway te da una URL pública automática (ej: `https://tu-proyecto.up.railway.app`).
Comparte esa URL con tus empleados.

---

## DESPLIEGUE EN RENDER (alternativa gratuita)

### Paso 1: Cuenta
- Ve a https://render.com y regístrate con GitHub

### Paso 2: Nuevo servicio
1. **New** → **Web Service**
2. Conecta tu repositorio de GitHub
3. Configura:
   - **Build Command:** `npm install`
   - **Start Command:** `npm start`
   - **Plan:** Free

### Paso 3: Variables de entorno
En Render, ve a **Environment** y agrega las mismas variables que en Railway.

### Nota sobre Render gratuito
El plan gratuito de Render "duerme" después de 15 minutos sin actividad.
El primer acceso del día puede tardar 30-60 segundos en cargar.
Para uso continuo, el plan Starter (~$7/mes) lo mantiene siempre activo.

---

## ESTRUCTURA DEL PROYECTO

```
control-asistencia/
├── server.js          → Servidor principal y rutas API
├── database.js        → Base de datos SQLite
├── package.json       → Dependencias
├── .env.example       → Plantilla de configuración
├── .env               → Tu configuración (NO subir a GitHub)
├── asistencia.db      → Base de datos (se crea automáticamente)
├── uploads/           → Fotos de empleados y registros
└── public/
    └── index.html     → Aplicación web completa
```

---

## AGREGAR A .gitignore

Crea un archivo `.gitignore` en la raíz con este contenido para no subir datos privados:

```
.env
asistencia.db
uploads/
node_modules/
```

---

## USO DEL SISTEMA

### Como Administrador
1. Ingresa con email/contraseña de admin
2. **Hoy**: Ve en tiempo real quién llegó, a qué hora y su foto
3. **Empleados**: Agrega o edita empleados con foto de perfil
4. **Historial**: Filtra por mes y empleado
5. **Reportes**: Descarga Excel con fotos o envía por correo ahora
6. El reporte se envía automáticamente a las **19:00** de lunes a sábado

### Como Empleado
1. Ingresa con tu email y contraseña
2. Permite el acceso a **cámara** y **ubicación** cuando el navegador lo pida
3. Captura tu foto con el botón "Capturar foto"
4. Presiona **Registrar entrada** al llegar o **Registrar salida** al irte
5. En "Mi historial" puedes ver tus propios registros del mes

---

## PREGUNTAS FRECUENTES

**¿Puedo usarlo desde el celular?**
Sí. La URL funciona en cualquier navegador moderno (Chrome, Safari, Edge).

**¿Qué pasa si un empleado no tiene cámara?**
El sistema acepta el registro pero sin foto. Se puede configurar para bloquear marcajes sin foto.

**¿Se puede usar en tablet fija en la entrada?**
Sí. Abre la URL en Chrome, pantalla completa. Cada empleado selecciona su nombre, se toma la foto y marca.

**¿Los datos se pierden si se reinicia el servidor?**
No, siempre que configures el volumen persistente en Railway (paso 5 de la guía).

**¿Puedo cambiar la hora del reporte automático?**
Sí. En `server.js`, busca `cron.schedule('0 19 * * 1-6'` y cambia `19` por la hora que quieras.

---

## SOPORTE Y MEJORAS FUTURAS

Ideas que puedes pedir agregar:
- Reconocimiento facial automático (compara la foto con el perfil)
- Notificación WhatsApp cuando alguien llega tarde
- Geofencing (bloquear marcaje si está fuera de la oficina)
- Control de turnos (mañana/tarde/noche)
- App móvil nativa (Android/iOS)
- Gestión de vacaciones y permisos
