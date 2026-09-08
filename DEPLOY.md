# Deploy en Vercel

1. Importa este proyecto en Vercel como proyecto estático.
2. En **Settings → Environment Variables**, agrega estas variables para Production, Preview y Development:

   - `TEAMCO_USER`: `teamco`
   - `TEAMCO_PASSWORD`: la clave entregada para Teamco
   - `AUTH_SECRET`: una cadena aleatoria larga y privada

3. Haz un nuevo deploy.

La aplicación usa una cookie `HttpOnly`, `Secure` y con expiración de 8 horas. La contraseña no está guardada en los archivos del proyecto; solo se compara en la función `/api/login` usando las variables de entorno de Vercel.
