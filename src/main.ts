import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';
import { json, urlencoded } from 'express'; // <-- 1. IMPORTAR ESTO

async function bootstrap() {
  const app = await NestFactory.create(AppModule);

  // ▼▼▼ 2. AÑADIR ESTAS LÍNEAS ▼▼▼
  // Aumentamos el límite para JSON (ej: el array de gastos)
  app.use(json({ limit: '50mb' }));
  // Aumentamos el límite para datos de formularios (aunque no lo uses, es buena práctica)
  app.use(urlencoded({ extended: true, limit: '50mb' }));
  // ▲▲▲ FIN DE LA MODIFICACIÓN ▲▲▲


  app.enableCors(); // Habilitar CORS para todas las rutas
  await app.listen(process.env.PORT ?? 3000);
}
bootstrap();
