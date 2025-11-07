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


  // ▼▼▼ MODIFICACIÓN DE CORS ▼▼▼
  app.enableCors({
    origin: [
      'http://localhost:4200', // Para tu desarrollo local
      // Añadiremos la URL de Vercel aquí después
    ],
    methods: 'GET,HEAD,PUT,PATCH,POST,DELETE',
  });
  // ▲▲▲ FIN DE LA MODIFICACIÓN ▲▲▲
  
  await app.listen(process.env.PORT ?? 3000);
}
bootstrap();
