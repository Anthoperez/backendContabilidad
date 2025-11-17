// src/main.ts
import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';
import { json, urlencoded } from 'express';

async function bootstrap() {
  const app = await NestFactory.create(AppModule);

  // Aumentar límites para importación de Excel
  app.use(json({ limit: '50mb' }));
  app.use(urlencoded({ extended: true, limit: '50mb' }));

  // CORS configurado para producción
  app.enableCors({
    origin: process.env.FRONTEND_URL || '*', // <-- Se configurará después
    credentials: true,
  });

  const port = process.env.PORT || 3000;
  await app.listen(port, '0.0.0.0'); // <-- IMPORTANTE: '0.0.0.0' para Render
  console.log(`🚀 Backend corriendo en puerto ${port}`);
}
bootstrap();