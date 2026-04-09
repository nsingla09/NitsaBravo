import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import { createServer as createViteServer } from 'vite';
import admin from 'firebase-admin';
import fs from 'fs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Load firebase config for admin initialization
const firebaseConfig = JSON.parse(fs.readFileSync(path.join(process.cwd(), 'firebase-applet-config.json'), 'utf-8'));

// Initialize Firebase Admin
try {
  if (!admin.apps.length) {
    admin.initializeApp({
      projectId: firebaseConfig.projectId,
    });
  }
} catch (error) {
  console.error('Firebase Admin initialization failed:', error);
}

async function startServer() {
  const app = express();
  const PORT = 3000;

  // API Route for Firestore Data Bundles
  app.get('/api/bundle', async (req, res) => {
    try {
      const db = admin.firestore(firebaseConfig.firestoreDatabaseId);
      const bundle = db.bundle('global-static-data');

      // Add queries to the bundle
      const rolesQuery = db.collection('roles').limit(50);
      const trainingQuery = db.collection('trainingMaterials').limit(100);
      const weeksQuery = db.collection('weeks').limit(100);
      const bdesQuery = db.collection('bdes').limit(100);
      const passwordsQuery = db.collection('pagePasswords');

      // Build the bundle
      const bundleBuffer = await bundle
        .add('roles-query', await rolesQuery.get())
        .add('training-query', await trainingQuery.get())
        .add('weeks-query', await weeksQuery.get())
        .add('bdes-query', await bdesQuery.get())
        .add('passwords-query', await passwordsQuery.get())
        .build();

      res.setHeader('Content-Type', 'application/octet-stream');
      res.setHeader('Cache-Control', 'public, max-age=3600'); // Cache for 1 hour
      res.send(bundleBuffer);
    } catch (error) {
      console.error('Error generating bundle:', error);
      res.status(500).json({ error: 'Failed to generate bundle' });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    // Serve static files from the dist directory in production
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    
    // Handle SPA routing: serve index.html for all non-file requests
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
