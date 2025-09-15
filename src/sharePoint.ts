import express from 'express';
import { ConfidentialClientApplication } from '@azure/msal-node';
import { Client } from '@microsoft/microsoft-graph-client';
import type { AuthenticationProvider } from '@microsoft/microsoft-graph-client';
import multer from 'multer';
import contentDisposition from 'content-disposition';
import 'isomorphic-fetch';
import dotenv from 'dotenv';
import axios from 'axios';

// Load .env early
dotenv.config();

// SharePoint configuration (read from environment variables)
const TENANT_ID = process.env.TENANT_ID || '';
const CLIENT_ID = process.env.CLIENT_ID || '';
const CLIENT_SECRET = process.env.CLIENT_SECRET || '';
const SITE_ID = process.env.SITE_ID || '';
const SHAREPOINT_SITE_URL = process.env.SHAREPOINT_SITE_URL || '';
const PORT = process.env.PORT || 3001;

// Basic validation — fail fast with a helpful message if required values are missing
const missing = [] as string[];
if (!TENANT_ID) missing.push('TENANT_ID');
if (!CLIENT_ID) missing.push('CLIENT_ID');
if (!CLIENT_SECRET) missing.push('CLIENT_SECRET');
if (!SITE_ID) missing.push('SITE_ID');
if (!SHAREPOINT_SITE_URL) missing.push('SHAREPOINT_SITE_URL');
if (missing.length > 0) {
  console.error('Missing required environment variables:', missing.join(', '));
  console.error('Please add them to your .env file or export them in the shell.');
  // Exit early so msal doesn't throw confusing errors about empty credentials
  process.exit(1);
}

// MSAL configuration for app-only authentication
const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    clientSecret: CLIENT_SECRET,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
  },
};

// Custom authentication provider for app-only access
class AppOnlyAuthProvider implements AuthenticationProvider {
  private msalInstance: ConfidentialClientApplication;

  constructor() {
    this.msalInstance = new ConfidentialClientApplication(msalConfig);
  }

  async getAccessToken(): Promise<string> {
    const clientCredentialRequest = {
      scopes: ['https://graph.microsoft.com/.default'],
      skipCache: false,
    };

    try {
      const response = await this.msalInstance.acquireTokenByClientCredential(
        clientCredentialRequest
      );
      return response?.accessToken || '';
    } catch (error) {
      console.error('Error acquiring token:', error);
      throw error;
    }
  }
}

// Initialize Graph client
const authProvider = new AppOnlyAuthProvider();
const graphClient = Client.initWithMiddleware({ authProvider });

// Express app setup
const app: express.Application = express();
app.use(express.json());

// Configure multer for file uploads (store in memory for processing)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 250 * 1024 * 1024 }, // 250MB
});


// Constants for upload handling
const SMALL_FILE_THRESHOLD = 4 * 1024 * 1024; // 4MB - threshold for simple vs resumable upload

// Helper function to get drive ID for the "Dokumenty" folder
async function getDriveId(): Promise<string> {
  try {
    const site = await graphClient.api(`/sites/${SITE_ID}`).get();
    const drives = await graphClient.api(`/sites/${SITE_ID}/drives`).get();
    
    // Look for the default document library (Dokumenty/Documents)
    const documentsDrive = drives.value.find((drive: any) => 
      drive.name === 'Documents' || drive.name === 'Dokumenty'
    );
    
    if (documentsDrive) {
      return documentsDrive.id;
    }
    
    // Fallback to first drive if Documents not found
    return drives.value[0]?.id;
  } catch (error) {
    console.error('Error getting drive ID:', error);
    throw error;
  }
}

// Helper function for small file upload (< 4MB)
async function uploadSmallFile(driveId: string, folder: string, fileName: string, buffer: Buffer): Promise<any> {
  return await graphClient
    .api(`/drives/${driveId}/root:/${folder}/${fileName}:/content`)
    .putStream(buffer);
}

// Helper function for large file upload (>= 4MB) using resumable upload
async function uploadLargeFile(driveId: string, folder: string, fileName: string, buffer: Buffer): Promise<any> {
  try {
        console.log('Starting large file upload...');
    // Create upload session
    const uploadSession = await graphClient
      .api(`/drives/${driveId}/root:/${folder}/${fileName}:/createUploadSession`)
      .post({
        item: {
          '@microsoft.graph.conflictBehavior': 'replace',
          name: fileName,
        },
      });

    const uploadUrl = uploadSession.uploadUrl;
    const fileSize = buffer.length;
    const chunkSize = 320 * 1024; // 320KB chunks (must be multiple of 320KB for Graph API)
    
    let bytesUploaded = 0;
    
    // Upload file in chunks
    while (bytesUploaded < fileSize) {
      const start = bytesUploaded;
      const end = Math.min(bytesUploaded + chunkSize, fileSize);
      const chunk = buffer.slice(start, end);
      
      const response = await fetch(uploadUrl, {
        method: 'PUT',
        headers: {
          'Content-Range': `bytes ${start}-${end - 1}/${fileSize}`,
          'Content-Length': chunk.length.toString(),
        },
        body: chunk,
      });

      if (response.status === 202) {
        // Continue uploading
        bytesUploaded = end;
        console.log(`Uploaded ${bytesUploaded}/${fileSize} bytes (${Math.round((bytesUploaded / fileSize) * 100)}%)`);
      } else if (response.status === 201 || response.status === 200) {
        // Upload complete
        const result = await response.json();
        console.log('Large file upload completed successfully');
        return result;
      } else {
        const errorText = await response.text();
        throw new Error(`Upload failed with status ${response.status}: ${errorText}`);
      }
    }
  } catch (error) {
    console.error('Large file upload error:', error);
    throw error;
  }
}

const HELP = `
SharePoint CRUD API server running on http://localhost:${PORT}
SharePoint Site: ${SHAREPOINT_SITE_URL}
Site ID: ${SITE_ID}

Available endpoints:
  GET    /                                           - API overview
  GET    /health                                     - Health check
  GET    /files[/<folder>]                           - List all files
  GET    /files[/<folder>]/<file>/detail             - Get file info
  POST   /files[/<folder>]/upload                    - Upload file
  PUT    /files[/<folder>]/<file>                    - Update file
  GET    /files[/<folder>]/<file>/download           - Download file
  GET    /files[/<folder>]/<file>/download-redirect  - Download using redirect
  DELETE /files[/<folder>]/<file>                    - Delete file
  GET    /folders[/<folder>]                         - List folders
  POST   /folders[/<subfolder>]/<folder>             - Create folder
  DELETE /folders[/<subfolder>]/<folder>             - Delete folder
`

// Routes

// GET / - API overview
app.get('/', (req, res) => {
  res.json({
    message: 'SharePoint CRUD API',
    help: HELP.split('\n').map(line => line.trim()),
    uploadInfo: {
      smallFiles: 'Files < 4MB use simple upload',
      largeFiles: 'Files >= 4MB use resumable upload with progress tracking',
      maxFileSize: '250MB',
      endpoint: '/files/upload (multipart form) or /files (JSON)',
    },
    siteInfo: {
      siteUrl: SHAREPOINT_SITE_URL,
      siteId: SITE_ID,
      targetFolder: 'Dokumenty',
    },
  });
});

// GET /files[/folder]/file/download - Download a file using axios streaming
app.get('/files/*splat/download', async (req, res) => {
  try {
    const { splat } = req.params;
    const filePath = splat.join('/');
    const driveId = await getDriveId();

    // Get file metadata to set headers
    const fileMeta = await graphClient
      .api(`/drives/${driveId}/root:/${encodeURIComponent(filePath)}`)
      .select('id,name,size')
      .get();

    if (!fileMeta || !fileMeta.id) {
      return res.status(404).json({ error: 'File not found' });
    }

    // Set headers for browser download
    res.setHeader('Content-Disposition', contentDisposition(fileMeta.name));
    res.setHeader('Content-Type', 'application/octet-stream');

    // Get access token for axios request
    const accessToken = await authProvider.getAccessToken();
    
    // Use axios to stream the file content
    const downloadUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodeURIComponent(filePath)}:/content`;


    const axiosResponse = await axios({
      method: 'GET',
      url: downloadUrl,
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Accept': 'application/octet-stream'
      },
      responseType: 'stream'
    });

    // Set content length if available
    const contentLength = axiosResponse.headers['content-length'];
    if (contentLength) {
      res.setHeader('Content-Length', contentLength);
    }

    // Pipe the axios response stream directly to the client
    axiosResponse.data.pipe(res);

    // Error handling
    axiosResponse.data.on('error', (err: any) => {
      console.error('Axios stream error:', err);
      if (!res.headersSent) {
        res.status(500).end('Error streaming file');
      }
    });

    // If client aborts, destroy the axios stream
    res.on('close', () => {
      try {
        if (axiosResponse.data.destroy) {
          axiosResponse.data.destroy();
        }
      } catch (e) {
        console.log('Error destroying axios stream:', (e as any)?.message || e);
      }
    });

  } catch (error: any) {
    console.error('Error downloading file with axios:', error);
    if (!res.headersSent) {
      if (error.response?.status === 404 || error.code === 'itemNotFound') {
        res.status(404).json({ error: 'File not found' });
      } else {
        res.status(500).json({ 
          error: 'Failed to download file', 
          details: error.message,
          status: error.response?.status 
        });
      }
    }
  }
});

// GET /files[/folder]/file/download-redirect - Direct redirect to SharePoint download URL (if available)
app.get('/files/*splat/download-redirect', async (req, res) => {
  try {
    const { splat } = req.params;
    const filePath = splat.join('/');
    const driveId = await getDriveId();
    
    console.log('Attempting to get direct download URL for:', filePath);
    
    // Get file info including potential download URL
    const file = await graphClient
      .api(`/drives/${driveId}/root:/${encodeURIComponent(filePath)}`)
      .select('id,name,size,@microsoft.graph.downloadUrl,@content.downloadUrl,webUrl')
      .get();

    // Check various possible download URL fields
    const downloadUrl = file['@microsoft.graph.downloadUrl'] || file['@content.downloadUrl'];
    
    if (downloadUrl) {
      console.log('Found download URL:', downloadUrl);
      // Redirect browser directly to Microsoft's download URL
      res.redirect(downloadUrl);
    } else {      
      res.status(404).json({ 
        error: 'Direct download URL not available',
        response: file,
      });
    }

  } catch (error: any) {
    console.error('Error getting download URL:', error);
    if (error.code === 'itemNotFound') {
      res.status(404).json({ error: 'File not found' });
    } else {
      res.status(500).json({ error: 'Failed to get download URL', details: error.message });
    }
  }
});

// GET /files[/folder]/file/detail - Get specific file info
app.get('/files/*splat/detail', async (req, res) => {
  try {
    const { splat }  = req.params;
    const filePath = splat.join('/');
    const driveId = await getDriveId();
    
    const file = await graphClient
      .api(`/drives/${driveId}/root:/${encodeURIComponent(filePath)}`)
      .select('id,name,size,lastModifiedDateTime,webUrl,@microsoft.graph.downloadUrl')
      .get();

    res.json({
      id: file.id,
      name: file.name,
      size: file.size,
      lastModified: file.lastModifiedDateTime,
      webUrl: file.webUrl,
      downloadUrl: file['@microsoft.graph.downloadUrl'],
    });
  } catch (error: any) {
    console.error('Error getting file:', error);
    if (error.code === 'itemNotFound') {
      res.status(404).json({ error: 'File not found' });
    } else {
      res.status(500).json({ error: 'Failed to get file', details: error });
    }
  }
});

// GET /files[/folder] - List all files
app.get('/files{/*splat}', async (req, res) => {
  try {
    const { splat } = req.params;
    const folder = (splat ?? []).join('/');
    const driveId = await getDriveId();
    const requestPath = !!folder ? 
    `/drives/${driveId}/root:/${folder}:/children` :
    `/drives/${driveId}/root/children`;
    const files = await graphClient
      .api(requestPath)
      .select('id,name,size,lastModifiedDateTime,webUrl,@microsoft.graph.downloadUrl,file')
      .get();

    // Filter only files (not folders)
    const fileItems = files.value.filter((item: any) => item.file);

    res.json({
      count: fileItems.length,
      files: fileItems.map((file: any) => ({
        id: file.id,
        name: file.name,
        size: file.size,
        lastModified: file.lastModifiedDateTime,
        webUrl: file.webUrl,
        downloadUrl: file['@microsoft.graph.downloadUrl'],
      })),
    });
  } catch (error) {
    console.error('Error listing files:', error);
    res.status(500).json({ error: 'Failed to list files', details: error });
  }
});

// POST /files[/folder]/upload - Upload a file using multipart/form-data
app.post(
  '/files{/*splat}/upload',
  upload.single('file'),
  async (req, res) => {
    try {
      const { splat } = req.params;
      const folder = (splat ?? []).join('/');

      if (!req.file) {
        return res.status(400).json({ error: 'No file provided. Use "file" field in multipart form.' });
      }

      const { originalname, buffer, size, mimetype } = req.file;
      const fileName = req.body.fileName || originalname;

      if (!fileName) {
        return res.status(400).json({ error: 'fileName is required' });
      }

      console.log(`Uploading file: ${fileName} (${size} bytes, ${mimetype})`);

      const driveId = await getDriveId();
      let uploadedFile;

      // Choose upload method based on file size
      if (size < SMALL_FILE_THRESHOLD) {
        console.log('Using simple upload for small file');
        uploadedFile = await uploadSmallFile(driveId, folder, fileName, buffer);
      } else {
        console.log('Using resumable upload for large file');
        uploadedFile = await uploadLargeFile(driveId, folder, fileName, buffer);
      }

      res.status(201).json({
        message: 'File uploaded successfully',
        response: uploadedFile,
        uploadMethod: size < SMALL_FILE_THRESHOLD ? 'simple' : 'resumable',
        originalSize: size,
      });
    } catch (error: any) {
      console.error('Error uploading file:', error);
      res.status(500).json({ 
        error: 'Failed to upload file', 
        details: error.message,
        stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
      });
    }
});

// PUT /files[/folder]/file - Update a file
app.put(
  '/files/*splat',
  upload.single('file'),
  async (req, res) => {
    try {
      const { splat } = req.params;
      const filePath = splat.join('/');
      
      if (!req.file) {
        return res.status(400).json({ error: 'No file provided. Use "file" field in multipart form.' });
      }
      const { buffer } = req.file;

      const driveId = await getDriveId();
      
      const updatedFile = await graphClient
        .api(`/drives/${driveId}/root:/${filePath}:/content`)
        .putStream(buffer);

      res.json({
        message: 'File updated successfully',
        response: updatedFile,
      });
    } catch (error: any) {
      console.error('Error updating file:', error);
      if (error.code === 'itemNotFound') {
        res.status(404).json({ error: 'File not found' });
      } else {
        res.status(500).json({ error: 'Failed to update file', details: error });
      }
    }
});

// DELETE /files[/folder]/file - Delete a file
app.delete('/files/*splat', async (req, res) => {
  try {
    const { splat } = req.params;
    const filePath = splat.join('/');
    const driveId = await getDriveId();

    const response = await graphClient.api(`/drives/${driveId}/root:/${filePath}`).delete();

    res.json({ message: 'File deleted successfully', filePath, response });
  } catch (error: any) {
    console.error('Error deleting file:', error);
    if (error.code === 'itemNotFound') {
      res.status(404).json({ error: 'File not found' });
    } else {
      res.status(500).json({ error: 'Failed to delete file', details: error });
    }
  }
});

// GET /folders[/folder] - List all folders
app.get('/folders{/*splat}', async (req, res) => {
  try {
    const { splat } = req.params;
    const folder = (splat ?? []).join('/');
    const driveId = await getDriveId();
    const requestPath = !!folder ? 
    `/drives/${driveId}/root:/${folder}:/children` :
    `/drives/${driveId}/root/children`;
    const items = await graphClient
      .api(requestPath)
      .select('id,name,lastModifiedDateTime,webUrl,folder')
      .get();

    // Filter only folders
    const folders = items.value.filter((item: any) => item.folder);

    res.json({
      count: folders.length,
      folders: folders.map((folder: any) => ({
        id: folder.id,
        name: folder.name,
        lastModified: folder.lastModifiedDateTime,
        webUrl: folder.webUrl,
        childCount: folder.folder.childCount,
      })),
    });
  } catch (error) {
    console.error('Error listing folders:', error);
    res.status(500).json({ error: 'Failed to list folders', details: error });
  }
});

// POST /folders[/subfolder]/folder - Create a new folder
app.post('/folders/*splat', async (req, res) => {
  try {
    const { splat } = req.params;
    const folderName = splat.join('/');

    if (!folderName) {
      return res.status(400).json({ error: 'folderName is required' });
    }

    const driveId = await getDriveId();

    const newFolder = await graphClient
      .api(`/drives/${driveId}/root:/${folderName}`)
      .patch({
        folder: {},
        '@microsoft.graph.conflictBehavior': 'fail',
      });

    res.status(201).json({
      message: 'Folder created successfully',
      folder: {
        id: newFolder.id,
        name: newFolder.name,
        webUrl: newFolder.webUrl,
      },
    });
  } catch (error) {
    console.error('Error creating folder:', error);
    res.status(500).json({ error: 'Failed to create folder', details: error });
  }
});

// DELETE /folders[/subfolder]/folder - Delete a folder
app.delete('/folders/*splat', async (req, res) => {
  try {
    const { splat } = req.params;
    const folderName = splat.join('/')
    
    if (!folderName) {
      return res.status(400).json({ error: 'folderName is required' });
    }

    const driveId = await getDriveId();

    await graphClient.api(`/drives/${driveId}/root:/${folderName}`).delete();

    res.json({ message: 'Folder deleted successfully', folderName });
  } catch (error: any) {
    console.error('Error deleting folder:', error);
    if (error.code === 'itemNotFound') {
      res.status(404).json({ error: 'Folder not found' });
    } else {
      res.status(500).json({ error: 'Failed to delete folder', details: error });
    }
  }
});

// Health check endpoint
app.get('/health', async (req, res) => {
  try {
    // Test the connection to SharePoint
    const site = await graphClient.api(`/sites/${SITE_ID}`).get();
    
    res.json({
      status: 'healthy',
      siteInfo: {
        id: site.id,
        name: site.name,
        webUrl: site.webUrl,
      },
      timestamp: new Date().toISOString(),
    });
  } catch (error) {
    console.error('Health check failed:', error);
    res.status(500).json({
      status: 'unhealthy',
      error: 'Failed to connect to SharePoint',
      timestamp: new Date().toISOString(),
    });
  }
});

// Error handling middleware
app.use((err: any, req: express.Request, res: express.Response, next: express.NextFunction) => {
  console.error('Unhandled error:', err);
  res.status(500).json({ error: 'Internal server error' });
});

// Start the server
app.listen(PORT, () => {
  console.log(HELP);
});

export default app;
