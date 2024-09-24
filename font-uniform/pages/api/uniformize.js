import { exec } from 'child_process';
import formidable from 'formidable';
import fs from 'fs';
import path from 'path';

const uploadDir = path.join(process.cwd(), 'uploads');

if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir);
}

export const config = {
  api: {
    bodyParser: false, // Disables Next.js's default body parsing
  },
};

export default async function handler(req, res) {
  if (req.method === 'POST') {
    const form = formidable({ uploadDir, keepExtensions: true });

    form.parse(req, (err, fields, files) => {
      if (err) {
        console.error('Error parsing the files:', err);
        res.status(500).json({ error: 'Error parsing the file' });
        return;
      }

      // Log files and fields to debug
      console.log('Fields:', fields);
      console.log('Files:', files);

      // Access the uploaded file
      const uploadedFile = files.file && files.file[0];

      if (!uploadedFile) {
        res.status(400).json({ error: 'No file uploaded' });
        return;
      }

      // Depending on the version of formidable, use filepath or path
      const filePath = uploadedFile.filepath || uploadedFile.path;

      if (!filePath) {
        res.status(500).json({ error: 'File path not found' });
        return;
      }

      const targetFont = fields.targetFont || 'Arial';

      const outputFilePath = filePath.replace('.pptx', '_updated.pptx');

      // Update the path to the Python script
      const pythonScriptPath = path.join(process.cwd(), 'powerpointfont.py');

      // Execute the Python script
      exec(
        `python3 "${pythonScriptPath}" "${filePath}" "${outputFilePath}" "${targetFont}"`,
        (error, stdout, stderr) => {
          if (error) {
            console.error('Error executing Python script:', stderr);
            res.status(500).json({ error: stderr });
            return;
          }

          console.log('Python script output:', stdout);

          res.status(200).json({
            message: 'File processed successfully',
            outputFilePath,
          });
        }
      );
    });
  } else {
    res.status(405).json({ error: 'Method not allowed' });
  }
}
