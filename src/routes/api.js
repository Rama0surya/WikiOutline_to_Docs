const express = require('express');
const router = express.Router();
const axios = require('axios');
const { buildDocx } = require('../docxBuilder');
const archiver = require('archiver');

// Helper: forward POST to Outline API
async function outlinePost(baseUrl, token, path, body) {
  const url = baseUrl.replace(/\/$/, '') + path;
  const res = await axios.post(url, body, {
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    timeout: 30000,
  });
  return res.data;
}

// GET all collections
router.post('/collections.list', async (req, res) => {
  const { outlineUrl, apiToken } = req.body;
  if (!outlineUrl || !apiToken)
    return res.status(400).json({ error: 'outlineUrl and apiToken are required' });
  try {
    const data = await outlinePost(outlineUrl, apiToken, '/api/collections.list', { limit: 100 });
    res.json(data);
  } catch (e) {
    res.status(502).json({ error: e.response?.data?.message || e.message });
  }
});

// GET all documents with auto-pagination
router.post('/documents.list', async (req, res) => {
  const { outlineUrl, apiToken, collectionId } = req.body;
  if (!outlineUrl || !apiToken)
    return res.status(400).json({ error: 'outlineUrl and apiToken are required' });
  try {
    let allDocs = [];
    let offset = 0;
    const limit = 100;
    while (true) {
      const body = { limit, offset };
      if (collectionId) body.collectionId = collectionId;
      const data = await outlinePost(outlineUrl, apiToken, '/api/documents.list', body);
      const docs = data.data || [];
      allDocs = allDocs.concat(docs);
      // Stop when we get less than limit (last page)
      if (docs.length < limit) break;
      offset += limit;
    }
    res.json({ data: allDocs });
  } catch (e) {
    res.status(502).json({ error: e.response?.data?.message || e.message });
  }
});

// Export one or multiple documents to .docx
// Returns: single .docx if 1 doc or merge=true; ZIP if multiple separate docs
router.post('/documents.export', async (req, res) => {
  const { outlineUrl, apiToken, ids, opts } = req.body;
  if (!outlineUrl || !apiToken || !ids?.length)
    return res.status(400).json({ error: 'outlineUrl, apiToken, and ids[] are required' });

  try {
    // Fetch all document contents in parallel (up to 5 concurrent)
    const fetchDoc = async (id) => {
      const [exportRes, infoRes] = await Promise.all([
        outlinePost(outlineUrl, apiToken, '/api/documents.export', { id }),
        outlinePost(outlineUrl, apiToken, '/api/documents.info', { id }),
      ]);
      return {
        id,
        title: infoRes.data?.title || 'Untitled',
        markdown: exportRes.data || '',
      };
    };

    // Batch fetch with concurrency limit of 5
    const docs = [];
    for (let i = 0; i < ids.length; i += 5) {
      const batch = ids.slice(i, i + 5);
      const results = await Promise.all(batch.map(fetchDoc));
      docs.push(...results);
    }

    const merge = opts?.merge === true;

    if (merge || docs.length === 1) {
      // Single .docx output
      const buffer = await buildDocx(docs, opts, outlineUrl, apiToken);
      const filename = docs.length === 1
        ? sanitizeFilename(docs[0].title) + '.docx'
        : 'outline-export.docx';
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
      res.send(buffer);
    } else {
      // Multiple .docx → ZIP
      res.setHeader('Content-Type', 'application/zip');
      res.setHeader('Content-Disposition', 'attachment; filename="outline-export.zip"');
      const archive = archiver('zip', { zlib: { level: 6 } });
      archive.pipe(res);
      for (const doc of docs) {
        const buf = await buildDocx([doc], opts, outlineUrl, apiToken);
        archive.append(buf, { name: sanitizeFilename(doc.title) + '.docx' });
      }
      await archive.finalize();
    }
  } catch (e) {
    console.error('Export error:', e.message);
    if (!res.headersSent) {
      res.status(502).json({ error: e.response?.data?.message || e.message });
    }
  }
});

function sanitizeFilename(name) {
  return name.replace(/[^\w\s-]/g, '').replace(/\s+/g, '-').toLowerCase().slice(0, 80) || 'document';
}

module.exports = router;
