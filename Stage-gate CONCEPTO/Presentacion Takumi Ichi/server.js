/**
 * Takumi Ichi - Centralized Server & API for Team Sales Projections
 * Node.js Express Server
 */
const express = require('express');
const fs = require('fs');
const path = require('path');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 8085;
const DB_FILE = path.join(__dirname, 'projections_db.json');

app.use(cors());
app.use(express.json());
app.use(express.static(__dirname));

// Ensure DB file exists
if (!fs.existsSync(DB_FILE)) {
  fs.writeFileSync(DB_FILE, JSON.stringify([], null, 2));
}

// GET all shared projections
app.get('/api/projections', (req, res) => {
  try {
    const data = JSON.parse(fs.readFileSync(DB_FILE, 'utf-8'));
    res.json({ success: true, data });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// POST save / update projections
app.post('/api/projections', (req, res) => {
  try {
    const newProjections = req.body;
    if (Array.isArray(newProjections)) {
      fs.writeFileSync(DB_FILE, JSON.stringify(newProjections, null, 2));
      res.json({ success: true, count: newProjections.length });
    } else {
      res.status(400).json({ success: false, error: 'Expected array of projections' });
    }
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// POST add single projection
app.post('/api/projections/add', (req, res) => {
  try {
    const current = JSON.parse(fs.readFileSync(DB_FILE, 'utf-8'));
    const item = req.body;
    current.unshift(item);
    fs.writeFileSync(DB_FILE, JSON.stringify(current, null, 2));
    res.json({ success: true, item });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// DELETE projection by ID
app.delete('/api/projections/:id', (req, res) => {
  try {
    const current = JSON.parse(fs.readFileSync(DB_FILE, 'utf-8'));
    const filtered = current.filter(p => p.id !== req.params.id);
    fs.writeFileSync(DB_FILE, JSON.stringify(filtered, null, 2));
    res.json({ success: true, remaining: filtered.length });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`=======================================================`);
  console.log(`  Takumi Ichi Server running at http://0.0.0.0:${PORT}`);
  console.log(`  Shared Database active: projections_db.json`);
  console.log(`=======================================================`);
});
