const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const excelHandler = require('./excel-handler');

const app = express();
app.use(express.json());

// Use memory storage for multer
const upload = multer({
    storage: multer.memoryStorage(),
    limits: {
        fileSize: 5 * 1024 * 1024 // 5MB
    }
});

// Upload and convert warehouse + priority Excel
app.post('/api/convert-location-priority-excel', upload.fields([
    { name: 'warehouseFile', maxCount: 1 },
    { name: 'priorityFile', maxCount: 1 }
]), async (req, res) => {
    try {
        if (!req.files?.warehouseFile || !req.files?.priorityFile || !req.body.eshipz_user_id) {
            return res.status(400).json({ error: 'Both Excel files and eshipz_user_id are required' });
        }

        const warehouseBuffer = req.files.warehouseFile[0].buffer;
        const priorityBuffer = req.files.priorityFile[0].buffer;
        const eshipzUserId = req.body.eshipz_user_id;

        const result = await excelHandler.convertLocationPriorityExcelFiles(
            warehouseBuffer,
            priorityBuffer,
            eshipzUserId
        );
     

        
      

        return res.status(200).json({
            success:true,
           data: result,
           
        });
    } catch (err) {
        console.error('Location Priority Error:', err);
        return res.status(500).json({ error: 'Processing error', details: err.message });
    }
});

// Convert warehouse Excel
app.post('/api/convert-warehouse-excel', upload.single('warehouseFile'), async (req, res) => {
    try {
        if (!req.file || !req.body.eshipz_user_id) {
            return res.status(400).json({ error: 'warehouseFile and eshipz_user_id required' });
        }

        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        const header = data[0];
        const rows = data.slice(1);

        const structured = rows.map(row => {
            const obj = {};
            row.forEach((val, i) => {
                if (header[i]) obj[header[i].trim()] = val;
            });
            return obj;
        }).filter(row => Object.keys(row).length > 0);

        const result = await excelHandler.convertWarehouseExcelToJson(structured, req.body.eshipz_user_id);

        return res.status(200).json({
            success:true,
            data: result,
            
        });
    } catch (err) {
        console.error('Warehouse Error:', err);
        return res.status(500).json({ error: 'Warehouse processing failed', details: err.message });
    }
});

// Convert customer SLA Excel
app.post('/api/convert-customersla-excel', upload.single('customerslaFile'), async (req, res) => {
    try {
        if (!req.file || !req.body.eshipz_user_id) {
            return res.status(400).json({ error: 'customerslaFile and eshipz_user_id required' });
        }

        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        const header = data[0];
        const rows = data.slice(1);

        const structured = rows.map(row => {
            const obj = {};
            row.forEach((val, i) => {
                if (header[i]) obj[header[i].trim()] = val;
            });
            return obj;
        }).filter(row => Object.keys(row).length > 0);

        const result = await excelHandler.convertCustomerslaExcelToJson(structured, req.body.eshipz_user_id);
        const allSlas = result.allSlas;

        return res.status(200).json({
            success: true,
            data: allSlas
        });
    } catch (err) {
        console.error('Customer SLA Error:', err);
        return res.status(500).json({ error: 'Customer SLA processing failed', details: err.message });
    }
});

app.get('/health', (req, res) => {
    res.status(200).json({ status: 'ok' });
});

// Start server
const PORT = process.env.PORT || 4000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on port ${PORT}`);
});