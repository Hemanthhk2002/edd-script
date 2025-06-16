const XLSX = require('xlsx');

// Process Location Priority Excel Files
async function convertLocationPriorityExcelFiles(warehouseFile, priorityFile,eshipzUserId) {
    try {
        // Read Excel files
        const warehouseWorkbook = XLSX.read(warehouseFile);
        const priorityWorkbook = XLSX.read(priorityFile);
        
        // Get first sheet
        const warehouseSheet = warehouseWorkbook.Sheets[warehouseWorkbook.SheetNames[0]];
        const prioritySheet = priorityWorkbook.Sheets[priorityWorkbook.SheetNames[0]];
        
        // Convert to JSON with explicit headers
        const warehouseData = XLSX.utils.sheet_to_json(warehouseSheet, { header: 1 });
        const priorityData = XLSX.utils.sheet_to_json(prioritySheet, { header: 1 });
        
        // Process data
        const warehouseMap = new Map();
        const priorityMap = new Map();
        
        // Skip header row and process warehouse pincodes
        warehouseData.slice(1).forEach(row => {
            if (row.length >= 2) { // Ensure we have both columns
                const warehouse = row[1]; // Mother_Warehouse is in second column
                const pincode = row[0].toString(); // Destination_Pincode is in first column
                if (!warehouseMap.has(warehouse)) {
                    warehouseMap.set(warehouse, []);
                }
                warehouseMap.get(warehouse).push(pincode);
            }
        });
        
        // Skip header row and process priorities
        priorityData.slice(1).forEach(row => {
            if (row.length >= 3) { // Ensure we have all columns
                const warehouse = row[0]; // Mother_Warehouse is in first column
                const shippingWarehouse = row[1]; // Shipping_Warehouse is in second column
                if (!priorityMap.has(warehouse)) {
                    priorityMap.set(warehouse, []);
                }
                priorityMap.get(warehouse).push(shippingWarehouse);
            }
        });
        
        // Sort priorities by their order in the Excel file
        const sortedPriorityMap = new Map();
        priorityMap.forEach((priorities, warehouse) => {
            const sortedPriorities = [warehouse, ...priorities];
            sortedPriorityMap.set(warehouse, sortedPriorities);
        });
        
        // Create documents
        const documents = [];
        // const promises = [];// for mongodb operations
        
        warehouseMap.forEach((pincodes, warehouse) => {
            const pickupLocations = sortedPriorityMap.get(warehouse) || [];
            
            // Split into chunks of 240
            for (let i = 0; i < pincodes.length; i += 240) {
                const chunk = pincodes.slice(i, i + 240);
                const setNumber = Math.floor(i / 240) + 1;
                
                const doc = {
                    eshipz_user_id: eshipzUserId,
                    delivery_pincodes_set: setNumber,
                    mother_warehouse: warehouse,
                    delivery_pincodes: chunk,
                    pickup_location_names: pickupLocations
                };
                documents.push(doc);
                

                //  Save to MongoDB (commented out for now)
                // promises.push(saveLocationPriority(doc));
            }
            
        });

        // Wait for all MongoDB operations to complete
        // await Promise.all(promises);

        // Save to JSON files
        // documents.forEach(doc => {
        //     const fileName = `${doc.mother_warehouse}_set${doc.delivery_pincodes_set}.json`;
        //     const filePath = path.join(OUTPUT_DIR, fileName);
        //     fs.writeFileSync(filePath, JSON.stringify(doc, null, 2));
        // });

        return documents;
    } catch (error) {
        console.error('Error processing Excel files:', error);
        throw error;
    }
}

// Convert Warehouse Excel to JSON
const convertWarehouseExcelToJson = async (excelData, eshipzUserId) => {
    const result = [];

    excelData.forEach(row => {
        if (!row['Warehouse Name'] && !row['warehouse_name'] &&
            !row['Pincode'] && !row['warehouse_location_pincode'] &&
            !row['Closing Time'] && !row['closing_time'] &&
            !row['Non-Operational Dates'] && !row['non_operational_dates']) {
            return;
        }

        const closingTime = row['Closing Time'] || row['closing_time'] || '';
        const hours = Math.floor(closingTime * 24);
        const minutes = Math.floor((closingTime * 24 - hours) * 60);
        const formattedTime = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;

        const warehouse = {
            eshipz_user_id: eshipzUserId,
            name: row['Warehouse Name'] || row['warehouse_name'] || '',
            pincode: row['Pincode'] || row['warehouse_location_pincode'] || '',
            closing_time: formattedTime,
            non_operational_dates: []
        };

        const datesStr = row['Non-Operational Dates'] || row['non_operational_dates'] || '';
        if (datesStr) {
            const dates = datesStr.split(',').map(date => date.trim()).filter(date => date);
            warehouse.non_operational_dates = dates.map(date => {
                const [dd, mm, yyyy] = date.split('-');
                return new Date(`${yyyy}-${mm}-${dd}T18:30:00.000+00:00`).toISOString();
            });
        }

        result.push(warehouse);
    });

    return {
      result
    };
};

// Convert SLA Excel to JSON
const convertCustomerslaExcelToJson = async (excelData, eshipzUserId) => {
    const result = [];
    const allSlas = [];

    excelData.forEach(row => {
        const sla = {
            slug: row['slug'] || '',
            vendor_id: row['vendor_id'] || '',
            service_type: row['service_type'] || '',
            source_pincode: row['source_pincode'] || '',
            destination_pincode: row['destination_pincode'] || '',
            sla: row['sla'] || 0,
            unit: 'day',
            eshipz_user_id: eshipzUserId,
            is_cod: row['cod_available'] === 'yes' || false
        };

        result.push(sla);
        allSlas.push(sla);
    });

    return {
       allSlas
    };
};

module.exports = {
    convertLocationPriorityExcelFiles,
    convertWarehouseExcelToJson,
    convertCustomerslaExcelToJson
};