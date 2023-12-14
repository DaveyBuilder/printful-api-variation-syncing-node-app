require('dotenv').config();
const path = require("path");
const express = require('express');
const fetch = require("node-fetch");
const hbs = require("hbs");
const bodyParser = require('body-parser');
const multer = require('multer');
const fs = require('fs'); 
const ExcelJS = require('exceljs');

////////////Product imports
const gildan_18000_sweatshirt = require('./gildan_18000_sweatshirt');
const gildan_64000_tee = require('./gildan_64000_tee');
const gildan_18500_hoodie = require('./gildan_18500_hoodie');

// Create express application & set handlebars as view engine
const app = express();
const { exec } = require('child_process');
app.set("view engine", "hbs");
app.set("views", path.join(__dirname, 'views'));
app.use(express.urlencoded({ extended: true }));

//defining the storage for multer
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
      cb(null, 'uploads/')
    },
    filename: function (req, file, cb) {
      cb(null, file.originalname);
    }
});
const upload = multer({ storage: storage });

const printfulApiToken = process.env.PRINTFUL_API_KEY;

// Initial render of `welcome.hbs` file.
app.get('/', async (req, res) => {
    res.render("welcome");
});


app.get('/createTemplate', async (req, res) => {

    //Get all parent listings that have zero variations already synced
    let parentListingsZeroSyncs = [];
    const limit = 100;
    let offset = 0;
    let stayInLoop = true;
    let listingCount;
    
    while (stayInLoop) {
        const requestOptions = {
            method: 'GET',
            headers: {
                Authorization: `Bearer ${printfulApiToken}`,
                'Content-Type': 'application/json'
            },
        };
        
        const listingsResponse = await fetch(
            `https://api.printful.com/sync/products?offset=${offset}&limit=${limit}`,
            requestOptions
            );
        
        let listingsData;
        if (listingsResponse.ok) {
            listingsData = await listingsResponse.json();
            listingsData.result.forEach(result => {
                if (result.synced === 0) {
                //if (result.name.includes("AUTOMATION TEST LISTING")) {
                    // Defining the listing object
                    let obj = {
                        id: result.id,
                        name: result.name
                    };
                    parentListingsZeroSyncs.push(obj);
                //}
                }
            });
        } else {
            const errorData = await listingsResponse.json();
            console.log('Error:', errorData);
        }
        listingCount = listingsData.paging.total;
        offset = offset + 100;
        if (offset >= listingCount) {
            stayInLoop = false;
        };
    }

    console.log("Total number of listings found with zero synced variations: " + parentListingsZeroSyncs.length)
    
    //SO NOW WE HAVE AN ARRAY OF ALL THE PARENT LISTINGS THAT HAVE ZERO SYNCED VARIATIONS
    //NOW WE NEED TO GET THE VARIATIONS FOR EACH OF THOSE PARENT LISTINGS

    console.log("Now getting the variation IDs for each of the listings...")

    let allVariations = [];
    let counter = 0;

    for (let listing of parentListingsZeroSyncs) {
        const id = listing.id;
        const requestOptionsVariations = {
            method: 'GET',
            headers: {
                Authorization: `Bearer ${printfulApiToken}`,
                'Content-Type': 'application/json'
            },
        };
    
        const responseVariations = await fetch(`https://api.printful.com/sync/products/${id}`, requestOptionsVariations);
        
        let tempVariations = [];
        if (responseVariations.ok) {
            const variationsData = await responseVariations.json();

            let teeArray = [];
            let sweatshirtArray = [];
            let hoodieArray = [];
            variationsData.result.sync_variants.forEach(variant => {
                if (variant.name.includes("Unisex TEE")) {
                    let size_color = variant.name.split("Unisex TEE")[1].trim();
                    let size = size_color.split('/')[0].replace('| ', '').trim();
                    let color = size_color.split('/')[1].trim();
                    let obj = {
                        id: variant.id,
                        name: variant.name,
                        type: "Tee",
                        size,
                        color
                    };
                    teeArray.push(obj);
                } else if (variant.name.includes("Unisex SWEATSHIRT")) {
                    let size_color = variant.name.split("Unisex SWEATSHIRT")[1].trim();
                    let size = size_color.split('/')[0].replace('| ', '').trim();
                    let color = size_color.split('/')[1].trim();
                    let obj = {
                        id: variant.id,
                        name: variant.name,
                        type: "Sweatshirt",
                        size,
                        color
                    };
                    sweatshirtArray.push(obj);
                } else if (variant.name.includes("Unisex HOODIE")) {
                    let size_color = variant.name.split("Unisex HOODIE")[1].trim();
                    let size = size_color.split('/')[0].replace('| ', '').trim();
                    let color = size_color.split('/')[1].trim();
                    let obj = {
                        id: variant.id,
                        name: variant.name,
                        type: "Hoodie",
                        size,
                        color
                    };
                    hoodieArray.push(obj);
                } else {
                    //console.log("This variant is not a tee, sweatshirt, or hoodie: " + variant.name);
                }
            });
            if (teeArray.length > 0) {
                let groupedByColor = teeArray.reduce((acc, item) => {
                    // If the color group doesn't exist, create it
                    if (!acc[item.color]) {
                        acc[item.color] = {
                            name: item.name.split("Unisex TEE")[0].trim(),
                            type: item.type,
                            color: item.color
                        };
                    }
                    // Add the size as a property and set its value to the item's id
                    acc[item.color][item.size] = item.id;
                    return acc;
                }, {});
                tempVariations.push(groupedByColor);
            }
            if (sweatshirtArray.length > 0) {
                let groupedByColor = sweatshirtArray.reduce((acc, item) => {
                    if (!acc[item.color]) {
                        acc[item.color] = {
                            name: item.name.split("Unisex SWEATSHIRT")[0].trim(),
                            type: item.type,
                            color: item.color
                        };
                    }
                    acc[item.color][item.size] = item.id;
                    return acc;
                }, {});
                tempVariations.push(groupedByColor);
            }
            if (hoodieArray.length > 0) {
                let groupedByColor = hoodieArray.reduce((acc, item) => {
                    if (!acc[item.color]) {
                        acc[item.color] = {
                            name: item.name.split("Unisex HOODIE")[0].trim(),
                            type: item.type,
                            color: item.color
                        };
                    }
                    acc[item.color][item.size] = item.id;
                    return acc;
                }, {});
                tempVariations.push(groupedByColor);
            }

            allVariations.push(tempVariations);

        } else {
            const errorData = await responseVariations.json();
            console.log('Error:', errorData);
        }

        counter++;
        if (counter >= 15) {
            break;
        }
    }

    //Now create a new array with the final objects containing all required info incl. Printful product IDs
    console.log("Finished getting the variation IDs for each of the listings")

    let printfulProducts = {
        "Sweatshirt": gildan_18000_sweatshirt,
        "Tee": gildan_64000_tee,
        "Hoodie": gildan_18500_hoodie
    };

    let finalResults = [];

    allVariations.forEach(variations => {
        let tempArray = [];
        variations.forEach(variant => {
            for (let color in variant) {
                let productType = variant[color].type;
                let sizes = Object.keys(variant[color]).filter(key => ["name", "type", "color"].indexOf(key) === -1);
                let matchFound = false; // Add a flag to check if a match was found
                let smallestSize, largestSize;
                let sizeRanking = ['S', 'M', 'L', 'XL', '2XL', '3XL', '4XL', '5XL']; // Add a size ranking array
                let sizeIdPairs = []; // Add an array to store the size/id pairs
                sizes.forEach(size => {
                    if (printfulProducts[productType][color] && printfulProducts[productType][color][size]) {
                        let sizeIdPair = `${variant[color][size]}:${printfulProducts[productType][color][size]}`;
                        sizeIdPairs.push(sizeIdPair); // Add the size/id pair to the array
                        variant[color][size] = `${variant[color][size]}:${printfulProducts[productType][color][size]}`;
                        matchFound = true; // Set the flag to true if a match is found
                        // Update smallest and largest sizes
                        if (!smallestSize || sizeRanking.indexOf(size) < sizeRanking.indexOf(smallestSize)) smallestSize = size;
                        if (!largestSize || sizeRanking.indexOf(size) > sizeRanking.indexOf(largestSize)) largestSize = size;
                    }
                });
                // If a match was found, create a new object and push it into the finalResults array
                if (matchFound) {
                    let newVariant = {
                        name: variant[color].name,
                        type: variant[color].type,
                        color: variant[color].color,
                        sizes: `${smallestSize}-${largestSize}`,
                        sizeIDPairs: sizeIdPairs.join('|')
                    };
                    tempArray.push(newVariant);
                }
            }
        });
        tempArray.sort((a, b) => a.color.localeCompare(b.color));
        finalResults.push(tempArray);
    });

    finalResults = finalResults.flat();

    console.log("Now creating the Excel file...")

    let fileCounter = 0;
    let newFileName;
    let fileExists = true;

    while (fileExists) {
        newFileName = `./zero_synced_listings${fileCounter ? fileCounter : ''}.xlsx`;
        fileExists = fs.existsSync(newFileName);
        if (fileExists) fileCounter++;
    }

    let workbook = new ExcelJS.Workbook();
    let worksheet = workbook.addWorksheet('Sheet1');

    // Define columns with widths
    worksheet.columns = [
        { header: 'Title', key: 'name', width: 132 },
        { header: 'Type', key: 'type', width: 11 },
        { header: 'Sizes', key: 'sizes', width: 10 },
        { header: 'Color', key: 'color', width: 25 },
        { header: 'File ID', key: 'fileID', width: 15 },
        { header: 'IDs', key: 'sizeIDPairs', width: 10 }
    ];

    // Center align the headers
    let headerRow = worksheet.getRow(1);
    headerRow.eachCell(cell => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { bold: true };
    });

    // Add rows
    finalResults.forEach(result => {
        let row = worksheet.addRow(result);
        row.eachCell((cell, colNumber) => {
            if (worksheet.columns[colNumber - 1].key === 'sizeIDPairs') {
                cell.alignment = { vertical: 'middle', horizontal: 'left' };
            } else {
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
            }
        });
    });

    let templateCreated;
    let templateFailed;
    // Write to file
    try {
        await workbook.xlsx.writeFile(`${newFileName}`);
        console.log('The Excel file was written successfully');
        templateCreated = true;
    } catch (err) {
        console.error('Error writing Excel file:', err);
        templateFailed = true;
    }

    res.render('welcome', {
        templateCreated,
        templateFailed,
        fileName: newFileName
    });
    
});

app.put("/syncPrintful", upload.single('xlsxfile'), async (req, res) => {

    const headers = {
        'Authorization': `Bearer ${printfulApiToken}`,
        'Content-Type': 'application/json'
    };

    let printfulErrorsArray = [];
    let printfulWarningsArray = [];

    const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

    console.log("STARTING PRINTFUL SYNCING PROCESS...")

    const readRowsArray = [];
    const csvErrorArray = [];

    let keyMapping = {
        'Title': 'name',
        'Color': 'color',
        'Type': 'type',
        'Sizes': 'sizes',
        'File ID': 'fileID',
        'IDs': 'sizeIDPairs',
        'Synced On Printful': 'syncedOnPrintful'
    };

    let uploadedFile = req.file.path;
    let workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile(uploadedFile);
        let worksheet = workbook.getWorksheet('Sheet1');
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            // Skip the header row
            if (rowNumber !== 1) {
                // Initialize rowObj with all keys
                let rowObj = {
                    name: null,
                    color: null,
                    type: null,
                    sizes: null,
                    fileID: null,
                    sizeIDPairs: null,
                    syncedOnPrintful: null
                };
                row.eachCell((cell, colNumber) => {
                    // The header row contains the keys for the objects (after usiing keyMapping)
                    let header = worksheet.getRow(1).getCell(colNumber).value;
                    let key = keyMapping[header];
                    rowObj[key] = cell.value;
                });
                //push an error message to the csvErrorArray if the File ID column is empty
                !rowObj['sizeIDPairs'] && csvErrorArray.push(`Row ${rowNumber}: Missing Printful IDs`);
            
                readRowsArray.push(rowObj);
            }
        });
    } catch (err) {
        console.error('Error reading Excel file:', err);
        csvErrorArray.push(`Error reading Excel file: ${err}`);
    }

    if (csvErrorArray.length > 0) {
        res.json({ csvErrorArray });
        return;
    }

    function getFormattedDate() {
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');
        const hours = String(now.getHours()).padStart(2, '0');
        const minutes = String(now.getMinutes()).padStart(2, '0');
        const seconds = String(now.getSeconds()).padStart(2, '0');
    
        return `${year}-${month}-${day}_${hours}-${minutes}-${seconds}`;
    }


    
    const processRow = async (currentObject) => {

        let idArray = currentObject['sizeIDPairs'].split('|');
        let fileID = currentObject['fileID'];

        try {
            for (let i = 0; i < idArray.length; i++) {
                await delay(7000);

                let variantID = idArray[i].split(':')[0];
                let printfulCatalogID = idArray[i].split(':')[1];
    
                const url = `https://api.printful.com/sync/variant/${variantID}`;
                const body = {
                    variant_id: printfulCatalogID,
                    files: [{
                        id: fileID
                    }],
                    options: [{
                        id: "embroidery_type",
                        value: "flat"
                    }]
                };
                const response = await fetch(url, {
                    method: 'PUT',
                    headers: headers,
                    body: JSON.stringify(body)
                });
    
                if (response.ok) {
                    await response.json();
                } else {
                    const errorData = await response.text();
                    if (response.status == 502) {
                        throw new Error(`Problem with Printful server. Re-run the automation in a few minutes time using the backup file and the app will attempt to re-link this variation.`);
                    } else {
                        throw new Error(`HTTP Status: ${response.status} - ${errorData}`);
                    }
                }
            }
            
            // if all operations complete successfully modify the syncedOnPrintful value
            currentObject['syncedOnPrintful'] = "Done";

        } catch (error) {

            console.error(`${currentObject['name']}, ${currentObject['color']}. Error linking the variations: ${error.message}`);
            printfulErrorsArray.push(`--- ${currentObject['name']}, ${currentObject['color']}. Error linking the variations: ${error.message}`);

            // if an error occurs modify the syncedOnPrintful value
            currentObject['syncedOnPrintful'] = "Error";
        }
        
    };

    const dateForFilename = getFormattedDate();
    const backupFileName = "printfulSyncBackup_" + dateForFilename + ".xlsx";
    const backupDirectory = './backups/';
    const backupFilePath = backupDirectory + backupFileName;

    for (let i = 0; i < readRowsArray.length; i++) {
        if (readRowsArray[i]['syncedOnPrintful'] != "Done"  && readRowsArray[i]['fileID']) {
            console.log(`Row ${i + 2} of ${readRowsArray.length + 1}...`)

            await processRow(readRowsArray[i]);

            //Now write the rowsArrayAfterEtsyDraftCreation to a new .xlsx file
            let workbook = new ExcelJS.Workbook();
            let worksheet = workbook.addWorksheet('Sheet1');

            // Define columns with widths
            worksheet.columns = [
                { header: 'Title', key: 'name', width: 132 },
                { header: 'Color', key: 'color', width: 11 },
                { header: 'Type', key: 'type', width: 10 },
                { header: 'Sizes', key: 'sizes', width: 25 },
                { header: 'File ID', key: 'fileID', width: 15 },
                { header: 'IDs', key: 'sizeIDPairs', width: 10 },
                { header: 'Synced On Printful', key: 'syncedOnPrintful', width: 16 }
            ];

            // Center align the headers
            let headerRow = worksheet.getRow(1);
            headerRow.eachCell(cell => {
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
                cell.font = { bold: true };
            });

            // Add rows (the data)
            readRowsArray.forEach(result => {
                let row = worksheet.addRow(result);
                row.eachCell((cell, colNumber) => {
                    if (worksheet.columns[colNumber - 1].key === 'sizeIDPairs') {
                        cell.alignment = { vertical: 'middle', horizontal: 'left' };
                    } else {
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    }
                });
            });

            // Write to file
            try {
                await workbook.xlsx.writeFile(backupFilePath);
                i === 0 ? console.log(`CREATED BACKUP FILE AT ${backupFilePath}`) : console.log(`UPDATED ${backupFilePath}`);
            } catch (err) {
                console.error(`Failed to update ${backupFilePath}`, err);
            }

        } else if (!readRowsArray[i]['fileID']) {           
            //push a warning message if the File ID column is empty
            printfulWarningsArray.push(`Row ${i + 2}: File ID column was empty so listing was not synced`);
        }
    }

    console.log("Finished linking the variations to Printful.")

    let allRowsSkipped = false;
    if (printfulWarningsArray.length === readRowsArray.length) {
        allRowsSkipped = true;
    }

    res.json({
        printfulWarningsArray,
        printfulErrorsArray,
        allRowsSkipped,
    });
    
});


const port = 3003;
app.listen(port, () => {
    console.log(`Hi! Go to the following link in your browser to start the app: http://localhost:${port}`);
    exec(`start http://localhost:${port}`, (err, stdout, stderr) => {
        if (err) {
            console.error(`exec error: ${err}`);
            return;
        }
    });
});