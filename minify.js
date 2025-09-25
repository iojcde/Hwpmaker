import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import minifyXML from 'minify-xml';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Directory containing XML files to minify
const contentsDir = path.join(__dirname, 'file', 'Contents');

// Minification options that collapse newlines in strings
const minifyOptions = {
    removeComments: true,
    removeWhitespaceBetweenTags: true,
    considerPreserveWhitespace: true,
    collapseWhitespaceInTags: true,
    collapseEmptyElements: true,
    trimWhitespaceFromTexts: false,
    collapseWhitespaceInTexts: true, // This helps collapse newlines in text content
    collapseWhitespaceInProlog: true,
    collapseWhitespaceInDocType: true,
    removeSchemaLocationAttributes: false,
    removeUnnecessaryStandaloneDeclaration: true,
    removeUnusedNamespaces: true,
    removeUnusedDefaultNamespace: true,
    shortenNamespaces: true,
    ignoreCData: true
};

async function minifyXMLFiles() {
    try {
        // Check if the Contents directory exists
        if (!fs.existsSync(contentsDir)) {
            console.error(`Directory ${contentsDir} does not exist`);
            return;
        }

        // Read all files in the Contents directory
        const files = fs.readdirSync(contentsDir);
        
        // Filter for XML files
        const xmlFiles = files.filter(file => file.endsWith('.xml'));
        
        if (xmlFiles.length === 0) {
            console.log('No XML files found in the Contents directory');
            return;
        }

        console.log(`Found ${xmlFiles.length} XML files to minify:`);
        xmlFiles.forEach(file => console.log(`  - ${file}`));

        // Process each XML file
        for (const xmlFile of xmlFiles) {
            const filePath = path.join(contentsDir, xmlFile);
            
            try {
                console.log(`\nMinifying ${xmlFile}...`);
                
                // Read the XML file
                const xmlContent = fs.readFileSync(filePath, 'utf8');
                
                // Get file stats before minification
                const originalSize = Buffer.byteLength(xmlContent, 'utf8');
                
                // Minify the XML
                const minifiedXML = minifyXML(xmlContent, minifyOptions);
                
                // Write the minified XML back to the file
                fs.writeFileSync(filePath, minifiedXML, 'utf8');
                
                // Get file stats after minification
                const minifiedSize = Buffer.byteLength(minifiedXML, 'utf8');
                const savings = originalSize - minifiedSize;
                const savingsPercent = ((savings / originalSize) * 100).toFixed(1);
                
                console.log(`  ✓ Original size: ${originalSize} bytes`);
                console.log(`  ✓ Minified size: ${minifiedSize} bytes`);
                console.log(`  ✓ Saved: ${savings} bytes (${savingsPercent}%)`);
                
            } catch (error) {
                console.error(`  ✗ Error minifying ${xmlFile}:`, error.message);
            }
        }
        
        console.log('\nMinification complete!');
        
    } catch (error) {
        console.error('Error during minification process:', error);
    }
}

// Run the minification
minifyXMLFiles();
