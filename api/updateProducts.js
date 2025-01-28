const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

module.exports = async (req, res) => {
  try {
    // Read the Excel file
    const workbook = xlsx.readFile(path.join(process.cwd(), 'products.xlsx'));
    
    // Read the Products sheet
    const productsSheet = workbook.Sheets['Products'];
    const products = xlsx.utils.sheet_to_json(productsSheet);
    
    // Read the Categories sheet
    const categoriesSheet = workbook.Sheets['Categories'];
    const categories = xlsx.utils.sheet_to_json(categoriesSheet);
    
    // Create the final data structure
    const data = {
      categories: categories,
      products: products
    };
    
    // Write to JSON file
    fs.writeFileSync(path.join(process.cwd(), 'public', 'products.json'), JSON.stringify(data, null, 2));
    
    res.status(200).json({ message: 'Products updated successfully' });
  } catch (error) {
    console.error('Error updating products:', error);
    res.status(500).json({ error: 'Failed to update products' });
  }
};