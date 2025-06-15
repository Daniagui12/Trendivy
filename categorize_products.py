import json
import urllib.parse
import pandas as pd
import re
import requests
from io import BytesIO
from PIL import Image
import os
import time
import shutil
import unicodedata

# Define category keywords for mapping
CATEGORY_KEYWORDS = {
    'MASCOTAS': [
        'mascota', 'perro', 'perros', 'gato', 'gatos', 'alimento', 'juguete', 'accesorio',
        'jaula', 'pecera', 'acuario', 'veterinario', 'peluquería', 'mascotas',
        'collar', 'arnés', 'comida', 'alimento', 'peluqueria', 'veterinaria', 'mascoteria',
        'mascotero', 'mascotera', 'mascoteros', 'mascoteras'
    ],
    'HOGAR': [
        'hogar', 'casa', 'decoración', 'mueble', 'lámpara', 'cortina', 'alfombra',
        'jardín', 'aire acondicionado', 'ventilador', 'limpieza', 'organizador',
        'herramienta', 'electricidad', 'iluminación', 'baño', 'cocina'
    ],
    'SALUD': [
        'salud', 'medicina', 'vitamina', 'suplemento', 'ejercicio', 'deporte',
        'bienestar', 'fitness', 'nutrición', 'cuidado', 'belleza', 'cosmético',
        'higiene', 'personal'
    ],
    'VIAJE': [
        'viaje', 'maleta', 'equipaje', 'mochila', 'bolso', 'accesorio', 'turismo',
        'hotel', 'camping', 'aventura', 'outdoor', 'aire libre'
    ],
    'COCINA': [
        'cocina', 'utensilio', 'electrodoméstico', 'vajilla', 'cubierto', 'olla',
        'sartén', 'batidora', 'licuadora', 'horno', 'microondas', 'refrigerador',
        'alimento', 'bebida', 'comida'
    ]
}

"""
def download_and_resize_image(url, max_size=(150, 150)):
    # Download and resize image to fit in Excel cell.
    try:
        response = requests.get(url)
        if response.status_code == 200:
            img = Image.open(BytesIO(response.content))

            # Calculate new size maintaining aspect ratio
            ratio = min(max_size[0]/img.size[0], max_size[1]/img.size[1])
            new_size = (int(img.size[0]*ratio), int(img.size[1]*ratio))

            img = img.resize(new_size, Image.Resampling.LANCZOS)
            return img, new_size
    except Exception as e:
        print(f"Error processing image {url}: {e}")
    return None, None

def get_image_urls(gallery):
    # Extract and format image URLs from the gallery.
    if not gallery:
        return []

    base_url = "https://d39ru7awumhhs2.cloudfront.net/"
    urls = []

    for image in gallery:
        if image.get('urlS3'):
            # URL encode the path to handle spaces and special characters
            encoded_path = urllib.parse.quote(image['urlS3'])
            urls.append(base_url + encoded_path)

    return urls
"""

def load_products():
    try:
        with open('./data/dropi_products.json', 'r', encoding='utf-8') as f:
            data = json.load(f)
            return data.get('objects', [])
    except Exception as e:
        print(f"Error loading products: {e}")
        return []

def clean_text(text):
    """Clean and normalize text for better matching."""
    # Convert to lowercase
    text = text.lower()
    # Remove special characters and extra spaces
    text = re.sub(r'[^\w\s]', ' ', text)
    # Remove extra spaces
    text = ' '.join(text.split())
    return text

def categorize_product(product):
    # Combine product name and categories for better matching
    text_to_match = clean_text(product['name'])
    if 'categories' in product:
        for category in product['categories']:
            text_to_match += ' ' + clean_text(category['name'])

    # First check for MASCOTAS category with more specific patterns
    mascotas_patterns = [
        r'\b(perro|perros|gato|gatos|mascota|mascotas)\b',
        r'\b(collar|correa|arnés|jaula|pecera|acuario)\b.*\b(perro|perros|gato|gatos|mascota|mascotas)\b',
        r'\b(alimento|comida)\b.*\b(perro|perros|gato|gatos|mascota|mascotas)\b',
        r'\b(mascoteria|mascotero|mascotera|mascoteros|mascoteras)\b'
    ]

    for pattern in mascotas_patterns:
        if re.search(pattern, text_to_match):
            return ['MASCOTAS']

    # Then check other categories in order
    for category, keywords in CATEGORY_KEYWORDS.items():
        if category != 'MASCOTAS':
            if any(keyword in text_to_match for keyword in keywords):
                return [category]

    return ['OTROS']

def get_total_stock(product):
    """Calculate total stock from warehouse_product first, then fall back to variations."""
    # First check warehouse_product array
    if 'warehouse_product' in product and product['warehouse_product']:
        for warehouse in product['warehouse_product']:
            if warehouse.get('stock', 0) > 0:
                return warehouse['stock']

    # If no stock in warehouse_product or it's empty, check variations
    total_stock = 0
    for variation in product.get('variations', []):
        total_stock += variation.get('stock', 0)
    return total_stock

def format_product_name_for_url(name):
    """Format product name for URL: remove accents, convert to lowercase, replace spaces with hyphens."""
    # Remove accents
    name = unicodedata.normalize('NFKD', name).encode('ASCII', 'ignore').decode('ASCII')
    # Convert to lowercase
    name = name.lower()
    # Replace spaces and special characters with hyphens
    name = re.sub(r'[^a-z0-9]+', '-', name)
    # Remove leading and trailing hyphens
    name = name.strip('-')
    return name

def main():
    products = load_products()
    excel_data = []
    miscategorized = []

    """
    # Create a temporary directory for images
    temp_dir = 'temp_images'
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)
    """

    # Process all products first
    for product in products:
        categories = categorize_product(product)
        category = categories[0]

        # Check for potential miscategorization
        if category != 'MASCOTAS' and any(keyword in clean_text(product['name']) for keyword in ['perro', 'perros', 'gato', 'gatos', 'mascota', 'mascotas']):
            miscategorized.append({
                'name': product['name'],
                'current_category': category,
                'suggested_category': 'MASCOTAS'
            })

        # Create product URL
        formatted_name = format_product_name_for_url(product['name'])
        product_url = f"app.dropi.co/dashboard/product-details/{product['id']}/{formatted_name}"

        """
        # Get first image URL
        image_urls = get_image_urls(product.get('gallery', []))
        image_path = None
        image_size = None
        if image_urls:
            image, image_size = download_and_resize_image(image_urls[0])
            if image:
                image_path = os.path.join(temp_dir, f'{product["id"]}.png')
                image.save(image_path)
        """

        excel_data.append({
            'Category': category,
            'ID': product['id'],
            'Name': product['name'],
            'Sale Price': product.get('sale_price', 0),
            'Suggested Price': product.get('suggested_price', 0),
            'Stock': get_total_stock(product),
            'Store Name': product.get('user', {}).get('store_name', 'N/A'),
            'Product URL': product_url
            # 'Image': image_path,
            # 'Image_Size': image_size
        })

    # Create DataFrame
    df = pd.DataFrame(excel_data)

    # Sort by Category and Name
    df = df.sort_values(['Category', 'Name'])

    # Create Excel writer with xlsxwriter engine
    writer = pd.ExcelWriter('./data/categorized_products.xlsx', engine='xlsxwriter')

    # Write DataFrame to Excel
    df.to_excel(writer, sheet_name='Products', index=False)

    # Get workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets['Products']

    # Define formats
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D7E4BC',
        'border': 1
    })

    # Add hyperlink format
    url_format = workbook.add_format({
        'color': 'blue',
        'underline': True
    })

    # Set column widths
    worksheet.set_column('A:A', 15)  # Category
    worksheet.set_column('B:B', 10)  # ID
    worksheet.set_column('C:C', 40)  # Name
    worksheet.set_column('D:E', 15)  # Prices
    worksheet.set_column('F:F', 10)  # Stock
    worksheet.set_column('G:G', 20)  # Store Name
    worksheet.set_column('H:H', 50)  # Product URL

    # Set row height for all rows
    worksheet.set_default_row(30)  # Reduced row height since we don't have images

    # Write header with format
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    # Write data with hyperlinks
    for row_num, row in enumerate(df.itertuples(), start=1):
        for col_num, value in enumerate(row[1:], start=0):
            if col_num == 7:  # Product URL column
                worksheet.write_url(row_num, col_num, f'https://{value}', url_format, value)
            else:
                worksheet.write(row_num, col_num, value)

    """
    # Create a dictionary to map IDs to row numbers
    id_to_row = {row['ID']: idx + 1 for idx, row in df.iterrows()}

    # Insert images in their correct positions
    for product in products:
        if product['id'] in id_to_row:
            row_num = id_to_row[product['id']]
            image_path = os.path.join(temp_dir, f'{product["id"]}.png')

            if os.path.exists(image_path):
                # Calculate cell dimensions
                cell_width = 25 * 7  # Convert column width to pixels
                cell_height = 120    # Row height in pixels

                # Get image size
                img = Image.open(image_path)
                img_width, img_height = img.size

                # Calculate scaling to fit cell while maintaining aspect ratio
                scale = min(cell_width/img_width, cell_height/img_height) * 0.8  # 80% of cell size

                # Insert image with calculated scaling
                worksheet.insert_image(
                    row_num, 7, image_path,
                    {
                        'x_scale': scale,
                        'y_scale': scale,
                        'x_offset': 5,
                        'y_offset': 5,
                        'positioning': 1  # Move and size with cells
                    }
                )
    """

    # Save the Excel file
    writer.close()

    # Print statistics
    print("\nProduct Categorization Statistics:")
    print("-" * 30)
    stats = df['Category'].value_counts()
    for category, count in stats.items():
        print(f"{category}: {count} products")

    # Print potential miscategorized items
    if miscategorized:
        print("\nPotential Miscategorized Items:")
        print("-" * 30)
        for item in miscategorized:
            print(f"Product: {item['name']}")
            print(f"Current Category: {item['current_category']}")
            print(f"Suggested Category: {item['suggested_category']}")
            print("-" * 20)

    print("\nExcel file 'categorized_products.xlsx' has been created successfully!")

    """
    # Clean up temporary images with retry mechanism
    print("\nCleaning up temporary files...")
    max_retries = 3
    retry_delay = 2  # seconds

    for attempt in range(max_retries):
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            print("Temporary files cleaned up successfully!")
            break
        except PermissionError:
            if attempt < max_retries - 1:
                print(f"Waiting for files to be released... (Attempt {attempt + 1}/{max_retries})")
                time.sleep(retry_delay)
            else:
                print("Warning: Could not clean up temporary files. You may need to delete the 'temp_images' folder manually.")
    """

if __name__ == "__main__":
    main()