import urllib2, openpyxl
from bs4 import BeautifulSoup as bs4
from openpyxl import Workbook
from openpyxl.cell import get_column_letter as col




url = "yoururl.com"

# Excel workbook
workbook = Workbook()
sheet = workbook.active

# Main product categories {Category Name: URL}
product_categories = {}

# Product URLs {Product Name: Product URL}
products = {}

# Product Details {Product URL: {name:product name,description:product description, images:[{url:product_url,alt:alt}], variants: {variant_name:[variant_name,variant_price]}}
product_meta = {}

def get_soup(url):
    """ Returns beautiful soup object of given url.

    get_soup(str) -> object(?)
    """

    req = urllib2.Request(url)
    response = urllib2.urlopen(req)
    html = response.read()
    soup = bs4(html)

    return soup

def get_products_from_category(category_url):

    try:
        product = get_soup(category_url)

        # Check to see if it is a sub-category
        sub_cat = True
        for n in product.find_all('p'):
            if n.has_attr("class"):
                if "top-right" in n['class']:
                    # page is not a sub-category
                    for n in product.find_all('h3'):
                        if n.getText() not in products:
                            products.update({n.getText():url+n.a.get('href')})
                        sub_cat = False

        if sub_cat == True:
            for n in product.findAll('div'):
                if n.has_attr("class"):
                    if 'thumbnail' in n['class']:
                        # Add it as a product category
                        product_categories.update({'url':url+n.a.get('href')})
                        # Remove it as a product
                        products.pop(category_url,None) # Remove as


    except Exception, e: print(str(e), category_url)


def get_product_info(product_url):
    product = get_soup(product_url)
    product_name = product.h1.getText()
    description = ""

    # Description
    for n in product.findAll('p'):
        if n.has_attr("class"):
            description= n.getText()
    try:
        price = product.find(id="display_price")['value'][1::]
    except TypeError:
        print(product_url)
        price = 0

    # images
    images = []
    for n in product.findAll('div'):
        if n.has_attr("class"):
            if 'thumbnail' in n['class']:
                try:
                    images.append({'url':url+n.a['data-image'],'alt':n.img['alt']})
                except KeyError, e: print(str(e), product_url)
    # variants
    variants = {}
    variant_menu = product.findAll('option')
    for n in variant_menu[1::]:     # First variant is a description ie "colour"

        variant_name = n.getText()
        try:
            variant_price = n['price']

            if variant_price == "":
                variant_price = price
        except KeyError: variant_price = price

        variants.update({variant_name:[variant_name,variant_price]})


    product_meta.update({product_url:{
    'name':product_name,
    'description':description,
    'price':price,
    'images':images,
    'variants': variants}})


# Get categories from side menu
home = get_soup(url+'/products')
for n in home.find_all('li'):
    if ' list-group-item' in str(n):
        product_categories.update({n.a.getText():url+n.a.get('href')})

# Get products from each category and check for sub-categories
# Some pages use thumbnails to link to subcategories instead of product collections
# Convert to list because the dictionary is going to increase when we find a sub-category
for n in list(product_categories):
    get_products_from_category(product_categories[n])

# Run it again to get products from sub-category pages
for n in product_categories:
    get_products_from_category(product_categories[n])

print(products
)
# Get product details for each product
row_count = 1
for n in products:
    col_index = 1
    get_product_info(products[n])
    p = product_meta[products[n]]
    print(p)
    sheet[col(col_index) + str(row_count)] = p['name']
    col_index += 1

    sheet[col(col_index) + str(row_count)] = p['description']
    col_index += 1

    for image in p['images']:
        sheet[col(col_index) + str(row_count)] = image['url']
        col_index += 1
        sheet[col(col_index) + str(row_count)] = image['alt']
        col_index += 1

    for variant in p['variants']:
        sheet[col(col_index) + str(row_count)] = p['variants'][variant][0]
        col_index += 1
        sheet[col(col_index) + str(row_count)] = p['variants'][variant][1]

    row_count += 1

    workbook.save('/Users/Joshua/Documents/My Documents/Websites/Latonas/products.xlsx')
