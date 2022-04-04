
import urllib3
from bs4 import BeautifulSoup
import xlwings as xw

# Search words
search_product = 'tv 40"'

# Excel Conexion.
#
# The .xlsx must be in the same .py directory.
def excel():
    """Excel sheet connection"""
    try:
        wbTest = xw.Book('amazon.xlsx')
        # Excel sheet name
        return wbTest.sheets['products']            
    except:
        print('excel error, something go wrong')
    
def get_url(search_product):
    """Create a URL from search words"""
    template = 'https://www.amazon.com.br/s?k={}&__mk_pt_BR=%C3%85M%C3%85%C5%BD%C3%95%C3%91&ref=nb_sb_noss_2'
    search = search_product.replace(' ','+')
    return template.format(search)

def next_page(page):
    """Create a URL for the next page"""
    return 'https://www.amazon.com.br' + page.a.get('href')
    
def product_info(item):
    """Information about the product"""
    atag = item.h2.a
    description = atag.text.strip()
    url_product = 'https://www.amazon.com.br' + atag.get('href')
   
    try:    
        # Only products with price will be saved
        price_parent = item.find('span','a-price')
        price = price_parent.find('span','a-price-whole').text + price_parent.find('span','a-price-fraction').text
    except AttributeError:
        return
    
    list_result = (description, price, url_product)   
    return list_result

def main(search):
        
    products = []   
    url = get_url(search_product)
    next_url = " "
    
    http = urllib3.PoolManager()
    
    excel_products = excel()
    
    # Maximum 20 pages
    for i in range(1,20):
      
        source = http.request('GET', url)
         
        soup = BeautifulSoup(source.data, "lxml")
        results = soup.find_all('div',{'data-component-type': 's-search-result'})
        page = soup.find('li','a-last') 
        
        print("page {}".format(i))
          
        for item in results:
            product = product_info(item)
            if product \
                and set(product[0].lower().split()[0:2]).intersection(search_product.split()) :
                #and len(set(product[0].lower().split()).intersection(search_product.split())) \ 
                # == len(search_product.split()):     
                
                '''
                If you add the two lines above, it will bring only the products
                that contains the two search words 
                
                '''
                products.append(product)
        
        # if the number of pages is less than 20, it  will break the loop
        try:
            next_url = next_page(page)
            url = next_url
        except:
            break

    # Saving data in excel
    excel_products.range('A1:D1').value = 'description', 'price', 'url_product'              
    excel_products.range('A2').value = products

main(search_product)
