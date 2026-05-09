import re
import pdfplumber
from db_manager import db

PDF_FILE = 'Invoice_esb-0085942_1772431065305  Proviant_ESVAGT DANA_esb-0101775{A2325E43-A13E-4D56-8B57-94173E6B8351}.PDF'

# Matches the start of an invoice item line:
# LINE_NUM  [ISSA_CODE]  ORDER_QTY  INV_QTY+UNIT  rest...
ITEM_RE = re.compile(
    r'^(\d+\.\d+)\s+'       # line number  e.g. 40.00
    r'(?:\d{5,8}\s+)?'      # optional ISSA/IMPA code
    r'\d+\.\d+\s+'          # order quantity
    r'(\d+\.\d+\s*[A-Za-z]+)\s+'  # invoice qty + unit  e.g. 5.00KG or 5.00 KG
    r'(.+)$'                # description + prices
)

# Lines to ignore when looking for wrapped description text
SKIP_RE = re.compile(
    r'^(ESVAGT|Invoice\s|Case\s+ID|Customer|DOKVEJ|DK-\d|DENMARK|VAT:|'
    r'Date:|Page:|Order\s+Invoice|Line\s+ISSA|Due\s+date|Payment\s+terms|'
    r'IMO\s+\d|Account\s+No|Our\s+ref\.|Delivery|and/or|'
    r'JJeennss|DDaarruummvveejj|DDaanniisshh\s+SSuuppppllyy|Subtotal|'
    r'Please\s+ensure|TRADE\s+SANCTIONS|TERMS\s+AND\s+CONDITIONS)',
    re.IGNORECASE
)


def sub_categorize(description):
    d = description.upper()
    if any(k in d for k in ('BEEF', 'PORK', 'CHICKEN', 'LAMB', 'TURKEY', 'DUCK',
                             'SAUSAGE', 'BACON', 'SALAMI', 'PEPPERONI', 'HAM',
                             'ROULADE', 'MEATBALL', 'MINCED', 'KEBAB', 'HEN',
                             'GOULASH', 'TENDERLOIN', 'SALT MEAT', 'NECK PULLED')):
        return 'Meat & Poultry'
    if any(k in d for k in ('COD', 'SALMON', 'PLAICE', 'FISH', 'SHRIMP',
                             'HERRING', 'TUNA', 'MACKEREL', 'SEJ')):
        return 'Fish & Seafood'
    if any(k in d for k in ('MILK', 'CHEESE', 'BUTTER', 'CREAM', 'YOGHURT',
                             'SKYR', 'MARGARINE', 'EGG', 'BUTTERMILK',
                             'PHILADELPHIA', 'PARMESAN', 'CHEDDAR', 'FETA',
                             'MASCARPONE', 'SAMSOE', 'MOZZARELLA')):
        return 'Dairy'
    if any(k in d for k in ('APPLE', 'ASPARAGUS', 'AVOCADO', 'BANANA', 'BASIL',
                             'BEETROOT', 'BELL PEPPER', 'BROCCOLI', 'CABBAGE',
                             'CARROT', 'CAULIFLOWER', 'CHILI', 'CHIVES', 'CORN',
                             'CRESS', 'CUCUMBER', 'DILL', 'EGGPLANT', 'FALAFEL',
                             'GARLIC', 'GINGER', 'GRAPE', 'KIWI', 'LEMON',
                             'LETTUCE', 'MELON', 'MUSHROOM', 'ONION', 'ORANGE',
                             'PARSLEY', 'PEAR', 'PEAS', 'PINEAPPLE', 'PLUM',
                             'POTATO', 'POTHERBS', 'RADISH', 'ROSEMARY', 'SQUASH',
                             'TANGERINE', 'THYME', 'TOMATO', 'WATERMELON',
                             'VEGETABLE', 'VEGETARIAN', 'BEANS GREEN', 'LUCERNE',
                             'PEPPER SPANISH', 'STEWED FRUIT', 'FRENCH FRIES', 'POMMES')):
        return 'Vegetables & Fruit'
    if any(k in d for k in ('BREAD', 'BUN', 'BUNS', 'CROISSANT', 'PASTRY',
                             'PANCAKE', 'ROLL W/', 'TORTILLA', 'WAFER',
                             'BISCUIT', 'COOKIE', 'CAKE', 'MACAROON')):
        return 'Bread & Bakery'
    if any(k in d for k in ('JUICE', 'COFFEE', 'TEA', 'OAT DRINK')):
        return 'Beverages'
    if any(k in d for k in ('OIL', 'SPICE', 'CURRY', 'CINNAMON', 'MUSTARD',
                             'KETCHUP', 'DRESSING', 'REMOULADE', 'MAYONNAISE',
                             'SAUCE', 'PESTO', 'HONEY', 'STOCK', 'SOUP',
                             'ESSENCE', 'TOPPING', 'ASPIC', 'PEPPER BLACK',
                             'COOKING SPRAY', 'SVAMPEFOND', 'COLD SALAD')):
        return 'Sauces, Spices & Oils'
    if any(k in d for k in ('FLOUR', 'RICE', 'SUGAR', 'CEREAL', 'OATMEAL',
                             'MUESLI', 'DATES', 'NUTS', 'SNACK', 'CHIPS',
                             'SPREAD', 'PEANUT', 'CHOCOLATE', 'CANNED',
                             'CUP NOODLE', 'YEAST', 'ICE CREAM', 'BEANS RED',
                             'BAKED BEANS', 'SALT STICK')):
        return 'Dry Goods & Cans'
    return 'Other'


def categorize(description):
    d = description.upper()
    if any(k in d for k in ('FROZEN', 'ICE CREAM', 'IQF')):
        return 'Frozen'
    if any(k in d for k in ('FRESH', 'CHILLED', 'MILK', 'YOGHURT', 'CHEESE',
                             'CREAM', 'BUTTER', 'MARGARINE', 'EGG', 'COLD SALAD', 'DELI')):
        return 'Chilled'
    if any(k in d for k in ('SODA', 'WATER', 'BEER', 'COLA', 'PEPSI', 'FAXE', 'SOFTDRINK')):
        return 'Softdrink'
    return 'Dry'


def extract_desc_and_amount(text):
    """Return (description, unit_price). Unit price is the second-to-last decimal;
    both trailing price columns are stripped from the description."""
    decimals = list(re.finditer(r'\d+\.\d+', text))
    if len(decimals) < 2:
        return text.strip(), None

    second_last = decimals[-2]
    return text[:second_last.start()].strip(), float(second_last.group())


def parse_invoice(pdf_path):
    items = []
    current = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            for line in text.splitlines():
                line = line.rstrip()
                if not line.strip():
                    continue

                # Stop at invoice summary section
                if re.match(r'^Subtotal', line.strip(), re.IGNORECASE):
                    break

                m = ITEM_RE.match(line)
                if m:
                    if current:
                        items.append(current)
                    desc, amount = extract_desc_and_amount(m.group(3))
                    unit = re.sub(r'[\d\s.]', '', m.group(2)).upper()
                    current = {'desc': desc, 'amount': amount, 'unit': unit,
                               'sub_category': sub_categorize(desc)}
                elif current and not SKIP_RE.match(line.strip()) and '||' not in line:
                    # Wrapped description text — append to current item
                    current['desc'] += ' ' + line.strip()

    if current:
        items.append(current)

    return items


def upload_items_to_firebase(items):
    total = len(items)
    for i, item in enumerate(items, 1):
        doc_id = item['desc'].replace('/', ' ')
        doc = {
            'name': item['desc'],
            'category': categorize(item['desc']),
            'sub_category': item.get('sub_category', 'Other'),
            'price': item['amount'] if item['amount'] is not None else 0.0,
            'unit': item.get('unit', ''),
            'stock': 0,
        }
        db.collection('inventory').document(doc_id).set(doc)
        if i % 50 == 0:
            print(f"Uploaded {i}/{total} items...")
    print(f"Done! All {total} items uploaded to Firebase.")


if __name__ == '__main__':
    print(f"Parsing: {PDF_FILE}\n")
    items = parse_invoice(PDF_FILE)
    print(f"Found {len(items)} line items. Starting upload...\n")
    upload_items_to_firebase(items)
