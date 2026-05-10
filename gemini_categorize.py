"""
One-shot script: asks Gemini to categorise every item in the Firebase
inventory collection and writes the result back to Firestore.

Usage:
    1. Put your key in .env  →  GEMINI_API_KEY=AIza...
    2. python gemini_categorize.py
"""

import os
import time
from dotenv import load_dotenv
from google import genai
from db_manager import db

load_dotenv()

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
if not GEMINI_API_KEY:
    raise SystemExit("ERROR: GEMINI_API_KEY is empty. Add your key to the .env file and try again.")

client = genai.Client(api_key=GEMINI_API_KEY)

SYSTEM_PROMPT = (
    "You are a strict data classifier for maritime provisioning. "
    "Classify the item into EXACTLY ONE of the following predefined combinations. "
    "You are strictly forbidden from inventing new categories.\n\n"
    "ALLOWED COMBINATIONS:\n"
    "Frozen|Meat & Seafood\n"
    "Frozen|Vegetables\n"
    "Frozen|Bread & Bakery\n"
    "Chilled|Dairy\n"
    "Chilled|Vegetables & Fruit\n"
    "Chilled|Meat & Deli\n"
    "Dry|Oils & Spices\n"
    "Dry|Grains & Baking\n"
    "Dry|Cans & Preserves\n"
    "Dry|Other\n"
    "Softdrink|Beverages\n\n"
    "CRITICAL RULES:\n"
    "- All fresh fruits (bananas, apples, etc.) and fresh greens → Chilled|Vegetables & Fruit\n"
    "- All juices, sodas, water, and beer → Softdrink|Beverages\n"
    "- Frozen fish, shrimp, and meat → Frozen|Meat & Seafood\n"
    "- Chilled sliced meats, sausages, deli items → Chilled|Meat & Deli\n\n"
    "Respond ONLY with the exact MainCategory|SubCategory string from the list above. No other text."
)

VALID_COMBOS = {
    ("Frozen",    "Meat & Seafood"),
    ("Frozen",    "Vegetables"),
    ("Frozen",    "Bread & Bakery"),
    ("Chilled",   "Dairy"),
    ("Chilled",   "Vegetables & Fruit"),
    ("Chilled",   "Meat & Deli"),
    ("Dry",       "Oils & Spices"),
    ("Dry",       "Grains & Baking"),
    ("Dry",       "Cans & Preserves"),
    ("Dry",       "Other"),
    ("Softdrink", "Beverages"),
}


def classify(item_name: str) -> tuple[str, str] | None:
    prompt = f"{SYSTEM_PROMPT}\n\nItem: {item_name}"
    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash", contents=prompt
        )
        text = response.text.strip()
        if "|" in text:
            main, sub = text.split("|", 1)
            main = main.strip()
            sub  = sub.strip()
            if (main, sub) in VALID_COMBOS:
                return main, sub
        print(f"  [WARN] Unexpected response for '{item_name}': {text!r}")
    except Exception as exc:
        print(f"  [ERROR] {item_name}: {exc}")
    return None


def main():
    print("Fetching inventory from Firestore...")
    docs = list(db.collection("inventory").stream())
    docs = [d for d in docs if d.to_dict().get("name")]
    total = len(docs)
    print(f"Found {total} items. Starting Gemini classification...\n")

    updated = skipped = 0

    for idx, doc in enumerate(docs, 1):
        data      = doc.to_dict()
        item_name = data.get("name", doc.id)

        result = classify(item_name)

        if result:
            main_cat, sub_cat = result
            db.collection("inventory").document(doc.id).update({
                "category":     main_cat,
                "sub_category": sub_cat,
            })
            updated += 1
            print(f"[{idx:>3}/{total}]  {main_cat:<10}  {sub_cat:<25}  {item_name}")
        else:
            skipped += 1
            print(f"[{idx:>3}/{total}]  SKIPPED  {item_name}")

        # Stay within Gemini free-tier rate limit (15 req/min)
        time.sleep(4.5)

    print(f"\nDone. Updated: {updated}  |  Skipped/Errors: {skipped}")


if __name__ == "__main__":
    main()
