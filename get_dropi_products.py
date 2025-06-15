import requests
import json
from datetime import datetime

def get_favorite_products():
    url = "https://api.dropi.co/api/products/v4/index"
    
    headers = {
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "en-US,en;q=0.9",
        "Connection": "keep-alive",
        "Content-Type": "application/json",
        "Origin": "https://app.dropi.co",
        "Referer": "https://app.dropi.co/",
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-site",
        "User-Agent": "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/137.0.0.0 Mobile Safari/537.36 Edg/137.0.0.0",
        "X-Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwOlwvXC9hcHAuZHJvcGkuY286ODAiLCJpYXQiOjE3NTAwMDY4NzgsImV4cCI6MTc1MDAyMTI3OCwibmJmIjoxNzUwMDA2ODc4LCJqdGkiOiJHWHplMmlya01BZVF2TXNsIiwic3ViIjo1MzE1OTEsInBydiI6Ijg3ZTBhZjFlZjlmZDE1ODEyZmRlYzk3MTUzYTE0ZTBiMDQ3NTQ2YWEiLCJhdWQiOiJEUk9QSSIsInRva2VuX3R5cGUiOiJEUk9QSSIsIndiX2lkIjoxfQ.NwtaK-BzNiBIyfAxNByrQ--GSeJyB-q4YY6aZ0f3pWI",
        "sec-ch-ua": "\"Microsoft Edge\";v=\"137\", \"Chromium\";v=\"137\", \"Not/A)Brand\";v=\"24\"",
        "sec-ch-ua-mobile": "?1",
        "sec-ch-ua-platform": "\"Android\""
    }
    
    payload = {
        "pageSize": 60,
        "startData": 0,
        "privated_product": False,
        "userVerified": False,
        "order_by": "created_at",
        "order_type": "desc",
        "favorite": True,
        "with_collection": True,
        "get_stock": False,
        "no_count": True
    }
    
    try:
        response = requests.post(url, headers=headers, json=payload)
        response.raise_for_status()  # Raise an exception for bad status codes
        
        # Save the response to a JSON file with timestamp
        filename = "./data/dropi_products.json"
        
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(response.json(), f, indent=2, ensure_ascii=False)
            
        print(f"Successfully retrieved products and saved to {filename}")
        return response.json()
        
    except requests.exceptions.RequestException as e:
        print(f"Error making request: {e}")
        return None

if __name__ == "__main__":
    get_favorite_products() 