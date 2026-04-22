"""
Mazrica API クライアント
案件一覧・商品内訳情報を取得するためのクライアント
"""
import time
import requests
from typing import Optional, Generator
from dataclasses import dataclass

from .config import Config


@dataclass
class ProductDetail:
    """商品内訳情報"""
    product_id: Optional[int]
    product_name: str
    quantity: Optional[float]
    unit_price: Optional[float]
    amount: Optional[float]
    custom_fields: dict


@dataclass
class Deal:
    """案件情報"""
    id: int
    name: str
    customer_name: Optional[str]
    customer_id: Optional[int]
    deal_type_id: Optional[int]
    deal_type_name: Optional[str]
    phase_name: Optional[str]
    amount: Optional[float]
    expected_contract_date: Optional[str]
    created_at: str
    updated_at: str
    user_name: Optional[str]
    product_name: Optional[str]  # 商品名（product.name）
    product_id: Optional[int]    # 商品ID（product.id）
    product_details: list[ProductDetail]  # 商品内訳詳細（dealProductDetails）
    custom_fields: dict
    customs_by_id: dict  # dealCustomItemId → 値（文字列化済み）


class MazricaAPIError(Exception):
    """Mazrica API エラー"""
    def __init__(self, status_code: int, message: str):
        self.status_code = status_code
        self.message = message
        super().__init__(f"Mazrica API Error ({status_code}): {message}")


class MazricaClient:
    """Mazrica APIクライアント"""
    
    def __init__(self, api_key: Optional[str] = None, base_url: Optional[str] = None):
        self.api_key = api_key or Config.MAZRICA_API_KEY
        self.base_url = base_url or Config.MAZRICA_BASE_URL
        self.rate_limit = Config.API_RATE_LIMIT
        self._last_request_time = 0.0
        
        if not self.api_key:
            raise ValueError("Mazrica API Key is required")
    
    def _wait_for_rate_limit(self):
        """レート制限を守るための待機"""
        elapsed = time.time() - self._last_request_time
        if elapsed < self.rate_limit:
            time.sleep(self.rate_limit - elapsed)
        self._last_request_time = time.time()
    
    def _request(
        self,
        method: str,
        endpoint: str,
        params: Optional[dict] = None,
        json_body: Optional[dict] = None,
    ) -> dict:
        """APIリクエストを実行"""
        self._wait_for_rate_limit()

        url = f"{self.base_url}{endpoint}"
        headers = {
            "X-Api-Key": self.api_key,
            "Content-Type": "application/json"
        }

        kwargs = {"method": method, "url": url, "headers": headers}
        if params:
            kwargs["params"] = params
        if json_body:
            kwargs["json"] = json_body

        response = requests.request(**kwargs)

        if response.status_code == 429:
            # レート制限に達した場合は待機してリトライ
            time.sleep(1)
            return self._request(method, endpoint, params, json_body)

        if response.status_code >= 400:
            raise MazricaAPIError(
                status_code=response.status_code,
                message=response.text
            )

        if response.status_code == 204:
            return {}

        return response.json()
    
    def get_deal_types(self) -> list[dict]:
        """案件タイプ一覧を取得"""
        result = self._request("GET", "/deal_types")
        return result.get("dealTypes", [])
    
    def get_deals(
        self,
        deal_type_id: Optional[int] = None,
        page: int = 1,
        limit: int = 100,
        sort: str = "-updatedAt"
    ) -> dict:
        """
        案件一覧を取得
        
        Args:
            deal_type_id: 案件タイプIDでフィルタ
            page: ページ番号
            limit: 1ページあたりの件数
            sort: ソート順（デフォルト: 更新日時降順）
        
        Returns:
            {"deals": [...], "totalCount": int, "page": int}
        """
        params = {
            "page": page,
            "limit": limit,
            "sort": sort
        }
        
        if deal_type_id:
            params["dealTypeId"] = deal_type_id
        
        return self._request("GET", "/deals", params)
    
    def get_all_deals(
        self,
        deal_type_id: Optional[int] = None,
        limit_per_page: int = 100
    ) -> Generator[dict, None, None]:
        """
        全案件を取得（ページネーション対応）
        
        Args:
            deal_type_id: 案件タイプIDでフィルタ
            limit_per_page: 1ページあたりの件数
        
        Yields:
            案件データ
        """
        page = 1
        while True:
            result = self.get_deals(
                deal_type_id=deal_type_id,
                page=page,
                limit=limit_per_page
            )
            
            deals = result.get("deals", [])
            if not deals:
                break
            
            for deal in deals:
                yield deal
            
            total_count = result.get("totalCount", 0)
            if page * limit_per_page >= total_count:
                break
            
            page += 1
    
    def _parse_product_detail(self, data: dict) -> ProductDetail:
        """商品内訳データをパース"""
        # dealProductDetails は price / quantity / subtotal 等のキーを使う
        # （過去の productName/unitPrice/amount キーとの両方に対応）
        return ProductDetail(
            product_id=data.get("productId") or data.get("dealProductDetailTypeId"),
            product_name=data.get("productName") or data.get("name", ""),
            quantity=data.get("quantity"),
            unit_price=data.get("unitPrice") if data.get("unitPrice") is not None else data.get("price"),
            amount=data.get("amount") if data.get("amount") is not None else data.get("subtotal"),
            custom_fields=data.get("customFields") or data.get("dealProductDetailCustoms", {}),
        )

    @staticmethod
    def _format_custom_value(custom: dict) -> str:
        """dealCustomsの1要素を文字列化"""
        item_type = custom.get("itemType")
        if item_type == "date":
            dt = custom.get("datetime") or custom.get("value")
            if not dt:
                return ""
            # "2026-04-06T00:00:00+09:00" → "2026-04-06"
            return str(dt).split("T", 1)[0] if "T" in str(dt) else str(dt)
        if item_type == "single_select":
            opt = custom.get("selectedOption") or {}
            return opt.get("label", "") or ""
        if item_type == "multi_select":
            opts = custom.get("selectedOptions") or []
            return ", ".join(o.get("label", "") for o in opts if o.get("label"))
        if item_type == "text":
            return custom.get("text") or ""
        if item_type == "number":
            n = custom.get("number")
            return str(n) if n is not None else ""
        if item_type == "decimal_number":
            n = custom.get("decimalNumber")
            return str(n) if n is not None else ""
        if item_type == "url":
            return custom.get("url") or ""
        # fallback
        v = custom.get("value")
        if isinstance(v, list):
            return ", ".join(str(x) for x in v)
        return str(v) if v is not None else ""
    
    def _parse_deal(self, data: dict) -> Deal:
        """案件データをパース"""
        # 取引先情報
        customer = data.get("customer", {}) or {}
        
        # 案件タイプ情報
        deal_type = data.get("dealType", {}) or {}
        
        # フェーズ情報
        phase = data.get("phase", {}) or {}
        
        # 担当者情報
        user = data.get("user", {}) or {}
        
        # 商品情報（product フィールド）
        product = data.get("product", {}) or {}
        
        # 商品内訳詳細（dealProductDetails フィールド）
        product_details = [
            self._parse_product_detail(pd)
            for pd in data.get("dealProductDetails", [])
        ]

        # カスタム項目を dealCustomItemId でインデックス化して文字列化
        customs_by_id = {}
        for c in data.get("dealCustoms", []) or []:
            cid = c.get("dealCustomItemId")
            if cid is not None:
                customs_by_id[cid] = self._format_custom_value(c)

        return Deal(
            id=data.get("id"),
            name=data.get("name", ""),
            customer_name=customer.get("name"),
            customer_id=customer.get("id"),
            deal_type_id=deal_type.get("id"),
            deal_type_name=deal_type.get("name"),
            phase_name=phase.get("name"),
            amount=data.get("amount"),
            expected_contract_date=data.get("expectedContractDate"),
            created_at=data.get("createdAt", ""),
            updated_at=data.get("updatedAt", ""),
            user_name=user.get("name"),
            product_name=product.get("name"),
            product_id=product.get("id"),
            product_details=product_details,
            custom_fields=data.get("customFields", {}),
            customs_by_id=customs_by_id,
        )
    
    # ============================================================
    # 取引先 (Customers) - 書き込み・検索
    # ============================================================

    def _get_all_customers_cached(self) -> list[dict]:
        """全取引先を取得してキャッシュ（初回のみAPI呼び出し）"""
        if not hasattr(self, "_customers_cache"):
            customers = []
            page = 1
            while True:
                result = self._request("GET", "/customers", params={"page": page, "limit": 100})
                batch = result.get("customers", [])
                customers.extend(batch)
                total = result.get("totalCount", 0)
                if page * 100 >= total or not batch:
                    break
                page += 1
            self._customers_cache = customers
        return self._customers_cache

    def search_customers(self, name: str) -> list[dict]:
        """取引先を名前でクライアント側検索"""
        all_customers = self._get_all_customers_cached()
        results = []
        for c in all_customers:
            c_name = c.get("name") or ""
            if name.lower() in c_name.lower():
                results.append(c)
        return results

    def create_customer(self, name: str, **kwargs) -> dict:
        """取引先を新規登録。返り値は作成されたレコード"""
        body = {"name": name}
        for key in ("address", "telNo", "webUrl", "employee", "capital",
                     "closingMonth", "ownerRoleId", "customerCustoms"):
            if key in kwargs and kwargs[key] is not None:
                body[key] = kwargs[key]
        return self._request("POST", "/customers", json_body=body)

    def update_customer(self, customer_id: int, **kwargs) -> dict:
        """取引先を更新"""
        body = {}
        for key in ("name", "address", "telNo", "webUrl", "employee", "capital",
                     "closingMonth", "ownerRoleId", "customerCustoms"):
            if key in kwargs and kwargs[key] is not None:
                body[key] = kwargs[key]
        if not body:
            return {}
        return self._request("PATCH", f"/customers/{customer_id}", json_body=body)

    # ============================================================
    # コンタクト (Contacts) - 書き込み・検索
    # ============================================================

    def _get_all_contacts_cached(self) -> list[dict]:
        """全コンタクトを取得してキャッシュ（初回のみAPI呼び出し）"""
        if not hasattr(self, "_contacts_cache"):
            contacts = []
            page = 1
            while True:
                result = self._request("GET", "/contacts", params={"page": page, "limit": 100})
                batch = result.get("contacts", [])
                contacts.extend(batch)
                total = result.get("totalCount", 0)
                if page * 100 >= total or not batch:
                    break
                page += 1
            self._contacts_cache = contacts
        return self._contacts_cache

    def search_contacts(self, name: Optional[str] = None, email: Optional[str] = None) -> list[dict]:
        """コンタクトを名前またはメールでクライアント側検索"""
        if not name and not email:
            return []
        all_contacts = self._get_all_contacts_cached()
        results = []
        for c in all_contacts:
            c_name = c.get("name") or ""
            c_email = c.get("email") or ""
            if name and name in c_name:
                results.append(c)
            elif email and email.lower() == c_email.lower():
                results.append(c)
        return results

    def create_contact(self, name: str, customer_id: int, **kwargs) -> dict:
        """コンタクトを新規登録"""
        body = {"name": name, "customerId": customer_id}
        for key in ("email", "tel", "mobileTel", "dept", "position", "address", "contactCustoms"):
            if key in kwargs and kwargs[key] is not None:
                body[key] = kwargs[key]
        return self._request("POST", "/contacts", json_body=body)

    def update_contact(self, contact_id: int, **kwargs) -> dict:
        """コンタクトを更新"""
        body = {}
        for key in ("name", "email", "tel", "mobileTel", "dept", "position",
                     "address", "customerId", "contactCustoms"):
            if key in kwargs and kwargs[key] is not None:
                body[key] = kwargs[key]
        if not body:
            return {}
        return self._request("PATCH", f"/contacts/{contact_id}", json_body=body)

    # ============================================================
    # 案件 (Deals) - 書き込み・検索
    # ============================================================

    def search_deals(self, customer_id: Optional[int] = None, name: Optional[str] = None) -> list[dict]:
        """案件を検索"""
        filters = []
        if customer_id:
            filters.append({
                "valueFilter": {
                    "field": {"reference": "customerId"},
                    "comparisonOperator": "=",
                    "value": {"integerValue": customer_id},
                }
            })
        if name:
            filters.append({
                "valueFilter": {
                    "field": {"reference": "name"},
                    "comparisonOperator": "=",
                    "value": {"stringValue": name},
                }
            })

        if len(filters) > 1:
            body = {
                "filter": {
                    "compositeFilter": {
                        "logicalOperator": "AND",
                        "subFilters": filters,
                    }
                },
                "limit": 10,
            }
        elif filters:
            body = {"filter": filters[0], "limit": 10}
        else:
            return []

        result = self._request("POST", "/deals/search", json_body=body)
        return result.get("deals", [])

    def create_deal(self, name: str, customer_id: int, deal_type_id: int, **kwargs) -> dict:
        """案件を新規登録"""
        body = {"name": name, "customerId": customer_id, "dealTypeId": deal_type_id}
        for key in ("amount", "phaseId", "userId", "expectedContractDate", "dealCustoms"):
            if key in kwargs and kwargs[key] is not None:
                body[key] = kwargs[key]
        return self._request("POST", "/deals", json_body=body)

    def update_deal(self, deal_id: int, **kwargs) -> dict:
        """案件を更新"""
        body = {}
        for key in ("name", "amount", "phaseId", "userId", "customerId",
                     "expectedContractDate", "dealCustoms"):
            if key in kwargs and kwargs[key] is not None:
                body[key] = kwargs[key]
        if not body:
            return {}
        return self._request("PATCH", f"/deals/{deal_id}", json_body=body)

    # ============================================================
    # 既存メソッド
    # ============================================================

    def fetch_deals_with_products(
        self,
        deal_type_id: Optional[int] = None
    ) -> list[Deal]:
        """
        案件一覧を商品内訳付きで取得
        
        Args:
            deal_type_id: 案件タイプIDでフィルタ
        
        Returns:
            案件リスト
        """
        deals = []
        for deal_data in self.get_all_deals(deal_type_id=deal_type_id):
            deal = self._parse_deal(deal_data)
            deals.append(deal)
        
        return deals


# テスト用
if __name__ == "__main__":
    from .config import load_dotenv
    load_dotenv()
    
    client = MazricaClient()
    
    # 案件タイプ一覧を取得
    print("=== 案件タイプ一覧 ===")
    deal_types = client.get_deal_types()
    for dt in deal_types:
        print(f"  ID: {dt.get('id')}, Name: {dt.get('name')}")
    
    # 案件一覧を取得（最初の10件）
    print("\n=== 案件一覧（最初の10件） ===")
    result = client.get_deals(limit=10)
    for deal_data in result.get("deals", [])[:10]:
        deal = client._parse_deal(deal_data)
        print(f"  ID: {deal.id}, Name: {deal.name}, Customer: {deal.customer_name}")
        if deal.product_details:
            for pd in deal.product_details:
                print(f"    - Product: {pd.product_name}, Amount: {pd.amount}")

